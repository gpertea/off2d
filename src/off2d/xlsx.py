"""Minimal, dependency-free xlsx reader (Python standard library only).

An xlsx file is a zip of XML parts. We read only what is needed to emit cell
*values* (not display formatting) as rows of strings:

  - workbook.xml + its rels  -> ordered (name, part) list of worksheets
  - sharedStrings.xml        -> shared string table
  - styles.xml               -> per-cell-style number-format id (for dates)
  - worksheets/sheetN.xml    -> the cells

Numbers are emitted with full stored precision (their cached <v> text), which
is what you want for data extraction. Cells whose style is a date format are
converted from the Excel serial to ISO 8601.
"""

import datetime
import re
import zipfile
import xml.etree.ElementTree as ET

# Built-in number-format ids that denote dates/times (ECMA-376, plus common
# East-Asian date formats). Custom formats are detected by their format code.
_DATE_BUILTIN_IDS = set(range(14, 23)) | set(range(27, 37)) | \
    set(range(45, 48)) | set(range(50, 59)) | {30, 36}

# Excel's day 0 is 1899-12-30 (the 1900 leap-year bug shifts the epoch back).
_EPOCH_1900 = datetime.datetime(1899, 12, 30)
_EPOCH_1904 = datetime.datetime(1904, 1, 1)


def _local(tag):
    """Strip the XML namespace from a tag: '{ns}row' -> 'row'."""
    return tag.rsplit("}", 1)[-1]


def _col_index(cell_ref):
    """'B3' -> 1 (0-based column index). 'AA10' -> 26."""
    letters = ""
    for ch in cell_ref:
        if ch.isalpha():
            letters += ch
        else:
            break
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def _looks_like_date_code(code):
    """Heuristic: a format code with date/time tokens outside quotes/brackets."""
    # Drop quoted literals and [bracketed] sections (colors, conditions, locale).
    stripped = re.sub(r'"[^"]*"', "", code)
    stripped = re.sub(r"\[[^\]]*\]", "", stripped)
    return bool(re.search(r"[ymdhsYMDHS]", stripped))


class Xlsx:
    def __init__(self, path):
        self._zip = zipfile.ZipFile(path)
        self._shared = None
        self._date_xf = None      # list: is cell-style index a date? (by xf order)
        self._sheets = None       # list of (name, part_path)
        self._date1904 = None

    def close(self):
        self._zip.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # -- part loading -------------------------------------------------------

    def _read(self, name):
        return self._zip.read(name)

    def _iter(self, name):
        with self._zip.open(name) as fh:
            for event, elem in ET.iterparse(fh, events=("end",)):
                yield elem

    # -- workbook / sheet listing ------------------------------------------

    def sheets(self):
        """Ordered list of (sheet_name, part_path) as they appear as tabs."""
        if self._sheets is not None:
            return self._sheets
        # rId -> target part
        rels = {}
        for elem in self._iter("xl/_rels/workbook.xml.rels"):
            if _local(elem.tag) == "Relationship":
                rid = elem.get("Id")
                target = elem.get("Target")
                if rid and target:
                    if not target.startswith("/"):
                        target = "xl/" + target.lstrip("./")
                    else:
                        target = target.lstrip("/")
                    rels[rid] = target
        sheets = []
        self._date1904 = False
        for elem in self._iter("xl/workbook.xml"):
            tag = _local(elem.tag)
            if tag == "workbookPr":
                val = elem.get("date1904")
                if val in ("1", "true", "True"):
                    self._date1904 = True
            elif tag == "sheet":
                name = elem.get("name") or ""
                rid = None
                for k, v in elem.attrib.items():
                    if _local(k) == "id":
                        rid = v
                        break
                part = rels.get(rid)
                if part:
                    sheets.append((name, part))
        self._sheets = sheets
        return sheets

    # -- shared strings -----------------------------------------------------

    def _shared_strings(self):
        if self._shared is not None:
            return self._shared
        result = []
        if "xl/sharedStrings.xml" not in self._zip.namelist():
            self._shared = result
            return result
        # Accumulate text of each <si>, concatenating its <t> runs.
        parts = []
        for elem in self._iter("xl/sharedStrings.xml"):
            tag = _local(elem.tag)
            if tag == "t":
                parts.append(elem.text or "")
            elif tag == "si":
                result.append("".join(parts))
                parts = []
        self._shared = result
        return result

    # -- styles (date detection) -------------------------------------------

    # styles.xml has both a cellStyleXfs and a cellXfs block of <xf> elements;
    # cell `s` indices refer to cellXfs only, so parse structurally.
    def _date_flags_structural(self):
        if self._date_xf is not None:
            return self._date_xf
        flags = []
        if "xl/styles.xml" not in self._zip.namelist():
            self._date_xf = flags
            return flags
        root = ET.fromstring(self._read("xl/styles.xml"))
        custom = {}
        for el in root.iter():
            if _local(el.tag) == "numFmt":
                fid = el.get("numFmtId")
                if fid is not None:
                    custom[int(fid)] = _looks_like_date_code(el.get("formatCode") or "")
        cellxfs = None
        for el in root:
            if _local(el.tag) == "cellXfs":
                cellxfs = el
                break
        if cellxfs is not None:
            for xf in cellxfs:
                if _local(xf.tag) != "xf":
                    continue
                fid = int(xf.get("numFmtId", "0"))
                flags.append(fid in _DATE_BUILTIN_IDS or custom.get(fid, False))
        self._date_xf = flags
        return flags

    def _serial_to_iso(self, serial):
        epoch = _EPOCH_1904 if self._date1904 else _EPOCH_1900
        try:
            dt = epoch + datetime.timedelta(days=serial)
        except (OverflowError, ValueError):
            return str(serial)
        if serial == int(serial):
            return dt.strftime("%Y-%m-%d")
        return dt.strftime("%Y-%m-%d %H:%M:%S")

    # -- row iteration ------------------------------------------------------

    def rows(self, part_path):
        """Yield rows (list of str) for a worksheet part. Internal empty cells
        are preserved; rows are padded to a common width by the caller."""
        shared = self._shared_strings()
        date_flags = self._date_flags_structural()
        for elem in self._iter(part_path):
            if _local(elem.tag) != "row":
                continue
            cells = {}
            max_col = -1
            for c in elem:
                if _local(c.tag) != "c":
                    continue
                ref = c.get("r") or ""
                col = _col_index(ref) if ref else (max_col + 1)
                ctype = c.get("t")
                style = c.get("s")
                value = self._cell_value(c, ctype, style, shared, date_flags)
                cells[col] = value
                if col > max_col:
                    max_col = col
            row = [cells.get(i, "") for i in range(max_col + 1)]
            yield row
            elem.clear()

    def _cell_value(self, c, ctype, style, shared, date_flags):
        v_text = None
        is_text = None  # inlineStr <is>
        for child in c:
            ct = _local(child.tag)
            if ct == "v":
                v_text = child.text
            elif ct == "is":
                is_text = "".join(
                    (t.text or "") for t in child.iter() if _local(t.tag) == "t"
                )
        if ctype == "s":  # shared string
            if v_text is None:
                return ""
            try:
                return shared[int(v_text)]
            except (ValueError, IndexError):
                return ""
        if ctype == "inlineStr":
            return is_text or ""
        if ctype == "str":  # formula string result
            return v_text or ""
        if ctype == "b":  # boolean
            return "TRUE" if (v_text or "").strip() in ("1", "true") else "FALSE"
        if ctype == "e":  # error
            return v_text or ""
        # default: numeric (t absent or "n")
        if v_text is None:
            return ""
        if style is not None and date_flags:
            try:
                if date_flags[int(style)]:
                    return self._serial_to_iso(float(v_text))
            except (ValueError, IndexError):
                pass
        return v_text


def read_sheet_rows(path, sheet_selector=None):
    """Open `path`, resolve the selected sheet, and return (name, rows).

    sheet_selector: None (first sheet), a sheet name, or a 1-based index (str
    or int). Name match wins over index if both are plausible.
    """
    xl = Xlsx(path)
    try:
        sheets = xl.sheets()
        if not sheets:
            raise ValueError("no worksheets found")
        idx = _resolve_sheet(sheets, sheet_selector)
        name, part = sheets[idx]
        # Materialize rows and pad to the widest row for rectangular output.
        rows = list(xl.rows(part))
        width = max((len(r) for r in rows), default=0)
        rows = [r + [""] * (width - len(r)) for r in rows]
        return name, rows
    finally:
        xl.close()


def _resolve_sheet(sheets, selector):
    if selector is None:
        return 0
    names = [n for n, _ in sheets]
    sel = str(selector)
    if sel in names:  # name match wins
        return names.index(sel)
    try:
        i = int(sel)
    except ValueError:
        raise ValueError("no sheet named %r" % sel)
    if 1 <= i <= len(sheets):
        return i - 1
    raise ValueError(
        "sheet index %d out of range (1..%d)" % (i, len(sheets))
    )
