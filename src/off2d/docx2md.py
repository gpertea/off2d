"""Minimal, vendored docx -> GitHub-flavored Markdown converter.

Standard library only. Handles the constructs that matter for readable
Markdown fidelity:

  - headings          (pStyle Heading N / Title, plus a bold+larger-font
                       heuristic for docs that only size their headings)
  - bold / italic     (w:b, w:i)
  - inline code        (character style whose name/id contains "Verbatim"/"Code")
  - fenced code blocks (paragraph style "SourceCode" / "Code")
  - hyperlinks         (external via rels, internal anchors)
  - lists              (numPr -> bullet/ordered, nested by ilvl)
  - tables             (GFM pipe tables; first row treated as header)
  - images             (opt-in extraction + Obsidian ![[name]] embeds)

Anything unrecognized degrades to its plain text.
"""

import os
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _w(tag):
    return "{%s}%s" % (W, tag)


def _r(tag):
    return "{%s}%s" % (R, tag)


# Characters that carry Markdown meaning in running text.
_ESCAPE_RE = re.compile(r"([\\`*_\[\]<>])")


def _escape(text):
    return _ESCAPE_RE.sub(r"\\\1", text)


class Docx:
    def __init__(self, path):
        self._zip = zipfile.ZipFile(path)
        self._rels = self._load_rels()
        self._numbering = self._load_numbering()
        self._code_char_styles, self._heading_styles = self._load_styles()

    def close(self):
        self._zip.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    def _read(self, name):
        return self._zip.read(name)

    def _has(self, name):
        return name in self._zip.namelist()

    # -- relationships (hyperlinks, images) --------------------------------

    def _load_rels(self):
        rels = {}
        name = "word/_rels/document.xml.rels"
        if not self._has(name):
            return rels
        root = ET.fromstring(self._read(name))
        for rel in root:
            rid = rel.get("Id")
            target = rel.get("Target")
            mode = rel.get("TargetMode", "Internal")
            rtype = rel.get("Type", "")
            if rid and target:
                rels[rid] = {"target": target, "mode": mode, "type": rtype}
        return rels

    # -- numbering (list format per numId/ilvl) ----------------------------

    def _load_numbering(self):
        """Return numId -> {ilvl: 'bullet'|'decimal'}."""
        result = {}
        if not self._has("word/numbering.xml"):
            return result
        root = ET.fromstring(self._read("word/numbering.xml"))
        abstract = {}  # abstractNumId -> {ilvl: fmt}
        for an in root.iter(_w("abstractNum")):
            aid = an.get(_w("abstractNumId"))
            levels = {}
            for lvl in an.iter(_w("lvl")):
                ilvl = lvl.get(_w("ilvl"))
                fmt_el = lvl.find(_w("numFmt"))
                fmt = fmt_el.get(_w("val")) if fmt_el is not None else "bullet"
                levels[ilvl] = fmt
            abstract[aid] = levels
        for num in root.iter(_w("num")):
            nid = num.get(_w("numId"))
            aref = num.find(_w("abstractNumId"))
            if aref is not None:
                result[nid] = abstract.get(aref.get(_w("val")), {})
        return result

    # -- styles (code character styles, heading paragraph styles) ----------

    def _load_styles(self):
        code_char = set()
        heading = {}  # styleId -> level
        if not self._has("word/styles.xml"):
            return code_char, heading
        root = ET.fromstring(self._read("word/styles.xml"))
        for st in root.iter(_w("style")):
            sid = st.get(_w("styleId")) or ""
            stype = st.get(_w("type")) or ""
            name_el = st.find(_w("name"))
            name = (name_el.get(_w("val")) if name_el is not None else "") or ""
            hay = (sid + " " + name).lower()
            if stype == "character" and ("verbatim" in hay or "code" in hay):
                code_char.add(sid)
            if stype == "paragraph":
                m = re.search(r"heading\s*([1-9])", hay)
                if m:
                    heading[sid] = int(m.group(1))
                elif "title" in hay:
                    heading[sid] = 1
                elif "subtitle" in hay:
                    heading[sid] = 2
        return code_char, heading

    # -- helpers ------------------------------------------------------------

    def _para_style(self, p):
        ppr = p.find(_w("pPr"))
        if ppr is None:
            return None
        pstyle = ppr.find(_w("pStyle"))
        return pstyle.get(_w("val")) if pstyle is not None else None

    def _is_code_para(self, p):
        sid = self._para_style(p) or ""
        return sid.lower() in ("sourcecode", "code")

    def _numpr(self, p):
        ppr = p.find(_w("pPr"))
        if ppr is None:
            return None
        numpr = ppr.find(_w("numPr"))
        if numpr is None:
            return None
        ilvl_el = numpr.find(_w("ilvl"))
        numid_el = numpr.find(_w("numId"))
        ilvl = int(ilvl_el.get(_w("val"))) if ilvl_el is not None else 0
        numid = numid_el.get(_w("val")) if numid_el is not None else None
        if numid is None:
            return None
        fmt = self._numbering.get(numid, {}).get(str(ilvl), "bullet")
        return {"ilvl": ilvl, "ordered": fmt != "bullet"}


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

class _Segment:
    __slots__ = ("text", "bold", "italic", "code", "href")

    def __init__(self, text, bold, italic, code, href):
        self.text = text
        self.bold = bold
        self.italic = italic
        self.code = code
        self.href = href

    def key(self):
        return (self.bold, self.italic, self.code, self.href)


class Converter:
    def __init__(self, docx, extract_images=False, image_dir=None):
        self.dx = docx
        self.extract_images = extract_images
        self.image_dir = image_dir
        self.images_written = []
        self._body = self._load_body()
        self._body_size, self._heading_size_level = self._analyze_headings()

    def _load_body(self):
        root = ET.fromstring(self.dx._read("word/document.xml"))
        return root.find(_w("body"))

    # -- heading size heuristic (first pass) --------------------------------

    def _run_size(self, r):
        rpr = r.find(_w("rPr"))
        if rpr is None:
            return None
        sz = rpr.find(_w("sz"))
        return int(sz.get(_w("val"))) if sz is not None else None

    def _para_profile(self, p):
        """Return (has_text, fully_bold, dominant_size)."""
        sizes = []
        bolds = []
        has_text = False
        for r in p.iter(_w("r")):
            txt = "".join(t.text or "" for t in r.iter(_w("t")))
            if not txt.strip():
                continue
            has_text = True
            rpr = r.find(_w("rPr"))
            bold = rpr is not None and rpr.find(_w("b")) is not None
            bolds.append(bold)
            sizes.append(self._run_size(r) or 22)
        if not has_text:
            return False, False, 22
        dominant = max(set(sizes), key=sizes.count)
        return True, all(bolds), dominant

    def _analyze_headings(self):
        """Body text size + a map {font_size -> heading level} for docs that
        signal headings by bold+size rather than by paragraph style."""
        body_sizes = {}
        heading_sizes = set()
        if self._body is None:
            return 22, {}
        for p in self._body.findall(_w("p")):
            if self.dx._is_code_para(p):
                continue
            has_text, fully_bold, size = self._para_profile(p)
            if not has_text:
                continue
            if not fully_bold:
                body_sizes[size] = body_sizes.get(size, 0) + 1
        body_size = max(body_sizes, key=body_sizes.get) if body_sizes else 22
        for p in self._body.findall(_w("p")):
            if self.dx._is_code_para(p) or self.dx._numpr(p):
                continue
            has_text, fully_bold, size = self._para_profile(p)
            if has_text and fully_bold and size > body_size:
                heading_sizes.add(size)
        levels = {}
        for i, size in enumerate(sorted(heading_sizes, reverse=True)):
            levels[size] = i + 1
        return body_size, levels

    # -- inline runs --------------------------------------------------------

    def _render_runs(self, container, href=None):
        """Convert the runs within a container (paragraph or hyperlink) into a
        list of _Segment, following nested hyperlinks."""
        segments = []
        for child in container:
            tag = child.tag
            if tag == _w("hyperlink"):
                url = self._hyperlink_target(child)
                segments.extend(self._render_runs(child, href=url or href))
            elif tag == _w("r"):
                segments.extend(self._render_one_run(child, href))
        return segments

    def _hyperlink_target(self, hl):
        rid = hl.get(_r("id"))
        if rid and rid in self.dx._rels:
            return self.dx._rels[rid]["target"]
        anchor = hl.get(_w("anchor"))
        if anchor:
            return "#" + anchor
        return None

    def _render_one_run(self, r, href):
        rpr = r.find(_w("rPr"))
        bold = rpr is not None and rpr.find(_w("b")) is not None
        italic = rpr is not None and rpr.find(_w("i")) is not None
        code = False
        if rpr is not None:
            rstyle = rpr.find(_w("rStyle"))
            if rstyle is not None and rstyle.get(_w("val")) in self.dx._code_char_styles:
                code = True
        out = []
        text_parts = []
        for node in r.iter():
            if node.tag == _w("t"):
                text_parts.append(node.text or "")
            elif node.tag == _w("tab"):
                text_parts.append("\t")
            elif node.tag == _w("br"):
                text_parts.append("\n")
        text = "".join(text_parts)
        if text:
            out.append(_Segment(text, bold, italic, code, href))
        # images inside the run
        for blip in r.iter("{%s}blip" % A):
            embed = blip.get(_r("embed"))
            img_md = self._image_markdown(embed)
            if img_md:
                out.append(_Segment(img_md, False, False, False, None))
        return out

    def _segments_to_md(self, segments, in_heading=False):
        # merge adjacent segments sharing formatting
        merged = []
        for s in segments:
            if merged and merged[-1].key() == s.key():
                merged[-1].text += s.text
            else:
                merged.append(_Segment(s.text, s.bold, s.italic, s.code, s.href))
        parts = []
        for s in merged:
            if s.code:
                # inline code: no escaping inside; keep raw text stripped of
                # newlines
                token = "`" + s.text.replace("`", "") + "`"
            elif s.text.startswith("![["):  # image embed, already markdown
                token = s.text
                parts.append(token)
                continue
            else:
                token = _escape(s.text)
                if s.bold and not in_heading:
                    token = "**" + token + "**"
                if s.italic:
                    token = "*" + token + "*"
            if s.href:
                token = "[%s](%s)" % (token, s.href)
            parts.append(token)
        return "".join(parts)

    # -- images -------------------------------------------------------------

    def _image_markdown(self, rid):
        if not rid or rid not in self.dx._rels:
            return None
        rel = self.dx._rels[rid]
        target = rel["target"]  # e.g. media/image1.png (relative to word/)
        if not self.extract_images:
            return None
        src = posixpath.normpath(posixpath.join("word", target))
        if src not in self.dx._zip.namelist():
            return None
        fname = posixpath.basename(target)
        out_dir = self.image_dir or "."
        try:
            os.makedirs(out_dir, exist_ok=True)
            dest = os.path.join(out_dir, fname)
            with open(dest, "wb") as fh:
                fh.write(self.dx._read(src))
            self.images_written.append(dest)
        except OSError:
            return None
        return "![[%s]]" % fname

    # -- block-level --------------------------------------------------------

    def _paragraph_kind(self, p):
        numpr = self.dx._numpr(p)
        if numpr is not None:
            return ("list", numpr)
        if self.dx._is_code_para(p):
            return ("code", None)
        sid = self.dx._para_style(p)
        if sid and sid in self.dx._heading_styles:
            return ("heading", self.dx._heading_styles[sid])
        has_text, fully_bold, size = self._para_profile(p)
        if has_text and fully_bold and size in self._heading_size_level:
            return ("heading", self._heading_size_level[size])
        return ("para", None)

    def _para_text(self, p, in_heading=False):
        return self._segments_to_md(self._render_runs(p), in_heading=in_heading)

    def convert(self):
        blocks = []
        code_buffer = None  # collecting consecutive code lines
        if self._body is None:
            return ""
        for el in list(self._body):
            if el.tag == _w("tbl"):
                if code_buffer is not None:
                    blocks.append("```\n" + "\n".join(code_buffer) + "\n```")
                    code_buffer = None
                blocks.append(self._render_table(el))
                continue
            if el.tag != _w("p"):
                continue
            kind, meta = self._paragraph_kind(el)
            if kind == "code":
                line = self._plain_text(el)
                if code_buffer is None:
                    code_buffer = []
                code_buffer.append(line)
                continue
            if code_buffer is not None:
                blocks.append("```\n" + "\n".join(code_buffer) + "\n```")
                code_buffer = None
            if kind == "heading":
                text = self._para_text(el, in_heading=True).strip()
                if text:
                    blocks.append("#" * min(meta, 6) + " " + text)
            elif kind == "list":
                indent = "  " * meta["ilvl"]
                marker = "1." if meta["ordered"] else "-"
                text = self._para_text(el).strip()
                blocks.append(("LIST", indent + marker + " " + text))
            else:
                text = self._para_text(el).strip()
                if text:
                    blocks.append(text)
        if code_buffer is not None:
            blocks.append("```\n" + "\n".join(code_buffer) + "\n```")
        return self._join_blocks(blocks)

    def _join_blocks(self, blocks):
        """Blank line between blocks, but keep list items tight together."""
        out = []
        prev_is_list = False
        for b in blocks:
            is_list = isinstance(b, tuple) and b[0] == "LIST"
            text = b[1] if is_list else b
            if out:
                if is_list and prev_is_list:
                    out.append("\n")
                else:
                    out.append("\n\n")
            out.append(text)
            prev_is_list = is_list
        return "".join(out) + "\n"

    def _plain_text(self, p):
        parts = []
        for node in p.iter():
            if node.tag == _w("t"):
                parts.append(node.text or "")
            elif node.tag == _w("tab"):
                parts.append("\t")
            elif node.tag in (_w("br"), _w("cr")):
                parts.append("\n")
        return "".join(parts)

    # -- tables -------------------------------------------------------------

    def _cell_md(self, tc):
        pieces = []
        for p in tc.findall(_w("p")):
            txt = self._para_text(p).strip()
            if txt:
                pieces.append(txt)
        # pipes and newlines break GFM cells
        return "<br>".join(pieces).replace("|", "\\|")

    def _render_table(self, tbl):
        rows = []
        for tr in tbl.findall(_w("tr")):
            cells = [self._cell_md(tc) for tc in tr.findall(_w("tc"))]
            rows.append(cells)
        if not rows:
            return ""
        width = max(len(r) for r in rows)
        rows = [r + [""] * (width - len(r)) for r in rows]
        header = rows[0]
        lines = ["| " + " | ".join(header) + " |"]
        lines.append("| " + " | ".join(["---"] * width) + " |")
        for r in rows[1:]:
            lines.append("| " + " | ".join(r) + " |")
        return "\n".join(lines)


def convert(path, out_path=None, extract_images=False, image_dir=None):
    """Convert a docx to Markdown. Writes to out_path or returns the string
    (caller prints to stdout). Returns (markdown, images_written)."""
    if extract_images and image_dir is None:
        if out_path:
            stem = os.path.splitext(out_path)[0]
        else:
            stem = os.path.splitext(os.path.basename(path))[0]
        image_dir = stem + "_media"
    with Docx(path) as dx:
        conv = Converter(dx, extract_images=extract_images, image_dir=image_dir)
        md = conv.convert()
        images = list(conv.images_written)
    if out_path:
        import gzip
        opener = gzip.open if out_path.lower().endswith(".gz") else open
        with opener(out_path, "wt", encoding="utf-8") as fh:
            fh.write(md)
        return md, images
    return md, images
