"""Input-type detection: extension first, zip-content sniff as fallback."""

import zipfile

DOCX = "docx"
XLSX = "xlsx"

_EXT = {
    ".docx": DOCX,
    ".docm": DOCX,
    ".xlsx": XLSX,
    ".xlsm": XLSX,
}


def detect_type(path):
    """Return DOCX, XLSX, or None for the given file path.

    Extension match wins. If the extension is unknown but the file is a zip
    container, sniff for the marker member of each Office format.
    """
    lower = str(path).lower()
    for ext, kind in _EXT.items():
        if lower.endswith(ext):
            return kind
    if zipfile.is_zipfile(path):
        try:
            with zipfile.ZipFile(path) as z:
                names = set(z.namelist())
        except (zipfile.BadZipFile, OSError):
            return None
        if "word/document.xml" in names:
            return DOCX
        if "xl/workbook.xml" in names:
            return XLSX
    return None
