"""Write spreadsheet rows as delimited text, with format/gzip inferred from
the output path extension."""

import csv
import gzip
import sys
import contextlib

from . import xlsx


def _guess_format(out_path):
    """Return (delimiter, gzip_bool) from an output filename.

    Strip a trailing '.gz' (-> gzip), then match the remaining extension:
    '.csv' -> comma; '.tab'/'.tsv'/'.txt' -> tab. Default is tab.
    """
    name = out_path.lower()
    use_gzip = name.endswith(".gz")
    if use_gzip:
        name = name[:-3]
    if name.endswith(".csv"):
        delim = ","
    elif name.endswith((".tab", ".tsv", ".txt")):
        delim = "\t"
    else:
        delim = "\t"
    return delim, use_gzip


@contextlib.contextmanager
def _open_output(out_path):
    """Yield (text_stream, delimiter). None path -> stdout, tab-delimited."""
    if out_path is None:
        yield sys.stdout, "\t"
        return
    delim, use_gzip = _guess_format(out_path)
    if use_gzip:
        fh = gzip.open(out_path, "wt", newline="", encoding="utf-8")
    else:
        fh = open(out_path, "w", newline="", encoding="utf-8")
    try:
        yield fh, delim
    finally:
        fh.close()


def convert(path, sheet_selector=None, out_path=None):
    """Convert one worksheet to delimited text. Returns the sheet name used."""
    name, rows = xlsx.read_sheet_rows(path, sheet_selector)
    with _open_output(out_path) as (fh, delim):
        writer = csv.writer(fh, delimiter=delim, lineterminator="\n")
        writer.writerows(rows)
    return name
