# CLAUDE.md - off2d

Guidance for working in this repository.

## What this is

`off2d` ("Office to text") is a cross-platform CLI that converts MS Office
files to plain text on stdout:

- `off2d file.docx`  -> GitHub-flavored Markdown to stdout
- `off2d sheet.xlsx` -> tab-delimited first worksheet to stdout

Input type is auto-detected from the file (extension first, zip-content
sniff as fallback). No long switches are required for the common case.

## Design principles (do not regress these)

1. **Zero system dependencies.** Pure-Python dependencies with prebuilt
   wheels only. Never add a dependency that needs a compiler, or an
   external binary (pandoc, libreoffice, etc.). Target install is an
   isolated `uv tool` / `uvx` env on a locked-down HPC (Rocky Linux 9).
2. **One obvious argument.** The primary interface is a single positional
   file path. Switches are optional and must be accepted **before or
   after** the positional (use `argparse.parse_intermixed_args`).
3. **stdout by default.** Conversions print to stdout so the tool composes
   in pipes. `-o PATH` redirects to a file.
4. **Guess intent from names.** Output format is inferred from the `-o`
   file extension, not a separate `--format` flag (though a flag may exist
   as an override). Input type inferred from the input file.
5. **Fidelity for docx.** Preserve links, tables, headings, emphasis, and
   lists in Markdown. Prefer semantic HTML (mammoth) -> GFM (markdownify)
   over lossy direct converters.

## Architecture

Stack: Python >= 3.9, packaged with `pyproject.toml`, entry point `off2d`.

Dependencies:
- xlsx path (M1): **none** - pure standard library (`zipfile`, `xml.etree`,
  `csv`, `gzip`). Do not add `openpyxl` or any third-party dep to this path.
- docx path (M2, planned): a small pure-Python/wheel-only set (`mammoth` +
  `markdownify`) under the optional `[docx]` extra, or a vendored minimal
  converter. Keep it out of the default install.

Module layout (`src/off2d/`):
- `cli.py`       - argparse (`parse_intermixed_args`), dispatch by input type
- `detect.py`    - input type detection (extension + zip sniff)
- `xlsx.py`      - stdlib xlsx reader (sheets, shared strings, styles, rows)
- `sheet2txt.py` - rows -> tab/csv; delimiter/gzip guessed from `-o` name
- `docx2md.py`   - docx -> markdown; optional image extraction (M2, TODO)

## CLI contract

```
off2d INPUT [options]

  -s, --sheet N|NAME   worksheet to convert (1-based index or name; default 1)
  -o, --output PATH    write to file; format guessed from extension
  -i, --images         (docx) extract images, reference Obsidian-style ![[img]]
```

Output-format guessing for `-o`:
- strip a trailing `.gz` -> gzip-compress the stream
- remaining ext `.csv` -> comma; `.tab`/`.tsv`/`.txt` -> tab (default tab)

## Testing

Fixtures in `test/`:
- `Modern-Genotype-Storage-Formats-for-Large-Cohorts.docx` - 28 hyperlinks,
  tables, lists, no embedded images.
- `ST7-coloc_results.xlsx` - multiple sheets (`coloc_pass`, `coloc_summary`).

Verify by running conversions and diffing stdout; assert links, tables, and
sheet selection survive. Do not commit large generated outputs.

## Conventions

- ASCII-only in source and docs (no smart quotes / em dashes / emoji).
- Keep the common path dependency-light and fast to start.
- New features stay opt-in; never change default stdout behavior.
