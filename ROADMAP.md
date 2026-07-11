# off2d roadmap

Cross-platform (Linux / Windows / macOS) CLI to convert MS Office files to
plain text, installable in locked-down user space (HPC Rocky Linux 9) via
`uv` with no system dependencies.

## Approach

Python + `uv`. Pure-Python, wheel-only dependencies so no compiler or
external binary is ever required.

- docx -> markdown: `mammoth` (docx -> HTML) + `markdownify` (HTML -> GFM)
- xlsx -> tab/csv: `openpyxl`
- packaging: `pyproject.toml`, console entry point `off2d`

Install:
```
uv tool install git+<repo-url>        # or: uvx --from git+<repo-url> off2d
uv tool install .                     # from a clone
pipx install off2d                    # also works (pip-compatible)
```

## Milestones

### M1 - Scaffold + spreadsheet path (MVP)  [DONE]
- [x] `pyproject.toml`, `src/off2d/` layout, `off2d` entry point
- [x] argparse CLI with `parse_intermixed_args` (switches after filename)
- [x] input-type detection (extension + zip-content sniff)
- [x] xlsx -> TSV of first sheet to stdout (stdlib-only reader; raw cached
      values at full precision; date cells -> ISO 8601)
- [x] `-s N|NAME` worksheet selection (1-based index or sheet name)
- [x] `-o PATH` with format guess: `.csv` comma, `.tab/.tsv/.txt` tab
- [x] `.gz` output extension -> gzip stream
- [x] tests against `test/ST7-coloc_results.xlsx` (`python -m unittest`)

Note: M1 dropped the planned `openpyxl` dependency entirely -- the xlsx
reader is pure standard library (`zipfile` + `xml.etree`), so the default
install has zero third-party dependencies.

### M2 - docx -> markdown
- [ ] mammoth -> HTML -> markdownify GFM to stdout
- [ ] preserve headings, bold/italic, links, bulleted/numbered lists, tables
- [ ] `-o PATH` writes `.md`; `.gz` supported
- [ ] tests against the genotype-formats docx (assert 28 links + tables survive)

### M3 - Images (opt-in)
- [ ] `-i` / `--images`: extract embedded media to a sibling folder
- [ ] rewrite references Obsidian-style `![[image.png]]`
- [ ] configurable image output dir; sensible default next to `-o` target

### M4 - Packaging & distribution
- [ ] publish to PyPI (or tagged git installs) for `uvx off2d`
- [ ] CI matrix: Linux / Windows / macOS x Python 3.9-3.13
- [ ] optional: standalone build (PyInstaller/shiv) for no-Python hosts

## Open questions / decisions

- Sheet selection: support both 1-based index and sheet name via `-s`.
  Ambiguity (a sheet literally named "2") resolves to name-match first,
  else index. Document the rule.
- csv quoting: default `csv.QUOTE_MINIMAL`; revisit if fields contain tabs.
- Empty-cell / merged-cell handling in xlsx: emit empty string; keep row
  rectangular to the max used column.
- Markdown table fidelity for merged/nested docx tables is best-effort.
- Should `-i` be implied when a docx has images and `-o` is a file? No -
  keep opt-in per design principle 5.
