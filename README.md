## off2d

Platform-independent command-line tool to convert MS Office documents and
spreadsheets (docx, xlsx) to plain text: Markdown for documents,
tab/CSV (optionally gzipped) for spreadsheets.

### Why

Convenient, dependency-light conversion that installs in locked-down user
space (e.g. an HPC cluster) via `uv`, with no system packages or external
binaries. One obvious argument; switches optional and position-independent.

### Usage

```
off2d file.docx              # GitHub-flavored Markdown to stdout
off2d sheet.xlsx             # tab-delimited first worksheet to stdout
off2d sheet.xlsx -s 2        # second worksheet (index or name)
off2d sheet.xlsx -o out.csv  # CSV to file (format guessed from extension)
off2d sheet.xlsx -o out.tab.gz   # gzipped tab-delimited
off2d file.docx -i -o doc.md     # extract images, Obsidian-style refs
```

Switches may appear before or after the input file.

### Install (planned)

```
uv tool install git+<repo-url>          # isolated tool env, no system deps
uvx --from git+<repo-url> off2d file.docx
```

### Status

Early development. See ROADMAP.md for the plan and CLAUDE.md for design
principles and architecture.
