"""off2d command-line entry point."""

import argparse
import sys

from . import __version__, detect, sheet2txt


def _build_parser():
    p = argparse.ArgumentParser(
        prog="off2d",
        description="Convert MS Office files to plain text. "
        "docx -> Markdown, xlsx -> tab/CSV. Output goes to stdout unless -o "
        "is given.",
    )
    p.add_argument("input", help="input file (.docx or .xlsx)")
    p.add_argument(
        "-s", "--sheet", metavar="N|NAME", default=None,
        help="worksheet to convert: 1-based index or sheet name (default: 1)",
    )
    p.add_argument(
        "-o", "--output", metavar="PATH", default=None,
        help="write to file; format guessed from extension "
        "(.csv=comma, .tab/.tsv/.txt=tab, trailing .gz=gzip)",
    )
    p.add_argument(
        "-i", "--images", action="store_true",
        help="(docx) extract images and reference them Obsidian-style "
        "[not yet implemented]",
    )
    p.add_argument("-V", "--version", action="version",
                   version="off2d " + __version__)
    return p


def main(argv=None):
    parser = _build_parser()
    # parse_intermixed_args lets switches appear before or after the filename.
    args = parser.parse_intermixed_args(argv)

    kind = detect.detect_type(args.input)
    if kind is None:
        parser.error(
            "cannot determine type of %r (expected .docx or .xlsx)"
            % args.input
        )

    try:
        if kind == detect.XLSX:
            sheet2txt.convert(args.input, args.sheet, args.output)
        elif kind == detect.DOCX:
            sys.stderr.write(
                "off2d: docx -> markdown is not implemented yet (M2)\n"
            )
            return 2
    except FileNotFoundError:
        parser.error("no such file: %r" % args.input)
    except (ValueError, OSError) as e:
        sys.stderr.write("off2d: %s\n" % e)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
