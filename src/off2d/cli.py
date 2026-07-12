"""off2d command-line entry point."""

import argparse
import os
import sys

from . import __version__, detect, sheet2txt, docx2md


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
        help="(docx) extract embedded images and reference them "
        "Obsidian-style: ![[image.png]]",
    )
    p.add_argument(
        "--image-dir", metavar="DIR", default=None,
        help="(docx, with -i) directory for extracted images "
        "(default: <output-stem>_media next to -o, else <input>_media)",
    )
    p.add_argument(
        "--no-cite-links", dest="cite_links", action="store_false",
        help="(docx) do not turn inline citation markers like [1] into "
        "links to their reference URL (linking is on by default)",
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
            md, images = docx2md.convert(
                args.input, out_path=args.output,
                extract_images=args.images, image_dir=args.image_dir,
                cite_links=args.cite_links,
            )
            if args.output is None:
                sys.stdout.write(md)
                if not md.endswith("\n"):
                    sys.stdout.write("\n")
            if images:
                sys.stderr.write(
                    "off2d: extracted %d image(s) to %s\n"
                    % (len(images), os.path.dirname(images[0]) or ".")
                )
    except FileNotFoundError:
        parser.error("no such file: %r" % args.input)
    except (ValueError, OSError) as e:
        sys.stderr.write("off2d: %s\n" % e)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
