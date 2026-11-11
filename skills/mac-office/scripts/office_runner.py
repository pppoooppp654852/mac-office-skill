#!/usr/bin/env python3
"""mac-office runner: headless inspection, native-app execution and PDF rendering.

Three modes:

  --inspect <file>   Read .docx/.xlsx/.pptx. Plain packages are parsed with zipfile +
                     ElementTree (milliseconds, no GUI); OLE2 containers (Word 97-2003 and
                     Microsoft Purview / RMS EncryptedPackage files, i.e. everything this
                     machine's Office writes) are read by the owning app over AppleScript.
  --render <file>    Export .docx/.xlsx/.pptx to PDF with the owning app, then rasterize the
                     PDF and report its text lines (PDFKit via JXA). A .pdf input is
                     rasterized directly.
  --app <app>        Compile and run an AppleScript through osascript with per-app alert
                     suppression, a per-run Group-Container stage, staged assets, timeout
                     watchdog and lazy write-back, reporting structured JSON.

Every Office call runs through ``macoffice.app_session``: serialized per app, with a stopped
app started hidden, a hidden app kept hidden and the user's focus restored if Office took it.

Environment variables exposed to AppleScript:
  system attribute "MAC_OFFICE_TARGET_FILE"   -> staged copy of --target-file
  system attribute "MAC_OFFICE_OUTPUT_FILE"   -> staged path for --output-file
  system attribute "MAC_OFFICE_STAGE_DIR"     -> this run's staging directory
  system attribute "MAC_OFFICE_ORIGINAL_FILE" -> the user's original target path
  system attribute "MAC_OFFICE_ASSET_COUNT"   -> number of --asset files (0 when none)
  system attribute "MAC_OFFICE_ASSET_<n>"     -> staged copy of the nth --asset file
"""

import argparse
import json
import sys

from macoffice import constants
from macoffice.execute import run_execute
from macoffice.inspect import run_inspect
from macoffice.render import DEFAULT_DPI, run_render


def emit(payload):
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if payload.get("success") else 1


def main():
    parser = argparse.ArgumentParser(
        description="Inspect Office files headlessly, render them to PDF, or run AppleScript "
        "against local Office apps."
    )
    parser.add_argument("--inspect", metavar="FILE", help="Headless structure inspection target")
    parser.add_argument("--render", metavar="FILE", help="Render a .docx/.xlsx/.pptx/.pdf to PNG + text lines")
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI, help=f"Render resolution (default {DEFAULT_DPI})")
    parser.add_argument("--output-dir", help="Render output directory (default: new temp dir)")
    parser.add_argument(
        "--max-items",
        type=int,
        default=200,
        help="Max items per collection (default 200)",
    )
    parser.add_argument("--app", choices=sorted(constants.APP_NAMES), help="App to drive: word, excel, powerpoint")
    parser.add_argument("--file", help="AppleScript file to execute (default: read stdin)")
    parser.add_argument("--target-file", help="Absolute path of the existing Office file to edit")
    parser.add_argument("--output-file", help="Absolute path for a newly created or exported file (e.g. .docx, .xlsx, .pptx, .pdf, .csv)")
    parser.add_argument(
        "--asset",
        action="append",
        metavar="PATH",
        help="Extra input file for the script (repeatable); exposed as MAC_OFFICE_ASSET_<n>",
    )
    parser.add_argument("--in-place", action="store_true", help="Edit the target file instead of a _revised copy")
    parser.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="osascript timeout in seconds (default 30, max 180; the inspect fallback gets at least 60)",
    )
    args = parser.parse_args()
    args.timeout = max(5, min(args.timeout, 180))

    if args.asset and not args.app:
        parser.error("--asset requires --app")
    if args.in_place and not args.target_file:
        parser.error("--in-place requires --target-file")
    if (args.output_dir or args.dpi != DEFAULT_DPI) and not args.render:
        parser.error("--output-dir and --dpi require --render")

    if args.inspect:
        return emit(run_inspect(args))
    if args.render:
        return emit(run_render(args))
    if args.app:
        return emit(run_execute(args))
    parser.error("Specify either --inspect FILE, --render FILE or --app <word|excel|powerpoint>")


if __name__ == "__main__":
    sys.exit(main())
