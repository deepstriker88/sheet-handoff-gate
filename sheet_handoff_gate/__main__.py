import argparse
import json
import sys
from sheet_handoff_gate import audit_batch, load_policy, InputError

def main():
    if sys.argv[1:] == ['--gui']:
        from .desktop import main as gui_main
        return gui_main()
    parser = argparse.ArgumentParser(description="Check selected workbook handoff signals locally. Not a malware or privacy certification.", epilog="Use --gui by itself for the optional local desktop interface.")
    parser.add_argument("--html", help="write a new local HTML report; never overwrite")
    parser.add_argument("--policy", help="JSON version-1 policy (all metric limits required)")
    parser.add_argument("files", nargs="+", help="1–20 standard XLSX/XLSM packages, max 8 MiB each")
    try:
        args = parser.parse_args()
        result = audit_batch(args.files, load_policy(args.policy))
        if args.html:
            from . import desktop
            from .desktop_ui import render_report, write_new_bytes
            write_new_bytes(args.html, render_report(desktop,desktop.describe(result)))
        print(json.dumps(result, indent=2))
        return 2 if result["summary"]["uninspected"] else int(bool(result["summary"]["fail"]))
    except (InputError, ValueError, OSError, RecursionError):
        print("Invalid policy or input limit exceeded", file=sys.stderr)
        return 2

if __name__ == "__main__":
    sys.exit(main())
