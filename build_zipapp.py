"""Build a reproducible offline archive with an exact member allowlist."""
from pathlib import Path
import io
import sys
import zipfile

ROOT = Path(__file__).resolve().parent
MEMBERS = ("README.md", "SECURITY.md", "LICENSE",
           "sheet_handoff_gate/__init__.py", "sheet_handoff_gate/__main__.py", "sheet_handoff_gate/desktop.py", "sheet_handoff_gate/desktop_ui.py")
ENTRY = b"import sys\nfrom sheet_handoff_gate.__main__ import main\nsys.exit(main())\n"


def build(output):
    output = Path(output)
    if output.suffix != ".pyz":
        raise ValueError("choose a .pyz output")
    data = io.BytesIO()
    with zipfile.ZipFile(data, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for name, content in sorted([("__main__.py", ENTRY)] +
                                    [(name, (ROOT / name).read_bytes()) for name in MEMBERS]):
            member = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            member.create_system = 3
            member.external_attr = 0o100644 << 16
            archive.writestr(member, content, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    with output.open("xb") as target:
        target.write(data.getvalue())


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python build_zipapp.py output.pyz")
    try:
        build(sys.argv[1])
    except (OSError, ValueError):
        raise SystemExit("Choose a new writable .pyz output")
