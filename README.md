# Sheet Handoff Gate

## Desktop quick start

Build or obtain the reviewed `sheet-handoff-gate.pyz` file. On Windows with
Python installed, open PowerShell in its directory and run:

    py sheet-handoff-gate.pyz --gui

On Linux/macOS:

    python3 sheet-handoff-gate.pyz --gui

Choose local inputs, click Run check, or Try example to run a real synthetic
fixture. View report shows the selected findings; Save report writes a new JSON or script-free HTML file
only. The interface never uploads inputs or launches a payment flow. Tkinter
must be available in your Python installation (it is optional on some Linux
distributions). If unavailable or on a headless machine, use the existing CLI.
The CLI needs no GUI dependencies. The GUI is a desktop interface, not a mobile
browser inspector. Free local edition; no paid plan or subscription.


A local, repeatable pre-handoff policy check for 1–20 standard-layout XLSX/XLSM workbooks.
It reports numeric findings by input position, not filenames, cells, sheet names, external URLs
or metadata values. No upload, Office application, formula evaluation, or macro execution.

## Run
Python 3.11+; no runtime dependencies. From a local source checkout:
    python -m sheet_handoff_gate workbook.xlsx another.xlsx
    python -m sheet_handoff_gate --policy examples/strict.json workbook.xlsx
Build a deterministic, exclusive-output local zipapp with the Python standard library:
    python build_zipapp.py /path/to/sheet-handoff-gate.pyz
    python /path/to/sheet-handoff-gate.pyz workbook.xlsx

Exit 0: every file passes the selected signal limits.
Exit 1: one or more files exceed a policy limit.
Exit 2: at least one file could not be inspected, or invocation/policy invalid.
The default limits are zero; define all eight numeric limits in a version-1 JSON policy
to permit intentionally accepted signals. An ordinary external hyperlink counts as an
external relationship; this conservative rule may flag normal workbooks.

## Scope and limits
Checks sheet visibility, external relationship records, known external-link, VBA/macro,
embedding/ActiveX, connection/query table and pivot cache part indicators. Macro declarations
are counted as at least one indicator even without a conventional VBA part.
Counts are selected package signals, not confirmed disclosure or malware counts. This is
not a complete OOXML validator, DLP solution, formula audit or certification of safe content.
Cells, comments, chart caches, document authors and formula text are not inspected.
Does not remove or modify any data. Unconventional workbook paths, unsupported namespaces,
encrypted archives, malformed inspected XML, and over-budget input are uninspected, not passed.
Other unsupported or corrupt workbook content outside the inspected parts may not be detected.
A passing report only attests that these limited checks stayed within the supplied limits.
Do not substitute it for Excel Document Inspector or your organization's security process.

Limits: 8 MiB compressed file, 64 MiB declared expanded size, 1,024 ZIP members,
512 KiB central directory, 1 MiB per inspected XML, 4 MiB total inspected XML,
20,000 XML elements and 32 levels per inspected part, 20 files per invocation.
The tool does not extract members or follow external targets. Review numerical metadata before sharing.
Use an updated Python runtime on a trusted system. No paid service or live billing is enabled.

## Human-readable local report

    python3 sheet-handoff-gate.pyz workbook.xlsx --html handoff-report.html

The command writes a new script-free HTML report and preserves the CLI JSON and exit code.
Choose a new filename outside inspected input folders; reports are not signed approvals.


## Installation and support

Extract the complete download ZIP; Windows uses the provided launcher or `py -3 -I -B sheet-handoff-gate.pyz --gui`. The CLI does not need Tkinter.

Use synthetic or redacted inputs in bug reports. Optional support: https://buymeacoffee.com/eztoolgo (no entitlement or subscription).

## Download and run

Use the versioned download ZIP attached to the corresponding release. Extract the complete archive and open index.html for the local interface and setup steps. No account or subscription is required.
