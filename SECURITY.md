# Security scope
Private MVP. Reads regular files, caps archive and XML parsing, rejects encrypted or ambiguous
ZIP entries, rejects DTD/entity declarations, never extracts archive members or executes macros.
Output contains only counts, input positions, fixed violation codes and supplied policy limits.
No network or persistence is implemented. Not an anti-malware scanner or complete privacy audit.
The original workbook may still contain sensitive content even when the selected policy passes.
Final independent review, name/IP checks and release approval remain pending.

## 2026-11-25 structural review
Unexpected or nested relationship/content-type elements, duplicate case-insensitive
content-type declarations and unrecognized worksheet relationship URIs are uninspected.
The recognized macro-sheet relationship types use Microsoft's documented namespace.
This does not validate every internal relationship target or inspect worksheet payloads.
