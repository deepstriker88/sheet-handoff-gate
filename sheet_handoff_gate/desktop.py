"""Local file-selection adapter for the existing workbook gate."""
from pathlib import Path
import tempfile
import zipfile
import sys
from . import audit_batch, load_policy, DEFAULT_POLICY
from .desktop_ui import Field, Result, launch

TITLE = "Sheet Handoff Gate"
LEAD = "Check a workbook batch for selected hidden content and external-package signals before sharing."
LIMIT = "Signal check only. No cell inspection, macro execution, cleanup, or certification."
MORE = """FREE LOCAL TOOL · $0

Select one to twenty XLSX or XLSM files and optionally a reviewed JSON policy.
The default policy flags every selected hidden-sheet, external relationship,
macro, embedding, connection and pivot-cache indicator above zero.

A passed workbook stayed within these policy limits. It is not proof that the
workbook is safe or free of hidden data. Cell values, formulas, comments,
document authors and chart caches are not inspected. Ordinary external
hyperlinks can be flagged. Unsupported, encrypted or malformed workbooks
are uninspected, never passed.

This tool complements Microsoft Excel Document Inspector; it is not a
replacement for Excel or ExcelJS. It does not modify files. Reports use
file positions and counts rather than filenames, cell contents or URLs.

Limits: 20 workbooks, 8 MiB each, 64 MiB declared expansion, 1,024 ZIP
members; selected bounded XML only. No network, account, telemetry or
paid service. For production use, run your organization's complete checks.
"""
FIELDS = (
    Field("files", "Workbooks", "files", filters=(("Excel workbooks","*.xlsx *.xlsm"),("All files","*")),maximum=20),
    Field("policy", "JSON policy", "file", optional=True, filters=(("JSON","*.json"),("All files","*"))),
)

SML="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL="http://schemas.openxmlformats.org/package/2006/relationships"
DOC="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT="http://schemas.openxmlformats.org/package/2006/content-types"


def make_example(path, hidden=False):
    """Create a small synthetic OOXML workbook, not a copy of user data."""
    entries={
        "[Content_Types].xml":f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        "_rels/.rels":f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{DOC}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml":f'<workbook xmlns="{SML}" xmlns:r="{DOC}"><sheets><sheet name="Example" sheetId="1" r:id="rId1" state="{"hidden" if hidden else "visible"}"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels":f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml":f'<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Example only</t></is></c></row></sheetData></worksheet>',
    }
    if hidden:
        entries["xl/workbook.xml"]=entries["xl/workbook.xml"].replace('state="hidden"','state="visible"').replace("</sheets>",'<sheet name="HiddenExample" sheetId="2" r:id="rId2" state="hidden"/></sheets>')
        entries["xl/_rels/workbook.xml.rels"]=entries["xl/_rels/workbook.xml.rels"].replace("</Relationships>",f'<Relationship Id="rId2" Type="{DOC}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>')
        entries["[Content_Types].xml"]=entries["[Content_Types].xml"].replace("</Types>",'<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        entries["xl/worksheets/sheet2.xml"]=f'<worksheet xmlns="{SML}"><sheetData/></worksheet>'
    with zipfile.ZipFile(path,"x",compression=zipfile.ZIP_DEFLATED) as z:
        for name,raw in sorted(entries.items()):
            info=zipfile.ZipInfo(name,(1980,1,1,0,0,0))
            info.compress_type=zipfile.ZIP_DEFLATED
            z.writestr(info,raw,compresslevel=9)


def describe(data):
    s=data["summary"]
    state="uninspected" if s["uninspected"] else ("review" if s["fail"] else "pass")
    titles={"uninspected":"Some files not inspected","review":"Review workbook signals","pass":"Selected checks passed"}
    details=[]
    for row in sorted(data["files"],key=lambda x:({"uninspected":0,"fail":1,"pass":2}[x["status"]],x["index"])):
        if row["status"]=="uninspected":
            details.append(f"File {row['index']} · Not inspected (unsupported or invalid)")
        else:
            detail=", ".join(f"{key.replace('_',' ')}: {row['counts'][key]}" for key in row["violations"])
            details.append(f"File {row['index']} · "+(detail or "Within selected limits"))
    return Result(state,titles[state],"Numeric-only selected signals; no workbook content is shown.",
                  (("Passed",s["pass"]),("Review",s["fail"]),("Uninspected",s["uninspected"])),
                  tuple(details),data)


def run(values):
    return describe(audit_batch(values["files"],load_policy(values.get("policy") or None)))


def demo():
    with tempfile.TemporaryDirectory(prefix="sheet-handoff-example-") as tmp:
        a,b=Path(tmp)/"clean.xlsx",Path(tmp)/"hidden.xlsx"
        make_example(a);make_example(b,True)
        return run({"files":(a,b),"policy":""})


def main():
    try:
        launch(sys.modules[__name__])
        return 0
    except ImportError:
        print("Desktop UI needs Python Tkinter. Use the CLI without --gui.",file=sys.stderr)
        return 2
    except Exception as e:
        if type(e).__name__=="TclError":
            print("Desktop display unavailable. Use the CLI without --gui.",file=sys.stderr)
            return 2
        raise
