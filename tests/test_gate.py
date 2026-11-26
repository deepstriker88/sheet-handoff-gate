import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile
import warnings
from sheet_handoff_gate import audit_batch, inspect_workbook, load_policy, DEFAULT_POLICY, InputError, MAX_FILE

SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

def fixture(path, state="visible", extra=None, workbook=None, content_type=None):
    typ = content_type or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    entries = {
        "[Content_Types].xml": f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{typ}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        "_rels/.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{DOC}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": workbook or f'<workbook xmlns="{SML}" xmlns:r="{DOC}"><sheets><sheet name="PRIVATE_SHEET" sheetId="1" r:id="rId1" state="{state}"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>PRIVATE_CELL</t></is></c></row></sheetData></worksheet>',
    }
    entries.update(extra or {})
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name, content in entries.items():
            z.writestr(name, content)

class GateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.a = self.root / "PRIVATE_A.xlsx"
        self.b = self.root / "PRIVATE_B.xlsx"

    def test_clean_selected_signals(self):
        fixture(self.a)
        r = audit_batch([self.a], DEFAULT_POLICY)
        self.assertEqual(r["summary"], {"pass":1,"fail":0,"uninspected":0})
        self.assertEqual(r["files"][0]["counts"]["sheets"], 1)
        self.assertNotIn("PRIVATE", json.dumps(r))

    def test_hidden_and_very_hidden(self):
        fixture(self.a, "hidden")
        fixture(self.b, "veryHidden")
        r = audit_batch([self.a,self.b],DEFAULT_POLICY)
        self.assertEqual(r["summary"]["fail"], 2)
        self.assertEqual(r["files"][0]["violations"], ["hidden_sheets"])
        self.assertEqual(r["files"][1]["violations"], ["very_hidden_sheets"])

    def test_external_relationships_and_known_parts(self):
        fixture(self.a, extra={
            "xl/worksheets/_rels/sheet1.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId2" Type="{DOC}/hyperlink" Target="https://example.invalid/PRIVATE" TargetMode="External"/></Relationships>',
            "xl/externalLinks/externalLink1.xml": "unused fixture",
            "xl/vbaProject.bin": b"fixture",
            "xl/embeddings/object.bin": b"fixture",
            "xl/connections.xml": "fixture",
            "xl/pivotCache/pivotCacheDefinition1.xml": "fixture",
        })
        r = audit_batch([self.a],DEFAULT_POLICY)
        self.assertEqual(set(r["files"][0]["violations"]),set(DEFAULT_POLICY)-{"hidden_sheets","very_hidden_sheets"})
        self.assertNotIn("PRIVATE",json.dumps(r))

    def test_macro_enabled_declaration_without_macro_part_is_flagged(self):
        fixture(self.a,content_type="application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        self.assertEqual(inspect_workbook(self.a)["macro_parts"],1)

    def test_custom_policy(self):
        fixture(self.a,"hidden")
        p = self.root / "policy.json"
        p.write_text(json.dumps({"version":1,"maximum":{**DEFAULT_POLICY,"hidden_sheets":1}}))
        self.assertEqual(audit_batch([self.a],load_policy(p))["summary"]["pass"],1)
        p.write_text('{"version":1,"version":1,"maximum":{}}')
        with self.assertRaises(InputError): load_policy(p)

    def test_policy_requires_exact_nonboolean_limits(self):
        p = self.root / "policy.json"
        for maximum in ({}, {**DEFAULT_POLICY,"hidden_sheets":True}, {**DEFAULT_POLICY,"unknown":0}):
            p.write_text(json.dumps({"version":1,"maximum":maximum}))
            with self.assertRaises(InputError): load_policy(p)

    def test_dtd_rejected(self):
        fixture(self.a,workbook=f'<!DOCTYPE workbook [<!ENTITY x "x">]><workbook xmlns="{SML}"><sheets><sheet name="&x;"/></sheets></workbook>')
        self.assertEqual(audit_batch([self.a],DEFAULT_POLICY)["summary"]["uninspected"],1)

    def test_nonstandard_namespace_and_invalid_state(self):
        fixture(self.a,workbook='<workbook><sheets><sheet/></sheets></workbook>')
        fixture(self.b,state="unknown")
        self.assertEqual(audit_batch([self.a,self.b],DEFAULT_POLICY)["summary"]["uninspected"],2)

    def test_invalid_relationship_rejected(self):
        fixture(self.a,extra={"_rels/.rels": f'<Relationships xmlns="{REL}"><Relationship Id="x" Type="{DOC}/officeDocument" Target="https://example.invalid" TargetMode="External"/></Relationships>'})
        self.assertEqual(audit_batch([self.a],DEFAULT_POLICY)["summary"]["uninspected"],1)

    def test_duplicate_case_and_traversal_members(self):
        fixture(self.a,extra={"XL/WORKBOOK.XML":"duplicate"})
        fixture(self.b,extra={"../private.txt":"bad"})
        self.assertEqual(audit_batch([self.a,self.b],DEFAULT_POLICY)["summary"]["uninspected"],2)

    def test_missing_sheet_relationship_is_not_a_pass(self):
        fixture(self.a, extra={"xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{REL}"/>'})
        self.assertEqual(audit_batch([self.a],DEFAULT_POLICY)["summary"]["uninspected"],1)

    def test_nonexistent_sheet_part_is_not_a_pass(self):
        fixture(self.a, extra={"xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/missing.xml"/></Relationships>'})
        self.assertEqual(audit_batch([self.a],DEFAULT_POLICY)["summary"]["uninspected"],1)

    def test_malformed_relationship_fields_are_uninspected(self):
        for extra in (
            f'<Relationship Id="rId1" Type="{DOC}/hyperlink" Target="https://example.invalid" TargetMode="External"/>',
            f'<Relationship Type="{DOC}/hyperlink" Target="https://example.invalid" TargetMode="External"/>',
            f'<Relationship Id="rId2" Type="{DOC}/hyperlink" TargetMode="External"/>',
        ):
            fixture(self.a, extra={"xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>{extra}</Relationships>'})
            self.assertEqual(audit_batch([self.a], DEFAULT_POLICY)["summary"]["uninspected"], 1)

    def test_nested_or_foreign_relationships_are_uninspected(self):
        valid = f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        for bad in (
            f'<wrapper><Relationship Id="rId2" Type="{DOC}/hyperlink" Target="https://example.invalid" TargetMode="External"/></wrapper>',
            '<Relationship xmlns="urn:unsupported" Id="rId2" Type="urn:link" Target="https://example.invalid" TargetMode="External"/>',
        ):
            with self.subTest(bad=bad):
                fixture(self.a, extra={"xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{REL}">{valid}{bad}</Relationships>'})
                self.assertEqual(audit_batch([self.a], DEFAULT_POLICY)["summary"]["uninspected"], 1)

    def test_foreign_sheet_relationship_type_is_uninspected(self):
        fixture(self.a, extra={"xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="urn:fake/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'})
        self.assertEqual(audit_batch([self.a], DEFAULT_POLICY)["summary"]["uninspected"], 1)

    def test_duplicate_or_malformed_content_types_are_uninspected(self):
        valid = '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        for bad in (
            valid,
            valid.replace("/xl/workbook.xml", "/XL/WORKBOOK.XML"),
            '<Default Extension="xml" ContentType="application/xml"/><Default Extension="XML" ContentType="application/xml"/>',
            '<Override ContentType="application/xml"/>',
            '<Unexpected/>',
        ):
            with self.subTest(bad=bad):
                fixture(self.a, extra={"[Content_Types].xml": f'<Types xmlns="{CT}">{valid}{bad}</Types>'})
                self.assertEqual(audit_batch([self.a], DEFAULT_POLICY)["summary"]["uninspected"], 1)

    def test_microsoft_macro_sheet_relationship_is_flagged(self):
        fixture(self.a, extra={"xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="worksheets/sheet1.xml"/></Relationships>'})
        self.assertEqual(inspect_workbook(self.a)["macro_parts"], 1)

    def test_duplicate_member(self):
        fixture(self.a)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with zipfile.ZipFile(self.a,"a") as z: z.writestr("xl/workbook.xml","duplicate")
        self.assertEqual(audit_batch([self.a],DEFAULT_POLICY)["summary"]["uninspected"],1)

    def test_size_and_xml_bounds(self):
        with self.a.open("wb") as f: f.truncate(MAX_FILE+1)
        fixture(self.b,extra={"xl/large.rels":"x"*(1024*1024+1)})
        self.assertEqual(audit_batch([self.a,self.b],DEFAULT_POLICY)["summary"]["uninspected"],2)

    def test_zip_member_limit(self):
        fixture(self.a,extra={f"xl/parts/p{i}":"x" for i in range(1024)})
        self.assertEqual(audit_batch([self.a],DEFAULT_POLICY)["summary"]["uninspected"],1)

    def test_corrupt_and_missing_file_not_passed(self):
        self.a.write_bytes(b"not a workbook")
        r=audit_batch([self.a,self.b,self.root],DEFAULT_POLICY)
        self.assertEqual(r["summary"]["uninspected"],3)

    def test_batch_limit(self):
        with self.assertRaises(InputError): audit_batch([self.a]*21,DEFAULT_POLICY)

    def test_packaged_zipapp_preserves_exit_codes(self):
        fixture(self.a)
        fixture(self.b,"hidden")
        app = self.root / "tool.pyz"
        subprocess.run([sys.executable,"build_zipapp.py",str(app)],check=True)
        with zipfile.ZipFile(app) as z:
            self.assertEqual(set(n for n in z.namelist() if not n.endswith("/")),
                {"__main__.py","sheet_handoff_gate/__init__.py","sheet_handoff_gate/__main__.py","sheet_handoff_gate/desktop.py","sheet_handoff_gate/desktop_ui.py",
                 "README.md","SECURITY.md","LICENSE"})
        for code, paths in [(0,[self.a]),(1,[self.b]),(2,[self.root/"PRIVATE_MISSING"])]:
            result=subprocess.run([sys.executable,str(app),*map(str,paths)],
                cwd=self.root,capture_output=True,text=True)
            self.assertEqual(result.returncode,code,result.stderr)
            self.assertNotIn("PRIVATE",result.stdout+result.stderr)

    def test_reproducible_zipapp_and_no_overwrite(self):
        one = self.root / "one.pyz"
        two = self.root / "two.pyz"
        for app in (one, two):
            subprocess.run([sys.executable, "build_zipapp.py", str(app)], check=True)
        original = one.read_bytes()
        self.assertEqual(original, two.read_bytes())
        result = subprocess.run([sys.executable, "build_zipapp.py", str(one)], capture_output=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(original, one.read_bytes())

    def test_cli_codes_and_privacy(self):
        fixture(self.a)
        fixture(self.b,"hidden")
        def run(*paths):
            return subprocess.run([sys.executable,"-m","sheet_handoff_gate",*map(str,paths)],capture_output=True,text=True)
        self.assertEqual(run(self.a).returncode,0)
        bad=run(self.a,self.b)
        self.assertEqual(bad.returncode,1)
        self.assertNotIn("PRIVATE",bad.stdout+bad.stderr)
        missing=run(self.root/"PRIVATE_MISSING")
        self.assertEqual(missing.returncode,2)
        self.assertNotIn("PRIVATE",missing.stdout+missing.stderr)

if __name__ == "__main__":
    unittest.main()
