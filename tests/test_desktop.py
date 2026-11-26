import json
import os
import tempfile
import unittest
from pathlib import Path

from sheet_handoff_gate import audit_batch, DEFAULT_POLICY
from sheet_handoff_gate import desktop
from sheet_handoff_gate.desktop_ui import write_new_json


class DesktopAdapterTests(unittest.TestCase):
    def test_real_example(self):
        r=desktop.demo()
        self.assertEqual(r.state,"review")
        self.assertEqual(r.data["summary"],{"pass":1,"fail":1,"uninspected":0})
        self.assertIn("hidden sheets: 1",r.details[0])

    def test_gui_matches_core_and_does_not_echo_inputs(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/"PRIVATE_FILENAME.xlsx"
            desktop.make_example(path)
            original=path.read_bytes()
            r=desktop.run({"files":[path],"policy":""})
            self.assertEqual(r.state,"pass")
            self.assertEqual(r.data,audit_batch([path],DEFAULT_POLICY))
            self.assertNotIn("PRIVATE_FILENAME",json.dumps(r.data))
            self.assertEqual(original,path.read_bytes())

    def test_bad_input_is_not_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/"bad.xlsx";path.write_bytes(b"not an xlsx")
            r=desktop.run({"files":[path],"policy":""})
            self.assertEqual(r.state,"uninspected")

    def test_policy_path_and_batch_limits(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/"hidden.xlsx";desktop.make_example(path,True)
            policy=Path(tmp)/"policy.json"
            policy.write_text(json.dumps({"version":1,"maximum":{**DEFAULT_POLICY,"hidden_sheets":1}}))
            self.assertEqual(desktop.run({"files":[path],"policy":policy}).state,"pass")
            with self.assertRaises(ValueError):desktop.run({"files":[path]*21,"policy":""})

    @unittest.skipUnless(os.name=="posix","POSIX link and FIFO")
    def test_special_file_or_symlink_not_inspected(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);path=root/"real.xlsx";desktop.make_example(path)
            link=root/"link.xlsx";link.symlink_to(path)
            fifo=root/"fifo.xlsx";os.mkfifo(fifo)
            r=desktop.run({"files":[link,fifo],"policy":""})
            self.assertEqual(r.data["summary"]["uninspected"],2)

    def test_export_new_file_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/"report.json"
            report=desktop.demo().data
            write_new_json(p,report)
            original=p.read_bytes()
            self.assertEqual(json.loads(original),report)
            with self.assertRaises(FileExistsError):write_new_json(p,{"changed":True})
            self.assertEqual(p.read_bytes(),original)
