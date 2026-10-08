# -*- coding: utf-8 -*-
"""python -m unittest test_launch_core.py （Windows 以外でも実行可）"""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import launch_core as core  # noqa: E402

handler = {"__file__": os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashrun_handler.pyw"), "__name__": "dashrun_handler"}
with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashrun_handler.pyw"), encoding="utf-8") as f:
    exec(compile(f.read(), "dashrun_handler.pyw", "exec"), handler)

ALWAYS = lambda p: True  # noqa: E731


class ValidateTest(unittest.TestCase):
    def ok(self, p):
        return core.validate_exe_path(p, exists=ALWAYS)[0]

    def test_valid(self):
        self.assertTrue(self.ok(r"C:\Tools\VirtualSplit\VirtualSplit.exe"))
        self.assertTrue(self.ok(r"d:\Program Files\App\app.EXE"))
        self.assertTrue(self.ok('"C:\\Tools\\a.exe"'))

    def test_invalid(self):
        for p in [r"\\server\share\a.exe", r"C:\a.bat", r"C:\a.exe /s", "C:/a.exe", "",
                  r"C:\a.exe\\", "a.exe", r"C:\x\..\..\..\a.cmd", "C:\\a\nb.exe", None]:
            self.assertFalse(self.ok(p), p)

    def test_dotdot_is_normalized(self):
        ok, path = core.validate_exe_path(r"C:\Tools\x\..\a.exe", exists=ALWAYS)
        self.assertTrue(ok)
        self.assertEqual(path, r"C:\Tools\a.exe")

    def test_missing_file(self):
        ok, msg = core.validate_exe_path(r"C:\none.exe", exists=lambda p: False)
        self.assertFalse(ok)


class ApprovalTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        core.APP_DIR = self.tmp.name
        core.APPROVED_FILE = os.path.join(self.tmp.name, "approved.json")
        core.LOG_FILE = os.path.join(self.tmp.name, "dashrun.log")
        self.launched = []

    def tearDown(self):
        self.tmp.cleanup()

    def run_req(self, path, answer):
        asked = []

        def ask(p):
            asked.append(p)
            return answer
        res = core.handle_launch_request(path, ask=ask, launcher=self.launched.append, exists=ALWAYS)
        return res, asked

    def test_first_time_asks_then_remembers(self):
        res, asked = self.run_req(r"C:\Tools\a.exe", True)
        self.assertEqual(res["status"], "started")
        self.assertEqual(len(asked), 1)
        res, asked = self.run_req(r"c:\tools\A.EXE", True)  # 大文字小文字違いも同じ扱い
        self.assertEqual(res["status"], "started")
        self.assertEqual(asked, [])
        self.assertEqual(len(self.launched), 2)

    def test_denied(self):
        res, _ = self.run_req(r"C:\Tools\b.exe", False)
        self.assertEqual(res["status"], "denied")
        self.assertEqual(self.launched, [])
        self.assertFalse(core.is_approved(r"C:\Tools\b.exe"))

    def test_invalid_never_asks(self):
        res, asked = self.run_req(r"\\evil\share\x.exe", True)
        self.assertEqual(res["status"], "invalid")
        self.assertEqual(asked, [])


class UrlTest(unittest.TestCase):
    parse = staticmethod(handler["parse_dashrun_url"])

    def test_ok(self):
        self.assertEqual(self.parse("dashrun://launch?path=C%3A%5CTools%5Ca%20b.exe"), r"C:\Tools\a b.exe")
        self.assertEqual(self.parse("dashrun://launch/?path=C%3A%5Ca.exe"), r"C:\a.exe")

    def test_rejects(self):
        for u in ["http://launch?path=x", "dashrun://run?path=x", "dashrun://launch?path=x&args=y",
                  "dashrun://launch?path=a&path=b", "dashrun://launch", "dashrun://launch/x?path=a",
                  "dashrun://launch?path=a#f"]:
            with self.assertRaises(ValueError, msg=u):
                self.parse(u)


if __name__ == "__main__":
    unittest.main()
