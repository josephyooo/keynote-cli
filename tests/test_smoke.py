"""Smoke tests for keynote-cli.

These are integration tests that require macOS with Keynote installed.
They run real AppleScript against the bundled my-template.key.
"""
from __future__ import annotations

import json
import subprocess
import sys
import textwrap
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CLI = [sys.executable, str(ROOT / "keynote-cli")]
TEMPLATE = ROOT / "my-template.key"
TMP_DIR = Path("/tmp/keynote-cli-tests")


def _run_cli(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [*CLI, *args],
        text=True,
        capture_output=True,
        timeout=60,
        check=check,
    )


def _write_script(name: str, lines: list[str]) -> Path:
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    path = TMP_DIR / f"{name}.keynote-script"
    path.write_text("\n".join(lines) + "\n")
    return path


def _output_path(name: str) -> Path:
    return TMP_DIR / f"{name}.key"


def setUpModule() -> None:
    if sys.platform != "darwin":
        raise unittest.SkipTest("Keynote tests require macOS")
    if not TEMPLATE.exists():
        raise unittest.SkipTest(f"Template not found: {TEMPLATE}")
    TMP_DIR.mkdir(parents=True, exist_ok=True)


class TestInspect(unittest.TestCase):
    def test_inspect_template(self) -> None:
        result = _run_cli("inspect", str(TEMPLATE))
        data = json.loads(result.stdout)
        self.assertGreater(data["slideCount"], 0)
        self.assertIn("Title", data["masters"])

    def test_inspect_masters(self) -> None:
        result = _run_cli("inspect-masters", str(TEMPLATE))
        data = json.loads(result.stdout)
        self.assertIsInstance(data, list)
        self.assertGreater(len(data), 0)
        names = [m["master"] for m in data]
        self.assertIn("Title", names)
        self.assertIn("Title & Bullets", names)
        for master in data:
            for ti in master["textItems"]:
                self.assertIn("target", ti)
                self.assertFalse(ti.get("hidden", False))


class TestRun(unittest.TestCase):
    def test_add_slide(self) -> None:
        out = _output_path("add-slide")
        script = _write_script("add-slide", [
            f"open {TEMPLATE} --output {out} --force",
            'add-slide --master "Title"',
            'set-text --slide 1 --target defaultTitleItem "Test Title"',
            "save",
        ])
        _run_cli("run", str(script))
        result = _run_cli("inspect", str(out))
        data = json.loads(result.stdout)
        self.assertEqual(data["slideCount"], 2)

    def test_doc_ops_only(self) -> None:
        out = _output_path("doc-ops-only")
        script = _write_script("doc-ops-only", [
            f"open {TEMPLATE} --output {out} --force",
            "duplicate-slide --slide 1",
            "save",
        ])
        _run_cli("run", str(script))
        result = _run_cli("inspect", str(out))
        data = json.loads(result.stdout)
        self.assertEqual(data["slideCount"], 2)

    def test_transition(self) -> None:
        out = _output_path("transition")
        script = _write_script("transition", [
            f"open {TEMPLATE} --output {out} --force",
            'add-slide --master "Title"',
            'set-transition --slide 1 --style dissolve --duration 2',
            'set-transition --slide 2 --style none',
            "save",
        ])
        _run_cli("run", str(script))
        self.assertTrue(out.exists())

    def test_check_template_catches_bad_set_master(self) -> None:
        out = _output_path("bad-set-master")
        script = _write_script("bad-set-master", [
            f"open {TEMPLATE} --output {out} --force",
            'set-master --slide 1 --master "Does Not Exist"',
            "save",
        ])
        result = _run_cli("run", str(script), "--check-template", check=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("missing required master", result.stderr.lower())

    def test_print_applescript(self) -> None:
        script = _write_script("print-as", [
            f"open {TEMPLATE} --output {_output_path('print-as')} --force",
            'add-slide --master "Title"',
            "save",
        ])
        result = _run_cli("run", str(script), "--print-applescript")
        self.assertIn("tell application", result.stdout)
        self.assertIn("make new slide", result.stdout)


class TestExport(unittest.TestCase):
    def test_export_pdf(self) -> None:
        out_pdf = TMP_DIR / "export-test.pdf"
        if out_pdf.exists():
            out_pdf.unlink()
        _run_cli("export", str(TEMPLATE), "--output", str(out_pdf), "--force")
        self.assertTrue(out_pdf.exists())
        self.assertGreater(out_pdf.stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
