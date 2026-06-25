import os
import sys
import json
import shutil
import subprocess

import pytest

APP_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", "..")
)
sys.path.insert(0, APP_DIR)
import calc as calc_mod  # noqa: E402

SETTINGS = {"weekday_hours": "8", "saturday_start": "09:00", "saturday_end": "14:00"}


def _cli():
    p = shutil.which("cli-anything-puantaj")
    if p:
        return [p]
    if os.environ.get("CLI_ANYTHING_FORCE_INSTALLED") == "1":
        raise RuntimeError("cli-anything-puantaj PATH'te yok. pip install -e .")
    return [sys.executable, "-m", "cli_anything.puantaj"]


class TestCalc:
    def test_weekday_overtime(self):
        # 2026-05-19 Sali, 08:00-19:00, 60dk mola -> 10s calisma, 2s fazla
        r = calc_mod.calc_day_hours("2026-05-19", "08:00", "19:00", 60, SETTINGS)
        assert r[0] == 10.0
        assert r[1] == 8.0
        assert r[2] == 2.0

    def test_no_overtime(self):
        r = calc_mod.calc_day_hours("2026-05-19", "09:00", "17:00", 0, SETTINGS)
        assert r[0] == 8.0
        assert r[2] == 0.0

    def test_break_deduction(self):
        r = calc_mod.calc_day_hours("2026-05-19", "08:00", "17:00", 90, SETTINGS)
        assert r[0] == pytest.approx(7.5)

    def test_night_hours(self):
        r = calc_mod.calc_day_hours("2026-05-19", "22:00", "06:00", 0, SETTINGS)
        assert r[3] > 0.0

    def test_special_day(self):
        r = calc_mod.calc_day_hours("2026-05-19", "08:00", "16:00", 0,
                                    SETTINGS, is_special=1)
        assert r[1] == 0.0
        assert r[5] == 8.0  # special_normal


class TestCLISubprocess:
    CLI = _cli()

    def _run(self, args):
        return subprocess.run(self.CLI + args, capture_output=True, text=True,
                              encoding="utf-8")

    def test_help(self):
        r = self._run(["--help"])
        assert r.returncode == 0
        assert "puantaj" in r.stdout

    def test_info_json(self):
        r = self._run(["--json", "info"])
        assert r.returncode == 0
        d = json.loads(r.stdout)
        assert "db_path" in d
        assert isinstance(d["employees"], int)

    def test_employee_list_json(self):
        r = self._run(["--json", "employee", "list"])
        assert r.returncode == 0
        assert isinstance(json.loads(r.stdout), list)

    def test_template_list_json(self):
        r = self._run(["--json", "template", "list"])
        assert r.returncode == 0
        assert isinstance(json.loads(r.stdout), list)

    def test_auto_template_calc_uses_department_rule(self):
        r = self._run([
            "--json", "timesheet", "calc",
            "--work-date", "2026-05-20",
            "--start", "09:30",
            "--end", "18:30",
            "--department", "LOJİSTİK",
        ])
        assert r.returncode == 0
        d = json.loads(r.stdout)
        assert d["worked"] == 8.0
        assert d["overtime"] == 0.0

    def test_auto_special_calc(self):
        r = self._run([
            "--json", "timesheet", "calc",
            "--work-date", "2026-05-19",
            "--start", "10:00",
            "--end", "18:00",
            "--department", "STANT",
        ])
        assert r.returncode == 0
        d = json.loads(r.stdout)
        assert d["scheduled"] == 0.0
        assert d["special_normal"] == 8.0

    def test_auto_special_kurban_bayrami_arefe(self):
        r = self._run([
            "--json", "timesheet", "calc",
            "--work-date", "2026-05-26",
            "--start", "09:30",
            "--end", "13:30",
            "--department", "LOJİSTİK",
        ])
        assert r.returncode == 0
        d = json.loads(r.stdout)
        assert d["scheduled"] == 0.0
        assert d["special_normal"] == 3.0

    def test_auto_special_kurban_bayrami(self):
        r = self._run([
            "--json", "timesheet", "calc",
            "--work-date", "2026-05-29",
            "--start", "10:00",
            "--end", "18:00",
            "--department", "STANT",
        ])
        assert r.returncode == 0
        d = json.loads(r.stdout)
        assert d["scheduled"] == 0.0
        assert d["special_normal"] == 8.0
