from pathlib import Path
from openpyxl import Workbook
from excel_diff.cli import main


def make(path: Path, value):
    wb = Workbook(); wb.active["A1"] = value; wb.save(path)


def test_cli_identical_returns_zero(tmp_path, capsys):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    make(a, 1); make(b, 1)
    assert main([str(a), str(b)]) == 0
    assert "identical" in capsys.readouterr().out


def test_cli_difference_returns_one_json(tmp_path, capsys):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    make(a, 1); make(b, 2)
    assert main([str(a), str(b), "--format", "json"]) == 1
    assert '"identical": false' in capsys.readouterr().out


def test_cli_bad_input_returns_two(capsys):
    assert main(["missing.xlsx", "other.xlsx"]) == 2
    assert "error:" in capsys.readouterr().err


def test_cli_writes_report(tmp_path):
    a, b, out = tmp_path / "a.xlsx", tmp_path / "b.xlsx", tmp_path / "report.json"
    make(a, "a"); make(b, "b")
    assert main([str(a), str(b), "--format", "json", "--output", str(out)]) == 1
    assert '"cell": "A1"' in out.read_text(encoding="utf-8")
