import json

from cellproof.model import Verdict
from cellproof.report import write_report


def test_report_counts_sections_and_dedupes_findings(tmp_path):
    verdicts = [
        Verdict(("S", 1, 1), "PASS", 1.0, 1.0),
        Verdict(("S", 2, 1), "FAIL", 5.0, 4.0),
        Verdict(("S", 3, 1), "UNRESOLVED", 1.0, None, "volatile function NOW"),
        Verdict(("S", 4, 1), "FAIL", 2.0, None),
    ]
    counts = write_report(verdicts, ["S!A9: x", "S!A9: x"], tmp_path, "book.xlsx")
    assert counts == {"PASS": 1, "FAIL": 2, "UNRESOLVED": 1}
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "| S!A2 | 5.0 | 4.0 |" in md
    assert "| S!A4 | 2.0 | (no value) |" in md
    assert "| S!A3 | volatile function NOW |" in md
    assert md.count("S!A9: x") == 1
    data = json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))
    assert data["counts"]["FAIL"] == 2 and data["cells"][1]["cell"] == "S!A2"


def test_empty_sections_are_left_out(tmp_path):
    write_report([Verdict(("S", 1, 1), "PASS", 1.0, 1.0)], [], tmp_path, "b.xlsx")
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "Failed cells" not in md and "Unresolved" not in md and "Findings" not in md


def test_regions_section_lists_the_notes(tmp_path):
    from cellproof.report import write_report
    write_report([], [], tmp_path, "b.xlsx", notes=["Data rows 2-9, columns B -> model region_data_2_9",
                                                    "Data column C rows 2-9 stays in cell mode: reads A1"])
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "## Regions" in md and "region_data_2_9" in md and "stays in cell mode" in md
    assert json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))["regions"][0].startswith("Data rows")


def test_no_regions_section_without_notes(tmp_path):
    from cellproof.report import write_report
    write_report([], [], tmp_path, "b.xlsx")
    assert "## Regions" not in (tmp_path / "REPORT.md").read_text(encoding="utf-8")
