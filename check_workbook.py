"""Independent check: every machine-owned value on the workbook's data sheets equals the
CSV archive. Columns are located by header text, not position, so a column-mapping bug in
the writer cannot hide itself. Read-only.

    python check_workbook.py <project_dir>
"""
import csv
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

XLSX = "LLM_Token_Expenditure_Index.xlsx"
DATE_KEY = (["date"], ["Date"])
RAMP_KEY = (["date_month", "breakdown", "dimension"], ["Month", "Breakdown", "Dimension"])
SHEETS = [
    ("Daily Index", "llm_token_index.csv", DATE_KEY,
     {"LLM Token": "llm_token", "Open LLM": "open_llm", "Proprietary LLM": "proprietary_llm"}),
    ("GPU Rental", "gpu_rental_index.csv", DATE_KEY,
     {"H100 Neo": "h100_neo", "A100 Neo": "a100_neo", "H200 Neo": "h200_neo", "MI300X Neo": "mi300x_neo",
      "B200 Neo": "b200_neo", "H100 Hyper": "h100_hyper", "A100 Hyper": "a100_hyper"}),
    ("Forward Curve", "forward_curve.csv", (["date", "gpu", "rate_type"], ["Date", "GPU", "Rate"]),
     {f"{t}m": f"t{t}" for t in range(37)}),
    ("Ramp Adoption", "ramp_adoption.csv", RAMP_KEY,
     {"Adoption %": "adoption_rate_pct", "MoM (pp)": "mom_change_pp", "YoY (pp)": "yoy_change_pp"}),
    ("Ramp Spend per Employee", "ramp_spend.csv", RAMP_KEY,
     {"Median $": "median_pepm", "P90 $": "p90_pepm", "P99 $": "p99_pepm",
      "Top 10% median $": "top_10_percent_median_pepm", "Top 1% median $": "top_1_percent_median_pepm",
      "Winsorized wtd $": "p99_winsorized_weighted_pepm"}),
]
ISSUES = ("stale", "blank_on_sheet", "formula_in_value_cell", "missing_rows", "extra_rows", "duplicate_keys")


def _text(v):
    return v.strftime("%Y-%m-%d") if isinstance(v, datetime) else ("" if v is None else str(v).strip())


def check(project_dir):
    project_dir = Path(project_dir)
    wb = openpyxl.load_workbook(project_dir / XLSX)
    report = {}
    for sheet, csv_name, (key_cols, key_labels), value_map in SHEETS:
        r = dict.fromkeys(ISSUES + ("compared", "sheet_only_values", "dod_formulas"), 0)
        r["examples"], r["errors"] = [], []
        report[sheet] = r
        if sheet not in wb.sheetnames:
            r["errors"].append("sheet missing")
            continue
        ws = wb[sheet]
        header = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
        absent = [l for l in key_labels + list(value_map) if l not in header]
        if absent:
            r["errors"].append(f"headers not found: {absent}")
            continue
        dod_cols = [c for l, c in header.items() if l.endswith("DoD %")]

        on_sheet = {}
        for row in ws.iter_rows(min_row=2):
            key = tuple(_text(row[header[l] - 1].value) for l in key_labels)
            if not key[0]:
                continue
            if key in on_sheet:
                r["duplicate_keys"] += 1
            on_sheet[key] = row
            r["dod_formulas"] += sum(1 for c in dod_cols
                                     if isinstance(row[c - 1].value, str) and row[c - 1].value.startswith("="))

        with open(project_dir / "data" / csv_name, newline="", encoding="utf-8") as fh:
            archive = {tuple((rec[k] or "").strip() for k in key_cols): rec for rec in csv.DictReader(fh)}

        r["missing_rows"] = len(archive.keys() - on_sheet.keys())
        r["extra_rows"] = len(on_sheet.keys() - archive.keys())
        for key in sorted(archive.keys() & on_sheet.keys()):
            row, rec = on_sheet[key], archive[key]
            for label, col in value_map.items():
                want = (rec.get(col) or "").strip()
                got = row[header[label] - 1].value
                if isinstance(got, str) and got.startswith("="):
                    r["formula_in_value_cell"] += 1
                    continue
                if not want:
                    r["sheet_only_values"] += got not in (None, "")
                    continue
                r["compared"] += 1
                if got in (None, ""):
                    r["blank_on_sheet"] += 1
                else:
                    try:
                        if abs(float(got) - float(want)) < 1e-9:
                            continue
                    except (TypeError, ValueError):
                        pass
                    r["stale"] += 1
                if len(r["examples"]) < 3:
                    r["examples"].append(f"{' / '.join(key)} {label}: sheet {got} vs archive {want}")
    wb.close()
    return report


def issues(report):
    return sum(sum(r[k] for k in ISSUES) + len(r["errors"]) for r in report.values())


def print_report(report):
    for sheet, r in report.items():
        print(f"  {sheet:24s} compared={r['compared']:<6d} stale={r['stale']:<4d} blank={r['blank_on_sheet']} "
              f"formula={r['formula_in_value_cell']} missing_rows={r['missing_rows']} extra_rows={r['extra_rows']} "
              f"dupes={r['duplicate_keys']} sheet_only={r['sheet_only_values']} dod_formulas={r['dod_formulas']}")
        for line in r["errors"] + r["examples"]:
            print("      " + line)
    n = issues(report)
    print("  RESULT:", "CLEAN" if n == 0 else f"{n} issue(s)")
    return n


if __name__ == "__main__":
    sys.exit(1 if print_report(check(sys.argv[1])) else 0)
