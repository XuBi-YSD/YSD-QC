#!/usr/bin/env python3
"""
Regenerate (or add to) data/field_map.json by diffing a blank Template
against a filled draft of the same workbook.

Usage:
    python3 build_field_map.py --template Template.xlsx --filled Filled.xlsx --label FormLabel

This merges the new diff into the existing data/field_map.json (creates it
if missing). Run from anywhere; paths are resolved relative to the current
working directory.
"""
import argparse
import datetime
import json
import os
import sys

import openpyxl


def infer_type(val):
    if isinstance(val, (datetime.date, datetime.datetime)):
        return "date"
    if isinstance(val, datetime.time):
        return "time"
    if isinstance(val, (int, float)):
        return "number"
    return "text"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True, help="Path to blank template .xlsx")
    ap.add_argument("--filled", required=True, help="Path to filled draft .xlsx")
    ap.add_argument("--label", required=True, help="Key to store this form under in field_map.json")
    ap.add_argument("--out", default=os.path.join(os.path.dirname(__file__), "..", "data", "field_map.json"))
    args = ap.parse_args()

    wt = openpyxl.load_workbook(args.template, data_only=True)
    wf = openpyxl.load_workbook(args.filled, data_only=True)

    sheets = {}
    for sn in wf.sheetnames:
        if sn not in wt.sheetnames:
            continue
        st, sf = wt[sn], wf[sn]
        maxr = max(st.max_row, sf.max_row)
        maxc = max(st.max_column, sf.max_column)
        fields = []
        for r in range(1, maxr + 1):
            for c in range(1, maxc + 1):
                vt = st.cell(row=r, column=c).value
                vf = sf.cell(row=r, column=c).value
                if vt != vf:
                    coord = sf.cell(row=r, column=c).coordinate
                    fields.append({
                        "cell": coord,
                        "template_value": None if vt in (None, "") else str(vt),
                        "sample_value": None if vf in (None, "") else str(vf),
                        "type": infer_type(vf if vf is not None else vt),
                    })
        if fields:
            sheets[sn] = fields
            print(f"  {sn}: {len(fields)} field(s)")

    out_path = os.path.abspath(args.out)
    existing = {}
    if os.path.exists(out_path):
        with open(out_path, encoding="utf-8") as f:
            existing = json.load(f)

    existing[args.label] = sheets
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    total = sum(len(v) for v in sheets.values())
    print(f"Saved {total} field(s) under '{args.label}' -> {out_path}")


if __name__ == "__main__":
    sys.exit(main())
