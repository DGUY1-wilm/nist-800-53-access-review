#!/usr/bin/env python3
"""Compare output/findings.xlsx with data/answer_key.csv (synthetic data only)."""
import pandas as pd

key = pd.read_csv("data/answer_key.csv")
found = pd.read_excel("output/findings.xlsx", sheet_name="Findings")
found_pairs = set(zip(found["Username"], found["Rule ID"]))
planted = set(zip(key["username"], key["planted_rule"]))
missed, extra = planted - found_pairs, found_pairs - planted
print(f"Planted: {len(planted)}  Caught: {len(planted) - len(missed)}  Missed: {len(missed)}")
print(f"Extra findings (check whether they are legitimate overlaps): {sorted(extra)}")
