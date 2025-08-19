```python
# ----- CONFIG -----
SHEET_NAME = "Data"   # <-- change this to your data sheet's name
ASSIGNEE_COL = "H"    # Assignees in column H
ISRC_COL     = "I"    # ISRC / "Not Found" in column I
FIRST_DATA_ROW = 2    # Row after the header
LAST_ROW = 1048576    # Excel max rows for safety
# -------------------

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Helper to pull a full column range from another worksheet
def col_range(sheet, col_letter, r1, r2):
    # xl() lets Python in Excel read a worksheet range by A1 notation
    # Example: xl("Data!H2:H1048576")
    return xl(f"{sheet}!{col_letter}{r1}:{col_letter}{r2}")

# Load columns H and I from the data sheet
assignees_raw = col_range(SHEET_NAME, ASSIGNEE_COL, FIRST_DATA_ROW, LAST_ROW)
isrcs_raw     = col_range(SHEET_NAME, ISRC_COL,     FIRST_DATA_ROW, LAST_ROW)

# Normalize into 1-D Series
a = pd.Series(assignees_raw.squeeze() if hasattr(assignees_raw, "squeeze") else assignees_raw)
i = pd.Series(isrcs_raw.squeeze() if hasattr(isrcs_raw, "squeeze") else isrcs_raw)

# Clean up values
a = a.astype(str).str.strip()
i = i.astype(str).str.strip()

# Bucket blank assignees (so blanks don't create separate NaN bars)
a = a.replace({"": np.nan}).fillna("Unassigned")

# Define completion:
# - "Not Found" (any case) -> NOT complete
# - Valid ISRC: 2 letters + 3 alnum + 7 digits (total 12). Case-insensitive.
is_not_found = i.str.casefold().eq("not found")
is_isrc = i.str.fullmatch(r"[A-Za-z]{2}[A-Za-z0-9]{3}\d{7}", case=False)

complete = is_isrc & ~is_not_found

# Consider a row "present" if either assignee or status has something
mask_nonempty = (a.ne("Unassigned")) | (i.ne(""))

# ===== Chart 1: Per-assignee progress =====
grp = pd.DataFrame({"assignee": a, "complete": complete})
grp = grp[mask_nonempty]

by_assignee = grp.groupby("assignee", dropna=False)["complete"].mean().sort_values(ascending=False)

plt.figure(figsize=(6, 3))
plt.bar(by_assignee.index.astype(str), by_assignee.values)
plt.ylim(0, 1)
plt.ylabel("Completion rate")
plt.title("Per-Assignee ISRC Progress")
plt.xticks(rotation=45, ha="right")
for idx, val in enumerate(by_assignee.values):
    plt.text(idx, val, f"{val:.0%}", ha="center", va="bottom")
plt.tight_layout()

# ===== Chart 2: Overall progress =====
overall = complete[mask_nonempty].mean() if mask_nonempty.any() else 0.0

plt.figure(figsize=(4, 0.8))
plt.barh([0], [overall], height=0.4)
plt.barh([0], [1], height=0.4, left=[0], fill=False)  # outline to 100%
plt.xlim(0, 1)
plt.yticks([])
plt.xlabel(f"{overall:.0%} complete")
plt.title("Overall ISRC Progress")
plt.tight_layout()

# ===== Chart 3: Found vs Not Found (counts) =====
found_count = int((complete & mask_nonempty).sum())
not_found_count = int((~complete & mask_nonempty).sum())

plt.figure(figsize=(4, 3))
cats = ["Found", "Not Found"]
vals = [found_count, not_found_count]
plt.bar(cats, vals)
plt.title("ISRCs: Found vs Not Found")
for idx, val in enumerate(vals):
    plt.text(idx, val, str(val), ha="center", va="bottom")
plt.tight_layout()
```