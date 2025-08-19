```python
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# --- Inputs expected: `assignees`, `isrcs` (bound to H and I columns respectively) ---

# Normalize into Series
a = pd.Series(assignees.squeeze() if hasattr(assignees, "squeeze") else assignees)
i = pd.Series(isrcs.squeeze() if hasattr(isrcs, "squeeze") else isrcs)

# Clean up values
a = a.astype(str).str.strip().replace({"": np.nan})
a = a.fillna("Unassigned")  # bucket blank assignees

s = i.astype(str).str.strip()

# Define validity:
# - "Not Found" (any case) is explicitly NOT complete
# - Valid ISRC pattern: 2 letters + 3 alnum + 7 digits (total 12)
is_not_found = s.str.casefold().eq("not found")
is_isrc = s.str.fullmatch(r"[A-Za-z]{2}[A-Za-z0-9]{3}\d{7}", case=False)

# "Complete" when cell looks like a valid ISRC and is not "Not Found"
complete = is_isrc & ~is_not_found

# ----- Chart 1: Per-assignee progress -----
grp = pd.DataFrame({"assignee": a, "complete": complete})
# Exclude rows that are entirely empty in both columns (optional)
# Keep if either assignee or status is present
mask_nonempty = a.ne("Unassigned") | s.ne("")
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

# ----- Chart 2: Total progress (overall) -----
overall = complete[mask_nonempty].mean() if mask_nonempty.any() else 0.0

plt.figure(figsize=(4, 0.8))
plt.barh([0], [overall], height=0.4)
plt.barh([0], [1], height=0.4, left=[0], fill=False)  # outline to 100%
plt.xlim(0, 1)
plt.yticks([])
plt.xlabel(f"{overall:.0%} complete")
plt.title("Overall ISRC Progress")
plt.tight_layout()

# ----- Chart 3: ISRCs Found vs Not Found (counts) -----
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