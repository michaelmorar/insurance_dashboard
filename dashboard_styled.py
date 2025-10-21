import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# -----------------------------
# Style / look & feel
# -----------------------------
face_colour = "#ADD8E6"
plt.style.use("fivethirtyeight")
plt.rc("font", family="Google Sans", size=12)

# Ensure output folder exists BEFORE any savefig
os.makedirs("outputs", exist_ok=True)

# -----------------------------
# 1) Locate the Excel file
# -----------------------------
files = [f for f in os.listdir(".") if f.lower().endswith(".xlsx")]
if len(files) == 1:
    excel_path = f"./{files[0]}"
else:
    # If multiple Excel files exist, specify the filename directly
    excel_path = "~/Downloads/aviva.xlsx"

# -----------------------------
# 2) Read sheets
# -----------------------------
df_oct = pd.read_excel(excel_path, sheet_name="DATA-Oct25")
df_sep = pd.read_excel(excel_path, sheet_name="Data_SEPT25")
df_aug = pd.read_excel(excel_path, sheet_name="Data_AUG25")

df_oct = df_oct.dropna(subset=['Resource Full Name'])
df_sep = df_sep.dropna(subset=['Resource Full Name'])
df_aug = df_aug.dropna(subset=['Resource Full Name'])

print(df_oct)
# -----------------------------
# 3) Clean column names & values
# -----------------------------
df_oct.columns = df_oct.columns.str.strip()
df_sep.columns = df_sep.columns.str.strip()
df_aug.columns = df_aug.columns.str.strip()

# Column mapping
categories = {
    "Project Type": "Project Type",
    "Business Unit": "Business Unit",
    "Capco Role": "Capco Role",
    "Capco Location": "Capco Location City",
    "Project Programme": "Project / Programme",
}

# Normalize values for all category columns (trim spaces, collapse multiple spaces)
def clean_col(series: pd.Series) -> pd.Series:
    # Keep missing values as <NA> (not the string "nan")
    s = series.astype("string")
    s = s.str.strip().str.replace(r"\s+", " ", regex=True)
    return s

for df in (df_oct, df_sep, df_aug):
    for col in categories.values():
        if col in df.columns:
            df[col] = clean_col(df[col])

# -----------------------------
# 4) Plot: Total on Account
# -----------------------------
total_oct = len(df_oct)
total_sep = len(df_sep)
total_aug = len(df_aug)

labels = ["Oct 2025", "Sep 2025", "Aug 2025"]
vals   = [total_oct, total_sep,   total_aug]

fig, ax = plt.subplots(figsize=(6, 4))
ax.set_facecolor(face_colour)
ax.bar(labels, vals, color=["blue", "green", "orange"], edgecolor="k", linewidth=0.8)
ax.set_title("Total Consulting Resources on Account")
ax.set_xlabel("Month")
ax.set_ylabel("Number of Resources")
for xcat, val in zip(labels, vals):
    ax.text(xcat, val + 1, str(val), ha="center", va="bottom")
plt.tight_layout()
plt.savefig("outputs/Total_on_account.png", dpi=200)
plt.close(fig)

# -----------------------------
# 5) Grouped comparisons (with spacing & correct offsets)
# -----------------------------
for label, col in categories.items():
    # Count per month (exclude NA)
    counts_oct = df_oct[col].dropna().value_counts()
    counts_sep = df_sep[col].dropna().value_counts()
    counts_aug = df_aug[col].dropna().value_counts()

    # For Project/Programme, take a stable Top-5 across ALL months combined
    if label == "Project Programme":
        all_counts = counts_oct.add(counts_sep, fill_value=0).add(counts_aug, fill_value=0)
        topN = all_counts.sort_values(ascending=False).head(10).index
        idx = list(topN)
        counts_oct = counts_oct.reindex(idx, fill_value=0)
        counts_sep = counts_sep.reindex(idx, fill_value=0)
        counts_aug = counts_aug.reindex(idx, fill_value=0)
    else:
        # Align to the union of categories for the other dimensions
        idx = sorted(set(counts_oct.index).union(counts_sep.index).union(counts_aug.index))
        counts_oct = counts_oct.reindex(idx, fill_value=0)
        counts_sep = counts_sep.reindex(idx, fill_value=0)
        counts_aug = counts_aug.reindex(idx, fill_value=0)

    # --- spacing controls (tweak to taste) ---
    n = 3               # Sep, Aug, Jul
    bar_width = 0.22    # width of each bar
    inner_gap = 0.08    # gap between bars within a group (x units)
    group_gap = 0.30    # extra gap between groups (x units)
    # -----------------------------------------

    x = np.arange(len(idx)) * (1 + group_gap)  # spread groups across the x-axis
    total_width = n * bar_width + (n - 1) * inner_gap
    start = x - total_width / 2 + bar_width / 2  # left bar center within each group

    oct_x = start
    sep_x = start + (bar_width + inner_gap)
    aug_x = start + 2 * (bar_width + inner_gap)

    # Plot
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.set_facecolor(face_colour)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.yaxis.grid(True, linestyle="--", linewidth=0.5, alpha=0.7)
    ax.set_axisbelow(True)

    ax.bar(oct_x, counts_oct.values, bar_width, label="Oct 2025", color="blue",   edgecolor="k", linewidth=0.8)
    ax.bar(sep_x, counts_sep.values, bar_width, label="Sep 2025", color="green",   edgecolor="k", linewidth=0.8)
    ax.bar(aug_x, counts_aug.values, bar_width, label="Aug 2025", color="orange",  edgecolor="k", linewidth=0.8)

    ax.set_title(f"Resources by {label}: Oct vs Sep vs Aug 2025", fontsize=16)
    ax.set_xlabel(label, fontsize=14)
    ax.set_ylabel("Number of Resources", fontsize=14)
    ax.set_xticks(x)
    ax.set_xticklabels(idx, rotation=45, ha="right")

    # Label bars at their own centers
    for cx, series in [(oct_x, counts_oct.values), (sep_x, counts_sep.values), (aug_x, counts_aug.values)]:
        for xi, yi in zip(cx, series):
            ax.text(xi, yi + 1, str(int(yi)), ha="center", va="bottom", fontsize=9)

    ax.legend()
    ax.margins(x=0.02)
    plt.tight_layout()
    plt.savefig(f"outputs/{label}.png", dpi=200)
    plt.close(fig)

# -----------------------------
# 6) (Optional) sanity check
# -----------------------------
# Uncomment to verify a specific item appears as expected:
target = "Project Rome"
print("OCT Project Rome =", (df_oct["Project / Programme"] == target).sum())
print("SEP Project Rome =", (df_sep["Project / Programme"] == target).sum())
print("AUG Project Rome =", (df_aug["Project / Programme"] == target).sum())
