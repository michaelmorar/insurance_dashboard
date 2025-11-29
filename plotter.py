import os
import numpy as np
import matplotlib.pyplot as plt
import config

def setup_style():
    """Sets up the matplotlib style."""
    plt.style.use(config.STYLE)
    plt.rc("font", family=config.FONT_FAMILY, size=config.FONT_SIZE)

def plot_total_on_account(df_nov, df_oct, df_sep):
    """Plots the Total Consulting Resources on Account."""
    total_nov = len(df_nov)
    total_oct = len(df_oct)
    total_sep = len(df_sep)

    labels = ["Nov 2025", "Oct 2025", "Sep 2025"]
    vals   = [total_nov, total_oct, total_sep]

    fig, ax = plt.subplots(figsize=(6, 4))
    ax.set_facecolor(config.FACE_COLOUR)
    ax.bar(labels, vals, color=["blue", "green", "orange"], edgecolor="k", linewidth=0.8)
    ax.set_title("Total Consulting Resources on Account")
    ax.set_xlabel("Month")
    ax.set_ylabel("Number of Resources")
    for xcat, val in zip(labels, vals):
        ax.text(xcat, val + 1, str(val), ha="center", va="bottom")
    plt.tight_layout()
    
    output_path = os.path.join(config.OUTPUT_DIR, "Total_on_account.png")
    plt.savefig(output_path, dpi=200)
    plt.close(fig)

def plot_grouped_comparisons(df_nov, df_oct, df_sep):
    """Plots grouped comparisons for various categories."""
    for label, col in config.CATEGORIES.items():
        # Count per month (exclude NA)
        counts_nov = df_nov[col].dropna().value_counts()
        counts_oct = df_oct[col].dropna().value_counts()
        counts_sep = df_sep[col].dropna().value_counts()

        # For Project/Programme, take a stable Top-5 across ALL months combined
        if label == "Project Programme":
            all_counts = counts_nov.add(counts_oct, fill_value=0).add(counts_sep, fill_value=0)
            topN = all_counts.sort_values(ascending=False).head(10).index
            idx = list(topN)
            counts_nov = counts_nov.reindex(idx, fill_value=0)
            counts_oct = counts_oct.reindex(idx, fill_value=0)
            counts_sep = counts_sep.reindex(idx, fill_value=0)
        else:
            # Align to the union of categories for the other dimensions
            idx = sorted(set(counts_nov.index).union(counts_oct.index).union(counts_sep.index))
            counts_nov = counts_nov.reindex(idx, fill_value=0)
            counts_oct = counts_oct.reindex(idx, fill_value=0)
            counts_sep = counts_sep.reindex(idx, fill_value=0)

        # --- spacing controls ---
        n = 3               
        bar_width = config.BAR_WIDTH
        inner_gap = config.INNER_GAP
        group_gap = config.GROUP_GAP
        # ------------------------

        x = np.arange(len(idx)) * (1 + group_gap)  # spread groups across the x-axis
        total_width = n * bar_width + (n - 1) * inner_gap
        start = x - total_width / 2 + bar_width / 2  # left bar center within each group

        nov_x = start
        oct_x = start + (bar_width + inner_gap)
        sep_x = start + 2 * (bar_width + inner_gap)

        # Plot
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.set_facecolor(config.FACE_COLOUR)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.yaxis.grid(True, linestyle="--", linewidth=0.5, alpha=0.7)
        ax.set_axisbelow(True)

        ax.bar(nov_x, counts_nov.values, bar_width, label="Nov 2025", color="blue",   edgecolor="k", linewidth=0.8)
        ax.bar(oct_x, counts_oct.values, bar_width, label="Oct 2025", color="green",   edgecolor="k", linewidth=0.8)
        ax.bar(sep_x, counts_sep.values, bar_width, label="Sep 2025", color="orange",  edgecolor="k", linewidth=0.8)

        ax.set_title(f"Resources by {label}: Nov vs Oct vs Sep 2025", fontsize=16)
        ax.set_xlabel(label, fontsize=14)
        ax.set_ylabel("Number of Resources", fontsize=14)
        ax.set_xticks(x)
        ax.set_xticklabels(idx, rotation=45, ha="right")

        # Label bars at their own centers
        for cx, series in [(nov_x, counts_nov.values), (oct_x, counts_oct.values), (sep_x, counts_sep.values)]:
            for xi, yi in zip(cx, series):
                ax.text(xi, yi + 1, str(int(yi)), ha="center", va="bottom", fontsize=9)

        ax.legend()
        ax.margins(x=0.02)
        plt.tight_layout()
        
        output_path = os.path.join(config.OUTPUT_DIR, f"{label}.png")
        plt.savefig(output_path, dpi=200)
        plt.close(fig)
