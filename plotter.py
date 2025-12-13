import os
import numpy as np
import matplotlib.pyplot as plt
import config

def setup_style():
    """Sets up the matplotlib style."""
    plt.style.use(config.STYLE)
    plt.rc("font", family=config.FONT_FAMILY, size=config.FONT_SIZE)

def plot_total_on_account(df_1, df_2, df_3):
    """Plots the Total Consulting Resources on Account."""
    total_1 = len(df_1)
    total_2 = len(df_2)
    total_3 = len(df_3)

    labels = [config.MONTH_1, config.MONTH_2, config.MONTH_3]
    vals   = [total_1, total_2, total_3]

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

def plot_grouped_comparisons(df_1, df_2, df_3):
    """Plots grouped comparisons for various categories."""
    for label, col in config.CATEGORIES.items():
        # Count per month (exclude NA)
        counts_1 = df_1[col].dropna().value_counts()
        counts_2 = df_2[col].dropna().value_counts()
        counts_3 = df_3[col].dropna().value_counts()

        # For Project/Programme, take a stable Top-5 across ALL months combined
        if label == "Project Programme":
            all_counts = counts_1.add(counts_2, fill_value=0).add(counts_3, fill_value=0)
            topN = all_counts.sort_values(ascending=False).head(10).index
            idx = list(topN)
            counts_1 = counts_1.reindex(idx, fill_value=0)
            counts_2 = counts_2.reindex(idx, fill_value=0)
            counts_3 = counts_3.reindex(idx, fill_value=0)
        else:
            # Align to the union of categories for the other dimensions
            idx = sorted(set(counts_1.index).union(counts_2.index).union(counts_3.index))
            counts_1 = counts_1.reindex(idx, fill_value=0)
            counts_2 = counts_2.reindex(idx, fill_value=0)
            counts_3 = counts_3.reindex(idx, fill_value=0)

        # --- spacing controls ---
        n = 3               
        bar_width = config.BAR_WIDTH
        inner_gap = config.INNER_GAP
        group_gap = config.GROUP_GAP
        # ------------------------

        x = np.arange(len(idx)) * (1 + group_gap)  # spread groups across the x-axis
        total_width = n * bar_width + (n - 1) * inner_gap
        start = x - total_width / 2 + bar_width / 2  # left bar center within each group

        month_1_x = start
        month_2_x = start + (bar_width + inner_gap)
        month_3_x = start + 2 * (bar_width + inner_gap)

        # Plot
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.set_facecolor(config.FACE_COLOUR)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.yaxis.grid(True, linestyle="--", linewidth=0.5, alpha=0.7)
        ax.set_axisbelow(True)

        ax.bar(month_1_x, counts_1.values, bar_width, label=config.MONTH_1, color="blue",   edgecolor="k", linewidth=0.8)
        ax.bar(month_2_x, counts_2.values, bar_width, label=config.MONTH_2, color="green",   edgecolor="k", linewidth=0.8)
        ax.bar(month_3_x, counts_3.values, bar_width, label=config.MONTH_3, color="orange",  edgecolor="k", linewidth=0.8)

        ax.set_title(f"Resources by {label}: {config.MONTH_1} vs {config.MONTH_2} vs {config.MONTH_3}", fontsize=16)
        ax.set_xlabel(label, fontsize=14)
        ax.set_ylabel("Number of Resources", fontsize=14)
        ax.set_xticks(x)
        ax.set_xticklabels(idx, rotation=45, ha="right")

        # Label bars at their own centers
        for cx, series in [(month_1_x, counts_1.values), (month_2_x, counts_2.values), (month_3_x, counts_3.values)]:
            for xi, yi in zip(cx, series):
                ax.text(xi, yi + 1, str(int(yi)), ha="center", va="bottom", fontsize=9)

        ax.legend()
        ax.margins(x=0.02)
        plt.tight_layout()
        
        output_path = os.path.join(config.OUTPUT_DIR, f"{label}.png")
        plt.savefig(output_path, dpi=200)
        plt.close(fig)
