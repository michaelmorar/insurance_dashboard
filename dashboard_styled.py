import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np

#Style 
face_colour="#ADD8E6"
plt.style.use('fivethirtyeight')
plt.rc('font', family='Google Sans', size=12)


# 1. Locate the Excel file (assumes only one .xlsx in /mnt/data)
files = [f for f in os.listdir('.') if f.lower().endswith('.xlsx')]
if len(files) == 1:
    excel_path = f"./{files[0]}"
else:
    # If multiple Excel files exist, specify the filename directly
    excel_path = "./aviva.xlsx"

# 2. Read August and July sheets
df_sep = pd.read_excel(excel_path, sheet_name='Data_SEPT25')
df_aug = pd.read_excel(excel_path, sheet_name='Data_AUG25')
df_jul = pd.read_excel(excel_path, sheet_name='Data_JUL25')

# 3. Clean column names
df_sep.columns = df_sep.columns.str.strip()
df_aug.columns = df_aug.columns.str.strip()
df_jul.columns = df_jul.columns.str.strip()

# 4. Define categories and corresponding column names
categories = {
    'Project Type': 'Project Type',
    'Business Unit': 'Business Unit',
    'Capco Role': 'Capco Role',
    'Capco Location': 'Capco Location City',
    'Project_Programme': 'Project / Programme'
}

# 5. Plot Total on Account
total_sep = len(df_sep)
total_aug = len(df_aug)
total_jul = len(df_jul)
fig, ax = plt.subplots(figsize=(6, 4))
ax.set_facecolor(face_colour)
#plt.figure(figsize=(6, 4))
ax.bar(['Sep 2025'], [total_sep], color='blue', label='Sep 2025')
ax.bar(['Aug 2025'], [total_aug], color='green', label='Aug 2025')
ax.bar(['Jul 2025'], [total_jul], color='orange', label='Jul 2025')
ax.set_title('Total Consulting Resources on Account')
ax.set_xlabel('Month')
ax.set_ylabel('Number of Resources')
for idx, val in enumerate([total_sep, total_aug, total_jul]):
    plt.text(idx, val + 1, str(val), ha='center', va='bottom')
plt.legend()
plt.tight_layout()
plt.savefig(f'outputs/Total_on_account.png')
#plt.show()

# 6. Plot comparisons for each category
for label, col in categories.items():
    # Count values
    counts_sep = df_sep[col].value_counts()
    counts_aug = df_aug[col].value_counts()
    counts_jul = df_jul[col].value_counts()
    
    # For Project/Programme, show top 10 only
    if label == 'Project_Programme':
        counts_sep = counts_sep.head(10)
        counts_aug = counts_aug.head(10)
        counts_jul = counts_jul.head(10)
    
    # Align indices
    idx = sorted(set(counts_aug.index).union(counts_jul.index).union(counts_sep.index))
    counts_sep = counts_sep.reindex(idx, fill_value=0)
    counts_aug = counts_aug.reindex(idx, fill_value=0)
    counts_jul = counts_jul.reindex(idx, fill_value=0)
    
    # Plot
    x = np.arange(len(idx))
    width = 0.35
    fig, ax = plt.subplots(figsize=(10, 6))

    sep_x = x - width        # left
    aug_x = x                # center
    jul_x = x + width        # right
    
#    ax.bar([i - width/2 for i in x], counts_sep, width, label='Sep 2025', color='blue', edgecolor='k', linewidth=0.8)
#    ax.bar([i - width/2 for i in x], counts_aug, width, label='Aug 2025', color='green', edgecolor='k', linewidth=0.8)
#    ax.bar([i + width/2 for i in x], counts_jul, width, label='Jul 2025', color='orange', edgecolor='k', linewidth=0.8)
    ax.bar(sep_x, counts_sep, width, label='Sep 2025', color='blue', edgecolor='k', linewidth=0.8)
    ax.bar(aug_x, counts_aug, width, label='Aug 2025', color='green', edgecolor='k', linewidth=0.8)
    ax.bar(jul_x, counts_jul, width, label='Jul 2025', color='orange', edgecolor='k', linewidth=0.8)

    ax.set_xticks(x)
    ax.set_xticklabels(idx, rotation=45, ha='right')

    # label bars at the correct x
    for i in range(len(x)):
        ax.text(sep_x[i], counts_sep.iloc[i] + 1, str(int(counts_sep.iloc[i])), ha='center', va='bottom', fontsize=9)
        ax.text(aug_x[i], counts_aug.iloc[i] + 1, str(int(counts_aug.iloc[i])), ha='center', va='bottom', fontsize=9)
        ax.text(jul_x[i], counts_jul.iloc[i] + 1, str(int(counts_jul.iloc[i])), ha='center', va='bottom', fontsize=9)



    ax.set_facecolor(face_colour)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.yaxis.grid(True, linestyle='--', linewidth=0.5, alpha=0.7)
    ax.set_axisbelow(True)
    
    ax.set_title(f'Resources by {label}: Sep vs Aug vs Jul 2025', fontsize=16)
    ax.set_xlabel(label, fontsize=14)
    ax.set_ylabel('Number of Resources', fontsize=14)
    ax.set_xticks(x)
    ax.set_xticklabels(idx, rotation=45, ha='right')
    ax.legend()
    
    
    plt.tight_layout()
    #plt.show()
    plt.savefig(f'outputs/{label}.png')