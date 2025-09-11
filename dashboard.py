import pandas as pd
import matplotlib.pyplot as plt
import os



# 1. Locate the Excel file (assumes only one .xlsx in /mnt/data)
files = [f for f in os.listdir('.') if f.lower().endswith('.xlsx')]
if len(files) == 1:
    excel_path = f"./{files[0]}"
else:
    # If multiple Excel files exist, specify the filename directly
    excel_path = "./aviva.xlsx"

# 2. Read August and July sheets
df_aug = pd.read_excel(excel_path, sheet_name='Data_AUG25')
df_jul = pd.read_excel(excel_path, sheet_name='Data_JUL25')

# 3. Clean column names
df_aug.columns = df_aug.columns.str.strip()
df_jul.columns = df_jul.columns.str.strip()

# 4. Define categories and corresponding column names
categories = {
    'Project Type': 'Project Type',
    'Business Unit': 'Business Unit',
    'Capco Role': 'Capco Role',
    'Capco Location': 'Capco Location City',
    'Project/Programme': 'Project / Programme'
}

# 5. Plot Total on Account
total_aug = len(df_aug)
total_jul = len(df_jul)
plt.figure(figsize=(6, 4))
plt.bar(['Aug 2025'], [total_aug], color='green', label='Aug 2025')
plt.bar(['Jul 2025'], [total_jul], color='orange', label='Jul 2025')
plt.title('Total Consulting Resources on Account')
plt.xlabel('Month')
plt.ylabel('Number of Resources')
for idx, val in enumerate([total_aug, total_jul]):
    plt.text(idx, val + 1, str(val), ha='center', va='bottom')
plt.legend()
plt.tight_layout()
plt.show()

# 6. Plot comparisons for each category
for label, col in categories.items():
    # Count values
    counts_aug = df_aug[col].value_counts()
    counts_jul = df_jul[col].value_counts()
    
    # For Project/Programme, show top 10 only
    if label == 'Project/Programme':
        counts_aug = counts_aug.head(10)
        counts_jul = counts_jul.head(10)
    
    # Align indices
    idx = sorted(set(counts_aug.index).union(counts_jul.index))
    counts_aug = counts_aug.reindex(idx, fill_value=0)
    counts_jul = counts_jul.reindex(idx, fill_value=0)
    
    # Plot
    x = range(len(idx))
    width = 0.35
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.bar([i - width/2 for i in x], counts_aug, width, label='Aug 2025', color='green')
    ax.bar([i + width/2 for i in x], counts_jul, width, label='Jul 2025', color='orange')
    
    ax.set_title(f'Resources by {label}: Aug vs Jul 2025')
    ax.set_xlabel(label)
    ax.set_ylabel('Number of Resources')
    ax.set_xticks(x)
    ax.set_xticklabels(idx, rotation=45, ha='right')
    ax.legend()
    
    # Label bars
    for i in x:
        ax.text(i - width/2, counts_aug.iloc[i] + 1, str(int(counts_aug.iloc[i])), 
                ha='center', va='bottom', fontsize=9)
        ax.text(i + width/2, counts_jul.iloc[i] + 1, str(int(counts_jul.iloc[i])), 
                ha='center', va='bottom', fontsize=9)
    
    plt.tight_layout()
    plt.show()
