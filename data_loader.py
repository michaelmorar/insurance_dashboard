import os
import pandas as pd
import config

def get_excel_path(base_dir="."):
    """Finds the Excel file in the current directory or defaults to a specific path."""
    files = [f for f in os.listdir(base_dir) if f.lower().endswith(".xlsx")]
    if len(files) == 1:
        return os.path.join(base_dir, files[0])
    else:
        # If multiple Excel files exist, specify the filename directly
        return "~/Downloads/aviva.xlsx"

def load_data(excel_path):
    """Loads the specific sheets from the Excel file."""
    df_nov = pd.read_excel(excel_path, sheet_name=config.SHEET_NOV)
    df_oct = pd.read_excel(excel_path, sheet_name=config.SHEET_OCT)
    df_sep = pd.read_excel(excel_path, sheet_name=config.SHEET_SEP)
    
    return df_nov, df_oct, df_sep
