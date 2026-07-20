import pandas as pd
import config

def clean_col(series: pd.Series) -> pd.Series:
    """Normalize values for all category columns (trim spaces, collapse multiple spaces)."""
    # Keep missing values as <NA> (not the string "nan")
    s = series.astype("string")
    s = s.str.strip().str.replace(r"\s+", " ", regex=True)
    return s

def clean_dataframes(df_1, df_2, df_3):
    """Applies cleaning logic to the dataframes."""
    
    # Drop rows with missing 'Resource Full Name'
    df_1 = df_1.dropna(subset=['Resource Full Name']).copy()
    df_2 = df_2.dropna(subset=['Resource Full Name']).copy()
    df_3 = df_3.dropna(subset=['Resource Full Name']).copy()

    # Clean column names
    df_1.columns = df_1.columns.str.strip()
    df_2.columns = df_2.columns.str.strip()
    df_3.columns = df_3.columns.str.strip()

    # Drop rows where Active is No
    df_1 = df_1[df_1['Active'].str.strip().str.lower() != 'no'].copy()
    df_2 = df_2[df_2['Active'].str.strip().str.lower() != 'no'].copy()
    df_3 = df_3[df_3['Active'].str.strip().str.lower() != 'no'].copy()

    # Normalize values for category columns
    for df in (df_1, df_2, df_3):
        for col in config.CATEGORIES.values():
            if col in df.columns:
                df[col] = clean_col(df[col])
                
    return df_1, df_2, df_3
