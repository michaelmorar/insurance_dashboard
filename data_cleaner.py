import pandas as pd
import config

def clean_col(series: pd.Series) -> pd.Series:
    """Normalize values for all category columns (trim spaces, collapse multiple spaces)."""
    # Keep missing values as <NA> (not the string "nan")
    s = series.astype("string")
    s = s.str.strip().str.replace(r"\s+", " ", regex=True)
    return s

def clean_dataframes(df_nov, df_oct, df_sep):
    """Applies cleaning logic to the dataframes."""
    
    # Drop rows with missing 'Resource Full Name'
    df_nov = df_nov.dropna(subset=['Resource Full Name']).copy()
    df_oct = df_oct.dropna(subset=['Resource Full Name']).copy()
    df_sep = df_sep.dropna(subset=['Resource Full Name']).copy()

    # Clean column names
    df_nov.columns = df_nov.columns.str.strip()
    df_oct.columns = df_oct.columns.str.strip()
    df_sep.columns = df_sep.columns.str.strip()

    # Normalize values for category columns
    for df in (df_nov, df_oct, df_sep):
        for col in config.CATEGORIES.values():
            if col in df.columns:
                df[col] = clean_col(df[col])
                
    return df_nov, df_oct, df_sep
