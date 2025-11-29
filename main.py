import os
import config
import data_loader
import data_cleaner
import plotter

def main():
    # Ensure output folder exists
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)

    # 1) Locate the Excel file
    excel_path = data_loader.get_excel_path()
    print(f"Using Excel file: {excel_path}")

    # 2) Read sheets
    print("Loading data...")
    df_nov, df_oct, df_sep = data_loader.load_data(excel_path)

    # 3) Clean data
    print("Cleaning data...")
    df_nov, df_oct, df_sep = data_cleaner.clean_dataframes(df_nov, df_oct, df_sep)
    
    print(f"Nov records: {len(df_nov)}")
    print(f"Oct records: {len(df_oct)}")
    print(f"Sep records: {len(df_sep)}")

    # 4) Plotting
    print("Generating plots...")
    plotter.setup_style()
    plotter.plot_total_on_account(df_nov, df_oct, df_sep)
    plotter.plot_grouped_comparisons(df_nov, df_oct, df_sep)

    # 5) Sanity check (Optional)
    target = "Project Rome"
    print("NOV Project Rome =", (df_nov["Project / Programme"] == target).sum())
    print("OCT Project Rome =", (df_oct["Project / Programme"] == target).sum())
    print("SEP Project Rome =", (df_sep["Project / Programme"] == target).sum())

    print("Done! Check the 'outputs' folder.")

if __name__ == "__main__":
    main()
