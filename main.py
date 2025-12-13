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
    df_1, df_2, df_3 = data_loader.load_data(excel_path)

    # 3) Clean data
    print("Cleaning data...")
    df_1, df_2, df_3 = data_cleaner.clean_dataframes(df_1, df_2, df_3)
    
    print(f"{config.MONTH_1} records: {len(df_1)}")
    print(f"{config.MONTH_2} records: {len(df_2)}")
    print(f"{config.MONTH_3} records: {len(df_3)}")

    # 4) Plotting
    print("Generating plots...")
    plotter.setup_style()
    plotter.plot_total_on_account(df_1, df_2, df_3)
    plotter.plot_grouped_comparisons(df_1, df_2, df_3)

    # 5) Sanity check (Optional)
    target = "Project Rome"
    print(f"{config.MONTH_1} Project Rome =", (df_1["Project / Programme"] == target).sum())
    print(f"{config.MONTH_2} Project Rome =", (df_2["Project / Programme"] == target).sum())
    print(f"{config.MONTH_3} Project Rome =", (df_3["Project / Programme"] == target).sum())

    print("Done! Check the 'outputs' folder.")

if __name__ == "__main__":
    main()
