import pandas as pd

def calculate_g_ratio(df, experiments):
    
    for exp_num, (ax_col, my_col) in enumerate(experiments, start=1):
        if ax_col not in df.columns or my_col not in df.columns:
            print(f"Required columns for CTL{exp_num} are missing. Skipping CTL{exp_num}...")
            continue

        g_ratio_col_name = f'G-Ratio_CTL{exp_num}'
        df[g_ratio_col_name] = (df[ax_col] / (df[ax_col] + df[my_col])).round(2)

    print("G-ratio calculations have been completed.")
    return df


