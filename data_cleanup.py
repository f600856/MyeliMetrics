import pandas as pd 
import numpy as np  


def data_cleanup(df, experiments):
    try:
        for ax_col, my_col in experiments:
            if ax_col in df.columns and my_col in df.columns:
                df.loc[df[ax_col] < 0.15, [ax_col, my_col]] = np.nan
                df.loc[df[my_col] < 0.03, [ax_col, my_col]] = np.nan
            else:
                return "Error", f"Columns {ax_col} and/or {my_col} not found in DataFrame."
        for col in df.columns:
            non_nan_values = df[col].dropna().reset_index(drop=True)
            df[col] = pd.Series(non_nan_values).reindex_like(df)
        return "Success", "Data cleaned successfully."
    except Exception as e:
        return "Error", f"An error occurred during data cleanup: {e}"




