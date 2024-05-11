import pandas as pd

def sort_fiber_diameter(df, experiments):
    sorted_df = pd.DataFrame(index=df.index)
    for fiber_diam_col, ax_col, my_col, g_ratio_col in experiments:
        current_df = df[[fiber_diam_col, ax_col, my_col, g_ratio_col]].dropna().sort_values(by=fiber_diam_col)
        sorted_df.loc[sorted_df.index[:len(current_df)], [fiber_diam_col, ax_col, my_col, g_ratio_col]] = current_df.values
    return sorted_df

