import pandas as pd


def process_fiber_data(df, experiments):
    if not isinstance(df, pd.DataFrame):
        raise ValueError("The 'df' parameter should be a pandas DataFrame.")
    
    if not isinstance(experiments, list) or not all(isinstance(pair, tuple) and len(pair) == 2 for pair in experiments):
        raise ValueError("The 'experiments' parameter should be a list of tuples, each with two elements.")

    try:
        # Calculate fiber diameter and add new columns
        for ax_col, my_col in experiments:
            if ax_col not in df.columns or my_col not in df.columns:
                raise ValueError(f"Columns {ax_col} and/or {my_col} are missing from the DataFrame.")
            fiber_diameter_col_name = f'FiberDiameter_{ax_col.split("_")[0]}'
            df[fiber_diameter_col_name] = (df[ax_col] + df[my_col]).round(2)

        # Initialize an empty list to store the desired order of columns
        sorted_columns = []

        # Append the new fiber diameter columns and the original columns to the list
        for ax_col, my_col in experiments:
            fiber_diameter_col_name = f'FiberDiameter_{ax_col.split("_")[0]}'
            sorted_columns.extend([fiber_diameter_col_name, ax_col, my_col])

        #add logic to handle any additional columns in 'df' that are not part of 'experiments'
        remaining_columns = [col for col in df.columns if col not in sorted_columns]
        sorted_columns.extend(remaining_columns)

        # Reorganize 'df' to match the order specified in 'sorted_columns'
        df = df[sorted_columns]

    except Exception as e:
        # Handling other unexpected errors
        raise RuntimeError(f"An error occurred: {str(e)}")

    return df