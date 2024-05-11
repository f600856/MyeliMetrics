import pandas as pd
from tkinter import filedialog, messagebox

def extract_and_organize_data(sheets, group_prefix):
    data_frames = []

    # Process each sheet according to its name and expected data structure
    for sheet_name, df in sheets.items():
        parts = sheet_name.split('_')
        if len(parts) == 2 and parts[0].startswith(group_prefix):
            group_number = parts[0].replace(group_prefix, '')
            subset_number = parts[1].replace('Subset', '')

            # Define the column names for extraction
            g_ratio_col_name = f"G-Ratio_{group_prefix}{group_number}"
            fiber_diameter_col_name = f"FiberDiameter_{group_prefix}{group_number}"

            # Extract if columns exist
            if g_ratio_col_name in df.columns and fiber_diameter_col_name in df.columns:
                # Extract the columns and maintain their full length
                extracted_data = df[[g_ratio_col_name, fiber_diameter_col_name]].copy()
                extracted_data.rename(columns={
                    g_ratio_col_name: f"{group_prefix}{group_number}_Subset{subset_number}_G-Ratio",
                    fiber_diameter_col_name: f"{group_prefix}{group_number}_Subset{subset_number}_Fiber Diameter"
                }, inplace=True)

                # Add calculations to the DataFrame safely using loc
                for index in extracted_data.index:
                    extracted_data.loc[index, f"{group_prefix}{group_number}_Subset{subset_number}_G-Ratio_Mean"] = extracted_data[f"{group_prefix}{group_number}_Subset{subset_number}_G-Ratio"].mean()
                    extracted_data.loc[index, f"{group_prefix}{group_number}_Subset{subset_number}_G-Ratio_Std"] = extracted_data[f"{group_prefix}{group_number}_Subset{subset_number}_G-Ratio"].std()
                    extracted_data.loc[index, f"{group_prefix}{group_number}_Subset{subset_number}_Fiber Diameter_Max"] = extracted_data[f"{group_prefix}{group_number}_Subset{subset_number}_Fiber Diameter"].max()

                data_frames.append(extracted_data)
            else:
                print(f"Columns not found in {sheet_name}: {g_ratio_col_name}, {fiber_diameter_col_name}")

    # Concatenate all DataFrame objects horizontally
    if data_frames:
        final_df = pd.concat(data_frames, axis=1)
        return final_df
    else:
        print("No data extracted; please check the sheet and column names.")
        return pd.DataFrame()



def process_data(ctl_path, exp_path, save_path):
    # Assuming the function 'extract_and_organize_data' is already defined as shown previously
    ctl_sheets = pd.read_excel(ctl_path, sheet_name=None)
    exp_sheets = pd.read_excel(exp_path, sheet_name=None)

    ctl_data = extract_and_organize_data(ctl_sheets, 'CTL')
    exp_data = extract_and_organize_data(exp_sheets, 'EXP')

    with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
        ctl_data.to_excel(writer, sheet_name='CTL_Data', index=False)
        exp_data.to_excel(writer, sheet_name='EXP_Data', index=False)

    print("Data has been successfully extracted and saved.")

