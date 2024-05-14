
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import shapiro, t



def analyze_g_ratio(excel_path, output_file, histogram_folder,results_sheet_name="Statistical_analysis"):
    xls = pd.ExcelFile(excel_path)
    results = []
    
    os.makedirs(histogram_folder, exist_ok=True)  

    for sheet_name in xls.sheet_names:
        if "Subset" in sheet_name:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            exp_number = sheet_name.split('_')[0]
            g_ratio_col = f'G-Ratio_{exp_number}'

            if g_ratio_col in df.columns:
                data = df[g_ratio_col].dropna()

                if len(data) < 3:
                    continue

                mean = np.mean(data)
                median = np.median(data)
                std = np.std(data, ddof=1)  
                ci = t.interval(0.95, len(data)-1, loc=mean, scale=std/np.sqrt(len(data)))
                ci_low, ci_up = ci
                shapiro_test = shapiro(data)
                shapiro_stat, shapiro_p_value = shapiro_test
                normality = 'Yes' if shapiro_p_value > 0.05 else 'No'

                results.append({
                    'EXP_name': sheet_name,
                    'Median': median,
                    'Mean': mean,
                    'Standard Deviation': std,
                    '95% CI Lower': ci_low,
                    '95% CI Upper': ci_up,
                    'Shapiro-Stat': shapiro_stat,
                    'Shapiro-p-value': shapiro_p_value,
                    'Normality': normality
                })

    


    results_df = pd.DataFrame(results)
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a' if os.path.exists(output_file) else 'w') as writer:
        results_df.to_excel(writer, index=False, sheet_name=results_sheet_name)
    print(f"Statistical results saved to {output_file} in the sheet named '{results_sheet_name}'")




def plot_g_ratio_frequency(df, g_ratio_col, subset_name, folder_path, display=False):
    if df.empty or len(df[g_ratio_col].dropna()) < 3:
        print(f"Not enough data available for {subset_name}. Skipping...")
        return

    freq, bins = np.histogram(df[g_ratio_col], bins='auto', range=(0, 1))
    freq = freq.astype(float) / freq.sum()  
    plt.figure(figsize=(10, 6))
    bars = plt.bar(bins[:-1], freq, width=np.diff(bins), align='edge', color='blue', edgecolor='black')
    for bar in bars:
        height = bar.get_height()
        if height > 0:
            plt.gca().annotate(f'{height:.1%}', xy=(bar.get_x() + bar.get_width() / 2, height),
                               xytext=(0, 3), textcoords='offset points', ha='center', va='bottom', rotation=90)
    plt.title(f'Frequency Distribution of G Ratios - {subset_name}')
    plt.xlabel('G Ratio Bin')
    plt.ylabel('Frequency (%)')
    plt.ylim(0, max(freq) * 1.2)

    if display:
        plt.show()

    os.makedirs(folder_path, exist_ok=True)
    plot_filename = f"{subset_name.replace(' ', '_')}.png"
    plot_path = os.path.join(folder_path, plot_filename)
    plt.savefig(plot_path)
    plt.close()
    print(f"Saved histogram for {subset_name} at {plot_path}")

def plot_all_g_ratio_frequencies(excel_path, folder_path, display=False):
    xls = pd.ExcelFile(excel_path)

    for sheet_name in xls.sheet_names:
        if "Subset" in sheet_name:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            exp_number = sheet_name.split('_')[0]  
            g_ratio_col = f'G-Ratio_{exp_number}'

            if g_ratio_col in df.columns:
                plot_g_ratio_frequency(df, g_ratio_col, sheet_name, folder_path, display)
            else:
                print(f"Column {g_ratio_col} not found in {sheet_name}. Skipping...")


