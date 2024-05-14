

import pandas as pd
import matplotlib.pyplot as plt
import os
from tkinter import filedialog, messagebox
import numpy as np


def extract_data(file_path):
    

    ctl_data = pd.read_excel(file_path, sheet_name='CTL_Data')
    exp_data = pd.read_excel(file_path, sheet_name='EXP_Data')
    return ctl_data, exp_data

def get_max_fiber_diameters(df):
   
    max_fibers = {}
    for col in df.columns:
        if "Fiber Diameter_Max" in col:
            parts = col.split('_')
            subset_index = int(parts[1].replace('Subset', ''))
            max_fibers[subset_index] = df[col].max()
    return max_fibers








def calculate_means_and_plot(ctl_df, exp_df, max_fibers, ctl_prefix, ctl_group, exp_prefix, exp_group): 
    fig, ax = plt.subplots()  
    datasets = {'CTL': ctl_df, 'EXP': exp_df}
    colors = {'CTL': 'blue', 'EXP': 'red'}
    markers = {'CTL': 'o', 'EXP': 's'}
    grand_means = {'CTL': [], 'EXP': []}
    label_added = {'CTL': False, 'EXP': False}

    # extract max fiber diameters from CTL1 subset columns
    ctl1_max_diameters = []
    for col in ctl_df.columns:
        if "CTL1" in col and "Fiber Diameter_Max" in col:
            subset_index = int(col.split('_')[1].replace('Subset', ''))
            ctl1_max_diameters.append((subset_index, ctl_df[col].max()))

    # Sort by subset index to ensure correct order on x-axis
    ctl1_max_diameters.sort()

    # Use only the max diameter values for x-axis plotting
    x_values = [diameter for _, diameter in ctl1_max_diameters]

    # Iterate over each dataset and each subset to calculate means and plot
    for prefix, df in datasets.items():
        for i, max_diameter in ctl1_max_diameters:  # Now looping through ctl1_max_diameters
            g_ratios = []
            for group_number in range(1, 5):  # For each group
                g_ratio_col = f"{prefix}{group_number}_Subset{i}_G-Ratio"
                if g_ratio_col in df.columns:
                    g_ratios.extend(df[g_ratio_col].dropna().tolist())

            if g_ratios:
                mean_g_ratio = sum(g_ratios) / len(g_ratios)
                std_dev = pd.Series(g_ratios).std()
                grand_means[prefix].append(mean_g_ratio)
                if not label_added[prefix]:
                    ax.errorbar(max_diameter, mean_g_ratio, yerr=std_dev, fmt=markers[prefix],
                                color=colors[prefix], label=prefix, capsize=5)
                    label_added[prefix] = True
                else:
                    ax.errorbar(max_diameter, mean_g_ratio, yerr=std_dev, fmt=markers[prefix],
                                color=colors[prefix], capsize=5)

    # Calculate and display the grand mean for each dataset
    for prefix, means in grand_means.items():
        if means:
            overall_grand_mean = sum(means) / len(means)
            ax.axhline(y=overall_grand_mean, color=colors[prefix], linestyle='-', linewidth=2)
            print(f"{prefix} Grand Mean across all subsets: {overall_grand_mean:.5f}")

    # Set x-axis limits based on the smallest and largest diameters from CTL1
    ax.set_xlim([min(x_values) - 0.1, max(x_values) + 0.2])

    ax.set_xlabel('Max Fiber Diameter (µm) from CTL1')
    ax.set_ylabel('G-Ratio')

    
    ax.legend(loc='upper right', title='Groups')

    plt.xticks(x_values)  
    plt.show()

    return fig


    
