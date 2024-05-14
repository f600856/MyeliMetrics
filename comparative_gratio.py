


import pandas as pd
import matplotlib.pyplot as plt
import os


def read_data(file_path):
    
    ctl_data = pd.read_excel(file_path, sheet_name='CTL_Data')
    exp_data = pd.read_excel(file_path, sheet_name='EXP_Data')
    return ctl_data, exp_data

def extract_max_fiber_diameters(df):
    """ Extract maximum fiber diameters from the CTL data. """
    max_fibers = {}
    for i in range(1, 7):  # Assuming there are 6 subsets
        for group_number in range(1, 5):  # Assuming there are 3 groups (CTL1, CTL2, CTL3)
            fiber_max = f"CTL{group_number}_Subset{i}_Fiber Diameter_Max"
            if fiber_max in df.columns:
                max_fibers[f"{group_number}_{i}"] = df[fiber_max].dropna().values[0]
    return max_fibers
def create_scatter_plot(df, max_fibers, title_prefix, group_prefix):
    """ Create the scatter plot using specific columns from the CTL or EXP DataFrame using CTL max fiber diameters. """
    fig, ax = plt.subplots()
    colors = ['red', 'green', 'blue', 'orange']  # Colors for group 1, group 2, group 3
    g_ratio_means = []
    data_plotted = False

    # Determine the overall x-axis range from CTL max fiber diameters
    min_diameter = min(max_fibers.values())
    max_diameter = max(max_fibers.values())

    # Loop through the subsets and plot each one
    for i in range(1, 7):  
        for group_number in range(1, 5):  
            g_ratio_mean_key = f"{group_prefix}{group_number}_Subset{i}_G-Ratio_Mean"
            g_ratio_std_key = f"{group_prefix}{group_number}_Subset{i}_G-Ratio_Std"
            max_fiber = max_fibers.get(f"{group_number}_{i}")

            if g_ratio_mean_key in df.columns and g_ratio_std_key in df.columns and max_fiber is not None:
                mean = df[g_ratio_mean_key].dropna().values[0]
                std = df[g_ratio_std_key].dropna().values[0]
                g_ratio_means.append(mean)
                ax.errorbar(max_fiber, mean, yerr=std, fmt='o', capsize=2,
                            color=colors[group_number-1], label=f"{group_prefix}{group_number}" if i == 1 else "")
                data_plotted = True
    ax.set_xlim(min_diameter, max_diameter)

    # Calculate and plot the grand mean line if data was plotted
    if data_plotted:
        grand_mean = sum(g_ratio_means) / len(g_ratio_means) if g_ratio_means else 0
        ax.axhline(y=grand_mean, color='gray', linestyle='-', label=f'Grand Mean: {grand_mean:.2f}')
        ax.legend()

    ax.set_xlabel('Max Fiber Diameter from CTL (µm)')
    ax.set_ylabel('G-Ratio')
    ax.set_title(f'{title_prefix}')
    
    if not data_plotted:
        ax.text(0.5, 0.5, 'No data available', horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)

    return fig
