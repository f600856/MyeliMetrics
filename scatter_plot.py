import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import pearsonr, spearmanr

def generate_scatter_plots(df, experiments):
    for fiber_diam_col, ax_col in experiments:
        # Filter out rows with NaN values for the current pair of columns
        valid_data = df[[fiber_diam_col, ax_col]].dropna()

        x = valid_data[fiber_diam_col]
        y = valid_data[ax_col]

        # Create scatter plot
        plt.figure()
        plt.scatter(x, y, label=f'{fiber_diam_col} vs. {ax_col}', s=10) 
        plt.xlabel('Fiber Diameter')
        plt.ylabel('Ax Diameter')
        plt.title(f'Scatter Plot and Correlation for {fiber_diam_col} vs. {ax_col}')

        # Calculate Pearson's and Spearman's correlation coefficients
        pearson_corr, _ = pearsonr(x, y)
        spearman_corr, _ = spearmanr(x, y)
        r_squared = pearson_corr ** 2  

        # Display correlation coefficients on the plot
        plt.legend(title=f'Pearson: {pearson_corr:.3f}, Spearman: {spearman_corr:.3f}, R²: {r_squared:.3f}')

        plt.show()
        plt.close()  
