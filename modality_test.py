import os
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import gaussian_kde
import diptest


def process_and_plot_kde(df, column_name, sheet_name, folder_name="Modality plots"):

    os.makedirs(folder_name, exist_ok=True)
    data = df[column_name].dropna()

    if len(data) <= 3:
        print(f"Not enough data points for Dip Test in {sheet_name}. Skipping...")
        return

    dip, p_value = diptest.diptest(data.values)
    kde = gaussian_kde(data)
    x_grid = np.linspace(data.min(), data.max(), 1000)
    kde_values = kde.evaluate(x_grid)

    plt.figure(figsize=(10, 6))
    plt.plot(x_grid, kde_values, color='blue', alpha=0.5, linewidth=2)
    data_kde_values = kde.evaluate(data)
    plt.scatter(data, data_kde_values, color='red', s=10, zorder=5)

    plt.title(f'{sheet_name} - KDE with Data Points\nDip Test p-value: {p_value:.4f}', fontsize=10)
    plt.xlabel('Fiber Diameter', fontsize=8)
    plt.ylabel('Density', fontsize=8)
    plt.grid(True)
    plt.tight_layout()
    plt.savefig(os.path.join(folder_name, f"{sheet_name}_KDE_Data_Points.png"))
    plt.close()
   