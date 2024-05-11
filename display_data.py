
import tkinter as tk
from tkinter import ttk



# Function to display the DataFrame in a new window
def display_dataframe(df, parent_window):
    display_window = tk.Toplevel(parent_window)
    display_window.title("Sorted Data")
    
    tree = ttk.Treeview(display_window)
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor=tk.CENTER)
    
    for index, row in df.iterrows():
        tree.insert("", tk.END, values=list(row))
    
    tree.pack(expand=True, fill='both', padx=10, pady=10)

    