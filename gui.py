import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from data_cleanup import data_cleanup
from fiber_diameter import process_fiber_data
from g_ratio import calculate_g_ratio
from scatter_plot import generate_scatter_plots
from sorted_fiber_diameter  import sort_fiber_diameter
from display_data import display_dataframe
from PIL import Image, ImageTk
from histogram import plot_all_g_ratio_frequencies,analyze_g_ratio
from ctl1_subset import create_subsets_and_save_ranges,save_ranges,create_subsets_based_on_ranges
import os
from modality_test import process_and_plot_kde
from exp_ctl_gratio_mean_std import process_data
from comparative_gratio import read_data, create_scatter_plot,extract_max_fiber_diameters
from visual_ctl_exp_plot import extract_data, get_max_fiber_diameters, calculate_means_and_plot
import matplotlib.pyplot as plt




# Initialize a global DataFrame
df = None
# Function to load file
def load_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            messagebox.showinfo("Success", "File successfully loaded!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

# Function to handle data cleanup
def cleanup_data_handler():
    global df
    if df is not None:
        experiments = [('CTL1_Ax', 'CTL1_My'), ('CTL2_Ax', 'CTL2_My'), ('CTL3_Ax', 'CTL3_My')]
        status, message = data_cleanup(df, experiments)
        if status == "Success":
            messagebox.showinfo(status, message)
        else:
            messagebox.showerror(status, message)
    else:
        messagebox.showerror("Error", "No data loaded.")

# function to handle fiber data processing
def process_fiber_data_handler():
    global df
    if df is not None:
        experiments = [('CTL1_Ax', 'CTL1_My'), ('CTL2_Ax', 'CTL2_My'), ('CTL3_Ax', 'CTL3_My')]
        df = process_fiber_data(df, experiments)
        messagebox.showinfo("Success", "Fiber data processed successfully.")
    else:
        messagebox.showerror("Error", "No data loaded to process.")
    
def calculate_g_ratio_handler():
    global df
    if df is not None:
        experiments = [('CTL1_Ax', 'CTL1_My'), ('CTL2_Ax', 'CTL2_My'), ('CTL3_Ax', 'CTL3_My')]
        df = calculate_g_ratio(df, experiments)
        messagebox.showinfo("Success", "G-ratio calculated successfully.")
    else:
        messagebox.showerror("Error", "No data loaded to calculate G-ratio.")

def show_scatter_plots_handler():
    global df
    if df is not None:
        experiments = [
            ('FiberDiameter_CTL1', 'CTL1_Ax'),
            ('FiberDiameter_CTL2', 'CTL2_Ax'),
            ('FiberDiameter_CTL3', 'CTL3_Ax'),
            
        ]
        generate_scatter_plots(df, experiments)
    else:
        messagebox.showerror("Error", "No data loaded to generate scatter plots.")


def sort_fiber_diameter_handler():
    global df
    if df is not None:
        experiments = [
            ('FiberDiameter_CTL1', 'CTL1_Ax', 'CTL1_My', 'G-Ratio_CTL1'),
            ('FiberDiameter_CTL2', 'CTL2_Ax', 'CTL2_My', 'G-Ratio_CTL2'),
            ('FiberDiameter_CTL3', 'CTL3_Ax', 'CTL3_My', 'G-Ratio_CTL3'),
            
        ]
        df = sort_fiber_diameter(df, experiments)
        messagebox.showinfo("Success", "Data sorted by 'FiberDiameter' for each experiment.")
    else:
        messagebox.showerror("Error", "No data loaded to sort.")


# Sort and display 
def sort_and_display_handler():
    global df
    if df is not None:
        experiments =  [
            ('FiberDiameter_CTL1', 'CTL1_Ax', 'CTL1_My', 'G-Ratio_CTL1'),
            ('FiberDiameter_CTL2', 'CTL2_Ax', 'CTL2_My', 'G-Ratio_CTL2'),
            ('FiberDiameter_CTL3', 'CTL3_Ax', 'CTL3_My', 'G-Ratio_CTL3'),
            
        ]
        sorted_df = sort_fiber_diameter(df, experiments)
        display_dataframe(sorted_df, root)
    else:
        messagebox.showerror("Error", "No data loaded.")



def save_sorted_data(df):
    if df is not None:
        # Open a file dialog to ask for the save location and file name
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All Files", "*.*")],
            title="Save Sorted Data"
        )
        if file_path:
            try:
                # Use ExcelWriter to save the DataFrame to the chosen path with a custom sheet name
                custom_sheet_name = "Sorted_Fiber_Diameter" 
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name=custom_sheet_name)
                messagebox.showinfo("Success", f"Data saved successfully to {file_path} in the sheet '{custom_sheet_name}'")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")
    else:
        messagebox.showerror("Error", "No data to save.")


def filter_subsets_by_count(df, column, min_count=5):
    return df if df[column].count() >= min_count else None


def process_and_save_subsets_command():
    global df
    if df is not None:
        # Inform the user to save the subsets file
        messagebox.showinfo("Save Subsets", "Please save the subsets file.")
        subsets_file_path = filedialog.asksaveasfilename(title="Save Subsets File", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        # Inform the user to save the statistical analysis result
        messagebox.showinfo("Save Ranges File", "Please save the CTL1 ranges file.")
        ranges_file_path = filedialog.asksaveasfilename(title="Save Ranges File", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        if subsets_file_path and ranges_file_path:
            # Handle subsets and ranges with an Excel writer and save them
            with pd.ExcelWriter(subsets_file_path, engine='openpyxl') as writer:
                
                ctl1_ranges = create_subsets_and_save_ranges(df, 'FiberDiameter_CTL1', writer, 'CTL1')
                create_subsets_based_on_ranges(df[['FiberDiameter_CTL2', 'CTL2_Ax', 'CTL2_My', 'G-Ratio_CTL2']].dropna(), 'FiberDiameter_CTL2', ctl1_ranges, writer, 'CTL2')
                create_subsets_based_on_ranges(df[['FiberDiameter_CTL3', 'CTL3_Ax', 'CTL3_My', 'G-Ratio_CTL3']].dropna(), 'FiberDiameter_CTL3', ctl1_ranges, writer, 'CTL3')
                create_subsets_based_on_ranges(df[['FiberDiameter_CTL4', 'CTL4_Ax', 'CTL4_My', 'G-Ratio_CTL4']].dropna(), 'FiberDiameter_CTL4', ctl1_ranges, writer, 'CTL4')

            # Handle ranges in a separate file
            if ctl1_ranges:
                save_ranges(ctl1_ranges, ranges_file_path, 'CTL1')

            messagebox.showinfo("Success", "Data subsets and ranges have been saved successfully.",
                                detail=f"Subsets saved in: {subsets_file_path}\nRanges saved in: {ranges_file_path}")
        else:
            messagebox.showerror("Error", "No file specified for saving. Please try again.")
    else:
        messagebox.showerror("Error", "No data loaded. Please load data before processing.")



def analyze_g_ratio_handler():
    messagebox.showinfo("Select File", "Please select the subsets file for G-Ratio analysis.")
    input_file_path = filedialog.askopenfilename(title="Open Excel file for analysis", filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")])
    
    if not input_file_path:
        messagebox.showerror("Error", "No input file selected. Please select a file to continue.")
        return

    messagebox.showinfo("Save Results", "Please select or create a file and then save the statistical analysis results.")
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save G-Ratio analysis results as")
    
    if not output_file_path:
        messagebox.showerror("Error", "No output file selected. Please select a file to save the results.")
        return

    # Assuming the code directory as the default folder for histograms
    code_directory = os.path.dirname(os.path.realpath(__file__))
    histogram_folder_path = os.path.join(code_directory, "Histograms")
    os.makedirs(histogram_folder_path, exist_ok=True)  # Ensure the directory exists

    try:
        if os.path.exists(output_file_path):
            # The file is removed if it exists to prevent appending
            os.remove(output_file_path)  

        analyze_g_ratio(input_file_path, output_file_path, histogram_folder_path, results_sheet_name="Statistical_analysis")
        plot_all_g_ratio_frequencies(input_file_path, histogram_folder_path, display=False)
        messagebox.showinfo("Success", "G-Ratio analysis and histograms saved successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    







def modality_test_handler():
   
    messagebox.showinfo("Open File", "Please open subsets file for the Modality Check.")
    
    # After the user clicks 'OK', proceed to open the file dialog
    file_path = filedialog.askopenfilename(title="Open File for Modality Check", filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    if not file_path:
        messagebox.showerror("Error", "No file has been loaded. Please load a file first.")
        return

    experiments = ['CTL1', 'CTL2', 'CTL3']  
    num_subsets = 6  
    success = True

    for experiment in experiments:
        xls = pd.ExcelFile(file_path)
        for i in range(1, num_subsets + 1):
            sheet_name = f'{experiment}_Subset{i}'
            if sheet_name in xls.sheet_names:
                try:
                    df_subset = pd.read_excel(xls, sheet_name=sheet_name)
                    fiber_diameter_column = f'FiberDiameter_{experiment}'
                    if fiber_diameter_column in df_subset.columns:
                       process_and_plot_kde(df_subset, fiber_diameter_column, sheet_name)
                    else:
                        print(f"Column {fiber_diameter_column} not found in {sheet_name}.")
                        success = False
                except Exception as e:
                    print(f"An unexpected error occurred while loading {sheet_name}: {e}")
                    success = False
            else:
                print(f"Sheet {sheet_name} does not exist in the file.")
                success = False

    if success:
        messagebox.showinfo("Success", "All subsets have been processed and plots are saved.")
    else:
        messagebox.showwarning("Partial Success", "Some subsets could not be processed. Please check the console for more information.")



def grand_mean_g_ratio_handler():
    
    messagebox.showinfo("Select Files", "Please select the CTL subset and EXP subset Excel files.")
    
    ctl_path = filedialog.askopenfilename(title="Select CTL Excel File", filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
    exp_path = filedialog.askopenfilename(title="Select EXP Excel File", filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
    if not ctl_path or not exp_path:
        messagebox.showerror("Error", "You need to select both CTL and EXP files to continue.")
        return

    save_path = filedialog.asksaveasfilename(title="Save Output", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
    if not save_path:
        messagebox.showerror("Error", "You need to select a file name to save the combined data.")
        return

    process_data(ctl_path, exp_path, save_path)
    
    # Assuming data has been saved and now calculating grand means
    ctl_data = pd.read_excel(save_path, sheet_name='CTL_Data')
    exp_data = pd.read_excel(save_path, sheet_name='EXP_Data')
    
    # Compute grand means for G-Ratio
    ctl_grand_mean = ctl_data.filter(like='G-Ratio').mean().mean()
    exp_grand_mean = exp_data.filter(like='G-Ratio').mean().mean()
    
   # messagebox.showinfo("Grand Mean Results", f"CTL Grand Mean G-Ratio: {ctl_grand_mean:.2f}\nEXP Grand Mean G-Ratio: {exp_grand_mean:.2f}")
    messagebox.showinfo("Success", "Data combined and saved successfully at: " + save_path)




def comparative_gratio_plot():
    

    messagebox.showinfo("Select File", "Please select the combined G-Ratio file for analysis.")
    file_path = filedialog.askopenfilename(title="Open Combined G-Ratio File", filetypes=[("Excel Files", "*.xlsx"), ("Excel Files", "*.xls")])
    
    if not file_path:
        messagebox.showerror("Error", "No file was selected.")
        return

    try:
        ctl_data, exp_data = read_data(file_path)
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return

    try:
        max_fibers_ctl = extract_max_fiber_diameters(ctl_data)
        fig_ctl = create_scatter_plot(ctl_data, max_fibers_ctl, "CTL", "CTL")
        fig_exp = create_scatter_plot(exp_data, max_fibers_ctl, "EXP", "EXP")

        # Save plots in the 'Plots' directory within the code directory
        code_directory = os.path.dirname(os.path.realpath(__file__))
        plot_directory = os.path.join(code_directory, "Plots")
        os.makedirs(plot_directory, exist_ok=True)

        fig_ctl.savefig(os.path.join(plot_directory, 'CTL_plot.png'))
        fig_exp.savefig(os.path.join(plot_directory, 'EXP_plot.png'))

        messagebox.showinfo("Success", "Plots have been saved successfully in the 'Plots' directory.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        plt.close('all')  
         




def Visualize_CTL_EXP_plot():
   

    messagebox.showinfo("Select File", "Please select the combined G-Ratio file for analysis.")
    file_path = filedialog.askopenfilename(title="Open Combined G-Ratio File", filetypes=[("Excel Files", "*.xlsx"), ("Excel Files", "*.xls")])
    
    if not file_path:
        messagebox.showerror("Error", "No file was selected.")
        return

    try:
        ctl_data, exp_data = extract_data(file_path)
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return

    try:
        max_fibers_ctl = get_max_fiber_diameters(ctl_data)
        fig_ctl_exp = calculate_means_and_plot(ctl_data, exp_data, max_fibers_ctl, "CTL", "CTL", "EXP", "EXP")
        # Save plots in the 'Plots' directory within the code directory
        code_directory = os.path.dirname(os.path.realpath(__file__))
        plot_directory = os.path.join(code_directory, "Plots")
        os.makedirs(plot_directory, exist_ok=True)
        
        fig_ctl_exp.savefig(os.path.join(plot_directory, 'CTL_EXP_plot.png'))
        messagebox.showinfo("Success", "Plots have been saved successfully in the 'Plots' directory.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        plt.close('all')  
        



# Create the GUI window
root = tk.Tk()
root.title("MyeliMetrics: A Comprehensive Tool for G-Ratio Calculation and Visualization")
# Optionally set a minimum window size
root.minsize(1375, 1200)

# Load the logo image
image_path = "logo01.png"  
image = Image.open(image_path)
input_path, output_path = "", ""

# Create a Canvas that will fill the entire window
canvas = tk.Canvas(root)
canvas.pack(fill="both", expand=True)

# Function to resize and set background image
def set_background():
    # Get the dimensions of the window
    window_width = root.winfo_width()
    window_height = root.winfo_height()
    if window_width <= 1 or window_height <= 1:
        # If window dimensions are not yet set, wait for 100ms and try again
        root.after(100, set_background)
        return

    # Resize the image to match the window size
    resized_logo_image = image.resize((window_width, window_height), Image.Resampling.LANCZOS)
    logo_photo = ImageTk.PhotoImage(resized_logo_image)

    # Use the Canvas to add the background image
    canvas.create_image(0, 0, image=logo_photo, anchor="nw")
    
    # Keep a reference to the image object
    canvas.image = logo_photo

# Call the function to set the background after the event loop starts
root.after(100, set_background)

# Set a common button size
button_width = 17
button_height = 1
button_color = "white"
text_color = "black"

# Create and place the buttons on the Canvas
load_button = tk.Button(root, text="Upload File", command=load_file, width=button_width, height=button_height, bg=button_color, fg=text_color)
cleanup_button = tk.Button(root, text="Clean Up Data", command=cleanup_data_handler, width=button_width, height=button_height, bg=button_color, fg=text_color)
process_fiber_button = tk.Button(root, text="Calculate Fiber Data", command=process_fiber_data_handler, width=button_width, height=button_height, bg=button_color, fg=text_color)
calculate_g_ratio_button = tk.Button(root, text="Calculate G-Ratio", command=calculate_g_ratio_handler, width=button_width, height=button_height, bg=button_color, fg=text_color)
show_scatter_plots_button = tk.Button(root, text="Show Scatter Plots", command=show_scatter_plots_handler, width=button_width, height=button_height, bg=button_color, fg=text_color)
sort_fiber_diameter_button = tk.Button(root, text="Sort Fiber Diameter", command=sort_fiber_diameter_handler, width=button_width, height=button_height, bg=button_color, fg=text_color)
sort_display_button = tk.Button(root, text="Display Data", command=sort_and_display_handler, width=button_width, height=button_height, bg=button_color, fg=text_color)
Save_Sorted_Data_button = tk.Button(root, text="Save Sorted Data", command=lambda: save_sorted_data(df), width=button_width, height=button_height, bg=button_color, fg=text_color)
Save_Sorted_Data_button.pack()
process_subsets_button = tk.Button(root, text="Create and Save Subsets", command=process_and_save_subsets_command, width=button_width, height=button_height, bg=button_color, fg=text_color)
analyze_g_ratio_button = tk.Button(root, text="Analyze G-Ratio", command=analyze_g_ratio_handler,width=button_width, height=button_height,bg=button_color, fg=text_color) 
analyze_g_ratio_button.pack(expand=True)
Modality_Test_button = tk.Button(root, text="Check Modality", command=modality_test_handler, width=button_width, height=button_height, bg=button_color, fg=text_color) 
Modality_Test_button.pack(expand=True)

# Add a button that will open the file dialog and run the analysis
combined_button = tk.Button(root, text="Grand Mean G-Ratio", command=grand_mean_g_ratio_handler,width=button_width, height=button_height, bg=button_color, fg=text_color)
combined_button.pack(expand=True)


comparative_button = tk.Button(root, text="comparative G-Ratio",  command=comparative_gratio_plot, width=button_width, height=button_height, bg=button_color, fg=text_color)
comparative_button.pack(expand=True)

final_plot_button= tk.Button(root, text="Show Combined Plots",  command=Visualize_CTL_EXP_plot, width=button_width, height=button_height, bg=button_color, fg=text_color)
final_plot_button.pack(expand=True)

# Bind this function to a button in GUI



# Place the buttons on the Canvas
button_positions = [(30, load_button), (76, cleanup_button), (122, process_fiber_button), 
                    (168, calculate_g_ratio_button), (214, show_scatter_plots_button), 
                    (260, sort_fiber_diameter_button), (306, sort_display_button), 
                    (353, Save_Sorted_Data_button),(399, process_subsets_button),(445,analyze_g_ratio_button),(491,Modality_Test_button),(537,combined_button),(583, comparative_button),(630,final_plot_button)]
                    
for y, button in button_positions:
    canvas.create_window(100, y, window=button)

input_label = tk.Label(root)
input_label.pack()


output_label = tk.Label(root)
output_label.pack()


# Start the GUI event loop
root.mainloop()


