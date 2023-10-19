import os
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk
from collections import defaultdict


def depth_first_search(node, column_prefix='', parsed_data={}):
    for child in node:
        new_key = f"{column_prefix}_{child.tag}"
        if child.text and child.text.strip():
            parsed_data[new_key] = child.text.strip()
        depth_first_search(child, column_prefix=new_key, parsed_data=parsed_data)
    return parsed_data


def clean_column_names(df):
    new_columns = {}
    for col in df.columns:
        new_col = col.split('_')[-1]
        new_columns[col] = new_col
    df.rename(columns=new_columns, inplace=True)


def handle_duplicate_columns(df):
    col_counter = defaultdict(int)
    new_columns = []
    for col in df.columns:
        col_counter[col] += 1
        if col_counter[col] > 1:
            new_columns.append(f"{col}_{col_counter[col]}")
        else:
            new_columns.append(col)
    df.columns = new_columns


def show_column_selector(columns):
    selected_columns = []
    root = tk.Tk()
    root.title("Select Columns")

    canvas = tk.Canvas(root)
    scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    def submit():
        for col, var in zip(columns, vars):
            if var.get():
                selected_columns.append(col)
        root.destroy()

    vars = []
    for col in columns:
        var = tk.IntVar()
        chk = tk.Checkbutton(scrollable_frame, text=col, variable=var)
        chk.pack(side=tk.TOP, anchor=tk.W)
        vars.append(var)

    btn_submit = tk.Button(root, text="Submit", command=submit)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    btn_submit.pack(side="bottom")

    root.mainloop()
    return selected_columns


# Initialize Tkinter window (it won't be shown)
root = tk.Tk()
root.withdraw()

# Open folder dialog to choose folder
folder_path = filedialog.askdirectory(title="Select Folder Containing XML Files")
if not folder_path:
    print("No folder selected. Exiting.")
    exit()
root.destroy() # Destroy the first Tk instance
# List XML Files
xml_files = [f for f in os.listdir(folder_path) if f.endswith('.xml')]

# Initialize an empty DataFrame
final_df = pd.DataFrame()

# Parse XML Files and populate DataFrame
for xml_file in xml_files:
    file_path = os.path.join(folder_path, xml_file)
    tree = ET.parse(file_path)
    root = tree.getroot()

    parsed_data = depth_first_search(root)
    parsed_data['XML_File'] = xml_file

    aligned_data = pd.DataFrame([parsed_data])
    final_df = pd.concat([final_df, aligned_data], ignore_index=True, sort=False)

# Clean up column names
clean_column_names(final_df)

# Handle duplicate columns
handle_duplicate_columns(final_df)

# Show a dialog to select relevant columns
# relevant_columns = show_column_selector(final_df.columns.tolist())
relevant_columns = ['StationID', 'TotalResult', 'StartDate', 'EndDate', 'StepNumber', 'StepTitle', 'Result', 'Utest', 'StepTitle_2', 'Ureal_2', 'Ireal_2', 'Result_2', 'File', 'ProgramFile', 'StepNumber_3', 'GoodTime', 'StepTitle_3', 'Unom', 'Frequency_2', 'StepNumber_4', 'StepTitle_4', 'StepNumber_5', 'PrintTitle_5', 'Result_5'
]
# Create the filtered DataFrame
filtered_df = final_df[relevant_columns]

# Open folder dialog to choose output folder
output_folder = filedialog.askdirectory(title="Select Folder to Save Excel File")
if not output_folder:
    print("No output folder selected. Saving to current directory.")
    output_folder = '.'

# Create Excel path
excel_path = os.path.join(output_folder, 'aligned_combined_data.xlsx')

# Export to Excel with two sheets
with pd.ExcelWriter(excel_path) as writer:
    final_df.to_excel(writer, sheet_name='Raw Data', index=False)
    filtered_df.to_excel(writer, sheet_name='Filtered Data', index=False)
