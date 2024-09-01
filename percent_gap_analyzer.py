import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
from tkinter.font import Font
from openpyxl import Workbook

# Function to load Excel file
def load_excel_data(filepath):
    df = pd.read_excel(filepath)  # Load the Excel file with headers
    return df

# Function to calculate percentage differences in a row
def calculate_percentage_difference(row):
    return [
        f"{(row[i] - row[i - 1]) / row[i - 1] * 100:.1f}%" if i > 0 and pd.notna(row[i]) and pd.notna(row[i - 1]) else None
        for i in range(len(row))
    ]

# Function to populate the Treeview with data
def populate_treeview(tree, df):
    tree.delete(*tree.get_children())  # Clear existing data

    # Add columns dynamically based on the DataFrame
    tree["column"] = list(df.columns)
    tree["show"] = "headings"

    # Add column headings and center align
    for col in tree["column"]:
        tree.heading(col, text=col, anchor="center")

    # Adjust column width and center align all cells
    for col in tree["column"]:
        tree.column(col, anchor="center", width=100)

    # Reset the formatted data list
    global formatted_data
    formatted_data = []

    # Add rows
    for index, row in df.iterrows():
        # Only process rows that have valid data (no NaN or invalid entries in any cell)
        if row.notna().all():
            # Format numeric values to 2 decimal places, and handle percentages
            formatted_row = [
                f"{value:.2f}" if pd.notna(value) and isinstance(value, (int, float)) else value
                for value in row
            ]
            
            # Insert the formatted row into the Treeview with light grey background
            tree.insert("", "end", values=formatted_row, tags=('price_row',))
            formatted_data.append(formatted_row)

            # Calculate and insert percentage differences if applicable (only for numeric columns)
            if df.shape[1] > 2:  # More than two columns (including at least one numeric)
                percentage_diffs = calculate_percentage_difference(row[1:].tolist())  # Ignore first column for difference
                percentage_row = ['% Diff'] + percentage_diffs
                tree.insert("", "end", values=percentage_row)
                formatted_data.append(percentage_row)

    # Apply the tag configuration for the price rows
    tree.tag_configure('price_row', background='#D3D3D3')  # Light grey background

def download_data():
    # Create a workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active

    # Add the column headers
    ws.append(tree["column"])

    # Write the data to the worksheet
    for row in formatted_data:
        ws.append(row)

    # Save the workbook to the specified path
    filepath = "C:/Users/Frank/Desktop/table_data.xlsx"
    wb.save(filepath)
    print(f"File saved to {filepath}")

def browse_file():
    # Reset the Treeview and formatted data before loading a new file
    tree.delete(*tree.get_children())
    global formatted_data
    formatted_data = []

    filepath = filedialog.askopenfilename(initialdir="/mnt/data/", title="Select file", filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
    if filepath:
        df = load_excel_data(filepath)
        populate_treeview(tree, df)

# Set up the main application window
root = tk.Tk()
root.title("Percentage Gap Analyzer")

# Set up the Treeview with a vertical scrollbar
tree_frame = tk.Frame(root)
tree_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

# Add a vertical scrollbar
tree_scroll = tk.Scrollbar(tree_frame)
tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set)
tree.pack(expand=True, fill=tk.BOTH)

# Configure the scrollbar
tree_scroll.config(command=tree.yview)

# Create a bold font for the headings
bold_font = Font(family="Helvetica", size=10, weight="bold")

# Apply the bold font to all headings
style = ttk.Style()
style.configure("Treeview.Heading", font=bold_font)

# Browse button
browse_button = tk.Button(root, text="Load Excel File", command=browse_file)
browse_button.pack(side=tk.LEFT, padx=10, pady=10)

# Download button
download_button = tk.Button(root, text="Download", command=download_data)
download_button.pack(side=tk.RIGHT, padx=10, pady=10)

# Start the application
root.geometry("800x400")
root.mainloop()
