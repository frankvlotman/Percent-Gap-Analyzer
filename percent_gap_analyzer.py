import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
from tkinter.font import Font
from openpyxl import Workbook
from PIL import Image

# Define the path for the blank icon
icon_path = 'C:\\Users\\Frank\\Desktop\\blank.ico'

# Create a blank (transparent) ICO file if it doesn't exist
def create_blank_ico(path):
    size = (16, 16)  # Size of the icon
    image = Image.new("RGBA", size, (255, 255, 255, 0))  # Transparent image
    image.save(path, format="ICO")

# Create the blank ICO file
create_blank_ico(icon_path)

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
        if row.notna().all():
            # Format numeric values to 2 decimal places
            formatted_row = [f"{value:.2f}" if isinstance(value, (int, float)) else value for value in row]
            tree.insert("", "end", values=formatted_row, tags=('price_row',))
            formatted_data.append(formatted_row)

            # Calculate and insert percentage differences if applicable (only for numeric columns)
            percentage_diffs = calculate_percentage_difference(row[1:].tolist())  # Ignore first column for difference
            percentage_row = ['% Diff'] + percentage_diffs
            tree.insert("", "end", values=percentage_row, tags=('percent_row',))
            formatted_data.append(percentage_row)

    # Apply the tag configuration
    tree.tag_configure('price_row', background='#D3D3D3')  # Light grey background for price rows
    tree.tag_configure('percent_row', background='#F5F5F5', foreground='blue')  # Light grey background for percentage rows

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

# Function to copy the selected row to the clipboard
def copy_selected_row(tree):
    selected_item = tree.selection()
    if not selected_item:
        return
    
    item_values = tree.item(selected_item)["values"]
    if not item_values:
        return
    
    # Copy the row values to the clipboard
    row_str = '\t'.join(str(value) for value in item_values)
    root.clipboard_clear()
    root.clipboard_append(row_str)
    root.update()  # Keeps the clipboard data available even after the application is closed

    print(f"Copied to clipboard: {row_str}")

def open_excel_formula_window():
    # Create a new window for displaying Excel formulas
    formula_window = tk.Toplevel(root)
    formula_window.title("Excel Formula")
    formula_window.iconbitmap(icon_path)  # Set the icon for the formula window

    # Create the text to display
    formula_text = (
        "Value A & Value B:\n"
        "=((New Value - Old Value) / Old Value) * 100\n\n"
        "Old Value & Percent Value:\n"
        "For Increase: =(Old Value * (1 + (Percent Value / 100)))\n"
        "For Decrease: =(Old Value * (1 - (Percent Value / 100)))"
    )

    # Display the text in a label
    label = tk.Label(formula_window, text=formula_text, justify=tk.LEFT, padx=10, pady=10)
    label.pack()

def open_review_window():
    selected_item = tree.selection()
    if not selected_item:
        return
    
    item_values = tree.item(selected_item)["values"]
    if not item_values or '% Diff' in item_values:
        return  # Skip if the selected item is a percentage row

    # Create a new window
    review_window = tk.Toplevel(root)
    review_window.title("Review Price Values")
    review_window.iconbitmap(icon_path)  # Set the icon for the review window

    # Set up the Treeview in the review window
    review_tree = ttk.Treeview(review_window, columns=tree["column"], show="headings")
    review_tree.pack(expand=True, fill=tk.BOTH)

    # Add column headings and center align
    for col in tree["column"]:
        review_tree.heading(col, text=col, anchor="center")
        review_tree.column(col, anchor="center", width=100)

    # Insert the selected row's values into the Treeview
    review_tree.insert("", "end", values=item_values)
    
    # Add entry fields for percentage inputs below the Treeview
    percent_entries = {}
    entry_frame = tk.Frame(review_window)
    entry_frame.pack(fill=tk.X, padx=5, pady=5)
    
    for i, col in enumerate(tree["column"]):
        if i == 0:  # Skip the first column (assuming it's not numeric)
            tk.Label(entry_frame, text=col, anchor="center", width=15).grid(row=0, column=i)
            continue
        
        tk.Label(entry_frame, text=f"{col} (%)", anchor="center", width=15).grid(row=0, column=i)
        entry = tk.Entry(entry_frame, width=10)
        entry.grid(row=1, column=i)
        percent_entries[col] = entry

    # Function to move focus to the next entry or button when Enter is pressed
    def focus_next_widget(event):
        event.widget.tk_focusNext().focus()
        return "break"

    # Store new values for the selected row to be applied to the main window later
    global pending_update_values
    pending_update_values = None

    # Function to apply percentage changes
    def apply_changes():
        new_values = [item_values[0]]  # First column value remains the same (assuming it's a non-numeric column like a label)
        
        # Initialize the previous value with the first numeric value (second column)
        previous_value = float(item_values[1])
        new_values.append(f"{previous_value:.2f}")  # Append the original second column value

        # Iterate over the item values starting from the third column
        for i in range(2, len(item_values)):
            try:
                # Retrieve the percentage entered for this column
                percent = float(percent_entries[tree["column"][i]].get())
                
                # Calculate the new value based on the updated previous column's value
                new_value = previous_value * (1 + percent / 100)
                new_values.append(f"{new_value:.2f}")
                
                # Update the previous value to be the newly calculated value
                previous_value = new_value
            except ValueError:
                new_values.append(item_values[i])  # Keep the original value if input is invalid
                previous_value = float(item_values[i])  # Update previous_value with the original if input is invalid

        # Update the Treeview with the new values in the review window only
        review_tree.item(review_tree.get_children()[0], values=new_values)
        
        # Store the updated values to apply to the main window when "Update to PGA" is clicked
        global pending_update_values
        pending_update_values = (selected_item, new_values)

        # Update the corresponding row in formatted_data
        item_index = tree.index(selected_item)
        formatted_data[item_index] = new_values

    # Function to update the main Treeview with the changes, including updating percentage differences
    def update_to_pga():
        if pending_update_values:
            item_id, new_values = pending_update_values
            
            # Update the main Treeview with the new values
            tree.item(item_id, values=new_values)
            
            # Update the corresponding percentage row in the main Treeview
            item_index = tree.index(item_id)
            percent_row_id = tree.get_children()[item_index + 1]  # Get the ID of the next row (percentage row)
            
            # Calculate new percentage differences
            original_values = [float(value) if value != 'None' else None for value in new_values[1:]]
            percentage_diffs = calculate_percentage_difference(original_values)
            percentage_row = ['% Diff'] + percentage_diffs
            
            # Update the percentage row in the main Treeview
            tree.item(percent_row_id, values=percentage_row)

            # Also update the percentage row in formatted_data
            formatted_data[item_index + 1] = percentage_row

    # Add buttons to apply the changes, update to the main Treeview, copy the row, and show Excel formulas
    apply_button = tk.Button(review_window, text="Apply Changes", command=apply_changes, bg='#d0e8f1')
    apply_button.pack(side=tk.LEFT, padx=5, pady=10)

    update_button = tk.Button(review_window, text="Update to PGA", command=update_to_pga, bg='#d0e8f1')
    update_button.pack(side=tk.LEFT, padx=5, pady=10)

    copy_button = tk.Button(review_window, text="Copy Row", command=lambda: copy_selected_row(review_tree), bg='#d0e8f1')
    copy_button.pack(side=tk.LEFT, padx=5, pady=10)

    formula_button = tk.Button(review_window, text="Excel Formula", command=open_excel_formula_window, bg='#d0e8f1')
    formula_button.pack(side=tk.LEFT, padx=5, pady=10)

    # Bind the Enter key to focus on the next widget
    for entry in percent_entries.values():
        entry.bind("<Return>", focus_next_widget)

    # Bind the Enter key to the Apply Changes button when it's in focus
    apply_button.bind("<Return>", lambda event: apply_changes())

# Set up the main application window
root = tk.Tk()
root.title("Percentage Gap Analyzer")
root.iconbitmap(icon_path)  # Set the icon for the main window

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
browse_button = tk.Button(root, text="Load Excel File", command=browse_file, bg='#d0e8f1')
browse_button.pack(side=tk.LEFT, padx=10, pady=10)

# Download button
download_button = tk.Button(root, text="Download", command=download_data, bg='#d0e8f1')
download_button.pack(side=tk.RIGHT, padx=10, pady=10)

# Review button
review_button = tk.Button(root, text="Review", command=open_review_window, bg='#d0e8f1')
review_button.pack(side=tk.LEFT, padx=10, pady=10)

copy_button_main = tk.Button(root, text="Copy Row", command=lambda: copy_selected_row(tree), bg='#d0e8f1')
copy_button_main.pack(side=tk.LEFT, padx=10, pady=10)

formula_button_main = tk.Button(root, text="Excel Formula", command=open_excel_formula_window, bg='#d0e8f1')
formula_button_main.pack(side=tk.RIGHT, padx=10, pady=10)

# Start the application
root.geometry("800x400")
root.mainloop()
