import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import xlsxwriter

# Variables to store the loaded dataframes
old_file = None
new_file = None

def load_file(file_var_name):
    global old_file, new_file
    file_path = filedialog.askopenfilename(
        title="Select a file",
        filetypes=[("Excel files", ".xlsx"), ("CSV files", ".csv")],  # Specify supported file types
        initialdir="/",  # Set initial directory
        multiple=False  # Allow only one file selection
    )
    if file_path:
        try:
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='latin1')
            else:
                raise ValueError("Unsupported file type")
            if file_var_name == 'old_file':
                old_file = df
                messagebox.showinfo("Success", "Old file loaded successfully")
            elif file_var_name == 'new_file':
                new_file = df
                messagebox.showinfo("Success", "New file loaded successfully")
                update_comboboxes()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while loading the file: {e}")
    else:
        messagebox.showinfo("Error", "No file selected")

# Functions to load the old and new files
def load_old_file():
    load_file('old_file')

def load_new_file():
    load_file('new_file')

# Function to update the combobox and listbox with the column names of the loaded files
def update_comboboxes():
    if old_file is not None and new_file is not None:
        common_columns = list(set(old_file.columns) & set(new_file.columns))
        primary_primarykey_column_combobox['values'] = common_columns
        old_columns_listbox.delete(0, tk.END)
        new_columns_listbox.delete(0, tk.END)
        for col in old_file.columns:
            old_columns_listbox.insert(tk.END, col)
        for col in new_file.columns:
            new_columns_listbox.insert(tk.END, col)

# Function to find the differences and matches between the old and new files
def find_differences(primarykey_column, old_columns_to_check, new_columns_to_check):
    try:
        if old_file is None or new_file is None:
            messagebox.showerror("Error", "Both files must be loaded first!")
            return [], []

        changes = []
        matches = []

        for idx, (old_col, new_col) in enumerate(zip(old_columns_to_check, new_columns_to_check)):
            old_data = old_file[old_col].tolist()
            new_data = new_file[new_col].tolist()

            for i, (old_val, new_val) in enumerate(zip(old_data, new_data)):
                if old_val != new_val:
                    changes.append({
                        primarykey_column: old_file.loc[i, primarykey_column],
                        'old_row': {old_col: old_val},
                        'new_row': {new_col: new_val}
                    })
                else:
                    matches.append({
                        primarykey_column: old_file.loc[i, primarykey_column],
                        'old_row': {old_col: old_val},
                        'new_row': {new_col: new_val}
                    })

        return changes, matches

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while finding differences: {e}")
        return [], []

# Function to display the changes in a new window
def display_changes(changes, matches):
    try:
        if not changes and not matches:
            messagebox.showinfo("No Changes", "No changes detected in the selected columns.")
        else:
            result_window = tk.Toplevel()
            result_window.title("Changes and Matches Detected")
            if changes:
                tk.Label(result_window, text="Changes:").pack()
                for change in changes:
                    tk.Label(result_window,
                             text=f"{primary_primarykey_column_combobox.get()}: {change[primary_primarykey_column_combobox.get()]} - Changes detected:").pack()
                    tk.Label(result_window, text=f"Old Row: {change['old_row']}").pack()
                    tk.Label(result_window, text=f"New Row: {change['new_row']}").pack()
                    tk.Label(result_window, text="-" * 40).pack()
            if matches:
                tk.Label(result_window, text="Matches:").pack()
                for match in matches:
                    tk.Label(result_window,
                             text=f"{primary_primarykey_column_combobox.get()}: {match[primary_primarykey_column_combobox.get()]} - Matching values:").pack()
                    tk.Label(result_window, text=f"Old Row: {match['old_row']}").pack()
                    tk.Label(result_window, text=f"New Row: {match['new_row']}").pack()
                    tk.Label(result_window, text="-" * 40).pack()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while displaying changes: {e}")

# Prompt for saving options
def prompt_save_options():
    return messagebox.askyesno("Save Options", "Do you want to save the entire row? (Yes for entire row, No for selected columns only)")

# Function to save the changes and matches to files
# Function to save the changes and matches to files
# Function to save the changes and matches to files
def save_changes_and_matches(changes, matches):
    try:
        if not changes and not matches:
            messagebox.showwarning("No Changes", "No changes or matches to save.")
            return

        save_entire_row = prompt_save_options()
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Changes and Matches", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                if changes:
                    for idx, change_column in enumerate(old_columns_listbox.curselection()):
                        change_column_name = old_columns_listbox.get(change_column)
                        sheet_name = f"Changes_{change_column_name}"
                        change_rows = []
                        for change in changes:
                            if change_column_name in change['old_row'] or change_column_name in change['new_row']:
                                if save_entire_row:
                                    change_row = new_file.loc[new_file[primary_primarykey_column_combobox.get()] == change[primary_primarykey_column_combobox.get()]].to_dict(orient='records')[0]
                                    change_rows.append(change_row)
                                else:
                                    combined_row = {primary_primarykey_column_combobox.get(): change[primary_primarykey_column_combobox.get()]}
                                    combined_row.update({f"{key} (Old)": change['old_row'].get(key, "") for key in change['old_row'].keys()})
                                    combined_row.update({f"{key} (New)": change['new_row'].get(key, "") for key in change['new_row'].keys()})
                                    change_rows.append(combined_row)
                        if change_rows:
                            change_df = pd.DataFrame(change_rows)
                            change_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            # Calculate number of changes and percentage
                            num_changes = len(change_df)
                            total_rows = len(new_file)
                            percentage_changes = (num_changes / total_rows) * 100
                            # Get the worksheet
                            worksheet = writer.sheets[sheet_name]
                            # Add bold format
                            bold_format = writer.book.add_format({'bold': True})
                            # Write number of changes and percentage in bold
                            worksheet.write(f'A{len(change_df) + 2}', f'Number of Changes: {num_changes}', bold_format)
                            worksheet.write(f'A{len(change_df) + 3}', f'Percentage of Changes: {percentage_changes:.2f}%', bold_format)

                if matches:
                    for idx, match_column in enumerate(old_columns_listbox.curselection()):
                        match_column_name = old_columns_listbox.get(match_column)
                        sheet_name = f"Matches_{match_column_name}"
                        match_rows = []
                        for match in matches:
                            if match_column_name in match['old_row'] or match_column_name in match['new_row']:
                                if save_entire_row:
                                    match_row = new_file.loc[new_file[primary_primarykey_column_combobox.get()] == match[primary_primarykey_column_combobox.get()]].to_dict(orient='records')[0]
                                    match_rows.append(match_row)
                                else:
                                    combined_row = {primary_primarykey_column_combobox.get(): match[primary_primarykey_column_combobox.get()]}
                                    combined_row.update({f"{key} (Old)": match['old_row'].get(key, "") for key in match['old_row'].keys()})
                                    combined_row.update({f"{key} (New)": match['new_row'].get(key, "") for key in match['new_row'].keys()})
                                    match_rows.append(combined_row)
                        if match_rows:
                            match_df = pd.DataFrame(match_rows)
                            match_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            # Calculate number of matches and percentage
                            num_matches = len(match_df)
                            total_rows = len(new_file)
                            percentage_matches = (num_matches / total_rows) * 100
                            # Get the worksheet
                            worksheet = writer.sheets[sheet_name]
                            # Add bold format
                            bold_format = writer.book.add_format({'bold': True})
                            # Write number of matches and percentage in bold
                            worksheet.write(f'A{len(match_df) + 2}', f'Number of Matches: {num_matches}', bold_format)
                            worksheet.write(f'A{len(match_df) + 3}', f'Percentage of Matches: {percentage_matches:.2f}%', bold_format)

                # Create a new sheet for the bar graph
                if changes or matches:
                    bar_sheet_name = "Changes_and_Matches_Percentage"
                    bar_data = []
                    for idx in old_columns_listbox.curselection():
                        column_name = old_columns_listbox.get(idx)
                        change_percentage = (len([change for change in changes if column_name in change['old_row'] or column_name in change['new_row']]) / len(new_file)) * 100
                        match_percentage = (len([match for match in matches if column_name in match['old_row'] or column_name in match['new_row']]) / len(new_file)) * 100
                        bar_data.append((column_name, change_percentage, match_percentage))
                    bar_df = pd.DataFrame(bar_data, columns=['Column', 'Change Percentage', 'Match Percentage'])
                    bar_df.to_excel(writer, sheet_name=bar_sheet_name, index=False)

                    # Create a column chart
                    column_chart = writer.book.add_chart({'type': 'column'})
                    column_chart.add_series({
                        'name': 'Change Percentage',
                        'categories': [bar_sheet_name, 1, 0, len(bar_df), 0],
                        'values': [bar_sheet_name, 1, 1, len(bar_df), 1],
                        'data_labels': {'value': True}
                    })
                    column_chart.add_series({
                        'name': 'Match Percentage',
                        'categories': [bar_sheet_name, 1, 0, len(bar_df), 0],
                        'values': [bar_sheet_name, 1, 2, len(bar_df), 2],
                        'data_labels': {'value': True}
                    })

                    column_chart.set_title({'name': 'Changes and Matches Percentage'})
                    column_chart.set_x_axis({'name': 'Column'})
                    column_chart.set_y_axis({'name': 'Percentage'})
                    column_chart.set_size({'width': 800, 'height': 600})
                    column_chart.set_legend({'position': 'top'})
                    bar_sheet = writer.sheets[bar_sheet_name]
                    bar_sheet.insert_chart('A10', column_chart)

            messagebox.showinfo("Success", "Changes and matches saved successfully")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving changes and matches: {e}")


def check_differences():
    primarykey_column = primary_primarykey_column_combobox.get()
    old_columns_to_check = [old_columns_listbox.get(idx) for idx in old_columns_listbox.curselection()]
    new_columns_to_check = [new_columns_listbox.get(idx) for idx in new_columns_listbox.curselection()]
    changes, matches = find_differences(primarykey_column, old_columns_to_check, new_columns_to_check)
    display_changes(changes, matches)
    if changes or matches:
        save_changes_and_matches(changes, matches)

def insert_values():
    try:
        if old_file is None or new_file is None:
            messagebox.showerror("Error", "Both files must be loaded first!")
            return

        # Get the selected columns from both old and new files
        old_columns_to_insert = [old_columns_listbox.get(idx) for idx in old_columns_listbox.curselection()]
        new_columns_to_insert = [new_columns_listbox.get(idx) for idx in new_columns_listbox.curselection()]

        # Iterate over each selected column pair
        for old_column, new_column in zip(old_columns_to_insert, new_columns_to_insert):
            for index, row in new_file.iterrows():
                if pd.isna(row[new_column]):
                    key_value = row[primary_primarykey_column_combobox.get()]
                    old_value = old_file.loc[old_file[primary_primarykey_column_combobox.get()] == key_value, old_column].values
                    if len(old_value) > 0:
                        if isinstance(old_value[0], str):
                            new_file.at[index, new_column] = str(old_value[0])  # Convert string value to string
                        else:
                            new_file.at[index, new_column] = old_value[0]  # Keep numerical value as is

        messagebox.showinfo("Success", "Values inserted successfully")

        # After inserting values, prompt user to save the modified new file
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Modified File", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            new_file.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "File saved successfully")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while inserting values: {e}")


root = tk.Tk()
root.maxsize(600, 500)
root.title("CSV/Excel Difference Checker")

# Set the background color for the main window
root.configure(bg="#000000")  # Black background color

# Define a bold font style
bold_font = ("Helvetica", 12, "bold")

# Add the centered, bold label with a custom color
title_label = tk.Label(root, text="DATA VALIDATOR", font=("Helvetica", 16, "bold"), fg="#ffffff", bg="#000000")
title_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

# Create and place the other widgets using grid
tk.Label(root, text="Load Old File", bg="#000000", fg="#ffffff", font=bold_font).grid(row=1, column=0, padx=10, pady=5, sticky="w")
tk.Button(root, text="Browse", command=load_old_file, bg="#ffffff", fg="#000000", font=bold_font).grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Load New File", bg="#000000", fg="#ffffff", font=bold_font).grid(row=2, column=0, padx=10, pady=5, sticky="w")
tk.Button(root, text="Browse", command=load_new_file, bg="#ffffff", fg="#000000", font=bold_font).grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Select the key column for comparison:", bg="#000000", fg="#ffffff", font=bold_font).grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="w")
primary_primarykey_column_combobox = ttk.Combobox(root, width=50)
primary_primarykey_column_combobox.grid(row=4, column=0, columnspan=2, padx=10, pady=5)

tk.Label(root, text="Select the columns from the old file:", bg="#000000", fg="#ffffff", font=bold_font).grid(row=5, column=0, padx=10, pady=5, sticky="w")
old_columns_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, exportselection=False, bg="#ffffff", fg="#000000", font=bold_font)  # White background for Listbox
old_columns_listbox.grid(row=6, column=0, padx=10, pady=5, sticky="nswe")

tk.Label(root, text="Select the columns from the new file:", bg="#000000", fg="#ffffff", font=bold_font).grid(row=5, column=1, padx=10, pady=5, sticky="w")
new_columns_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, exportselection=False, bg="#ffffff", fg="#000000", font=bold_font)  # White background for Listbox
new_columns_listbox.grid(row=6, column=1, padx=10, pady=5, sticky="nswe")

tk.Button(root, text="Check Differences", command=check_differences, bg="#ffffff", fg="#000000", font=bold_font).grid(row=7, column=0, padx=10, pady=10, sticky="w")
tk.Button(root, text="Insert Values", command=insert_values, bg="#ffffff", fg="#000000", font=bold_font).grid(row=7, column=1, padx=10, pady=10, sticky="e")

root.mainloop()