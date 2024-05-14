import os
import re
import pandas as pd
from collections import defaultdict
from tkinter import Tk, Label, Button, Entry, StringVar, filedialog, Listbox, Scrollbar, END
from datetime import datetime

def find_duplicates(path):
    pattern = re.compile(r'^[A-Za-z]?[0-9]{6}[A-Za-z]?$')  # Regex pattern for filenames
    files_map = defaultdict(list)
    for dirpath, dirnames, filenames in os.walk(path):
        for filename in filenames:
            base_name, ext = os.path.splitext(filename)  # Split filename and extension
            if ext.lower() in ('.dxf', '.pdf') and pattern.match(base_name):  # Filter by file type and pattern
                full_path = os.path.join(dirpath, filename)
                mod_time = os.path.getmtime(full_path)
                mod_date = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
                files_map[filename].append((full_path, mod_date))

    duplicates = {filename: details for filename, details in files_map.items() if len(details) > 1}
    return duplicates

def browse_folder():
    initial_dir = "\\\\SERVEUR2019\\Dessins\\PENDING"  # Set default directory path
    directory = filedialog.askdirectory(initialdir=initial_dir)
    path_entry.set(directory)

def scan_for_duplicates():
    folder_path = path_entry.get()
    duplicates = find_duplicates(folder_path)
    duplicates_list.delete(0, END)  # Clear the listbox

    # Sort by modification date, descending for each filename
    for filename, details in sorted(duplicates.items(), key=lambda x: x[1][0][1], reverse=True):
        # Sort the details by modification date within each filename, descending
        sorted_details = sorted(details, key=lambda x: x[1], reverse=True)
        duplicates_list.insert(END, f"{filename}:")
        for path, mod_date in sorted_details:
            duplicates_list.insert(END, f"  Last Modified: {mod_date} - Path: {path}")
        duplicates_list.insert(END, "")  # Add a blank line for readability

def choose_output_folder():
    default_output_dir = os.path.expanduser("~/Desktop")  # Default to user's desktop
    output_directory = filedialog.askdirectory(initialdir=default_output_dir)
    output_path_entry.set(output_directory)  # Set the chosen directory or keep default

def save_to_excel(duplicates, excel_path):
    data = []
    for filename, details in sorted(duplicates.items(), key=lambda x: x[1][0][1], reverse=True):
        sorted_details = sorted(details, key=lambda x: x[1], reverse=True)
        for path, mod_date in sorted_details:
            data.append({"Filename": filename, "Modified Date": mod_date, "Path": path})

    df = pd.DataFrame(data)
    if not excel_path.endswith(".xlsx"):
        excel_path += "/duplicates_list.xlsx"
    df.to_excel(excel_path, index=False)

def save_results():
    folder_path = path_entry.get()
    output_path = output_path_entry.get() if output_path_entry.get() else os.path.expanduser("~/Desktop")
    excel_file_path = os.path.join(output_path, 'duplicates_list.xlsx')
    duplicates = find_duplicates(folder_path)
    if duplicates:
        save_to_excel(duplicates, excel_file_path)
        result_label.config(text=f"Results saved to {excel_file_path}")
    else:
        result_label.config(text="No duplicates to save.")

root = Tk()
root.title("Duplicate File Finder for DXF and PDF with Specific Naming Pattern")

# Entries and labels for paths
path_label = Label(root, text="Select Folder to Check:")
path_label.pack()

path_entry = StringVar(root)
entry = Entry(root, textvariable=path_entry, width=50)
entry.pack()

browse_button = Button(root, text="Browse", command=browse_folder)
browse_button.pack()

output_label = Label(root, text="Select Output Folder for Excel File:")
output_label.pack()

output_path_entry = StringVar(root)
output_entry = Entry(root, textvariable=output_path_entry, width=50)
output_entry.pack()

output_button = Button(root, text="Choose Output Folder", command=choose_output_folder)
output_button.pack()

scan_button = Button(root, text="Scan for Duplicates", command=scan_for_duplicates)
scan_button.pack()

save_button = Button(root, text="Save to Excel", command=save_results)
save_button.pack()

result_label = Label(root, text="")
result_label.pack()

# Scrollbar for the listbox
scrollbar = Scrollbar(root)
scrollbar.pack(side="right", fill="y")

duplicates_list = Listbox(root, yscrollcommand=scrollbar.set, width=100, height=20)
duplicates_list.pack()

scrollbar.config(command=duplicates_list.yview)


output_path_entry.set(os.path.expanduser("~/Desktop"))
path_entry.set("\\\\SERVEUR2019\\Dessins\\PENDING")
root.mainloop()