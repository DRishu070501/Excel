import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Google Sheets setup
def google_sheets_setup():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    return client

# Load data from Google Sheets
def load_from_google_sheets():
    try:
        client = google_sheets_setup()
        sheet = client.open("Your Google Sheet Name").sheet1
        global full_data
        full_data = pd.DataFrame(sheet.get_all_values())
        create_dynamic_grid(full_data)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Save data to Google Sheets
def save_to_google_sheets():
    try:
        client = google_sheets_setup()
        sheet = client.open("Your Google Sheet Name").sheet1
        data = [[grid[r][c].get() for c in range(len(grid[0]))] for r in range(len(grid))]
        sheet.update("A1", data)
        messagebox.showinfo("Success", "Data saved to Google Sheets!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Save to CSV
def save_to_csv():
    data = [[grid[r][c].get() for c in range(len(grid[0]))] for r in range(len(grid))]
    df = pd.DataFrame(data)
    filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filepath:
        df.to_csv(filepath, index=False, header=False)
        messagebox.showinfo("Success", "File saved successfully!")

# Load from CSV
def load_from_csv():
    filepath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if filepath:
        global full_data
        full_data = pd.read_csv(filepath, header=None)
        create_dynamic_grid(full_data)

# Load from Excel or CSV
def load_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if filepath:
        global full_data
        try:
            if filepath.endswith((".xlsx", ".xls")):
                # Load Excel file
                full_data = pd.read_excel(filepath, header=None)
            elif filepath.endswith(".csv"):
                # Load CSV file
                full_data = pd.read_csv(filepath, header=None)
            else:
                messagebox.showerror("Error", "Unsupported file format. Please select a CSV or Excel file.")
                return
            create_dynamic_grid(full_data)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")


# Display data in the grid
def create_dynamic_grid(data):
    # Clear the grid frame
    for widget in grid_frame.winfo_children():
        widget.destroy()

    # Create a dynamic grid
    global grid
    grid = [[tk.Entry(grid_frame, width=10) for _ in range(data.shape[1])] for _ in range(data.shape[0])]
    for r, row in enumerate(data.values):
        for c, value in enumerate(row):
            entry = grid[r][c]
            entry.grid(row=r, column=c)
            entry.insert(0, value)

# Apply filter
def apply_filter():
    try:
        global full_data
        filtered_data = full_data.copy()

        # Apply filters for each column
        for c in range(len(filter_entries)):
            filter_value = filter_entries[c].get().strip()
            if filter_value:
                if filter_conditions[c].get() == "Contains":
                    filtered_data = filtered_data[filtered_data[c].astype(str).str.contains(filter_value, na=False)]
                elif filter_conditions[c].get() == "Equals":
                    filtered_data = filtered_data[filtered_data[c].astype(str) == filter_value]
                elif filter_conditions[c].get() == "Range":
                    try:
                        min_val, max_val = map(float, filter_value.split("-"))
                        filtered_data = filtered_data[(filtered_data[c].astype(float) >= min_val) &
                                                      (filtered_data[c].astype(float) <= max_val)]
                    except ValueError:
                        messagebox.showerror("Error", "Invalid range format. Use 'min-max'.")
                        return

        create_dynamic_grid(filtered_data)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Clear filters
def clear_filters():
    for entry in filter_entries:
        entry.delete(0, tk.END)
    create_dynamic_grid(full_data)

# Initialize global variables
filter_entries = []
filter_conditions = []
grid = []
full_data = pd.DataFrame()

# GUI Setup
root = tk.Tk()
root.title("Dynamic Spreadsheet with Filters")

# Frames for layout
filter_frame = tk.Frame(root)
filter_frame.pack(fill=tk.X)

grid_frame = tk.Frame(root)
grid_frame.pack(fill=tk.BOTH, expand=True)

button_frame = tk.Frame(root)
button_frame.pack(fill=tk.X)

# Add filter options
def setup_filters(column_count):
    for widget in filter_frame.winfo_children():
        widget.destroy()

    filter_entries.clear()
    filter_conditions.clear()

    for c in range(column_count):
        filter_condition = tk.StringVar(root)
        filter_condition.set("Contains")
        ttk.OptionMenu(filter_frame, filter_condition, "Contains", "Contains", "Equals", "Range").grid(row=0, column=c)
        filter_entry = tk.Entry(filter_frame, width=10)
        filter_entry.grid(row=1, column=c)
        filter_entries.append(filter_entry)
        filter_conditions.append(filter_condition)

# Buttons
tk.Button(button_frame, text="Load from Google Sheets", command=load_from_google_sheets).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Save to Google Sheets", command=save_to_google_sheets).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Save to CSV", command=save_to_csv).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Load from CSV", command=load_from_csv).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Apply Filter", command=apply_filter).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Clear Filters", command=clear_filters).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Load File (Excel/CSV)", command=load_file).pack(side=tk.LEFT, padx=5, pady=5)

root.mainloop()
