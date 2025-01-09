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
        data = [[grid[r][c].get() for c in range(cols)] for r in range(rows)]
        sheet.update("A1", data)
        messagebox.showinfo("Success", "Data saved to Google Sheets!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Save to CSV
def save_to_csv():
    data = [[grid[r][c].get() for c in range(cols)] for r in range(rows)]
    df = pd.DataFrame(data)
    filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filepath:
        df.to_csv(filepath, index=False, header=False)
        messagebox.showinfo("Success", "File saved successfully!")

# Load from Excel or CSV
def load_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if filepath:
        global full_data
        try:
            if filepath.endswith((".xlsx", ".xls")):
                full_data = pd.read_excel(filepath, header=None)
            elif filepath.endswith(".csv"):
                full_data = pd.read_csv(filepath, header=None)
            else:
                messagebox.showerror("Error", "Unsupported file format. Please select a CSV or Excel file.")
                return
            create_dynamic_grid(full_data)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")

# Apply filter
def apply_filter():
    try:
        global full_data
        filtered_data = full_data.copy()

        for c in range(cols):
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

# Create a dynamic grid with scrolling
def create_dynamic_grid(data):
    global rows, cols, grid
    rows, cols = data.shape

    # Clear the canvas
    for widget in canvas_frame.winfo_children():
        widget.destroy()

    # Add the new grid
    grid = [[tk.Entry(canvas_frame, width=10) for c in range(cols)] for r in range(rows)]
    for r, row in enumerate(data.values):
        for c, value in enumerate(row):
            entry = grid[r][c]
            entry.insert(0, value)
            entry.grid(row=r, column=c, sticky="nsew")

    # Adjust canvas scroll region
    canvas_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

# GUI Setup
root = tk.Tk()
root.title("Enhanced Spreadsheet with Filters")

# Scrollable Canvas
canvas = tk.Canvas(root)
scroll_y = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
scroll_x = tk.Scrollbar(root, orient="horizontal", command=canvas.xview)
canvas_frame = tk.Frame(canvas)

canvas.create_window((0, 0), window=canvas_frame, anchor="nw")
canvas.config(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

# Layout
canvas.grid(row=0, column=0, sticky="nsew")
scroll_y.grid(row=0, column=1, sticky="ns")
scroll_x.grid(row=1, column=0, sticky="ew")

# Filter Row
filter_frame = tk.Frame(root)
filter_frame.grid(row=2, column=0, sticky="ew")

filter_entries = [tk.Entry(filter_frame, width=10) for _ in range(5)]
filter_conditions = [tk.StringVar(filter_frame) for _ in range(5)]

for c in range(5):
    ttk.OptionMenu(filter_frame, filter_conditions[c], "Contains", "Contains", "Equals", "Range").grid(row=0, column=c)
    filter_entries[c].grid(row=1, column=c)

# Buttons
button_frame = tk.Frame(root)
button_frame.grid(row=3, column=0, columnspan=2, sticky="ew")

tk.Button(button_frame, text="Load File (Excel/CSV)", command=load_file).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Load from Google Sheets", command=load_from_google_sheets).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Save to Google Sheets", command=save_to_google_sheets).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Apply Filter", command=apply_filter).pack(side=tk.LEFT, padx=5, pady=5)
tk.Button(button_frame, text="Clear Filters", command=clear_filters).pack(side=tk.LEFT, padx=5, pady=5)

# Configure row/column weights
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# Initialize full_data as empty
full_data = pd.DataFrame()

root.mainloop()
