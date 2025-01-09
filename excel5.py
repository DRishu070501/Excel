import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
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

# Load file (Excel or CSV)
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

# Apply filter for individual columns
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

# Highlight selected cells/rows/columns
def highlight_cells_or_rows():
    try:
        color = colorchooser.askcolor(title="Choose a highlight color")[1]
        if not color:
            return

        # Highlight based on the mode selected (Cell, Row, Column)
        if highlight_mode.get() == "Cell" and selected_cell:
            r, c = selected_cell
            grid[r][c].config(bg=color)
            highlighted_cells.add((r, c))

        elif highlight_mode.get() == "Row" and selected_row is not None:
            for c in range(cols):
                grid[selected_row][c].config(bg=color)
            highlighted_rows.add(selected_row)

        elif highlight_mode.get() == "Column" and selected_col is not None:
            for r in range(rows):
                grid[r][selected_col].config(bg=color)
            highlighted_columns.add(selected_col)

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Reset Highlight (Remove all highlights)
def reset_highlight():
    global highlighted_cells, highlighted_rows, highlighted_columns

    # Reset highlighted cells
    for r, c in highlighted_cells:
        grid[r][c].config(bg="white")
    highlighted_cells.clear()

    # Reset highlighted rows
    for r in highlighted_rows:
        for c in range(cols):
            grid[r][c].config(bg="white")
    highlighted_rows.clear()

    # Reset highlighted columns
    for c in highlighted_columns:
        for r in range(rows):
            grid[r][c].config(bg="white")
    highlighted_columns.clear()

# Create a dynamic grid with scrolling
def create_dynamic_grid(data):
    global rows, cols, grid
    rows, cols = data.shape

    # Clear the canvas
    for widget in canvas_frame.winfo_children():
        widget.destroy()

    # Add the new grid
    grid = [[tk.Entry(canvas_frame, width=10, font=("Arial", 12)) for c in range(cols)] for r in range(rows)]
    for r, row in enumerate(data.values):
        for c, value in enumerate(row):
            entry = grid[r][c]
            entry.insert(0, value)
            entry.grid(row=r, column=c, sticky="nsew")

            # Bind mouse click event to capture row, column, or cell
            entry.bind("<Button-1>", lambda event, r=r, c=c: on_cell_click(event, r, c))

    # Adjust canvas scroll region
    canvas_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    # Create dynamic filter entries based on the number of columns
    update_filter_entries(cols)

# Dynamically update filter entries based on column count
def update_filter_entries(column_count):
    global filter_entries, filter_conditions
    for widget in filter_frame.winfo_children():
        widget.destroy()

    filter_entries = [tk.Entry(filter_frame, width=15, font=("Arial", 12)) for _ in range(column_count)]
    filter_conditions = [tk.StringVar(filter_frame) for _ in range(column_count)]

    for c in range(column_count):
        ttk.OptionMenu(filter_frame, filter_conditions[c], "Contains", "Contains", "Equals", "Range").grid(row=0, column=c, padx=5, pady=5)
        filter_entries[c].grid(row=1, column=c, padx=5, pady=5)

# Handle cell, row, or column selection
def on_cell_click(event, r, c):
    global selected_cell, selected_row, selected_col

    # Update the selected cell, row, or column
    selected_cell = (r, c)
    selected_row = r
    selected_col = c

    # Highlight the clicked cell with a temporary color
    if highlight_mode.get() == "Cell":
        grid[r][c].config(bg="lightblue")  # Highlight clicked cell
    elif highlight_mode.get() == "Row":
        for col in range(cols):
            grid[r][col].config(bg="lightblue")  # Highlight the whole row
    elif highlight_mode.get() == "Column":
        for row in range(rows):
            grid[row][c].config(bg="lightblue")  # Highlight the whole column

# GUI Setup
root = tk.Tk()
root.title("Enhanced Spreadsheet with Filters and Highlights")
root.geometry("900x600")  # Resize the window for better UI

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
filter_frame.grid(row=2, column=0, sticky="ew", pady=10)

# Buttons and Options
button_frame = tk.Frame(root)
button_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)

tk.Button(button_frame, text="Load File (Excel/CSV)", command=load_file, width=20).pack(side=tk.LEFT, padx=10, pady=5)
tk.Button(button_frame, text="Load from Google Sheets", command=load_from_google_sheets, width=20).pack(side=tk.LEFT, padx=10, pady=5)
tk.Button(button_frame, text="Save to Google Sheets", command=save_to_google_sheets, width=20).pack(side=tk.LEFT, padx=10, pady=5)
tk.Button(button_frame, text="Apply Filter", command=apply_filter, width=20).pack(side=tk.LEFT, padx=10, pady=5)
tk.Button(button_frame, text="Clear Filters", command=clear_filters, width=20).pack(side=tk.LEFT, padx=10, pady=5)
tk.Button(button_frame, text="Highlight", command=highlight_cells_or_rows, width=20).pack(side=tk.LEFT, padx=10, pady=5)
tk.Button(button_frame, text="Reset Highlight", command=reset_highlight, width=20).pack(side=tk.LEFT, padx=10, pady=5)

# Highlight Mode (Cell, Row, Column)
highlight_mode = tk.StringVar(value="Cell")
highlight_mode_frame = tk.Frame(root)
highlight_mode_frame.grid(row=4, column=0, sticky="ew", pady=10)
ttk.Label(highlight_mode_frame, text="Highlight Mode:").pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(highlight_mode_frame, text="Cell", variable=highlight_mode, value="Cell").pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(highlight_mode_frame, text="Row", variable=highlight_mode, value="Row").pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(highlight_mode_frame, text="Column", variable=highlight_mode, value="Column").pack(side=tk.LEFT, padx=5)

# Initialize variables for tracking selected cell, row, and column
selected_cell = None
selected_row = None
selected_col = None
highlighted_cells = set()
highlighted_rows = set()
highlighted_columns = set()

root.mainloop()
