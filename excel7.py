import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from tkinter import simpledialog
from tkinter import colorchooser

# Global Variables
full_data = pd.DataFrame()
rows, cols = 0, 0
highlighted_cells = set()
highlighted_rows = set()
highlighted_columns = set()
selected_cell = None
selected_row = None
selected_col = None

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

# Save file (Excel or CSV)
def save_file():
    filepath = filedialog.asksaveasfilename(
        defaultextension=".csv", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    )
    if filepath:
        try:
            if filepath.endswith(".xlsx"):
                full_data.to_excel(filepath, index=False, header=False)
            elif filepath.endswith(".csv"):
                full_data.to_csv(filepath, index=False, header=False)
            else:
                messagebox.showerror("Error", "Unsupported file format.")
                return
            messagebox.showinfo("Success", f"File saved to {filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")

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

    # Clear the canvas and grid frame
    for widget in canvas_frame.winfo_children():
        widget.destroy()

    # Add the new grid
    grid = [[tk.Entry(canvas_frame, width=15, font=("Arial", 10)) for c in range(cols)] for r in range(rows)]
    for r, row in enumerate(data.values):
        for c, value in enumerate(row):
            entry = grid[r][c]
            entry.insert(0, value)
            entry.grid(row=r, column=c, sticky="nsew", padx=5, pady=5)

            # Bind mouse click event to capture row, column, or cell
            entry.bind("<Button-1>", lambda event, r=r, c=c: on_cell_click(event, r, c))

    # Update canvas scroll region after the grid is created
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
root.geometry("1000x600")  # Resize the window for better UI

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

# Buttons and Options (smaller buttons)
button_frame = tk.Frame(root)
button_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)

def style_button(btn):
    btn.config(width=12, height=1, font=("Arial", 10), relief="solid", bg="#4CAF50", fg="white", bd=1)

buttons = [
    ("Open File", load_file),
    ("Save File", save_file),
    ("Apply Filter", apply_filter),
    ("Clear Filters", clear_filters),
    ("Highlight", highlight_cells_or_rows),
    ("Reset Highlight", reset_highlight),
]

for text, command in buttons:
    btn = tk.Button(button_frame, text=text, command=command)
    style_button(btn)
    btn.pack(side="left", padx=5)

highlight_mode = tk.StringVar(value="Cell")
highlight_modes = ["Cell", "Row", "Column"]
highlight_mode_menu = ttk.OptionMenu(button_frame, highlight_mode, *highlight_modes)
highlight_mode_menu.pack(side="left", padx=5)

root.mainloop()
