import pandas as pd
import ezdxf
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)
        load_sheets(file_path)

def load_sheets(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        sheet_dropdown['values'] = xl.sheet_names
        if xl.sheet_names:
            sheet_dropdown.current(0)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load sheets:\n{e}")

def create_dxf():
    try:
        file_path = file_entry.get()
        sheet = sheet_var.get()
        x_col = x_col_entry.get().upper()
        y_col = y_col_entry.get().upper()
        x_from = int(x_from_entry.get()) - 1
        x_to = int(x_to_entry.get()) - 1
        y_from = int(y_from_entry.get()) - 1
        y_to = int(y_to_entry.get()) - 1
        color = int(color_entry.get())

        if not file_path or not sheet:
            messagebox.showerror("Input Error", "File path or sheet not specified.")
            return

        df = pd.read_excel(file_path, sheet_name=sheet, header=None)

        x_index = ord(x_col) - ord('A')
        y_index = ord(y_col) - ord('A')

        x_vals = df.iloc[x_from:x_to+1, x_index].values
        y_vals = df.iloc[y_from:y_to+1, y_index].values

        if len(x_vals) != len(y_vals):
            messagebox.showerror("Data Error", "X and Y row ranges do not match in length.")
            return

        points = list(zip(x_vals, y_vals))
        points.sort(key=lambda pt: pt[0])

        doc = ezdxf.new()
        msp = doc.modelspace()
        layer_name = "PolylineLayer"
        if layer_name not in doc.layers:
            doc.layers.new(name=layer_name, dxfattribs={"color": color})
        msp.add_lwpolyline(points, dxfattribs={"layer": layer_name})

        output_path = "output_polyline.dxf"
        doc.saveas(output_path)
        messagebox.showinfo("Success", f"DXF created and saved as {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")

# --- GUI Setup ---
root = tk.Tk()
root.title("Excel to AutoCAD Polyline (DXF Creator)")
root.geometry("600x400")

frame = ttk.Frame(root, padding=10)
frame.pack(fill=tk.BOTH, expand=True)

# File Selection
ttk.Label(frame, text="Excel File:").grid(row=0, column=0, sticky="e")
file_entry = ttk.Entry(frame, width=50)
file_entry.grid(row=0, column=1)
ttk.Button(frame, text="Browse", command=select_file).grid(row=0, column=2)

# Sheet Dropdown
ttk.Label(frame, text="Sheet:").grid(row=1, column=0, sticky="e")
sheet_var = tk.StringVar()
sheet_dropdown = ttk.Combobox(frame, textvariable=sheet_var, state="readonly")
sheet_dropdown.grid(row=1, column=1, columnspan=2, sticky="we")

# X Settings
ttk.Label(frame, text="X Column (e.g., A):").grid(row=2, column=0, sticky="e")
x_col_entry = ttk.Entry(frame, width=5)
x_col_entry.grid(row=2, column=1, sticky="w")

ttk.Label(frame, text="X Row From:").grid(row=3, column=0, sticky="e")
x_from_entry = ttk.Entry(frame, width=5)
x_from_entry.grid(row=3, column=1, sticky="w")

ttk.Label(frame, text="X Row To:").grid(row=4, column=0, sticky="e")
x_to_entry = ttk.Entry(frame, width=5)
x_to_entry.grid(row=4, column=1, sticky="w")

# Y Settings
ttk.Label(frame, text="Y Column (e.g., B):").grid(row=2, column=2, sticky="e")
y_col_entry = ttk.Entry(frame, width=5)
y_col_entry.grid(row=2, column=3, sticky="w")

ttk.Label(frame, text="Y Row From:").grid(row=3, column=2, sticky="e")
y_from_entry = ttk.Entry(frame, width=5)
y_from_entry.grid(row=3, column=3, sticky="w")

ttk.Label(frame, text="Y Row To:").grid(row=4, column=2, sticky="e")
y_to_entry = ttk.Entry(frame, width=5)
y_to_entry.grid(row=4, column=3, sticky="w")

# Color
ttk.Label(frame, text="Layer Color (1-255):").grid(row=5, column=0, sticky="e")
color_entry = ttk.Entry(frame, width=5)
color_entry.grid(row=5, column=1, sticky="w")
color_entry.insert(0, "1")  # Default red

# Create Button
ttk.Button(frame, text="Create DXF", command=create_dxf).grid(row=6, column=0, columnspan=4, pady=20)

root.mainloop()
#pip install pandas ezdxf openpyxl
