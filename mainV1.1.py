import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd
import ezdxf
import os

def rgb_to_aci(r, g, b):
    color_map = {
        (255, 0, 0): 1, (255, 255, 0): 2, (0, 255, 0): 3,
        (0, 255, 255): 4, (0, 0, 255): 5, (255, 0, 255): 6,
        (255, 255, 255): 7
    }
    return color_map.get((r, g, b), 7)

class PolylineGroup:
    def __init__(self, parent, index, remove_callback):
        self.frame = ttk.LabelFrame(parent, text=f"Polyline {index+1}")
        self.index = index
        self.remove_callback = remove_callback
        self.color_rgb = (255, 0, 0)
        self.color_aci = 1

        self.x_col = self._add_row("X Column (A-Z):", 0)
        self.x_from = self._add_row("X Row From:", 1)
        self.x_to = self._add_row("X Row To:", 2)

        self.y_col = self._add_row("Y Column (A-Z):", 0, col_offset=2)
        self.y_from = self._add_row("Y Row From:", 1, col_offset=2)
        self.y_to = self._add_row("Y Row To:", 2, col_offset=2)

        ttk.Label(self.frame, text="Color:").grid(row=3, column=0, sticky="e")
        self.color_button = ttk.Button(self.frame, text="Pick Color", command=self.pick_color)
        self.color_button.grid(row=3, column=1, sticky="w")

        self.plot_points_var = tk.BooleanVar()
        ttk.Checkbutton(self.frame, text="Plot Points", variable=self.plot_points_var).grid(row=3, column=2, sticky="w")

        self.remove_btn = ttk.Button(self.frame, text="Remove", command=self.remove)
        self.remove_btn.grid(row=3, column=3, sticky="e")

        self.frame.grid_columnconfigure(1, weight=1)
        self.frame.grid_columnconfigure(3, weight=1)

    def _add_row(self, label, row, col_offset=0):
        ttk.Label(self.frame, text=label).grid(row=row, column=col_offset, sticky="e")
        entry = ttk.Entry(self.frame, width=10)
        entry.grid(row=row, column=col_offset + 1, sticky="w")
        return entry

    def pick_color(self):
        color_code = colorchooser.askcolor(title="Choose Color", initialcolor=self.color_rgb)
        if color_code[0]:
            self.color_rgb = tuple(map(int, color_code[0]))
            self.color_aci = rgb_to_aci(*self.color_rgb)
            self.color_button.configure(text=f"Color ACI: {self.color_aci}")

    def remove(self):
        self.frame.destroy()
        self.remove_callback(self)

    def get_data(self):
        try:
            x_col = ord(self.x_col.get().upper()) - ord('A')
            y_col = ord(self.y_col.get().upper()) - ord('A')
            x_from = int(self.x_from.get()) - 1
            x_to = int(self.x_to.get()) - 1
            y_from = int(self.y_from.get()) - 1
            y_to = int(self.y_to.get()) - 1
            return x_col, x_from, x_to, y_col, y_from, y_to, self.color_aci, self.plot_points_var.get()
        except Exception as e:
            raise ValueError(f"Invalid input in Polyline {self.index+1}: {e}")

class DXFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Add Polyline to AutoCad via Excel | v1.1 by Pendar")
        self.groups = []

        top_frame = ttk.Frame(root, padding=10)
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text="Excel File:").pack(side=tk.LEFT)
        self.file_entry = ttk.Entry(top_frame, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="Browse", command=self.select_file).pack(side=tk.LEFT)

        sheet_frame = ttk.Frame(root, padding=10)
        sheet_frame.pack(fill=tk.X)
        ttk.Label(sheet_frame, text="Sheet:").pack(side=tk.LEFT)
        self.sheet_var = tk.StringVar()
        self.sheet_dropdown = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, state="readonly")
        self.sheet_dropdown.pack(side=tk.LEFT, fill=tk.X, expand=True)

        options_frame = ttk.Frame(root, padding=10)
        options_frame.pack(fill=tk.X)

        ttk.Label(options_frame, text="Output DXF Name:").pack(side=tk.LEFT)
        self.output_entry = ttk.Entry(options_frame, width=30)
        self.output_entry.pack(side=tk.LEFT, padx=5)
        self.output_entry.insert(0, "output")

        self.debug_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Debug Mode", variable=self.debug_var).pack(side=tk.LEFT, padx=5)

        self.group_container = ttk.Frame(root, padding=10)
        self.group_container.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(root, padding=10)
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="➕ Add Polyline", command=self.add_group).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Generate DXF", command=self.generate_dxf).pack(side=tk.RIGHT)

        self.add_group()

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)
            try:
                xl = pd.ExcelFile(path)
                self.sheet_dropdown['values'] = xl.sheet_names
                if xl.sheet_names:
                    self.sheet_dropdown.current(0)
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def add_group(self):
        index = len(self.groups)
        group = PolylineGroup(self.group_container, index, self.remove_group)
        group.frame.pack(fill=tk.X, pady=5)
        self.groups.append(group)

    def remove_group(self, group):
        self.groups.remove(group)

    def generate_dxf(self):
        try:
            file_path = self.file_entry.get()
            sheet = self.sheet_var.get()
            output_base = self.output_entry.get().strip()
            debug = self.debug_var.get()

            if not os.path.exists(file_path):
                raise FileNotFoundError("Excel file not found.")
            if not sheet:
                raise ValueError("No sheet selected.")
            if not output_base:
                raise ValueError("Output filename cannot be empty.")

            output_name = output_base + ".dxf"

            df = pd.read_excel(file_path, sheet_name=sheet, header=None)
            all_points = []
            layer_colors = []
            point_flags = []

            for group in self.groups:
                x_col, x_from, x_to, y_col, y_from, y_to, color, plot_points = group.get_data()
                x_vals = df.iloc[x_from:x_to+1, x_col].values
                y_vals = df.iloc[y_from:y_to+1, y_col].values
                if len(x_vals) != len(y_vals):
                    raise ValueError(f"X and Y values length mismatch in group {group.index + 1}")
                points = list(zip(x_vals, y_vals))
                points.sort(key=lambda pt: pt[0])
                all_points.append(points)
                layer_colors.append(color)
                point_flags.append(plot_points)

            doc = ezdxf.new()
            msp = doc.modelspace()

            for i, (pts, color, plot_pts) in enumerate(zip(all_points, layer_colors, point_flags)):
                layer_name = f"Polyline_{i+1}"
                if layer_name not in doc.layers:
                    doc.layers.new(name=layer_name, dxfattribs={"color": color})
                msp.add_lwpolyline(pts, dxfattribs={"layer": layer_name})

                if plot_pts:
                    for pt in pts:
                        msp.add_point(pt, dxfattribs={"layer": layer_name, "color": color})

            doc.saveas(output_name)

            if debug:
                self.debug_compare(output_name, all_points)

            messagebox.showinfo("Success", f"DXF saved as '{output_name}' with {len(all_points)} polylines.")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def debug_compare(self, dxf_path, input_data):
        try:
            doc = ezdxf.readfile(dxf_path)
            msp = doc.modelspace()
            debug_text = "Debug Comparison:\n\n"

            for i, (pline, input_points) in enumerate(zip(msp.query("LWPOLYLINE"), input_data)):
                dxf_points = list(pline.get_points("xy"))
                debug_text += f"Polyline {i+1}:\n"
                debug_text += f"  Input Points   ({len(input_points)}): {input_points}\n"
                debug_text += f"  DXF File Points({len(dxf_points)}): {dxf_points}\n"

                mismatch = any(abs(ix - dx) > 1e-6 or abs(iy - dy) > 1e-6
                               for (ix, iy), (dx, dy) in zip(input_points, dxf_points))
                debug_text += f"  ✅ Match: {'No mismatch' if not mismatch else '❌ MISMATCH!'}\n\n"

            top = tk.Toplevel(self.root)
            top.title("Debug Report")
            txt = tk.Text(top, wrap=tk.WORD, height=30, width=100)
            txt.insert(tk.END, debug_text)
            txt.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        except Exception as e:
            messagebox.showerror("Debug Error", f"Failed to read DXF: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DXFApp(root)
    root.mainloop()
