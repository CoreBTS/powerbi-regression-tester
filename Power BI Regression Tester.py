
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog, messagebox
from PowerBIRegressionTester import PowerBIRegressionTester
from tabulate import tabulate
import os
import json

CONFIG_FILE = os.path.expanduser("~/.pbi_regression_tester_gui.json")

# Save config on exit
def on_closing():
    save_config()
    root.destroy()

def save_config():
    config = {
        "project_folder": project_folder_var.get(),
        "connection_string": connection_string_var.get(),
        "pbi_report_folder": pbi_report_folder_var.get(),
        "instance_name": instance_name_var.get()
    }
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
            project_folder_var.set(config.get("project_folder", ""))
            connection_string_var.set(config.get("connection_string", ""))
            pbi_report_folder_var.set(config.get("pbi_report_folder", ""))
            instance_name_var.set(config.get("instance_name", ""))

def run_baseline():
    tester = PowerBIRegressionTester(
        project_folder_var.get(),
        connection_string_var.get(),
        pbi_report_folder_var.get() if pbi_report_folder_var.get() else ""
    )
    if tester.baseline_exists():
        if not messagebox.askyesno("Overwrite?", "Baseline exists. Overwrite?"):
            return
    df = tester.run_baseline()
    show_result(df)

def run_instance():
    instance_name = instance_name_var.get().strip()
    if not instance_name:
        messagebox.showerror("Error", "Instance name required.")
        return
    tester = PowerBIRegressionTester(
        project_folder_var.get(),
        connection_string_var.get(),
        pbi_report_folder_var.get() if pbi_report_folder_var.get() else ""
    )
    if tester.instance_exists(instance_name):
        if not messagebox.askyesno("Overwrite?", f"Instance '{instance_name}' exists. Overwrite?"):
            return
    df = tester.run_instance(instance_name)
    show_result(df)

def browse_folder(var):
    folder = filedialog.askdirectory()
    if folder:
        var.set(folder)

def show_result(df):
    if df is not None and not df.empty:
        result_window = tk.Toplevel(root)
        result_window.title("Result Table")

        frame = ttk.Frame(result_window)
        frame.pack(fill='both', expand=True)
        tree_scroll_y = ttk.Scrollbar(frame, orient="vertical")
        tree_scroll_y.pack(side='right', fill='y')
        tree_scroll_x = ttk.Scrollbar(frame, orient="horizontal")
        tree_scroll_x.pack(side='bottom', fill='x')

        tree = ttk.Treeview(frame, columns=list(df.columns), show='headings',
                            yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        tree.pack(fill='both', expand=True)
        tree_scroll_y.config(command=tree.yview)
        tree_scroll_x.config(command=tree.xview)

        # Store full values for copy
        cell_values = {}

        for col in df.columns:
            tree.heading(col, text=col)
            width = 200 if col.lower() == "query" else 100
            tree.column(col, width=width, anchor='w')

        for row_idx, row in df.iterrows():
            values = []
            for col in df.columns:
                val = row[col]
                display_val = val
                if col.lower() == "query" and isinstance(val, str):
                    display_val = val[:20] + "..." if len(val) > 20 else val
                    display_val = display_val.replace('\n\t', ' ').replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
                values.append(display_val)
                cell_values[(row_idx, col)] = val  # Store full value
            tree.insert('', 'end', iid=row_idx, values=values)

        # Context menu for copying cell value
        menu = tk.Menu(result_window, tearoff=0)
        menu.add_command(label="Copy", command=lambda: copy_cell())

        def copy_cell():
            if hasattr(tree, 'clicked_cell'):
                row_idx, col_idx = tree.clicked_cell
                col_name = df.columns[col_idx]
                value = cell_values.get((row_idx, col_name), "")
                result_window.clipboard_clear()
                result_window.clipboard_append(str(value))

        def on_right_click(event):
            region = tree.identify("region", event.x, event.y)
            if region == "cell":
                row_id = tree.identify_row(event.y)
                col_id = tree.identify_column(event.x)
                if row_id and col_id:
                    row_idx = int(row_id)
                    col_idx = int(col_id.replace("#", "")) - 1
                    tree.clicked_cell = (row_idx, col_idx)
                    menu.tk_popup(event.x_root, event.y_root)

        tree.bind("<Button-3>", on_right_click)  # Right-click

        # Double-click to show/copy full text
        def on_double_click(event):
            item = tree.selection()
            if item:
                row = tree.item(item, "values")
                col = tree.identify_column(event.x)
                col_index = int(col.replace("#", "")) - 1
                col_name = df.columns[col_index]
                full_text = df.iloc[tree.index(item[0]), col_index]
                if col_name.lower() == "query" and isinstance(full_text, str):
                    # Show full text in a popup with copy option
                    popup = tk.Toplevel(result_window)
                    popup.title(f"Full text: {col_name}")
                    text_box = tk.Text(popup, wrap='word', width=80, height=10)
                    text_box.insert('1.0', str(full_text))
                    text_box.pack(expand=True, fill='both')
                    text_box.config(state='normal')
                    def copy_to_clipboard():
                        popup.clipboard_clear()
                        popup.clipboard_append(str(full_text))
                    copy_btn = tk.Button(popup, text="Copy", command=copy_to_clipboard)
                    copy_btn.pack()

        tree.bind("<Double-1>", on_double_click)

def show_result_OLD(df):
    if df is not None and not df.empty:
        # Create a new window
        result_window = tk.Toplevel(root)
        result_window.title("Result Table")

        # Scrollbars
        frame = ttk.Frame(result_window)
        frame.pack(fill='both', expand=True)
        tree_scroll_y = ttk.Scrollbar(frame, orient="vertical")
        tree_scroll_y.pack(side='right', fill='y')
        tree_scroll_x = ttk.Scrollbar(frame, orient="horizontal")
        tree_scroll_x.pack(side='bottom', fill='x')

        # Treeview setup
        tree = ttk.Treeview(frame, columns=list(df.columns), show='headings',
                            yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        tree.pack(fill='both', expand=True)
        tree_scroll_y.config(command=tree.yview)
        tree_scroll_x.config(command=tree.xview)

        # Set up columns and headings
        for col in df.columns:
            tree.heading(col, text=col)
            # Set width, make long text columns shorter
            width = 200 if col.lower() == "query" else 100
            tree.column(col, width=width, anchor='w')

        # Insert data (truncate long text columns)
        for _, row in df.iterrows():
            values = []
            for col in df.columns:
                val = row[col]
                if col.lower() == "query" and isinstance(val, str):
                    display_val = val[:20] + "..." if len(val) > 20 else val
                else:
                    display_val = val
                values.append(display_val)
            tree.insert('', 'end', values=values)

        # Double-click to show/copy full text
        def on_double_click(event):
            item = tree.selection()
            if item:
                row = tree.item(item, "values")
                col = tree.identify_column(event.x)
                col_index = int(col.replace("#", "")) - 1
                col_name = df.columns[col_index]
                full_text = df.iloc[tree.index(item[0]), col_index]
                # Show full text in a popup with copy option
                popup = tk.Toplevel(result_window)
                popup.title(f"Full text: {col_name}")
                text_box = tk.Text(popup, wrap='word', width=80, height=10)
                text_box.insert('1.0', str(full_text))
                text_box.pack(expand=True, fill='both')
                text_box.config(state='normal')
                def copy_to_clipboard():
                    popup.clipboard_clear()
                    popup.clipboard_append(str(full_text))
                copy_btn = tk.Button(popup, text="Copy", command=copy_to_clipboard)
                copy_btn.pack()
        tree.bind("<Double-1>", on_double_click)

root = tk.Tk()
root.title("Power BI Regression Testing")

project_folder_var = tk.StringVar()
connection_string_var = tk.StringVar()
pbi_report_folder_var = tk.StringVar()
instance_name_var = tk.StringVar()

fields = [
    ("Project Folder", project_folder_var),
    ("Connection String", connection_string_var),
    ("PBI Report Folder (optional)", pbi_report_folder_var),
    ("Instance Name (for instance only)", instance_name_var)
]

for i, (label, var) in enumerate(fields):
    tk.Label(root, text=label).grid(row=i, column=0, sticky='e')
    entry = tk.Entry(root, textvariable=var, width=60)
    entry.grid(row=i, column=1, padx=5, pady=2)
    if "Folder" in label:
        tk.Button(root, text="Browse", command=lambda v=var: browse_folder(v)).grid(row=i, column=2)

tk.Button(root, text="Create Baseline", command=run_baseline, bg="lightblue").grid(row=len(fields), column=0, pady=10)
tk.Button(root, text="Create Instance", command=run_instance, bg="lightgreen").grid(row=len(fields), column=1, pady=10)

load_config()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()



