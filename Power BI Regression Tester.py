
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from PowerBIRegressionTester import PowerBIRegressionTester
from tabulate import tabulate
import os
import json

CONFIG_FILE = os.path.expanduser("~/.pbi_regression_tester_gui3.json")

def save_all_configs():
    with open(CONFIG_FILE, "w") as f:
        json.dump(configs, f)

def load_all_configs():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

def update_config_dropdown():
    config_names = list(configs.keys())
    config_dropdown['values'] = config_names
    if config_names:
        config_dropdown.current(0)
    else:
        config_dropdown.set("")

def load_config_to_fields(name):
    config = configs.get(name, {})
    project_folder_var.set(config.get("project_folder", ""))
    connection_string_var.set(config.get("connection_string", ""))
    pbi_report_folder_var.set(config.get("pbi_report_folder", ""))
    instance_name_var.set(config.get("instance_name", ""))

def on_config_select(event=None):
    name = config_dropdown.get()
    if name in configs:
        load_config_to_fields(name)

def save_current_config():
    name = config_dropdown.get().strip()
    if not name:
        name = simpledialog.askstring("Config Name", "Enter a name for this configuration:")
        if not name:
            return
        config_dropdown.set(name)
    configs[name] = {
        "project_folder": project_folder_var.get(),
        "connection_string": connection_string_var.get(),
        "pbi_report_folder": pbi_report_folder_var.get(),
        "instance_name": instance_name_var.get()
    }
    save_all_configs()
    update_config_dropdown()
    config_dropdown.set(name)
    messagebox.showinfo("Saved", f"Configuration '{name}' saved.")

def delete_current_config():
    name = config_dropdown.get()
    if name in configs:
        if messagebox.askyesno("Delete", f"Delete configuration '{name}'?"):
            del configs[name]
            save_all_configs()
            update_config_dropdown()
            
            # If there are configs left, load the selected one; else, clear fields
            current = config_dropdown.get()
            if current in configs:
                load_config_to_fields(current)
            else:
                project_folder_var.set("")
                connection_string_var.set("")
                pbi_report_folder_var.set("")
                instance_name_var.set("")
            messagebox.showinfo("Deleted", f"Configuration '{name}' deleted.")

def on_closing():
    save_all_configs()
    root.destroy()

def browse_folder(var):
    folder = filedialog.askdirectory()
    if folder:
        var.set(folder)

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

def run_compare():
    project_folder = project_folder_var.get()
    instance_name = instance_name_var.get().strip()
    if not instance_name:
        instance_name = simpledialog.askstring("Instance Name", "Enter an instance name to compare:")
        if not instance_name:
            return
        instance_name_var.set(instance_name)
    tester = PowerBIRegressionTester.for_compare_only(project_folder)
    if tester.instance_exists(instance_name):
        df = tester.compare(instance_name)
        if df is not None:
            show_result(df)
        else:
            messagebox.showinfo("Compare", "No differences found or comparison failed.")
    else:
        messagebox.showerror("Error", f"The instance '{instance_name}' does not exist.")
        
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

root = tk.Tk()
root.title("Power BI Regression Testing")

# Load configs
configs = load_all_configs()

project_folder_var = tk.StringVar()
connection_string_var = tk.StringVar()
pbi_report_folder_var = tk.StringVar()
instance_name_var = tk.StringVar()

# Config dropdown
tk.Label(root, text="Configuration:").grid(row=0, column=0, sticky='e')
config_dropdown = ttk.Combobox(root, width=40, state="normal")
config_dropdown.grid(row=0, column=1, padx=5, pady=2)
config_dropdown.bind("<<ComboboxSelected>>", on_config_select)

tk.Button(root, text="Save", command=save_current_config).grid(row=0, column=2, padx=2)
tk.Button(root, text="Delete", command=delete_current_config).grid(row=0, column=3, padx=2)

fields = [
    ("Project Folder", project_folder_var),
    ("Connection String", connection_string_var),
    ("PBI Report Folder (optional)", pbi_report_folder_var),
    ("Instance Name (for instance only)", instance_name_var)
]

for i, (label, var) in enumerate(fields, start=1):  # <-- start=1
    tk.Label(root, text=label).grid(row=i, column=0, sticky='e')
    entry = tk.Entry(root, textvariable=var, width=60)
    entry.grid(row=i, column=1, padx=5, pady=2)
    if "Folder" in label:
        tk.Button(root, text="Browse", command=lambda v=var: browse_folder(v)).grid(row=i, column=2)

tk.Button(root, text="Create Baseline", command=run_baseline, bg="lightblue").grid(row=len(fields)+1, column=0, pady=10)
tk.Button(root, text="Create Instance", command=run_instance, bg="lightgreen").grid(row=len(fields)+1, column=1, pady=10)
tk.Button(root, text="Compare", command=run_compare, bg="orange").grid(row=len(fields)+1, column=2, pady=10)

update_config_dropdown()
if config_dropdown['values']:
    load_config_to_fields(config_dropdown.get())

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()



