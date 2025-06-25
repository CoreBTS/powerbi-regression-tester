import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import json
import shutil
from PowerBIRegressionTester import PowerBIRegressionTester

class PowerBIRegressionTesterApp:
    CONFIG_FILE = os.path.expanduser("~/.pbi_regression_tester_gui5.json")

    def __init__(self, root):
        self.root = root
        self.root.title("Power BI Regression Testing")
        self.configs = self.load_all_configs()

        # Variables
        self.project_folder_var = tk.StringVar()
        self.connection_string_var = tk.StringVar()
        self.pbi_report_folder_var = tk.StringVar()
        self.instance_name_var = tk.StringVar()

        # Widgets
        self.setup_widgets()
        self.update_project_folder_dropdown()
        if self.project_folder_dropdown['values']:
            self.load_config_to_fields(self.project_folder_dropdown.get())

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_widgets(self):
        # Config dropdown
        tk.Label(self.root, text="Project Folder:").grid(row=0, column=0, sticky='e')
        self.project_folder_dropdown = ttk.Combobox(self.root, textvariable=self.project_folder_var, width=40, state="readonly")
        self.project_folder_dropdown.grid(row=0, column=1, padx=5, pady=2, sticky='w')
        self.project_folder_dropdown.bind("<<ComboboxSelected>>", self.on_project_folder_select)
        tk.Button(self.root, text="Save", command=self.save_current_config).grid(row=0, column=2, padx=2, sticky='w')
        tk.Button(self.root, text="Delete", command=self.delete_current_config).grid(row=0, column=3, padx=2, sticky='w')
        tk.Button(self.root, text="New", command=self.create_new_config).grid(row=0, column=4, padx=2, sticky='w')

        # Fields
        fields = [
            ("Connection String", self.connection_string_var),
            ("PBI Report Folder (optional)", self.pbi_report_folder_var),
            ("Instance Name (for instance only)", None)
        ]
        for i, (label, var) in enumerate(fields, start=1):
            tk.Label(self.root, text=label).grid(row=i, column=0, sticky='e')
            if label.startswith("Instance Name"):
                self.instance_dropdown = ttk.Combobox(self.root, textvariable=self.instance_name_var, width=40, state="readonly")
                self.instance_dropdown.grid(row=i, column=1, padx=5, pady=2, sticky='w')
                tk.Button(self.root, text="Delete", command=self.delete_current_instance, fg="red").grid(row=i, column=2, padx=2, sticky='w')
                # tk.Button(self.root, text="View Instance", command=self.view_instance, bg="white").grid(row=len(fields)+1, column=3, pady=10, sticky='w')
                # tk.Button(self.root, text="Create Instance", command=self.run_instance, bg="lightgreen").grid(row=i, column=3, padx=2, sticky='w')
                tk.Button(self.root, text="View Instance", command=self.view_instance, bg="white").grid(row=i, column=3, padx=2, sticky='w')
                tk.Button(self.root, text="Create Instance", command=self.run_instance, bg="lightgreen").grid(row=i, column=4, padx=2, sticky='w')
            else:
                entry = tk.Entry(self.root, textvariable=var, width=60)
                entry.grid(row=i, column=1, padx=5, pady=2)
                if "Folder" in label:
                    tk.Button(self.root, text="Browse", command=lambda v=var: self.browse_folder(v)).grid(row=i, column=2, sticky='w')

        # Action buttons
        # tk.Button(self.root, text="Create Baseline", command=self.run_baseline, bg="lightblue").grid(row=len(fields)+1, column=0, pady=10)
        # # tk.Button(self.root, text="Create Instance", command=self.run_instance, bg="lightgreen").grid(row=len(fields)+1, column=1, pady=10)
        # tk.Button(self.root, text="Compare", command=self.run_compare, bg="orange").grid(row=len(fields)+1, column=2, pady=10)

        # Action buttons
        tk.Button(self.root, text="Create Baseline", command=self.run_baseline, bg="lightblue").grid(row=len(fields)+1, column=0, pady=10)
        tk.Button(self.root, text="View Baseline", command=self.view_baseline, bg="white").grid(row=len(fields)+1, column=1, pady=10, sticky='w')
        tk.Button(self.root, text="Compare", command=self.run_compare, bg="orange").grid(row=len(fields)+1, column=3, pady=10, sticky='w')

        self.project_folder_var.trace_add("write", lambda *args: self.update_instance_dropdown())

    # --- All your logic functions become methods below ---
    def browse_folder(self, var):
        folder = filedialog.askdirectory()
        if folder:
            var.set(folder)

    def update_instance_dropdown(self):
        instance_base = os.path.join(os.getcwd(), PowerBIRegressionTester.PROJECT_FOLDER_BASE, self.project_folder_var.get(), PowerBIRegressionTester.INSTANCE_FOLDER_NAME)
        # instance_base = os.path.join(project_path)
        instance_names = []
        if os.path.isdir(instance_base):
            instance_names = [name for name in os.listdir(instance_base)
                              if os.path.isdir(os.path.join(instance_base, name))]
        self.instance_dropdown['values'] = instance_names
        if instance_names:
            self.instance_dropdown.current(0)
        else:
            self.instance_dropdown.set("")

    def update_config_dropdown(self):
        config_names = list(self.configs.keys())
        self.config_dropdown['values'] = config_names
        if config_names:
            self.config_dropdown.current(0)
        else:
            self.config_dropdown.set("")

    def update_project_folder_dropdown(self):
        project_folders = list(self.configs.keys())
        self.project_folder_dropdown['values'] = project_folders
        if project_folders:
            self.project_folder_dropdown.current(0)
        else:
            self.project_folder_dropdown.set("")

    def create_new_config(self):
        new_name = simpledialog.askstring("New Project Folder", "Enter a name for the new project folder:")
        if not new_name:
            return
        if new_name in self.configs:
            messagebox.showerror("Error", f"Project folder '{new_name}' already exists.")
            return
        # Add new config with empty/default values
        self.configs[new_name] = {
            "connection_string": "",
            "pbi_report_folder": "",
            "instance_name": ""
        }
        self.save_all_configs()
        PowerBIRegressionTester.create_project_skeleton(new_name)
        self.update_project_folder_dropdown()
        self.project_folder_dropdown.set(new_name)
        self.load_config_to_fields(new_name)

    def load_config_to_fields(self, project_folder):
        config = self.configs.get(project_folder, {})
        self.project_folder_var.set(project_folder)
        self.connection_string_var.set(config.get("connection_string", ""))
        self.pbi_report_folder_var.set(config.get("pbi_report_folder", ""))
        self.update_instance_dropdown()
        instance_name = config.get("instance_name", "")
        self.instance_dropdown.set(instance_name)

    def on_config_select(self, event=None):
        name = self.config_dropdown.get()
        if name in self.configs:
            self.load_config_to_fields(name)

    def on_project_folder_select(self, event=None):
        project_folder = self.project_folder_var.get()
        if project_folder in self.configs:
            self.load_config_to_fields(project_folder)
    
    def save_current_config(self):
        project_folder = self.project_folder_var.get().strip()
        if not project_folder:
            project_folder = simpledialog.askstring("Config Name", "Enter a name for this configuration:")
            if not project_folder:
                return
            self.project_folder_dropdown.set(project_folder)
        self.configs[project_folder] = {
            "connection_string": self.connection_string_var.get(),
            "pbi_report_folder": self.pbi_report_folder_var.get(),
            "instance_name": self.instance_dropdown.get().strip()}
        self.save_all_configs()
        self.update_project_folder_dropdown()
        self.project_folder_dropdown.set(project_folder)
        PowerBIRegressionTester.create_project_skeleton(project_folder)
        messagebox.showinfo("Saved", f"Configuration '{project_folder}' saved.")

    def delete_current_config(self):
        project_folder = self.project_folder_dropdown.get()
        if project_folder in self.configs:
            if messagebox.askyesno("Delete", f"Delete configuration '{project_folder}'?"):
                del self.configs[project_folder]
                self.save_all_configs()
                self.update_project_folder_dropdown()
                current = self.update_project_folder_dropdown.get()
                if current in self.configs:
                    self.load_config_to_fields(current)
                else:
                    self.project_folder_var.set("")
                    self.connection_string_var.set("")
                    self.pbi_report_folder_var.set("")
                    self.instance_dropdown.set("")
                    self.instance_dropdown['values'] = []
                messagebox.showinfo("Deleted", f"Configuration '{project_folder}' deleted.")

    def delete_current_instance(self):
        instance_name = self.instance_dropdown.get().strip()
        if not instance_name:
            messagebox.showerror("Error", "No instance selected to delete.")
            return
        project_path = os.path.join(os.getcwd(), self.project_folder_var.get())
        instance_base = os.path.join(project_path, "instance")
        instance_path = os.path.join(instance_base, instance_name)
        if not os.path.isdir(instance_path):
            messagebox.showerror("Error", f"Instance folder '{instance_name}' does not exist.")
            return
        if messagebox.askyesno("Delete Instance", f"Are you sure you want to delete instance '{instance_name}'? This cannot be undone."):
            try:
                shutil.rmtree(instance_path)
                self.update_instance_dropdown()
                self.instance_dropdown.set("")
                messagebox.showinfo("Deleted", f"Instance '{instance_name}' deleted.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete instance: {e}")


    def save_current_instance(self):
        name = self.instance_dropdown.get().strip()
        if name and name not in self.instance_dropdown['values']:
            # Add new instance name to dropdown
            values = list(self.instance_dropdown['values'])
            values.append(name)
            self.instance_dropdown['values'] = values
            self.instance_dropdown.set(name)
            
    def save_all_configs(self):
        with open(self.CONFIG_FILE, "w") as f:
            json.dump(self.configs, f)

    def load_all_configs(self):
        if os.path.exists(self.CONFIG_FILE):
            with open(self.CONFIG_FILE, "r") as f:
                return json.load(f)
        return {}

    def on_closing(self):
        self.save_all_configs()
        self.root.destroy()

    def show_result(self, df):
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
    
    def view_baseline(self):
        tester = PowerBIRegressionTester(
            self.project_folder_var.get(),
            self.connection_string_var.get(),
            self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        )
        if not tester.baseline_exists():
            messagebox.showerror("Error", "No baseline exists for this project.")
            return
        df = tester.load_baseline_df()
        self.show_result(df)

    def view_instance(self):
        instance_name = self.instance_dropdown.get().strip()
        if not instance_name:
            messagebox.showerror("Error", "No instance selected.")
            return
        # tester = PowerBIRegressionTester(
        #     self.project_folder_var.get(),
        #     self.connection_string_var.get(),
        #     self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        # )
        tester = PowerBIRegressionTester.for_compare_only(self.project_folder_var.get())
        if not tester.instance_exists(instance_name):
            messagebox.showerror("Error", f"Instance '{instance_name}' does not exist.")
            return
        df = tester.load_instance_df(instance_name)
        self.show_result(df)
        
    def run_baseline(self):
        # tester = PowerBIRegressionTester(
        #     self.project_folder_var.get(),
        #     self.connection_string_var.get(),
        #     self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        # )
        tester = PowerBIRegressionTester.for_compare_only(self.project_folder_var.get())
        if tester.baseline_exists():
            if not messagebox.askyesno("Overwrite?", "Baseline exists. Overwrite?"):
                return
        df = tester.run_baseline()
        self.show_result(df)

    def run_instance(self):
        # instance_name = self.instance_dropdown.get().strip()
        # if not instance_name:
        #     messagebox.showerror("Error", "Instance name required.")
        #     return
        
        # Prompt for new instance name if none selected
        instance_name = simpledialog.askstring("New Instance Name", "Enter a name for the new instance:")
        if not instance_name:
            return
        self.instance_dropdown.set(instance_name)

        tester = PowerBIRegressionTester(
            self.project_folder_var.get(),
            self.connection_string_var.get(),
            self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        )
        if tester.instance_exists(instance_name):
            if not messagebox.askyesno("Overwrite?", f"Instance '{instance_name}' exists. Overwrite?"):
                return
        df = tester.run_instance(instance_name)
        self.save_current_instance()
        self.show_result(df)

    def run_compare(self):
        project_folder = self.project_folder_var.get()
        instance_name = self.instance_dropdown.get().strip()
        if not instance_name:
            instance_name = simpledialog.askstring("Instance Name", "Enter an instance name to compare:")
            if not instance_name:
                return
            self.instance_dropdown.set(instance_name)
        tester = PowerBIRegressionTester.for_compare_only(project_folder)
        if tester.instance_exists(instance_name):
            df = tester.compare(instance_name)
            if df is not None:
                self.show_result(df)
            else:
                messagebox.showinfo("Compare", "No differences found or comparison failed.")
        else:
            messagebox.showerror("Error", f"The instance '{instance_name}' does not exist.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PowerBIRegressionTesterApp(root)
    root.mainloop()