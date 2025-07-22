import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import json
import shutil
import base64
import sys

if sys.platform == "win32":
    import win32crypt
from PowerBIRegressionTester import PowerBIRegressionTester

class PowerBIRegressionTesterApp:
    CONFIG_FILE = os.path.expanduser("~/.pbi_regression_tester_gui5.json")

    def __init__(self, root):
        self.root = root
        self.root.title("Power BI Regression Testing")
        self.configs = self.load_all_configs()

        # Variables
        self.project_folder_var = tk.StringVar()
        self.pbi_report_folder_var = tk.StringVar()
        self.instance_name_var = tk.StringVar()

        # Widgets
        self.setup_widgets()
        self.update_project_folder_dropdown()
        if self.project_folder_dropdown['values']:
            self.load_config_to_fields(self.project_folder_dropdown.get())

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def encrypt_for_user(self, plaintext):
        if sys.platform != "win32":
            return plaintext  # fallback: no encryption
        if not plaintext:
            return ""
        encrypted = win32crypt.CryptProtectData(
            plaintext.encode("utf-8"), None, None, None, None, 0)
        return base64.b64encode(encrypted).decode("utf-8")

    def decrypt_for_user(self, ciphertext):
        if sys.platform != "win32":
            return ciphertext  # fallback: no decryption
        if not ciphertext:
            return ""
        try:
            encrypted = base64.b64decode(ciphertext)
            decrypted = win32crypt.CryptUnprotectData(encrypted, None, None, None, 0)[1]
            return decrypted.decode("utf-8")
        except Exception:
            return ""  # or raise
        
    def setup_widgets(self):
        tk.Label(self.root, text="Project:").grid(row=0, column=0, sticky='e')
        self.project_folder_dropdown = ttk.Combobox(self.root, textvariable=self.project_folder_var, width=40, state="readonly")
        self.project_folder_dropdown.grid(row=0, column=1, padx=5, pady=2, sticky='w')
        self.project_folder_dropdown.bind("<<ComboboxSelected>>", self.on_project_folder_select)
        tk.Button(self.root, text="Save", command=self.save_current_config).grid(row=0, column=2, padx=2, sticky='w')
        tk.Button(self.root, text="Delete", command=self.delete_current_config).grid(row=0, column=3, padx=2, sticky='w')
        tk.Button(self.root, text="New", command=self.create_new_config).grid(row=0, column=4, padx=2, sticky='w')

        # Project folder field
        tk.Label(self.root, text="PBI Report Folder (optional)").grid(row=1, column=0, sticky='e')
        entry = tk.Entry(self.root, textvariable=self.pbi_report_folder_var, width=60)
        entry.grid(row=1, column=1, padx=5, pady=2)
        tk.Button(self.root, text="Browse", command=lambda v=self.pbi_report_folder_var: self.browse_folder(v)).grid(row=1, column=2, sticky='w')

        # Instance dropdown
        tk.Label(self.root, text="Instance Name").grid(row=2, column=0, sticky='e')
        self.instance_dropdown = ttk.Combobox(self.root, textvariable=self.instance_name_var, width=40, state="readonly")
        self.instance_dropdown.grid(row=2, column=1, padx=5, pady=2, sticky='w')
        tk.Button(self.root, text="Delete", command=self.delete_current_instance, fg="red").grid(row=2, column=2, padx=2, sticky='w')
        tk.Button(self.root, text="View Instance", command=self.view_instance, bg="white").grid(row=2, column=3, padx=2, sticky='w')
        tk.Button(self.root, text="Create Instance", command=self.run_instance, bg="lightgreen").grid(row=2, column=4, padx=2, sticky='w')

        # Action buttons
        tk.Button(self.root, text="Create Baseline", command=self.run_baseline, bg="lightblue").grid(row=3, column=0, pady=10)
        tk.Button(self.root, text="View Baseline", command=self.view_baseline, bg="white").grid(row=3, column=1, pady=10, sticky='w')
        tk.Button(self.root, text="Compare", command=self.run_compare, bg="orange").grid(row=3, column=3, pady=10, sticky='w')
        tk.Button(self.root, text="Compare Instance To", command=self.compare_to_dialog, bg="orange").grid(row=3, column=4, padx=2, sticky='w')

        tk.Button(self.root, text="Edit Baseline", command=self.edit_baseline, bg="lightyellow").grid(row=4, column=0, pady=5)
        tk.Button(self.root, text="Edit Instance", command=self.edit_selected_instance, bg="lightyellow").grid(row=4, column=1, pady=5)

        tk.Button(self.root, text="Run Selected Instance", command=self.run_selected_instance, bg="lightgreen").grid(row=4, column=2, pady=5)

        self.project_folder_var.trace_add("write", lambda *args: self.update_instance_dropdown())

    def browse_folder(self, var):
        folder = filedialog.askdirectory()
        if folder:
            var.set(folder)

    def run_selected_instance(self):
        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return

        instance_name = self.instance_dropdown.get().strip()
        if not instance_name:
            messagebox.showerror("Error", "No instance selected.")
            return

        instance = next((inst for inst in project.get("instances", []) if inst["instance_name"] == instance_name), None)
        if not instance:
            messagebox.showerror("Error", f"Instance '{instance_name}' not found.")
            return


        # if instance.get("interactive"):
        #     conn_str = None
        # else:
        #     conn_str = (
        #         f"Provider=MSOLAP;Data Source={instance['server_name']};"
        #         f"Initial Catalog={instance['database_name']};"
        #         f"User ID={instance['user_id']};Password={self.decrypt_for_user(instance['password'])}"
        #     )

        # tester = PowerBIRegressionTester(
        #     self.project_folder_var.get(),
        #     conn_str,
        #     self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        # )

        tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance)

        if tester.instance_exists(instance_name):
            if not messagebox.askyesno("Overwrite?", f"Instance '{instance_name}' exists. Overwrite?"):
                return

        df = tester.run_instance(instance_name)
        if df is not None and not df.empty:
            self.show_result(df)
        else:
            messagebox.showinfo("Compare", "No differences found.")
            
    def update_instance_dropdown(self):
        project = self.get_project(self.project_folder_var.get())
        instance_names = []
        if project and "instances" in project:
            instance_names = [
                inst["instance_name"]
                for inst in project["instances"]
                if inst["instance_name"].lower() != "baseline"
            ]
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
        project_folders = [proj["name"] for proj in self.configs.get("projects", [])]
        self.project_folder_dropdown['values'] = project_folders
        if project_folders:
            self.project_folder_dropdown.current(0)
        else:
            self.project_folder_dropdown.set("")

    def create_new_config(self):
        new_name = simpledialog.askstring("New Project", "Enter a name for the new project:")
        if not new_name:
            return
        if self.get_project(new_name):
            messagebox.showerror("Error", f"Project '{new_name}' already exists.")
            return
        pbi_report_folder = simpledialog.askstring("PBI Report Folder", "Enter PBI Report Folder (optional):") or ""
        self.configs.setdefault("projects", []).append({
            "name": new_name,
            "pbi_report_folder": pbi_report_folder,
            "instances": []
        })
        self.save_all_configs()
        self.update_project_folder_dropdown()
        self.project_folder_dropdown.set(new_name)
        self.pbi_report_folder_var.set(pbi_report_folder)
        self.update_instance_dropdown()

    def load_config_to_fields(self, project_name):
        project = self.get_project(project_name)
        self.project_folder_var.set(project_name)
        self.pbi_report_folder_var.set(project.get("pbi_report_folder", "") if project else "")
        self.update_instance_dropdown()
        # Only set to first non-baseline instance
        if project and project.get("instances"):
            non_baseline = [
                inst["instance_name"]
                for inst in project["instances"]
                if inst["instance_name"].lower() != "baseline"
            ]
            if non_baseline:
                self.instance_dropdown.set(non_baseline[0])
            else:
                self.instance_dropdown.set("")
        else:
            self.instance_dropdown.set("")

    def on_config_select(self, event=None):
        name = self.config_dropdown.get()
        if name in self.configs:
            self.load_config_to_fields(name)

    def on_project_folder_select(self, event=None):
        project_folder = self.project_folder_var.get()
        if project_folder in self.configs:
            self.load_config_to_fields(project_folder)
    
    def save_current_config(self):
        project_name = self.project_folder_var.get().strip()
        if not project_name:
            project_name = simpledialog.askstring("Project Name", "Enter a name for this project:")
            if not project_name:
                return
            self.project_folder_dropdown.set(project_name)
        project = self.get_project(project_name)
        if not project:
            project = {"name": project_name, "pbi_report_folder": self.pbi_report_folder_var.get(), "instances": []}
            self.configs.setdefault("projects", []).append(project)
        project["pbi_report_folder"] = self.pbi_report_folder_var.get()
        self.save_all_configs()
        self.update_project_folder_dropdown()
        self.project_folder_dropdown.set(project_name)
        messagebox.showinfo("Saved", f"Project '{project_name}' saved.")

    def delete_current_config(self):
        project_name = self.project_folder_var.get()
        projects = self.configs.get("projects", [])
        idx = next((i for i, p in enumerate(projects) if p["name"] == project_name), None)
        if idx is not None:
            if messagebox.askyesno("Delete", f"Delete project '{project_name}'?"):
                del projects[idx]
                self.save_all_configs()
                self.update_project_folder_dropdown()
                if self.project_folder_dropdown['values']:
                    self.load_config_to_fields(self.project_folder_dropdown.get())
                else:
                    self.project_folder_var.set("")
                    self.pbi_report_folder_var.set("")
                    self.instance_dropdown.set("")
                    self.instance_dropdown['values'] = []
                messagebox.showinfo("Deleted", f"Project '{project_name}' deleted.")

    def delete_current_instance(self):
        project = self.get_project(self.project_folder_var.get())
        instance_name = self.instance_dropdown.get().strip()
        if not instance_name or not project or "instances" not in project:
            messagebox.showerror("Error", "No instance selected to delete.")
            return
        idx = next((i for i, inst in enumerate(project["instances"]) if inst["instance_name"] == instance_name), None)
        if idx is not None:
            if messagebox.askyesno("Delete Instance", f"Are you sure you want to delete instance '{instance_name}'? This cannot be undone."):
                del project["instances"][idx]
                self.save_all_configs()
                self.update_instance_dropdown()
                self.instance_dropdown.set("")
                messagebox.showinfo("Deleted", f"Instance '{instance_name}' deleted.")

    def prompt_instance_details(self, default_name=""):
        dialog = tk.Toplevel(self.root)
        dialog.title("Instance Details")
        dialog.grab_set()
        dialog.resizable(False, False)

        # Variables
        instance_name_var = tk.StringVar(value=default_name)
        server_name_var = tk.StringVar()
        database_name_var = tk.StringVar()
        user_id_var = tk.StringVar()
        password_var = tk.StringVar()
        interactive_var = tk.BooleanVar(value=False)
        xmla_endpoint_var = tk.BooleanVar(value=False)
        local_instance_var = tk.BooleanVar(value=False)

        # Layout
        tk.Label(dialog, text="Instance Name:").grid(row=0, column=0, sticky="e")
        name_entry = tk.Entry(dialog, textvariable=instance_name_var, width=40)
        name_entry.grid(row=0, column=1, padx=5, pady=2)

        tk.Label(dialog, text="Server Name:").grid(row=1, column=0, sticky="e")
        server_entry = tk.Entry(dialog, textvariable=server_name_var, width=40)
        server_entry.grid(row=1, column=1, padx=5, pady=2)

        tk.Label(dialog, text="Database Name:").grid(row=2, column=0, sticky="e")
        db_entry = tk.Entry(dialog, textvariable=database_name_var, width=40)
        db_entry.grid(row=2, column=1, padx=5, pady=2)

        tk.Label(dialog, text="User ID:").grid(row=3, column=0, sticky="e")
        user_entry = tk.Entry(dialog, textvariable=user_id_var, width=40)
        user_entry.grid(row=3, column=1, padx=5, pady=2)

        tk.Label(dialog, text="Password:").grid(row=4, column=0, sticky="e")
        pass_entry = tk.Entry(dialog, textvariable=password_var, width=40, show="*")
        pass_entry.grid(row=4, column=1, padx=5, pady=2)

        interactive_chk = tk.Checkbutton(dialog, text="Interactive", variable=interactive_var)
        interactive_chk.grid(row=5, column=0, columnspan=2, pady=5)

        xmla_chk = tk.Checkbutton(dialog, text="XMLA Endpoint", variable=xmla_endpoint_var)
        xmla_chk.grid(row=6, column=0, columnspan=2, pady=5)

        local_chk = tk.Checkbutton(dialog, text="Local Instance", variable=local_instance_var)
        local_chk.grid(row=7, column=0, columnspan=2, pady=5)

        # Disable/enable fields based on checkbox
        def toggle_fields(*args):
            # Only disable User ID and Password if Interactive is checked
            user_entry.config(state="disabled" if interactive_var.get() else "normal")
            pass_entry.config(state="disabled" if interactive_var.get() else "normal")
            # Server and DB always enabled
            server_entry.config(state="normal")
            db_entry.config(state="normal")

            # Disable/clear XMLA and Interactive if Local Instance is checked
            if local_instance_var.get():
                xmla_endpoint_var.set(False)
                interactive_var.set(False)
                xmla_chk.config(state="disabled")
                interactive_chk.config(state="disabled")
            else:
                xmla_chk.config(state="normal")
                interactive_chk.config(state="normal")

        interactive_var.trace_add("write", toggle_fields)
        local_instance_var.trace_add("write", toggle_fields)
        toggle_fields()  # Set initial state

        # Buttons
        result = {}
        def on_ok():
            result["instance_name"] = instance_name_var.get().strip()
            result["interactive"] = interactive_var.get()
            if not result["instance_name"]:
                messagebox.showerror("Error", "Instance Name is required.", parent=dialog)
                return
            
            if not (interactive_var.get() or xmla_endpoint_var.get() or local_instance_var.get()):
                messagebox.showerror("Error", "Either 'Interactive' or 'XMLA Endpoint' must be checked.", parent=dialog)
                return

            result["server_name"] = server_name_var.get().strip()
            result["database_name"] = database_name_var.get().strip()
            result["xmla_endpoint"] = xmla_endpoint_var.get()
            result["local_instance"] = local_instance_var.get()

            if not result["interactive"]:
                result["user_id"] = user_id_var.get().strip()
                result["password"] = self.encrypt_for_user(password_var.get())
                if not all([result["server_name"], result["database_name"]]):
                    messagebox.showerror("Error", "All fields are required unless Interactive is checked.", parent=dialog)
                    return
            dialog.destroy()

        def on_cancel():
            result.clear()
            dialog.destroy()

        tk.Button(dialog, text="OK", command=on_ok).grid(row=7, column=0, pady=10)
        tk.Button(dialog, text="Cancel", command=on_cancel).grid(row=7, column=1, pady=10)

        dialog.wait_window()
        return result if result else None
        
    def edit_selected_instance(self):
        project = self.get_project(self.project_folder_var.get())
        instance_name = self.instance_dropdown.get().strip()
        if not project or not instance_name:
            messagebox.showerror("Error", "No instance selected to edit.")
            return
        instance = next((inst for inst in project.get("instances", []) if inst["instance_name"] == instance_name), None)
        if not instance:
            messagebox.showerror("Error", f"Instance '{instance_name}' not found.")
            return

        # Prompt with current values
        edited = self.prompt_instance_details_prefilled(instance)
        if not edited:
            return
        # Overwrite the instance
        idx = next((i for i, inst in enumerate(project["instances"]) if inst["instance_name"] == instance_name), None)
        if idx is not None:
            project["instances"][idx] = edited
            self.save_all_configs()
            self.update_instance_dropdown()
            self.instance_dropdown.set(edited["instance_name"])
            messagebox.showinfo("Saved", f"Instance '{edited['instance_name']}' updated.")

    def edit_baseline(self):
        # For baseline, you may want to use a convention (e.g., instance_name == "Baseline")
        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return
        instance = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == "baseline"), None)
        if not instance:
            messagebox.showerror("Error", "No baseline instance found.")
            return
        edited = self.prompt_instance_details_prefilled(instance)
        if not edited:
            return
        idx = next((i for i, inst in enumerate(project["instances"]) if inst["instance_name"].lower() == "baseline"), None)
        if idx is not None:
            project["instances"][idx] = edited
            self.save_all_configs()
            self.update_instance_dropdown()
            if edited["instance_name"] != "Baseline":
                self.instance_dropdown.set(edited["instance_name"])
            messagebox.showinfo("Saved", "Baseline instance updated.")

    def prompt_instance_details_prefilled(self, instance):
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Instance Details")
        dialog.grab_set()
        dialog.resizable(False, False)

        # Variables
        instance_name_var = tk.StringVar(value=instance.get("instance_name", ""))
        server_name_var = tk.StringVar(value=instance.get("server_name", ""))
        database_name_var = tk.StringVar(value=instance.get("database_name", ""))
        user_id_var = tk.StringVar(value=instance.get("user_id", ""))
        password_var = tk.StringVar(value=self.decrypt_for_user(instance.get("password", "")))
        interactive_var = tk.BooleanVar(value=instance.get("interactive", False))
        xmla_endpoint_var = tk.BooleanVar(value=instance.get("xmla_endpoint", False) if "instance" in locals() else False)
        local_instance_var = tk.BooleanVar(value=instance.get("local_instance", False) if "instance" in locals() else False)

        # Layout
        tk.Label(dialog, text="Instance Name:").grid(row=0, column=0, sticky="e")
        name_entry = tk.Entry(dialog, textvariable=instance_name_var, width=40)
        name_entry.config(state="disabled" if instance.get("instance_name", "") == 'Baseline' else "normal")
        name_entry.grid(row=0, column=1, padx=5, pady=2)

        tk.Label(dialog, text="Server Name:").grid(row=1, column=0, sticky="e")
        server_entry = tk.Entry(dialog, textvariable=server_name_var, width=40)
        server_entry.grid(row=1, column=1, padx=5, pady=2)

        tk.Label(dialog, text="Database Name:").grid(row=2, column=0, sticky="e")
        db_entry = tk.Entry(dialog, textvariable=database_name_var, width=40)
        db_entry.grid(row=2, column=1, padx=5, pady=2)

        tk.Label(dialog, text="User ID:").grid(row=3, column=0, sticky="e")
        user_entry = tk.Entry(dialog, textvariable=user_id_var, width=40)
        user_entry.grid(row=3, column=1, padx=5, pady=2)

        tk.Label(dialog, text="Password:").grid(row=4, column=0, sticky="e")
        pass_entry = tk.Entry(dialog, textvariable=password_var, width=40, show="*")
        pass_entry.grid(row=4, column=1, padx=5, pady=2)

        interactive_chk = tk.Checkbutton(dialog, text="Interactive", variable=interactive_var)
        interactive_chk.grid(row=5, column=0, columnspan=2, pady=5)

        xmla_chk = tk.Checkbutton(dialog, text="XMLA Endpoint", variable=xmla_endpoint_var)
        xmla_chk.grid(row=6, column=0, columnspan=2, pady=5)

        local_chk = tk.Checkbutton(dialog, text="Local Instance", variable=local_instance_var)
        local_chk.grid(row=7, column=0, columnspan=2, pady=5)

        # Disable/enable fields based on checkbox
        def toggle_fields(*args):
            # Only disable User ID and Password if Interactive is checked
            user_entry.config(state="disabled" if interactive_var.get() else "normal")
            pass_entry.config(state="disabled" if interactive_var.get() else "normal")
            # Server Name and Database Name always enabled
            server_entry.config(state="normal")
            db_entry.config(state="normal")

            # Disable/clear XMLA and Interactive if Local Instance is checked
            if local_instance_var.get():
                xmla_endpoint_var.set(False)
                interactive_var.set(False)
                xmla_chk.config(state="disabled")
                interactive_chk.config(state="disabled")
            else:
                xmla_chk.config(state="normal")
                interactive_chk.config(state="normal")

        interactive_var.trace_add("write", toggle_fields)
        local_instance_var.trace_add("write", toggle_fields)
        toggle_fields()  # Set initial state

        # Buttons
        result = {}
        def on_ok():
            result["instance_name"] = instance_name_var.get().strip()
            result["interactive"] = interactive_var.get()
            if not result["instance_name"]:
                messagebox.showerror("Error", "Instance Name is required.", parent=dialog)
                return

            if not (interactive_var.get() or xmla_endpoint_var.get() or local_instance_var.get()):
                messagebox.showerror("Error", "Either 'Interactive' or 'XMLA Endpoint' must be checked.", parent=dialog)
                return
            
            result["server_name"] = server_name_var.get().strip()
            result["database_name"] = database_name_var.get().strip()
            result["xmla_endpoint"] = xmla_endpoint_var.get()
            result["local_instance"] = local_instance_var.get()

            if not result["interactive"]:
                result["user_id"] = user_id_var.get().strip()
                result["password"] = self.encrypt_for_user(password_var.get())
                if not all([result["server_name"], result["database_name"]]):
                    messagebox.showerror("Error", "All fields are required unless Interactive is checked.", parent=dialog)
                    return
            dialog.destroy()

        def on_cancel():
            result.clear()
            dialog.destroy()

        tk.Button(dialog, text="OK", command=on_ok).grid(row=7, column=0, pady=10)
        tk.Button(dialog, text="Cancel", command=on_cancel).grid(row=7, column=1, pady=10)

        dialog.wait_window()
        return result if result else None

    def save_instance_to_project(self, project_name, instance):
        project = self.get_project(project_name)
        if not project:
            return
        # Overwrite if instance with same name exists
        instances = project.setdefault("instances", [])
        idx = next((i for i, inst in enumerate(instances) if inst["instance_name"] == instance["instance_name"]), None)
        if idx is not None:
            instances[idx] = instance
        else:
            instances.append(instance)
        self.save_all_configs()
        self.update_instance_dropdown()
        self.instance_dropdown.set(instance["instance_name"])

    def save_current_instance(self):
        name = self.instance_dropdown.get().strip()
        if name and name not in self.instance_dropdown['values']:
            # Add new instance name to dropdown
            values = list(self.instance_dropdown['values'])
            values.append(name)
            self.instance_dropdown['values'] = values
            self.instance_dropdown.set(name)
            
    def get_project(self, name):
        for proj in self.configs.get("projects", []):
            if proj["name"] == name:
                return proj
        return None

    def save_all_configs(self):
        with open(self.CONFIG_FILE, "w") as f:
            json.dump(self.configs, f, indent=2)

    def load_all_configs(self):
        if os.path.exists(self.CONFIG_FILE):
            with open(self.CONFIG_FILE, "r") as f:
                return json.load(f)
        return {"projects": []}
    
    def on_closing(self):
        self.save_all_configs()
        self.root.destroy()

    def compare_to_dialog(self):
        # Get the current instance name
        instance1 = self.instance_dropdown.get().strip()
        if not instance1:
            messagebox.showerror("Error", "Please select the first instance from the main dropdown.")
            return

        # Get available instances (excluding the current one)
        instances = list(self.instance_dropdown['values'])
        if instance1 in instances:
            instances.remove(instance1)
        if not instances:
            messagebox.showerror("Error", "No other instances available to compare.")
            return

        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Instance to Compare To")
        tk.Label(dialog, text="Compare to instance:").grid(row=0, column=0, padx=10, pady=10)
        compare_var = tk.StringVar()
        compare_dropdown = ttk.Combobox(dialog, textvariable=compare_var, values=instances, state="readonly", width=40)
        compare_dropdown.grid(row=0, column=1, padx=10, pady=10)
        compare_dropdown.current(0)

        def do_compare():
            instance2 = compare_var.get().strip()
            if not instance2:
                messagebox.showerror("Error", "Please select an instance to compare to.")
                return
            tester = PowerBIRegressionTester.for_compare_only(self.project_folder_var.get())
            project = self.get_project(self.project_folder_var.get())
            ignore_list = project.get("queriesToIgnore", []) if project else []

            if not tester.instance_exists(instance1) or not tester.instance_exists(instance2):
                messagebox.showerror("Error", "One or both selected instances do not exist.")
                return
            df = tester.compare_instances(instance1, instance2, ignore_list)
            if df is not None and not df.empty:
                self.show_result(df)
            else:
                messagebox.showinfo("Compare", "No differences found.")
            dialog.destroy()

        tk.Button(dialog, text="Compare", command=do_compare, bg="orange").grid(row=1, column=0, columnspan=2, pady=10)
        dialog.grab_set()
        dialog.transient(self.root)
        dialog.wait_window()

    def show_result(self, df):
        if df is not None and not df.empty:
            result_window = tk.Toplevel(root)
            result_window.title(f"Result Table ({len(df)} rows)")

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
            menu.add_command(label="Ignore Query", command=lambda: ignore_query())
            menu.add_command(label="Run Query on Both Instances", command=lambda: run_query_on_both())

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
                        col_name = df.columns[col_idx]
                        # Enable/disable Ignore Query menu item
                        if col_name == "ID":
                            menu.entryconfig("Ignore Query", state="normal")
                        else:
                            menu.entryconfig("Ignore Query", state="disabled")
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

            def ignore_query():
                if hasattr(tree, 'clicked_cell'):
                    row_idx, col_idx = tree.clicked_cell
                    col_name = df.columns[col_idx]
                    if col_name == "ID":
                        query_id = cell_values.get((row_idx, col_name), None)
                        if query_id:
                            # Find the current project
                            project_name = self.project_folder_var.get()
                            project = self.get_project(project_name)
                            if project is not None:
                                ignore_list = project.setdefault("queriesToIgnore", [])
                                if query_id not in ignore_list:
                                    ignore_list.append(query_id)
                                    self.save_all_configs()
                                    messagebox.showinfo("Ignored", f"Query ID '{query_id}' added to ignore list for project '{project_name}'.")
                                else:
                                    messagebox.showinfo("Ignored", f"Query ID '{query_id}' is already in the ignore list.")            

            def run_query_on_both():
                if hasattr(tree, 'clicked_cell'):
                    row_idx, col_idx = tree.clicked_cell
                    col_name = df.columns[col_idx]
                    # Only allow if the column is "Query"
                    if col_name.lower() == "query":
                        query_text = cell_values.get((row_idx, col_name), None)
                        # Get instance names from the selected rows
                        id_col = "ID" if "ID" in df.columns else "id"
                        instance1 = self.instance_dropdown.get().strip()
                        # Prompt user for second instance
                        project = self.get_project(self.project_folder_var.get())
                        instances = [inst["instance_name"] for inst in project.get("instances", []) if inst["instance_name"] != instance1]
                        if not instances:
                            messagebox.showerror("Error", "No other instances available to run the query.")
                            return
                        # Simple dialog to select second instance
                        dialog = tk.Toplevel(result_window)
                        dialog.title("Select Second Instance")
                        tk.Label(dialog, text="Second Instance:").grid(row=0, column=0, padx=10, pady=10)
                        compare_var = tk.StringVar()
                        compare_dropdown = ttk.Combobox(dialog, textvariable=compare_var, values=instances, state="readonly", width=40)
                        compare_dropdown.grid(row=0, column=1, padx=10, pady=10)
                        compare_dropdown.current(0)
                        
                        def do_run():
                            instance2 = compare_var.get().strip()
                            dialog.destroy()
                            # Run query on both instances
                            tester1 = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, next(inst for inst in project["instances"] if inst["instance_name"] == instance1))
                            tester2 = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, next(inst for inst in project["instances"] if inst["instance_name"] == instance2))
                            df1 = tester1.run_single_query(query_text)
                            df2 = tester2.run_single_query(query_text)
                            # Show results in new window
                            compare_window = tk.Toplevel(self.root)
                            compare_window.title(f"Query Results: {instance1} vs {instance2}")
                            # Instance 1 grid
                            tk.Label(compare_window, text=f"{instance1} Result").pack()
                            frame1 = ttk.Frame(compare_window)
                            frame1.pack(fill='both', expand=True)
                            tree1 = ttk.Treeview(frame1, columns=list(df1.columns), show='headings')
                            tree1.pack(fill='both', expand=True)
                            for col in df1.columns:
                                tree1.heading(col, text=col)
                                tree1.column(col, width=120)
                            for _, row in df1.iterrows():
                                tree1.insert('', 'end', values=list(row))
                            # Instance 2 grid
                            tk.Label(compare_window, text=f"{instance2} Result").pack()
                            frame2 = ttk.Frame(compare_window)
                            frame2.pack(fill='both', expand=True)
                            tree2 = ttk.Treeview(frame2, columns=list(df2.columns), show='headings')
                            tree2.pack(fill='both', expand=True)
                            for col in df2.columns:
                                tree2.heading(col, text=col)
                                tree2.column(col, width=120)
                            for _, row in df2.iterrows():
                                tree2.insert('', 'end', values=list(row))
                        tk.Button(dialog, text="Run", command=do_run).grid(row=1, column=0, columnspan=2, pady=10)
                        dialog.grab_set()
                        dialog.transient(result_window)
                        dialog.wait_window()
                        
            tree.bind("<Double-1>", on_double_click)
    
    def view_baseline(self):
        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return
        instance = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == "baseline"), None)
        if not instance:
            messagebox.showerror("Error", "No baseline exists for this project.")
            return

        # if instance.get("interactive"):
        #     conn_str = None  # Or handle interactive logic in your tester
        # else:
        #     conn_str = (
        #         f"Provider=MSOLAP;Data Source={instance['server_name']};"
        #         f"Initial Catalog={instance['database_name']};"
        #         f"User ID={instance['user_id']};Password={self.decrypt_for_user(instance['password'])}"
        #     )

        # tester = PowerBIRegressionTester(
        #     self.project_folder_var.get(),
        #     conn_str,
        #     self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        # )

        tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance)

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

        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return

        instance = next((inst for inst in project.get("instances", []) if inst["instance_name"] == instance_name), None)
        if not instance:
            messagebox.showerror("Error", f"Instance '{instance_name}' does not exist.")
            return

        # if instance.get("interactive"):
        #     conn_str = None  # Or handle interactive logic in your tester
        # else:
        #     conn_str = (
        #         f"Provider=MSOLAP;Data Source={instance['server_name']};"
        #         f"Initial Catalog={instance['database_name']};"
        #         f"User ID={instance['user_id']};Password={self.decrypt_for_user(instance['password'])}"
        #     )

        # tester = PowerBIRegressionTester(
        #     self.project_folder_var.get(),
        #     conn_str,
        #     self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        # )

        tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance)

        if not tester.instance_exists(instance_name):
            messagebox.showerror("Error", f"Instance '{instance_name}' does not exist in the file system.")
            return
        df = tester.load_instance_df(instance_name)
        self.show_result(df)
        
    def run_baseline(self):
        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return

        # Find existing baseline instance
        instance = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == "baseline"), None)

        if not instance:
            # No baseline exists, prompt for details
            instance = self.prompt_instance_details("Baseline")
            if not instance:
                return
            self.save_instance_to_project(self.project_folder_var.get(), instance)
        else:
            # Baseline exists, ask to overwrite
            if not messagebox.askyesno("Overwrite?", "Baseline exists. Overwrite?"):
                return
            # Use the existing baseline config (do not prompt again)

        # Build connection string
        # if instance.get("interactive"):
        #     conn_str = None
        # else:
        #     conn_str = (
        #         f"Provider=MSOLAP;Data Source={instance['server_name']};"
        #         f"Initial Catalog={instance['database_name']};"
        #         f"User ID={instance['user_id']};Password={self.decrypt_for_user(instance['password'])}"
        #     )

        # tester = PowerBIRegressionTester(
        #     self.project_folder_var.get(),
        #     conn_str,
        #     self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        # )

        tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance)

        df = tester.run_baseline()
        self.show_result(df)

    def run_instance(self):
        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return

        # Always pre-fill with baseline if available, otherwise empty
        baseline = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == "baseline"), None)
        prefill = {k: v for k, v in baseline.items() if k != "instance_name"} if baseline else {}
        prefill["instance_name"] = ""  # Always leave instance name empty

        # Prompt for instance details, prefilled from baseline if available
        instance_data = self.prompt_instance_details_prefilled(prefill)
        if not instance_data:
            return

        self.save_instance_to_project(self.project_folder_var.get(), instance_data)
        # if instance_data.get("interactive"):
        #     conn_str = None  # Or handle interactive logic in your tester
        # else:
        #     conn_str = (
        #         f"Provider=MSOLAP;Data Source={instance_data['server_name']};"
        #         f"Initial Catalog={instance_data['database_name']};"
        #         f"User ID={instance_data['user_id']};Password={self.decrypt_for_user(instance_data['password'])}"
        #     )
        # tester = PowerBIRegressionTester(
        #     self.project_folder_var.get(),
        #     conn_str,
        #     self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        # )

        tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance_data)

        # Check if instance already exists
        if tester.instance_exists(instance_data["instance_name"]):
            if not messagebox.askyesno("Overwrite?", f"Instance '{instance_data['instance_name']}' exists. Overwrite?"):
                return
            
        ignore_list = project.get("queriesToIgnore", []) if project else []
        df = tester.run_instance(instance_data["instance_name"], ignore_list)
        if df is not None and not df.empty:
            self.show_result(df)
        else:
            messagebox.showinfo("Compare", "No differences found.")

    def run_compare(self):
        project_folder = self.project_folder_var.get()
        instance_name = self.instance_dropdown.get().strip()
        if not instance_name:
            instance_name = simpledialog.askstring("Instance Name", "Enter an instance name to compare:")
            if not instance_name:
                return
            self.instance_dropdown.set(instance_name)
        tester = PowerBIRegressionTester.for_compare_only(project_folder)
        project = self.get_project(project_folder)
        ignore_list = project.get("queriesToIgnore", []) if project else []

        if tester.instance_exists(instance_name):
            df = tester.compare(instance_name, ignore_list)
            if df is not None and not df.empty:
                self.show_result(df)
            else:
                messagebox.showinfo("Compare", "No differences found.")
        else:
            messagebox.showerror("Error", f"The instance '{instance_name}' does not exist.")

    def create_regession_tester(self, project_folder_var, project_report_folder_var, instance):
        data_source = instance.get('server_name')
        database = instance.get('database_name')
        user_id = instance.get('user_id')
        password = self.decrypt_for_user(instance.get('password', ''))
        interactive = instance.get("interactive", False)
        xmla_endpoint = instance.get("xmla_endpoint", False)
        local_instance = instance.get("local_instance", False)

        tester = PowerBIRegressionTester(
            self.project_folder_var.get(),
            self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else "",
            datasource=data_source,
            database=database,
            user_id=user_id,
            password=password,
            interactive=interactive,
            xmla_endpoint=xmla_endpoint,
            local_instance=local_instance
        )

        return tester

if __name__ == "__main__":
    root = tk.Tk()
    app = PowerBIRegressionTesterApp(root)
    root.mainloop()