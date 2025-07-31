import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import json
import shutil
import base64
import sys
import pandas as pd
import re

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
        # self.setup_widgets()
        self.setup_main_screen()
        self.update_project_folder_dropdown()
        if self.project_folder_dropdown['values']:
            self.load_config_to_fields(self.project_folder_dropdown.get())

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # PowerBIRegressionTester.set_tenant_id('e39cce29-5716-43ba-b27d-1bdd8fd67901')
        # self.ensure_tenant_id()

    # def ensure_tenant_id(self):
    #     # Load config (assuming self.configs is your loaded JSON dict)
    #     tenant_id = self.configs.get("tenant_id", None)
    #     if tenant_id is None:
    #         tenant_id = self.prompt_for_tenant_id()
    #         if tenant_id is not None:  # User didn't cancel
    #             self.configs["tenant_id"] = tenant_id
    #             self.save_all_configs()  # Save back to file
    #     # Set it on your class/static property if needed
    #     PowerBIRegressionTester.set_tenant_id(tenant_id or "")

    # def prompt_for_tenant_id(self):
    #     import tkinter as tk
    #     from tkinter import simpledialog, messagebox

    #     dialog = tk.Toplevel(self.root)
    #     dialog.title("Enter Tenant ID")
    #     dialog.grab_set()
    #     dialog.resizable(False, False)
    #     dialog.geometry("400x150")

    #     tk.Label(dialog, text="Please enter your Azure Tenant ID (can be blank):").pack(pady=10)
    #     tenant_var = tk.StringVar()
    #     entry = tk.Entry(dialog, textvariable=tenant_var, width=40)
    #     entry.pack(pady=5)
    #     entry.focus_set()

    #     result = {"value": None}

    #     def on_ok():
    #         result["value"] = tenant_var.get().strip()
    #         dialog.destroy()

    #     def on_cancel():
    #         dialog.destroy()

    #     btn_frame = tk.Frame(dialog)
    #     btn_frame.pack(pady=10)
    #     tk.Button(btn_frame, text="OK", width=10, command=on_ok).pack(side="left", padx=5)
    #     tk.Button(btn_frame, text="Cancel", width=10, command=on_cancel).pack(side="left", padx=5)

    #     dialog.wait_window()
    #     return result["value"]

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

    def setup_main_screen(self):
        # Clear existing widgets if needed
        for widget in self.root.winfo_children():
            widget.destroy()

        # --- Project Selection Frame ---
        project_frame = ttk.LabelFrame(self.root, text="Project")
        project_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        project_frame.columnconfigure(1, weight=1)

        ttk.Label(project_frame, text="Project:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.project_folder_dropdown = ttk.Combobox(project_frame, textvariable=self.project_folder_var, width=40, state="readonly")
        self.project_folder_dropdown.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.project_folder_dropdown.bind("<<ComboboxSelected>>", self.on_project_folder_select)
        ttk.Button(project_frame, text="New", command=self.create_new_config).grid(row=0, column=2, padx=2, sticky='w')
        ttk.Button(project_frame, text="Save", command=self.save_current_config).grid(row=0, column=3, padx=2, sticky='w')
        ttk.Button(project_frame, text="Delete", command=self.delete_current_config).grid(row=0, column=4, padx=2, sticky='w')

        # --- PBI Report Folder ---
        ttk.Label(project_frame, text="PBI Report Folder (optional):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        entry = ttk.Entry(project_frame, textvariable=self.pbi_report_folder_var, width=60)
        entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(project_frame, text="Browse", command=lambda v=self.pbi_report_folder_var: self.browse_folder(v)).grid(row=1, column=2, sticky='w')

        # --- Instance Selection Frame ---
        instance_frame = ttk.LabelFrame(self.root, text="Instance")
        instance_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        instance_frame.columnconfigure(1, weight=1)

        ttk.Label(instance_frame, text="Instance:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.instance_dropdown = ttk.Combobox(instance_frame, textvariable=self.instance_name_var, width=40, state="readonly")
        self.instance_dropdown.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(instance_frame, text="Create", command=self.create_instance).grid(row=0, column=2, padx=2, sticky='w')
        ttk.Button(instance_frame, text="Edit", command=self.edit_selected_instance).grid(row=0, column=3, padx=2, sticky='w')
        ttk.Button(instance_frame, text="Delete", command=self.delete_current_instance).grid(row=0, column=4, padx=2, sticky='w')
        ttk.Button(instance_frame, text="View", command=self.view_instance).grid(row=0, column=5, padx=2, sticky='w')

        # --- Action Buttons Frame ---
        action_frame = ttk.LabelFrame(self.root, text="Actions")
        action_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        action_frame.columnconfigure(0, weight=1)

        ttk.Button(action_frame, text="Create Baseline", command=self.run_baseline).grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Button(action_frame, text="Edit Baseline", command=self.edit_baseline).grid(row=0, column=1, padx=5, pady=5, sticky='w')
        ttk.Button(action_frame, text="View Baseline", command=self.view_baseline).grid(row=0, column=2, padx=5, pady=5, sticky='w')
        ttk.Button(action_frame, text="Compare", command=self.run_compare).grid(row=0, column=3, padx=5, pady=5, sticky='w')
        ttk.Button(action_frame, text="Compare Instance To", command=self.compare_to_dialog).grid(row=0, column=4, padx=5, pady=5, sticky='w')
        ttk.Button(action_frame, text="Update Selected Instance", command=self.run_selected_instance).grid(row=0, column=5, padx=5, pady=5, sticky='w')

        # --- Status Bar ---
        # self.status_var = tk.StringVar()
        # status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        # status_bar.grid(row=99, column=0, sticky="ew", padx=0, pady=0)

        # Make the main window expand columns
        self.root.columnconfigure(0, weight=1)

    # def setup_widgets(self):
    #     tk.Label(self.root, text="Project:").grid(row=0, column=0, sticky='e')
    #     self.project_folder_dropdown = ttk.Combobox(self.root, textvariable=self.project_folder_var, width=40, state="readonly")
    #     self.project_folder_dropdown.grid(row=0, column=1, padx=5, pady=2, sticky='w')
    #     self.project_folder_dropdown.bind("<<ComboboxSelected>>", self.on_project_folder_select)
    #     tk.Button(self.root, text="Save", command=self.save_current_config).grid(row=0, column=2, padx=2, sticky='w')
    #     tk.Button(self.root, text="Delete", command=self.delete_current_config).grid(row=0, column=3, padx=2, sticky='w')
    #     tk.Button(self.root, text="New", command=self.create_new_config).grid(row=0, column=4, padx=2, sticky='w')

    #     # Project folder field
    #     tk.Label(self.root, text="PBI Report Folder (optional)").grid(row=1, column=0, sticky='e')
    #     entry = tk.Entry(self.root, textvariable=self.pbi_report_folder_var, width=60)
    #     entry.grid(row=1, column=1, padx=5, pady=2)
    #     tk.Button(self.root, text="Browse", command=lambda v=self.pbi_report_folder_var: self.browse_folder(v)).grid(row=1, column=2, sticky='w')

    #     # Instance dropdown
    #     tk.Label(self.root, text="Instance Name").grid(row=2, column=0, sticky='e')
    #     self.instance_dropdown = ttk.Combobox(self.root, textvariable=self.instance_name_var, width=40, state="readonly")
    #     self.instance_dropdown.grid(row=2, column=1, padx=5, pady=2, sticky='w')
    #     tk.Button(self.root, text="Delete", command=self.delete_current_instance, fg="red").grid(row=2, column=2, padx=2, sticky='w')
    #     tk.Button(self.root, text="View Instance", command=self.view_instance, bg="white").grid(row=2, column=3, padx=2, sticky='w')
    #     tk.Button(self.root, text="Create Instance", command=self.create_instance, bg="lightgreen").grid(row=2, column=4, padx=2, sticky='w')

    #     # Action buttons
    #     tk.Button(self.root, text="Create Baseline", command=self.run_baseline, bg="lightblue").grid(row=3, column=0, pady=10)
    #     tk.Button(self.root, text="View Baseline", command=self.view_baseline, bg="white").grid(row=3, column=1, pady=10, sticky='w')
    #     tk.Button(self.root, text="Compare", command=self.run_compare, bg="orange").grid(row=3, column=3, pady=10, sticky='w')
    #     tk.Button(self.root, text="Compare Instance To", command=self.compare_to_dialog, bg="orange").grid(row=3, column=4, padx=2, sticky='w')

    #     tk.Button(self.root, text="Edit Baseline", command=self.edit_baseline, bg="lightyellow").grid(row=4, column=0, pady=5)
    #     tk.Button(self.root, text="Edit Instance", command=self.edit_selected_instance, bg="lightyellow").grid(row=4, column=1, pady=5)

    #     tk.Button(self.root, text="Update Selected Instance", command=self.run_selected_instance, bg="lightgreen").grid(row=4, column=2, pady=5)

    #     self.project_folder_var.trace_add("write", lambda *args: self.update_instance_dropdown())

    def browse_folder(self, var):
        folder = filedialog.askdirectory()
        if folder:
            var.set(folder)
            
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

    # def update_config_dropdown(self):
    #     config_names = list(self.configs.keys())
    #     self.config_dropdown['values'] = config_names
    #     if config_names:
    #         self.config_dropdown.current(0)
    #     else:
    #         self.config_dropdown.set("")

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

        PowerBIRegressionTester.create_project_skeleton(new_name)

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

    # def on_config_select(self, event=None):
    #     name = self.config_dropdown.get()
    #     if name in self.configs:
    #         self.load_config_to_fields(name)

    def on_project_folder_select(self, event=None):
        project_folder = self.project_folder_var.get()
        # if project_folder in self.configs:
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
            # --- Start: Custom dialog with checkbox ---
            dialog = tk.Toplevel(self.root)
            dialog.title("Delete Instance")
            dialog.grab_set()
            tk.Label(dialog, text=f"Are you sure you want to delete instance '{instance_name}'?\nThis cannot be undone.").pack(padx=20, pady=10)
            delete_folder_var = tk.BooleanVar()
            chk = tk.Checkbutton(dialog, text="Also delete instance folder structure on disk", variable=delete_folder_var)
            chk.pack(pady=5)
            btn_frame = tk.Frame(dialog)
            btn_frame.pack(pady=10)
            def do_delete():
                del project["instances"][idx]
                self.save_all_configs()
                self.update_instance_dropdown()
                self.instance_dropdown.set("")
                dialog.destroy()
                # Delete folder structure if checked
                if delete_folder_var.get():
                    # Example: assume folder is under project_folder/instance_name
                    folder_path = os.path.join(os.getcwd(), 'Projects', self.project_folder_var.get(), 'instance', instance_name)
                    if os.path.exists(folder_path):
                        try:
                            shutil.rmtree(folder_path)
                            messagebox.showinfo("Deleted", f"Instance '{instance_name}' and its folder were deleted.")
                        except Exception as e:
                            messagebox.showwarning("Warning", f"Instance deleted, but failed to delete folder:\n{e}")
                    else:
                        messagebox.showinfo("Deleted", f"Instance '{instance_name}' deleted (no folder found).")
                else:
                    messagebox.showinfo("Deleted", f"Instance '{instance_name}' deleted.")
            def do_cancel():
                dialog.destroy()
            tk.Button(btn_frame, text="Delete", command=do_delete, fg="red").pack(side="left", padx=10)
            tk.Button(btn_frame, text="Cancel", command=do_cancel).pack(side="left", padx=10)
            dialog.wait_window()
            # --- End: Custom dialog with checkbox ---

    def prompt_instance_details(self, instance=None):
        instance = instance or {}
        dialog = tk.Toplevel(self.root)
        dialog.title(f"{instance.get('instance_name', 'Baseline')} Details")
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.geometry("470x360")  # Wider window

        # --- Menu Bar ---
        menu_bar = tk.Menu(dialog)
        template_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Template", menu=template_menu)

        def apply_pro_template():
            server_name_var.set("(blank)")
            database_name_var.set("WorkspaceGUID")
            user_id_var.set("")
            password_var.set("")
            tenant_id_var.set("TenantID")
            interactive_var.set(True)
            xmla_endpoint_var.set(False)
            local_instance_var.set(False)

        def apply_xmla_template():
            server_name_var.set("powerbi://api.powerbi.com/v1.0/myorg/YourWorkspace")
            database_name_var.set("ModelName")
            user_id_var.set("AppID")
            password_var.set("YourSecret")
            tenant_id_var.set("TenantID")
            interactive_var.set(False)
            xmla_endpoint_var.set(True)
            local_instance_var.set(False)

        def apply_local_template():
            server_name_var.set("localhost:PortNumber")
            database_name_var.set("ModelGUID")
            user_id_var.set("")
            password_var.set("")
            tenant_id_var.set("TenantID")
            interactive_var.set(False)
            xmla_endpoint_var.set(False)
            local_instance_var.set(True)

        template_menu.add_command(label="Power BI Pro Template", command=apply_pro_template)
        template_menu.add_command(label="XMLA Endpoint Template", command=apply_xmla_template)
        template_menu.add_command(label="Local Instance Template", command=apply_local_template)

        dialog.config(menu=menu_bar)

        # Variables
        instance_name_var = tk.StringVar(value=instance.get("instance_name", "Baseline"))
        server_name_var = tk.StringVar(value=instance.get("server_name", ""))
        database_name_var = tk.StringVar(value=instance.get("database_name", ""))
        user_id_var = tk.StringVar(value=instance.get("user_id", ""))
        password_var = tk.StringVar(value=self.decrypt_for_user(instance.get("password", "")))
        tenant_id_var = tk.StringVar(value=instance.get("tenant_id", ""))  # NEW
        interactive_var = tk.BooleanVar(value=instance.get("interactive", False))
        xmla_endpoint_var = tk.BooleanVar(value=instance.get("xmla_endpoint", False))
        local_instance_var = tk.BooleanVar(value=instance.get("local_instance", False))

        # --- Main Frame ---
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill="both", expand=True)

        # --- Details Frame ---
        details = ttk.LabelFrame(main_frame, text="Instance Details", padding=(10, 10))
        details.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        details.columnconfigure(1, weight=1)

        entry_width = 48  # Wider entry fields

        ttk.Label(details, text="Instance Name:").grid(row=0, column=0, sticky="e", padx=5, pady=4)
        name_entry = ttk.Entry(details, textvariable=instance_name_var, width=entry_width)
        name_entry.config(state="normal" if instance.get("instance_name", "") == "" else "disabled")
        name_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=4)
        help_btn = tk.Label(details, text="?", foreground="blue", cursor="question_arrow")
        help_btn.grid(row=0, column=2, sticky="w", padx=2)
        ToolTip(help_btn, "The name of this instance. Cannot be changed for Baseline.")

        ttk.Label(details, text="Server Name:").grid(row=1, column=0, sticky="e", padx=5, pady=4)
        server_entry = ttk.Entry(details, textvariable=server_name_var, width=entry_width)
        server_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=4)
        help_btn = tk.Label(details, text="?", foreground="blue", cursor="question_arrow")
        help_btn.grid(row=1, column=2, sticky="w", padx=2)
        ToolTip(help_btn, "The Power BI server address or XMLA endpoint.")

        ttk.Label(details, text="Database Name:").grid(row=2, column=0, sticky="e", padx=5, pady=4)
        db_entry = ttk.Entry(details, textvariable=database_name_var, width=entry_width)
        db_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=4)
        help_btn = tk.Label(details, text="?", foreground="blue", cursor="question_arrow")
        help_btn.grid(row=2, column=2, sticky="w", padx=2)
        ToolTip(help_btn, "If using an XMLA Endpoing, use semantic model name.\n\nIf using Power BI Pro, to use semantic model GUID exposed in the service.\n\nIf using a local instance, use the model GUID.\nThis can be obtained by running a DMV query in DAX Studio ( select * from $SYSTEM.DBSCHEMA_CATALOGS )")

        ttk.Label(details, text="User ID:").grid(row=3, column=0, sticky="e", padx=5, pady=4)
        user_entry = ttk.Entry(details, textvariable=user_id_var, width=entry_width)
        user_entry.grid(row=3, column=1, sticky="ew", padx=5, pady=4)
        help_btn = tk.Label(details, text="?", foreground="blue", cursor="question_arrow")
        help_btn.grid(row=3, column=2, sticky="w", padx=2)
        ToolTip(help_btn, "App ID when using an XMLA Endpoint (leave blank for interactive).")

        ttk.Label(details, text="Password:").grid(row=4, column=0, sticky="e", padx=5, pady=4)
        pass_entry = ttk.Entry(details, textvariable=password_var, width=entry_width, show="*")
        pass_entry.grid(row=4, column=1, sticky="ew", padx=5, pady=4)
        help_btn = tk.Label(details, text="?", foreground="blue", cursor="question_arrow")
        help_btn.grid(row=4, column=2, sticky="w", padx=2)
        ToolTip(help_btn, "Secret when using an XMLA Endpoint (leave blank for interactive).")

        # --- Tenant ID field ---
        ttk.Label(details, text="Tenant ID:").grid(row=5, column=0, sticky="e", padx=5, pady=4)
        tenant_entry = ttk.Entry(details, textvariable=tenant_id_var, width=entry_width)
        tenant_entry.grid(row=5, column=1, sticky="ew", padx=5, pady=4)
        help_btn = tk.Label(details, text="?", foreground="blue", cursor="question_arrow")
        help_btn.grid(row=5, column=2, sticky="w", padx=2)
        ToolTip(help_btn, "Azure Tenant ID (required for interactive authentication).")

        # --- Options Frame ---
        options = ttk.LabelFrame(main_frame, text="Connection Options", padding=(10, 10))
        options.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))
        options.columnconfigure(0, weight=1)
        options.columnconfigure(1, weight=1)
        options.columnconfigure(2, weight=1)

        # Helper function to create a checkbox with a right-aligned tooltip icon
        def add_checkbox_with_tooltip(parent, row, text, variable, tooltip_text):
            frame = ttk.Frame(parent)
            frame.grid(row=0, column=row, sticky="w", padx=2, pady=2)
            chk = ttk.Checkbutton(frame, text=text, variable=variable)
            chk.pack(side="left")
            help_btn = tk.Label(frame, text="?", foreground="blue", cursor="question_arrow")
            help_btn.pack(side="left", padx=(2, 0))  # Small space between checkbox and icon
            ToolTip(help_btn, tooltip_text)
            return chk

        # Usage:
        interactive_chk = add_checkbox_with_tooltip(
            options, 0, "Interactive", interactive_var,
            "Enable Interactive authentication to the Power BI Service (XMLA Endpoints and Pro Workspace)."
        )
        xmla_chk = add_checkbox_with_tooltip(
            options, 1, "XMLA Endpoint", xmla_endpoint_var,
            "Connect using an XMLA endpoint."
        )
        local_chk = add_checkbox_with_tooltip(
            options, 2, "Local Instance", local_instance_var,
            "Connect to a local Semantic Model instance (requires local model GUID as the Database Name)."
        )

        # Configure columns for even distribution
        # for i in range(3):
        #     options.columnconfigure(i, weight=1)

        # interactive_chk = ttk.Checkbutton(options, text="Interactive", variable=interactive_var)
        # interactive_chk.grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        # help_btn = tk.Label(options, text="?", foreground="blue", cursor="question_arrow")
        # help_btn.grid(row=0, column=1, sticky="w", padx=2)
        # ToolTip(help_btn, "Enable Interactive authentication to the Power BI Service (XMLA Endpoints and Pro Workspace).")

        # xmla_chk = ttk.Checkbutton(options, text="XMLA Endpoint", variable=xmla_endpoint_var)
        # xmla_chk.grid(row=0, column=2, sticky="ew", padx=2, pady=2)
        # help_btn = tk.Label(options, text="?", foreground="blue", cursor="question_arrow")
        # help_btn.grid(row=0, column=3, sticky="w", padx=2)
        # ToolTip(help_btn, "Connect using an XMLA endpoint.")

        # local_chk = ttk.Checkbutton(options, text="Local Instance", variable=local_instance_var)
        # local_chk.grid(row=0, column=4, sticky="ew", padx=2, pady=2)
        # help_btn_local = tk.Label(options, text="?", foreground="blue", cursor="question_arrow")
        # help_btn_local.grid(row=0, column=5, sticky="w", padx=2)
        # ToolTip(help_btn_local, "Connect to a local Semantic Model instance (requires local model GUID as the Database Name).")

        # --- Field Enable/Disable Logic ---
        def toggle_fields(*args):
            user_entry.config(state="disabled" if interactive_var.get() or local_instance_var.get() else "normal")
            pass_entry.config(state="disabled" if interactive_var.get() or local_instance_var.get() else "normal")
            tenant_entry.config(state="disabled" if local_instance_var.get() else "normal")  # Disable for Local Instance
            server_entry.config(state="normal")
            db_entry.config(state="normal")
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

        # --- Buttons Frame ---
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, sticky="e", pady=(10, 0))
        result = {}

        def on_ok():
            result["instance_name"] = instance_name_var.get().strip()
            result["interactive"] = interactive_var.get()
            if not result["instance_name"]:
                messagebox.showerror("Error", "Instance Name is required.", parent=dialog)
                return
            if not (interactive_var.get() or xmla_endpoint_var.get() or local_instance_var.get()):
                messagebox.showerror("Error", "Either 'Interactive', 'XMLA Endpoint', or 'Local Instance' must be checked.", parent=dialog)
                return
            result["server_name"] = server_name_var.get().strip()
            result["database_name"] = database_name_var.get().strip()
            result["user_id"] = user_id_var.get().strip()
            result["password"] = self.encrypt_for_user(password_var.get())
            result["tenant_id"] = tenant_id_var.get().strip()  # NEW
            result["xmla_endpoint"] = xmla_endpoint_var.get()
            result["local_instance"] = local_instance_var.get()

            if result["local_instance"]:
                if not all([result["server_name"], result["database_name"]]):
                    messagebox.showerror("Error", "Server Name and Database Name are required for a Local Instance.", parent=dialog)
                    return
            else:
                if result["interactive"]:
                    if not all([result["tenant_id"]]):
                        messagebox.showerror("Error", "Tenant ID is required for interactive authentication.", parent=dialog)
                        return
                    
                    guid_regex = re.compile(
                        r'^[{(]?[0-9a-fA-F]{8}-'
                        r'[0-9a-fA-F]{4}-'
                        r'[0-9a-fA-F]{4}-'
                        r'[0-9a-fA-F]{4}-'
                        r'[0-9a-fA-F]{12}[)}]?$'
                    )
                    if not isinstance(result["tenant_id"], str) or not guid_regex.match(result["tenant_id"]):
                        messagebox.showerror("Error", f"Tenant ID '{result['tenant_id']}' is not a valid GUID.", parent=dialog)
                        return
                else:
                    if result["xmla_endpoint"]:
                        if not all([result["server_name"], result["database_name"], result["user_id"], result["password"], result["tenant_id"]]):
                            messagebox.showerror("Error", "Server Name, Database Name, User ID, Password and Tenant ID are required for XMLA Endpoint.", parent=dialog)
                            return

                # if not result["interactive"]:
                #     result["user_id"] = user_id_var.get().strip()
                #     result["password"] = self.encrypt_for_user(password_var.get())
                #     if not all([result["server_name"], result["database_name"]]):
                #         messagebox.showerror("Error", "User ID and Password are required unless using Interactive.", parent=dialog)
                #         return
                # elif result["interactive"]:
                #     if not all([result["tenant_id"]]):
                #         messagebox.showerror("Error", "Tenant ID is required for interactive authentication.", parent=dialog)
                #         return

            dialog.destroy()

        def on_cancel():
            result.clear()
            dialog.destroy()

        ttk.Button(btn_frame, text="OK", command=on_ok).grid(row=0, column=0, padx=10)
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).grid(row=0, column=1, padx=10)

        dialog.wait_window()
        return result if result else None

    def edit_selected_instance(self):
        instance_name = self.instance_dropdown.get().strip()
        self.edit_instance(instance_name)

        # project = self.get_project(self.project_folder_var.get())
        # instance_name = self.instance_dropdown.get().strip()


        # if not project or not instance_name:
        #     messagebox.showerror("Error", "No instance selected to edit.")
        #     return
        # instance = next((inst for inst in project.get("instances", []) if inst["instance_name"] == instance_name), None)
        # if not instance:
        #     messagebox.showerror("Error", f"Instance '{instance_name}' not found.")
        #     return

        # # Prompt with current values
        # edited = self.prompt_instance_details(instance)
        # if not edited:
        #     return
        # # Overwrite the instance
        # idx = next((i for i, inst in enumerate(project["instances"]) if inst["instance_name"] == instance_name), None)
        # if idx is not None:
        #     project["instances"][idx] = edited
        #     self.save_all_configs()
        #     self.update_instance_dropdown()
        #     self.instance_dropdown.set(edited["instance_name"])
        #     messagebox.showinfo("Saved", f"Instance '{edited['instance_name']}' updated.")

    def edit_baseline(self):
        self.edit_instance("Baseline")
        # # For baseline, you may want to use a convention (e.g., instance_name == "Baseline")
        # project = self.get_project(self.project_folder_var.get())
        # if not project:
        #     messagebox.showerror("Error", "No project selected.")
        #     return
        # instance = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == "baseline"), None)
        # if not instance:
        #     messagebox.showerror("Error", "No baseline instance found.")
        #     return
        # edited = self.prompt_instance_details(instance)
        # if not edited:
        #     return
        # idx = next((i for i, inst in enumerate(project["instances"]) if inst["instance_name"].lower() == "baseline"), None)
        # if idx is not None:
        #     project["instances"][idx] = edited
        #     self.save_all_configs()
        #     self.update_instance_dropdown()
        #     if edited["instance_name"] != "Baseline":
        #         self.instance_dropdown.set(edited["instance_name"])
        #     messagebox.showinfo("Saved", "Baseline instance updated.")

    def edit_instance(self, instance_name):
        # For baseline, you may want to use a convention (e.g., instance_name == "Baseline")
        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return
        # instance = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == "baseline"), None)
        instance = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == instance_name.lower()), None)
        if not instance:
            messagebox.showerror("Error", f"{instance_name} not found.")
            return
        edited = self.prompt_instance_details(instance)
        if not edited:
            return
        # idx = next((i for i, inst in enumerate(project["instances"]) if inst["instance_name"].lower() == "baseline"), None)
        idx = next((i for i, inst in enumerate(project["instances"]) if inst["instance_name"].lower() == instance_name.lower()), None)
        if idx is not None:
            project["instances"][idx] = edited
            self.save_all_configs()
            self.update_instance_dropdown()
            if edited["instance_name"].lower() != "baseline":
                self.instance_dropdown.set(edited["instance_name"])
            #messagebox.showinfo("Saved", "Baseline instance updated.")
            messagebox.showinfo("Saved", f"'{edited['instance_name']}' updated.")

    # def prompt_instance_details_prefilled(self, instance):
    #     dialog = tk.Toplevel(self.root)
    #     dialog.title("Edit Instance Details")
    #     dialog.grab_set()
    #     dialog.resizable(False, False)

    #     # Variables
    #     instance_name_var = tk.StringVar(value=instance.get("instance_name", ""))
    #     server_name_var = tk.StringVar(value=instance.get("server_name", ""))
    #     database_name_var = tk.StringVar(value=instance.get("database_name", ""))
    #     user_id_var = tk.StringVar(value=instance.get("user_id", ""))
    #     password_var = tk.StringVar(value=self.decrypt_for_user(instance.get("password", "")))
    #     interactive_var = tk.BooleanVar(value=instance.get("interactive", False))
    #     xmla_endpoint_var = tk.BooleanVar(value=instance.get("xmla_endpoint", False) if "instance" in locals() else False)
    #     local_instance_var = tk.BooleanVar(value=instance.get("local_instance", False) if "instance" in locals() else False)

    #     # Layout
    #     tk.Label(dialog, text="Instance Name:").grid(row=0, column=0, sticky="e")
    #     name_entry = tk.Entry(dialog, textvariable=instance_name_var, width=40)
    #     name_entry.config(state="disabled" if instance.get("instance_name", "") == 'Baseline' else "normal")
    #     name_entry.grid(row=0, column=1, padx=5, pady=2)

    #     tk.Label(dialog, text="Server Name:").grid(row=1, column=0, sticky="e")
    #     server_entry = tk.Entry(dialog, textvariable=server_name_var, width=40)
    #     server_entry.grid(row=1, column=1, padx=5, pady=2)

    #     tk.Label(dialog, text="Database Name:").grid(row=2, column=0, sticky="e")
    #     db_entry = tk.Entry(dialog, textvariable=database_name_var, width=40)
    #     db_entry.grid(row=2, column=1, padx=5, pady=2)

    #     tk.Label(dialog, text="User ID:").grid(row=3, column=0, sticky="e")
    #     user_entry = tk.Entry(dialog, textvariable=user_id_var, width=40)
    #     user_entry.grid(row=3, column=1, padx=5, pady=2)

    #     tk.Label(dialog, text="Password:").grid(row=4, column=0, sticky="e")
    #     pass_entry = tk.Entry(dialog, textvariable=password_var, width=40, show="*")
    #     pass_entry.grid(row=4, column=1, padx=5, pady=2)

    #     interactive_chk = tk.Checkbutton(dialog, text="Interactive", variable=interactive_var)
    #     interactive_chk.grid(row=5, column=0, columnspan=2, pady=5)

    #     xmla_chk = tk.Checkbutton(dialog, text="XMLA Endpoint", variable=xmla_endpoint_var)
    #     xmla_chk.grid(row=6, column=0, columnspan=2, pady=5)

    #     local_chk = tk.Checkbutton(dialog, text="Local Instance", variable=local_instance_var)
    #     local_chk.grid(row=7, column=0, columnspan=2, pady=5)

    #     # Disable/enable fields based on checkbox
    #     def toggle_fields(*args):
    #         # Only disable User ID and Password if Interactive is checked
    #         user_entry.config(state="disabled" if interactive_var.get() else "normal")
    #         pass_entry.config(state="disabled" if interactive_var.get() else "normal")
    #         # Server Name and Database Name always enabled
    #         server_entry.config(state="normal")
    #         db_entry.config(state="normal")

    #         # Disable/clear XMLA and Interactive if Local Instance is checked
    #         if local_instance_var.get():
    #             xmla_endpoint_var.set(False)
    #             interactive_var.set(False)
    #             xmla_chk.config(state="disabled")
    #             interactive_chk.config(state="disabled")
    #         else:
    #             xmla_chk.config(state="normal")
    #             interactive_chk.config(state="normal")

    #     interactive_var.trace_add("write", toggle_fields)
    #     local_instance_var.trace_add("write", toggle_fields)
    #     toggle_fields()  # Set initial state

    #     # Buttons
    #     result = {}
    #     def on_ok():
    #         result["instance_name"] = instance_name_var.get().strip()
    #         result["interactive"] = interactive_var.get()
    #         if not result["instance_name"]:
    #             messagebox.showerror("Error", "Instance Name is required.", parent=dialog)
    #             return

    #         if not (interactive_var.get() or xmla_endpoint_var.get() or local_instance_var.get()):
    #             messagebox.showerror("Error", "Either 'Interactive' or 'XMLA Endpoint' must be checked.", parent=dialog)
    #             return
            
    #         result["server_name"] = server_name_var.get().strip()
    #         result["database_name"] = database_name_var.get().strip()
    #         result["xmla_endpoint"] = xmla_endpoint_var.get()
    #         result["local_instance"] = local_instance_var.get()

    #         if not result["interactive"]:
    #             result["user_id"] = user_id_var.get().strip()
    #             result["password"] = self.encrypt_for_user(password_var.get())
    #             if not all([result["server_name"], result["database_name"]]):
    #                 messagebox.showerror("Error", "All fields are required unless Interactive is checked.", parent=dialog)
    #                 return
    #         dialog.destroy()

    #     def on_cancel():
    #         result.clear()
    #         dialog.destroy()

    #     tk.Button(dialog, text="OK", command=on_ok).grid(row=7, column=0, pady=10)
    #     tk.Button(dialog, text="Cancel", command=on_cancel).grid(row=7, column=1, pady=10)

    #     dialog.wait_window()
    #     return result if result else None

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
                self.show_result(df, instance1, instance2)
            else:
                messagebox.showinfo("Compare", "No differences found.")
            dialog.destroy()

        tk.Button(dialog, text="Compare", command=do_compare, bg="orange").grid(row=1, column=0, columnspan=2, pady=10)
        dialog.grab_set()
        dialog.transient(self.root)
        dialog.wait_window()

    def show_result(self, df, instance1=None, instance2=None):
        if df is not None and not df.empty:
            # if "Query Hash" in df.columns and "Query Hash Baseline" in df.columns:
            #     df["Hash Match"] = df["Query Hash"] == df["Query Hash Baseline"]
            #     df["Hash Match"] = df["Hash Match"].replace({True: "Match", False: "Mismatch"})
            diff_count = (df["Hash Match"] == False).sum() if "Hash Match" in df.columns else 0
            total_count = len(df)
            diff_summary = f"Differences found: {diff_count} of {total_count}"

            result_window = tk.Toplevel(self.root)
            result_window.instance1 = instance1
            result_window.instance2 = instance2
            result_window.geometry("1400x700")

            two_instances = instance1 and instance2

            def on_view_all_toggle():
                # Remove all rows from the tree
                for item in tree.get_children():
                    tree.delete(item)
                # Choose which DataFrame to display
                if two_instances and view_all_var.get():
                    display_df = df
                else:
                    display_df = df[df["Hash Match"] == False] if "Hash Match" in df.columns else df

                # Re-insert rows
                for row_idx, row in display_df.iterrows():
                    tag = "no_compare"
                    if two_instances:
                        if "Query Hash" in display_df.columns and "Query Hash Baseline" in display_df.columns:
                            tag = "diff" if row["Query Hash"] != row["Query Hash Baseline"] else "same"
                        elif "Row Hash" in display_df.columns and "Row Hash_baseline" in display_df.columns:
                            tag = "diff" if row["Row Hash"] != row["Row Hash_baseline"] else "same"
                        else:
                            tag = "diff" if len(set([str(row[col]) for col in columns])) > 1 else "same"
                    values = [
                        (
                            str(row[col])[:20] + "..." if col.lower() == "query" and len(str(row[col])) > 20 else str(row[col])
                        ).replace('\n\t', ' ').replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
                        for col in columns
                    ]
                    tree.insert("", "end", iid=row_idx, values=values, tags=(tag,))
                # Optionally update status bar
                # status = "Showing all queries" if view_all_var.get() else "Showing filtered queries"
                # status_var.set(status)


            # --- Header ---
            header_frame = ttk.Frame(result_window)
            header_frame.pack(fill='x', padx=10, pady=10)

            result_window.title(f"Results: {instance1} vs {instance2} ({len(df)} rows)")
            ttk.Label(header_frame, text=f"Instance 1: {instance1}", font=("Arial", 12, "bold")).pack(side='left', padx=10)
            ttk.Label(header_frame, text=f"Instance 2: {instance2}", font=("Arial", 12, "bold")).pack(side='left', padx=10)
            ttk.Label(header_frame, text=f"{diff_summary}", font=("Arial", 12)).pack(side='right', padx=10)
            # Add "View All Queries" checkbox
            view_all_var = tk.BooleanVar(value=False)
            view_all_checkbox = ttk.Checkbutton(header_frame, text="View All Queries", variable=view_all_var, command=lambda: on_view_all_toggle())
            view_all_checkbox.pack(side='right', padx=10)

            if two_instances:
                result_window.title(f"Results: {instance1} vs {instance2} ({len(df)} rows)")
            else:
                for widget in header_frame.winfo_children():
                    if (
                        isinstance(widget, ttk.Checkbutton)
                        or (isinstance(widget, ttk.Label) and "Instance 2" in widget.cget("text"))
                        or (isinstance(widget, ttk.Label) and diff_summary in widget.cget("text"))
                    ):
                        widget.pack_forget()
                result_window.title(f"Results: {instance1} ({len(df)} rows)")

            # --- Treeview Frame ---
            tree_frame = ttk.Frame(result_window)
            tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

            columns = list(df.columns)
            tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=30, selectmode="browse")
            tree.pack(side="left", fill="both", expand=True)

            # Add scrollbars
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            vsb.pack(side="right", fill="y")
            tree.configure(yscrollcommand=vsb.set)

            hsb = ttk.Scrollbar(result_window, orient="horizontal", command=tree.xview)
            hsb.pack(fill="x")
            tree.configure(xscrollcommand=hsb.set)

            # Setup columns
            # Exclude "Hash Match" column from display
            display_columns = [col for col in columns if col != "Hash Match"]
            tree["columns"] = display_columns
            for col in display_columns:
                tree.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree, _col, False))
                tree.column(col, width=180 if col.lower() == "query" else 120, anchor="w", stretch=True)

            # Color tags
            tree.tag_configure("diff", background="#f4d4d4")
            tree.tag_configure("same", background="#d4f4dd")
            tree.tag_configure("no_compare", background="#ffffff", foreground="#000000")

            on_view_all_toggle()
            # Insert data
            # Determine which DataFrame to use based on the "View All Queries" checkbox
            # display_df = df if (not two_instances or view_all_var.get()) else df[df["Hash Match"] == "Mismatch"] if "Hash Match" in df.columns else df

            # for row_idx, row in display_df.iterrows():
            #     tag = "same"
            #     if two_instances:
            #         if "Query Hash" in display_df.columns and "Query Hash Baseline" in display_df.columns:
            #             tag = "diff" if row["Query Hash"] != row["Query Hash Baseline"] else "same"
            #         elif "row_hash" in display_df.columns and "row_hash_baseline" in display_df.columns:
            #             tag = "diff" if row["row_hash"] != row["row_hash_baseline"] else "same"
            #         else:
            #             tag = "diff" if len(set([str(row[col]) for col in columns])) > 1 else "same"

            #     values = [
            #         (
            #             str(row[col])[:20] + "..." if col.lower() == "query" and len(str(row[col])) > 20 else str(row[col])
            #         ).replace('\n\t', ' ').replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
            #         for col in columns
            #     ]
            #     tree.insert("", "end", iid=row_idx, values=values, tags=(tag,))
            # for row_idx, row in df.iterrows():
            #     # Example: color row red if any value differs between instances, else green
            #     tag = "same"
            #     if two_instances:
            #         # If you have specific columns to compare, adjust here
            #         # For now, mark as "diff" if any value in the row is different between instances
            #         # (Assumes columns are aligned for both instances)
            #         if "Query Hash" in df.columns and "Query Hash Baseline" in df.columns:
            #             tag = "diff" if row["Query Hash"] != row["Query Hash Baseline"] else "same"
            #         elif "row_hash" in df.columns and "row_hash_baseline" in df.columns:
            #             tag = "diff" if row["row_hash"] != row["row_hash_baseline"] else "same"
            #         else:
            #             # Fallback: mark as diff if any value is not equal to the first column
            #             # values = [str(row[col]) for col in columns]
            #             tag = "diff" if len(set(values)) > 1 else "same"

            #     values = [
            #         (
            #             str(row[col])[:20] + "..." if col.lower() == "query" and len(str(row[col])) > 20 else str(row[col])
            #         ).replace('\n\t', ' ').replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
            #         for col in columns
            #     ]
            #     # values = [row[col] for col in columns]
            #     tree.insert("", "end", iid=row_idx, values=values, tags=(tag,))

            # --- Context Menu ---
            menu = tk.Menu(result_window, tearoff=0)
            menu.add_command(label="Copy", command=lambda: copy_cell())
            menu.add_command(label="Ignore Query", command=lambda: ignore_query())
            menu.add_command(label="View Query", command=lambda: on_double_click())
            if two_instances:
                menu.add_command(label="Run Query on Both Instances", command=lambda: run_query())
            else:
                menu.add_command(label="Run Query", command=lambda: run_query())

            # Store full values for copy
            cell_values = {}
            for row_idx, row in df.iterrows():
                for col in df.columns:
                    cell_values[(row_idx, col)] = row[col]

            def treeview_sort_column(tree, col, reverse):
                # Get all values in the column and sort them
                l = [(tree.set(k, col), k) for k in tree.get_children('')]
                try:
                    l.sort(key=lambda t: float(t[0]) if t[0].replace('.', '', 1).isdigit() else t[0], reverse=reverse)
                except Exception:
                    l.sort(key=lambda t: t[0], reverse=reverse)
                # Rearrange items in sorted positions
                for index, (val, k) in enumerate(l):
                    tree.move(k, '', index)
                # Reverse sort next time
                tree.heading(col, command=lambda: treeview_sort_column(tree, col, not reverse))
                
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

            tree.bind("<Button-3>", on_right_click)

            def on_double_click():
                if hasattr(tree, 'clicked_cell'):
                    row_idx, col_idx = tree.clicked_cell
                    col_name = df.columns[col_idx]
                    full_text = cell_values.get((row_idx, col_name), "")
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
                    query_id = cell_values.get((row_idx, col_name), None)
                    if query_id:
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
        
            def run_query():
                if hasattr(tree, 'clicked_cell'):
                            
                    row_idx, col_idx = tree.clicked_cell
                    col_name = df.columns[col_idx]
                    # if col_name.lower() == "query":
                    query_text = cell_values.get((row_idx, "Query"), None)
                    # Get the value in the "ID" column for this row
                    id_value = cell_values.get((row_idx, "ID"), None)
                    # Get instance names from window attributes
                    instance1 = getattr(result_window, "instance1", None)
                    instance2 = getattr(result_window, "instance2", None)
                    # Determine if this was a compare (two instances) or just a view (one instance)
                    two_instances = instance1 and instance2


                    def on_right_click_query(event):
                        region = event.widget.identify("region", event.x, event.y)
                        if region == "cell":
                            if tree2_list:
                                tree_x = event.widget
                                tree1 = None
                                tree2 = None

                                if tree_x in tree2_list:
                                    tree2 = tree_x
                                    tree_x_idx = tree2_list.index(tree_x)
                                    tree1 = tree1_list[tree_x_idx] if tree1_list else None
                                else:
                                    tree1 = tree_x
                                    tree_x_idx = tree1_list.index(tree_x)
                                    tree2 = tree2_list[tree_x_idx] if tree2_list else None

                                region = tree_x.identify("region", event.x, event.y)
                                left_frame.tree1 = tree1
                                right_frame.tree2 = tree2
    
                                query_menu.tk_popup(event.x_root, event.y_root)

                    def compare_row():
                        if hasattr(left_frame, 'tree1') and hasattr(right_frame, 'tree2'):
                            tree1 = left_frame.tree1
                            tree2 = right_frame.tree2

                            if tree1 and tree2:
                                row1_values = None              
                                row2_values = None

                                selected1 = tree1.selection()
                                selected2 = tree2.selection()
                                if selected1 and len(selected1) == 1 and selected2 and len(selected2) == 1:
                                    row_id1 = selected1[0]
                                    row1_values = tree1.item(row_id1, "values")
                                    row_id2 = selected2[0]
                                    row2_values = tree2.item(row_id2, "values") 
                                else:
                                    messagebox.showerror("Error", "Please select one row from each tree to compare.")
                                    return

                                if row1_values and row2_values:
                                    columns = [col for col in tree1["columns"] if col not in ("Row Hash", "Result Set Hash")]

                                    # --- Compare Popup Window ---
                                    compare_popup = tk.Toplevel(self.root)
                                    compare_popup.title("Row Comparison")
                                    compare_popup.geometry("1000x600")

                                    # --- Header Frame ---
                                    header_frame = ttk.Frame(compare_popup)
                                    header_frame.pack(fill='x', padx=10, pady=10)
                                    instance1_label = ttk.Label(header_frame, text=f"{instance1} Row", font=("Arial", 12, "bold"))
                                    instance1_label.pack(side='left', padx=10)
                                    instance2_label = ttk.Label(header_frame, text=f"{instance2} Row", font=("Arial", 12, "bold"))
                                    instance2_label.pack(side='left', padx=10)
                                    # ttk.Label(header_frame, text="Column Comparison", font=("Arial", 12)).pack(side='right', padx=10)

                                    # --- Main Frame for Treeview ---
                                    main_frame = ttk.Frame(compare_popup)
                                    main_frame.pack(fill='both', expand=True, padx=10, pady=5)

                                    compare_tree = ttk.Treeview(main_frame, columns=("Column", "Query 1 Value", "Query 2 Value"), show='headings', selectmode="browse")
                                    compare_tree.grid(row=0, column=0, sticky="nsew")

                                    vsb = ttk.Scrollbar(main_frame, orient="vertical", command=compare_tree.yview)
                                    vsb.grid(row=0, column=1, sticky="ns")
                                    compare_tree.configure(yscrollcommand=vsb.set)

                                    hsb = ttk.Scrollbar(main_frame, orient="horizontal", command=compare_tree.xview)
                                    hsb.grid(row=1, column=0, sticky="ew")
                                    compare_tree.configure(xscrollcommand=hsb.set)

                                    main_frame.rowconfigure(0, weight=1)
                                    main_frame.columnconfigure(0, weight=1)

                                    compare_tree.heading("Column", text="Column")
                                    compare_tree.heading("Query 1 Value", text=f"{instance1} Value")
                                    compare_tree.heading("Query 2 Value", text=f"{instance2} Value")
                                    compare_tree.column("Column", width=200, anchor="w")
                                    compare_tree.column("Query 1 Value", width=350, anchor="w")
                                    compare_tree.column("Query 2 Value", width=350, anchor="w")

                                    compare_tree.tag_configure('green_row', background='#d4f4dd')
                                    compare_tree.tag_configure('red_row', background='#f4d4d4')

                                    matched_values = 0
                                    for idx, col in enumerate(columns):
                                        val1 = row1_values[idx] if idx < len(row1_values) else ""
                                        val2 = row2_values[idx] if idx < len(row2_values) else ""
                                        # tag = 'green_row' if val1 == val2 else 'red_row'
                                        if val1 == val2:
                                            matched_values += 1
                                            tag = 'green_row'
                                        else:
                                            tag = 'red_row'
                                        compare_tree.insert('', 'end', values=(col, val1, val2), tags=(tag,))

                                    unmatched_values = len(columns) - matched_values
                                    instance1_label.config(text=f"Equal columns: {matched_values}")
                                    instance2_label.config(text=f"Different columns: {unmatched_values}")


                    project = self.get_project(self.project_folder_var.get())
                    inst1_obj = next((inst for inst in project["instances"] if inst["instance_name"] == instance1), None)
                    inst2_obj = None
                    if two_instances:
                        inst2_obj = next((inst for inst in project["instances"] if inst["instance_name"] == instance2), None)

                    if not inst1_obj:
                        messagebox.showerror("Error", "The first instance config was not found.")
                        return

                    if not inst2_obj and two_instances:
                        messagebox.showerror("Error", "Neither instance config not found.")
                        return
                    
                    # Create testers for each instance
                    tester1 = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, inst1_obj)
                    tester2 = None
                    if two_instances:
                        tester2 = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, inst2_obj)

                    # After you get lists of dataframes from both testers:
                    df_list1 = tester1.run_single_query(query_text)  # List of DataFrames
                    df_list2 = []
                    if two_instances:
                        df_list2 = tester2.run_single_query(query_text)  # List of DataFrames

                    if (not df_list1 or all(df.empty for df in df_list1)) and (not df_list2 or all(df.empty for df in df_list2)):
                        messagebox.showinfo("Query Results", "No results returned for this query on either instance.")
                        return

                    compare_window = tk.Toplevel(self.root)
                    if two_instances:
                        compare_window.title(f"Query Results: {instance1} vs {instance2} for Query ID: {id_value}")
                    else:
                        compare_window.title(f"Query Results: {instance1}")
                    compare_window.geometry("1400x700")

                    notebook = ttk.Notebook(compare_window)
                    notebook.pack(fill='both', expand=True)

                    if two_instances:
                        # Context menu for copying cell value
                        query_menu = tk.Menu(result_window, tearoff=0)
                        query_menu.add_command(label="Compare Rows", command=lambda: compare_row())

                    num_tabs = max(len(df_list1), len(df_list2))
                    tree1_list = []  # Add this before your loop
                    tree2_list = []  # Add this before your loop
                    
                    for idx in range(num_tabs):
                        tab_name = f"Result {idx+1}"
                        frame = ttk.Frame(notebook, width=1200, height=500)
                        frame.pack(fill='both', expand=True)
                        frame.pack_propagate(False)

                        # Check if the df_list1 has this index
                        instance1_df = pd.DataFrame()
                        if idx < len(df_list1):
                            instance1_df = df_list1[idx]

                        # Check if the df_list2 has this index
                        instance2_df = pd.DataFrame()
                        if idx < len(df_list2):
                            instance2_df = df_list2[idx]

                        # Set tab color based on Query Hash comparison
                        tab_color = "black"
                        if not instance1_df.empty and not instance2_df.empty:
                            hash1 = instance1_df.iloc[0].get("Result Set Hash", "")
                            hash2 = instance2_df.iloc[0].get("Result Set Hash", "")
                            if hash1 == hash2:
                                tab_color = "green"
                            else:
                                tab_color = "red"
                        # elif instance1_df.empty and instance2_df.empty:
                        #     tab_color = "green"

                        # Label for each instance
                        label_frame = ttk.Frame(frame)
                        label_frame.pack(fill='x')
                        tk.Label(label_frame, text=f"{instance1} ({len(instance1_df)})", font=("Arial", 11, "bold"), fg=tab_color).pack(side='left', padx=10, pady=5)
                        if two_instances:
                            tk.Label(label_frame, text=f"{instance2} ({len(instance2_df)})", font=("Arial", 11, "bold"), fg=tab_color).pack(side='left', padx=10, pady=5)

                        # Frames for each grid
                        grid_frame = ttk.Frame(frame)
                        grid_frame.pack(fill='both', expand=True)

                        # Left grid (tester1)
                        left_frame = ttk.Frame(grid_frame, width=600, height=500)
                        left_frame.pack(side='left', fill='both', expand=True)
                        left_frame.pack_propagate(False)
                        tree1_scroll_y = ttk.Scrollbar(left_frame, orient="vertical")
                        tree1_scroll_y.pack(side='right', fill='y')
                        tree1_scroll_x = ttk.Scrollbar(left_frame, orient="horizontal")
                        tree1_scroll_x.pack(side='bottom', fill='x')
                        tree1 = ttk.Treeview(left_frame, columns=list(instance1_df.columns), show='headings',
                                            yscrollcommand=tree1_scroll_y.set, xscrollcommand=tree1_scroll_x.set, height=20, selectmode="browse")
                        tree1.bind("<Button-3>", on_right_click_query)  # Right-click
                        tree1_list.append(tree1)
                        
                        # After creating your Treeview (tree), add this for each column:
                        for col in instance1_df.columns:
                            tree1.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree1, _col, False))

                        tree1.tag_configure('green_text', background='#d4f4dd')
                        tree1.tag_configure('red_text', background='#f4d4d4')

                        # Build a set of row_hashes from instance2_df
                        instance2_hashes = set()
                        if two_instances and not instance2_df.empty:
                            instance2_hashes = set(instance2_df["Row Hash"]) if "Row Hash" in instance2_df.columns else set()

                        for row_idx, row in instance1_df.iterrows():
                            values = []
                            tags = []
                            for col in instance1_df.columns:
                                val = row[col]
                                display_val = val
                                values.append(display_val)
                            if two_instances:    
                                # Mark row if its Row Hash is in instance2_df
                                if "Row Hash" in instance1_df.columns:
                                    if row["Row Hash"] in instance2_hashes:
                                        tags.append('green_text')
                                    else:
                                        tags.append('red_text')
                            # else:
                            #     # If only one instance, mark all rows as green
                            #     tags.append('green_text')
                            tree1.insert('', 'end', iid=row_idx, values=values, tags=tags)

                        tree1.pack(fill='both', expand=True)
                        tree1_scroll_y.config(command=tree1.yview)
                        tree1_scroll_x.config(command=tree1.xview)
                        for col in instance1_df.columns:
                            tree1.heading(col, text=col)
                            tree1.column(col, width=120, stretch=False)

                        if two_instances:
                            # Right grid (tester2)
                            right_frame = ttk.Frame(grid_frame, width=600, height=500)
                            right_frame.pack(side='right', fill='both', expand=True)
                            right_frame.pack_propagate(False)
                            tree2_scroll_y = ttk.Scrollbar(right_frame, orient="vertical")
                            tree2_scroll_y.pack(side='right', fill='y')
                            tree2_scroll_x = ttk.Scrollbar(right_frame, orient="horizontal")
                            tree2_scroll_x.pack(side='bottom', fill='x')
                            tree2 = ttk.Treeview(right_frame, columns=list(instance2_df.columns), show='headings',
                                                yscrollcommand=tree2_scroll_y.set, xscrollcommand=tree2_scroll_x.set, height=20, selectmode="browse")
                            tree2.bind("<Button-3>", on_right_click_query)  # Right-click
                            tree2_list.append(tree2)

                            # tree1.tree1_list = tree1_list
                            # tree1.tree2_list = tree2_list
                            # tree2.tree1_list = tree1_list
                            # tree2.tree2_list = tree2_list

                            # After creating your Treeview (tree), add this for each column:
                            for col in instance2_df.columns:
                                tree2.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree2, _col, False))

                            tree2.tag_configure('green_text', background='#d4f4dd')
                            tree2.tag_configure('red_text', background='#f4d4d4')

                            # Build a set of row_hashes from instance2_df
                            instance1_hashes = set(instance1_df["Row Hash"]) if "Row Hash" in instance1_df.columns else set()

                            for row_idx, row in instance2_df.iterrows():
                                values = []
                                tags = []
                                for col in instance2_df.columns:
                                    val = row[col]
                                    display_val = val
                                    values.append(display_val)
                                # Mark row if its Row Hash is in instance1_df
                                if "Row Hash" in instance2_df.columns:
                                    if row["Row Hash"] in instance1_hashes:
                                        tags.append('green_text')
                                    else:
                                        tags.append('red_text')
                                tree2.insert('', 'end', iid=row_idx, values=values, tags=tags)

                            tree2.pack(fill='both', expand=True)
                            tree2_scroll_y.config(command=tree2.yview)
                            tree2_scroll_x.config(command=tree2.xview)
                            for col in instance2_df.columns:
                                tree2.heading(col, text=col)
                                tree2.column(col, width=120, stretch=False)

                        notebook.add(frame, text=tab_name)
                        
            # --- Status Bar ---
            # status_var = tk.StringVar(value="Ready")
            # status_bar = ttk.Label(result_window, textvariable=status_var, relief="sunken", anchor="w")
            # status_bar.pack(fill="x", side="bottom")
                                                        
    
    def view_baseline(self):
        self.view("Baseline")
        # project = self.get_project(self.project_folder_var.get())
        # if not project:
        #     messagebox.showerror("Error", "No project selected.")
        #     return
        # instance = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == "baseline"), None)
        # if not instance:
        #     messagebox.showerror("Error", "No baseline exists for this project.")
        #     return

        # # if instance.get("interactive"):
        # #     conn_str = None  # Or handle interactive logic in your tester
        # # else:
        # #     conn_str = (
        # #         f"Provider=MSOLAP;Data Source={instance['server_name']};"
        # #         f"Initial Catalog={instance['database_name']};"
        # #         f"User ID={instance['user_id']};Password={self.decrypt_for_user(instance['password'])}"
        # #     )

        # # tester = PowerBIRegressionTester(
        # #     self.project_folder_var.get(),
        # #     conn_str,
        # #     self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        # # )

        # tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance)

        # if not tester.baseline_exists():
        #     messagebox.showerror("Error", "No baseline exists for this project.")
        #     return
        # df = tester.load_baseline_df()
        # self.show_result(df, instance["instance_name"])

    def view_instance(self):
        instance_name = self.instance_dropdown.get().strip()
        self.view(instance_name)
        # if not instance_name:
        #     messagebox.showerror("Error", "No instance selected.")
        #     return

        # project = self.get_project(self.project_folder_var.get())
        # if not project:
        #     messagebox.showerror("Error", "No project selected.")
        #     return

        # instance = next((inst for inst in project.get("instances", []) if inst["instance_name"] == instance_name), None)
        # if not instance:
        #     messagebox.showerror("Error", f"Instance '{instance_name}' does not exist.")
        #     return

        # # if instance.get("interactive"):
        # #     conn_str = None  # Or handle interactive logic in your tester
        # # else:
        # #     conn_str = (
        # #         f"Provider=MSOLAP;Data Source={instance['server_name']};"
        # #         f"Initial Catalog={instance['database_name']};"
        # #         f"User ID={instance['user_id']};Password={self.decrypt_for_user(instance['password'])}"
        # #     )

        # # tester = PowerBIRegressionTester(
        # #     self.project_folder_var.get(),
        # #     conn_str,
        # #     self.pbi_report_folder_var.get() if self.pbi_report_folder_var.get() else ""
        # # )

        # tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance)

        # if not tester.instance_exists(instance_name):
        #     messagebox.showerror("Error", f"Instance '{instance_name}' does not exist in the file system.")
        #     return
        # df = tester.load_instance_df(instance_name)
        # self.show_result(df, instance_name)
        

    def view(self, instance_name):
        # instance_name = self.instance_dropdown.get().strip()
        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return

        if not instance_name:
            messagebox.showerror("Error", "No baseline or instance selected.")
            return

        instance = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == instance_name.lower()), None)
        if not instance:
            messagebox.showerror("Error", f"'{instance_name}' does not exist.")
            return

        tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance)

        if instance_name.lower() == "baseline":
            if not tester.baseline_exists():
                messagebox.showerror("Error", "No baseline exists for this project.")
                return
            df = tester.load_baseline_df()
        else:
            if not tester.instance_exists(instance_name):
                messagebox.showerror("Error", f"Instance '{instance_name}' does not exist in the file system.")
                return
            df = tester.load_instance_df(instance_name)

        self.show_result(df, instance_name)

    def run_baseline(self):
        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return

        # Find existing baseline instance
        instance = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == "baseline"), None)

        if not instance:
            # No baseline exists, prompt for details
            instance = self.prompt_instance_details()
            if not instance:
                return
            self.save_instance_to_project(self.project_folder_var.get(), instance)
        else:
            # Baseline exists, ask to overwrite
            if not messagebox.askyesno("Overwrite?", "Baseline exists. Overwrite?"):
                return

        tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance)

        df = tester.run_baseline()
        self.show_result(df, 'Baseline')

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

        tester = self.create_regession_tester(self.project_folder_var, self.pbi_report_folder_var, instance)

        if tester.instance_exists(instance_name):
            if not messagebox.askyesno("Overwrite?", f"Instance '{instance_name}' exists. Overwrite?"):
                return

        df = tester.run_instance(instance_name)
        if df is not None and not df.empty:
            self.show_result(df, 'Baseline', instance_name)
        else:
            messagebox.showinfo("Compare", "No differences found.")

    def create_instance(self):
        project = self.get_project(self.project_folder_var.get())
        if not project:
            messagebox.showerror("Error", "No project selected.")
            return

        # Always pre-fill with baseline if available, otherwise empty
        baseline = next((inst for inst in project.get("instances", []) if inst["instance_name"].lower() == "baseline"), None)
        prefill = {k: v for k, v in baseline.items() if k != "instance_name"} if baseline else {}
        prefill["instance_name"] = ""  # Always leave instance name empty

        # Prompt for instance details, prefilled from baseline if available
        instance_data = self.prompt_instance_details(prefill)
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
            self.show_result(df, 'Baseline', instance_data["instance_name"])
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
                self.show_result(df, 'Baseline', instance_name)
            else:
                messagebox.showinfo("Compare", "No differences found.")
        else:
            messagebox.showerror("Error", f"The instance '{instance_name}' does not exist.")

    def create_regession_tester(self, project_folder_var, project_report_folder_var, instance):
        data_source = instance.get('server_name')
        database = instance.get('database_name')
        user_id = instance.get('user_id')
        password = self.decrypt_for_user(instance.get('password', ''))
        tenant_id = instance.get('tenant_id', '')
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
            tenant_id=tenant_id,
            interactive=interactive,
            xmla_endpoint=xmla_endpoint,
            local_instance=local_instance
        )

        return tester

class ToolTip(object):
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        # Position below and to the right of the widget
        x, y, _, cy = self.widget.bbox("insert") if self.widget.winfo_class() == 'Entry' else (0, 0, 0, 0)
        x = x + self.widget.winfo_rootx() + 30
        y = y + cy + self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        # Modern style: larger font, more padding, rounded border, shadow
        label = tk.Label(
            tw,
            text=self.text,
            justify='left',
            background="#222",
            foreground="#fff",
            relief='solid',
            borderwidth=2,
            font=("Segoe UI", 12, "normal"),
            padx=16,
            pady=10,
            wraplength=350
        )
        label.pack(ipadx=1, ipady=1)
        # Optional: add a drop shadow effect (simulate with another label)
        # shadow = tk.Label(tw, background="#444", borderwidth=0)
        # shadow.place(x=2, y=2)

    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PowerBIRegressionTesterApp(root)
    root.mainloop()


