import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox
import os
import csv
import sys
import zipfile
import shutil
import subprocess
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from threading import Thread
from Registry import Registry
import datetime
import struct
import json


BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
TOOLS_DIR = os.path.join(BASE_DIR, "tools")
JLECMD_PATH = os.path.join(TOOLS_DIR, "JLECmd", "JLECmd.exe")
SBECMD_PATH = os.path.join(TOOLS_DIR, "SBECmd", "SBECmd.exe")
PECMD_PATH = os.path.join(TOOLS_DIR, "PECmd", "PECmd.exe")

class ForensicParserApp:
    def __init__(self, root):
        self.root = root
        root.title("RegParser v2.0")
        root.configure(bg="#f0f0f0")
        root.state('zoomed')

        self.reg_folder_var = tk.StringVar()
        self.jump_folder_var = tk.StringVar()
        self.prefetch_folder_var = tk.StringVar()
        self.output_folder_var = tk.StringVar()
        self.cancel_flag = False
        self.logo_path_var = tk.StringVar()
        self.temp_zip_dir = None
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)


        
        # Configuration for case information
        self.case_info = {
            'case_name': tk.StringVar(),
            'examiner': tk.StringVar(),
            'date': tk.StringVar(value=datetime.datetime.now().strftime("%Y-%m-%d")),
            'organization': tk.StringVar(),
            'logo_path': tk.StringVar()
        }

        self.create_menu()
        self.create_case_info_frame()
        self.create_frames()
        self.create_console()
        self.load_config()
    
    def browse_logo(self):
        """Browse for organization logo file"""
        logo_path = filedialog.askopenfilename(
            title="Select Organization Logo",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg *.jpeg"),
                ("All files", "*.*")
            ]
        )
        if logo_path:
            self.case_info['logo_path'].set(logo_path)


    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Save Configuration", command=self.save_config)
        file_menu.add_command(label="Load Configuration", command=self.load_config)
        file_menu.add_separator()
        file_menu.add_command(label="Export Report", command=self.export_report)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Verify Tool Paths", command=self.verify_tools)
        tools_menu.add_command(label="View Output Folder", command=self.open_output_folder)

    def create_case_info_frame(self):
        case_frame = tk.LabelFrame(self.root, text="Case Information", bg="#f0f0f0", fg="red", font=("Arial", 12, "bold"))
        case_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
    
        # First row of case info fields
        info_container1 = tk.Frame(case_frame, bg="#f0f0f0")
        info_container1.pack(fill='x', padx=5, pady=2)
    
        tk.Label(info_container1, text="Case Name:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, sticky='w')
        tk.Entry(info_container1, textvariable=self.case_info['case_name'], width=25).grid(row=0, column=1, padx=5)
    
        tk.Label(info_container1, text="Examiner:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, sticky='w')
        tk.Entry(info_container1, textvariable=self.case_info['examiner'], width=20).grid(row=0, column=3, padx=5)
    
        tk.Label(info_container1, text="Date:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=5, sticky='w')
        tk.Entry(info_container1, textvariable=self.case_info['date'], width=15).grid(row=0, column=5, padx=5)
    
        # Second row for organization and logo
        info_container2 = tk.Frame(case_frame, bg="#f0f0f0")
        info_container2.pack(fill='x', padx=5, pady=2)
    
        tk.Label(info_container2, text="Organization:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, sticky='w')
        tk.Entry(info_container2, textvariable=self.case_info['organization'], width=30).grid(row=0, column=1, padx=5)
    
        tk.Label(info_container2, text="Logo:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, sticky='w')
        logo_entry = tk.Entry(info_container2, textvariable=self.case_info['logo_path'], width=25)
        logo_entry.grid(row=0, column=3, padx=5)
        tk.Button(info_container2, text="Browse", command=self.browse_logo, bg="#FFC107", fg="black").grid(row=0, column=4, padx=5)

    def create_frames(self):
        padding = {'padx': 10, 'pady': 10}

        # Registry frame with enhanced features
        reg_frame = tk.LabelFrame(self.root, text="Registry Hive Analysis", bg="#f0f0f0", fg="blue", font=("Arial", 14))
        reg_frame.grid(row=1, column=0, sticky="nsew", **padding)
        self.add_label_entry_button(reg_frame, "Registry Hive Folder:", self.browse_reg_folder, self.reg_folder_var)
        
        # Button frame for registry operations
        reg_buttons = tk.Frame(reg_frame, bg="#f0f0f0")
        reg_buttons.pack(fill='x', pady=5)
        
        tk.Button(reg_buttons, text="Scan for Hives", command=self.scan_hives, bg="#4CAF50", fg="white").pack(side='left', padx=2)
        tk.Button(reg_buttons, text="Select All", command=self.select_all_hives, bg="#607D8B", fg="white").pack(side='left', padx=2)
        tk.Button(reg_buttons, text="Clear Selection", command=self.clear_hive_selection, bg="#607D8B", fg="white").pack(side='left', padx=2)
        tk.Button(reg_buttons, text="Load from ZIP", command=self.load_zip_and_scan, bg="#03A9F4", fg="white").pack(side='left', padx=2)
        
        # Hives listbox with scrollbar
        hives_frame = tk.Frame(reg_frame, bg="#f0f0f0")
        hives_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.hives_listbox = tk.Listbox(hives_frame, selectmode=tk.MULTIPLE, width=100, height=8)
        hives_scrollbar = ttk.Scrollbar(hives_frame, orient="vertical", style="Vertical.TScrollbar")
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Vertical.TScrollbar", width=20)  # Adjust width here

        self.hives_listbox.config(yscrollcommand=hives_scrollbar.set)
        hives_scrollbar.config(command=self.hives_listbox.yview)
        
        self.hives_listbox.pack(side="left", fill="both", expand=True)
        hives_scrollbar.pack(side="right", fill="y")
        
        # Registry analysis buttons
        reg_analysis_frame = tk.Frame(reg_frame, bg="#f0f0f0")
        reg_analysis_frame.pack(fill='x', pady=5)
        
        tk.Button(reg_analysis_frame, text="Parse Selected Hives", command=self.start_parse_hives, bg="#2196F3", fg="white").pack(side='left', padx=2)
        tk.Button(reg_analysis_frame, text="Parse Shellbags", command=self.start_parse_shellbags, bg="#9C27B0", fg="white").pack(side='left', padx=2)
        tk.Button(reg_analysis_frame, text="Parse USB Devices", command=self.start_parse_usb_devices, bg="#FF9800", fg="white").pack(side='left', padx=2)
        
        reg_analysis_frame2 = tk.Frame(reg_frame, bg="#f0f0f0")
        reg_analysis_frame2.pack(fill='x', pady=2)
        
        tk.Button(reg_analysis_frame2, text="Parse Bluetooth", command=self.start_parse_bluetooth, bg="#00796B", fg="white").pack(side='left', padx=2)
        tk.Button(reg_analysis_frame2, text="Parse Network", command=self.start_parse_network, bg="#33691E", fg="white").pack(side='left', padx=2)


        # Jump Lists frame
        jump_frame = tk.LabelFrame(self.root, text="Jump Lists Analysis", bg="#f0f0f0", fg="purple", font=("Arial", 14))
        jump_frame.grid(row=1, column=1, sticky="nsew", **padding)
        self.add_label_entry_button(jump_frame, "Jump Lists Folder:", self.browse_jump_folder, self.jump_folder_var)
        tk.Button(jump_frame, text="Parse Jump Lists", command=self.start_parse_jump_lists, bg="#FF5722", fg="white").pack(pady=5)

        # Prefetch frame
        prefetch_frame = tk.LabelFrame(self.root, text="Prefetch Analysis", bg="#f0f0f0", fg="blue", font=("Arial", 14))
        prefetch_frame.grid(row=2, column=0, sticky="nsew", **padding)
        self.add_label_entry_button(prefetch_frame, "Prefetch Folder:", self.browse_prefetch_folder, self.prefetch_folder_var)
        tk.Button(prefetch_frame, text="Parse Prefetch", command=self.start_parse_prefetch, bg="#795548", fg="white").pack(pady=5)

        # Control frame with enhanced options
        control_frame = tk.LabelFrame(self.root, text="Output & Control", bg="#f0f0f0", fg="green", font=("Arial", 14))
        control_frame.grid(row=2, column=1, sticky="nsew", **padding)
        self.add_label_entry_button(control_frame, "Output Folder:", self.browse_output_folder, self.output_folder_var)
        
        control_buttons = tk.Frame(control_frame, bg="#f0f0f0")
        control_buttons.pack(fill='x', pady=5)
        
        tk.Button(control_buttons, text="Cancel", bg="#f44336", fg="white", command=self.cancel_parsing).pack(side='left', padx=5)
        tk.Button(control_buttons, text="Clear Log", bg="#9E9E9E", fg="white", command=self.clear_log).pack(side='left', padx=5)
        tk.Button(control_buttons, text="Exit", bg="black", fg="white", command=self.on_close).pack(side='left', padx=5)


        self.root.grid_columnconfigure((0,1), weight=1)
        self.root.grid_rowconfigure((1,2), weight=1)

    def add_label_entry_button(self, frame, text, browse_command, text_var):
        sub_frame = tk.Frame(frame, bg="#f0f0f0")
        sub_frame.pack(fill='x', padx=5, pady=5)
        tk.Label(sub_frame, text=text, bg="#f0f0f0", width=20, anchor="w").pack(side="left")
        entry = tk.Entry(sub_frame, textvariable=text_var, width=40)
        entry.pack(side="left", padx=5, fill='x', expand=True)
        tk.Button(sub_frame, text="Browse", command=browse_command, bg="#FFC107", fg="black").pack(side="left", padx=5)

    def create_console(self):
        console_frame = tk.LabelFrame(self.root, text="Output Log", bg="#f0f0f0", font=("Arial", 14))
        console_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
        
        # Console with improved styling
        self.output_console = scrolledtext.ScrolledText(console_frame, width=150, height=12, 
                                                       bg="#f0f0f0", font=("Courier", 9))
        self.output_console.pack(padx=5, pady=5, fill='both', expand=True)
        
        # Progress bar with better styling
        progress_frame = tk.Frame(console_frame, bg="#f0f0f0")
        progress_frame.pack(fill='x', padx=5, pady=5)
        
        self.progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=1000, mode='determinate')
        self.progress.pack(side='left', fill='x', expand=True)
        
        self.progress_label = tk.Label(progress_frame, text="0%", bg="#f0f0f0", width=5)
        self.progress_label.pack(side='right', padx=5)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = tk.Label(console_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w', bg="#e0e0e0")
        status_bar.pack(fill='x')

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        
        self.output_console.config(state='normal')
        self.output_console.insert(tk.END, formatted_message + "\n")
        self.output_console.see(tk.END)
        self.output_console.config(state='disabled')
        self.status_var.set(message)
        self.root.update_idletasks()

    def clear_log(self):
        self.output_console.config(state='normal')
        self.output_console.delete(1.0, tk.END)
        self.output_console.config(state='disabled')

    def select_all_hives(self):
        self.hives_listbox.select_set(0, tk.END)

    def clear_hive_selection(self):
        self.hives_listbox.selection_clear(0, tk.END)

    def update_progress(self, current, total):
        if total > 0:
            percentage = int((current / total) * 100)
            self.progress["value"] = percentage
            self.progress_label.config(text=f"{percentage}%")
            self.root.update_idletasks()

    def save_config(self):
        config = {
            'reg_folder': self.reg_folder_var.get(),
            'jump_folder': self.jump_folder_var.get(),
            'prefetch_folder': self.prefetch_folder_var.get(),
            'output_folder': self.output_folder_var.get(),
            'case_name': self.case_info['case_name'].get(),
            'examiner': self.case_info['examiner'].get(),
            'date': self.case_info['date'].get(),
            'organization': self.case_info['organization'].get(),
            'logo_path': self.case_info['logo_path'].get()
        }
    
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
    
        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(config, f, indent=2)
                self.log(f"✅ Configuration saved to {filename}")
            except Exception as e:
                self.log(f"❌ Failed to save configuration: {e}")

    def load_config(self):
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
    
        if filename and os.path.exists(filename):
            try:
                with open(filename, 'r') as f:
                    config = json.load(f)
            
                self.reg_folder_var.set(config.get('reg_folder', ''))
                self.jump_folder_var.set(config.get('jump_folder', ''))
                self.prefetch_folder_var.set(config.get('prefetch_folder', ''))
                self.output_folder_var.set(config.get('output_folder', ''))
                self.case_info['case_name'].set(config.get('case_name', ''))
                self.case_info['examiner'].set(config.get('examiner', ''))
                self.case_info['date'].set(config.get('date', ''))
                self.case_info['organization'].set(config.get('organization', ''))
                self.case_info['logo_path'].set(config.get('logo_path', ''))
            
                self.log(f"✅ Configuration loaded from {filename}")
            except Exception as e:
                self.log(f"❌ Failed to load configuration: {e}")
    def copy_app_logo_to_output(self, output_dir):
        src_path = os.path.join(BASE_DIR, "app_logo.png")
        if os.path.exists(src_path):
            try:
                import shutil
                dest_path = os.path.join(output_dir, "app_logo.png")
                shutil.copy2(src_path, dest_path)
                return "app_logo.png"  # Relative path
            except Exception as e:
                print(f"⚠️ Failed to copy app logo: {e}")
        return None

    def copy_logo_to_output(self, logo_path, output_dir):
        """Copy logo file to output directory and return relative path"""
        if not logo_path or not os.path.exists(logo_path):
            return None
    
        try:
            logo_filename = os.path.basename(logo_path)
            logo_dest = os.path.join(output_dir, logo_filename)
        
            # Copy logo file to output directory
            import shutil
            shutil.copy2(logo_path, logo_dest)
        
            return logo_filename  # Return relative path
        except Exception as e:
            self.log(f"⚠️ Failed to copy logo: {e}")
            return None
    def get_analysis_summary(self):
        """Generate analysis summary for the report"""
        summary = {
            'registry_files': self.hives_listbox.size(),
            'output_folders': []
        }
    
        output_base = self.output_folder_var.get()
        if output_base:
            # Check which output folders exist
            potential_folders = ['Registry', 'JumpLists', 'Prefetch', 'Shellbags', 'USB_Devices', 'Bluetooth_Devices', 'Network_Connections']
            for folder in potential_folders:
                folder_path = os.path.join(output_base, folder)
                if os.path.exists(folder_path):
                    file_count = len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
                    summary['output_folders'].append({'name': folder, 'file_count': file_count})
    
        return summary
    
    def verify_tools(self):
        tools = {
            "JLECmd": JLECMD_PATH,
            "SBECmd": SBECMD_PATH,
            "PECmd": PECMD_PATH
        }
        
        missing_tools = []
        for tool_name, tool_path in tools.items():
            if not os.path.exists(tool_path):
                missing_tools.append(f"{tool_name}: {tool_path}")
        
        if missing_tools:
            message = "Missing tools:\n" + "\n".join(missing_tools)
            messagebox.showwarning("Missing Tools", message)
            self.log("⚠️ Some tools are missing. Check paths in TOOLS_DIR.")
        else:
            messagebox.showinfo("Tool Verification", "All tools found successfully!")
            self.log("✅ All forensic tools verified.")

    def open_output_folder(self):
        output_folder = self.output_folder_var.get()
        if output_folder and os.path.exists(output_folder):
            os.startfile(output_folder)
        else:
            messagebox.showwarning("Output Folder", "Output folder not set or doesn't exist.")

    def export_report(self):
        if not self.output_folder_var.get():
            messagebox.showwarning("Export Report", "Please set an output folder first.")
            return

        # Create a pop-up window
        export_window = tk.Toplevel(self.root)
        export_window.title("Select Export Format")
        export_window.geometry("300x160")
        export_window.resizable(False, False)

        html_var = tk.BooleanVar(value=True)
        pdf_var = tk.BooleanVar(value=True)

        tk.Label(export_window, text="Choose report format:", font=("Helvetica", 12)).pack(pady=10)
        tk.Checkbutton(export_window, text="Export as HTML", variable=html_var).pack(anchor='w', padx=20)
        tk.Checkbutton(export_window, text="Export as PDF", variable=pdf_var).pack(anchor='w', padx=20)

        def perform_export():
            base_path = os.path.join(self.output_folder_var.get(), "forensic_analysis_report")
            try:
                if html_var.get():
                    self.generate_html_report(base_path + ".html")
                if pdf_var.get():
                    self.generate_pdf_report(base_path + ".pdf")

                formats = []
                if html_var.get(): formats.append("HTML")
                if pdf_var.get(): formats.append("PDF")
                if formats:
                    self.log(f"✅ Report exported as: {', '.join(formats)}")
                    messagebox.showinfo("Export Success", f"Exported report as: {', '.join(formats)}")
                else:
                    self.log("⚠️ No format selected for export.")
                    messagebox.showwarning("No Format", "No format selected.")

            except Exception as e:
                self.log(f"❌ Failed to export report: {e}")
                messagebox.showerror("Export Error", f"Failed to export report:\n{e}")
            export_window.destroy()

        tk.Button(export_window, text="Export", command=perform_export, bg="#4CAF50", fg="white").pack(pady=10)

    
    def generate_html_report(self, output_path):
        output_dir = os.path.dirname(output_path)
        logo_filename = self.copy_logo_to_output(self.case_info['logo_path'].get(), output_dir)
        app_logo_filename = self.copy_app_logo_to_output(output_dir)
        analysis_summary = self.get_analysis_summary()
    
        # Generate logo HTML
        logo_html = ""
        if logo_filename:
            logo_html = f'<img src="{logo_filename}" alt="Organization Logo" style="max-height: 120px; float: right;">'
        
        app_logo_html = f'<img src="{app_logo_filename}" alt="App Logo" style="max-height: 100px;">' if app_logo_filename else ""

    
        # Generate analysis summary HTML
        summary_html = ""
        if analysis_summary['output_folders']:
            summary_html = "<ul>"
            for folder in analysis_summary['output_folders']:
                summary_html += f"<li>{folder['name']}: {folder['file_count']} files generated</li>"
            summary_html += "</ul>"
        else:
            summary_html = "<p>No output files generated yet.</p>"

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Forensic Analysis Report</title>
            <meta charset="UTF-8">
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }}
                .header {{ background-color: #f8f9fa; padding: 20px; border-radius: 5px; border-left: 5px solid #007bff; margin-bottom: 20px; }}
                .organization-info {{ overflow: hidden; margin-bottom: 15px; }}
                .case-details {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 15px; }}
                .section {{ margin: 20px 0; }}
                .section h2 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }}
                .artifact {{ background-color: #f9f9f9; padding: 15px; margin: 10px 0; border-left: 4px solid #2196F3; border-radius: 4px; }}
                .artifact h3 {{ margin-top: 0; color: #34495e; }}
                table {{ border-collapse: collapse; width: 100%; margin: 15px 0; }}
                th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
                th {{ background-color: #f2f2f2; font-weight: bold; }}
                tr:nth-child(even) {{ background-color: #f9f9f9; }}
                .path-cell {{ font-family: 'Courier New', monospace; font-size: 0.9em; word-break: break-all; }}
                .status-complete {{ color: #27ae60; font-weight: bold; }}
                .status-pending {{ color: #f39c12; font-weight: bold; }}
                .status-missing {{ color: #e74c3c; font-weight: bold; }}
                .footer {{ margin-top: 40px; padding: 20px; background-color: #ecf0f1; border-radius: 5px; text-align: center; font-size: 0.9em; color: #7f8c8d; }}
            </style>
        </head>
        <body>
            <div class="header">
                <div class="organization-info">
                    {logo_html}
                    <h1>Forensic Analysis Report</h1>
                    {f'<h2 style="color: #2980b9; margin: 5px 0;">{self.case_info["organization"].get()}</h2>' if self.case_info["organization"].get() else ''}
                </div>
                <div class="case-details">
                    <div>
                        <p><strong>Case Name:</strong> {self.case_info['case_name'].get() or 'Not specified'}</p>
                        <p><strong>Examiner:</strong> {self.case_info['examiner'].get() or 'Not specified'}</p>
                    </div>
                    <div>
                        <p><strong>Analysis Date:</strong> {self.case_info['date'].get()}</p>
                        <p><strong>Report Generated:</strong> {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
                    </div>
                </div>
            </div>
        
            <div class="section">
                <h2>Executive Summary</h2>
                <div class="artifact">
                    <h3>Analysis Overview</h3>
                    <p>This report contains the results of digital forensic analysis performed on various system artifacts including registry hives, jump lists, prefetch files, and other system artifacts.</p>
                    <p>All steps have been taken to maintain integrity of evidence.</p>
                    <h4>Processed Artifacts Summary:</h4>
                    {summary_html}
                </div>
            </div>
        
            <div class="section">
            <p> </p> 
                <h2>Evidence Sources</h2>
                <table>
                    <tr>
                        <th>Artifact Type</th>
                        <th>Source Path</th>
                        <th>Status</th>
                        <th>Output Location</th>
                    </tr>
                    <tr>
                        <td>Registry Hives</td>
                        <td class="path-cell">{self.reg_folder_var.get() or 'Not specified'}</td>
                        <td class="{'status-complete' if self.reg_folder_var.get() else 'status-missing'}">{f'{analysis_summary["registry_files"]} files detected' if self.reg_folder_var.get() else 'Not configured'}</td>
                        <td class="path-cell">{os.path.join(self.output_folder_var.get(), 'Registry') if self.output_folder_var.get() else 'Not set'}</td>
                    </tr>
                    <tr>
                        <td>Jump Lists</td>
                        <td class="path-cell">{self.jump_folder_var.get() or 'Not specified'}</td>
                        <td class="{'status-complete' if self.jump_folder_var.get() else 'status-missing'}">{'Parsed successfully' if self.jump_folder_var.get() else 'Not configured'}</td>
                        <td class="path-cell">{os.path.join(self.output_folder_var.get(), 'JumpLists') if self.output_folder_var.get() else 'Not set'}</td>
                    </tr>
                    <tr>
                        <td>Prefetch Files</td>
                        <td class="path-cell">{self.prefetch_folder_var.get() or 'Not specified'}</td>
                        <td class="{'status-complete' if self.prefetch_folder_var.get() else 'status-missing'}">{'Parsed successfully' if self.prefetch_folder_var.get() else 'Not configured'}</td>
                        <td class="path-cell">{os.path.join(self.output_folder_var.get(), 'Prefetch') if self.output_folder_var.get() else 'Not set'}</td>
                    </tr>
                    <tr>
                        <td>Shellbags</td>
                        <td class="path-cell">{self.reg_folder_var.get() or 'Not specified'}</td>
                        <td class="{'status-complete' if self.reg_folder_var.get() else 'status-missing'}">{'Available from Registry' if self.reg_folder_var.get() else 'Requires Registry'}</td>
                        <td class="path-cell">{os.path.join(self.output_folder_var.get(), 'Shellbags') if self.output_folder_var.get() else 'Not set'}</td>
                    </tr>
                    <tr>
                        <td>USB Devices</td>
                        <td class="path-cell">{self.reg_folder_var.get() or 'Not specified'}</td>
                        <td class="{'status-complete' if self.reg_folder_var.get() else 'status-missing'}">{'Available from SYSTEM hive' if self.reg_folder_var.get() else 'Requires SYSTEM hive'}</td>
                        <td class="path-cell">{os.path.join(self.output_folder_var.get(), 'USB_Devices') if self.output_folder_var.get() else 'Not set'}</td>
                    </tr>
                    <tr>
                        <td>Bluetooth Devices</td>
                        <td class="path-cell">{self.reg_folder_var.get() or 'Not specified'}</td>
                        <td class="{'status-complete' if self.reg_folder_var.get() else 'status-missing'}">{'Available from SYSTEM hive' if self.reg_folder_var.get() else 'Requires SYSTEM hive'}</td>
                        <td class="path-cell">{os.path.join(self.output_folder_var.get(), 'Bluetooth_Devices') if self.output_folder_var.get() else 'Not set'}</td>
                    </tr>
                    <tr>
                        <td>Network Profiles</td>
                        <td class="path-cell">{self.reg_folder_var.get() or 'Not specified'}</td>
                        <td class="{'status-complete' if self.reg_folder_var.get() else 'status-missing'}">{'Available from SOFTWARE hive' if self.reg_folder_var.get() else 'Requires SOFTWARE hive'}</td>
                        <td class="path-cell">{os.path.join(self.output_folder_var.get(), 'Network_Connections') if self.output_folder_var.get() else 'Not set'}</td>
                    </tr>
                </table>
            </div>
        
            <div class="section">
                <h2>Tool Information</h2>
                <div class="artifact">
                    <h3>Forensic Tools Used</h3>
                    <ul>
                        <li><strong>Registry Analysis:</strong> Custom Python parser using python-registry library</li>
                        <li><strong>Jump Lists:</strong> JLECmd.exe 1.5.1 (Eric Zimmerman Tools)</li>
                        <li><strong>Shellbags:</strong> SBECmd.exe 2.1.0 (Eric Zimmerman Tools)</li>
                        <li><strong>Prefetch:</strong> PECmd.exe 1.5.1 (Eric Zimmerman Tools)</li>
                        <li><strong>Report Generation:</strong> RegParser v2.0</li>
                    </ul>
                </div>
            </div>
        
            <div class="footer">
                {app_logo_html}
                <p>This report was generated by RegParser v2.0 developed by Soukarya Sur. Follow me on linkedin.com/in/soukarya-sur-096589256</p>
                <p>For questions about this analysis, please contact: {self.case_info['examiner'].get() or 'the assigned examiner'}</p>
                <p>RegParser v2.0 will not be responsible for any data loss.</p>
            </div>
        </body>
        </html>
        """
    
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

    def generate_pdf_report(self, output_path):
        output_dir = os.path.dirname(output_path)
        summary = self.get_analysis_summary()
        logo_filename = self.copy_logo_to_output(self.case_info['logo_path'].get(), output_dir)
        app_logo_filename = self.copy_app_logo_to_output(output_dir)

        pdf = canvas.Canvas(output_path, pagesize=A4)
        width, height = A4
        y = height - 40

        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawString(40, y, "Forensic Analysis Report")
        y -= 25

        pdf.setFont("Helvetica", 10)
        pdf.drawString(40, y, f"Organization: {self.case_info['organization'].get()}")
        y -= 15
        pdf.drawString(40, y, f"Case Name: {self.case_info['case_name'].get()}")
        y -= 15
        pdf.drawString(40, y, f"Examiner: {self.case_info['examiner'].get()}")
        y -= 15
        pdf.drawString(40, y, f"Date: {self.case_info['date'].get()}")
        y -= 30

        if logo_filename:
            try:
                logo_path = os.path.join(output_dir, logo_filename)
                # Draw org logo at top-right corner
                pdf.drawImage(logo_path, width - 100, 400, width=70, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                self.log(f"⚠️ Failed to draw organization logo: {e}")

        # Executive Summary
        pdf.setFont("Helvetica-Bold", 12)
        pdf.drawString(40, y, "Executive Summary")
        y -= 18
        pdf.setFont("Helvetica", 10)
        pdf.drawString(40, y, "This report contains the results of forensic analysis on various artifacts.")
        y -= 30

        # Summary
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(40, y, "Processed Artifacts:")
        y -= 18
        pdf.setFont("Helvetica", 10)

        for folder in summary['output_folders']:
            if y < 60:
                pdf.showPage()
                y = height - 40
                pdf.setFont("Helvetica", 10)
            pdf.drawString(50, y, f"- {folder['name']}: {folder['file_count']} files")
            y -= 15

        y -= 10
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(40, y, "Evidence Sources:")
        y -= 20
        pdf.setFont("Helvetica", 10)

        def draw_source(label, input_path, output_folder):
            nonlocal y
            if y < 60:
                pdf.showPage()
                y = height - 40
                pdf.setFont("Helvetica", 10)

            # Check if the tool has run (folder created and has files)
            if os.path.exists(output_folder) and os.listdir(output_folder):
                pdf.drawString(50, y, f"{label}:")
                y -= 15
                pdf.drawString(70, y, f"Input Path: {input_path or 'Not set'}")
                y -= 15
                pdf.drawString(70, y, f"Output Folder: {output_folder}")
                y -= 20
            else:
                pdf.drawString(50, y, f"{label}: Not parsed")
                y -= 20

        draw_source("Registry Hives", self.reg_folder_var.get(), os.path.join(self.output_folder_var.get(), 'Registry'))
        draw_source("Jump Lists", self.jump_folder_var.get(), os.path.join(self.output_folder_var.get(), 'JumpLists'))
        draw_source("Prefetch", self.prefetch_folder_var.get(), os.path.join(self.output_folder_var.get(), 'Prefetch'))
        draw_source("USB Devices", self.reg_folder_var.get(), os.path.join(self.output_folder_var.get(), 'USB_Devices'))
        draw_source("Shellbags", self.reg_folder_var.get(), os.path.join(self.output_folder_var.get(), 'Shellbags'))
        draw_source("Bluetooth Devices", self.reg_folder_var.get(), os.path.join(self.output_folder_var.get(), 'Bluetooth_Devices'))
        draw_source("Network Profiles", self.reg_folder_var.get(), os.path.join(self.output_folder_var.get(), 'Network_Connections'))

        y -= 10
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(40, y, "Tool Info:")
        y -= 15
        pdf.setFont("Helvetica", 10)
        tools = [
            "Registry Analysis: Internal Python Parser",
            "Jump Lists: JLECmd.exe",
            "Shellbags: SBECmd.exe",
            "Prefetch: PECmd.exe",
            "Report Generation: RegParser v2.0"
        ]
        for tool in tools:
            pdf.drawString(50, y, f"- {tool}")
            y -= 15

        y -= 10
        pdf.setFont("Helvetica-Oblique", 8)
        pdf.drawString(40, y, "Generated by RegParser v2.0 developed by Soukarya Sur")
        y -= 10
        pdf.drawString(40, y, f"LinkedIn: linkedin.com/in/soukarya-sur-096589256")
        y -= 10
        pdf.drawString(40, y, f"For queries contact: {self.case_info['examiner'].get() or 'N/A'}")

        if app_logo_filename:
            try:
                app_logo_path = os.path.join(output_dir, app_logo_filename)
                # Draw app logo at bottom-right corner (leave margin)
                pdf.drawImage(app_logo_path, width - 100, 20, width=70, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                self.log(f"⚠️ Failed to draw app logo: {e}")

        pdf.save()

    def browse_reg_folder(self): self.reg_folder_var.set(filedialog.askdirectory() or "")
    def browse_jump_folder(self):
        start_dir = ""
        if self.temp_zip_dir:
            potential = os.path.join(self.temp_zip_dir)
            if os.path.exists(potential):
                start_dir = potential

        selected = filedialog.askdirectory(title="Select JumpLists folder", initialdir=start_dir or None)
        if selected:
            self.jump_folder_var.set(selected)


    def browse_prefetch_folder(self):
        start_dir = ""
        if self.temp_zip_dir:
            potential = os.path.join(self.temp_zip_dir)
            if os.path.exists(potential):
                start_dir = potential

        selected = filedialog.askdirectory(title="Select Prefetch folder", initialdir=start_dir or None)
        if selected:
            self.prefetch_folder_var.set(selected)


    def browse_output_folder(self): self.output_folder_var.set(filedialog.askdirectory() or "")

    def cancel_parsing(self):
        self.cancel_flag = True
        self.log("🛑 Cancel requested. Waiting for threads to stop...")

    def start_parse_hives(self): self.start_thread(self.thread_parse_hives)
    def start_parse_jump_lists(self): self.start_thread(self.thread_parse_jump_lists)
    def start_parse_shellbags(self): self.start_thread(self.thread_parse_shellbags)
    def start_parse_prefetch(self): self.start_thread(self.thread_parse_prefetch)
    def start_parse_usb_devices(self): self.start_thread(self.thread_parse_usb_devices)
    def start_parse_bluetooth(self): self.start_thread(self.thread_parse_bluetooth)
    def start_parse_network(self): self.start_thread(self.thread_parse_network)

    def start_thread(self, target_func):
        self.cancel_flag = False
        thread = Thread(target=target_func)
        thread.daemon = True
        thread.start()

    def thread_parse_hives(self):
        output = self.output_folder_var.get()
        indices = self.hives_listbox.curselection()
        if not (output and indices):
            self.log("⚠️ Missing output folder or hive selection.")
            return
        
        out_dir = os.path.join(output, "Registry")
        os.makedirs(out_dir, exist_ok=True)
        
        total_hives = len(indices)
        self.progress["maximum"] = 100
        
        for idx, i in enumerate(indices, 1):
            if self.cancel_flag:
                self.log("🛑 Hive parsing canceled.")
                break
                
            hive_path = self.hives_listbox.get(i)
            hive_name = os.path.basename(hive_path)
            out_file = os.path.join(out_dir, f"{hive_name}.csv")
            
            try:
                self.log(f"🔍 Parsing {hive_name} ({idx}/{total_hives})")
                parse_registry_hive(hive_path, out_file)
                self.log(f"✅ Saved to {out_file}")
            except Exception as e:
                self.log(f"❌ Failed parsing {hive_name}: {e}")
            
            self.update_progress(idx, total_hives)
            
        self.log(f"✅ Registry parsing complete. Processed {len(indices)} hives.")
        self.status_var.set("Registry parsing complete.")

    def thread_parse_jump_lists(self):
        folder, output = self.jump_folder_var.get(), self.output_folder_var.get()
        if not (folder and output):
            self.log("⚠️ Missing jump lists folder or output folder.")
            return
            
        out_dir = os.path.join(output, "JumpLists")
        os.makedirs(out_dir, exist_ok=True)
        
        try:
            self.log("🔍 Parsing Jump Lists...")
            self.progress.config(mode='indeterminate')
            self.progress.start()
            
            result = subprocess.run([JLECMD_PATH, "-d", folder, "--csv", out_dir], 
                                  check=True, capture_output=True, text=True)
            
            self.progress.stop()
            self.progress.config(mode='determinate')
            self.log(f"✅ Jump Lists parsed successfully. Output: {out_dir}")
            
        except subprocess.CalledProcessError as e:
            self.log(f"❌ JLECmd failed: {e}")
        except Exception as e:
            self.log(f"❌ Unexpected error: {e}")
        finally:
            self.progress.stop()
            self.progress.config(mode='determinate')
            
        self.status_var.set("Jump Lists parsing complete.")

    def thread_parse_shellbags(self):
        folder, output = self.reg_folder_var.get(), self.output_folder_var.get()
        if not (folder and output):
            self.log("⚠️ Missing registry folder or output folder.")
            return
            
        out_dir = os.path.join(output, "Shellbags")
        os.makedirs(out_dir, exist_ok=True)
        
        try:
            self.log("🔍 Parsing Shellbags...")
            self.progress.config(mode='indeterminate')
            self.progress.start()
            
            subprocess.run([SBECMD_PATH, "-d", folder, "--csv", out_dir], check=True)
            self.log(f"✅ Shellbags parsed successfully. Output: {out_dir}")
            
        except Exception as e:
            self.log(f"❌ Shellbags parsing failed: {e}")
        finally:
            self.progress.stop()
            self.progress.config(mode='determinate')
            
        self.status_var.set("Shellbags parsing complete.")

    def thread_parse_prefetch(self):
        folder, output = self.prefetch_folder_var.get(), self.output_folder_var.get()
        if not (folder and output):
            self.log("⚠️ Missing prefetch folder or output folder.")
            return
            
        out_dir = os.path.join(output, "Prefetch")
        os.makedirs(out_dir, exist_ok=True)
        
        try:
            self.log("🔍 Parsing Prefetch files...")
            self.progress.config(mode='indeterminate')
            self.progress.start()
            
            subprocess.run([PECMD_PATH, "-d", folder, "--csv", out_dir], check=True)
            self.log(f"✅ Prefetch files parsed successfully. Output: {out_dir}")
            
        except Exception as e:
            self.log(f"❌ Prefetch parsing failed: {e}")
        finally:
            self.progress.stop()
            self.progress.config(mode='determinate')
            
        self.status_var.set("Prefetch parsing complete.")

    def thread_parse_usb_devices(self):
        output = self.output_folder_var.get()
        indices = self.hives_listbox.curselection()
        if not (output and indices):
            self.log("⚠️ Missing output folder or hive selection.")
            return
            
        system_hive_path = None
        for i in indices:
            path = self.hives_listbox.get(i)
            if os.path.basename(path).upper() == "SYSTEM":
                system_hive_path = path
                break
                
        if not system_hive_path:
            self.log("⚠️ Please select the SYSTEM hive to parse USB devices.")
            return
            
        out_dir = os.path.join(output, "USB_Devices")
        os.makedirs(out_dir, exist_ok=True)
        out_file = os.path.join(out_dir, "USB_Devices.csv")
        
        try:
            self.log(f"🔍 Parsing USB devices from {os.path.basename(system_hive_path)}")
            self.progress.config(mode='indeterminate')
            self.progress.start()
            
            parse_usb_devices_from_system_hive(system_hive_path, out_file)
            self.log(f"✅ USB device information saved to {out_file}")
            
        except Exception as e:
            self.log(f"❌ Failed parsing USB devices: {e}")
        finally:
            self.progress.stop()
            self.progress.config(mode='determinate')
            
        self.status_var.set("USB device parsing complete.")
        
    def thread_parse_bluetooth(self):

        output = self.output_folder_var.get()
        indices = self.hives_listbox.curselection()
        if not (output and indices):
            self.log("⚠️ Select output folder and SYSTEM hive.")
            return

        out_bt_dir = os.path.join(output, "Bluetooth_Devices")
        os.makedirs(out_bt_dir, exist_ok=True)
        bt_file = os.path.join(out_bt_dir, "Bluetooth_SYSTEM.csv")

        def filetime_to_dt(ft):
            try:
                if isinstance(ft, bytes) and len(ft) == 8:
                    ts = struct.unpack("<Q", ft)[0]
                elif isinstance(ft, int):
                    ts = ft
                else:
                    return ""
                timestamp = (ts - 116444736000000000) / 10000000
                dt = datetime.datetime.utcfromtimestamp(timestamp)
                return dt.strftime('%Y-%m-%d %H:%M:%S UTC')
            except:
                return ""

        def decode_device_name(raw_bytes):
            if not raw_bytes:
                return ""

            # Try UTF-16 first (some devices use it)
            try:
                name = raw_bytes.decode('utf-16-le').strip('\x00')
                if name and all(32 <= ord(c) < 127 or c.isspace() for c in name):  # basic ASCII printable
                    return name
            except:
                pass

            # Try UTF-8 next
            try:
                return raw_bytes.decode('utf-8', errors='replace').strip('\x00')
            except:
                pass

            

            return raw_bytes.hex()  # Raw hex fallback

        def parse_cod(cod):
            major_device_classes = {
                0x00: 'Miscellaneous',
                0x01: 'Computer',
                0x02: 'Phone',
                0x03: 'LAN/Network Access Point',
                0x04: 'Audio/Video',
                0x05: 'Peripheral',
                0x06: 'Imaging',
                0x07: 'Wearable',
                0x08: 'Toy',
                0x09: 'Health',
            }
            try:
                major = (cod >> 8) & 0x1F
                return major_device_classes.get(major, 'Unknown')
            except:
                return ""

        try:
            self.log("🔍 Parsing Bluetooth devices...")
            self.progress.config(mode='indeterminate')
            self.progress.start()

            with open(bt_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Hive', 'MAC Address', 'Name', 'ClassOfDevice', 'Device Type', 'LastSeen', 'LastConnected'])

                device_count = 0
                for i in indices:
                    hive_path = self.hives_listbox.get(i)
                    if os.path.basename(hive_path).upper() != "SYSTEM":
                        continue
                    try:
                        reg = Registry.Registry(hive_path)
                        try:
                            root = reg.open("ControlSet001\\Services\\BTHPORT\\Parameters\\Devices")
                        except:
                            self.log(f"⚠️ Devices key not found in {os.path.basename(hive_path)}")
                            continue

                        for dev in root.subkeys():
                            mac = dev.name()

                            def get_value(name):
                                try:
                                    return dev.value(name).value()
                                except:
                                    return None

                            name_bin = get_value("Name")
                            name = decode_device_name(name_bin) if name_bin else ""
                            class_of_device = get_value("COD")
                            device_type = parse_cod(class_of_device) if class_of_device else ""
                            last_seen = filetime_to_dt(get_value("LastSeen"))
                            last_conn = filetime_to_dt(get_value("LastConnected"))

                            writer.writerow([
                                os.path.basename(hive_path),
                                mac,
                                name,
                                class_of_device if class_of_device else "",
                                device_type,
                                last_seen,
                                last_conn
                            ])
                            device_count += 1

                        self.log(f"✅ Bluetooth devices parsed from {os.path.basename(hive_path)}")
                    except Exception as e:
                        self.log(f"❌ Bluetooth parse failed for {os.path.basename(hive_path)}: {e}")

            self.log(f"✅ Found {device_count} Bluetooth devices. Output: {bt_file}")
        except Exception as e:
            self.log(f"❌ Bluetooth parsing failed: {e}")
        finally:
            self.progress.stop()
            self.progress.config(mode='determinate')

        self.status_var.set("Bluetooth parsing complete.")

    def thread_parse_network(self):
        output = self.output_folder_var.get()
        indices = self.hives_listbox.curselection()
        if not (output and indices):
            self.log("⚠️ Select output folder and SOFTWARE hive.")
            return

        out_net_dir = os.path.join(output, "Network_Connections")
        os.makedirs(out_net_dir, exist_ok=True)
        net_file = os.path.join(out_net_dir, "NetworkProfiles_SOFTWARE.csv")

        def systemtime_to_dt(data):
            try:
                if isinstance(data, bytes) and len(data) >= 16:
                    year = int.from_bytes(data[0:2], 'little')
                    month = int.from_bytes(data[2:4], 'little')
                    day = int.from_bytes(data[6:8], 'little')
                    hour = int.from_bytes(data[8:10], 'little')
                    minute = int.from_bytes(data[10:12], 'little')
                    second = int.from_bytes(data[12:14], 'little')
                    millisecond = int.from_bytes(data[14:16], 'little')

                    dt = datetime.datetime(year, month, day, hour, minute, second, millisecond * 1000)
                    return dt.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3] + " UTC"
                else:
                    return "Invalid SYSTEMTIME"
            except Exception as e:
                return f"Error: {e}"

        def filetime_to_dt(ft):
            try:
                if isinstance(ft, bytes) and len(ft) == 8:
                    ts = struct.unpack("<Q", ft)[0]
                elif isinstance(ft, int):
                    ts = ft
                else:
                    return "Invalid time format"

                timestamp = (ts - 116444736000000000) / 10000000
                dt = datetime.datetime.utcfromtimestamp(timestamp)
                return dt.strftime('%Y-%m-%d %H:%M:%S UTC')

            except Exception as e:
                return f"Error: {e}"

        def parse_timestamp(value):
            if isinstance(value, bytes):
                if len(value) == 8:
                    return filetime_to_dt(value)
                elif len(value) >= 16:
                    return systemtime_to_dt(value)
                else:
                    return "Invalid binary time"
            elif isinstance(value, int):
                return filetime_to_dt(value)
            else:
                return "N/A"

        try:
            self.log("🔍 Parsing network profiles...")
            self.progress.config(mode='indeterminate')
            self.progress.start()

            with open(net_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Hive', 'ProfileName', 'Description', 'DateCreated', 'Managed', 'DateLastConnected'])

                profile_count = 0
                for i in indices:
                    hive_path = self.hives_listbox.get(i)
                    if os.path.basename(hive_path).upper() != "SOFTWARE":
                        continue
                    try:
                        reg = Registry.Registry(hive_path)
                        profiles = reg.open("Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Profiles")
                        for profile in profiles.subkeys():
                            def get_val(name):
                                try:
                                    return profile.value(name).value()
                                except:
                                    return None

                            name = get_val("ProfileName")
                            desc = get_val("Description")
                            date_created = parse_timestamp(get_val("DateCreated"))
                            managed = get_val("Managed")
                            date_last_connected = parse_timestamp(get_val("DateLastConnected"))

                            writer.writerow([os.path.basename(hive_path), name, desc, date_created, managed, date_last_connected])
                            profile_count += 1

                        self.log(f"✅ Network profiles parsed from {os.path.basename(hive_path)}")
                    except Exception as e:
                        self.log(f"❌ Network parse failed for {os.path.basename(hive_path)}: {e}")

            self.log(f"✅ Found {profile_count} network profiles. Output: {net_file}")
        except Exception as e:
            self.log(f"❌ Network parsing failed: {e}")
        finally:
            self.progress.stop()
            self.progress.config(mode='determinate')

        self.status_var.set("Network profile parsing complete.")

    def load_zip_and_scan(self):
        zip_path = filedialog.askopenfilename(
            title="Select ZIP File",
            filetypes=[("ZIP files", "*.zip")]
        )
        if not zip_path:
            return

        # Default base path = current working directory
        default_base_path = os.getcwd()
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        default_extract_path = os.path.join(default_base_path, f"zip_extract_{timestamp}")

        # Create a pop-up dialog to let the user view/change the base path
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Extraction Location")
        dialog.geometry("500x160")
        dialog.resizable(False, False)

        tk.Label(dialog, text="Base Folder for Extraction:", font=("Helvetica", 10)).pack(pady=(10, 0))

        base_path_var = tk.StringVar(value=default_base_path)

        entry = tk.Entry(dialog, textvariable=base_path_var, width=60)
        entry.pack(pady=5)

        def browse_folder():
            selected = filedialog.askdirectory(title="Select Folder")
            if selected:
                base_path_var.set(selected)

        tk.Button(dialog, text="Browse", command=browse_folder).pack()

        def confirm_path():
            base_path = base_path_var.get().strip()
            if not base_path or not os.path.exists(base_path):
                messagebox.showerror("Invalid Path", "Please select a valid folder.")
                return

            dialog.destroy()
            subfolder_name = f"zip_extract_{timestamp}"
            self.temp_zip_dir = os.path.join(base_path, subfolder_name)

            try:
                os.makedirs(self.temp_zip_dir, exist_ok=True)
                self.log(f"📦 Extracting ZIP file: {zip_path}")
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(self.temp_zip_dir)

                self.log("📂 ZIP extracted successfully.")
                self.log(f"📁 Extraction path: {self.temp_zip_dir}")
                self.log("📡 Now scanning hives from extracted content...")

                self.reg_folder_var.set(self.temp_zip_dir)
                self.scan_hives()

            except Exception as e:
                self.log(f"❌ Failed to extract ZIP: {e}")

        tk.Button(dialog, text="Extract", command=confirm_path, bg="#4CAF50", fg="white").pack(pady=10)




    def cleanup_temp_zip(self):
        if self.temp_zip_dir and os.path.exists(self.temp_zip_dir):
            try:
                shutil.rmtree(self.temp_zip_dir)
                self.log(f"🧹 Auto-deleted temp ZIP folder: {self.temp_zip_dir}")
            except Exception as e:
                self.log(f"⚠️ Failed to delete temp folder: {e}")

    def on_close(self):
        if self.temp_zip_dir and os.path.exists(self.temp_zip_dir):
            answer = messagebox.askyesno(
                "Save Extracted Files?",
                "Do you want to keep the extracted ZIP folder?\n\nIf you click No, it will be permanently deleted."
            )
            if not answer:
                self.cleanup_temp_zip()
            else:
                self.log(f"📁 Extracted folder kept: {self.temp_zip_dir}")
        self.root.destroy()

            
        


    def scan_hives(self):
        folder = self.reg_folder_var.get()
        self.hives_listbox.delete(0, tk.END)
        if not folder or not os.path.isdir(folder):
            self.log("⚠️ Select a valid registry hive folder.")
            return

        self.log(f"🔎 Scanning for registry hives in {folder}")

        # Enhanced hive detection with more file types
        known_hive_names = [
            'SYSTEM', 'SOFTWARE', 'SAM', 'SECURITY', 'NTUSER.DAT', 'USRCLASS.DAT', 
            'AMCACHE.HVE', 'DRIVERS', 'usrClass.dat', 'BBI', 'BCD', 'COMPONENTS',
            'DEFAULT', 'ELAM', 'SCHEMA.DAT'
        ]

        hive_count = 0
        for root_dir, dirs, files in os.walk(folder):
            for file in files:
                # Check for known hive names (case-insensitive)
                if any(file.upper() == hive.upper() for hive in known_hive_names):
                    full_path = os.path.join(root_dir, file)
                    self.hives_listbox.insert(tk.END, full_path)
                    hive_count += 1
                # Also check for files without extensions that might be hives
                elif '.' not in file and len(file) > 2:
                    full_path = os.path.join(root_dir, file)
                    # Basic heuristic: check file size (registry hives are typically > 10KB)
                    try:
                        if os.path.getsize(full_path) > 10240:  # 10KB
                            self.hives_listbox.insert(tk.END, full_path)
                            hive_count += 1
                    except:
                        pass

        self.log(f"✅ Found {hive_count} potential registry hives.")
        if hive_count > 0:
            self.log("💡 Tip: Select specific hives or use 'Select All' button.")
        self.log(f"🧾 Total hive entries: {self.hives_listbox.size()}")


def parse_registry_hive(hive_path, output_csv):
    """Enhanced registry hive parser with better error handling"""
    reg = Registry.Registry(hive_path)
    
    with open(output_csv, 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Key Path', 'Value Name', 'Value Type', 'Value Data', 'Last Modified'])
        
        def get_value_type(value):
            """Convert registry value type to readable string"""
            type_map = {
                Registry.RegSZ: "REG_SZ",
                Registry.RegExpandSZ: "REG_EXPAND_SZ", 
                Registry.RegBin: "REG_BINARY",
                Registry.RegDWord: "REG_DWORD",
                Registry.RegMultiSZ: "REG_MULTI_SZ",
                Registry.RegQWord: "REG_QWORD"
            }
            return type_map.get(value.value_type(), f"Unknown({value.value_type()})")
        
        def recursive_parse(key, path=""):
            current_path = path + "\\" + key.name() if path else key.name()
            
            # Parse values in current key
            for value in key.values():
                try:
                    val_name = value.name() or "(Default)"
                    val_type = get_value_type(value)
                    
                    # Handle different data types
                    try:
                        val_data = str(value.value())
                        # Truncate very long binary data
                        if val_type == "REG_BINARY" and len(val_data) > 100:
                            val_data = val_data[:100] + "... (truncated)"
                    except:
                        val_data = "[Error reading value]"
                    
                    last_modified = key.timestamp().strftime("%Y-%m-%d %H:%M:%S")
                    
                    writer.writerow([current_path, val_name, val_type, val_data, last_modified])
                except Exception as e:
                    # Log error but continue processing
                    writer.writerow([current_path, "[Error]", "ERROR", f"Failed to read: {e}", ""])
            
            # Recursively process subkeys
            for subkey in key.subkeys():
                try:
                    recursive_parse(subkey, current_path)
                except Exception as e:
                    # Log error but continue with other subkeys
                    writer.writerow([current_path + "\\" + subkey.name(), "[Error]", "ERROR", 
                                   f"Failed to access subkey: {e}", ""])
        
        recursive_parse(reg.root())

def parse_usb_devices_from_system_hive(hive_path, output_csv):
    """Enhanced USB device parser with more comprehensive data extraction"""
    reg = Registry.Registry(hive_path)
    
    # Try multiple ControlSets for comprehensive coverage
    control_sets = ["ControlSet001", "ControlSet002", "CurrentControlSet"]
    usbstor_key = None
    usb_key = None

    for cs in control_sets:
        if usbstor_key is None:
            try:
                usbstor_key = reg.open(f"{cs}\\Enum\\USBSTOR")
            except Registry.RegistryKeyNotFoundException:
                pass
        if usb_key is None:
            try:
                usb_key = reg.open(f"{cs}\\Enum\\USB")
            except Registry.RegistryKeyNotFoundException:
                pass
        if usbstor_key and usb_key:
            break

    with open(output_csv, 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.writer(csvfile)
        # Enhanced column set with additional forensically relevant fields
        writer.writerow([
            'Type', 'Device ID', 'Instance ID', 'Key Last Modified', 'Device Description', 
            'Friendly Name', 'Service', 'Class GUID', 'Parent ID Prefix', 'Serial Number',
            'Hardware IDs', 'Compatible IDs', 'Driver', 'Manufacturer', 'Location Information'
        ])

        def parse_usb_keys(key, key_type):
            if not key:
                return
            
            device_count = 0
            for device_id_key in key.subkeys():
                device_id = device_id_key.name()
                for instance_key in device_id_key.subkeys():
                    instance_id = instance_key.name()

                    # Enhanced timestamp handling
                    try:
                        last_modified = instance_key.timestamp().strftime("%Y-%m-%d %H:%M:%S UTC")
                    except Exception:
                        last_modified = "N/A"

                    # Comprehensive value extraction with error handling
                    def get_value_safe(key, val_name):
                        try:
                            value = key.value(val_name).value()
                            if isinstance(value, list):
                                return "; ".join(str(v) for v in value)
                            return str(value)
                        except Registry.RegistryValueNotFoundException:
                            return ""
                        except Exception:
                            return "[Error reading value]"

                    device_desc = get_value_safe(instance_key, "DeviceDesc")
                    friendly_name = get_value_safe(instance_key, "FriendlyName")
                    service = get_value_safe(instance_key, "Service")
                    class_guid = get_value_safe(instance_key, "ClassGUID")
                    parent_id_prefix = get_value_safe(instance_key, "ParentIdPrefix")
                    hardware_ids = get_value_safe(instance_key, "HardwareID")
                    compatible_ids = get_value_safe(instance_key, "CompatibleIDs")
                    driver = get_value_safe(instance_key, "Driver")
                    manufacturer = get_value_safe(instance_key, "Mfg")
                    location_info = get_value_safe(instance_key, "LocationInformation")

                    # Enhanced serial number extraction
                    serial_number = instance_id  # Default to instance ID
                    serial_number_val = get_value_safe(instance_key, "SerialNumber")
                    if serial_number_val:
                        serial_number = serial_number_val

                    writer.writerow([
                        key_type, device_id, instance_id, last_modified, device_desc,
                        friendly_name, service, class_guid, parent_id_prefix, serial_number,
                        hardware_ids, compatible_ids, driver, manufacturer, location_info
                    ])
                    device_count += 1
            
            return device_count

        usbstor_count = 0
        usb_count = 0
        
        if usbstor_key:
            usbstor_count = parse_usb_keys(usbstor_key, "USBSTOR") or 0
        else:
            writer.writerow(["USBSTOR", "No devices found or key missing", "", "", "", "", "", "", "", "", "", "", "", "", ""])

        if usb_key:
            usb_count = parse_usb_keys(usb_key, "USB") or 0
        else:
            writer.writerow(["USB", "No devices found or key missing", "", "", "", "", "", "", "", "", "", "", "", "", ""])

    return usbstor_count + usb_count



if __name__ == "__main__":
    root = tk.Tk()
    app = ForensicParserApp(root)
    root.mainloop()