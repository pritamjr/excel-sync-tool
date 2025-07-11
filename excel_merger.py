import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import time
import os
import json
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import hashlib

# Codespaces-specific setup
if os.environ.get('CODESPACES') == 'true':
    # Set default paths for cloud environment
    DEFAULT_SOURCE = "/workspaces/excel-sync-tool/source.xlsx"
    DEFAULT_TARGET = "/workspaces/excel-sync-tool/target.xlsx"
    
    # Create sample files if they don't exist
    if not os.path.exists(DEFAULT_SOURCE):
        pd.DataFrame([["Sample", "Data"]], columns=["Name", "Value"]).to_excel(DEFAULT_SOURCE, index=False)
    if not os.path.exists(DEFAULT_TARGET):
        pd.DataFrame([["Sample", ""]], columns=["Name", "Value"]).to_excel(DEFAULT_TARGET, index=False)

CONFIG_FILE = "excel_sync_config.json"

class ExcelSyncApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Auto-Sync Tool v3.0")
        
        # Setup GUI
        self.create_widgets()
        
        # Sync variables
        self.source_path = ""
        self.target_path = ""
        self.observer = None
        self.sync_active = False
        self.last_hash = None
        self.last_sync_time = 0
        
        # Load previous selections
        self.load_config()

    def create_widgets(self):
        # Source Selection
        tk.Label(self.root, text="Source Excel (Master File):").pack(pady=(10,0))
        self.source_frame = tk.Frame(self.root)
        self.source_frame.pack()
        self.source_entry = tk.Entry(self.source_frame, width=50)
        self.source_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(self.source_frame, text="Browse", command=self.select_source).pack(side=tk.LEFT)
        
        # Target Selection
        tk.Label(self.root, text="Target Excel (To Update):").pack(pady=(10,0))
        self.target_frame = tk.Frame(self.root)
        self.target_frame.pack()
        self.target_entry = tk.Entry(self.target_frame, width=50)
        self.target_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(self.target_frame, text="Browse", command=self.select_target).pack(side=tk.LEFT)
        
        # Sync Controls
        self.control_frame = tk.Frame(self.root)
        self.control_frame.pack(pady=10)
        self.sync_btn = tk.Button(self.control_frame, text="Start Auto-Sync", command=self.toggle_sync)
        self.sync_btn.pack(side=tk.LEFT, padx=5)
        self.manual_sync_btn = tk.Button(self.control_frame, text="Sync Now", command=self.manual_sync)
        self.manual_sync_btn.pack(side=tk.LEFT)
        
        # Status Area
        self.status_var = tk.StringVar()
        self.status_var.set("Status: Waiting to start sync")
        tk.Label(self.root, textvariable=self.status_var).pack(pady=(10,0))
        
        # Logging
        self.log_text = tk.Text(self.root, height=10, width=70, state=tk.DISABLED)
        self.log_text.pack(pady=10, padx=10)
        self.log("Application started. Select files to begin.")

    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def select_source(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if path:
            self.source_path = path
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, path)
            self.check_ready()
            self.last_hash = self.get_file_hash(path)
            self.log(f"Source file set: {os.path.basename(path)}")

    def select_target(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if path:
            self.target_path = path
            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, path)
            self.check_ready()
            self.log(f"Target file set: {os.path.basename(path)}")

    def get_file_hash(self, filepath):
        """Generate a hash of the file contents for change detection"""
        try:
            with open(filepath, 'rb') as f:
                return hashlib.md5(f.read()).hexdigest()
        except:
            return None

    def check_ready(self):
        if self.source_path and self.target_path:
            self.save_config()
            self.manual_sync_btn.config(state=tk.NORMAL)

    def save_config(self):
        with open(CONFIG_FILE, 'w') as f:
            json.dump({
                'source': self.source_path,
                'target': self.target_path
            }, f)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    self.source_path = config.get('source', '')
                    self.target_path = config.get('target', '')
                    
                    if self.source_path:
                        self.source_entry.insert(0, self.source_path)
                        self.last_hash = self.get_file_hash(self.source_path)
                    if self.target_path:
                        self.target_entry.insert(0, self.target_path)
                    
                    self.log("Loaded previous configuration")
            except Exception as e:
                self.log(f"Error loading config: {str(e)}")

    def toggle_sync(self):
        if not self.sync_active:
            self.start_sync()
        else:
            self.stop_sync()

    def start_sync(self):
        if not all(os.path.exists(f) for f in [self.source_path, self.target_path]):
            messagebox.showerror("Error", "One or both files no longer exist")
            return

        self.sync_active = True
        self.sync_btn.config(text="Stop Auto-Sync")
        self.status_var.set("Status: ACTIVE - Monitoring for changes")
        self.log("Auto-sync started. Monitoring for changes...")

        # Start the file observer in a separate thread
        self.observer = Observer()
        event_handler = SyncHandler(self)
        self.observer.schedule(event_handler, path=os.path.dirname(self.source_path))
        self.observer.start()

        # Start periodic checking as backup
        self.root.after(5000, self.periodic_check)

    def periodic_check(self):
        """Backup check in case file system events are missed"""
        if self.sync_active:
            current_hash = self.get_file_hash(self.source_path)
            if current_hash and current_hash != self.last_hash:
                self.log("Periodic check detected changes")
                self.perform_sync()
                self.last_hash = current_hash
            self.root.after(5000, self.periodic_check)

    def stop_sync(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
        self.sync_active = False
        self.sync_btn.config(text="Start Auto-Sync")
        self.status_var.set("Status: INACTIVE")
        self.log("Auto-sync stopped")

    def manual_sync(self):
        self.log("Manual sync initiated...")
        threading.Thread(target=self.perform_sync, daemon=True).start()

    def perform_sync(self):
        try:
            if time.time() - self.last_sync_time < 2:  # 2-second cooldown
                return

            start_time = time.time()
            self.last_sync_time = time.time()
            self.root.after(0, lambda: self.status_var.set("Status: Syncing..."))
            
            # Read files
            df_source = pd.read_excel(self.source_path)
            df_target = pd.read_excel(self.target_path)
            
            # Process duplicates (keep last)
            df_source = df_source.drop_duplicates(subset=[df_source.columns[0]], keep='last')
            source_map = df_source.set_index(df_source.columns[0]).to_dict('index')
            
            # Update target
            update_count = 0
            for idx, row in df_target.iterrows():
                name = row[0]
                if name in source_map:
                    for col in df_target.columns[1:]:
                        if col in source_map[name]:
                            old_value = df_target.at[idx, col]
                            new_value = source_map[name][col]
                            if pd.isna(old_value) or old_value != new_value:
                                df_target.at[idx, col] = new_value
                                update_count += 1
            
            if update_count > 0:
                df_target.to_excel(self.target_path, index=False)
                elapsed = time.time() - start_time
                self.root.after(0, lambda: self.status_var.set(
                    f"Status: Synced {update_count} changes ({elapsed:.2f}s)"
                ))
                self.log(f"Synced {update_count} cells to target")
                self.last_hash = self.get_file_hash(self.source_path)
            else:
                self.root.after(0, lambda: self.status_var.set("Status: No changes needed"))
                self.log("No changes detected in source file")
                
        except PermissionError:
            self.root.after(0, lambda: messagebox.showerror(
                "Error", 
                "Please close the target Excel file before syncing"
            ))
            self.log("Sync failed: Target file is locked")
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror(
                "Error", 
                f"Sync failed: {str(e)}"
            ))
            self.log(f"Sync error: {str(e)}")

    def on_close(self):
        self.stop_sync()
        self.root.destroy()

class SyncHandler(FileSystemEventHandler):
    def __init__(self, app):
        self.app = app
        self.last_trigger = 0
    
    def on_modified(self, event):
        if not event.is_directory and event.src_path == self.app.source_path:
            current_time = time.time()
            if current_time - self.last_trigger > 3:  # 3-second cooldown
                self.last_trigger = current_time
                new_hash = self.app.get_file_hash(event.src_path)
                if new_hash and new_hash != self.app.last_hash:
                    self.app.log("File system event detected changes")
                    self.app.perform_sync()
                    self.app.last_hash = new_hash

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSyncApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()