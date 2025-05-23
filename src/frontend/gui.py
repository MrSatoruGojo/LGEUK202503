# src/frontend/gui.py
"""
Desktop GUI for selecting input and output files using Tkinter.
Supports browsing Excel files including .xlsx, .xls, .xlsb, CSV, JSON.
Remembers prior selections, shows live progress & logs, and notifies when done.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import json
import os
import threading
import traceback
from datetime import datetime
from pathlib import Path

from tksheet import Sheet
from src.main import main
import pandas as pd

import config.config as cfg

# Optional desktop toast
try:
    from plyer import notification
except ImportError:
    notification = None


class FileSelectorApp:
    def __init__(self):
        # Main window
        self.root = tk.Tk()
        # right after self.root = tk.Tk():
        self.show_intermediates = tk.BooleanVar(value=False)
         # Allow window to be resized and fill large screens
        self.root.resizable(True, True)
        try:
            self.root.state('zoomed')
        except:
            pass
        self.root.title("Claim Verificator")
       

        # Storage for file selections
        self.file_paths = {}
        # These keys must match cfg.INPUT_PARTS
        self.fields = {key: label for key, label in cfg.INPUT_PARTS}

        # Build UI
        self.create_widgets()
        self.load_previous_selections()

        # Start the GUI loop
        self.root.mainloop()

    
    def create_widgets(self):
    # ─── Scrollable file-pickers ─────────────────────────────────────────
        container = tk.Frame(self.root)
        container.grid(row=0, column=0, columnspan=2, sticky='nsew')
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        canvas = tk.Canvas(container)
        scrollbar = tk.Scrollbar(container, orient="vertical",
                                command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0,0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Place each file-picker
        for idx, (key, label) in enumerate(self.fields.items()):
            btn = tk.Button(scroll_frame, text=label, width=25,
                            command=lambda k=key: self.select_path(k))
            btn.grid(row=idx, column=0, padx=10, pady=5, sticky='w')
            lbl = tk.Label(scroll_frame, text="[Not selected]", width=40,
                        anchor='w', wraplength=350)
            lbl.grid(row=idx, column=1, padx=10, pady=5, sticky='w')
                        # configure-sheet button (disabled until a file is chosen)
            cfg_btn = tk.Button(scroll_frame, text="⚙️",
                                command=lambda k=key: self.open_sheet_config(k),
                                state='disabled', width=3)
            cfg_btn.grid(row=idx, column=2, padx=(0,10), pady=5)

             # record into our dict

            self.file_paths[key] = {
                'widget': lbl,
                'value': None,
                'config_btn': cfg_btn,
                'sheets': {}       # <-- use the same dict key everywhere
            }
    
            


        base_row = len(self.fields)  # first free row below scroll area

# ─── Notebook for previews ───────────────────────────────────────────
        self.preview_nb = ttk.Notebook(self.root)
        self.preview_nb.grid(row=base_row, column=0, columnspan=2,
                         sticky='nsew', padx=10, pady=(10,5))

    # ─── Control buttons ─────────────────────────────────────────────────
        term_btn = tk.Button(self.root, text="Edit Terminology",
                         command=self.open_terminology_editor,
                         bg='#2196F3', fg='white', width=25)
        term_btn.grid(row=base_row+1, column=0, columnspan=2, pady=(5,10))

        self.exec_btn = tk.Button(self.root, text="Run Processing",
                              command=self.run_processing,
                              bg='#4CAF50', fg='white', width=25)
        self.exec_btn.grid(row=base_row+2, column=0, padx=10, pady=(0,10))
        exit_btn = tk.Button(self.root, text="Exit",
                         command=self.root.destroy,
                         bg='#f44336', fg='white', width=25)
        exit_btn.grid(row=base_row+2, column=1, padx=10, pady=(0,10))

        clear_btn = tk.Button(self.root, text="Clear Selections",
                          command=self.reset_selections,
                          bg='#FF9800', fg='white', width=25)
        clear_btn.grid(row=base_row+3, column=0, columnspan=2,
                   padx=10, pady=(0,10))

    # ─── Checkbox to toggle previews ──────────────────────────────────────
        chk = tk.Checkbutton(self.root,
                         text="Show intermediate tables",
                         variable=self.show_intermediates)
        chk.grid(row=base_row+4, column=0, columnspan=2, pady=(0,10))

    # ─── Progress & log area ─────────────────────────────────────────────
        self.progress = ttk.Progressbar(self.root,
                                    orient='horizontal',
                                    length=600,
                                    mode='determinate',
                                    maximum=7)
        self.progress.grid(row=base_row+5, column=0, columnspan=2,
                       padx=10, pady=(5,2))

        self.status_lbl = tk.Label(self.root, text="Ready", anchor='w')
        self.status_lbl.grid(row=base_row+6, column=0, columnspan=2,
                         padx=10, pady=(0,10), sticky='w')

        tk.Label(self.root, text="Log Output:").grid(
            row=base_row+7, column=0, columnspan=2,
            padx=10, pady=(0,0), sticky='w'
        )
        self.log_text = scrolledtext.ScrolledText(
        self.root, width=80, height=12, state='disabled', wrap='word'
        )
        self.log_text.grid(row=base_row+8, column=0, columnspan=2,
                       padx=10, pady=(0,10), sticky='nsew')
        self.root.grid_rowconfigure(base_row+8, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

# --- end of create_widgets() ---


    def load_previous_selections(self):
        """Load saved file paths and update labels."""
        config_path = Path.home() / ".claim_verificator_config.json"
        if not config_path.exists():
            return
        try:
            data = json.loads(config_path.read_text())
            for key, info in data.items():
                if key not in self.file_paths:
                    continue
                # value could be either a single path or a dict with sheet/header
                if isinstance(info, dict):
                    path   = info.get('path')
                    sheets = info.get('sheets', {})
                else:
                    path   = info
                    sheets = {}

                if not path:
                    continue

                self.file_paths[key]['value']  = path
                self.file_paths[key]['sheets'] = sheets
                self.file_paths[key]['config_btn'].config(state='normal')

                # rebuild display: show filename and any sheet@row selections
                disp = os.path.basename(path)
                if sheets:
                   disp += " ▾" + ",".join(f"{s}@r{r}" for s, r in sheets.items())
                self.file_paths[key]['widget'].config(text=disp)
                
        except Exception:
            pass

    def save_selections(self):
        """Persist current file_paths to JSON."""
        try:
            config_path = Path.home() / ".claim_verificator_config.json"
            out = {}
            for k, v in self.file_paths.items():
                path   = v['value']
                sheets = v.get('sheets', {})

                if not path:
                    out[k] = None
                elif sheets:
                    out[k] = {'path': path, 'sheets': sheets}
                else:
                    out[k] = path
            config_path.write_text(json.dumps(out, indent=2))
        except Exception:
            pass

    def select_path(self, key):
        """Open the appropriate dialog for the given input key."""
        last = self.file_paths[key]['value']
        initialdir = os.getcwd()
        initialfile = None
        if last:
            if isinstance(last, list) and last:
                initialdir = os.path.dirname(last[0])
            elif isinstance(last, str):
                initialdir = os.path.dirname(last)
                initialfile = os.path.basename(last)

        # Closed orders → multi‐select
        if key.startswith('CLOSED_ORDERS_'):
            paths = filedialog.askopenfilenames(
                title=self.fields[key],
                filetypes=[('Excel/CSV/JSON', '*.xls;*.xlsx;*.xlsb;*.csv;*.json'),
                           ('All files','*.*')],
                initialdir=initialdir
            )
            path = list(paths) if paths else None

        # OUTPUT_DIR → pick directory
        elif key == 'OUTPUT_DIR':
            path = filedialog.askdirectory(
                title=self.fields[key],
                initialdir=initialdir
            )

        # Otherwise → single file
        else:
            file = filedialog.askopenfilename(
                title=self.fields[key],
                filetypes=[('Excel/CSV/JSON', '*.xls;*.xlsx;*.xlsb;*.csv;*.json'),
                           ('All files','*.*')],
                initialdir=initialdir,
                initialfile=initialfile
            )
            path = file or None

        if path:
                        # 1) store & persist
            self.file_paths[key]['value'] = path
            self.save_selections()

            # 2) figure out extension (first item if it's a list)
            if isinstance(path, (list, tuple)) and path:
                ext = os.path.splitext(path[0])[1].lower()
            else:
                ext = os.path.splitext(path)[1].lower()

            # 3) enable/disable the ⚙️ configure-sheet button
            if ext in ('.xls', '.xlsx', '.xlsb', '.csv', '.json'):
                self.file_paths[key]['config_btn'].config(state='normal')
            else:
                self.file_paths[key]['config_btn'].config(state='disabled')

            # 4) update display text
            if isinstance(path, (list, tuple)):
                disp = f"{len(path)} files selected"
            else:
                disp = os.path.basename(path)
            self.file_paths[key]['widget'].config(text=disp)



    def configure_sheet(self, key):
        # TODO: pop up a window that:
        #   - opens self.file_paths[key]['value']
        #   - lists sheets, previews first few rows
        #   - lets user pick sheet name(s) + header row
        #   - stores those choices (e.g. in self.file_paths[key]['sheet_info'])
        pass

    def _log(self, message: str):
        """Append a timestamped message to the log widget."""
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state='normal')
        self.log_text.insert('end', f"[{ts}] {message}\n")
        self.log_text.configure(state='disabled')
        self.log_text.yview('end')

    def run_processing(self):
        """Launch the main() pipeline on a background thread."""
         # 1) Validate required inputs
        required = ['CLAIM_FILE', 'SPMS_FILE', 'PSI_FILE', 'CLOSED_ORDERS_DIR']
        missing_req = [k for k in required if not self.file_paths[k]['value']]
        if missing_req:
            messagebox.showerror(
                "Missing Required Inputs",
                f"Please select: {', '.join(self.fields[k] for k in missing_req)}"
            )
            return

        # Warn about missing optional inputs (but continue)
        optional = [k for k in self.file_paths if k not in required]
        missing_opt = [k for k in optional if not self.file_paths[k]['value']]
        if missing_opt:
            messagebox.showwarning(
                "Optional Inputs Missing",
                "The following optional inputs are not selected and will be skipped:\n"
                + "\n".join(f"- {self.fields[k]}" for k in missing_opt)
            )

        # Persist into cfg
        for k, info in self.file_paths.items():
            # 1) always store the path
            setattr(cfg, k, info['value'])
            # 2) if the user chose one or more sheets, store that dict
            sheets = info.get('sheets', {})
            if sheets:
                setattr(cfg, f"{k}_SHEETS", sheets)


        # Disable Run
        self.exec_btn.config(state='disabled')
        self._log("Starting processing…")

        def gui_progress(label, step, elapsed, df=None):
            self.progress['value'] = step
            status = f"{step}/7: {label} — {elapsed:.1f}s"
            self.status_lbl.config(text=status)
            self._log(status)
            # if toggled, preview this DataFrame
            if self.show_intermediates.get() and df is not None:
                self.root.after(0, lambda: self._preview_df(label, df))
            
                # fire off the pipeline using the imported main()
        threading.Thread(
            target=lambda: main(progress_callback=gui_progress),
            daemon=True
        ).start()
            

        def worker():
            try:
                from src.main import main
                start = datetime.now()
                main(progress_callback=gui_progress)
                total = (datetime.now() - start).total_seconds()
                self._log(f"All steps finished in {total:.1f}s")
                self.status_lbl.config(text="Completed!")
                try:
                    self.root.bell()
                except:
                    pass
                if notification:
                    notification.notify(
                        title="Claim Verificator",
                        message=f"Finished in {total:.1f} seconds",
                        timeout=5
                    )
                messagebox.showinfo(
                    "Done", f"Finished in {total:.1f} seconds"
                )
            except Exception:
                err = traceback.format_exc()
                self._log("ERROR during processing:")
                self._log(err)
                messagebox.showerror("Error", "An error occurred. See log.")
            finally:
                self.exec_btn.config(state='normal')

        threading.Thread(target=worker, daemon=True).start()
        
    def reset_selections(self):
        """Clear all selected paths and update the UI."""
        for key, info in self.file_paths.items():
            info['value'] = None
            info['widget'].config(text='[Not selected]')
        # Persist the cleared state
        try:
            config_path = Path.home() / ".claim_verificator_config.json"
            empty = {k: None for k in self.file_paths}
            config_path.write_text(json.dumps(empty, indent=2))
        except Exception:
            pass

    def open_terminology_editor(self):
        """Edit your Sub-category / Name table in an Excel-like grid."""
        # Sample initial data; replace with your actual draft
        data = {
            "Sub-category": [
                "Inputs", "Inputs", "Inputs",
                "Closed Orders", "Closed Orders", "Closed Orders",
                # ...
            ],
            "Name": [
                "Claim File", "SPMS Files", "PSI File",
                "Closed Orders 2020", "Closed Orders 2021", "Closed Orders 2022",
                # ...
            ]
        }
        df = pd.DataFrame(data)

        # Popup window
        win = tk.Toplevel(self.root)
        win.title("Terminology Editor")
        win.geometry("600x400")

        # tksheet grid
        sheet = Sheet(
            win,
            data=df.values.tolist(),
            headers=list(df.columns),
            width=580, height=330,
            show_row_index=False
        )
        sheet.grid(row=0, column=0, padx=10, pady=10)

        def save_terms():
            new_data = sheet.get_sheet_data(return_copy=True)
            new_df = pd.DataFrame(new_data[1:], columns=new_data[0])
            new_df.to_excel("terminology.xlsx", index=False)
            messagebox.showinfo("Saved", "Terminology saved to terminology.xlsx")
            win.destroy()

        save_btn = tk.Button(win, text="Save", command=save_terms,
                             bg='#4CAF50', fg='white', width=20)
        save_btn.grid(row=1, column=0, pady=(0,10))


   

    # Updated `open_sheet_config` Method
  
    def open_sheet_config(self, key):
        """Let the user pick one or more sheets + header rows, then preview."""
        path = self.file_paths[key]['value']
        if not path:
            return

        try:
            xlsx = pd.ExcelFile(path)
            sheets = xlsx.sheet_names
        except Exception:
            messagebox.showerror("Error", f"Cannot read sheets from {path}")
            return

        win = tk.Toplevel(self.root)
        # allow normal window decorations (minimize/maximize/close) 
        win.resizable(True, True)
        # on Windows, start maximized
        try:
            win.state('zoomed')
        except Exception:
            pass

        win.title(f"Configure “{os.path.basename(path)}”")
        win.geometry("700x600")
        win.resizable(True, True)

        # Variables per sheet
        use_vars = {}
        header_vars = {}

        # Container for sheet configs
        cfg_frame = tk.Frame(win)
        cfg_frame.pack(fill="x", padx=10, pady=10)

        # Header row labels
        tk.Label(cfg_frame, text="Sheet", width=30, anchor="w").grid(row=0, column=0)
        tk.Label(cfg_frame, text="Use?", width=5).grid(row=0, column=1)
        tk.Label(
            cfg_frame,
            text="Header row (Excel row containing your column names):",
            width=30,
            anchor="w"
        ).grid(row=0, column=2, padx=5)


        # Create one row per sheet
        for i, sheet_name in enumerate(sheets, start=1):
            # 1) Sheet name (clickable)
            lbl = tk.Label(cfg_frame, text=sheet_name, fg="blue", cursor="hand2")
            lbl.grid(row=i, column=0, sticky="w", pady=2)
            lbl.bind("<Button-1>", lambda e, s=sheet_name: load_preview(s))

            # 2) Use? checkbox
            use_vars[sheet_name] = tk.BooleanVar(value=False)
            chk = tk.Checkbutton(cfg_frame, variable=use_vars[sheet_name])
            chk.grid(row=i, column=1)

            # 3) Header-row dropdown (1–5)
            header_vars[sheet_name] = tk.StringVar(value="1")
            combo = ttk.Combobox(
                cfg_frame,
                textvariable=header_vars[sheet_name],
                values=[str(n) for n in range(1, 6)],
                width=5,
                state='readonly'
            )
            combo.grid(row=i, column=2, padx=5)
            combo.bind("<<ComboboxSelected>>", lambda e, s=sheet_name: load_preview(s))
        # Preview area (raw, no header row)
        # ─── Preview area ─────────────────────────────────
        preview_frame = tk.Frame(win)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=(0,10))

        # 3) Label showing which sheet is being previewed
        preview_label = tk.Label(preview_frame, text="", anchor="w")
        preview_label.pack(fill="x")

        preview = Sheet(preview_frame,
                        height=300, width=650,
                        show_row_index=True)  # keep row-index on left
        preview.pack(fill="both", expand=True)

        # Maximize this preview window
        try:
            win.state('zoomed')
        except:
            pass

        def load_preview(sheet):
            try:
                # always read raw rows (no header) so user can see exactly what's on row 1,2,3…
                # read absolutely raw so row 1 in Excel is row 1 in preview
                raw = pd.read_excel(path, sheet_name=sheet, header=None)
                data = raw.head(20).values.tolist()

                # build Excel-style column headers A, B, C… for as many columns as you have
                def excel_col(n):
                    s = ""
                    while n >= 0:
                        s = chr(n % 26 + ord('A')) + s
                        n = n // 26 - 1
                    return s
                cols = [ excel_col(i) for i in range(raw.shape[1]) ]

                # 1) set the column-letter headers:
                preview.headers(cols)

                # 2) populate exactly the raw rows 1–20:
                preview.set_sheet_data(data)

            except Exception as err:
                messagebox.showerror("Preview failed", str(err))

        # Bottom buttons
        btn_frame = tk.Frame(win)
        btn_frame.pack(fill="x", pady=(0,10))
        tk.Button(btn_frame, text="Cancel",
              command=win.destroy).pack(side="right", padx=5)
        def save_and_close():
            # Collect only chosen sheets
            chosen = [
                (s, int(header_vars[s].get()))
                for s, v in use_vars.items() if v.get()
            ]
            # new:
            self.file_paths[key]['sheets'] = { s: hdr for s, hdr in chosen }
            self.save_selections()   # immediately persist
            # Update display to show e.g. "f.xlsx ▾S1@r4,S2@r2"
            disp = os.path.basename(path)
            if chosen:
                disp += " ▾" + ",".join(f"{s}@r{r}" for s,r in chosen)
            self.file_paths[key]['widget'].config(text=disp)
            self.save_selections()
            win.destroy()

        tk.Button(btn_frame, text="OK", bg='#4CAF50', fg='white',
                command=save_and_close).pack(side="right")

        # Initial preview of first sheet
        if sheets:
            load_preview(sheets[0])
    

    def _preview_df(self, label, df):
        """Add or replace a Notebook tab showing df.head(50)."""
    # Remove any existing tab with this label
        for tab_id in self.preview_nb.tabs():
            if self.preview_nb.tab(tab_id, 'text') == label:
                self.preview_nb.forget(tab_id)
                break

        frame = tk.Frame(self.preview_nb)
        self.preview_nb.add(frame, text=label)

        import pandas as _pd
        if not isinstance(df, _pd.DataFrame):
            self._log(f"Skipping preview for “{label}”: not a DataFrame.")
            return
        data = df.head(50).values.tolist()
        headers = list(df.columns)
        sheet = Sheet(frame,
                  data=data,
                  headers=headers,
                  width=600,
                  height=200,
                  show_row_index=True)
        sheet.grid(row=0, column=0, sticky='nsew')
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

if __name__ == "__main__":
    FileSelectorApp()
