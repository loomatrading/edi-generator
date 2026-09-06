import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class CUSCARGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CUSCAR EDI / TXT Generator")
        self.root.geometry("850x600")
        
        # Data Variables
        self.excel_file_path = ""
        
        # Build UI
        self.create_widgets()

    def create_widgets(self):
        # Header Frame
        header_frame = ttk.Frame(self.root, padding=10)
        header_frame.pack(fill=tk.X)
        
        ttk.Label(header_frame, text="CUSCAR EDI / TXT Generator", font=("Arial", 16, "bold")).pack(anchor=tk.W)
        
        # File Selection Frame
        file_frame = ttk.LabelFrame(self.root, text="Select Excel File", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=70, state="readonly").pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="Browse Excel", command=self.load_excel).pack(side=tk.LEFT, padx=5)

        # Manifest Details Frame
        details_frame = ttk.LabelFrame(self.root, text="Manifest Details", padding=10)
        details_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Fields (Registration Number Removed)
        self.fields = {
            "Vessel Name": tk.StringVar(),
            "Voyage Number": tk.StringVar(),
            "Port of Loading": tk.StringVar(),
            "Port of Discharge": tk.StringVar(),
            "Carrier / Agent": tk.StringVar(),
            "Total Packages": tk.StringVar(),
            "Total Weight (KG)": tk.StringVar()
        }

        row = 0
        for label_text, var in self.fields.items():
            ttk.Label(details_frame, text=f"{label_text}:", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky=tk.W, pady=5, padx=5)
            ttk.Entry(details_frame, textvariable=var, width=50).grid(row=row, column=1, sticky=tk.W, pady=5, padx=5)
            row += 1

        # Action Buttons Frame
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Generate CUSCAR TXT", command=self.generate_cuscar_txt).pack(side=tk.RIGHT, padx=5)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return

        self.excel_file_path = file_path
        self.file_path_var.set(file_path)

        # Reset / Clear old values before loading new file
        self.reset_fields()

        # Parse Excel with smart keyword mapping
        try:
            self.parse_excel_smart(file_path)
        except Exception as e:
            messagebox.showerror("Error Reading File", f"Failed to parse Excel file:\n{str(e)}")

    def reset_fields(self):
        for var in self.fields.values():
            var.set("")

    def parse_excel_smart(self, file_path):
        """
        Smart heuristic parser to search Excel sheets for keywords and auto-extract manifest info.
        """
        xls = pd.ExcelFile(file_path)
        
        extracted = {
            "Vessel Name": "",
            "Voyage Number": "",
            "Port of Loading": "",
            "Port of Discharge": "",
            "Carrier / Agent": "",
            "Total Packages": "",
            "Total Weight (KG)": ""
        }

        # Multilingual & flexible keywords dictionary
        keywords = {
            "Vessel Name": ["vessel", "ship", "vessel name", "ship name", "اسم السفينة", "المركب"],
            "Voyage Number": ["voyage", "voy", "voyage no", "رقم الرحلة"],
            "Port of Loading": ["pol", "port of loading", "loading port", "ميناء التحميل"],
            "Port of Discharge": ["pod", "port of discharge", "discharge port", "ميناء التفريغ"],
            "Carrier / Agent": ["carrier", "line", "agent", "shipping line", "الوكيل"],
            "Total Packages": ["total packages", "packages", "pkgs", "إجمالي الطرود"],
            "Total Weight (KG)": ["total weight", "gross weight", "weight", "الوزن"]
        }

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            
            for r in range(min(len(df), 35)): # Scan first 35 rows
                for c in range(min(len(df.columns), 15)):
                    val = str(df.iloc[r, c]).strip().lower() if pd.notna(df.iloc[r, c]) else ""
                    if not val:
                        continue

                    for field, keys in keywords.items():
                        if not extracted[field] and any(k in val for k in keys):
                            # Try getting value from adjacent right cell
                            if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c + 1]):
                                extracted[field] = str(df.iloc[r, c + 1]).strip()
                            elif r + 1 < len(df) and pd.notna(df.iloc[r + 1, c]):
                                extracted[field] = str(df.iloc[r + 1, c]).strip()

        # Update GUI fields with extracted values
        for field, val in extracted.items():
            if val:
                self.fields[field].set(val)

    def generate_cuscar_txt(self):
        vessel_name = self.fields["Vessel Name"].get().strip() or "MANIFEST"
        
        # Clean vessel name to use as default filename
        clean_filename = "".join([c for c in vessel_name if c.isalnum() or c in (' ', '_', '-')]).rstrip()
        default_filename = f"{clean_filename}.txt"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=default_filename,
            filetypes=[("Text Files", "*.txt")]
        )

        if not save_path:
            return

        # Build CUSCAR TXT Structure
        content = []
        content.append("UNB+UNOA:2+SENDER+RECEIVER'")
        content.append("UNH+1+CUSCAR:D:95B:UN'")
        content.append(f"BGM+850+{self.fields['Voyage Number'].get()}+9'")
        content.append(f"TDT+20+{self.fields['Voyage Number'].get()}+1++++{self.fields['Vessel Name'].get()}'")
        content.append(f"LOC+7+{self.fields['Port of Loading'].get()}'")
        content.append(f"LOC+11+{self.fields['Port of Discharge'].get()}'")
        content.append(f"NAD+CA+{self.fields['Carrier / Agent'].get()}'")
        content.append(f"CNT+11:{self.fields['Total Packages'].get()}'")
        content.append(f"CNT+15:{self.fields['Total Weight (KG)'].get()}'")
        content.append("UNT+9+1'")
        content.append("UNZ+1+1'")

        with open(save_path, "w", encoding="utf-8") as f:
            f.write("\n".join(content))

        # Quiet save: No alert box shown on success

if __name__ == "__main__":
    root = tk.Tk()
    app = CUSCARGeneratorApp(root)
    root.mainloop()
