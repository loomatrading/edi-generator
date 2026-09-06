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
        
        self.excel_file_path = ""
        self.create_widgets()

    def create_widgets(self):
        # Header
        header_frame = ttk.Frame(self.root, padding=10)
        header_frame.pack(fill=tk.X)
        ttk.Label(header_frame, text="CUSCAR EDI / TXT Generator", font=("Arial", 16, "bold")).pack(anchor=tk.W)
        
        # File Selection
        file_frame = ttk.LabelFrame(self.root, text="Select Excel File", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=70, state="readonly").pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="Browse Excel", command=self.load_excel).pack(side=tk.LEFT, padx=5)

        # Manifest Details
        details_frame = ttk.LabelFrame(self.root, text="Manifest Details", padding=10)
        details_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

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

        # Action Buttons
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="Generate CUSCAR TXT", command=self.generate_cuscar_txt).pack(side=tk.RIGHT, padx=5)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return

        self.excel_file_path = file_path
        self.file_path_var.set(file_path)

        # Always reset fields before loading new file
        self.reset_fields()

        try:
            self.parse_excel_smart(file_path)
        except Exception as e:
            messagebox.showerror("Error Reading File", f"Failed to parse Excel file:\n{str(e)}")

    def reset_fields(self):
        for var in self.fields.values():
            var.set("")

    def parse_excel_smart(self, file_path):
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

        # 1. Try Tabular Detection First (For Manifest Lists like VARADA.xls)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            
            # Find the header row by searching for common column titles
            header_row_idx = -1
            for r in range(min(15, len(df))):
                row_vals = [str(x).strip().lower() for x in df.iloc[r].tolist() if pd.notna(x)]
                if any(k in row_vals for k in ['vessel name', 'vessel', 'b/l no', 'container no.']):
                    header_row_idx = r
                    break
            
            if header_row_idx != -1:
                # Read DataFrame with detected header
                data_df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_row_idx)
                # Clean column names
                data_df.columns = [str(c).replace('\n', ' ').strip() for c in data_df.columns]
                
                # Column Name Mapping Candidates
                vessel_cols = [c for c in data_df.columns if 'vessel name' in c.lower() or c.lower() == 'vessel']
                voyage_cols = [c for c in data_df.columns if 'voyage' in c.lower() or 'voy' in c.lower()]
                pol_cols = [c for c in data_df.columns if 'pol' in c.lower() or 'loading' in c.lower()]
                pod_cols = [c for c in data_df.columns if 'pod' in c.lower() or 'discharge' in c.lower()]
                carrier_cols = [c for c in data_df.columns if any(k in c.lower() for k in ['carrier', 'agent', 'line', 'shipper'])]
                pkg_cols = [c for c in data_df.columns if 'packages' in c.lower() or 'pkgs' in c.lower()]
                weight_cols = [c for c in data_df.columns if 'gross weight' in c.lower() or 'gross & tare' in c.lower() or 'weight' in c.lower()]

                # Extract first non-null values for Vessel, Voyage, POL, POD
                if vessel_cols and not extracted["Vessel Name"]:
                    vals = data_df[vessel_cols[0]].dropna().astype(str).str.strip().tolist()
                    vals = [v for v in vals if v.lower() not in ['vessel name', 'vessel', 'nan']]
                    if vals: extracted["Vessel Name"] = vals[0]

                if voyage_cols and not extracted["Voyage Number"]:
                    vals = data_df[voyage_cols[0]].dropna().astype(str).str.strip().tolist()
                    vals = [v for v in vals if v.lower() not in ['voyage', 'voy', 'nan']]
                    if vals: extracted["Voyage Number"] = vals[0]

                if pol_cols and not extracted["Port of Loading"]:
                    vals = data_df[pol_cols[0]].dropna().astype(str).str.strip().tolist()
                    vals = [v for v in vals if v.lower() not in ['pol', 'v.pol', 'nan']]
                    if vals: extracted["Port of Loading"] = vals[0]

                if pod_cols and not extracted["Port of Discharge"]:
                    vals = data_df[pod_cols[0]].dropna().astype(str).str.strip().tolist()
                    vals = [v for v in vals if v.lower() not in ['pod', 'nan']]
                    if vals: extracted["Port of Discharge"] = vals[0]

                if carrier_cols and not extracted["Carrier / Agent"]:
                    vals = data_df[carrier_cols[0]].dropna().astype(str).str.strip().tolist()
                    vals = [v for v in vals if v.lower() not in ['carrier', 'agent', 'nan']]
                    if vals: extracted["Carrier / Agent"] = vals[0].replace('\n', ' ')

                # Sum Total Packages
                if pkg_cols and not extracted["Total Packages"]:
                    p_vals = pd.to_numeric(data_df[pkg_cols[0]], errors='coerce').dropna()
                    if not p_vals.empty:
                        extracted["Total Packages"] = str(int(p_vals.sum()))

                # Sum Total Gross Weight
                if weight_cols and not extracted["Total Weight (KG)"]:
                    w_vals = pd.to_numeric(data_df[weight_cols[0]], errors='coerce').dropna()
                    if not w_vals.empty:
                        extracted["Total Weight (KG)"] = str(round(w_vals.sum(), 2))

        # 2. Fallback Key-Value Search (If tabular detection didn't find all fields)
        keywords = {
            "Vessel Name": ["vessel name", "vessel", "ship name", "اسم السفينة"],
            "Voyage Number": ["voyage no", "voyage", "voy", "رقم الرحلة"],
            "Port of Loading": ["port of loading", "pol", "loading port"],
            "Port of Discharge": ["port of discharge", "pod", "discharge port"],
            "Carrier / Agent": ["carrier", "shipping line", "agent"],
            "Total Packages": ["total packages", "total pkgs"],
            "Total Weight (KG)": ["total weight", "gross weight"]
        }

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            for r in range(min(len(df), 30)):
                for c in range(min(len(df.columns), 15)):
                    val = str(df.iloc[r, c]).strip().lower() if pd.notna(df.iloc[r, c]) else ""
                    if not val:
                        continue
                    for field, keys in keywords.items():
                        if not extracted[field] and any(k == val for k in keys):
                            if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c + 1]):
                                extracted[field] = str(df.iloc[r, c + 1]).strip()

        # Fill GUI Fields
        for field, val in extracted.items():
            if val:
                self.fields[field].set(val)

    def generate_cuscar_txt(self):
        vessel_name = self.fields["Vessel Name"].get().strip() or "MANIFEST"
        
        # Clean vessel name for filename
        clean_filename = "".join([c for c in vessel_name if c.isalnum() or c in (' ', '_', '-')]).rstrip()
        default_filename = f"{clean_filename}.txt"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=default_filename,
            filetypes=[("Text Files", "*.txt")]
        )

        if not save_path:
            return

        # Build CUSCAR TXT Output
        content = [
            "UNB+UNOA:2+SENDER+RECEIVER'",
            "UNH+1+CUSCAR:D:95B:UN'",
            f"BGM+850+{self.fields['Voyage Number'].get()}+9'",
            f"TDT+20+{self.fields['Voyage Number'].get()}+1++++{self.fields['Vessel Name'].get()}'",
            f"LOC+7+{self.fields['Port of Loading'].get()}'",
            f"LOC+11+{self.fields['Port of Discharge'].get()}'",
            f"NAD+CA+{self.fields['Carrier / Agent'].get()}'",
            f"CNT+11:{self.fields['Total Packages'].get()}'",
            f"CNT+15:{self.fields['Total Weight (KG)'].get()}'",
            "UNT+9+1'",
            "UNZ+1+1'"
        ]

        with open(save_path, "w", encoding="utf-8") as f:
            f.write("\n".join(content))

if __name__ == "__main__":
    root = tk.Tk()
    app = CUSCARGeneratorApp(root)
    root.mainloop()
