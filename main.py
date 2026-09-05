import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime

class CUSCAREdiGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CUSCAR EDI / TXT Generator - مولد ملفات المانيفست")
        self.root.geometry("700x680")
        
        self.excel_path = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        # Header Title
        title_label = tk.Label(self.root, text="نظام تحويل بيانات الشحن إلى EDI / TXT (CUSCAR)", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)

        # Step 1: Browse Excel
        excel_frame = tk.LabelFrame(self.root, text=" 1. اختيار ملف الإكسيل (Excel Input) ", font=("Arial", 11, "bold"))
        excel_frame.pack(fill="x", padx=15, pady=5)

        tk.Entry(excel_frame, textvariable=self.excel_path, width=52).pack(side="left", padx=10, pady=10)
        tk.Button(excel_frame, text="استعراض (Browse)", command=self.browse_excel, bg="#2196F3", fg="white", font=("Arial", 9, "bold")).pack(side="left", padx=5)

        # Step 2: EDI Basic Header Details
        header_frame = tk.LabelFrame(self.root, text=" 2. البيانات الأساسية للمانيفست (Header Info) ", font=("Arial", 11, "bold"))
        header_frame.pack(fill="x", padx=15, pady=5)

        self.entries = {}
        fields = [
            ("Sender Code (مرسل):", "MAR"),
            ("Recipient Code (مستقبل):", "DW SOKHNA"),
            ("Vessel Name (اسم السفينة):", "X-PRESS KARAKORAM"),
            ("Voyage No (رقم الرحلة):", "FXKA2607W"),
            ("Call Sign (رمز النداء):", "9V2444"),
            ("Declarant Tax ID (رقم التسجيل):", "10001914565")
        ]

        for i, (label_text, default_val) in enumerate(fields):
            row = i // 2
            col = (i % 2) * 2
            tk.Label(header_frame, text=label_text, font=("Arial", 9)).grid(row=row, column=col, sticky="w", padx=5, pady=5)
            entry = tk.Entry(header_frame, width=22)
            entry.insert(0, default_val)
            entry.grid(row=row, column=col+1, padx=5, pady=5)
            self.entries[label_text] = entry

        # Actions Buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="قراءة ومعاينة الإكسيل", command=self.load_excel_preview, font=("Arial", 10, "bold"), bg="#4CAF50", fg="white", padx=12, pady=5).pack(side="left", padx=10)
        tk.Button(btn_frame, text="تصدير وحفظ كـ EDI / TXT", command=self.generate_and_save_edi, font=("Arial", 10, "bold"), bg="#FF9800", fg="white", padx=12, pady=5).pack(side="left", padx=10)

        # Step 3: Data Preview Table
        preview_frame = tk.LabelFrame(self.root, text=" معاينة البيانات المقروءة ", font=("Arial", 11, "bold"))
        preview_frame.pack(fill="both", expand=True, padx=15, pady=10)

        self.tree = ttk.Treeview(preview_frame, columns=("Container", "BL", "Weight", "Description"), show="headings")
        self.tree.heading("Container", text="رقم الحاوية")
        self.tree.heading("BL", text="رقم البوليصة")
        self.tree.heading("Weight", text="الوزن القائم")
        self.tree.heading("Description", text="وصف البضاعة")
        
        self.tree.column("Container", width=120)
        self.tree.column("BL", width=130)
        self.tree.column("Weight", width=100)
        self.tree.column("Description", width=250)
        
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.excel_path.set(filename)

    def load_excel_preview(self):
        path = self.excel_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("خطأ", "يرجى اختيار ملف Excel صحيح أولاً")
            return
        
        try:
            df = pd.read_excel(path)
            for i in self.tree.get_children():
                self.tree.delete(i)
                
            for _, row in df.iterrows():
                cntr = str(row.get("Container No", row.get("Container", row.get("CNTR", ""))))
                bl = str(row.get("BL No", row.get("Master BL", row.get("BL", ""))))
                weight = str(row.get("Weight", row.get("Gross Weight", "")))
                desc = str(row.get("Description", row.get("Cargo", "")))
                self.tree.insert("", "end", values=(cntr, bl, weight, desc))
                
            messagebox.showinfo("نجاح", f"تم قراءة الملف بنجاح! إجمالي الأسطر: {len(df)}")
        except Exception as e:
            messagebox.showerror("خطأ في قراءة الملف", str(e))

    def generate_and_save_edi(self):
        path = self.excel_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("خطأ", "قم باختيار ملف إكسيل أولاً")
            return

        save_file = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("EDI Files", "*.edi"), ("All Files", "*.*")]
        )
        if not save_file:
            return

        try:
            df = pd.read_excel(path)
            now = datetime.now()
            doc_num = "000001109"
            
            lines = []
            sender = self.entries["Sender Code (مرسل):"].get()
            recipient = self.entries["Recipient Code (مستقبل):"].get()
            
            lines.append(f"UNB+UNOA:1+{sender}+{recipient}+{now.strftime('%y%m%d:%H%M')}+{doc_num}'")
            lines.append("UNH+1+CUSCAR:D:95B:UN:1.4'")
            lines.append(f"BGM+785+EGSOKI{now.strftime('%Y%m%d%H%M%S')}+9'")
            lines.append(f"DTM+137:{now.strftime('%y%m%d0000')}:203'")
            
            tax_id = self.entries["Declarant Tax ID (رقم التسجيل):"].get()
            lines.append(f"NAD+MS+{tax_id}'")
            
            vessel = self.entries["Vessel Name (اسم السفينة):"].get()
            voyage = self.entries["Voyage No (رقم الرحلة):"].get()
            callsign = self.entries["Call Sign (رمز النداء):"].get()
            lines.append(f"TDT+20+{voyage}+1++HMM:::DPS+++{callsign}:::{vessel}:SG'")
            lines.append(f"DTM+132:{now.strftime('%y%m%d0000')}:201'")
            
            for _, row in df.iterrows():
                cntr = str(row.get("Container No", row.get("Container", row.get("CNTR", ""))))
                weight = str(row.get("Weight", row.get("Gross Weight", "25146.000")))
                volume = str(row.get("Volume", "37.1"))
                seal = str(row.get("Seal No", "IPC97722"))
                
                lines.append(f"EQD+CN+{cntr}+4500::5++3+5'")
                lines.append(f"MEA+AAE+G+KGM:{weight}'")
                lines.append(f"MEA+AAE+AAW+MTQ:{volume}'")
                lines.append(f"SEL+{seal}+SH'")
            
            lines.append(f"CNT+16:{len(df)}'")
            
            segment_count = len(lines) - 1 + 2
            lines.append(f"UNT+{segment_count}+1'")
            lines.append(f"UNZ+1+{doc_num}'")
            
            with open(save_file, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
                
            messagebox.showinfo("تم الحفظ", f"تم إنشاء وتصدير الملف بنجاح إلى:\n{save_file}")
            
        except Exception as e:
            messagebox.showerror("خطأ أثناء التصدير", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = CUSCAREdiGeneratorApp(root)
    root.mainloop()
