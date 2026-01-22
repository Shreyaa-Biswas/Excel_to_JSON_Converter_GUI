import sys
sys.dont_write_bytecode = True

import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

class ExcelToJsonConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to JSON Converter")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        self.excel_path = None
        self.output_path = None
        
        self.setup_ui()
    
    def setup_ui(self):
        title_label = tk.Label(
            self.root, 
            text="Excel to JSON Converter", 
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=20)
        
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Label(file_frame, text="Excel File:", font=("Arial", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.excel_entry = tk.Entry(file_frame, width=50)
        self.excel_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Browse", command=self.browse_excel, width=10).grid(row=0, column=2, pady=5)
        
        tk.Label(file_frame, text="Output JSON:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", pady=5)
        self.output_entry = tk.Entry(file_frame, width=50)
        self.output_entry.grid(row=1, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Browse", command=self.browse_output, width=10).grid(row=1, column=2, pady=5)
        
        options_frame = tk.Frame(self.root)
        options_frame.pack(pady=10)
        
        self.include_sheet_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame, 
            text="Include sheet name in records", 
            variable=self.include_sheet_var,
            font=("Arial", 10)
        ).pack()
        
        self.convert_btn = tk.Button(
            self.root, 
            text="Convert to JSON", 
            command=self.start_conversion,
            font=("Arial", 12, "bold"),
            bg="#4CAF50",
            fg="white",
            width=20,
            height=2
        )
        self.convert_btn.pack(pady=20)
        
        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=400)
        self.progress.pack(pady=10)
        
        self.status_label = tk.Label(
            self.root, 
            text="Ready", 
            font=("Arial", 10),
            fg="gray"
        )
        self.status_label.pack(pady=10)
    
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path = filename
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, filename)
            
            if not self.output_path:
                default_output = filename.rsplit('.', 1)[0] + '.json'
                self.output_entry.delete(0, tk.END)
                self.output_entry.insert(0, default_output)
                self.output_path = default_output
    
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save JSON As",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.output_path = filename
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)
    
    def start_conversion(self):
        if not self.excel_path:
            messagebox.showerror("Error", "Please select an Excel file")
            return
        
        if not self.output_path:
            messagebox.showerror("Error", "Please select output JSON file")
            return
        
        self.convert_btn.config(state="disabled")
        self.progress.start()
        self.status_label.config(text="Converting...", fg="blue")
        
        thread = threading.Thread(target=self.convert)
        thread.start()
    
    def convert(self):
        try:
            xls = pd.ExcelFile(self.excel_path, engine="openpyxl")
            all_records = []
            
            for sheet in xls.sheet_names:
                self.update_status(f"Processing sheet: {sheet}")
                df = xls.parse(sheet)
                
                for col in df.select_dtypes(include=['datetime64']):
                    df[col] = df[col].astype(str)
                
                records = df.to_dict('records')
                
                for record in records:
                    clean_record = {}
                    for k, v in record.items():
                        if pd.isna(v):
                            clean_record[k] = None
                        elif hasattr(v, 'isoformat'):
                            clean_record[k] = v.isoformat()
                        elif isinstance(v, (pd.Timestamp, pd.Timedelta)):
                            clean_record[k] = str(v)
                        else:
                            clean_record[k] = v
                    
                    if self.include_sheet_var.get():
                        clean_record["_sheet"] = sheet
                    
                    all_records.append(clean_record)
            
            self.update_status(f"Writing {len(all_records)} records to JSON...")
            
            with open(self.output_path, "w", encoding="utf-8") as f:
                f.write("[")
                
                for i, record in enumerate(all_records):
                    if i == 0:
                        json_str = json.dumps(record, ensure_ascii=False)
                        f.write(json_str)
                    else:
                        f.write(",\n\n ")
                        json_str = json.dumps(record, ensure_ascii=False, indent=1)
                        lines = json_str.split('\n')
                        formatted_lines = [lines[0]] + [' ' + line for line in lines[1:]]
                        f.write('\n '.join(formatted_lines))
                
                f.write("]")
            
            self.conversion_complete(len(all_records))
            
        except Exception as e:
            self.conversion_error(str(e))
    
    def update_status(self, message):
        self.root.after(0, lambda: self.status_label.config(text=message))
    
    def conversion_complete(self, record_count):
        self.root.after(0, lambda: self.progress.stop())
        self.root.after(0, lambda: self.convert_btn.config(state="normal"))
        self.root.after(0, lambda: self.status_label.config(
            text=f"Success! {record_count} records converted", 
            fg="green"
        ))
        self.root.after(0, lambda: messagebox.showinfo(
            "Success", 
            f"Conversion completed!\n{record_count} records written to:\n{self.output_path}"
        ))
    
    def conversion_error(self, error_msg):
        self.root.after(0, lambda: self.progress.stop())
        self.root.after(0, lambda: self.convert_btn.config(state="normal"))
        self.root.after(0, lambda: self.status_label.config(text="Error occurred", fg="red"))
        self.root.after(0, lambda: messagebox.showerror("Error", f"Conversion failed:\n{error_msg}"))

def main():
    root = tk.Tk()
    app = ExcelToJsonConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()