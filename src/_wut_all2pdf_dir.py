import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from win32com import client
from datetime import datetime

class FileSelectorApp(tk.Tk):
    def __init__(self, initial_dir=None):
        super().__init__()
        self.title("Multi File Selector [MultiType to PDF]")
        self.geometry("600x800")

        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.browse_button = ttk.Button(self.main_frame, text="Browse Directory", command=self.browse_directory)
        self.browse_button.pack(pady=5)

        self.tree = ttk.Treeview(self.main_frame, columns=("Status", "Filename"), show='headings', selectmode="browse")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Filename", text="Filename")
        self.tree.column("Status", width=50, anchor="center")
        self.tree.column("Filename", width=550)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.convert_button = ttk.Button(self.main_frame, text="Convert to PDF", command=self.convert_to_pdf)
        self.convert_button.pack(pady=5)

        self.footer_label = ttk.Label(self, text="■Directory: None", anchor="w")
        self.footer_label.pack(fill=tk.X, padx=5, pady=5)

        self.file_vars = []
        self.current_directory = initial_dir

        if initial_dir:
            self.display_files(initial_dir)

    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.current_directory = directory
            self.display_files(directory)

    def display_files(self, directory):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.file_vars.clear()

        files = os.listdir(directory)
        files = sorted(files)
        short_directory = directory[-70:]
        self.footer_label.config(text=f"■Directory: {short_directory}")
        for file in files:
            if file.endswith('.doc') or file.endswith('.docx') or file.endswith('.xls') or file.endswith('.xlsx'):
                var = tk.BooleanVar(value=True)
                self.file_vars.append(var)
                self.tree.insert("", tk.END, values=("☑", file))

    def convert_to_pdf(self):
        selected_files = [self.tree.item(item)["values"][1] for item, var in zip(self.tree.get_children(), self.file_vars) if var.get()]
        user_dir = os.path.expanduser('~')
        output_dir = os.path.join(user_dir, 'WindUpTool', 'output')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        new_dir = os.path.join(output_dir, timestamp)
        os.makedirs(new_dir, exist_ok=True)

        if selected_files:
            word = client.Dispatch("Word.Application")
            excel = client.Dispatch("Excel.Application")
            excel.Visible = False
            for file in selected_files:
                full_path = os.path.join(self.current_directory, file)
                full_path = os.path.abspath(full_path)
                basename = os.path.basename(full_path)
                basename_without_ext = os.path.splitext(basename)[0]

                try:
                    # Wordファイルの場合
                    if file.endswith('.doc') or file.endswith('.docx'):
                        output_file = os.path.join(new_dir, basename_without_ext + ".pdf")
                        if os.path.exists(output_file):
                            os.remove(output_file)
                        doc = word.Documents.Open(full_path)
                        doc.SaveAs(output_file, FileFormat=17)
                        doc.Close()

                    # Excelファイルの場合（シートごとに出力）
                    elif file.endswith('.xls') or file.endswith('.xlsx'):
                        workbook = excel.Workbooks.Open(full_path)
                        for sheet in workbook.Sheets:
                            sheet_name = sheet.Name
                            output_file = os.path.join(new_dir, f"{basename_without_ext}_{sheet_name}.pdf")
                            if os.path.exists(output_file):
                                os.remove(output_file)
                            sheet.ExportAsFixedFormat(0, output_file)
                        workbook.Close()

                except Exception as e:
                    messagebox.showerror("Conversion Error", f"Error converting {file}: {e}")
            word.Quit()
            excel.Quit()
            os.startfile(new_dir)
            messagebox.showinfo("Conversion Complete", "Selected files have been converted to PDF.")
        else:
            messagebox.showwarning("No Selection", "No files selected for conversion")

if __name__ == "__main__":
    initial_dir = sys.argv[1] if len(sys.argv) > 1 else None
    app = FileSelectorApp(initial_dir)
    app.mainloop()
