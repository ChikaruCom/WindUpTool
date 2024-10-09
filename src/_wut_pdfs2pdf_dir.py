# file:.
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from win32com import client
import PyPDF2
from datetime import datetime

class FileSelectorApp(tk.Tk):
    def __init__(self, initial_dir=None):
        super().__init__()
        self.title("Multi File Selector [PDFs to PDF]")
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

        #self.convert_button = ttk.Button(self.main_frame, text="Convert to PDF", command=self.convert_to_pdf)
        #self.convert_button.pack(pady=5)

        self.merge_button = ttk.Button(self.main_frame, text="PDFs to PDF", command=self.convert_and_merge)
        self.merge_button.pack(pady=5)

        # フッターのラベルを追加してディレクトリパスを表示
        self.footer_label = ttk.Label(self, text="■Directory: None", anchor="w")
        self.footer_label.pack(fill=tk.X, padx=5, pady=5)

        self.file_vars = []
        self.current_directory = initial_dir

        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)
        self.tree.bind("<Control-Up>", self.move_up)
        self.tree.bind("<Control-Down>", self.move_down)
        self.tree.bind("<Return>", self.on_enter_key)
        self.tree.bind("<Delete>", self.on_delete_key)

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
        # フッターにディレクトリパスを表示
        short_directory = directory[-70:]
        self.footer_label.config(text=f"■Directory: {short_directory}")
        for file in files:
            if file.endswith('.pdf') or file.endswith('.pdf'):
                var = tk.BooleanVar(value=True)
                self.file_vars.append(var)
                self.tree.insert("", tk.END, values=("☑", file))

    def toggle_checkbox(self, item):
        values = self.tree.item(item, 'values')
        var_index = self.tree.index(item)
        var = self.file_vars[var_index]
        var.set(not var.get())
        self.tree.item(item, values=("☑" if var.get() else "☐", values[1]))

    def set_checkbox(self, item, state):
        values = self.tree.item(item, 'values')
        var_index = self.tree.index(item)
        var = self.file_vars[var_index]
        var.set(state)
        self.tree.item(item, values=("☑" if state else "☐", values[1]))

    def show_selected_files(self):
        selected_files = [self.tree.item(item)["values"][1] for item, var in zip(self.tree.get_children(), self.file_vars) if var.get()]
        if selected_files:
            result = "\n".join(selected_files)
            messagebox.showinfo("Selected Files", result)
        else:
            messagebox.showwarning("No Selection", "No files selected")

    def convert_to_pdf(self):
        selected_files = [self.tree.item(item)["values"][1] for item, var in zip(self.tree.get_children(), self.file_vars) if var.get()]
        print(selected_files)

        # ユーザーディレクトリを取得
        user_dir = os.path.expanduser('~')

        # WindUpTool/output のディレクトリパスを設定
        output_dir = os.path.join(user_dir, 'WindUpTool', 'output')

        # 現在の日時をフォーマット（例: 20241006_151230）
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        # 日時ディレクトリを生成
        new_dir = os.path.join(output_dir, timestamp)

        # ディレクトリを作成
        os.makedirs(new_dir, exist_ok=True)
        output_path = ""
        if selected_files:
            word = client.Dispatch("Word.Application")
            for file in selected_files:
                try:
                    full_path = os.path.join(self.current_directory, file)
                    full_path = os.path.abspath(full_path)
                    basename = os.path.basename(full_path)
                    basename_without_ext = os.path.splitext(basename)[0]
                    print(basename_without_ext)
                    output_file = os.path.join(new_dir, basename_without_ext + ".pdf")
                    #output_file = os.path.splitext(full_path)[0] + ".pdf"
                    print(output_file)
                    if os.path.exists(output_file):
                        os.remove(output_file)
                    doc = word.Documents.Open(full_path)
                    doc.SaveAs(output_file, FileFormat=17)
                    doc.Close()
                except Exception as e:
                    messagebox.showerror("Conversion Error", f"Error converting {file}: {e}")
            word.Quit()
            os.startfile(new_dir)
            messagebox.showinfo("Conversion Complete", "Selected files have been converted to PDF.")
        else:
            messagebox.showwarning("No Selection", "No files selected for conversion")

    def convert_and_merge(self):
        selected_files = [self.tree.item(item)["values"][1] for item, var in zip(self.tree.get_children(), self.file_vars) if var.get()]
        if selected_files:
            try:
                merger = PyPDF2.PdfMerger()
                for file in selected_files:
                    full_path = os.path.join(self.current_directory, file)
                    pdf_path = os.path.splitext(full_path)[0] + ".pdf"
                    #self.convert_to_pdf_file(full_path, pdf_path)
                    merger.append(pdf_path)

                output_name = os.path.splitext(selected_files[0])[0]
                output_name = output_name + "-merged"

                output_pdf = filedialog.asksaveasfilename(initialfile=output_name, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialdir=self.current_directory)
                if output_pdf:
                    with open(output_pdf, 'wb') as fout:
                        merger.write(fout)
                    messagebox.showinfo("Merge Complete", f"Selected files have been converted and merged into {output_pdf}")
            except Exception as e:
                messagebox.showerror("Merge Error", f"Error merging PDFs: {e}")
        else:
            messagebox.showwarning("No Selection", "No files selected for conversion and merge")

    def convert_to_pdf_file(self, input_doc, output_pdf):
        try:
            word = client.Dispatch("Word.Application")
            doc = word.Documents.Open(input_doc)
            doc.SaveAs(output_pdf, FileFormat=17)
            doc.Close()
            word.Quit()
        except Exception as e:
            raise RuntimeError(f"Error converting {input_doc} to PDF: {e}")

    def on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if col == "#1" and item:  # Status column
            self.toggle_checkbox(item)

    def move_up(self, event):
        item = self.tree.focus()
        if item:
            index = self.tree.index(item)
            if index > 0:
                self.tree.move(item, "", index - 1)
                self.tree.selection_set(item)
        return "break"

    def move_down(self, event):
        item = self.tree.focus()
        if item:
            index = self.tree.index(item)
            if index < len(self.tree.get_children("")) - 1:
                self.tree.move(item, "", index + 1)
                self.tree.selection_set(item)
        return "break"

    def on_enter_key(self, event):
        item = self.tree.focus()
        if item:
            self.set_checkbox(item, True)
        return "break"

    def on_delete_key(self, event):
        item = self.tree.focus()
        if item:
            self.set_checkbox(item, False)
        return "break"

if __name__ == "__main__":
    initial_dir = sys.argv[1] if len(sys.argv) > 1 else None
    app = FileSelectorApp(initial_dir)
    app.mainloop()
