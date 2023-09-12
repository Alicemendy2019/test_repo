import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from moduler import combine_files

class SummaryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("まとめるくん")

        # row 1
        self.label = ttk.Label(self.root, text="ファイルのパスを指定")
        self.label.grid(row=1, column=1)
        self.button = ttk.Button(self.root, text="ファイルを選択", command=self.select_directory)
        self.button.grid(row=1, column=2)
        self.dir_path = tk.StringVar()
        self.path_label = ttk.Label(self.root, textvariable=self.dir_path)
        self.path_label.grid(row=1, column=3)

        # row 2
        self.ext_label = ttk.Label(self.root, text="拡張子を選択")
        self.ext_label.grid(row=2, column=1)
        self.ext_var = tk.StringVar()
        self.ext_var.set("csv")  # default value
        self.ext_dropdown = ttk.Combobox(self.root, textvariable=self.ext_var, values=["csv", "xlsx", "xls"])
        self.ext_dropdown.grid(row=2, column=2)
        self.ext_dropdown.bind("<<ComboboxSelected>>", lambda e: self.execute_action())

        # row 3
        self.file_list = tk.Text(self.root, height=10, width=50)
        self.file_list.grid(row=3, column=1, columnspan=3)

        # row 4
        self.output_format_label = ttk.Label(self.root, text="出力形式")
        self.output_format_label.grid(row=4, column=1)
        self.output_format_var = tk.StringVar()
        self.output_format_var.set("CSV")  # default value
        self.output_format_dropdown = ttk.Combobox(self.root, textvariable=self.output_format_var, values=["CSV", "EXCEL"])
        self.output_format_dropdown.grid(row=4, column=2)

        self.execute_btn = ttk.Button(self.root, text="実行", command=self.show_confirmation)
        self.execute_btn.grid(row=4, column=3)

    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.dir_path.set(directory)
            self.selected_path = Path(directory)
            self.execute_action()

    def execute_action(self):
        if hasattr(self, 'selected_path'):
            files = list(self.selected_path.glob(f"*.{self.ext_var.get()}"))
            if self.ext_var.get() == "csv":
                files += list(self.selected_path.glob("*.CSV"))
            file_names = [f.name for f in files]
            self.file_list.delete('1.0', tk.END)
            for name in file_names:
                self.file_list.insert(tk.END, name + "\n")

    def show_confirmation(self):
        if not hasattr(self, 'selected_path'):
            messagebox.showerror("File not selected", "フォルダを選択してください")
            return 
        selected_ext_files = "\n".join([f.name for f in self.selected_path.glob(f"*.{self.ext_var.get()}")])
        if self.ext_var.get() == "csv":
            selected_ext_files += "\n".join([f.name for f in self.selected_path.glob(f"*.{self.ext_var.get()}")])
        if len(selected_ext_files) == 0:
            messagebox.showerror("File not found", "指定されたフォルダに指定形式のファイルがありません")
            return 
        output_format = self.output_format_var.get()
        res = messagebox.askokcancel("確認", f"選択されたファイル:\n{selected_ext_files}\n\n出力形式: {output_format}\n\n実行してよろしいですか？")
        if res:
            try:
                res2 = combine_files(
                    self.selected_path,
                    # selected_ext_files,
                    self.ext_var.get(),
                    output_format)
                messagebox.showinfo("完了", f"出力が完了しました\n{res2}")
            except Exception as e:
                messagebox.showerror("error", e)

if __name__ == "__main__":
    root = tk.Tk()
    app = SummaryApp(root)
    root.mainloop()
