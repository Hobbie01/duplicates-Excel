import pandas as pd
import os
from tkinter import Tk, filedialog, Button, Label, Listbox, Scrollbar, messagebox, END, MULTIPLE

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Merge Multiple Excels & Remove Duplicates")

        self.file_paths = []   # เก็บ path ไฟล์ที่เลือก

        # UI components
        self.label = Label(root, text="เลือกไฟล์ Excel เพื่อนำมารวม", pady=10)
        self.label.pack()

        self.select_btn = Button(root, text="เพิ่มไฟล์ Excel", command=self.open_files_dialog, width=25)
        self.select_btn.pack(pady=5)

        # listbox แสดงไฟล์ + โหมด MULTIPLE
        self.listbox = Listbox(root, selectmode=MULTIPLE, width=80)
        self.listbox.pack(pady=5)

        self.scrollbar = Scrollbar(root, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")

        self.remove_btn = Button(root, text="ลบไฟล์ที่เลือก", command=self.remove_selected_files, width=25)
        self.remove_btn.pack(pady=5)

        self.generate_btn = Button(root, text="Generate และ Export", command=self.generate_and_export, width=25)
        self.generate_btn.pack(pady=5)

    def open_files_dialog(self):
        new_files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
        if new_files:
            self.file_paths.extend(new_files)
            self.refresh_file_list()

    def refresh_file_list(self):
        self.listbox.delete(0, END)
        for file_path in self.file_paths:
            self.listbox.insert(END, os.path.basename(file_path))

    def remove_selected_files(self):
        selected_indices = list(self.listbox.curselection())
        if selected_indices:
            for index in sorted(selected_indices, reverse=True):
                del self.file_paths[index]
            self.refresh_file_list()

    def process_multiple_excels(self):
        all_data = []
        for file_path in self.file_paths:
            df = pd.read_excel(file_path)
            all_data.append(df)

        merged = pd.concat(all_data, ignore_index=True)
        merged_unique = merged.drop_duplicates()
        return merged_unique

    def generate_and_export(self):
        if not self.file_paths:
            messagebox.showwarning("ยังไม่ได้เลือกไฟล์", "กรุณาเลือกไฟล์ Excel ก่อน")
            return

        try:
            merged_df = self.process_multiple_excels()
            export_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save merged Excel as...",
                initialfile="merged_result.xlsx"  # <<<<<< ตั้งชื่อไฟล์ default ตรงนี้!
            )
            if export_path:
                merged_df.to_excel(export_path, index=False)
                messagebox.showinfo("สำเร็จ", f"ไฟล์ถูกบันทึกที่:\n{export_path}")
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เกิดข้อผิดพลาด: {e}")

# เริ่มโปรแกรม
if __name__ == "__main__":
    root = Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
