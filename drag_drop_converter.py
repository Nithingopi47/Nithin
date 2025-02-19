import tkinter as tk
from tkinter import ttk
import os
from word_to_pdf_converter import convert_word_to_pdf
from tkinterdnd2 import DND_FILES, TkinterDnD

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word to PDF Converter")
        self.root.geometry("400x300")

        # Create and configure the main frame
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Create drop zone
        self.drop_frame = ttk.LabelFrame(self.main_frame, text="Drop Word Document Here")
        self.drop_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.drop_label = ttk.Label(self.drop_frame, text="Drag and drop .doc/.docx files here")
        self.drop_label.pack(expand=True)

        # Status label
        self.status_label = ttk.Label(self.main_frame, text="")
        self.status_label.pack(pady=10)

        # Bind drag and drop events
        self.drop_frame.bind('<Enter>', self.on_enter)
        self.drop_frame.bind('<Leave>', self.on_leave)
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)

    def on_enter(self, event):
        self.drop_frame.configure(style='Hover.TLabelframe')

    def on_leave(self, event):
        self.drop_frame.configure(style='TLabelframe')

    def on_drop(self, event):
        file_path = event.data
        if file_path.lower().endswith(('.doc', '.docx')):
            try:
                # Get output path (same directory as input, but with .pdf extension)
                output_path = os.path.splitext(file_path)[0] + '.pdf'
                convert_word_to_pdf(file_path)
                self.status_label.configure(text=f"Successfully converted to: {output_path}")
            except Exception as e:
                self.status_label.configure(text=f"Error: {str(e)}")
        else:
            self.status_label.configure(text="Please drop a Word document (.doc or .docx)")

def main():
    root = TkinterDnD.Tk()
    app = PDFConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()