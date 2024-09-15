import tkinter as tk
import tkinter.filedialog
from spire.pdf.common import *
from spire.pdf import *


class MainGUI:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.geometry("400x400")
        self.root.title("File Converter")

        self.label = tk.Label(self.root, text="File Converter", font=("Arial", 18))
        self.label.pack()

        self.button = tk.Button(
            self.root,
            text="Select a file",
            font=("Arial", 16),
            command=self.get_filepath,
        )
        self.button.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        self.root.mainloop()

    def get_filepath(self):
        filename = tkinter.filedialog.askopenfilename()

        doc_filename = filename.replace(".pdf", ".docx")

        doc = PdfDocument()
        doc.LoadFromFile(filename)
        doc.SaveToFile(doc_filename, FileFormat.DOCX)
        doc.Close()
        tk.messagebox.showinfo(
            title="Conversion Done",
            message=f"File has been converted to {doc_filename}",
        )


MainGUI()
