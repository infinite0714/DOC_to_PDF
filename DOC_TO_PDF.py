import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert

def convert_docx_to_pdf(input_file_paths):
    for input_file_path in input_file_paths:
        try:
            output_file_path = input_file_path.replace('.docx', '.pdf')
            convert(input_file_path, output_file_path)
            print(f"Conversion of '{input_file_path}' complete!")
        except Exception as e:
            print(f"Error converting '{input_file_path}': {str(e)}")

def select_files_and_convert():
    input_file_paths = filedialog.askopenfilenames(filetypes=[("Word Files", "*.docx")])
    if input_file_paths:
        convert_docx_to_pdf(input_file_paths)
        status_label.config(text="All conversions completed!", fg="green")
    else:
        status_label.config(text="No files selected!", fg="red")

# GUI setup
root = tk.Tk()
root.title("Batch Word to PDF Converter")

convert_button = tk.Button(root, text="Convert Files", command=select_files_and_convert)
convert_button.pack(pady=20)

status_label = tk.Label(root, text="", fg="black")
status_label.pack()

root.mainloop()