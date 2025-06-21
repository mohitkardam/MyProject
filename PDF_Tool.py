#!/usr/bin/env python
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF3 import PdfFileReader, PdfFileWriter, PdfFileMerger
from pdf2docx import Converter
from pdf2image import convert_from_path, convert_from_bytes
from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)
from PIL import Image
import tabula
import os
import sys
import pandas as pd
import openpyxl
import pypandoc
import subprocess



#POPPLER_PATH = resource_path("poppler/Library/bin")
"""
Use command--->pyinstaller --onefile --add-data "logo.png;." Remote diagnostic.pyw
pyinstaller --onefile --add-data "poppler\\Library\\bin;poppler/Library/bin" test_pdf_jpg.py
pyinstaller --onefile^ --add-data "poppler\\Library\\bin;poppler/Library/bin"^ --add-data "logo.png;."^ test_pdf_jpg.py
pyinstaller --onefile^ --add-data "poppler\\Library\\bin;poppler/Library/bin"^ --add-data "logo.png;."^ --add-data "tabula.jar;."^ PDF_Tool_Updated.py
"""


# ---------- UTILITY FUNCTIONS ----------
def browse_file(entry):
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF Files", "*.pdf")],
        title="Select a PDF file"
    )
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def browse_saveas_jpg():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".jpg",
        filetypes=[("JPEG files", "*.jpg")],
        title="Save as JPG"
    )
    # Ensure it ends with .jpg even if user removes it
    if file_path and not file_path.lower().endswith(".jpg"):
        file_path += ".jpg"
    return file_path
   


def browse_saveas():
    return filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]) 

def browse_saveas_excel():
    return filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("XLSX files", "*.xlsx")]) 


def browse_saveas_docx():
    return filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("DOC files", "*.docx")])


def browse_directory():
    return filedialog.askdirectory()


# ---------- TOOL FUNCTIONS ----------
def merge_pdfs():
    pdf1_path = entry_pdf1.get()
    pdf2_path = entry_pdf2.get()
    if not pdf1_path or not pdf2_path:
        messagebox.showerror("Error", "Please select both PDF files!")
        return
    try:
        merger = PdfFileMerger()
        merger.append(pdf1_path)
        merger.append(pdf2_path)
        output_path = browse_saveas()
        if output_path:
            merger.write(output_path)
            merger.close()
            messagebox.showinfo("Success", "PDFs merged successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def rotate_pdf():
    pdf_path = entry_rotate.get()
    angle = angle_var.get()
    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file!")
        return
    try:
        reader = PdfFileReader(pdf_path)
        writer = PdfFileWriter()
        for page_num in range(reader.getNumPages()):
            page = reader.getPage(page_num)
            page.rotateClockwise(angle)
            writer.addPage(page)
        output = browse_saveas()
        if output:
            with open(output, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", "PDF rotated successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def pdf_to_doc():
    pdf_path = entry_pdf_doc.get()
    output = browse_saveas_docx().replace(".pdf", ".docx")
    try:
        cv = Converter(pdf_path)
        cv.convert(output, start=0, end=None)
        cv.close()
        messagebox.showinfo("Success", "Converted to DOCX!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


# Function to resolve the resource path (for PyInstaller EXE)
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # for PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)




def pdf_to_excel():
    pdf_path = entry_pdf_excel.get()
    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file.")
        return

    try:
        output = browse_saveas_excel().replace(".pdf", ".xlsx")

        # Get path to tabula JAR
        tabula_jar = resource_path("tabula.jar")

        # Set JAVA options to use the custom tabula.jar
        tabula.environment_info()
        tabula.java_options = ["-Dfile.encoding=UTF8"]
        os.environ["TABULA_JAR"] = tabula_jar  # this is the key step

        # Read PDF tables
        dfs = tabula.read_pdf(
            pdf_path,
            pages="all",
            multiple_tables=True
        )

        if not dfs or all(df.empty for df in dfs):
            messagebox.showerror("Error", "No tables found in the PDF.")
            return

        with pd.ExcelWriter(output) as writer:
            for i, df in enumerate(dfs):
                if not df.empty:
                    df.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)

        messagebox.showinfo("Success", "Converted to Excel!")
    except Exception as e:
        messagebox.showerror("Error", f"Conversion failed:\n{e}")






def split_pdf():
    pdf_path = entry_split.get()
    page_range = entry_split_range.get()
    try:
        reader = PdfFileReader(pdf_path)
        writer = PdfFileWriter()
        if '-' in page_range:
            start, end = map(int, page_range.split('-'))
            pages = range(start-1, end)
        else:
            pages = [int(page_range)-1]
        for p in pages:
            writer.addPage(reader.getPage(p))
        output = browse_saveas()
        if output:
            with open(output, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", "PDF split successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def compress_pdf():
    pdf_path = entry_compress.get()
    try:
        reader = PdfFileReader(pdf_path)
        writer = PdfFileWriter()
        for i in range(reader.numPages):
            page = reader.getPage(i)
            page.compressContentStreams()
            writer.addPage(page)
        output = browse_saveas()
        if output:
            with open(output, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", "Compressed successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

'''
def pdf_to_jpg():
    pdf_path = entry_pdf_jpg.get()
    entry_pdf_jpg.delete(0, tk.END)
    entry_pdf_jpg.insert(0, pdf_path)

    output = browse_saveas_jpg()
    if not output:
        return

    try:
        output_folder = os.path.dirname(pdf_path)
        base_filename = os.path.splitext(os.path.basename(pdf_path))[0]

        images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)

        for i, image in enumerate(images):
            image_path = os.path.join(output_folder, f"{base_filename}_page_{i+1}.jpg")
            image.save(image_path, "JPEG")

        messagebox.showinfo("Success", f"Converted {len(images)} page(s) to JPG successfully in the same folder!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert PDF to JPG:\n{e}")
'''




def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS  # used when bundled with PyInstaller
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Set the Poppler path (for convert_from_path)
#POPPLER_PATH = resource_path("poppler/Library/bin")
POPPLER_PATH = resource_path("poppler/Library/bin")

def pdf_to_jpg():
    pdf_path = entry_pdf_jpg.get()
    entry_pdf_jpg.delete(0, tk.END)
    entry_pdf_jpg.insert(0, pdf_path)

    output = browse_saveas_jpg()
    if not output:
        return

    try:
        output_folder = os.path.dirname(output)
        base_filename = os.path.splitext(os.path.basename(output))[0]

        images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
        for i, image in enumerate(images):
            image_path = os.path.join(output_folder, f"{base_filename}_page_{i + 1}.jpg")
            image.save(image_path, "JPEG")

        messagebox.showinfo("Success", f"Converted {len(images)} page(s) to JPG successfully in the same folder!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert PDF to JPG:\n{e}")


def jpg_to_pdf():
    files = filedialog.askopenfilenames(filetypes=[("JPG Files", "*.jpg")])
    if files:
        try:
            output = browse_saveas()
            images = [Image.open(f).convert('RGB') for f in files]
            images[0].save(output, save_all=True, append_images=images[1:])
            messagebox.showinfo("Success", "JPGs converted to PDF!")
        except Exception as e:
            messagebox.showerror("Error", str(e))


def remove_pages():
    pdf_path = entry_remove.get()
    pages_to_remove = entry_remove_pages.get()
    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file!")
        return
    if not pages_to_remove:
        messagebox.showerror("Error", "Please specify pages to remove!")
        return
    try:
        reader = PdfFileReader(pdf_path)
        writer = PdfFileWriter()
        total_pages = reader.getNumPages()

        # Parse the pages to remove
        remove_set = set()
        parts = pages_to_remove.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                remove_set.update(range(start, end + 1))
            else:
                remove_set.add(int(part))

        # Rebuild PDF without specified pages
        for page_num in range(total_pages):
            if (page_num + 1) not in remove_set:
                writer.addPage(reader.getPage(page_num))

        output = browse_saveas()
        if output:
            with open(output, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", "Selected pages removed successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))







def convert_pdf_to_odt():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if not pdf_file:
        return

    output_dir = os.path.dirname(pdf_file)
    base_name = os.path.splitext(os.path.basename(pdf_file))[0]
    intermediate_docx = os.path.join(output_dir, base_name + ".docx")
    final_odt = os.path.join(output_dir, base_name + ".odt")

    try:
        # Step 1: Convert PDF to DOCX using pdf2docx
        cv = Converter(pdf_file)
        cv.convert(intermediate_docx, start=0, end=None)
        cv.close()

        # Step 2: Convert DOCX to ODT using LibreOffice
        soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
        result = subprocess.run(
            [soffice_path, "--headless", "--convert-to", "odt", intermediate_docx, "--outdir", output_dir],
            capture_output=True,
            text=True
        )

        if os.path.exists(final_odt):
            messagebox.showinfo("Success", f"Converted to ODT: {final_odt}")
            os.remove(intermediate_docx)
        else:
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            messagebox.showerror("Error", "ODT conversion failed.")
    except Exception as e:
        messagebox.showerror("Error", str(e))









# ---------- GUI SETUP ----------
master = tk.Tk()
master.title("PDF Tool")
master.geometry("500x300+650+400") #XxY+from X padding+from Y padding
tool_var = tk.StringVar(value="Select Tool")
frame_dict = {}
all_frames = []
#------------ICON PATH(Logo.png)-------------
try:
    icon_path = resource_path("logo.png")  
    icon_image = tk.PhotoImage(file=icon_path)
    master.iconphoto(True, icon_image)
except Exception as e:
    print(f"Error loading icon: {e}")
    
#-------CLEAR FRAME-----------------
def clear_frames():
    for frame in all_frames:
        frame.grid_remove()

#------TOOL INTERFACE---------------
def show_tool_interface(tool_name):
    clear_frames()
    frame = frame_dict.get(tool_name)
    if frame:
        frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10)


top_frame = tk.Frame(master)
top_frame.grid(row=0, column=0, pady=40,padx=90)
top_frame.grid_columnconfigure(0, weight=1)

#-------------TOOL MENU----------------------------
tk.Label(top_frame, text="Select Tool:", font=("Arial", 12)).grid(row=0, column=0, padx=5, pady=5)
tool_menu = tk.OptionMenu(top_frame, tool_var,
                          "Merge PDFs", "Rotate PDF", "PDF to DOC",
                          "PDF to Excel", "Split PDF", "Compress PDF",
                          "PDF to JPG", "JPG to PDF","Remove Pages","PDF to ODT",
                          command=show_tool_interface)
tool_menu.grid(row=0, column=1, padx=5, pady=5)

# ---------- Tool Frames ----------
# Merge
merge_frame = tk.Frame(master)
entry_pdf1 = tk.Entry(merge_frame, width=40)
entry_pdf2 = tk.Entry(merge_frame, width=40)
tk.Label(merge_frame, text="PDF 1:").grid(row=0, column=0)
entry_pdf1.grid(row=0, column=1)
tk.Button(merge_frame, text="Browse", command=lambda: browse_file(entry_pdf1), bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=2)
tk.Label(merge_frame, text="PDF 2:").grid(row=1, column=0)
entry_pdf2.grid(row=1, column=1)
tk.Button(merge_frame, text="Browse", command=lambda: browse_file(entry_pdf2), bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=1, column=2)
tk.Button(merge_frame, text="Merge PDFs", command=merge_pdfs, bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=2, column=1)
frame_dict["Merge PDFs"] = merge_frame
all_frames.append(merge_frame)

# Rotate
rotate_frame = tk.Frame(master)
entry_rotate = tk.Entry(rotate_frame, width=40)
angle_var = tk.IntVar(value=90)
tk.Label(rotate_frame, text="PDF File:").grid(row=0, column=0)
entry_rotate.grid(row=0, column=1)
tk.Button(rotate_frame, text="Browse", command=lambda: browse_file(entry_rotate), bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=2)
tk.Label(rotate_frame, text="Angle:").grid(row=1, column=0)
tk.OptionMenu(rotate_frame, angle_var, 90, 180, 270).grid(row=1, column=1)
tk.Button(rotate_frame, text="Rotate",command=rotate_pdf, bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=2, column=1)
frame_dict["Rotate PDF"] = rotate_frame
all_frames.append(rotate_frame)

# PDF to DOC
convert_frame = tk.Frame(master)
entry_pdf_doc = tk.Entry(convert_frame, width=40)
tk.Label(convert_frame, text="PDF File:").grid(row=0, column=0)
entry_pdf_doc.grid(row=0, column=1)
tk.Button(convert_frame, text="Browse", command=lambda: browse_file(entry_pdf_doc),bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=2)
tk.Button(convert_frame, text="Convert to DOCX", command=pdf_to_doc,bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=1, column=1)
frame_dict["PDF to DOC"] = convert_frame
all_frames.append(convert_frame)

# PDF to Excel
excel_frame = tk.Frame(master)
entry_pdf_excel = tk.Entry(excel_frame, width=40)
tk.Label(excel_frame, text="PDF File:").grid(row=0, column=0)
entry_pdf_excel.grid(row=0, column=1)
tk.Button(excel_frame, text="Browse", command=lambda: browse_file(entry_pdf_excel),bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=2)
tk.Button(excel_frame, text="Convert to Excel", command=pdf_to_excel,bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=1, column=1)
frame_dict["PDF to Excel"] = excel_frame
all_frames.append(excel_frame)

# Split PDF
split_frame = tk.Frame(master)
entry_split = tk.Entry(split_frame, width=40)
entry_split_range = tk.Entry(split_frame, width=20)
tk.Label(split_frame, text="PDF File:").grid(row=0, column=0)
entry_split.grid(row=0, column=1)
tk.Button(split_frame, text="Browse", command=lambda: browse_file(entry_split),bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=2)
tk.Label(split_frame, text="Pages (e.g., 1-3 or 5):").grid(row=1, column=0)
entry_split_range.grid(row=1, column=1)
tk.Button(split_frame, text="Split PDF", command=split_pdf,bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=2, column=1)
frame_dict["Split PDF"] = split_frame
all_frames.append(split_frame)

# Compress PDF
compress_frame = tk.Frame(master)
entry_compress = tk.Entry(compress_frame, width=40)
tk.Label(compress_frame, text="PDF File:").grid(row=0, column=0)
entry_compress.grid(row=0, column=1)
tk.Button(compress_frame, text="Browse", command=lambda: browse_file(entry_compress),bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=2)
tk.Button(compress_frame, text="Compress", command=compress_pdf,bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=1, column=1)
frame_dict["Compress PDF"] = compress_frame
all_frames.append(compress_frame)

# PDF to JPG
jpg_frame = tk.Frame(master)
entry_pdf_jpg = tk.Entry(jpg_frame, width=40)
tk.Label(jpg_frame, text="PDF File:").grid(row=0, column=0)
entry_pdf_jpg.grid(row=0, column=1)
tk.Button(jpg_frame, text="Browse", command=lambda: browse_file(entry_pdf_jpg),bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=2)

tk.Button(jpg_frame, text="Convert to JPG", command=pdf_to_jpg,bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=1, column=1)

frame_dict["PDF to JPG"] = jpg_frame
all_frames.append(jpg_frame)

# JPG to PDF
jpg_to_pdf_frame = tk.Frame(master)
tk.Button(jpg_to_pdf_frame, text="Select JPG Files & Convert", command=jpg_to_pdf,bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=1)
frame_dict["JPG to PDF"] = jpg_to_pdf_frame
all_frames.append(jpg_to_pdf_frame)


# PDF to ODT Frame
odt_frame = tk.Frame(master)
tk.Label(odt_frame, text="Select PDF to Convert to ODT:").grid(row=0, column=0, padx=5, pady=5)
tk.Button(odt_frame, text="Convert to ODT", command=convert_pdf_to_odt, bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=1)
frame_dict["PDF to ODT"] = odt_frame
all_frames.append(odt_frame)


# Remove Pages
remove_frame = tk.Frame(master)
entry_remove = tk.Entry(remove_frame, width=40)
entry_remove_pages = tk.Entry(remove_frame, width=20)
tk.Label(remove_frame, text="PDF File:").grid(row=0, column=0)
entry_remove.grid(row=0, column=1)
tk.Button(remove_frame, text="Browse", command=lambda: browse_file(entry_remove), bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=0, column=2)
tk.Label(remove_frame, text="Pages to remove (e.g., 2,4-6):").grid(row=1, column=0)
entry_remove_pages.grid(row=1, column=1)
tk.Button(remove_frame, text="Remove Pages", command=remove_pages, bg="#A9A9A9", fg="black", font=("Arial", 10, "bold")).grid(row=2, column=1)
frame_dict["Remove Pages"] = remove_frame
all_frames.append(remove_frame)






# Start GUI
master.mainloop()
