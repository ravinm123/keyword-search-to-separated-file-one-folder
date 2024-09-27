import os
import shutil
import fitz
import pandas as pd
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# Global dictionary to hold keywords and folder mappings
keywords_folders = {}

def ensure_target_folders(source_dir):
    for folder in set(keywords_folders.values()):
        target_folder = os.path.join(source_dir, folder)
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

def read_pdf(file_path):
    content = ""
    with fitz.open(file_path) as doc:
        for page in doc:
            content += page.get_text()
    return content.lower()

def read_excel(file_path):
    content = ""
    xls = pd.ExcelFile(file_path)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name)
        content += df.to_string()
    return content.lower()

def read_word(file_path):
    content = ""
    doc = Document(file_path)
    for para in doc.paragraphs:
        content += para.text
    return content.lower()

def process_file(file_path, source_dir):
    file_name = os.path.basename(file_path)
    extension = file_name.lower().split('.')[-1]

    if extension == 'pdf':
        content = read_pdf(file_path)
    elif extension in ['xls', 'xlsx']:
        content = read_excel(file_path)
    elif extension == 'docx':
        content = read_word(file_path)
    else:
        return f"Unsupported file format: {file_name}"

    for keywords, folder in keywords_folders.items():
        if any(keyword in content for keyword in keywords):
            target_folder_path = os.path.join(source_dir, folder)
            shutil.move(file_path, os.path.join(target_folder_path, file_name))
            return f"Moved file {file_name} to folder '{folder}' based on keyword(s) {keywords}"
    
    return f"No keyword found in file: {file_name}"

def organize_files(source_dir):
    ensure_target_folders(source_dir)
    messages = []

    for file_name in os.listdir(source_dir):
        file_path = os.path.join(source_dir, file_name)

        if os.path.isfile(file_path):
            result = process_file(file_path, source_dir)
            messages.append(result)

    return "\n".join(messages)

def select_directory():
    directory = filedialog.askdirectory()
    if directory:
        result = organize_files(directory)
        messagebox.showinfo("Organizing Complete", result)

def add_keyword_folder():
    keyword = simpledialog.askstring("Input", "Enter keyword(s) (comma separated):")
    folder = simpledialog.askstring("Input", "Enter folder name:")
    
    if keyword and folder:
        keywords_list = tuple(keyword.lower().strip() for keyword in keyword.split(','))
        keywords_folders[keywords_list] = folder
        messagebox.showinfo("Success", f"Added keywords: {keywords_list} to folder: {folder}")
    else:
        messagebox.showwarning("Input Error", "Both fields are required!")

# Setting up the Tkinter GUI
root = tk.Tk()
root.title("File Organizer")
root.geometry("400x200")



btn_add_keyword = tk.Button(root, text="Add Keyword and Folder", command=add_keyword_folder)
btn_add_keyword.pack(pady=10)

btn_select_dir = tk.Button(root, text="Select Directory", command=select_directory)
btn_select_dir.pack(pady=10)

root.mainloop()
