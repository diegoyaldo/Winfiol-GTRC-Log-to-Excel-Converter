import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

UPLOAD_FOLDER = 'uploads'
SPC_FOLDER = 'spc_files'
ALLOWED_EXTENSIONS = {'txt', 'log'}
SPC_FILENAME = 'SPC.txt'

# Ensure the necessary directories exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(SPC_FOLDER):
    os.makedirs(SPC_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def clean_file(filename):
    with open(filename, "r") as file:
        lines = file.readlines()

    cleaned_lines = []
    extrastuff_done = False

    for line in lines:
        if "PTERM" in line:
            extrastuff_done = True

        if "ACTIVE:" in line and extrastuff_done:
            cleaned_lines.append(line[8:] + "\n")

        if "END" in line:
            extrastuff_done = False

    base, ext = os.path.splitext(filename)
    new_filename = f"{base}_cleaned{ext}"

    with open(new_filename, "w") as file:
        file.writelines(cleaned_lines)

    return new_filename

def clean_psp_ssp(value):
    if pd.isna(value):
        return value
    return value.split('-')[-1]

def parse_spc(spc_filename):
    spc_mapping = {}
    with open(spc_filename, "r", encoding="utf-8", errors="replace") as file:
        for line in file:
            parts = line.split(maxsplit=1)
            spc_mapping[parts[0]] = parts[1].strip()

    return spc_mapping

def parse(filename, spc_mapping):
    columns = [
        "GTRC",
        "PSP",
        "PTERM",
        "PINTER",
        "PSSN",
        "SSP",
        "STERM",
        "SINTER",
        "SSSN",
        "LSH",
    ]

    df = pd.read_fwf(filename, sep=r"^\s*$", names=columns)

    df = df.replace(r"^\s*$", pd.NA, regex=True)

    df['PSP_Cleaned'] = df['PSP'].apply(clean_psp_ssp)
    df['SSP_Cleaned'] = df['SSP'].apply(clean_psp_ssp)

    df["PSP_Name"] = df["PSP_Cleaned"].map(spc_mapping)
    df["SSP_Name"] = df["SSP_Cleaned"].map(spc_mapping)

    # delete it
    del df['PSP_Cleaned'] 
    del df['SSP_Cleaned']

    output_file = "output.xlsx"
    df.to_excel(output_file, index=False)

    return output_file

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text and Log files", "*.txt *.log")])
    if not file_path:
        messagebox.showwarning("No file selected", "Please select a log file.")
        return
    
    filename = os.path.basename(file_path)

    if allowed_file(filename):
        upload_path = os.path.join(UPLOAD_FOLDER, filename)
        spc_path = os.path.join(SPC_FOLDER, SPC_FILENAME)
        
        if not os.path.exists(spc_path):
            messagebox.showwarning("SPC file not found", f"Default SPC file {SPC_FILENAME} not found in {SPC_FOLDER}.")
            return
        
        shutil.copy(file_path, upload_path)
        
        # Process the file and generate the excel file
        cleaned_file = clean_file(upload_path)
        spc_mapping = parse_spc(spc_path)
        output_file = parse(cleaned_file, spc_mapping)
        
        # Ask the user where to save the output file
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            shutil.move(output_file, save_path)
            messagebox.showinfo("Success", f"File processed and saved as {save_path}")
        
        # Clear the uploads cache
        shutil.rmtree(UPLOAD_FOLDER)
        os.makedirs(UPLOAD_FOLDER)
    else:
        messagebox.showwarning("Invalid file", "Only .txt and .log files are allowed.")

# Setup Tkinter window
root = tk.Tk()
root.title("Winfiol GTRC Routing Log to Excel Converter")
root.geometry("700x400")

# Set the icon (if you have an icon file, place it in the same directory)
icon_path = os.path.join("assets", "world.ico")
root.iconbitmap(icon_path)

title_label = tk.Label(root, text="Winfiol GTRC Routing Log to Excel Converter", font=("Segoe UI", 24))
title_label.pack(pady=10)

subtitle_label = tk.Label(root, text="Select a log file to process", font=("Segoe UI", 16))
subtitle_label.pack(pady=5)

upload_btn = tk.Button(root, text="Upload File", command=upload_file, bg="gray", fg="white", font=("Segoe UI", 14))
upload_btn.pack(pady=20)

footer_label = tk.Label(root, text="This tool processes log files and maps them to SPC data using a default SPC.txt file.\nPlease ensure the SPC.txt file is located in the spc_files folder.", font=("Segoe UI", 10), fg="#6c757d", justify=tk.CENTER)
footer_label.pack(pady=10)

root.mainloop()
