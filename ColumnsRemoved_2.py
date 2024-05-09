import os
import csv
import tkinter as tk
from tkinter import filedialog
from docx import Document
import re
import win32com.client

def select_directory():
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Extraction Directory")
    return folder_path

def extract_text_from_doc(doc_path):
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    abs_path = os.path.abspath(doc_path)
    doc = word_app.Documents.Open(abs_path)
    text = doc.Content.Text
    doc.Close(False)
    if word_app.Documents.Count == 0:
        word_app.Quit()
    return text

def extract_patient_info(file_path):
    _, file_extension = os.path.splitext(file_path)
    if file_extension.lower() == '.docx':
        doc = Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + " "
    elif file_extension.lower() == '.doc':
        text = extract_text_from_doc(file_path)
    else:
        return None
    match = re.search(r"PATIENT\s*NAME:\s*([^\n]*)\s*DOB", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    else:
        re_match = re.search(r"RE:\s*([^\n]*)\s*DOB:", text, re.IGNORECASE)
        if re_match:
            return re_match.group(1).strip()
    return None

def split_files_by_delimiter(folder_path):
    if not os.path.isdir(folder_path):
        print("Invalid folder path.")
        return
    files = os.listdir(folder_path)
    total_files = len(files)
    processed_files = 0
    file_info_list = []

    for idx, filename in enumerate(files, start=1):
        file_name, file_extension = os.path.splitext(filename)
        file_path = os.path.join(folder_path, filename)

        if file_extension.lower() not in ['.doc', '.docx']:
            print(f"Delimiter Pattern Not Found, Skipping: {filename}")
            continue

        patient_info = extract_patient_info(file_path)
        print(f"Processing file {idx} of {total_files}: {filename}, Patient Info: {patient_info}")
        file_info_list.append([file_name] + file_name.split('_') + [patient_info])
        processed_files += 1

    folder_name = os.path.basename(folder_path)
    csv_file_name = f"{folder_name}_output.csv"
    csv_file_path = os.path.normpath(os.path.join(folder_path, csv_file_name))
    with open(csv_file_path, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['File Name', 'Audio #', 'Dictation ID', 'Date', 'Visit Type', 'ID', 'Visit Type', 'Patient Name'])
        writer.writerows(file_info_list)

    print(f"File information saved to {csv_file_path}")
    print(f"Number of files processed: {processed_files}")

def remove_columns_from_csv(csv_file_path, columns_to_remove):
    if os.path.exists(csv_file_path):
        if os.path.getsize(csv_file_path) > 0:
            with open(csv_file_path, 'r') as csvfile:
                reader = csv.reader(csvfile)
                header = next(reader, None)
                if header is not None:
                    columns_to_keep_indices = [index for index, col in enumerate(header) if col not in columns_to_remove]
                    if columns_to_keep_indices:
                        rows_to_keep = [[row[i] for i in columns_to_keep_indices] for row in reader]
                        with open(csv_file_path, 'w', newline='') as csvfile:
                            writer = csv.writer(csvfile)
                            writer.writerow([header[i] for i in columns_to_keep_indices])
                            writer.writerows(rows_to_keep)
                        print(f"Columns removed from {csv_file_path}")
                    else:
                        print("No columns to keep.")
                else:
                    print("CSV file is empty.")
        else:
            print("CSV file is empty.")
    else:
        print("CSV file does not exist.")

folder_path = select_directory()
if folder_path:
    split_files_by_delimiter(folder_path)
    folder_name = os.path.basename(folder_path)
    csv_file_name = f"{folder_name}_output.csv"
    csv_file_path = os.path.normpath(os.path.join(folder_path, csv_file_name))
    columns_to_remove = ['Audio #', 'Dictation ID', 'ID']
    remove_columns_from_csv(csv_file_path, columns_to_remove)
else:
    print("No directory selected. Exiting.")