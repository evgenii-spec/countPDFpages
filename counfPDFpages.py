import os
import PyPDF2
import pandas as pd

def get_pdf_page_count(pdf_path):
    try:
        with open(pdf_path, "rb") as fr:
            reader = PyPDF2.PdfReader(fr)
            return len(reader.pages)
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return None

def get_folder_name_from_path(path):
    parts = path.split(os.sep)
    if len(parts) > 1:
        return parts[-2] + '/' + parts[-1]
    return parts[-1]

def list_pdf_pages_in_directory(directory, output_file):
    data = {"Folder Name": [], "File Name": [], "Number of Pages": []}
    
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(".pdf"):
                full_path = os.path.join(root, file)
                page_count = get_pdf_page_count(full_path)
                if page_count is not None:
                    data["Folder Name"].append(get_folder_name_from_path(root))
                    data["File Name"].append(file)
                    data["Number of Pages"].append(page_count)
                    
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False, engine='openpyxl')

if __name__ == "__main__":
    dir_path = input("Enter the directory path: ")
    output_path = os.path.join(dir_path, "pdf_results.xlsx")
    list_pdf_pages_in_directory(dir_path, output_path)
    print(f"Results written to {output_path}")
