import os
import win32com.client
from win32com.client import constants
import time

def convert_docx_to_pdf(docx_path, pdf_path):
    # Normalize the file paths
    #docx_path = os.path.normpath(docx_path)
    #pdf_path = os.path.normpath(pdf_path)

    # Add double quotes around file paths to handle spaces
    docx_path = f'{docx_path}'
    pdf_path = f'{pdf_path}'

    # Create a new Word Application object
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    time.sleep(1)

    try:
        # Open the source DOCX file
        print(docx_path)
        print(pdf_path)
        doc = word.Documents.Open(docx_path)
        
        print("after open")
        time.sleep(1)
        # Save the file as PDF
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 corresponds to the PDF format
        print("after save")
        
        # Close the source document
        doc.Close()

        print(f"{docx_path} converted to PDF: {pdf_path}")
    except Exception as e:
        print(f"Error converting {docx_path} to PDF: {e}")
    finally:
        # Quit the Word Application
        word.Quit()

def convert_folder_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file in os.listdir(input_folder):
        if file.endswith(".docx"):
            docx_path = os.path.join(input_folder, file)
            pdf_path = os.path.join(output_folder, f"{os.path.splitext(file)[0]}.pdf")
            print(pdf_path)
            convert_docx_to_pdf(docx_path, pdf_path)

if __name__ == "__main__":
    input_folder = "input\\folder"
    output_folder = "output\\folder"
    convert_folder_to_pdf(input_folder, output_folder)