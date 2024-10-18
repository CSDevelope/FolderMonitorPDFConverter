import os
import time
import threading
import tempfile  # Add this import
import json
import tkinter as tk
from tkinter import filedialog, Button, messagebox, Label
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from docx import Document
from fpdf import FPDF
from PIL import Image, ImageTk
import win32com.client
from pywintypes import com_error
import pythoncom


CONFIG_FILE = 'config.json'

if not os.path.exists(CONFIG_FILE):
    config = {
        'pdf_folder': './pdfs',
        'watch_folder': './watch_folder'
    }
    with open(CONFIG_FILE, 'w') as config_file:
        json.dump(config, config_file, indent=4)
else:
    with open(CONFIG_FILE, 'r') as config_file:
        config = json.load(config_file)

PDF_FOLDER = os.path.abspath(config['pdf_folder'])
WATCH_FOLDER = os.path.abspath(config['watch_folder'])

os.makedirs(PDF_FOLDER, exist_ok=True)
os.makedirs(WATCH_FOLDER, exist_ok=True)
os.makedirs('./fonts', exist_ok=True)

class FileHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            filepath = event.src_path
            file_extension = os.path.splitext(filepath)[1].lower()
            pdf_filename = os.path.basename(filepath).rsplit('.', 1)[0] + '.pdf'
            pdf_filepath = os.path.join(PDF_FOLDER, pdf_filename)
            
            try:
                if file_extension == '.docx':
                    convert_docx_to_pdf(filepath, pdf_filepath)
                elif file_extension in ['.png', '.jpg', '.jpeg']:
                    convert_image_to_pdf(filepath, pdf_filepath)
                elif file_extension in ['.xls', '.xlsx']:
                    convert_excel_to_pdf(filepath, pdf_filepath)
                elif file_extension == '.tmp':
                    # Ignore temporary files
                    print(f"Ignored temporary file: {file_extension}")
                else:
                    print(f"Unsupported file type: {file_extension}. Deleting file.")
                    time.sleep(1)  # Delay to ensure file is fully written before deletion
                    os.remove(filepath)
            except Exception as e:
                print(f"Failed to convert {filepath}: {e}")


def convert_docx_to_pdf(docx_path, pdf_path):
    doc = Document(docx_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'fonts/DejaVuSans.ttf')
    pdf.set_font('DejaVu', size=10)

    for para in doc.paragraphs:
        if para.text.strip():
            pdf.multi_cell(190, 10, para.text)
            pdf.ln(5)

    pdf.output(pdf_path)

def convert_image_to_pdf(image_path, pdf_path):
    image = Image.open(image_path)
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'fonts/DejaVuSans.ttf')
    pdf.set_font('DejaVu', size=10)
    pdf.image(image_path, x=10, y=10, w=190)
    pdf.output(pdf_path)

def convert_excel_to_pdf(excel_path, pdf_path):
    excel_path = os.path.abspath(excel_path)  # Get the absolute path for the Excel file
    pdf_path = os.path.abspath(pdf_path)      # Get the absolute path for the PDF output
    
    pythoncom.CoInitialize()  # Initialize COM library
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(excel_path)
        temp_pdf_path = os.path.join(tempfile.gettempdir(), os.path.basename(pdf_path))
        wb.WorkSheets.Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, temp_pdf_path)
        # Move the temporary PDF to the desired location
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        os.rename(temp_pdf_path, pdf_path)
    except com_error as e:
        print('Excel to PDF conversion failed:', e)
    finally:
        if 'wb' in locals():  
            wb.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()  


def choose_watch_folder():
    global WATCH_FOLDER, config
    watch_folder = filedialog.askdirectory(title="Select Folder to Monitor")
    if watch_folder:
        WATCH_FOLDER = os.path.abspath(watch_folder)
        config['watch_folder'] = WATCH_FOLDER
        os.makedirs(WATCH_FOLDER, exist_ok=True)
        with open(CONFIG_FILE, 'w') as config_file:
            json.dump(config, config_file, indent=4)


def choose_output_folder():
    global PDF_FOLDER, config
    pdf_folder = filedialog.askdirectory(title="Select Output Folder for PDFs")
    if pdf_folder:
        PDF_FOLDER = os.path.abspath(pdf_folder)
        config['pdf_folder'] = PDF_FOLDER
        os.makedirs(PDF_FOLDER, exist_ok=True)
        with open(CONFIG_FILE, 'w') as config_file:
            json.dump(config, config_file, indent=4)


def show_information():
    instructions = (
        "1. Click 'Choose Monitored Folder' to select the folder where the original files are located.\n"
        "2. Click 'Choose Output Folder' to select the folder where you want the converted PDFs to be saved.\n"
        "3. The application will automatically convert supported files (.docx, .png, .jpg, .jpeg, .xls, .xlsx) to PDFs and save them in the output folder.\n"
        "4. Unsupported files will be deleted, and temporary files will be ignored."
    )
    messagebox.showinfo("Information", instructions)


def start_monitoring():
    event_handler = FileHandler()
    observer = Observer()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    
    root = tk.Tk()
    root.title("Convert Files to PDF")
    root.geometry("600x400")
    root.iconbitmap('./images/dhl.ico')

    
    background_image = Image.open('./images/f1.png')
    background_photo = ImageTk.PhotoImage(background_image)
    background_label = tk.Label(root, image=background_photo)
    background_label.place(relwidth=1, relheight=1)

    
    choose_watch_button = Button(root, text="Choose Monitored Folder", command=choose_watch_folder)
    choose_watch_button.place(relx=0.5, rely=0.4, anchor='center')

    choose_output_button = Button(root, text="Choose Output Folder", command=choose_output_folder)
    choose_output_button.place(relx=0.5, rely=0.5, anchor='center')

    
    information_button = Button(root, text="Information", command=show_information)
    information_button.place(relx=0.5, rely=0.6, anchor='center')

    
    contact_label = Label(root, text="Questions? Contact andrew.tufarella@dhl.com", font=("Helvetica", 10), fg="#555555", bg="#f1f1f1")
    contact_label.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-10)

    
    monitoring_thread = threading.Thread(target=start_monitoring)
    monitoring_thread.daemon = True
    monitoring_thread.start()

    
    root.mainloop()
