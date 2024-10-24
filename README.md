# weekly-report
import os
import fitz  # PyMuPDF for PDFs
import pytesseract  # OCR for images
from pdf2image import convert_from_path
import pandas as pd
from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import logging
import spacy
from reportlab.pdfgen import canvas
from textblob import TextBlob
import re
import datetime

# Set up logging
logging.basicConfig(filename='project_report.log', level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

nlp = spacy.load("en_core_web_sm")

# Function to identify document type based on text content
def identify_document_type(text):
    doc = nlp(text.lower())
    if re.search(r"\binspection request\b|\bIR\b", text, re.I):
        return "IR"
    elif re.search(r"\bmaterial inspection request\b|\bMIR\b", text, re.I):
        return "MIR"
    elif re.search(r"\bpreconstruction\b|\bPreCon\b", text, re.I):
        return "PreCon"
    elif re.search(r"\bcorrective action request\b|\bCAR\b", text, re.I):
        return "CAR"
    else:
        return "Unknown"

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text += page.get_text("text")
        logging.info(f"Successfully extracted text from PDF: {pdf_path}")
        return text
    except Exception as e:
        logging.error(f"Failed to process PDF {pdf_path}: {str(e)}")
        return ""

# Function to extract text from image using OCR
def extract_text_from_image(image_path):
    try:
        image = Image.open(image_path)
        text = pytesseract.image_to_string(image)
        logging.info(f"Successfully extracted text from image: {image_path}")
        return text
    except Exception as e:
        logging.error(f"Failed to process image {image_path}: {str(e)}")
        return ""

# Function to handle PDF as images (for PDFs that need OCR)
def extract_text_from_pdf_as_image(pdf_path):
    try:
        images = convert_from_path(pdf_path)
        text = ""
        for image in images:
            text += pytesseract.image_to_string(image)
        logging.info(f"Successfully performed OCR on PDF as image: {pdf_path}")
        return text
    except Exception as e:
        logging.error(f"Failed to process PDF as image {pdf_path}: {str(e)}")
        return ""

# Function to extract dates with better regex patterns
def extract_dates(text):
    date_patterns = [
        r"\b(\d{4}[-/]\d{2}[-/]\d{2})\b",  # YYYY-MM-DD
        r"\b(\d{2}[-/]\d{2}[-/]\d{4})\b"   # DD/MM/YYYY
    ]
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    return "Date Not Found"

# Function to extract IR data (example for structured data extraction)
def extract_ir_data(text):
    ir_data = []
    lines = text.split("\n")
    discipline, submittal_no, status, date_issued = None, None, None, None
    for line in lines:
        if "Discipline:" in line:
            discipline = re.search(r"Discipline: (.+)", line).group(1)
        if "Submittal No.:" in line:
            submittal_no = re.search(r"Submittal No.: (.+)", line).group(1)
        if "Status:" in line:
            status = re.search(r"Status: (.+)", line).group(1)
        if "Date Issued:" in line:
            date_issued = extract_dates(line)  # Improved date extraction
        if all([discipline, submittal_no, status, date_issued]):
            ir_data.append([discipline, submittal_no, status, date_issued])
            discipline, submittal_no, status, date_issued = None, None, None, None
    return pd.DataFrame(ir_data, columns=['Discipline', 'Submittal No.', 'Status', 'Date Issued'])

# Function to generate PDF report
def generate_pdf_report(ir_data, output_pdf):
    try:
        c = canvas.Canvas(output_pdf)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, 800, "Weekly Progress Report - BoSS Cash Center Project")

        # Adding sections to the report
        c.setFont("Helvetica", 12)
        c.drawString(50, 770, "1. Inspection Requests (IRs):")
        y = 750
        for index, row in ir_data.iterrows():
            c.drawString(50, y, f"Discipline: {row['Discipline']}, Submittal No.: {row['Submittal No.']}, Status: {row['Status']}, Date Issued: {row['Date Issued']}")
            y -= 20
        c.save()
        logging.info(f"PDF report generated successfully: {output_pdf}")
    except Exception as e:
        logging.error(f"Failed to generate PDF report: {str(e)}")

# Main processing function to handle files
def process_files(file_list):
    ir_data_list = []
    for file in file_list:
        if file.endswith(".pdf"):
            text = extract_text_from_pdf(file)
            if not text.strip():
                text = extract_text_from_pdf_as_image(file)
        elif file.endswith((".png", ".jpg", ".jpeg")):
            text = extract_text_from_image(file)
        else:
            logging.warning(f"Unsupported file format: {file}")
            continue
        
        doc_type = identify_document_type(text)
        if doc_type == "IR":
            ir_data_list.append(extract_ir_data(text))
        else:
            logging.warning(f"Document type {doc_type} not handled: {file}")
    
    # Combine all IR data
    ir_data_combined = pd.concat(ir_data_list, ignore_index=True) if ir_data_list else pd.DataFrame()
    
    # Generate PDF report
    output_pdf_path = './weekly_report.pdf'
    generate_pdf_report(ir_data_combined, output_pdf_path)

# GUI using Tkinter
class ReportGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Project Report Generator")
        self.geometry("500x300")
        
        self.file_list = []
        
        self.create_widgets()

    def create_widgets(self):
        # Label and Buttons
        self.label = tk.Label(self, text="BoSS Cash Center Report Generator", font=("Helvetica", 16))
        self.label.pack(pady=10)
        
        self.upload_btn = tk.Button(self, text="Upload Files", command=self.upload_files)
        self.upload_btn.pack(pady=5)

        self.generate_btn = tk.Button(self, text="Generate Report", command=self.generate_report)
        self.generate_btn.pack(pady=5)

        # Status box
        self.status_box = tk.Text(self, height=10, width=60)
        self.status_box.pack(pady=10)
        
    def upload_files(self):
        # Select multiple files
        files = filedialog.askopenfilenames(title="Select Project Documents", filetypes=[("All files", "*.*")])
        self.file_list = list(files)
        self.status_box.insert(tk.END, f"Selected files: {len(self.file_list)}\n")
        
    def generate_report(self):
        if not self.file_list:
            messagebox.showwarning("No Files", "Please upload files before generating the report.")
            return
        
        self.status_box.insert(tk.END, "Processing files...\n")
        self.update()
        
        # Call processing function
        process_files(self.file_list)
        
        self.status_box.insert(tk.END, "Report generated successfully! Check 'weekly_report.pdf'.\n")
        self.update()

if __name__ == "__main__":
    app = ReportGeneratorApp()
    app.mainloop()
