from pymongo import MongoClient
import PyPDF2
import pandas as pd
from fpdf import FPDF
import os

# MongoDB Connection
MONGO_URI = "mongodb://localhost:27017/"
DB_NAME = "email_processor"
COLLECTION_NAME = "pdf_data"

def connect_mongo():
    """Connect to MongoDB and return the collection."""
    client = MongoClient(MONGO_URI)
    db = client[DB_NAME]
    return db[COLLECTION_NAME]

def extract_pdf_text(pdf_path):
    """Extract text from a PDF file."""
    text = ""
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page in pdf_reader.pages:
            text += page.extract_text()
    return text

def save_to_mongo(collection, data):
    """Save data to MongoDB."""
    collection.insert_one(data)

def save_as_excel(text, output_path):
    """Save text content into an Excel file."""
    df = pd.DataFrame({'Content': [text]})
    df.to_excel(output_path, index=False)

def save_as_pdf(text, output_path):
    """Save text content into a PDF file with Unicode support."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Add a Unicode font (replace with the actual path to a TTF font file)
    font_path = r'C:/Windows/Fonts/times.ttf'  # Replace with the full path to a Unicode-compatible font
    pdf.add_font('ArialUnicode', '', font_path, uni=True)
    pdf.set_font('ArialUnicode', size=12)

    pdf.multi_cell(0, 10, text)
    pdf.output(output_path)

def process_pdf(filepath):
    """Extract, store, and convert PDF data."""
    collection = connect_mongo()

    # Extract text from PDF
    text = extract_pdf_text(filepath)

    # Save to MongoDB
    document = {
        "filename": os.path.basename(filepath),
        "content": text
    }
    save_to_mongo(collection, document)
    print(f"Data saved to MongoDB: {document['filename']}")

    # Save to Excel and PDF
    excel_path = filepath.replace(".pdf", ".xlsx")
    pdf_path = filepath.replace(".pdf", "_converted.pdf")
    save_as_excel(text, excel_path)
    save_as_pdf(text, pdf_path)

if __name__ == "__main__":
    # Example file path
    pdf_file_path = r"D:\Downloads\generated_output_with_image.pdf"  # Replace with the actual file path

    try:
        process_pdf(pdf_file_path)
    except Exception as e:
        print(f"An error occurred: {e}")