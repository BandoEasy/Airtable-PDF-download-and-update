import os
import requests
import pdfplumber
from pyairtable import Api
from io import BytesIO
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from docx import Document
from fpdf import FPDF

# Airtable Configuration
AIRTABLE_API_KEY = 'patvS8osFDngCE0iW.4728217ddb2f14ecb1a3267e685ea11382100608c5f39205d0249eae7b8356c3'  # Replace with your Airtable Personal Access Token (PAT)
BASE_ID = 'appMoJrdRNUc086rC'  # Replace with your Airtable Base ID
TABLE_NAME = 'Bandi online '  # Ensure the table name is correct
VIEW_NAME = 'Open'  # Ensure the view name is correct

# Airtable column names
LINK_BANDO_COL = 'Link Bando'  # Column where the PDF URL is stored
PDF_ATTACHMENT_COL = 'PDF'  # Column for PDF attachments
PDF_TEXT_COL = 'PDF text'  # Column for extracted text
ERROR_HANDLING_COL = 'PDF status'  # Column for handling error outcomes (Multiple Select)

# Directory for saving temporary text files and PDFs
output_directory = '/Users/it/desktop/PDF_files'  # Adjust this to where you want to save files

# Global variables for settings
MAX_AIRTABLE_TEXT_SIZE = 100000  # Airtable's limit on long text fields
PDF_STATUS_OPTIONS = ["Success", "Link Bando is empty", "Unsupported Document Type", "Error Downloading Document", "Error Converting Document", "Others", "Scanned PDF", "Text-based PDF"]

# Starting options (use one or the other):
start_row = 384-1  # Set this to the row index you want to start from (e.g., start from row 10)
start_codice = None  # Set this to a specific 'Codice' to start from (e.g., '12345'). Leave as None to use `start_row`

# Initialize the Airtable API client
api = Api(AIRTABLE_API_KEY)
table = api.table(BASE_ID, TABLE_NAME)

# Ensure the directory exists
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Function definitions
def download_document(url):
    """Download a document from a given URL."""
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()

        content_type = response.headers.get('Content-Type', '')
        if 'application/pdf' in content_type:
            return BytesIO(response.content), "PDF", "Success"
        elif 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' in content_type:
            return BytesIO(response.content), "DOCX", "Success"
        elif 'image/' in content_type:
            return BytesIO(response.content), "Image", "Success"
        else:
            return None, None, "Unsupported Document Type"
    except requests.exceptions.RequestException as e:
        return None, None, "Error Downloading Document"

def download_pdf_from_airtable(pdf_attachment_url):
    """Download a PDF file directly from Airtable's attachment field."""
    try:
        response = requests.get(pdf_attachment_url, stream=True)
        response.raise_for_status()
        return BytesIO(response.content)
    except requests.exceptions.RequestException as e:
        print(f"Error downloading PDF from Airtable: {e}")
        return None

def convert_docx_to_pdf(docx_stream, pdf_path):
    """Convert a Word document (DOCX) to a PDF."""
    try:
        document = Document(docx_stream)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for paragraph in document.paragraphs:
            pdf.cell(200, 10, txt=paragraph.text, ln=True)

        pdf.output(pdf_path)
        return "Success"
    except Exception as e:
        return f"Error converting DOCX to PDF: {e}"

def convert_image_to_pdf(image_stream, pdf_path):
    """Convert an image (JPG, PNG) to a PDF."""
    try:
        image = Image.open(image_stream)
        image.convert('RGB').save(pdf_path)
        return "Success"
    except Exception as e:
        return f"Error converting Image to PDF: {e}"

def convert_to_pdf(document_stream, doc_type, pdf_path):
    """Convert various document types to PDF."""
    if doc_type == "DOCX":
        return convert_docx_to_pdf(document_stream, pdf_path)
    elif doc_type == "Image":
        return convert_image_to_pdf(document_stream, pdf_path)
    else:
        return "Unsupported Document Type"

def is_scanned_pdf(pdf_path):
    """Check if the PDF is scanned by attempting to extract text from the first page."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                if page.extract_text():
                    return False  # Text-based PDF
        return True  # Scanned PDF (image-based)
    except Exception as e:
        print(f"Error checking if PDF is scanned: {e}")
        return True  # Assume it's scanned if there's an error

def ocr_image_from_pdf(pdf_path):
    """Convert each page of a PDF into an image and extract text using OCR."""
    extracted_text = ""
    pages = convert_from_path(pdf_path)  # Convert PDF pages to images
    for i, page_image in enumerate(pages):
        text = pytesseract.image_to_string(page_image, lang='ita')  # You can change 'ita' to another language
        extracted_text += f"--- Page {i+1} ---\n{text}\n"
    return extracted_text

def extract_text_from_pdf(pdf_stream):
    """Extract text from a PDF using pdfplumber."""
    extracted_text = ""
    with pdfplumber.open(pdf_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                extracted_text += text
    return extracted_text

def process_grant_document(pdf_url, save_to, pdf_temp_path, use_airtable_pdf=False, pdf_stream=None):
    """Download and process a document, converting to PDF if needed."""
    try:
        if use_airtable_pdf:
            with open(pdf_temp_path, 'wb') as temp_pdf_file:
                temp_pdf_file.write(pdf_stream.read())
        else:
            document_stream, doc_type, download_status = download_document(pdf_url)
            if not document_stream:
                return None, download_status

            if doc_type != "PDF":
                convert_status = convert_to_pdf(document_stream, doc_type, pdf_temp_path)
                if convert_status != "Success":
                    return None, convert_status
            else:
                with open(pdf_temp_path, 'wb') as temp_pdf_file:
                    temp_pdf_file.write(document_stream.read())

        if is_scanned_pdf(pdf_temp_path):
            print(f"The document at {pdf_url} appears to be scanned. Using OCR to extract text.")
            extracted_text = ocr_image_from_pdf(pdf_temp_path)
            error_handling = ["Success", "Scanned PDF"]
        else:
            print(f"The document at {pdf_url} contains text. Extracting text directly.")
            extracted_text = extract_text_from_pdf(pdf_temp_path)
            error_handling = ["Success", "Text-based PDF"]

        with open(save_to, 'w', encoding='utf-8') as text_file:
            text_file.write(extracted_text)

        return extracted_text, error_handling
    except Exception as e:
        print(f"Error processing document from {pdf_url}: {e}")
        return None, ["Others"]

def upload_to_airtable(record_id, pdf_url, extracted_text, error_handling):
    """Upload the processed PDF URL, text, and error handling status to Airtable."""
    try:
        if len(extracted_text) > MAX_AIRTABLE_TEXT_SIZE:
            extracted_text = "Text is too large, please refer to the PDF for analysis."

        table.update(record_id, {PDF_ATTACHMENT_COL: [{'url': pdf_url}]})
        table.update(record_id, {
            PDF_TEXT_COL: extracted_text,
            ERROR_HANDLING_COL: error_handling
        })
        print(f"Successfully updated record {record_id}")
    except Exception as e:
        print(f"Error uploading to Airtable for record {record_id}: {e}")
        update_airtable_status(record_id, f"Failed to upload: {e}")

def update_airtable_status(record_id, status_message):
    """Update the PDF status column in Airtable for failed or successful records."""
    if status_message not in PDF_STATUS_OPTIONS:
        print(f"Warning: Status '{status_message}' is not a valid option in the 'PDF status' field.")
        status_message = "Others"
    try:
        table.update(record_id, {ERROR_HANDLING_COL: [status_message]})
        print(f"Updated status for record {record_id} with message: {status_message}")
    except Exception as e:
        print(f"Error updating Airtable status for record {record_id}: {e}")

def process_pdfs_from_airtable(airtable_records, start_row=0, start_codice=None):
    """Process documents (PDFs or others) from Airtable records, starting from a specific row or codice."""
    start_index = 0

    if start_codice:
        for i, record in enumerate(airtable_records):
            if record['fields'].get('Codice') == start_codice:
                start_index = i
                break
    else:
        start_index = start_row

    records_to_process = airtable_records[start_index:]

    for record in records_to_process:
        pdf_attachment = record['fields'].get(PDF_ATTACHMENT_COL)
        pdf_url = record['fields'].get(LINK_BANDO_COL)
        codice = record['fields'].get('Codice')

        sanitized_codice = codice.replace(' ', '_') if codice else 'temp'
        output_file_path = os.path.join(output_directory, f"{sanitized_codice}.txt")
        pdf_temp_path = os.path.join(output_directory, f"{sanitized_codice}.pdf")

        if pdf_attachment:
            pdf_attachment_url = pdf_attachment[0]['url']
            print(f"Processing PDF from Airtable attachment for record {record['id']}.")
            pdf_stream = download_pdf_from_airtable(pdf_attachment_url)
            if pdf_stream:
                extracted_text, error_handling = process_grant_document(pdf_attachment_url, output_file_path, pdf_temp_path, use_airtable_pdf=True, pdf_stream=pdf_stream)
                if extracted_text:
                    upload_to_airtable(record['id'], pdf_attachment_url, extracted_text, error_handling)
                else:
                    update_airtable_status(record['id'], "Error processing Airtable PDF")
            else:
                update_airtable_status(record['id'], "Error Downloading Document from Airtable Attachment")
        elif pdf_url:
            print(f"Processing document from Link Bando URL for record {record['id']}.")
            extracted_text, error_handling = process_grant_document(pdf_url, output_file_path, pdf_temp_path)
            if extracted_text:
                upload_to_airtable(record['id'], pdf_url, extracted_text, error_handling)
            else:
                update_airtable_status(record['id'], error_handling)
        else:
            print(f"Skipping record {record['id']} as both PDF attachment and Link Bando URL are missing.")
            update_airtable_status(record['id'], "Link Bando is empty")

# Main execution
if __name__ == "__main__":
    airtable_records = table.all(view=VIEW_NAME)
    if airtable_records:
        print(f"Fetched {len(airtable_records)} records from Airtable.")
        process_pdfs_from_airtable(airtable_records, start_row=start_row, start_codice=start_codice)
    else:
        print("No records fetched from Airtable.")
