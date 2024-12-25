import imaplib
import email
import os
from tabula import read_pdf
from reportlab.pdfgen import canvas
from django.core.mail import EmailMessage

def fetch_pdf_emails_debug(imap_server, email_user, email_password, download_folder):
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_user, email_password)
        mail.select("inbox")
        status, messages = mail.search(None, "ALL")
        
        if status != "OK" or not messages[0]:
            print("No emails found.")
            return None
        
        for num in messages[0].split():
            status, msg_data = mail.fetch(num, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    print(f"Subject: {msg['subject']}")
                    for part in msg.walk():
                        if part.get_content_maintype() == "multipart":
                            continue
                        if part.get("Content-Disposition"):
                            filename = part.get_filename()
                            if filename and filename.endswith(".pdf"):
                                print(f"Found PDF: {filename}")
                                return filename
        print("No PDF attachments found.")
        return None
    except Exception as e:
        print(f"Error fetching emails: {e}")
        return None

# Test fetching
fetch_pdf_emails_debug("imap.gmail.com", "your_email@example.com", "your_app_password", "downloads")


# Function to extract table data and calculate total sales
def extract_table_data(pdf_path):
    """
    Extracts table data from a PDF and calculates the total sales.

    Args:
        pdf_path (str): Path to the PDF file.

    Returns:
        float: Total sales value.
    """
    try:
        tables = read_pdf(pdf_path, pages="all", multiple_tables=True)
        total_sales = 0
        for table in tables:
            if "Sales" in table.columns:
                total_sales += table["Sales"].sum()
        print(f"Total Sales: {total_sales}")
        return total_sales
    except Exception as e:
        print(f"Error extracting table data: {e}")
        return 0

# Function to generate the output PDF
def generate_output_pdf(output_path, total_sales):
    """
    Generates an output PDF with total sales data.

    Args:
        output_path (str): Path to save the output PDF.
        total_sales (float): Total sales value.
    """
    try:
        c = canvas.Canvas(output_path)
        c.drawString(100, 750, f"Total Sales: {total_sales}")
        c.save()
        print(f"Output PDF generated: {output_path}")
    except Exception as e:
        print(f"Error generating output PDF: {e}")

# Function to send the output PDF via email
def send_email_with_attachment(to_email, subject, body, attachment_path):
    """
    Sends an email with the output PDF as an attachment.

    Args:
        to_email (str): Recipient's email address.
        subject (str): Email subject.
        body (str): Email body.
        attachment_path (str): Path to the PDF attachment.
    """
    try:
        email = EmailMessage(subject, body, to=[to_email])
        email.attach_file(attachment_path)
        email.send()
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Error sending email: {e}")

# Main workflow function
def process_sales_email():
    """
    End-to-end process to fetch, process, and respond to sales email.
    """
    # Email settings
    imap_server = "imap.gmail.com"
    email_user = "anumurugan045@gmail.com"  # Replace with your email
    email_password = "zhkq atkb cesv gst"   # Replace with your app password
    download_folder = "downloads"
    
    # Step 1: Fetch PDF
    pdf_path = fetch_pdf_emails_debug(imap_server, email_user, email_password, download_folder)  # Updated function name
    if not pdf_path:
        print("No PDF to process. Exiting.")
        return

    # Remaining steps...

    
    # Step 1: Fetch PDF
    pdf_path = fetch_pdf_emails(imap_server, email_user, email_password, download_folder)
    if not pdf_path:
        print("No PDF to process. Exiting.")
        return

    # Step 2: Extract data and calculate sales
    total_sales = extract_table_data(pdf_path)

    # Step 3: Generate output PDF
    output_path = "outputs/sales_report.pdf"
    os.makedirs("outputs", exist_ok=True)  # Ensure the outputs directory exists
    generate_output_pdf(output_path, total_sales)

    # Step 4: Send output email
    send_email_with_attachment(
        to_email="anumurugan045@gmail.com",  # Replace with client's email
        subject="Sales Report",
        body="Here is your sales report.",
        attachment_path=output_path
    )

# Run the process
if __name__ == "__main__":
    process_sales_email()
