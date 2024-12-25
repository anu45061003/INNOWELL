import os
import pdfplumber
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from openpyxl import Workbook

# Function to extract the first image from the input PDF
def extract_image_from_pdf(input_pdf_path, output_image_path):
    with pdfplumber.open(input_pdf_path) as pdf:
        for page in pdf.pages:
            if page.images:
                image_data = page.images[0]  # Get the first image
                x0, y0, x1, y1 = image_data["x0"], image_data["y0"], image_data["x1"], image_data["y1"]
                cropped_image = page.within_bbox((x0, y0, x1, y1)).to_image()
                cropped_image.save(output_image_path, format="PNG")
                return True
    return False

# Function to create a structured PDF including an image and details
def create_pdf_with_image(output_path, spec_sheet_data, header_info, image_path):
    pdf = SimpleDocTemplate(output_path, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    # Add header information
    for line in header_info:
        elements.append(Paragraph(line, styles['Normal']))
    elements.append(Spacer(1, 0.2 * inch))

    # Add image if available
    if os.path.exists(image_path):
        elements.append(Image(image_path, width=2 * inch, height=2 * inch))
        elements.append(Spacer(1, 0.2 * inch))

    # Create a table with the specification sheet data
    table = Table(spec_sheet_data)

    # Set table style
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    table.setStyle(style)

    # Add the table to the elements list
    elements.append(table)

    # Build the PDF
    pdf.build(elements)

# Function to create an Excel file
def create_excel(output_path, spec_sheet_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Vegetable Data"

    for row in spec_sheet_data:
        ws.append(row)

    wb.save(output_path)

# Function to send email
def send_email(to_email, subject, body, attachments=[]):
    from_email = "anusudha045@gmail.com"  # Replace with your email
    password = "mudb pfnn fkzc zqak"  # Use the App Password generated from Google

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach files
    for file_path in attachments:
        with open(file_path, "rb") as attachment:
            part = MIMEApplication(attachment.read(), Name=os.path.basename(file_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
            msg.attach(part)

    # Send the email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(from_email, password)
            server.send_message(msg)
            print("Email sent successfully!")
    except Exception as e:
        print(f"Error sending email: {e}")

# Main procedure
def main():
    input_pdf_path = r"D:\\Downloads\\Vegetable Analysis Report.pdf"  # Ensure this path is correct
    output_pdf_path = r"D:\\Downloads\\generated_output_with_image.pdf"
    extracted_image_path = r"D:\\Downloads\\extracted_image.png"
    output_excel_path = r"D:\\Downloads\\vegetable_data.xlsx"

    # Extract image from the input PDF
    image_extracted = extract_image_from_pdf(input_pdf_path, extracted_image_path)

    # Prepare the header information
    header_info = [
        "<b>Vegetable Analysis Report</b>",
        "Properties:",
        "Product Type: Fresh Vegetables                                Order number: 700696773017",
        "Brand: MORL                                                                Focus: pricing",
        "Commodity: Vegetables                                            E-mail: sudhaanu563@outlook.com",
        "Address: Tower-B, 7th Floor, IBC Knowledge Park, Nagar, S.G. Palya, Bengaluru, Karnataka – 560029",
    ]

    # Prepare the specification sheet data
    spec_sheet_data = [
        ["Vegetable", "Quantity (MT)", "Average Price (INR/KG)", "Seasonal Availability", "Total Value (INR)"]
    ]

    vegetable_data = [
        ["Tomato", 198, 10, "Year-round"],
        ["Onion", 228, 7, "Kharif and rabi seasons"],
        ["Potato", 491, 8, "Primarily December to March"],
        ["Green Peas", 35, 9, "Winter months"],
        ["Carrot", 15, 18, "Mainly winter"],
        ["Cabbage", 135, 20, "Kharif and rabi seasons"],
        ["Spinach", 25, 10, "Winter months"],
        ["Cauliflower", 12, 15, "Mainly winter"],
        ["Bell Pepper", 8, 27, "Mainly winter"],
        ["Radish", 9, 30, "Winter months"],
        ["Cucumber", 7, 45, "Summer"],
    ]

    for veg in vegetable_data:
        vegetable, quantity, avg_price, availability = veg
        total_value = quantity * avg_price
        spec_sheet_data.append([vegetable, quantity, avg_price, availability, total_value])

    # Create a PDF with image and header information
    create_pdf_with_image(output_pdf_path, spec_sheet_data, header_info, extracted_image_path)

    # Create an Excel file
    create_excel(output_excel_path, spec_sheet_data)

    # Send email with both PDF and Excel as attachments
    client_email = "anumurugan045@gmail.com.com"  # Replace with the recipient's email
    subject = "Vegetable Analysis Report"
    body = "Please find attached the vegetable analysis report."

    send_email(client_email, subject, body, [output_pdf_path, output_excel_path])

if __name__ == "__main__":
    main()
