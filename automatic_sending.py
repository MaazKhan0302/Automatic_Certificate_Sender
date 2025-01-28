import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Function to embed the name into the certificate PDF
def create_certificate(name, template_path, output_path, font_path):
    pdfmetrics.registerFont(TTFont('Tomorrow', font_path))
    c = canvas.Canvas("temp_overlay.pdf", pagesize=letter)
    font_size = 28.3
    c.setFont("Tomorrow", font_size)
    x_position = 11.50 * 28.35 #Adjust the x-axis of placement accordingly
    y_position = 10.20 * 28.35#Adjust the y-axis of placement accordingly
    text_width = c.stringWidth(name, "Tomorrow", font_size)

    if text_width > 200:
        font_size = (200 / text_width) * font_size
        c.setFont("Tomorrow", font_size)

    c.drawString(x_position, y_position, name)
    c.save()

    template_pdf = PdfReader(template_path)
    overlay_pdf = PdfReader("temp_overlay.pdf")
    output_pdf = PdfWriter()

    for i in range(len(template_pdf.pages)):
        page = template_pdf.pages[i]
        if i == 0:
            page.merge_page(overlay_pdf.pages[0])
        output_pdf.add_page(page)

    with open(output_path, "wb") as out_f:
        output_pdf.write(out_f)

# Function to send the certificate via email
def send_email(to_email, name, attachment_path):
    from_email = "Your Email Address" #Enter Your Email Address
    from_password = "Your Password" # Enter You Email Password

    # Create a personalized email message
    subject = "Congratulations on Your Achievement!"
    body = (f"Dear {name},\n\n"
            "🎉 Congratulations on achieving your certificate! 🎉\n\n"
            "This accomplishment reflects your hard work and dedication. "
            "We are proud of your achievement and excited to see what you will accomplish next.\n\n"
            "Please find your certificate attached.\n\n"
            "Best wishes,\n"
            "KCITES Team")

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    msg.attach(MIMEBase('application', 'octet-stream'))
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={attachment_path}')
        msg.attach(part)

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(from_email, from_password)
        server.send_message(msg)

if __name__ == "__main__":
    data = pd.read_csv("Feedback Form Responses - Form Responses 1.csv")
    template_path = "Certificate.pdf"
    font_path = "Tomorrow-Bold.ttf"

    for index, row in data.iterrows():
        name = row['Full Name']
        output_path = f"{name.replace(' ', '_')}_Certificate.pdf"

        create_certificate(name, template_path, output_path, font_path)
        recipient_email = row['Email']
        # Send email with the certificate
        send_email(recipient_email, name, output_path)

        # Delete the certificate after sending
        os.remove(output_path)
