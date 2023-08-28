import pandas as pd
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Read Excel data
excel_file = 'names.xlsx'
df = pd.read_excel(excel_file)

def set_font(c, font, font_size, rgb):
    c.setFont(font, font_size)
    c.setFillColorRGB(rgb[0], rgb[1], rgb[2])

def draw_centered_text(c, text, y, font, font_size, rgb):
    set_font(c, font, font_size, rgb)
    page_width, _ = c._pagesize  # Get the width of the current page size
    text_width = c.stringWidth(text, font, font_size)  # Calculate the width of the text
    text_center_x = (page_width - text_width) / 2  # Calculate the x-coordinate for center alignment
    c.drawString(text_center_x, y, text)

def draw_centered_text_two_colors(c, text1, text2, y, font, font_size, rgb1, rgb2):
    set_font(c, font, font_size, rgb1)
    page_width, _ = c._pagesize  # Get the width of the current page size
    text_width = c.stringWidth(text1+text2, font, font_size)  # Calculate the width of the total text
    text_width_1 = c.stringWidth(text1, font, font_size)
    
    text_center_x = (page_width - text_width) / 2  # Calculate the x-coordinate for center alignment
    c.drawString(text_center_x, y, text1)
    
    set_font(c, font, font_size, rgb2)
    c.drawString(text_center_x + text_width_1, y, text2)

def add_scaled_image(c, img_path, target_height, x, y):
    img = Image(img_path)
    height = target_height
    img_width = img.drawWidth
    img_height = img.drawHeight
    # Calculate the scaling factor to fit the image within the header height
    scaling_factor = height / img_height
    scaled_width = img_width * scaling_factor
    scaled_height = height
    c.drawImage(img_path, x, y, width=scaled_width, height=scaled_height)

def draw_dash_line(c, rgb, x1, y1, x2, y2):
    # Define dash pattern (on, off) and draw line
    dash_pattern = (6, 4)  # 6 points on, 4 points off
    c.setDash(dash_pattern, 0)  # Set the dash pattern
    c.setStrokeColorRGB(*rgb)
    c.line(x1, y1, x2, y2)

# Loop through each row and generate PDF certificate
for index, row in df.iterrows():
    first_name = row['Name']
    last_name = row['Last Name']
    program = row['Program']
    session = row['Session']
    duration = row['Duration'] # in minutes
    pdf_filename = f'Certificado{first_name}{last_name}.pdf'

    # Font, size and color
    # Check https://www.dafont.com/ for more fonts
    pdfmetrics.registerFont(TTFont('HandwrittenFont', 'fonts/Handwritten.ttf'))
    pdfmetrics.registerFont(TTFont('RegularFont', 'fonts/CaviarDreams.ttf'))
    pdfmetrics.registerFont(TTFont('RegularFontBold', 'fonts/CaviarDreams_Bold.ttf'))
    font_size = 22
    font_color = (0, 0, 0)
    special_color = (0, 0, 0.7)

    # Create a PDF with landscape orientation
    c = canvas.Canvas(pdf_filename, pagesize=landscape(letter))
    page_height = 792
    page_width = 612+200

    # Define header dimensions 
    header_width = page_width
    background_color = (0, 0, 0) # Black color
    context_beginning = 400
    c.setFillColorRGB(*background_color)
    c.rect(0, context_beginning, page_width, page_height, fill=True, stroke=False)
    # Add logo or content within the header area
    header_path = 'photos/header.jpg'
    add_scaled_image(c, header_path, page_height - context_beginning, 0, 400)
    
    # Define content size
    background_color = (1, 1, 1)  # white color
    c.setFillColorRGB(*background_color)
    c.rect(0, 0, page_width, context_beginning, fill=True, stroke=False)

    # Draw text below the header
    companyName = "Company Name"
    draw_centered_text_two_colors(c, companyName, " certifies that:", 330, "RegularFontBold", font_size, special_color,font_color)

    draw_dash_line(c, (0.7, 0.7, 0.7), 0, 270, page_width, 270)
    draw_centered_text(c, f'{first_name} {last_name}', 280, "HandwrittenFont", font_size + 20, font_color)
    
    draw_centered_text(c, f'attended the training for {program} on', 220, "RegularFont", font_size+4, font_color)
    draw_centered_text(c, f'{session}', 180, "RegularFontBold", font_size+6, font_color)
    draw_centered_text(c, f'Duration: {duration} minutes', 150, "RegularFont", font_size - 4, font_color)

    # Add logo or content within the footer area
    footer_path = 'photos/logo.png'
    add_scaled_image(c, footer_path, 50, page_height/1.3, 10)


    c.save()

    print(f'Generated {pdf_filename}')

    # # Send email with the generated PDF as an attachment
    # sender_email = "email@gmail.com"  # Replace with your sender email
    # sender_password = "password"  # Replace with your sender email password
    # recipient_email = row['Mail']  # Get the recipient's email address from the Excel row

    # msg = MIMEMultipart()
    # msg['From'] = sender_email
    # msg['To'] = recipient_email
    # # Add email content
    # email_content = f"Dear {first_name},\n\nThank you for participating in the {program} training session. " \
    #                 f"Please find your certificate attached.\n\nBest regards,\nYour Organization"

    # # Attach the email content
    # msg.attach(MIMEText(email_content, 'plain'))

    # # Attach the PDF to the email
    # with open(pdf_filename, "rb") as attachment:
    #     part = MIMEApplication(attachment.read(), Name=pdf_filename)
    #     part['Content-Disposition'] = f'attachment; filename="{pdf_filename}"'
    #     msg.attach(part)

    # # Send the email
    # with smtplib.SMTP('smtp.gmail.com', 587) as server:  # Replace with your SMTP server details
    #     server.starttls()
    #     server.login(sender_email, sender_password)
    #     server.sendmail(sender_email, recipient_email, msg.as_string())

    # print(f'Sent email with {pdf_filename} to {recipient_email}')

print('PDF generation complete.')


    # (0,792) -------------- (612,792)
    #    |                      |
    #    |                      |
    #    |        Page          |
    #    |                      |
    #    |                      |
    # (0,0) ----------------- (612,0)
