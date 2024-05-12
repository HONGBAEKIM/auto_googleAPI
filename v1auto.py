#Google APIs using a JSON key
from oauth2client.service_account import ServiceAccountCredentials

#Google Sheets API
import gspread 

#make requests to the Google API service (Google Drive, Google Sheets)
from googleapiclient.discovery import build

#argv as perameters
import sys 

#User pdf_path
import os

#uploading files 
from googleapiclient.http import MediaFileUpload

#making it suitable for constructing complex email messages
from email.mime.multipart import MIMEMultipart

#create email messages with plain text content
from email.mime.text import MIMEText

#mimebase is Python's email package. It provides methods and attributes common to all MIME types.
from email.mime.base import MIMEBase

#For safety, When you attach a file to an email message, it needs to be encoded. 
from email import encoders

#SMTP is the standard protocol for sending emails across the internet.
import smtplib



# Define the scope and credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('/Users/hongbae/etc/sinuous-set-419712-9f824ab71b8a.json', scope)

# Authenticate with Google Sheets API
client = gspread.authorize(credentials)

def update_spreadsheet(day):

    # Open the Google Sheets spreadsheet by title or URL
    spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/10ehmE9EjW3D5edrAI9vYqFrsmd0FfqfshsRXqGNyNnQ')

    # Access the first worksheet in the spreadsheet
    worksheet = spreadsheet.sheet1

    # Update cell B23 with day
    worksheet.update([[day]], 'B22')  

    return worksheet
    

def export_spreadsheet_to_pdf(day):
    pdf_path = None  # Initialize pdf_path variable
    try:
        # Build the Google Drive service
        drive_service = build('drive', 'v3', credentials=credentials)

        # Define the PDF export URL with landscape orientation
        pdf_export_url = 'https://docs.google.com/spreadsheets/d/10ehmE9EjW3D5edrAI9vYqFrsmd0FfqfshsRXqGNyNnQ/export?format=pdf&portrait=false&size=A4'

        # Fetch the PDF content
        response = drive_service.files().export(fileId='10ehmE9EjW3D5edrAI9vYqFrsmd0FfqfshsRXqGNyNnQ', mimeType='application/pdf').execute()

        # Construct the full path where the PDF file will be saved
        pdf_directory = '/Users/hongbae/etc'
        pdf_file = f"{day.replace('/', '.')}_contract.pdf"
        pdf_path = os.path.join(pdf_directory, pdf_file)

        # Check if the response contains the content
        if response:
            # Write the content to a PDF file
            with open(pdf_path, 'wb') as output_file:
                output_file.write(response)
            print("PDF saved successfully.")
        else:
            print("An error occurred: No content in the response.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

    return pdf_path



def upload_to_drive(pdf_file):
    # Authenticate with Google Drive API
    drive_service = build('drive', 'v3', credentials=credentials)
    
    # Set file metadata including the parent folder ID
    file_metadata = {
        'name': pdf_file,
        'parents': ['1z-ZUY9rkucRmLintb3aEzf1Pz0jgsiJo']
    }

    # Upload file to Google Drive
    media = MediaFileUpload(pdf_file, mimetype='application/pdf')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print('File ID:', file.get('id'))



def read_password_from_file(file_path):
    with open(file_path, 'r') as file:
        password = file.read().strip()
    return password

# Import other required libraries and functions
def send_email(pdf_file, day, recipient_email):
    # Email configuration
    ###########################################
    sender_email = "xxxxxxxx@gmail.com"  # Replace with your email address
    ###########################################
    password = read_password_from_file('/Users/hongbae/pem/pass.txt')
    subject = f" subject "
    body = f"   Hello.\n\
    \n\
    \n\
    Have a great day!\n

    # Create message container
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Attach body to the email
    msg.attach(MIMEText(body, 'plain'))

    # Attach PDF file to the email
    attachment = open(pdf_file, "rb")

    # Create the attachment part
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % pdf_file)

    # Attach the attachment to the message
    msg.attach(part)

    # Connect to SMTP server and send email
    server = smtplib.SMTP('smtp.gmail.com', 587)  # Change the SMTP server accordingly
    server.starttls()
    server.login(sender_email, password)
    text = msg.as_string()
    server.sendmail(sender_email, recipient_email, text)
    server.quit()

    print("Email sent successfully to", recipient_email)



if __name__ == "__main__":

    day = sys.argv[1]
    recipient_email = sys.argv[2]
    
    # Call the function to update the spreadsheet
    worksheet = update_spreadsheet(day)

    # Call the function to export the spreadsheet to PDF
    pdf_file = export_spreadsheet_to_pdf(day)

    # Upload PDF to Google Drive
    upload_to_drive(pdf_file)

    # Send email with PDF attachment
    send_email(pdf_file, day, recipient_email)

