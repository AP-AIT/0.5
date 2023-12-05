import streamlit as st
import imaplib
import email
from bs4 import BeautifulSoup
import re
import pandas as pd
from io import BytesIO
import base64
import PyPDF2
import pytesseract
from PIL import Image

# Streamlit app title
st.header("Developed by MKSSS-AIT")
st.title("Automate2Excel: Simplified Data Transfer")

# Create input fields for the user and password
user = st.text_input("Enter your email address")
password = st.text_input("Enter your email password", type="password")

# Create input field for the email address to search
search_email = st.text_input("Enter the email address to search", value="info@uidanceexpo.com")

# Function to extract information from HTML content
def extract_info_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    name_element = soup.find(string=re.compile(r'Name', re.IGNORECASE))
    email_element = soup.find(string=re.compile(r'Email', re.IGNORECASE))
    workshop_element = soup.find(string=re.compile(r'Workshop Detail', re.IGNORECASE))
    date_element = soup.find(string=re.compile(r'Date', re.IGNORECASE))
    mobile_element = soup.find(string=re.compile(r'Mobile No\.', re.IGNORECASE))

    info = {
        "Name": None,
        "Email": None,
        "Workshop Detail": None,
        "Date": None,
        "Mobile No.": None
    }

    if name_element and name_element.find_next('td'):
        info["Name"] = name_element.find_next('td').get_text().strip()

    if email_element and email_element.find_next('td'):
        info["Email"] = email_element.find_next('td').get_text().strip()

    if workshop_element and workshop_element.find_next('td'):
        info["Workshop Detail"] = workshop_element.find_next('td').get_text().strip()

    if date_element and date_element.find_next('td'):
        info["Date"] = date_element.find_next('td').get_text().strip()

    if mobile_element and mobile_element.find_next('td'):
        info["Mobile No."] = mobile_element.find_next('td').get_text().strip()

    return info


# Function to extract text from PDF file
def extract_text_from_pdf(pdf_content):
    pdf_reader = PyPDF2.PdfFileReader(BytesIO(pdf_content))
    text = ""
    for page_num in range(pdf_reader.numPages):
        text += pdf_reader.getPage(page_num).extractText()
    return text

# Function to extract text from image
def extract_text_from_image(image_content):
    image = Image.open(BytesIO(image_content))
    text = pytesseract.image_to_string(image)
    return text

if st.button("Fetch"):
    try:
        # URL for IMAP connection
        imap_url = 'imap.gmail.com'

        # Connection with GMAIL using SSL
        my_mail = imaplib.IMAP4_SSL(imap_url)

        # Log in using user and password
        my_mail.login(user, password)

        # Select the Inbox to fetch messages
        my_mail.select('inbox')

        # Define the key and value for email search
        key = 'FROM'
        value = search_email
        _, data = my_mail.search(None, key, value)

        mail_id_list = data[0].split()

        info_list = []

        # Iterate through messages and extract information from HTML content, PDFs, and images
        for num in mail_id_list:
            typ, data = my_mail.fetch(num, '(RFC822)')
            msg = email.message_from_bytes(data[0][1])

            # Extract subject
            subject = msg["Subject"]

            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    html_content = part.get_payload(decode=True).decode('utf-8')
                    info = extract_info_from_html(html_content)

                elif part.get_content_type() == 'application/pdf':
                    pdf_content = part.get_payload(decode=True)
                    text = extract_text_from_pdf(pdf_content)
                    info["PDF Content"] = text

                elif part.get_content_type().startswith('image/'):
                    image_content = part.get_payload(decode=True)
                    text = extract_text_from_image(image_content)
                    info["Image Content"] = text

            # Extract and add the received date and subject
            date = msg["Date"]
            info["Received Date"] = date
            info["Subject"] = subject

            info_list.append(info)

        # Create a DataFrame from the info_list
        df = pd.DataFrame(info_list)

        # Display the data in the Streamlit app
        st.write("Data extracted from emails:")
        st.write(df)

        # Download the DataFrame as an Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        output.seek(0)
        st.write("Downloading Excel file...")
        st.download_button(
            label="Download Excel File",
            data=output,
            key="download_excel",
            on_click=None,
            file_name="EXPO_leads.xlsx"  # Specify the file name
        )

    except Exception as e:
        st.error(f"Error: {e}")
