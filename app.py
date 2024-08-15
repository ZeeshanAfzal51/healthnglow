import fitz  # PyMuPDF
from pdf2image import convert_from_bytes
import pytesseract
from PIL import Image
from openpyxl import load_workbook
import google.generativeai as genai
import streamlit as st
import os
import pandas as pd

# Set the API key for Google Generative AI
os.environ["GEMINI_API_KEY"] = "AIzaSyDI2DelJZlGyXEPG3_b-Szo-ixRvaB0ydY"
genai.configure(api_key=os.environ["GEMINI_API_KEY"])

# Title
st.title("Invoice PDF Processor")

# File Upload: Invoice PDFs
st.markdown("**Upload the Invoice PDFs**")
pdf_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)

# File Upload: Local Master Excel File
st.markdown("**Upload the Local Master Excel File**")
excel_file = st.file_uploader("Choose Excel file", type="xlsx")

if pdf_files and excel_file:
    # Load the workbook and select the active sheet
    workbook = load_workbook(excel_file)
    worksheet = workbook.active

    # Insert headers in the first row
    headers = ["PO Number", "Invoice Number", "Invoice Amount", "Invoice Date", "CGST Amount", 
               "SGST Amount", "IGST Amount", "Total Tax Amount", "Taxable Amount", "TCS Amount", 
               "IRN Number", "Receiver GSTIN", "Receiver Name", "Vendor GSTIN", "Vendor Name", 
               "Remarks", "Vendor Code"]
    worksheet.append(headers)

    # Define the prompt
    prompt = ("Before we begin please disregard, ignore, and forget any previous context or cumulative data. "
              "The following is OCR extracted text from a single invoice PDF. "
              "Please use the OCR extracted text to give a structured summary. "
              "The structured summary should consider information such as PO Number, Invoice Number, Invoice Amount, "
              "Invoice Date, CGST Amount, SGST Amount, IGST Amount, Total Tax Amount, Taxable Amount, TCS Amount, "
              "IRN Number, Receiver GSTIN, Receiver Name, Vendor GSTIN, Vendor Name, Remarks, and Vendor Code. "
              "If any of this information is not available or present, then NA must be denoted next to the value. "
              "Please do not give any additional information.")

    def extract_text_from_pdf(pdf_stream):
        doc = fitz.open(stream=pdf_stream.read(), filetype="pdf")
        text_data = []
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text("text")
            text_data.append(text)
        return text_data

    def convert_pdf_to_images_and_ocr(pdf_stream):
        pdf_stream.seek(0)
        images = convert_from_bytes(pdf_stream.read())  # Use convert_from_bytes
        ocr_results = [pytesseract.image_to_string(image) for image in images]
        return ocr_results

    def combine_text_and_ocr_results(text_data, ocr_results):
        combined_results = []
        for text, ocr_text in zip(text_data, ocr_results):
            combined_results.append(text + "\n" + ocr_text)
        combined_text = "\n".join(combined_results)
        return combined_text

    def extract_parameters_from_response(response_text):
        parameters = {
            "PO Number": "NA",
            "Invoice Number": "NA",
            "Invoice Amount": "NA",
            "Invoice Date": "NA",
            "CGST Amount": "NA",
            "SGST Amount": "NA",
            "IGST Amount": "NA",
            "Total Tax Amount": "NA",
            "Taxable Amount": "NA",
            "TCS Amount": "NA",
            "IRN Number": "NA",
            "Receiver GSTIN": "NA",
            "Receiver Name": "NA",
            "Vendor GSTIN": "NA",
            "Vendor Name": "NA",
            "Remarks": "NA",
            "Vendor Code": "NA"
        }
        
        lines = response_text.splitlines()
        for line in lines:
            for key in parameters.keys():
                if key in line:
                    value = line.split(":")[-1].strip()
                    # Remove surrounding double quotes and trailing commas
                    value = value.strip('"').strip(',')
                    # Remove trailing double quotes
                    value = value.rstrip('"')
                    parameters[key] = value

        return parameters
    
    all_parameters = []

    for pdf_file in pdf_files:
        text_data = extract_text_from_pdf(pdf_file)
        ocr_results = convert_pdf_to_images_and_ocr(pdf_file)
        combined_text = combine_text_and_ocr_results(text_data, ocr_results)
        
        # Combine the prompt and the extracted text
        input_text = f"{prompt}\n\n{combined_text}"

        # Creating the model configuration
        generation_config = {
            "temperature": 1,
            "top_p": 0.95,
            "top_k": 64,
            "max_output_tokens": 8192,
            "response_mime_type": "text/plain",
        }

        # Initializing the model
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            generation_config=generation_config,
        )

        # Starting a chat session
        chat_session = model.start_chat(
            history=[]
        )

        # Send the combined text as a message
        response = chat_session.send_message(input_text)

        # Extract the relevant data from the response
        parameters = extract_parameters_from_response(response.text)
        all_parameters.append(parameters)

        # Add data to the Excel file
        row_data = [parameters[key] for key in parameters.keys()]
        worksheet.append(row_data)

        # Display success message
        st.markdown(f"**Data from {pdf_file.name} has been successfully added to the Excel file**")
    
    # Save the updated Excel file with the same name as the uploaded file
    updated_excel_file_name = excel_file.name
    workbook.save(updated_excel_file_name)

    # Provide download link for the updated Excel file
    with open(updated_excel_file_name, "rb") as file:
        st.download_button(
            label="Download Updated Excel File",
            data=file,
            file_name=updated_excel_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
