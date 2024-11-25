import os
import json
import openai
import pandas as pd
import streamlit as st
from PyPDF2 import PdfReader
from PIL import Image
import pytesseract
from io import BytesIO
from dotenv import load_dotenv
load_dotenv()

# Set OpenAI API key
openai.api_key = os.environ.get("API_KEY")


# Function to extract text from uploaded files
def extract_text_from_file(file):
    file_type = file.type
    try:
        if file_type == "application/pdf":
            pdf_reader = PdfReader(file)
            return " ".join([page.extract_text() for page in pdf_reader.pages])
        elif file_type.startswith("image/"):
            image = Image.open(file)

            return pytesseract.image_to_string(image)
        elif file_type in ["text/plain", "application/json"]:
            return file.read().decode("utf-8")
        else:
            st.error("Unsupported file type. Please upload a PDF, image, or text file.")
            return None
    except Exception as e:
        st.error(f"Error extracting text: {e}")
        return None
def process_json(data, parent_key=""):
    """
    Flatten JSON recursively and return a dictionary where keys represent columns
    and values represent rows, suitable for conversion to a DataFrame.
    """
    flattened_data = {}

    def flatten(item, key_prefix=""):
        if isinstance(item, dict):
            for key, value in item.items():
                flatten(value, f"{key_prefix}{key}." if key_prefix else f"{key}.")
        elif isinstance(item, list):
            for i, sub_item in enumerate(item):
                flatten(sub_item, f"{key_prefix}{i}.")
        else:
            flattened_data[key_prefix[:-1]] = item

    flatten(data)
    return flattened_data

# Function to convert JSON to Excel
def convert_json_to_excel(json_data):
    """
    Converts JSON data into an Excel file with well-structured columns and rows.
    """
    try:
        data = json.loads(json_data)
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Process the main JSON into a DataFrame
            if isinstance(data, dict):
                flattened_dict = process_json(data)
                df = pd.DataFrame([flattened_dict])  # Single-row DataFrame
            elif isinstance(data, list):
                df = pd.DataFrame([process_json(item) for item in data])
            else:
                raise ValueError("Unsupported JSON format: Root must be an object or array.")

            # Write the main data to the first sheet
            df.to_excel(writer, sheet_name="Main Data", index=False)

        output.seek(0)
        return output

    except Exception as e:
        st.error(f"Error converting JSON to Excel: {e}")
        return None

# Function to validate JSON
def validate_json(json_data):
    try:
        return json.loads(json_data)
    except json.JSONDecodeError:
        st.error("Invalid JSON response. Please check the input.")
        return None

# Function to generate structured JSON using OpenAI
def generate_json_response(prompt):
    full_prompt = f"""
    You are a document classification and entity extraction assistant. Your task is to process input text, classify the type of document, and extract specific information based on the classification. Follow these instructions strictly and also do not limited to below mentioned entities, give all possible necessary enities:

    1. **Document Classification**:
    - First, identify the type of document. Possible types include:
        - **Bill/Invoice**
        - **Identity Document**
        - **Result/Grade Sheet**
        - **Other** (if none of the above apply)

    2. **Entity Extraction**:
    - Based on the identified document type, extract only relevant entities.
    - The response must be a flat JSON structure with no nested objects or arrays.

    **For Bill/Invoice**:
    - `document_type`: "Bill/Invoice"
    - `sender`: Name of the organization or person issuing the bill
    - `receiver`: Name of the recipient
    - `invoice_number`: Unique identifier for the invoice
    - `invoice_date`: Date of the invoice
    - `due_date`: Payment due date
    - `total_amount`: Total amount on the invoice
    - `currency`: Currency of the transaction
    - `billing_address`: Address where the bill is issued
    - `shipping_address`: Address for delivery
    - `payment_method`: Method of payment (if available)

    **For Identity Document**:
    - `document_type`: "Identity Document"
    - `name`: Name of the person
    - `id_number`: Unique identification number
    - `date_of_birth`: Date of birth
    - `issue_date`: Date of issue
    - `expiry_date`: Expiry date
    - `address`: Address of the person (if available)
    - `nationality`: Nationality (if mentioned)

    **For Result/Grade Sheet**:
    - `document_type`: "Result/Grade Sheet"
    - `student_name`: Name of the student
    - `roll_number`: Unique identifier for the student
    - `exam_name`: Name of the exam
    - `date_of_issue`: Date when the result was issued
    - `subjects`: Comma-separated list of subjects
    - `grades`: Comma-separated list of grades corresponding to the subjects
    - `overall_result`: Pass/Fail or other summary

    **For Other Documents**:
    - `document_type`: "Other"
    - Extract any relevant details such as `document_title`, `issue_date`, or any specific identifiers if available.

    3. **Formatting Requirements**:
    - The response must be valid JSON without ```json only with braces.
    - Do not use nested structures or arrays.
    - Ensure that all keys are lowercase and use underscores as separators.
    - Do not only limited to above mentioned keys gather max information you can
    - Make sure keys are in english



    Input Text:
    {prompt}
    



    """
    
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a helpful assistant extracting entities from text to JSON."},
                {"role": "user", "content": full_prompt}
            ],
            max_tokens=2000,
            temperature=0.5
        )
        print(response['choices'][0]['message']['content'])
        return response['choices'][0]['message']['content']
    except Exception as e:
        st.error(f"Error generating JSON: {e}")
        return None

# Streamlit App UI
st.title("Document Processor")
st.write("Upload a file (PDF, image, or text) to extract structured data as Excel.")

uploaded_file = st.file_uploader("Upload a file", type=["pdf", "txt", "png", "jpg", "jpeg"])

if uploaded_file:
    st.write("Processing file...")
    extracted_text = extract_text_from_file(uploaded_file)
    
    if extracted_text:
        # st.text_area("Extracted Text", extracted_text, height=200)
        pass
        
        if st.button("Generate Excel"):
            st.write("Extracting text and generating excel file...")
            generated_json = generate_json_response(extracted_text)
            
            if generated_json:
                validated_json = validate_json(generated_json)
                if validated_json:
                    # st.json(validated_json)
                    pass

                    excel_file = convert_json_to_excel(generated_json)
                    if excel_file:
                        st.download_button(
                            label="Download Excel",
                            data=excel_file,
                            file_name="output.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
