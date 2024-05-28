import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

def extract_text_from_pdf(pdf,start_string,end_string):
    text = ""
    with pdfplumber.open(pdf) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

    start_index = text.find(start_string)

    # Find the start index of "Total"
    end_index = text.find(end_string)

    final_text = text[start_index:end_index].strip()

    return final_text

def normalize_header(header):
    return header

def parse_text_to_structure(text):
    raw_headers = ["Invoice Dt", "Invoice No.", "Description", "Amt. Paid", "WH Tax", "Total Invoice Amt."]
    headers = {normalize_header(header): header for header in raw_headers}
    data = {key: [] for key in headers}

    lines = text.split('\n')

    current_header = None
    # for line in lines:
    #     normalized_line = normalize_header(line)
    #     if normalized_line in headers:
    #         current_header = normalized_line
    #         st.write(f"Detected header: {headers[current_header]}")
    #     elif current_header:
    #         data[current_header].append(line.strip())

    # structured_data = []
    # max_length = max(len(v) for v in data.values())
    # for i in range(max_length):
    #     row = {headers[header]: (data[header][i] if i < len(data[header]) else "") for header in data}
    #     structured_data.append(row)

    # return structured_data

    structured_data = []

    # for line in lines:
    for index, line in enumerate(lines):
        if index!=0:
            normalized_line = normalize_header(line)
            line_words = normalized_line.split()
            structured_data.append(line_words)

    return structured_data

def save_data_to_excel(structured_data):
    df = pd.DataFrame(structured_data)
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()  # Correctly close the writer to save the file
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title("PDF Table Extractor")

    start_substring = st.text_input("Start Substring", "InvoiceDt")
    end_substring = st.text_input("End Substring", "Total (USD)")

    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

    if uploaded_file is not None:
        text = extract_text_from_pdf(uploaded_file,start_substring,end_substring)

        if not text:
            st.error("Failed to extract text from the PDF.")
            return

        st.write("Extracted Text:")
        st.write(text)

        structured_data = parse_text_to_structure(text)

        if not structured_data:
            st.error("No data found in the PDF")
            return

        st.write("Parsed Data:")
        st.write(pd.DataFrame(structured_data))

        excel_data = save_data_to_excel(structured_data)

        st.download_button(label='Download Excel file',
                           data=excel_data,
                           file_name='extracted_data.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    main()
