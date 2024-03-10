import hashlib
import streamlit as st
import spacy
import camelot
from docx import Document
import os

# Check if model is already present
if not os.path.exists("path/to/your/model/directory/en_core_web_sm"):
    try:
        spacy.cli.download("en_core_web_sm")
        print("Model downloaded successfully.")
    except Exception as e: 
        print(f"Error downloading model: {e}")

# Replace with the path to your spaCy model within Colab or consider downloading the model.
SPACY_MODEL = "en_core_web_sm"

@st.cache(hash_funcs={type(lambda: 0): lambda _: None})  # Caching to avoid reloading the model
def load_nlp_model():
    return spacy.load(SPACY_MODEL)

nlp = load_nlp_model()

def extract_pdf_premiums(pdf_file):
    # Read tables from PDF
    tables = camelot.read_pdf(pdf_file, pages="all")

    # Assuming your premium table is the first 
    premium_table = tables[0].df  # Access table as a DataFrame

    # Extract headers (assuming they are in the first row)
    headers = premium_table.iloc[0].tolist()

    # Extract data row (assuming it's the second row)
    data_row = premium_table.iloc[1].tolist()

    # Create a dictionary mapping headers to corresponding data
    premiums = dict(zip(headers, data_row))

    return premiums

def update_docx_table(docx_file, premiums):
    doc = Document(docx_file)

    # Find your target table (adjust search criteria as needed)
    for table in doc.tables:
        if "Plan" in table.rows[0].cells[0].text and "Premium" in table.rows[0].cells[1].text:
            target_table = table
            break

    # Update target table cells based on matching headers
    for i, cell in enumerate(target_table.cell(0).cells):  # Assuming headers are in the first row
        header = cell.text.strip().lower()
        if header in premiums:
            target_table.cell(1, i).text = premiums[header]  # Assuming data in the second row

    # Temporarily saving the DOCX (consider a download option for the user)
    doc.save("updated_premium_summary.docx")

def main():
    st.title("Document Processing App")

    pdf_file = st.file_uploader("Upload PDF", type=["pdf"])
    docx_file = st.file_uploader("Upload DOCX", type=["docx"])

    if pdf_file and docx_file:
        # Process the files
        premiums = extract_pdf_premiums(pdf_file)
        update_docx_table(docx_file, premiums)

        st.success("DOCX file updated! (updated_premium_summary.docx)")

if __name__ == "__main__":
    main()
