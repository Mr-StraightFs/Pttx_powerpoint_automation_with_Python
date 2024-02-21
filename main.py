# Streamlit app
import streamlit as st
import pandas as pd
import pptx
from datetime import datetime
from pathlib import Path
import base64
from io import BytesIO
from constants import INPUT_TEMPLATE_PATH, COLUMNS_TO_CONTROL_NOMEMCLATURE, COLUMNS_TO_CONTROL_ODOO, COMMON_COLUMN_NAME
from presentation_generator import generate_presentation_from_template
from data_preparation import merge_and_clean_dataframes, control_columns_in_input



# Function to check if the uploaded file is in CSV or Excel format
def is_valid_file_type(file):
    return file.type == "application/vnd.ms-excel" or file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Helper functions for handling file downloads
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df, filename, text):
    val = to_excel(df)
    b64 = base64.b64encode(val)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}">{text}</a>'

def get_pptx_download_link(pptx_data, filename, text):
    b64 = base64.b64encode(pptx_data).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}">{text}</a>'

# Function to save the presentation as a .pptx file
def save_presentation_as_pptx(prs, filename):
    prs.save(filename)

def save_presentation(prs):
    with BytesIO() as output:
        prs.save(output)
        data = output.getvalue()
    return data

# Main Streamlit app
def main():
    st.set_page_config(page_title="Who's Who Document Generator", page_icon="✍️", layout="centered", initial_sidebar_state="expanded")

    st.markdown(
        """
        <style>
        .stButton > button:first-child {
            background-color: #008CBA;
            color: white;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("Who's Who Document Generator")
    st.write("Welcome to the Who's Who Document Generator. This tool allows you to create a presentation based on input data. Please follow the steps below:")

    st.header("Step 1: Upload Job Nomenclature")
    nomenclature_file = st.file_uploader("Upload the Job Nomenclature Excel file (CSV or Excel)", type=["xlsx", "csv"])

    if nomenclature_file:
        try:
            nomenclature_df = pd.read_excel(nomenclature_file) if nomenclature_file.name.endswith("xlsx") else pd.read_csv(nomenclature_file)
            control_columns_in_input(nomenclature_df, COLUMNS_TO_CONTROL_NOMEMCLATURE)
            st.success("Job Nomenclature file uploaded successfully!")
        except Exception as e:
            st.error(f"Error: {e}")

    st.header("Step 2: Upload Odoo Extraction")
    odoo_extraction_file = st.file_uploader("Upload the Odoo Extraction Excel file (CSV or Excel)", type=["xlsx", "csv"])

    if odoo_extraction_file:
        try:
            odoo_extraction_df = pd.read_excel(odoo_extraction_file) if odoo_extraction_file.name.endswith("xlsx") else pd.read_csv(odoo_extraction_file)
            control_columns_in_input(odoo_extraction_df, COLUMNS_TO_CONTROL_ODOO)
            st.success("Odoo Extraction file uploaded successfully!")
        except Exception as e:
            st.error(f"Error: {e}")

    st.header("Step 3: Generate Presentation")
    if st.button("Generate Presentation"):
        if 'nomenclature_df' in locals() and 'odoo_extraction_df' in locals():
            common_column_name = COMMON_COLUMN_NAME
            cleaned_merged_df = merge_and_clean_dataframes(odoo_extraction_df, nomenclature_df, common_column_name)
            cleaned_merged_df.to_excel('cleaned_merged_df.xlsx',)

            prs = pptx.Presentation(INPUT_TEMPLATE_PATH)

            # Check if the presentation contains slides
            if len(prs.slides) > 0:
                # Get today's date
                today = datetime.now().strftime('%m-%d-%Y')

                prs_2 = generate_presentation_from_template(prs, cleaned_merged_df)

                # Get today's date
                today = datetime.now().strftime('%m-%d-%Y')
                # Update the file path with today's date
                file_path = f'./output_data/auto_generated_whoswho_{today}.pptx'
                prs_2.save(file_path)

                # Provide a download link for the presentation
                ppt_data = open(file_path, 'rb').read()
                st.markdown(get_pptx_download_link(ppt_data, "presentation.pptx", "Download Presentation"),
                            unsafe_allow_html=True)
                st.success("Presentation generated successfully!")
            else:
                st.error("No slides found in the presentation. Check your data and template.")



        else:
            st.warning("Please upload both Job Nomenclature and Odoo Extraction files before generating the presentation.")

if __name__ == "__main__":
    main()
