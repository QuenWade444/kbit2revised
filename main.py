import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import os
import shutil
from zipfile import ZipFile
import base64

def generate_reports(csv_file, template_file, reports_dir):
    # Load template
    doc = DocxTemplate(template_file)

    # Read CSV file
    data = pd.read_csv(csv_file)

    # Create a directory for saving reports
    os.makedirs(reports_dir, exist_ok=True)

    # Iterate through each row in the CSV
    for index, row in data.iterrows():
        # Render template with row data
        context = {col: row[col] for col in data.columns}
        doc.render(context)

        # Save the rendered report to the reports directory
        report_filename = os.path.join(reports_dir, f"report_{index + 1}.docx")
        doc.save(report_filename)

def main():
    st.title('CSV to Reports Generator')

    # Select CSV file
    csv_file = st.file_uploader("Select CSV file:", type=['csv'])

    if csv_file is not None:
        # Define the path to the template file
        template_file = "template.docx"

        # Define the directory for saving reports
        reports_dir = "reports"

        # Generate reports
        generate_reports(csv_file, template_file, reports_dir)
        st.success("Reports generated successfully!")

        # Zip the reports directory
        zip_file_path = "reports.zip"
        with ZipFile(zip_file_path, 'w') as zipf:
            for root, dirs, files in os.walk(reports_dir):
                for file in files:
                    zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), reports_dir))

        # Provide download link for the zip file
        st.write("Download the zip file containing the reports:")
        st.markdown(f"[Download Reports](data:application/zip;base64,{base64.b64encode(open(zip_file_path, 'rb').read()).decode()})")

if __name__ == "__main__":
    main()
