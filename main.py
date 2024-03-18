import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import os

def generate_reports(csv_file, template_file, reports_dir):
    # Load template
    doc = DocxTemplate(template_file)

    # Read CSV file
    data = pd.read_csv(csv_file)

    # Create a directory for saving reports on the user's desktop if it doesn't exist
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

        # Get the path to the user's desktop
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        reports_dir = os.path.join(desktop_path, 'reports')

        # Generate reports
        generate_reports(csv_file, template_file, reports_dir)
        st.success("Reports generated successfully!")
        st.write(f"Check the 'reports' directory on your desktop for the generated files")

if __name__ == "__main__":
    main()
