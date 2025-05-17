import streamlit as st
import os
import zipfile
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn

st.set_page_config(page_title="Question Paper Formatter", layout="centered")
st.title("📝 Question Paper Formatter")

st.markdown("""
Upload multiple `.docx` files below and choose your formatting options. This tool supports English, Hindi, and Sanskrit documents.
""")

# Sidebar for formatting options
st.sidebar.header("⚙️ Formatting Options")
font_name = st.sidebar.selectbox("Font Name", ["Times New Roman", "Arial", "Mangal", "Kruti Dev"])
font_size = st.sidebar.slider("Font Size (pt)", 10, 18, 12)
line_spacing = st.sidebar.selectbox("Line Spacing", [1.0, 1.15, 1.5, 2.0])
margin_inch = st.sidebar.slider("Margins (inches)", 0.5, 2.0, 1.0)

add_header = st.sidebar.checkbox("Add Header/Footer", value=True)
school_name = st.sidebar.text_input("School Name", "Your School Name")
exam_name = st.sidebar.text_input("Exam Name", "Term 1 Examination")

uploaded_files = st.file_uploader("📂 Upload DOCX Files", type="docx", accept_multiple_files=True)

# Function to format a single docx file
def format_docx(file, settings):
    doc = Document(file)

    # Page margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(settings['margin'])
        section.bottom_margin = Inches(settings['margin'])
        section.left_margin = Inches(settings['margin'])
        section.right_margin = Inches(settings['margin'])

    # Header/Footer
    if settings['add_header']:
        for section in sections:
            header = section.header
            footer = section.footer
            header.paragraphs[0].text = settings['school'] + " | " + settings['exam']
            footer.paragraphs[0].text = "Page "

    # Apply font and spacing to all paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            run.font.name = settings['font']
            run._element.rPr.rFonts.set(qn('w:eastAsia'), settings['font'])
            run.font.size = Pt(settings['size'])
        para.paragraph_format.line_spacing = settings['spacing']

    return doc

if uploaded_files:
    formatted_files = []
    for uploaded_file in uploaded_files:
        formatted_doc = format_docx(uploaded_file, {
            'font': font_name,
            'size': font_size,
            'spacing': line_spacing,
            'margin': margin_inch,
            'add_header': add_header,
            'school': school_name,
            'exam': exam_name
        })

        out_stream = BytesIO()
        formatted_doc.save(out_stream)
        out_stream.seek(0)
        formatted_files.append((uploaded_file.name, out_stream))

    if formatted_files:
        # Create ZIP
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for name, file_data in formatted_files:
                zip_file.writestr(name, file_data.read())
        zip_buffer.seek(0)

        st.success("✅ All files formatted successfully!")
        st.download_button("📦 Download All as ZIP", zip_buffer, file_name="Formatted_Questions.zip")
