import streamlit as st
import zipfile
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Question Paper Formatter", layout="centered")
st.title("📝 Question Paper Formatter")

st.markdown("""
Upload multiple `.docx` files below and choose your formatting options. This tool supports English, Hindi, and Sanskrit documents.
""")

# Sidebar for formatting options
st.sidebar.header("⚙️ Formatting Options")
font_name = st.sidebar.selectbox("Font Name", ["Times New Roman", "Arial", "Calibri", "Mangal", "Kruti Dev"])
font_size = st.sidebar.slider("Font Size (pt)", 10, 18, 12)
line_spacing = st.sidebar.selectbox("Line Spacing", [1.0, 1.15, 1.5, 2.0])
margin_inch = st.sidebar.slider("Margins (inches)", 0.5, 2.0, 1.0)
bold_sections = st.sidebar.checkbox("Bold Section Headers", value=True)
auto_indent = st.sidebar.checkbox("Auto Indent Paragraphs", value=True)

add_header = st.sidebar.checkbox("Add Header/Footer", value=True)
school_name = st.sidebar.text_input("School Name", "Your School Name")
exam_name = st.sidebar.text_input("Exam Name", "Term 1 Examination")

uploaded_files = st.file_uploader("📂 Upload DOCX Files", type="docx", accept_multiple_files=True)

# Utility to detect section headers
def is_section_header(text):
    keywords = ["section", "instructions", "general", "note"]
    return any(text.lower().strip().startswith(k) for k in keywords)

# Align marks to right
def align_marks_right(paragraph):
    if "(" in paragraph.text and ")" in paragraph.text:
        text = paragraph.text
        parts = text.rsplit("(", 1)
        if len(parts) == 2 and ")" in parts[1]:
            question, marks = parts[0].strip(), "(" + parts[1].strip()
            paragraph.clear()
            run_q = paragraph.add_run(question + " ")
            run_m = paragraph.add_run(marks)

            # Font settings
            for run in [run_q, run_m]:
                run.font.name = font_name
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                run.font.size = Pt(font_size)

            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# Main formatter
def format_docx(file, settings):
    doc = Document(file)

    for section in doc.sections:
        section.top_margin = section.bottom_margin = Inches(settings['margin'])
        section.left_margin = section.right_margin = Inches(settings['margin'])

        if settings['add_header']:
            section.header.paragraphs[0].text = f"{settings['school']} | {settings['exam']}"
            footer = section.footer.paragraphs[0]
            footer.clear()
            footer.add_run("Page ").bold = True
            footer.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for para in doc.paragraphs:
        para.paragraph_format.line_spacing = settings['spacing']
        if settings['indent']:
            para.paragraph_format.left_indent = Inches(0.25)

        if settings['bold_section'] and is_section_header(para.text):
            for run in para.runs:
                run.bold = True

        for run in para.runs:
            run.font.name = settings['font']
            run._element.rPr.rFonts.set(qn('w:eastAsia'), settings['font'])
            run.font.size = Pt(settings['size'])

        align_marks_right(para)

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
            'exam': exam_name,
            'bold_section': bold_sections,
            'indent': auto_indent
        })

        out_stream = BytesIO()
        formatted_doc.save(out_stream)
        out_stream.seek(0)
        formatted_files.append((uploaded_file.name, out_stream))

    if formatted_files:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for name, file_data in formatted_files:
                zip_file.writestr(name, file_data.read())
        zip_buffer.seek(0)

        st.success("✅ All files formatted successfully!")
        st.download_button("📦 Download All as ZIP", zip_buffer, file_name="Formatted_Questions.zip")
