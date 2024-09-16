import fitz
import io
import os
import pdfplumber
import re
import streamlit as st
import subprocess
import tempfile
import win32com.client
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from docx2pdf import convert
from openai import OpenAI


# Initialize the OpenAI client
def initialize_openai_client(api_key):
    return OpenAI(api_key=api_key)


def extract_content_from_pdf(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text


def clean_text(text):
    text = re.sub(r"\s+", " ", text)
    text = "".join(char for char in text if char.isprintable() or char in ["\n", "\t"])
    return text.strip()


def wrap_keywords_in_b_tags(text, keywords):
    # Sort keywords by length in descending order to avoid partial replacements
    keywords = sorted(keywords, key=len, reverse=True)

    for keyword in keywords:
        # Use re.sub with a lambda to replace the matched keyword with <b>wrapped keyword</b>, ignoring case
        text = re.sub(
            f"(?i)({re.escape(keyword)})",
            lambda match: f"<b>{match.group(0)}</b>",
            text,
        )

    return text


def reformat_resume(content, keywords):
    Format = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Candidate Resume</title>
</head>
<body>

    <h1>Candidate Name</h1>
    <role_title>role title</role_title>

    <h2>PROFESSIONAL SUMMARY</h2>
    <ul>
        <li>professional summary 1</li>
        <li>professional summary 2</li>
        <li>professional summary 3</li>
        <!-- Add more summaries as needed -->
    </ul>


    <h2>TECHNICAL SKILLS</h2>
    <table border="1" cellpadding="5" cellspacing="0">
        <thead>
            <tr>
                <th>Category</th>
                <th>Tools & Technologies</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>Category 1</td>
                <td>Tools || Technologies</td>
            </tr>
            <tr>
                <td>Category 2</td>
                <td>Tools || Technologies</td>
            </tr>
            <tr>
                <td>Category 3</td>
                <td>Tools || Technologies</td>
            </tr>
            <!-- Add more rows as needed -->
        </tbody>
    </table>

    <h2>EDUCATION/QUALIFICATION</h2>
    <ul>
        <li>Education 1</li>
        <li>Education 2</li>
        <li>Education 3</li>
        <!-- Add more education qualifications as needed -->
    </ul>

    <h2>CERTIFICATION/TRAINING</h2>
    <ul>
        <li>Certification 1</li>
        <li>Certification 2</li>
        <li>Certification 3</li>
        <!-- Add more certifications as needed -->
    </ul>


    <!-- Strict instruction: Use the format MM/YY - MM/YY (Total Months) and calculate the total months or set 'Present' if applicable. If the candidate is currently working, ensure all responsibilities are written in the present tense. If the candidate is no longer working, ensure all responsibilities are written in the past tense. Double-check for consistent tense usage across all responsibilities. -->

    <h2>WORK HISTORY</h2>

    <strong>Date: MM/YY - MM/YY (Total Months)</strong> 
    <strong>Company: Company Name</strong>

        <p><strong>Client:</strong> Client Name</p>
        <p><strong>Title:</strong> Job Title</p>
        <p><strong>Tools and Technologies:</strong> Tools, Technologies</p>
        <p><strong>Description:</strong> Job description goes here</p>

        <p><strong>Roles and Responsibilities</strong></p> 
        <ol>
            <li>Responsibility 1</li> 
            <li>Responsibility 2</li> 
            <li>Responsibility 3</li> 
            <!-- Add more responsibilities as needed -->
        </ol>


    <!-- Repeat the WORK HISTORY section for each job -->

</body>
</html>
"""

    messages = [
        {
            "role": "system",
            "content": """
                                You are an expert in resume reformatting, specializing in adhering to predefined templates with high accuracy. Your role involves parsing and reformatting resumes according to a specific format, ensuring all details are captured correctly and presented in a clean, HTML-compatible format. You excel in grammatical precision, avoiding unnecessary details, and adding required prerequisites when needed.
                                """,
        },
        {
            "role": "user",
            "content": f"""
                Please parse the provided resume details according to the following format and deliver the output in HTML format:                           

                Instructions:
                    1. Professional Summary: Craft a concise summary that highlights the candidate's skills and achievements.
                    2. Technical Skills: List skills in a clear, tabular format.
                    3. Work History: Include dates, company, title, environment, job description, and responsibilities for each job. Provide details on projects related to each job.
                    4. Projects: For each project, include client, designation, environment, description, and responsibilities.
                        -While parsing the description, roles & responsibilities from the resume, ensure that all details are accurately captured. And also make sure that none of the details are missed out, no matter how small or long they are, you have to format each and every line of the resume as it is. You are not allowed to change the content of the resume, also don't summarize the content.
                        -Date: MM/YY - MM/YY (Total Months) and calculate the total months or set 'Present' if applicable. If the candidate is currently working, ensure all responsibilities are written in the present tense. If the candidate is no longer working, ensure all responsibilities are written in the past tense. Double-check for consistent tense usage across all responsibilities.
                    5. Education: List qualifications in a structured format.
                    6. Certification/Training: Include all relevant certifications and training.
                    7. Additionally I have a very critical requirement of grammatical precision, so you have to make sure that the content is grammatically correct and there are no tense errors in the content. For example, if the candidate is currently working, all responsibilities should be written in the present tense. If the candidate is no longer working, all responsibilities should be written in the past tense.
                        - Be sure to double-check for consistent tense usage across the whole resume.
                        - If the resume is not grammatically correct, you have to check the content and correct it as per the grammatical rules. Be very much careful about the existing content, you are not allowed to change the content, you can only correct the grammatical errors, as the output is used for official purpose directly to the clients.

                Formatting Rules:
                    - Use `<h2>` tags for section titles.
                    - Use `<p>` tags for text and descriptions.
                    - Use `<strong>` tags for any text that needs to be bold.
                    - Use bullet points for lists, wrapped in `<ul> or <ol>` and `<li>` tags.
                    - Align sections clearly and avoid adding unnecessary information.
                                
                Resume: [{content}]
                Format: [{Format}]

                            """,
        },
    ]

    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=messages,
        temperature=0.15,
        n=1,
    )

    reply = completion.choices[0].message.content

    if "```html" in reply and "```" in reply:
        reply = reply.replace("```html", "").replace("```", "").strip()
    if keywords:
        return wrap_keywords_in_b_tags(reply, keywords)
    return reply


# /////////////////////////// DOCX to HTML ///////////////////////////

# Constants
FONT_NAME = "Times New Roman"
LOGO_PATH = "kyralogo.png"
CONTACT_INFO = """3673 Coolidge Ct.,
Tallahassee, FL 32311
Phone: (850) 459-5854
Email: vpatel@KyraSolutions.com"""


def add_header_with_logo_and_contact(doc):
    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False

        # Create table with 1 row and 2 columns
        table = header.add_table(1, 2, width=Cm(20.32))
        table.alignment = WD_ALIGN_PARAGRAPH.LEFT
        table.columns[0].width = Cm(7.62)  # Adjust column width for the logo
        table.columns[1].width = Cm(12.7)  # Adjust column width for the contact info

        # Left cell for logo
        left_cell = table.cell(0, 0)
        left_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        if os.path.exists(LOGO_PATH):
            paragraph = left_cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(
                LOGO_PATH, width=Cm(4.67), height=Cm(2.3)
            )  # Adjust logo size
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            print(f"Logo file not found at {LOGO_PATH}")

        # Right cell for contact info
        right_cell = table.cell(0, 1)
        right_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        contact_paragraph = right_cell.paragraphs[0]
        contact_run = contact_paragraph.add_run(CONTACT_INFO)
        contact_run.font.size = Pt(10)
        contact_run.font.name = FONT_NAME
        contact_paragraph.alignment = (
            WD_ALIGN_PARAGRAPH.RIGHT
        )  # Align text to the right

    # Add a line break after the logo
    doc.add_paragraph()

    # Set margins
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.left_margin = Cm(1.27)
        section.right_margin = Cm(1.27)

    # Remove extra space before/after paragraphs
    doc.styles["Normal"].paragraph_format.space_before = Pt(0)
    doc.styles["Normal"].paragraph_format.space_after = Pt(0)


def add_paragraph(doc, text, style=None, bold=False, alignment=None, color=None):
    p = doc.add_paragraph(text, style=style)
    if alignment:
        p.alignment = alignment
    run = p.runs[0]
    run.bold = bold
    run.font.color.rgb = color if color else RGBColor(0, 0, 0)  # Set text color
    run.font.name = FONT_NAME  # Set font to Times New Roman
    run.font.size = Pt(10)  # Set font size to 10


def add_list_item(doc, element, indent):
    p = doc.add_paragraph(style="List Bullet")
    for child in element.children:
        handle_element(doc, child, p)
    p.paragraph_format.left_indent = Cm(indent)  # Adjust the indent as needed


def handle_element(doc, element, parent_paragraph=None):
    if isinstance(element, str):
        if parent_paragraph:
            run = parent_paragraph.add_run(element)
            run.font.name = FONT_NAME
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0, 0, 0)
        return

    # Handle <b> and <strong> tags for bold text
    if element.name in ["b", "strong"]:
        if parent_paragraph:
            run = parent_paragraph.add_run(element.get_text())
            run.bold = True
            run.font.name = FONT_NAME
            run.font.size = Pt(10)  # Set font size to 10
            run.font.color.rgb = RGBColor(0, 0, 0)
            if element.name == "strong":
                parent_paragraph.paragraph_format.left_indent = Cm(0.64)
                parent_paragraph.paragraph_format.space_before = Pt(12)
                parent_paragraph.paragraph_format.space_after = Pt(12)
        else:
            add_paragraph(doc, element.get_text(), bold=True)
        return

    if element.name == "h1":
        add_paragraph(
            doc, element.get_text(), bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER
        )
    elif element.name == "role_title":
        add_paragraph(
            doc,
            element.get_text(),
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            color=RGBColor(0, 0, 255),
        )
    elif element.name == "h2":
        p = doc.add_paragraph(element.get_text())
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_format = p.paragraph_format
        p_format.space_before = Pt(12)  # Add space before h2
        p_format.space_after = Pt(12)  # Add space after h2
        run = p.runs[0]
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)  # Set text color to black
        run.font.name = FONT_NAME  # Set font to Times New Roman
        run.font.size = Pt(10)  # Set font size to 10
    elif element.name == "h3":
        add_paragraph(doc, element.get_text(), style="Heading 2")
    elif element.name == "h4":
        add_paragraph(doc, element.get_text(), style="Heading 3")
    elif element.name == "p":
        p = doc.add_paragraph()
        for child in element.children:
            handle_element(doc, child, p)
    elif element.name == "table":
        # Create table with a specific number of columns (based on <th> tags)
        table = doc.add_table(
            rows=1, cols=len(element.find_all("th")), style="Table Grid"
        )

        # Disable autofit to allow manual control over the column widths
        table.autofit = False

        # Set the column widths (e.g., 1 inch for each column, adjust as needed)
        for column in table.columns:
            for cell in column.cells:
                cell.width = Cm(2.54)  # Adjust column width as needed

        # Set header row with custom font and size
        hdr_cells = table.rows[0].cells
        for idx, th in enumerate(element.find_all("th")):
            hdr_cells[idx].text = th.get_text()
            for paragraph in hdr_cells[idx].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)  # Set font size to 10pt
                    run.font.name = FONT_NAME  # Set font to Times New Roman

        # Add table rows and handle table data
        for tr in element.find_all("tr")[1:]:
            row_cells = table.add_row().cells
            for idx, td in enumerate(tr.find_all("td")):
                cell_paragraph = row_cells[idx].paragraphs[0]
                handle_element(doc, td, cell_paragraph)

                # Align paragraph and vertically center content in each cell
                row_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                row_cells[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Apply indentation to the table (e.g., 1 inch indent)
        tbl = table._tbl  # Access the underlying table element
        tblPr = tbl.tblPr  # Get the table properties
        tblInd = OxmlElement("w:tblInd")  # Create table indentation element
        tblInd.set(qn("w:w"), "500")  # Set indentation value in twips (1440 = 1 inch)
        tblInd.set(qn("w:type"), "dxa")  # Set measurement type to dxa (twips)

        # Append the indentation element to the table properties
        tblPr.append(tblInd)
    elif element.name in ["ul", "ol"]:
        indent = 1.27
        if element.name == "ol":
            indent = 2.12
        for li in element.find_all("li"):
            add_list_item(doc, li, indent)
        # Add space after the list
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(12)
    elif element.name == "br":
        doc.add_paragraph()
    else:
        for child in element.children:
            handle_element(doc, child, parent_paragraph)


def convert_html_to_docx(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    doc = Document()

    add_header_with_logo_and_contact(doc)

    for element in soup.body:
        handle_element(doc, element)

    return doc
    # doc.save(docx_filename)


# //////////////////////////////////////////////////////////////////////
def read_pdf(file_path):
    """Reads a .pdf file and returns its content as a string.

    Args:
        file_path (str): The path to the PDF file.

    Returns:
        str: The content of the PDF file as a string.
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")

    try:
        # Open the PDF file
        doc = fitz.open(file_path)
        content = []

        # Iterate through each page
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            content.append(page.get_text())

        return "\n".join(content)

    except Exception as e:
        raise RuntimeError(f"An error occurred while reading the PDF file: {e}")


def convert_doc_to_docx(doc_path, docx_path):
    """Converts a .doc file to .docx.

    Args:
        doc_path (str): The path to the .doc file.
        docx_path (str): The path to save the .docx file.
    """
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"File not found: {doc_path}")

    word = win32com.client.Dispatch("Word.Application")

    try:
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(docx_path, FileFormat=16)  # FileFormat=16 for .docx
        doc.Close()
    except Exception as e:
        raise RuntimeError(f"An error occurred while converting .doc to .docx: {e}")
    finally:
        word.Quit()


def convert_and_read(file_path):
    """Converts .doc or .docx files to .pdf and reads the .pdf content, or reads .pdf content directly.

    Args:
        file_path (str): The path to the input file (.doc, .docx, or .pdf).

    Returns:
        str: The content of the PDF file as a string.
    """
    # Handle .doc files by converting them to .docx
    if file_path.endswith(".doc"):
        docx_path = file_path.replace(".doc", ".docx")
        convert_doc_to_docx(file_path, docx_path)
        file_path = docx_path

    # Convert .docx files to .pdf
    if file_path.endswith(".docx"):
        pdf_path = file_path.replace(".docx", ".pdf")
        try:
            convert(file_path, pdf_path)
        except Exception as e:
            raise RuntimeError(f"An error occurred during the conversion: {e}")
        file_path = pdf_path

    # Handle .pdf files
    if file_path.endswith(".pdf"):
        return read_pdf(file_path)

    raise ValueError(
        "Unsupported file type. Please provide a .doc, .docx, or .pdf file."
    )


def handle_temp_file(uploaded_file, suffix):
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_file_path = tmp_file.name

    resume = convert_and_read(tmp_file_path)
    os.unlink(tmp_file_path)
    return resume


# Streamlit UI

st.title("Enhanced Resume Reformatter")
st.write(
    "Upload a resume in DOC,DOCX or PDF format to convert it to the predefined format."
)

api_key = st.text_input("Enter your OpenAI API key:", type="password")
keywords_input = st.text_input("Enter keywords (comma-separated)")
uploaded_file = st.file_uploader("Choose a file", type=["docx", "pdf"])
keywords = [keyword.strip() for keyword in keywords_input.split(",") if keyword.strip()]

if uploaded_file is not None:
    try:
        client = initialize_openai_client(api_key)
        file_type = uploaded_file.type

        if (
            file_type
            == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ):
            file_content = handle_temp_file(uploaded_file, ".docx")

        elif file_type == "application/msword":
            file_content = handle_temp_file(uploaded_file, ".doc")

        elif file_type == "application/pdf":
            file_content = uploaded_file.getvalue()
            st.write("PDF file uploaded successfully.")

        else:
            st.error("Unsupported file format. Please upload a PDF, DOC, or DOCX file.")
            st.stop()

        # Extract content from PDF
        resume_content = extract_content_from_pdf(io.BytesIO(file_content))
        cleaned_resume_content = clean_text(resume_content)

        with st.spinner("Reformatting resume..."):
            formatted_resume = reformat_resume(cleaned_resume_content, keywords)

        final_formatted_doc = convert_html_to_docx(formatted_resume)

        # Save as DOCX
        docx_buffer = io.BytesIO()
        final_formatted_doc.save(docx_buffer)
        docx_buffer.seek(0)

        st.download_button(
            label="Download Final Formatted Resume (DOCX)",
            data=docx_buffer,
            file_name="Final_Formatted_Resume.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        st.error("Please try uploading the file again or use a different file.")
