import os
import docx
import pdfkit

def convert_docx_to_html(docx_path):
    doc = docx.Document(docx_path)
    html = ""
    for paragraph in doc.paragraphs:
        html += f"<p>{paragraph.text}</p>"
    return html

def convert_html_to_pdf(html_content, pdf_path):
    pdfkit.from_string(html_content, pdf_path)

def iterate_and_convert(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            docx_path = os.path.join(folder_path, filename)
            pdf_path = os.path.join(folder_path, filename.replace(".docx", ".pdf"))
            html_content = convert_docx_to_html(docx_path)
            convert_html_to_pdf(html_content, pdf_path)
            print(f"Converted {docx_path} to {pdf_path}")

if __name__ == "__main__":
    folder_path = input("Enter the path to the folder containing .docx files: ")
    iterate_and_convert(folder_path)

