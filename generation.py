import openpyxl
from docxtpl import DocxTemplate
from docx2pdf import convert
import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont

def generate_certificates(excel_file, template_file, output_folder, font_path="arialbd.ttf", font_size=100):
    data = pd.read_excel(excel_file)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    font_name = ImageFont.truetype(font_path, font_size)
    for index, row in data.iterrows():
        name = row["Name"]
        certificate = Image.open(template_file)
        draw = ImageDraw.Draw(certificate)
        if len(name) >= 15 and len(name) < 25:
            name_position = (550, 600)
        elif len(name) >= 10 and len(name) < 15:
            name_position = (700, 600)
        else:
            name_position = (730, 600)
        draw.text(name_position, name, fill="orange", font=font_name)
        output_path = os.path.join(output_folder, "certificate_" + name + ".png")
        certificate.save(output_path)
        print("Certificate generated for " + name + " and saved to " + output_path)
    print("All certificates have been generated!")
# Example usage
generate_certificates(
    excel_file="Certificate.xlsx",
    template_file="certificate.png",
    output_folder="Generate_certificate"
)


def load_student_data(file_name):
    """Load student data from an Excel file."""
    workbook = openpyxl.load_workbook(file_name)
    sheet = workbook.active
    return list(sheet.values)
def create_directories(*dirs):
    """Create directories if they don't exist."""
    for directory in dirs:
        os.makedirs(directory, exist_ok=True)
def generate_documents(template_path, student_list, doc_dir, pdf_dir):
    student_info = DocxTemplate(template_path)
    for student in student_list[1:]:  # Skip the header row
        context = {
            'name_kh': student[2],
            'g1': student[4],
            'id_kh': student[0],
            'name_e': student[3],
            'g2': student[5],
            'id_e': student[1],
            'dob_kh': student[6],
            'pro_kh': student[8],
            'dob_e': student[7],
            'pro_e': student[9],
            'ed_kh': student[10],
            'ed_e': student[11],
            # 'cur_date': student[12]
        }        
        # Render and save the .docx file
        doc_name = os.path.join(doc_dir, f"{student[1]}.docx")
        student_info.render(context)
        student_info.save(doc_name)
        print(f"{doc_name} has been created.")      
        # Convert the .docx file to PDF
        pdf_name = os.path.join(pdf_dir, f"{student[1]}.pdf")
        convert(doc_name, pdf_name)
        print(f"{pdf_name} has been created.")
# Main logic directly executed
excel_file = 'Template associate degree - 2024.xlsx'
template_file = 'WEP Temporary Certificate - Template.docx'
pdf_output_dir = "pdf_outputs"
doc_output_dir = "doc_name"
# Load data and create directories
student_list = load_student_data(excel_file)
create_directories(pdf_output_dir, doc_output_dir)
# Generate documents
generate_documents(template_file, student_list, doc_output_dir, pdf_output_dir)
print("All documents have been processed!")



def excel(filename):
    """Load the Excel file and return the sheet values."""
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    return list(sheet.values)

def save_document(template, docx_directory, value):
    """Render the document with the given values and save it as a .docx file."""
    doc = DocxTemplate(template)
    doc.render({
        "student_id": value[0],
        "first_name": value[1],
        "last_name": value[2],
        "logic": value[3],
        "l_g": value[4],
        "bcum": value[5],
        "bc_g": value[6],
        "design": value[7],
        "d_g": value[8],
        "p1": value[9],
        "p1_g": value[10],
        "e1": value[11],
        "e1_g": value[12],
        "wd": value[13],
        "wd_g": value[14],
        "algo": value[15],
        "al_g": value[16],
        "p2": value[17],
        "p2_g": value[18],
        "e2": value[19],
        "e2_g": value[20],
        "sd": value[21],
        "sd_g": value[22],
        "js": value[23],
        "js_g": value[24],
        "php": value[25],
        "ph_g": value[26],
        "db": value[27],
        "db_g": value[28],
        "vc1": value[29],
        "v1_g": value[30],
        "node": value[31],
        "no_g": value[32],
        "e3": value[33],
        "e3_g": value[34],
        "p3": value[35],
        "p3_g": value[36],
        "oop": value[37],
        "op_g": value[38],
        "lar": value[39],
        "la_g": value[40],
        "vue": value[41],
        "vu_g": value[42],
        "vc2": value[43],
        "v2_g": value[44],
        "e4": value[45],
        "e4_g": value[46],
        "p4": value[47],
        "p4_g": value[48],
        "int": value[49],
        "in_g": value[50]
    })
    doc_name = os.path.join(docx_directory, value[1] + ".docx")
    doc.save(doc_name)
    return doc_name

def convert_to_pdf(doc_path, pdf_directory):
    """Convert the .docx file to a .pdf file."""
    pdf_name = os.path.join(pdf_directory, os.path.splitext(os.path.basename(doc_path))[0] + ".pdf")
    convert(doc_path, pdf_name)

def main():
    filename = "data.xlsx"
    template = "template-pnc.docx"
    docx_directory = "Template_Documents"
    pdf_directory = "Template_PDF"

    os.makedirs(docx_directory, exist_ok=True)
    os.makedirs(pdf_directory, exist_ok=True)

    name_data = excel(filename)

    for value_tuple in name_data[1:]:
        doc_path = save_document(template, docx_directory, value_tuple)
        convert_to_pdf(doc_path, pdf_directory)

    print("All documents have been processed!")

# Call the main function directly
main()
