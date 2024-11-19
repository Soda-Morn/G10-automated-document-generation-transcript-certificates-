
import openpyxl
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
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
            'cur_date': student[12]
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

