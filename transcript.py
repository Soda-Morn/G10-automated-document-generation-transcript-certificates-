# import openpyxl
# from docxtpl import DocxTemplate
# from docx2pdf import convert
# import os



# # Load the Excel file
# filename = "data.xlsx"
# workbook = openpyxl.load_workbook(filename)
# sheet = workbook.active
# name = list(sheet.values)

# # Load the Word template
# doc = DocxTemplate("template-pnc.docx")

# # Directory paths to save the documents and PDFs
# docx_directory = "Template_Documents" 
# pdf_directory = "Template_PDF"
# # Create directories if they don't exist
# os.makedirs(docx_directory, exist_ok=True)
# os.makedirs(pdf_directory, exist_ok=True)
# # Process each row in the Excel sheet
# for value_tuple in name[1:]:
#     doc.render({
#         "student_id": value_tuple[0],
#         "first_name": value_tuple[1],
#         "last_name": value_tuple[2],
#         "logic": value_tuple[3],
#         "l_g": value_tuple[4],
#         "bcum": value_tuple[5],
#         "bc_g": value_tuple[6],
#         "design": value_tuple[7],
#         "d_g": value_tuple[8],
#         "p1": value_tuple[9],
#         "p1_g": value_tuple[10],
#         "e1": value_tuple[11],
#         "e1_g": value_tuple[12],
#         "wd": value_tuple[13],
#         "wd_g": value_tuple[14],
#         "algo": value_tuple[15],
#         "al_g": value_tuple[16],
#         "p2": value_tuple[17],
#         "p2_g": value_tuple[18],
#         "e2": value_tuple[19],
#         "e2_g": value_tuple[20],
#         "sd": value_tuple[21],
#         "sd_g": value_tuple[22],
#         "js": value_tuple[23],
#         "js_g": value_tuple[24],
#         "php": value_tuple[25],
#         "ph_g": value_tuple[26],
#         "db": value_tuple[27],
#         "db_g": value_tuple[28],
#         "vc1": value_tuple[29],
#         "v1_g": value_tuple[30],
#         "node": value_tuple[31],
#         "no_g": value_tuple[32],
#         "e3": value_tuple[33],
#         "e3_g": value_tuple[34],
#         "p3": value_tuple[35],
#         "p3_g": value_tuple[36],
#         "oop":value_tuple[37],
#         "op_g":value_tuple[38],
#         "lar":value_tuple[39],
#         "la_g":value_tuple[40],
#         "vue":value_tuple[41],
#         "vu_g":value_tuple[42],
#         "vc2":value_tuple[43],
#         "v2_g":value_tuple[44],
#         "e4":value_tuple[45],
#         "e4_g":value_tuple[46],
#         "p4":value_tuple[47],
#         "p4_g":value_tuple[48],
#         "int":value_tuple[49],
#         "in_g":value_tuple[50]
#     })
    
#     #Create folder for put document
#     Template_Documents = os.path.join(docx_directory, f"{value_tuple[1]}.docx")  # Filename based on student ID
#     doc.save(Template_Documents)
#     print(f"{Template_Documents} has been created.")
    
#     # Convert the Word document to PDF and save in the `pdf_output_dir` directory
#     Template_PDF = os.path.join(pdf_directory, f"{value_tuple[1]}.pdf")
#     convert(Template_Documents, Template_PDF)
#     print(f"{Template_PDF} has been created.")

# print("All documents have been processed!")


import openpyxl
from docxtpl import DocxTemplate
from docx2pdf import convert
import os

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
        "first_name": value[1],"last_name": value[2],
        "logic": value[3],"l_g": value[4],
        "bcum": value[5],"bc_g": value[6],
        "design": value[7],"d_g": value[8],
        "p1": value[9],"p1_g": value[10],
        "e1": value[11],"e1_g": value[12],
        "wd": value[13],"wd_g": value[14],
        "algo": value[15],"al_g": value[16],
        "p2": value[17],"p2_g": value[18],
        "e2": value[19],"e2_g": value[20],
        "sd": value[21],"sd_g": value[22],
        "js": value[23],"js_g": value[24],
        "php": value[25],"ph_g": value[26],
        "db": value[27],"db_g": value[28],
        "vc1": value[29],"v1_g": value[30],
        "node": value[31],"no_g": value[32],
        "e3": value[33],"e3_g": value[34],
        "p3": value[35],"p3_g": value[36],
        "oop": value[37],"op_g": value[38],
        "lar": value[39],"la_g": value[40],
        "vue": value[41],"vu_g": value[42],
        "vc2": value[43],"v2_g": value[44],
        "e4": value[45],"e4_g": value[46],
        "p4": value[47],"p4_g": value[48],
        "int": value[49],"in_g": value[50]
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
