import openpyxl
from docxtpl import DocxTemplate
from docx2pdf import convert
import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont

# function for generate certificate
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

# function for create associate
def AssociateExcel(filename):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    return list(sheet.values)

def Associate_doc(template, docx_directory, student):
    doc = DocxTemplate(template)
    doc.render({
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
            'ed_e': student[11]
    })
    doc_name = os.path.join(docx_directory, str(student[3]) + ".docx")
    doc.save(doc_name)
    return doc_name

def Associate_convertToPDF(doc_path, pdf_directory):
    pdf_name = os.path.join(pdf_directory, os.path.splitext(os.path.basename(doc_path))[0] + ".pdf")
    convert(doc_path, pdf_name)

def AssociateGenerate(generate_pdf = True):
    filename = "Associate.xlsx"
    template = "WEP Temporary Certificate - Template.docx"
    docx_directory = "Associate degree_Documents"
    pdf_directory = "Associate degree_PDF"

    os.makedirs(docx_directory, exist_ok=True)
    os.makedirs(pdf_directory, exist_ok=True)

    name_data = AssociateExcel(filename)

    for value_tuple in name_data[1:]:
        doc_path = Associate_doc(template, docx_directory, value_tuple)
        if generate_pdf:
            Associate_convertToPDF(doc_path, pdf_directory)
    if generate_pdf:
        print("All documents (DOC and PDF) have been processed!")
    else:
        print("All DOC files have been processed!")

#function for generate transcript
def TranscriptExcel(filename):
    """Load the Excel file and return the sheet values."""
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    return list(sheet.values)


def Transcript_document(template, docx_directory, value):
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

def TranscriptToPDF(doc_path, pdf_directory):
    pdf_name = os.path.join(pdf_directory, os.path.splitext(os.path.basename(doc_path))[0] + ".pdf")
    convert(doc_path, pdf_name)

def Transcript(generate_PDF):
    filename = "data.xlsx"
    template = "template-pnc.docx"
    docx_directory = "Template_Documents"
    pdf_directory = "Template_PDF"

    os.makedirs(docx_directory, exist_ok=True)
    os.makedirs(pdf_directory, exist_ok=True)

    name_data = TranscriptExcel(filename)

    for value_tuple in name_data[1:]:
        doc_path = Transcript_document(template, docx_directory, value_tuple)
        if generate_PDF:
            Associate_convertToPDF(doc_path, pdf_directory)
    if generate_PDF:
        print("All documents (DOC and PDF) have been processed!")
    else:
        print("All DOC files have been processed!")


#Function have choice for user
def GenetateChoice():
    print("Choose an option to generate documents:")
    print("1. Associate Degree")
    print("2. Certificates")
    print("3. Transcripts")
    choice = input("Enter your choice (1, 2, or 3): ")
    
    if choice == "1":
        print("Do you want to generate:")
        print("1. Only DOC files")
        print("2. Only PDF file")
        print("3. Generate PDF and DOC")
        associate_choice = input("Enter your choice (1 or 2): ")
        if associate_choice == "1":
            AssociateGenerate(generate_pdf = False)
        elif associate_choice == "2":
            AssociateGenerate(generate_pdf= True)
        elif associate_choice == "3":
            AssociateGenerate()
        else:
            print("Invalid choice for Associate Degree. Exiting.")
    elif choice == "2":
        generate_certificates(
            excel_file="Certificate.xlsx",
            template_file="certificate.png",
            output_folder="Certificates"
        )
    elif choice == "3":
        print("Do you want to generate:")
        print("1. Only DOC files")
        print("2. Only PDF file")
        print("3. Generate PDF and DOC")
        Transcript_choice = input("Enter your choice (1 or 2 or 3): ")
        if Transcript_choice == "1":
            Transcript(generate_PDF = False)
        elif Transcript_choice == "2":
            Transcript(generate_PDF= True)
        elif  Transcript_choice == "3":
            Transcript(generate_PDF)
        else:
            print("Invalid choice for Associate Degree. Exiting.")
    else:
        print("Invalid choice. Exiting.")
GenetateChoice()
