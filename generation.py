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
