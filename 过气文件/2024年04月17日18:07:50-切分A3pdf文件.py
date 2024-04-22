import fitz  # PyMuPDF
from PIL import Image #pillow
#from reportlab.pdfgen import canvas
import os
import img2pdf #img2pdf

def convert_pdf_to_png(pdf_path, output_folder, zoom_x=2, zoom_y=2):
    pdf_document = fitz.open(pdf_path)
    for page_number in range(len(pdf_document)):
        page = pdf_document[page_number]

        mat = fitz.Matrix(zoom_x, zoom_y)

        pix = page.get_pixmap(matrix=mat, alpha=False)

        output_filename = f"{output_folder}/page_{page_number + 1}.png"

        pix.save(output_filename)

    pdf_document.close()


pdf_file_path = '/Users/zhangyidong/Desktop/A3文件.pdf'  # PDF文件路径
output_folder_path = '/Users/zhangyidong/Desktop'  # 输出文件夹路径
convert_pdf_to_png(pdf_file_path, output_folder_path, zoom_x=4, zoom_y=4)

##########

def split_image_vertically(input_image_path, output_folder):
    with Image.open(input_image_path) as img:
        width, height = img.size

        top_part = img.crop((0, 0, width, height // 2))
        bottom_part = img.crop((0, height // 2, width, height))

        top_part.save(f"{output_folder}/top_part.png")
        bottom_part.save(f"{output_folder}/bottom_part.png")


input_image_path = '/Users/zhangyidong/Desktop/page_1.png'  # PNG图片路径
output_folder_path = '/Users/zhangyidong/Desktop'  # 输出文件夹路径
split_image_vertically(input_image_path, output_folder_path)

##########

# 图片文件路径
image_files = ['/Users/zhangyidong/Desktop/top_part.png','/Users/zhangyidong/Desktop/bottom_part.png']

# 将图片列表转换为PDF字节流
pdf_bytes = img2pdf.convert(image_files)

# 将PDF字节流写入文件
with open('/Users/zhangyidong/Desktop/一刀两断.pdf', 'wb') as f:
    f.write(pdf_bytes)

print("PDF文件已创建。")