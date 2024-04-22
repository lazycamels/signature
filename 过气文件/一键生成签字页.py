# zhangyidong
# 时间：2024/4/14 13:33
import os
import fitz
import tkinter as tk
import shutil
from PIL import Image
from tkinter import filedialog, messagebox

window1 = tk.Tk()
window1.withdraw()

pdfPath = fitz.open(filedialog.askopenfilename(title="选择已赋码的PDF文件", filetypes=[("PDF files", "*.PDF"), ("All files", "*.*")]))
imagePath = '/Users/zhangyidong/Desktop/缓存/二维码'

def covert2pic(file_path, zoom, png_path, page_number=1):
    doc = fitz.open(file_path)
    if page_number <= doc.page_count:
        page = doc[page_number - 1]  # 使用-1是因为索引从0开始
        zoom = int(zoom)
        rotate = int(0)

        trans = fitz.Matrix(zoom / 60.0, zoom / 60.0).prerotate(rotate)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        if not os.path.exists(png_path):
            os.mkdir(png_path)
        save = os.path.join(png_path, f'{page_number}.png')
        pm.save(save)
    doc.close()

covert2pic(pdfPath, 200, imagePath, 1)

def crop_image(image_path, crop_area, output_path):
    img = Image.open('/Users/zhangyidong/Desktop/缓存/二维码/1.png')
    cropped_img = img.crop(crop_area)
    cropped_img.save('/Users/zhangyidong/Desktop/缓存/二维码/二维码.png')

img = Image.open('/Users/zhangyidong/Desktop/缓存/二维码/1.png')
width, height = img.size

crop_area = (width-228, height-220, width-47, height-38)

crop_image('original_image.png', crop_area, 'cropped_image.png')

def covert2pic(file_path, zoom, png_path):
    doc = fitz.open(file_path)
    total = doc.page_count
    for pg in range(total):
        page = doc[pg]
        zoom = int(zoom)
        rotate = int(0)

        trans = fitz.Matrix(zoom / 60.0, zoom / 60.0).prerotate(rotate)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        if not os.path.exists(png_path):
            os.mkdir(png_path)
        save = os.path.join(png_path, '%s.png' %(pg+1))
        pm.save(save)
    doc.close()

def ask_for_pdf():
    window2 = tk.Tk()
    window2.withdraw()

    pdfPath = filedialog.askopenfilename(
        title="选择未赋码的PDF文件",
        filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
    )
    return pdfPath

pdfPath = ask_for_pdf()
if pdfPath:
    zoom = 200
    png_path = '/Users/zhangyidong/Desktop/缓存/签字页'
    covert2pic(pdfPath, zoom, png_path)

def select_image(title):
    file_path = filedialog.askopenfilename(title=title)
    return Image.open(file_path)

def combine_images(image1_path, image2_path):
    image1 = Image.open(image1_path)
    image2 = Image.open(image2_path)

    combined_img = Image.new('RGB', image2.size, (255, 255, 255))
    combined_img.paste(image2, (0, 0))

    position = (image2.width - image1.width - 44, image2.height - image1.height - 44)
    combined_img.paste(image1, position)

    output_pdf_path = '/Users/zhangyidong/Desktop/签字页.pdf'

    if combined_img.save(output_pdf_path, "PDF", resolution=100.0):
        print(f"图片合成并保存为PDF成功！文件已保存为：{output_pdf_path}")
    else:
        print("操作已取消。")

image1_path = '/Users/zhangyidong/Desktop/缓存/二维码/二维码.png'
image2_path = '/Users/zhangyidong/Desktop/缓存/签字页/3.png'

combine_images(image1_path, image2_path)

#删除文件夹
shutil.rmtree('/Users/zhangyidong/Desktop/缓存/签字页')
shutil.rmtree('/Users/zhangyidong/Desktop/缓存/二维码')