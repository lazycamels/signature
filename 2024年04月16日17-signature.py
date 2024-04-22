import os
import fitz
import tkinter as tk
import shutil
from PIL import Image
from tkinter import filedialog, messagebox

script_dir = os.path.abspath(os.path.dirname(__file__))# 获取当前脚本的绝对路径
cache_dir = os.path.join(script_dir, 'cache')# 在脚本目录下创建一个名为 'cache' 的子目录用于存放缓存文件
if not os.path.exists(cache_dir):
    os.makedirs(cache_dir)

window1 = tk.Tk()
window1.withdraw()

pdfPath = fitz.open(filedialog.askopenfilename(title="选择已赋码的PDF文件",
                                               filetypes=[("PDF files", "*.PDF"), ("All files", "*.*")]))
imagePath = cache_dir

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
    img = Image.open(os.path.join(cache_dir, '1.png'))
    cropped_img = img.crop(crop_area)
    cropped_img.save(os.path.join(cache_dir, '二维码.png'))

img = Image.open(os.path.join(cache_dir, '1.png'))
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
    png_path = os.path.join(cache_dir)
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

    home_directory = os.path.expanduser('~')
    desktop_directory = os.path.join(home_directory, 'Desktop')
    output_pdf_path = os.path.join(desktop_directory, '签字页.pdf' )

    if os.path.exists(output_pdf_path):
        os.remove(output_pdf_path)
        print(f"文件'{os.path.join(desktop_directory, '签字页.pdf' )}'已经存在于桌面目录中，即将被覆盖。")
    try:
        combined_img.save(output_pdf_path, "pdf", resolution=100.0)
        print(f"图片合成并保存为PDF成功！文件已保存为：{output_pdf_path}")
    except Exception as e:
        print(f"保存PDF文件失败：{e}")

image1_path = os.path.join(cache_dir, '二维码.png')
image2_path = os.path.join(cache_dir, '3.png')
combine_images(image1_path, image2_path)

os.remove(os.path.join(cache_dir, '二维码.png'))
os.remove(os.path.join(cache_dir, '1.png'))
os.remove(os.path.join(cache_dir, '2.png'))
os.remove(os.path.join(cache_dir, '3.png'))