# zhangyidong
# 时间：2024/4/10 20:49
import os
import fitz
import tkinter as tk
from tkinter import filedialog
from PIL import Image

root = tk.Tk()
root.withdraw()

def covert2pic(file_path, zoom, png_path, page_number=1):
    doc = fitz.open(file_path)
    # 确保page_number在页面总数范围内
    if page_number <= doc.page_count:
        page = doc[page_number - 1]  # 使用-1是因为索引从0开始
        zoom = int(zoom)  # 值越小，分辨率越高，文件越清晰
        rotate = int(0)

        trans = fitz.Matrix(zoom / 60.0, zoom / 60.0).prerotate(rotate)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        if not os.path.exists(png_path):
            os.mkdir(png_path)
        save = os.path.join(png_path, f'{page_number}.png')  # 使用f-string格式化字符串
        pm.save(save)
    doc.close()

if __name__ == "__main__":

    pdfPath = fitz.open(filedialog.askopenfilename(title="选择pdf文件", filetypes=[("PDF files", "*.PDF"), ("All files", "*.*")]))
    imagePath = '/Users/zhangyidong/Desktop/缓存/二维码'
    # 指定只保存第一页
    covert2pic(pdfPath, 200, imagePath, 1)

def crop_image(image_path, crop_area, output_path):
    # 打开图片
    img = Image.open('/Users/zhangyidong/Desktop/缓存/二维码/1.png')
    # 裁切区域，格式为(左上角x坐标, 左上角y坐标, 右下角x坐标, 右下角y坐标)
    cropped_img = img.crop(crop_area)
    # 保存裁切后的图片
    cropped_img.save('/Users/zhangyidong/Desktop/缓存/二维码/二维码.png')

# 指定裁切区域，假设我们想要裁切右下角200x200像素的区域
# 首先需要获取原始图片的尺寸
img = Image.open('/Users/zhangyidong/Desktop/缓存/二维码/1.png')
width, height = img.size

# 定义裁切区域的坐标，这里以右下角200x200像素为例
# 起始坐标为(width - 200, height - 200)，结束坐标为(width, height)
crop_area = (width-228, height-220, width-47, height-38)

# 调用裁切函数
crop_image('original_image.png', crop_area, 'cropped_image.png')