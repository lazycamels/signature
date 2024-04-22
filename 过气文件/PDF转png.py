# zhang
# 时间：2024/3/31 15:15
import os
import fitz
import tkinter as tk
from tkinter import filedialog

def covert2pic(file_path, zoom, png_path):
    doc = fitz.open(file_path)
    total = doc.page_count
    for pg in range(total):
        page = doc[pg]
        zoom = int(zoom)  # 值越小，分辨率越高，文件越清晰
        rotate = int(0)

        trans = fitz.Matrix(zoom / 60.0, zoom / 60.0).prerotate(rotate)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        if not os.path.exists(png_path):
            os.mkdir(png_path)
        save = os.path.join(png_path, '%s.png' %(pg+1))
        pm.save(save)
    doc.close()

def ask_for_pdf():
    # 创建一个新的Tkinter窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 使用filedialog模块打开文件选择对话框
    pdfPath = filedialog.askopenfilename(
        title="选择PDF文件",
        filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
    )
    return pdfPath

if __name__ == "__main__":
    # 调用ask_for_pdf函数来获取用户选择的PDF文件路径
    pdfPath = ask_for_pdf()
    if pdfPath:
        zoom = 200  # 你可以根据需要调整这个值
        png_path = '/Users/zhangyidong/Desktop/缓存/签字页'
        covert2pic(pdfPath, zoom, png_path)