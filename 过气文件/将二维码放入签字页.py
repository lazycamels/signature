# zhangyidong
# 时间：2024/4/13 15:25
from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox

def select_image(title):
    file_path = filedialog.askopenfilename(title=title)
    return Image.open(file_path)

def combine_images(image1, image2):
    # 创建一个新的图像，尺寸和图片2相同，背景为白色
    combined_img = Image.new('RGB', image2.size, (255, 255, 255))
    # 将图片2粘贴到新图片上
    combined_img.paste(image2, (0, 0))
    # 计算图片1应该放置的位置，距离顶点14像素
    position = (image2.width - image1.width - 44, image2.height - image1.height - 44)
    # 将图片1粘贴到计算出的位置
    combined_img.paste(image1, position)

    # 弹出保存对话框，选择输出PDF的路径
    output_pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                    filetypes=[("PDF files", "*.pdf")],title="保存PDF文件")
    if output_pdf_path:
        # 保存为PDF格式
        combined_img.save(output_pdf_path, "PDF", resolution=100.0)
        messagebox.showinfo("完成", f"图片合成并保存为PDF成功！文件已保存为：{output_pdf_path}")
    else:
        messagebox.showwarning("取消", "操作已取消。")

# 运行程序
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 选择图片1
    image1 = select_image("选择二维码")
    if image1 is not None:
        # 选择图片2
        image2 = select_image("选择签字页")
        if image2 is not None:
            combine_images(image1, image2)