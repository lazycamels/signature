# zhangyidong
# 时间：2024/4/13 16:10
import subprocess
import os


# 定义要运行的脚本列表
scripts = ['将赋码PDF中的二维码提取出来.py', 'PDF转png.py','将二维码放入签字页.py']

# 依次运行脚本
for script in scripts:
    subprocess.run(['python', script])

os.remove('/Users/zhangyidong/Desktop/缓存/二维码/1.png')
os.remove('/Users/zhangyidong/Desktop/缓存/签字页/1.png')
os.remove('/Users/zhangyidong/Desktop/缓存/签字页/2.png')
os.remove('/Users/zhangyidong/Desktop/缓存/签字页/3.png')
