import os
import tkinter as tk
from tkinter.simpledialog import askstring
from docx import Document

old_file_path = '/Users/zhangyidong/desktop/docx/'
new_file_path = '/Users/zhangyidong/desktop/new_docx/'

def check_and_change(document, replace_dict):
    for para in document.paragraphs:
        for run in para.runs:
            for key, value in replace_dict.items():
                if key in run.text:
                    run.text = run.text.replace(key, value)
    return document

def get_input_from_user(prompt, title):
    root = tk.Tk()
    root.withdraw()
    value = askstring(title, prompt)
    return value

replace_dict = {
    '-公司名称': '',
    '-报告年度': '',
    '-报告号': '',
    '-报告日期': '',
    '-成立日期': '',
    '-公司住所': '',
    '-注册资本': '',
    '-信用代码': '',
    '-法定代表人': '',
    '-经营范围': ''
}

for key in replace_dict:
    value = get_input_from_user(f'Enter new value for {key}', replace_dict[key])
    if value:
        replace_dict[key] = value

for name in os.listdir(old_file_path):
    if name.endswith('.docx'):
        old_file = os.path.join(old_file_path, name)
        document = Document(old_file)
        document = check_and_change(document, replace_dict)
        new_file = os.path.join(new_file_path, name)
        document.save(new_file)
        print(f'处理结束 {name}.')

