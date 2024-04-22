# 时间：2024/4/16 20:46
from docx import Document
import re
import tkinter as tk
from docx import Document

doc = Document('/Users/zhangyidong/Desktop/docx/企业信息.docx')

rules =[
    (r'统一社会信用代码：(.*?)企业名称：','信用代码'),
    (r'企业名称：(.*?)注册号：','公司名称'),
    (r'法定代表人(.*?)类型：','法定代表人'),
    (r'成立日期：(.*?)注册资本：','成立日期'),
    (r'注册资本：(.*?)核准日期：','注册资本'),
    (r'住所：(.*?)经营范围：','住所'),
    (r'经营范围：(.*?)提示：','经营范围')
]

# 用于存储提取信息的字典
extracted_data = {}

# 遍历文档中的段落
for paragraph in doc.paragraphs:
    text = paragraph.text

    # 重置每段落的变量字典
    paragraph_data = {}

    # 尝试匹配每个规则
    for pattern, variable_name in rules:
        match = re.search(pattern, text, re.DOTALL)

        # 如果找到匹配项，则将其添加到当前段落的变量字典中
        if match:
            paragraph_data[variable_name] = match.group(1).strip()

    # 如果当前段落中有匹配的数据，将其添加到总的提取信息字典中
    if paragraph_data:
        extracted_data.update(paragraph_data)

for key, value in extracted_data.items():
    print(f'{key}: {value}')