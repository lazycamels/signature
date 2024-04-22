# coding=utf-8

import os
from docx import Document

old_file_path = '/Users/zhangyidong/desktop/docx/'
new_file_path = '/Users/zhangyidong/desktop/new_docx/'

replace_dict = {'-信用代码': '91410000170001401D', '-公司名称': '宇通客车股份有限公司', '-法定代表人': '汤玉祥', '-成立日期': '1997年01月08日', '-注册资本': '221393.922300万人民币', '-公司住所': '郑州市管城回族区宇通路6号', '-经营范围': '经营本企业自产产品及相关技术的出口业务；经营本企业生产、科研所需的原辅材料、机械设备、仪器仪表、零配件及相关技术的进口业务；经营本企业的进料加工和“三来一补”业务；改装汽车、挂车、客车及配件附件、客车底盘、信息安全设备、智能车载设备的设计、生产与销售；机械加工、汽车整车及零部件的技术开发、转让、咨询与服务；通用仪器仪表制造与销售；质检技术服务；摩托车、旧车及配件、机电产品、五金交电、百货、互联网汽车、化工产品（不含易燃易爆化学危险品）、润滑油的销售；汽车维修（限分支机构凭证经营）；住宿、饮食服务（限其分支机构凭证经营）；普通货运；仓储（除可燃物资）；租赁业；旅游服务；公路旅客运输；县际非定线旅游、市际非定线旅游；软件和信息技术，互联网平台、安全、数据、信息服务；第二类增值电信业务中的信息服务业务（不含固定网电话信息服务和互联网信息服务）；经营第Ⅱ类、第Ⅲ类医疗器械（详见许可证）；保险兼业代理；对外承包工程业务；工程（建设及）管理服务；新能源配套基础设施的设计咨询、建设及运营维护；通讯设备、警用装备、检测设备的销售；计算机信息系统集成。涉及许可经营项目，应取得相关部门许可后方可经营',
     '-报告年度' : '2023年',
     '-报告号' : '0411号',
     '-报告日期' : '2024年04月17日'
                }

def check_and_change(document, replace_dict):
    for para in document.paragraphs:
        for i in range(len(para.runs)):
            for key,value in replace_dict.items():
                if key in para.runs[i].text:
                    print(key+'-->'+value)
                    para.runs[i].text = para.runs[i].text.replace(key, value)
    return document


for name in os.listdir(old_file_path):
    print(name)
    old_file = old_file_path+name
    new_file = new_file_path+name
    if old_file.split('.')[1] == 'docx':
        document = Document(old_file)
        document = check_and_change(document, replace_dict)
        document.save(new_file)
    print('^'*30)

