from pptx import Presentation
import re
import pprint

# 通用函数
def only_chinese(content):
    # 处理前进行相关的处理，包括转换成Unicode等
    pattern = re.compile('[^\u4e00-\u9fa50-9]')  # 中文的编码范围是：\u4e00到\u9fa5
    zh_str = "".join(pattern.split(content))
    return zh_str



def get_pptx_content(pptx_path: str) -> dict:
    prs = Presentation(pptx_path)

    backup_content_dict: dict = {}
    content_dict: dict = {}
    i = 0

    for slide in prs.slides:
        text: str = ''
        backup_text: str = ''
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    backup_text += run.text
                    text += run.text.replace(' ','').replace('_','')
        content_dict[i] = text
        backup_content_dict[i] = backup_text
        i += 1

    pprint.pprint(content_dict)
    return content_dict


def get_slide_sort(content_dict: dict) -> dict:
    slide_sort_dict: dict = {
        '胰岛素规范临床实践' : [],
        'Content' : [],
        'BaseLine' : []
    }


    for k,v in content_dict.items():
        pass


def audit_slide(doc):
    

    # ----------------------- 首页 -----------------------
    # 检查汇报人姓名
    ppt_page_numb  # TODO:首页的文本
    cover_text: str = ''  # TODO:首页的文本
    temp_text: str = only_chinese(cover_text)
    del_list: list = ['胰岛素规范临床实践','总结','报告','汇报人']
    for del_str in del_list:
        temp_text = temp_text.replace(del_str,'')
    if len(temp_text):
        real_name: str = ''  #TODO:获取提交报告的医生姓名
        if real_name not in cover_text:
            audit_result[1].append(f'【第{ppt_page_numb}页】汇报医生与提交报告的医生姓名不一致！')

    # ----------------------- 内容目录 -----------------------
    title_content_text: str = '' # TODO:内容目录的文本
    lack_title_list: list = []
    title_list: list = ['患者情况汇总', '治疗方案', '治疗结果', '典型病例分享', '胰岛素规范实践的获益', '胰岛素规范实践临床展望']
    for title_str in title_list:
        if not title_str in title_content_text:
            lack_title_list.append(title_str)
    if 


def audit_report() -> list:
    
    pptx_path: str = r'D:\\1.pptx'
    get_pptx_content(pptx_path)
    


audit_result: list = [['审核结果'], ['修改记录'], ['错误记录']]
audit_report()
