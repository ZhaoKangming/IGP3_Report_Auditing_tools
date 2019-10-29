from pptx import Presentation
import re
import pprint

# 通用函数
def only_chinese(content):
    # 处理前进行相关的处理，包括转换成Unicode等
    pattern = re.compile('[^\u4e00-\u9fa50-9]')  # 中文的编码范围是：\u4e00到\u9fa5
    zh_str = "".join(pattern.split(content))
    return zh_str


# 【获取报告的文本内容】-----------------------------------------------------------------------------------

def get_pptx_content(rep_numb: int) -> dict:
    '''
    【功能】依据编号获取pptx文件内的文本内容
    :param rep_numb：报告在reports_info_list中的index值
    '''
    pptx_path: str = reports_info_list[i]  # TODO:
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
                    text += run.text.replace(' ', '').replace('_', '')
        content_dict[i] = text
        backup_content_dict[i] = backup_text
        i += 1

    pprint.pprint(content_dict)
    return content_dict

# -----------------------------------------------------------------------------------

def audit_slide():
    slide_sort_dict: dict = {
        '首页': ['胰岛素规范临床实践', []],
        '注意事项': ['注意事项', []],
        'BaseLine': []
    }

    # ----------------------- 首页 -----------------------
    if '胰岛素规范临床实践' in content_dict[0]:  # 获取首页页码
        slide_sort_dict['首页'][1].append(0)

    temp_text: str = only_chinese("胰岛素规范临床实践总结报告汇报人：范晓东")   # TODO:

    # 检查汇报人姓名
    del_list: list = ['胰岛素规范临床实践', '总结', '报告', '汇报人']
    for del_str in del_list:
        temp_text = temp_text.replace(del_str, '')
    if len(temp_text) > 0:
        if '医生姓名' not in temp_text:  # TODO:
            audit_result[2].append('【首页】汇报人姓名与上传报告医生姓名不一致！')



    # ----------------------- 注意事项 -----------------------
    title_content_text: str = '' # TODO:内容目录的文本
    lack_title_list: list = []
    title_list: list = ['患者情况汇总', '治疗方案', '治疗结果', '典型病例分享', '胰岛素规范实践的获益', '胰岛素规范实践临床展望']
    for title_str in title_list:
        if not title_str in title_content_text:
            lack_title_list.append(title_str)
    if 



# ------------------ 程序主体调用部分 ------------------
audit_result: list = [['审核结果'], ['修改记录'], ['错误记录']]

rep_numb: int = 1  # TODO:
content_dict: dict = get_pptx_content(rep_numb)

audit_slide()

print(audit_result)
