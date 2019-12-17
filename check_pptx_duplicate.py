import os
import shutil
import win32com.client as win32
from pptx import Presentation
import re


# 定义公共变量
script_path: str = os.path.dirname(os.path.realpath(__file__))
passed_report_folder: str = os.path.join(script_path, '..', 'reports','合格报告')
summary_csv_path: str = os.path.join(script_path, 'records', 'summary_data.csv')
summary_dict: dict = {}


def ppt_to_pptx():
    '''
    【功能】删除隐藏文件、转化ppt为pptx
    '''
    app = win32.gencache.EnsureDispatch('PowerPoint.Application')

    for file in os.listdir(passed_report_folder):
        file_path: str = os.path.join(passed_report_folder, file)
        if file[0] == '.':
            os.remove(file_path)
        else:
            if os.path.splitext(file)[1] == '.ppt':
                print(f'正在转换文件 -- {file} --')
                new_pptx_path: str = file_path + 'x'
                backup_ppt_path: str = os.path.join(script_path, '合格ppt格式备份', file)
                shutil.copyfile(file_path, backup_ppt_path)   # 备份一下合格ppt报告
                office_obj = app.Presentations.Open(file_path, WithWindow=False)
                office_obj.SaveAs(new_pptx_path)
                office_obj.Close()
                os.remove(file_path)

    app.Quit()
    print('转换文件格式完成！')


def load_summary_data():
    '''
    【功能】载入总结数据csv
    '''
    global summary_dict
    with open(summary_csv_path, 'r', encoding='utf-8-sig') as csv_reader:
        content = csv_reader.read()
        rows = content.split('\n')
        for row in rows:
            if row != '':
                summary_dict[row.split(',')[0]] = row.split(',')[1]
        summary_dict.pop('报告文件名')


def update_summary_data():
    '''
    【功能】更新报告的总结数据
    '''
    global summary_dict
    with open(summary_csv_path, 'a+', encoding='utf-8-sig') as csv_writer:
        for file in os.listdir(passed_report_folder):
            if file[0] != '.' and os.path.splitext(file)[1] == '.pptx':
                if not file in summary_dict.keys():
                    summary_content: dict = get_pptx_content(os.path.join(passed_report_folder, file))



def get_pptx_content(pptx_path: str) -> dict:
    '''
    【功能】从报告中提取报告的总结部分
    【输出】内容字典，key为页码数，value为文本内容的连接
    :param pptx_path: 报告pptx的文件路径
    '''
    prs = Presentation(pptx_path)
    backup_content_dict: dict = {}
    content_dict: dict = {}
    i: int = 1

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
    
    return content_dict


def get_content_summary(content_dict: dict) -> list:
    '''
    【功能】在报告内容字典中获取医生自己写的总结部分
    '''
    have_summary: bool = False
    have_typical_case: bool = False
    start_page_numb: int = 0        # 总结与展望开始的页码
    end_page_numb: int = 0          # 总结与展望结束的页码
    summary_text: str = ''          # 总结与展望的文本
    sentence_list: list = []

    for i in range(1, len(content_dict)+1):
        if '胰岛素规范实践的获益与展望' in content_dict[i]:
            have_summary = True
            start_page_numb: int = i
            for j in range(i+1, i +4):
                if '典型病例分享' in content_dict[j]:
                    have_typical_case = True
                    end_page_numb: int = j - 1
                    break
            break
    
    if have_summary == True:
        if have_typical_case == False:
            end_page_numb: int = start_page_numb    # 如果在 “总结页” 后面的三页内都没有 “典型病例分享”，我们就默认总结只有一页

        # 合并总结部分的文本
        for i in range(start_page_numb, end_page_numb+1):
            summary_text += content_dict[i]

        



def compute_similarity():
    '''
    【功能】计算总结部分的相似度
    '''
    pass

#TODO:首先核对该医生的报告名称是否已经在字典中
