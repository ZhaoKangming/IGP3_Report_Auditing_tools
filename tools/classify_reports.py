# -- coding:utf-8 --

'''
author: ZhaoKangming
version: 0.4
功能：爬虫赋能启航第二期后台未审核报告信息
'''

import requests
import sys
import io
from bs4 import BeautifulSoup
import os
import urllib.request
import shutil
import lxml

script_path: str = os.path.dirname(os.path.realpath(__file__))


def login_get_docInfoList() -> list:
    '''
    【功能】爬虫模拟登陆赋能起航二期后台，获取报告页网页内容
    '''
    sys.stdout = io.TextIOWrapper(
        sys.stdout.buffer, encoding='utf8')  # 改变标准输出的默认编码

    #登录时需要POST的数据
    data = {'user_name': 'admin', 'user_password': '123456'}
    headers = {
        'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'}
    login_url = 'http://ydszn2nd.91huayi.com/pc/Manage/login'  # 登录时表单提交到的地址
    session = requests.Session()  # 构造Session
    #在session中发送登录请求，此后这个session里就存储了cookie
    #可以用print(session.cookies.get_dict())查看
    resp = session.post(login_url, data)

    # 获取最大的页码数
    page1_url = 'http://ydszn2nd.91huayi.com/pc/Manage/ReportsAudit?txtUserName=&txtDoctorName=&radAuditStatus=1&txtDateBegin=&txtDateEnd=&radReportType='  # 登录后才能访问的网页
    resp = session.get(page1_url)  # 发送访问请求
    page1_url_content: str = resp.content.decode('utf-8')
    soup = BeautifulSoup(page1_url_content, 'lxml')
    li_text_list = soup.find("ul", "pagination").find_all("li")
    page_numb_list: list = []
    for li in li_text_list:
        a = li.a.string
        if a == '»»':
            pn = li.a['href'].replace('/pc/Manage/ReportsAudit?page=','').replace('&radAuditStatus=1','')
            page_numb_list.append(int(pn))
        elif not a in ['»', '…', '']:
            page_numb_list.append(int(a))
    max_page_numb: int = max(page_numb_list)
    reports_info_list: list = get_reports_info(page1_url_content)

    # 获取后续页面的数据
    if max_page_numb > 1:
        for page_numb in range(2, max_page_numb + 1):
            url = f'http://ydszn2nd.91huayi.com/pc/Manage/ReportsAudit?page={page_numb}&radAuditStatus=1'
            url_content: str = session.get(url).content.decode('utf-8')
            temp_list: list = get_reports_info(url_content)
            reports_info_list += temp_list

    # print(len(reports_info_list))
    return [max_page_numb, reports_info_list]


def get_reports_info(content_text: str) -> list:
    '''
    【功能】从网页内容中解析出来报告信息
    :param content_text: requests的response网页内容
    '''
    soup = BeautifulSoup(content_text, 'lxml')
    tr_text = soup.tbody.find_all('tr')
    reports_info_list: list = []
    for tr in tr_text:
        td_list = []
        td = tr.find_all('td')
        for d in td:
            if td.index(d) < 4:
                td_list.append(d.string)
            elif td.index(d) == 4:
                td_list.append('http://ydszn2nd.91huayi.com' + d.a['href'])
            elif td.index(d) == 6:
                td_list.append(d.string)
            elif td.index(d) == 7:
                td_list.append(d.button['value'])
        reports_info_list.append(td_list)

    # 获取医生的报告文件名
    for file_info_list in reports_info_list:
        if '.' in file_info_list[4][-8:]:
            file_extension_name: str = '.' + file_info_list[4][-8:].split('.')[1]
        else:
            file_extension_name: str = ''
        file_name: str = file_info_list[0] + '_' + \
            file_info_list[1] + '_' + \
            file_info_list[2].replace('报告', 'R') + '_' + \
            file_info_list[3].replace('2019', '19').replace('-', '') + \
            file_extension_name
        file_info_list.append(file_name)

    return reports_info_list


def get_resubmit_reports() -> list:
    reports_info_list: list = login_get_docInfoList()[1]
    resubmit_list: list = []
    # print(reports_info_list)
    for i in reports_info_list:
        if i[5]:
            if '.' in i[7]:
                resubmit_list.append(i[7].split('.')[0])
    return resubmit_list


def classify_reports():
    resubmit_list: list = get_resubmit_reports()
    resubmit_folder: str = os.path.join(script_path,'再次提交')
    if not os.path.exists(resubmit_folder):
        os.makedirs(resubmit_folder)
    for report_file in os.listdir(script_path):
        rp: str = os.path.splitext(report_file)[0]
        if rp in resubmit_list:
            print(f'  [RESUBMIT] ---> {rp}')
            src_file: str = os.path.join(script_path, rp)
            dst_file: str = os.path.join(resubmit_folder, rp)
            shutil.move(src_file,dst_file)
    print('\n' + '— — '*20 + '\n 处理完毕！')


classify_reports()
f = input('按任意按键退出！')
if f:
    sys.exit()                                                                                                                                                                                                                                                
