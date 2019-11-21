# -- coding:utf-8 --  
# author:ZhaoKangming
# 功能：爬虫之模拟登录赋能启航第二期后台

import requests
import sys
import io
from bs4 import BeautifulSoup
import os
import urllib.request
import shutil


def login_get_urlcontent() -> str:
    '''
    【功能】模拟登陆赋能起航二期后台，获取报告页网页内容
    '''
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf8') #改变标准输出的默认编码

    #登录时需要POST的数据
    data = {'user_name': 'admin', 'user_password': '123456'}
    headers = {'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'}
    login_url = 'http://ydszn2nd.91huayi.com/pc/Manage/login'  # 登录时表单提交到的地址

    session = requests.Session()  # 构造Session

    #在session中发送登录请求，此后这个session里就存储了cookie
    #可以用print(session.cookies.get_dict())查看
    resp = session.post(login_url, data)

    url = 'http://ydszn2nd.91huayi.com/pc/Manage/ReportsAudit?txtUserName=&txtDoctorName=&radAuditStatus=1&txtDateBegin=&txtDateEnd=&radReportType='  # 登录后才能访问的网页
    # 'http://ydszn2nd.91huayi.com/pc/Manage/ReportsAudit?page=2&radAuditStatus=1'
    resp = session.get(url)  # 发送访问请求
    url_content: str = resp.content.decode('utf-8')

    return url_content


def get_reports_info(content_text: str) -> list:
    '''
    【功能】从网页内容中解析出来报告信息
    :param content_text: requests的response网页内容
    '''
    soup = BeautifulSoup(content_text,'lxml')
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


def download_file(file_info_list: list) -> str:
    '''
    【功能】从列表中取值，下载文件并命名
    :param file_info_list: 报告信息列表
    '''
    ppt_extension_list: list = ['ppt', 'pptx']
    file_name : str = file_info_list[6]
    if '.' in file_name:
        file_extension_name: str = file_name.split('.')[1]
        backup_path: str = f'../reports/原始报告/{file_name}'
        urllib.request.urlretrieve(file_info_list[4], backup_path)
        if file_extension_name in ppt_extension_list:
            temp_path: str = f'../reports/temp_reports/{file_name}'
            shutil.copy(backup_path, temp_path)
            urllib.request.urlretrieve(file_info_list[4], temp_path)
            download_state: str = '已下载'
        else:
            download_state: str = f'非PPT文件'
    else:
        download_state: str = '无后缀名'

    return download_state

