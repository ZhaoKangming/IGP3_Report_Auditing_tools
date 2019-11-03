# -*- coding:UTF-8 -*-

def get_error_dict(csv_path : str) -> dict:
    '''
    【功能】从csv文件中读取错误码与话术，输出一个字典
    :param csv_path: 存储错误码与话术的csv文件路径
    '''
    f = open(csv_path, 'r', encoding='utf-8-sig')
    content = f.read()
    final_dict: dict = {}
    rows = content.split('\n')
    for row in rows:
        if row != '':
            final_dict[row.split(',')[1]] = row.split(',')[2]
    final_dict.pop('错误码')
    return final_dict


def get_drugname_dict(csv_path : str) -> dict:
    '''
    【功能】从csv文件中读取药品的商品名与通用名，输出一个字典
    :param csv_path: 存储药品的商品名与通用名对应关系的csv文件路径
    '''
    f = open(csv_path, 'r', encoding='utf-8-sig')
    content = f.read()
    drugname_dict: dict = {}
    rows = content.split('\n')
    for row in rows:
        if row != '':
            drugname_dict[row.split(',')[0]] = row.split(',')[1]
    drugname_dict.pop('商品名')
    return drugname_dict
