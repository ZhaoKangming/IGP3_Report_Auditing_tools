def get_errordict(csv_path: str) -> dict:
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