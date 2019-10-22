#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from PyQt5.QtGui import *
from ui_report_checker import Ui_MainWindow
import get_reports
import read_errorcsv
import subprocess
import sys
import os
import datetime
import requests
import io


#TODO: 一个人一天内多次提交
#TODO: 重命名提交给诺和的功能
#TODO: 导出审核记录、

class Main(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(Main, self).__init__()
        self.setupUi(self)
        
        # 绑定信号与槽
        self.load_report_list_btn.clicked.connect(self.load_report_list)
        self.clear_report_list_btn.clicked.connect(self.clear_report_list)
        self.original_folder_btn.clicked.connect(self.open_original_folder)
        self.passed_folder_btn.clicked.connect(self.open_passed_folder)
        self.open_ppt_btn.clicked.connect(self.open_selected_ppt)
        self.download_selected_btn.clicked.connect(self.download_selected_report)
        self.dwonload_page_btn.clicked.connect(self.download_page_report)
        self.download_all_btn.clicked.connect(self.download_all_report)
        self.submit_result_btn.clicked.connect(self.submit_result)

    def get_selected_rows(self) -> list:
        selected_rows_list: list = []
        item = self.report_info_table.selectedItems()
        for i in item:
            if self.report_info_table.indexFromItem(i).row() not in selected_rows_list:
                selected_rows_list.append(self.report_info_table.indexFromItem(i).row())
        return selected_rows_list


    def load_report_list(self):
        # reports_info_list: list = get_reports.get_reports_info_list()
        row: int = len(reports_info_list)  # 取得记录个数，用于设置表格的行数
        self.report_info_table.setRowCount(row)

        for i in range(row):
            audit_cb = QComboBox()
            audit_cb.addItem("     通过")  # 多余的空格是为了居中
            audit_cb.addItem("     退回")  # 多余的空格是为了居中
            self.report_info_table.setCellWidget(i, 5, audit_cb)
            for j in range(4): # 只需要显示前四项参数
                temp_data = reports_info_list[i][j]  # 临时记录，不能直接插入表格
                data = QTableWidgetItem(str(temp_data))  # 转换后可插入表格
                self.report_info_table.setItem(i, j, data)
                self.report_info_table.item(i, j).setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                


    def clear_report_list(self):
        '''
        【功能】清空报告列表
        '''
        reply = QMessageBox.question(self, 'Message', '确定要清空列表么?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.report_info_table.setRowCount(0)

    def download_feedback(self, dst_report_numb_list: list):
        '''
        【功能】尝试下载报告文件，并返回下载情况，并依据不同的下载状况返回不同颜色的下载状况到表格中
        :param  dst_report_numb_list: 报告序号列表
        '''
        information = QMessageBox.information(self, '温馨提醒', '报告下载中，请耐心等待！', QMessageBox.Yes, QMessageBox.Yes)
        for report_numb in dst_report_numb_list:
            download_state: str = get_reports.download_file(reports_info_list[report_numb])
            self.report_info_table.setItem(report_numb, 4, QTableWidgetItem(download_state))
            self.report_info_table.item(report_numb, 4).setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
            if download_state == '已下载':
                self.report_info_table.item(report_numb, 4).setForeground(QBrush(QColor(66, 184, 131))) # 绿色
            elif download_state == '无后缀名':
                errorcode: str = 'A2'
                self.report_info_table.item(report_numb, 4).setForeground(QBrush(QColor(178, 34, 34)))  # 红色
                self.report_info_table.setItem(report_numb, 6, QTableWidgetItem(error_dict[errorcode]))
                self.report_info_table.item(report_numb, 6).setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                self.report_info_table.cellWidget(report_numb, 5).setValue("     退回")
            elif download_state == '非PPT文件':
                errorcode: str = 'A1'
                self.report_info_table.item(report_numb, 4).setForeground(QBrush(QColor(178, 34, 34)))  # 红色
                self.report_info_table.setItem(report_numb, 6, QTableWidgetItem(error_dict[errorcode]))
                self.report_info_table.item(report_numb, 6).setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                self.report_info_table.cellWidget(report_numb, 5).setCurrentText("     退回")

    def download_selected_report(self):
        '''
        【功能】下载选中的报告文件
        '''
        dst_report_numb_list: list = Main.get_selected_rows(self)
        Main.download_feedback(self, dst_report_numb_list)
        


    def download_page_report(self):
        '''
        【功能】下载当前页所有的报告文件
        '''
        

    def download_all_report(self):
        '''
        【功能】下载全部报告文件
        '''
        dst_report_numb_list: list = list(range(len(reports_info_list)))
        Main.download_feedback(self, dst_report_numb_list)


    def submit_result(self):
        '''
        【功能】提交报告审核结果
        '''
        submit_report_numb_list: list = Main.get_selected_rows(self)
        today_date: str = str(datetime.date.today().strftime("%Y-%m-%d"))

        def post_report_result(rep_numb: int, operation_mode: int, back_reason: str):
            '''
            【功能】提交报告审核结果
            :param rep_numb：报告在表格中或者说是reports_info_list中的index值
            :param operation_mode: 操作模式，2代表退回，3代表通过
            :param back_reason: 退回原因，通过的话，back_reason = ''
            '''
            if operation_mode == 3:
                self.report_info_table.setItem(rep_numb, 6, '通过审核')
                self.report_info_table.item(rep_numb, 6).setForeground(QBrush(QColor(66, 184, 131)))  # 绿色
                reports_info_list[rep_numb].append(today_date)
                reports_info_list[rep_numb].append('通过')
                reports_info_list[rep_numb].append('--')
            elif operation_mode == 2:
                self.report_info_table.item(rep_numb, 6).setForeground(QBrush(QColor(178, 34, 34)))  # 红色
                reports_info_list[rep_numb].append(today_date)
                reports_info_list[rep_numb].append('退回')
                reports_info_list[rep_numb].append(back_reason)

            # 构造session进行登录
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')
            data = {'user_name': 'admin', 'user_password': '123456'}
            headers = {
                'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'}
            login_url = 'http://ydszn2nd.91huayi.com/pc/Manage/login'  # 登录时表单提交地址
            session = requests.Session()
            resp = session.post(login_url, data)

            # 进行post审核结果
            url = 'http://ydszn2nd.91huayi.com/pc/Manage/ReportsAudit'  # 报告审核页
            result_data = {'Rid': reports_info_list[rep_numb][5],
                            'Operation': operation_mode,
                            'AuditMsg': back_reason}
            resp2 = session.post(url, result_data)
            url_content: str = resp2.content.decode('utf-8')
            #TODO: {"error":0,"data":"2"} 判读返回值是否出错以及弹出警告


        if len(submit_report_numb_list) == 0:
            reply = QMessageBox.warning(self, '警告', '未选择待提交审核结果的行', QMessageBox.Yes, QMessageBox.Yes)
        else:    
            for submit_report_numb in submit_report_numb_list:
                audit_result: str = self.report_info_table.cellWidget(submit_report_numb, 5).currentText().replace(' ','')
                
                if audit_result == '通过':
                    if self.report_info_table.item(submit_report_numb, 6):
                        back_reason: str = self.report_info_table.item(submit_report_numb, 6).text().replace(' ','')
                        if len(back_reason) == 0:
                            post_report_result(submit_report_numb, 3, '')
                        else:
                            self.report_info_table.item(submit_report_numb, 6).setForeground(QBrush(QColor(6, 82, 121)))  # 靛蓝 
                            reply = QMessageBox.warning(self, '警告', f'第{submit_report_numb + 1}行，审核结果设为-通过，却有不合格原因！', QMessageBox.Yes, QMessageBox.Yes)
                    else:
                        post_report_result(submit_report_numb, 3, '')
                elif audit_result == '退回':
                    if self.report_info_table.item(submit_report_numb, 6):
                        back_reason: str = self.report_info_table.item(submit_report_numb, 6).text().replace(' ','')
                        if len(back_reason) == 0:
                            self.report_info_table.setItem(submit_report_numb, 6, '缺少不合格原因')
                            self.report_info_table.item(submit_report_numb, 6).setForeground(QBrush(QColor(6, 82, 121)))  # 靛蓝 
                            reply = QMessageBox.warning(self, '警告', f'第{submit_report_numb + 1}行，审核结果设为-退回，却无不合格原因！', QMessageBox.Yes, QMessageBox.Yes)
                        else:
                            post_report_result(submit_report_numb, 2, back_reason)
                    else:
                        self.report_info_table.setItem(submit_report_numb, 6, '缺少不合格原因')
                        self.report_info_table.item(submit_report_numb, 6).setForeground(QBrush(QColor(6, 82, 121)))  # 靛蓝 
                        reply = QMessageBox.warning(self, '警告', f'第{submit_report_numb + 1}行，审核结果设为-退回，却无不合格原因！', QMessageBox.Yes, QMessageBox.Yes)



    def open_original_folder(self):
        '''
        【功能】打开存储原始报告的文件夹
        '''
        os.system('start explorer '  + '..\\reports\\原始报告\\')


    def open_passed_folder(self):
        '''
        【功能】打开存储合格报告的文件夹
        '''
        os.system('start explorer ' + '..\\reports\\合格报告\\')

    def open_selected_ppt(self):
        '''
        【功能】打开所选项的原始文件
        '''
        dst_report_numb_list: list = Main.get_selected_rows(self)
        for dst_report_numb in dst_report_numb_list:
            if '.' in reports_info_list[dst_report_numb][4][-8:]:
                file_extension_name: str = reports_info_list[dst_report_numb][4][-8:].split('.')[1]
                file_name: str = reports_info_list[dst_report_numb][0] + '_' + \
                            reports_info_list[dst_report_numb][1] + '_' + \
                            reports_info_list[dst_report_numb][2].replace('报告', 'R') + '_' + \
                            reports_info_list[dst_report_numb][3].replace('20', '').replace('-', '') + '.' + \
                            file_extension_name
                if os.name == 'nt':
                    os.startfile(f"..\\reports\\原始报告\\{file_name}")
                elif os.name == 'posix':
                    subprocess.call(["open", f"..\\reports\\原始报告\\{file_name}"])
            else:
                reply = QMessageBox.warning(self, '警告', '此报告无扩展名！无法打开', QMessageBox.Yes, QMessageBox.Yes)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = Main()
    main.show()

    # 公用变量
    reports_info_list: list = get_reports.get_reports_info_list()
    error_dict: dict = read_errorcsv.get_errordict("records\\error_list.csv")
    
    sys.exit(app.exec_())

