import sys
import os
from PyQt5 import QtWidgets,QtGui,QtCore,Qt
from PyQt5.QtWidgets import QFileDialog,QProgressBar
import pandas as pd
import re
import threading
class GUI(QtWidgets.QWidget):
    def __init__(self):
        #初始化————init__
        super().__init__()
        self.desktop_path = os.path.join(os.path.expanduser("~"),'Desktop')
        self.initGUI()
    def initGUI(self):
        #设置窗口大小
        self.resize(580,400)
        #设置窗口位置(下面配置的是居于屏幕中间)
        qr = self.frameGeometry()
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        #设置窗口标题和图标
        self.setWindowTitle('editExcel1.0版')
        #######配置文件
        #设置按钮
        self.btn_pz =QtWidgets.QPushButton('添加配置文件路径',self)
        #大小
        self.btn_pz.resize(150,30)
        #位置
        self.btn_pz.move(20,20)
        
        #设置输入框
        self.textbox_pz = Qt.QLineEdit(self)
        #大小
        self.textbox_pz.resize(300, 30)
        #w位置
        self.textbox_pz.move(210, 20)
        #点击鼠标触发事件
        self.btn_pz.clicked.connect(self.select_pz_file_path)
        
        
        ######excel文件
        #设置按钮
        self.btn_excel =QtWidgets.QPushButton('添加excel文件路径',self)
        #大小
        self.btn_excel.resize(150,30)
        #位置
        self.btn_excel.move(20,70)
        
        #设置输入框
        self.textbox_excel = Qt.QLineEdit(self)
        #大小
        self.textbox_excel.resize(300, 30)
        #w位置
        self.textbox_excel.move(210, 70)
        #点击鼠标触发事件
        self.btn_excel.clicked.connect(self.select_excel_file_path)
        
        
        # 构建一个进度条
        self.pbar = QProgressBar(self)
        # 从左上角30-50的界面，显示一个200*25的界面
        self.pbar.setGeometry(70, 150, 450, 40)  # 设置进度条的位置

        ##执行button
        #设置按钮
        self.btn_act =QtWidgets.QPushButton('执行',self)
        #大小
        self.btn_act.resize(80,50)
        #位置
        self.btn_act.move(250,220)
        self.btn_act.clicked.connect(self.thread_deal)

        #展示窗口
        self.show();
    
    #点击鼠标触发函数
    def clickbtn(self):
        #打印出输入框的信息
        textboxValue = self.textbox.text()
        QtWidgets.QMessageBox.question(self, "信息", '你输入的输入框内容为:' + textboxValue,QtWidgets.QMessageBox.Ok, QtWidgets.QMessageBox.Ok)
        #清空输入框信息
        self.textbox.setText('')
        
    #关闭窗口事件重写
    def closeEvent(self, QCloseEvent):
        reply = QtWidgets.QMessageBox.question(self, '警告',"确定关闭当前窗口?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            QCloseEvent.accept()
        else:
            QCloseEvent.ignore()
            
    def select_pz_file_path(self):
        pz_path = QFileDialog.getOpenFileName(None,'选择配置文件',self.desktop_path,'*.xlsx')
        self.textboxValue_pz = pz_path[0]
        self.textbox_pz.setText(self.textboxValue_pz)
        
    def select_excel_file_path(self):
        excel_path = QFileDialog.getExistingDirectory(None,'选择需要处理的excel文件',self.desktop_path)
        self.textboxValue_excel = excel_path
        self.textbox_excel.setText(self.textboxValue_excel)
        
    def print_bar(self,get_num):
        percentage_values = (get_num/self.need_deal_count)*100
        self.pbar.setValue(percentage_values)

          
        
    def thread_deal(self):
        self.status = 0
        t1 = threading.Thread(target=self.deal_excel_file)
        t1.start()

        
    def deal_excel_file(self):
        
        file_data = pd.read_excel(self.textboxValue_pz)
        wExcel_path = self.textboxValue_excel
        all_excel_name = os.listdir(self.textboxValue_excel)
        tag = 0
        try:
            extract_col = file_data['excel表头名称'].values[0]
            extract_col = extract_col.strip()
            if ',' in extract_col and not '，' in extract_col:
                extract_col = extract_col.split(',')
            elif '，' in extract_col and not ',' in extract_col:
                extract_col = extract_col.split('，')
            tag = 1
        except:
            reply1 = QtWidgets.QMessageBox.critical(self, '配置文件错误',"更改excel表头名称")
            tag = 0
        if tag == 1:
            try:
                sheet_name_ = file_data['excel_sheet名称'].values[0]
                sheet_name_ = sheet_name_.strip()
                tag = 1
            except:
                reply2 = QtWidgets.QMessageBox.critical(self, '配置文件错误',"更改excel_sheet名称")
                tag = 0
        if tag == 1:    
            save_path = os.path.join(self.desktop_path,'提取')
            if not os.path.exists(save_path):
                os.mkdir(save_path)
            self.need_deal_count = len(all_excel_name)
            count = 0
            for i in all_excel_name:
    #            left_deal_count = self.need_deal_count - count
                count = count + 1
                self.print_bar(count)
    
                i_excel_name = re.findall('(.*)\.',i)[0]
                i_excel_path = os.path.join(wExcel_path,i)
                try:
                    i_df = pd.read_excel(i_excel_path,sheet_name=sheet_name_)
                except:
                    pass
                i_new_df = pd.DataFrame(columns = extract_col)
                for j in extract_col:
                    try:
                        i_new_df[j] = i_df[j]
                    except:
                        pass
                i_excel_new_name = i_excel_name + '_提取_.xlsx'
                i_save_path = os.path.join(save_path,i_excel_new_name)
                i_new_df.to_excel(i_save_path,index=None)

            
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    gui = GUI()
    sys.exit(app.exec_())


                
                
                
                
                
                
                
                
                
                

