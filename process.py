import os
import sys

from PyQt5.QtWidgets import QFileDialog, QDialog, QApplication
from PyQt5.uic.properties import QtWidgets

from calculate import Ui_MainWindow
from algorithm import read_excelfile

#主界面初始化
class MyPyQT_Form(Ui_MainWindow):
    def __init__(self):
        super().__init__()
    #继承calculate中定义的Ui)MainWindow类
    def setupUi(self, MainWindow):
        super().setupUi(MainWindow)
    #按下计算按钮，将触发“process”函数
        self.pushButton_2.clicked.connect(self.process)

#将文本框文本清除，并触发openfile函数
    def process(self):
        self.textBrowser.clear()
        self.openfile()

#openfile功能：
#根据起始年月日拼凑所有范围内的excel的文件名
    def openfile(self):
        start_year = int(self.comboBox_2.currentText())
        start_month = int(self.comboBox_3.currentText())
        end_year = int(self.comboBox_5.currentText())
        end_month = int(self.comboBox_4.currentText())
        self.filename = []
        year_cha = end_year - start_year
        year_add = 0
        #分两种情况读数
        #1.起始年都在一年内
        if(start_year == end_year):
            for num in range(start_month,end_month+1):
                if(num < 10):
                    self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year) + str(0) + str(num) + '.xlsx')
                else:
                    self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year) + str(num) + '.xlsx')
        else:
        # 2.起始年不在一年内
            for num in range(start_month, 13):
                if(num < 10):
                    self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year) + str(0) + str(num) + '.xlsx')
                else:
                    self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year) + str(num) + '.xlsx')

            while(year_cha > 1):
                year_add = year_add + 1
                for num in range(1, 13):
                    if(num < 10):
                        self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year + year_add) + str(0) + str(num) + '.xlsx')
                    else:
                        self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year + year_add) + str(num) + '.xlsx')
                year_cha = year_cha - 1

            for num in range(1, end_month+1):
                if(num < 10):
                    self.filename.append(os.getcwd() + '\excel\财务报表' + str(end_year) + str(0) + str(num) + '.xlsx')
                else:
                    self.filename.append(os.getcwd() + '\excel\财务报表' + str(end_year) + str(num) + '.xlsx')
        #根据下拉框索引选择不同的处理方式
        result = self.switch_case(self.comboBox.currentIndex())
        self.textBrowser.append(str(result))


        #self.excelfile.readmulti()

    def switch_case(self,value):
        switcher = {
            0: self.loadingrate()
        }
        return switcher.get(value, 'wrong value')

    #负债率计算，返回值为平均负债率
    def loadingrate(self):
        #负债率所用参数都在‘合并资产负载表’
        excelfile = read_excelfile(self.filename, '合并资产负债表')
        excelfile.readmulti()
        return excelfile.loadrate()


