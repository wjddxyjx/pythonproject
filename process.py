import os
import sys

from PyQt5.QtWidgets import QFileDialog, QDialog, QApplication, QMessageBox
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
        self.data = []
        year_cha = end_year - start_year
        year_add = 0
        #分两种情况读数
        #1.起始年都在一年内
        if(start_year == end_year):
            if(start_month <= end_month):
                for num in range(start_month,end_month+1):
                    if(num < 10):
                        self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year) + str(0) + str(num) + '.xlsx')
                        self.data.append(str(start_year)+'年'+str(num)+'月:')
                    else:
                        self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year) + str(num) + '.xlsx')
                        self.data.append(str(start_year) + '年' + str(num) + '月:')
            else:
                QMessageBox.warning(None, '警告', '请确认好起始年月份是否有误！')
                return
        else:
        # 2.起始年不在一年内
            if (start_year < end_year):
                for num in range(start_month, 13):
                    if(num < 10):
                        self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year) + str(0) + str(num) + '.xlsx')
                        self.data.append(str(start_year) + '年' + str(num) + '月:')
                    else:
                        self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year) + str(num) + '.xlsx')
                        self.data.append(str(start_year) + '年' + str(num) + '月:')

                while(year_cha > 1):
                    year_add = year_add + 1
                    for num in range(1, 13):
                        if(num < 10):
                            self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year + year_add) + str(0) + str(num) + '.xlsx')
                            self.data.append(str(start_year + year_add) + '年' + str(num) + '月:')
                        else:
                            self.filename.append(os.getcwd() + '\excel\财务报表' + str(start_year + year_add) + str(num) + '.xlsx')
                            self.data.append(str(start_year + year_add) + '年' + str(num) + '月:')
                    year_cha = year_cha - 1

                for num in range(1, end_month+1):
                    if(num < 10):
                        self.filename.append(os.getcwd() + '\excel\财务报表' + str(end_year) + str(0) + str(num) + '.xlsx')
                        self.data.append(str(end_year) + '年' + str(num) + '月:')
                    else:
                        self.filename.append(os.getcwd() + '\excel\财务报表' + str(end_year) + str(num) + '.xlsx')
                        self.data.append(str(end_year) + '年' + str(num) + '月:')
            else:
                QMessageBox.warning(None, '警告', '请确认好起始年月份是否有误！')
                return
            #根据不同页数选择不同的处理方式
        tabindex = self.tabWidget.currentIndex()
        result = self.switch_case(tabindex)
        if(result == -1):
            return
        else:
            linenum = len(result)
            for num in range(0, linenum):
                self.textBrowser.append((self.data[num]))
                self.textBrowser.append((result[num]))

        #self.excelfile.readmulti()

    def switch_case(self,value):
       if(value == 0):
        return self.multistats()
       elif(value == 1):
        return self.parentstats()


    #勾选不同的指标
    def multistats(self):
        result = 0
        if(self.radioButton_2.isChecked()):
            numindex = self.comboBox_6.currentIndex()
            result = self.switch_index(numindex)
        elif(self.radioButton_3.isChecked()):
            numindex = self.comboBox_7.currentIndex()
            result = self.switch_index(numindex)
        else:
            QMessageBox.warning(None,'警告','请勾选要计算的指标！')
            result = -1
        return result

    def parentstats(self):
        result = 0
        if (self.radioButton_4.isChecked()):
            numindex = self.comboBox_6.currentIndex()
            result = self.switch_index(numindex)
        elif (self.radioButton_5.isChecked()):
            numindex = self.comboBox_7.currentIndex()
            result = self.switch_index(numindex)
        else:

            QMessageBox.warning(None, '警告', '请勾选要计算的指标！')
            result = -1
        return result

    def switch_index(self,value):
        tabindex = self.tabWidget.currentIndex()
        if(tabindex == 0):
            excelfile = read_excelfile(self.filename, '合并资产负债表')
            fid = excelfile.readmulti()
            if(fid == -1):
                return -1
            else:
                return self.switch_textindex(excelfile, value)
        elif(tabindex == 1):
            excelfile = read_excelfile(self.filename, '母公司资产负债表')
            fid = excelfile.readmulti()
            if(fid == -1):
                return -1
            else:
                return self.switch_textindex(excelfile, value)

    def switch_textindex(self,read_excelfile ,value):
        if (value == 0):
            return read_excelfile.absolute1('资产总计')
        elif (value == 1):
            return read_excelfile.absolute1('负债合计')
        elif (value == 2):
            return read_excelfile.absolute('')
        elif (value == 3):
            return read_excelfile.absolute('')

    #负债率计算，返回值为平均负债率
    def loadingrate(self):
        #负债率所用参数都在‘合并资产负载表’
        excelfile = read_excelfile(self.filename, '合并资产负债表')
        excelfile.readmulti()
        return excelfile.loadrate()


