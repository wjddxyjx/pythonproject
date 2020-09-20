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
        #根据不同页数选择不同的处理方式
        tabindex = self.tabWidget.currentIndex()
        #分两种情况读数
        #1.起始年都在一年内
        if(tabindex == 2):
            excelfilename = '\excel\考核表\考核表'
        else:
            excelfilename = '\excel\财务报表\财务报表'

        if(start_year == end_year):
            if(start_month <= end_month):
                for num in range(start_month,end_month+1):
                    if(num < 10):
                        self.filename.append(os.getcwd() + excelfilename + str(start_year) + str(0) + str(num) + '.xlsx')
                        self.data.append(str(start_year)+'年'+str(num)+'月:')
                    else:
                        self.filename.append(os.getcwd() + excelfilename + str(start_year) + str(num) + '.xlsx')
                        self.data.append(str(start_year) + '年' + str(num) + '月:')
            else:
                QMessageBox.warning(None, '警告', '请确认好起始年月份是否有误！')
                return
        else:
        # 2.起始年不在一年内
            if (start_year < end_year):
                for num in range(start_month, 13):
                    if(num < 10):
                        self.filename.append(os.getcwd() + excelfilename + str(start_year) + str(0) + str(num) + '.xlsx')
                        self.data.append(str(start_year) + '年' + str(num) + '月:')
                    else:
                        self.filename.append(os.getcwd() + excelfilename + str(start_year) + str(num) + '.xlsx')
                        self.data.append(str(start_year) + '年' + str(num) + '月:')

                while(year_cha > 1):
                    year_add = year_add + 1
                    for num in range(1, 13):
                        if(num < 10):
                            self.filename.append(os.getcwd() + excelfilename + str(start_year + year_add) + str(0) + str(num) + '.xlsx')
                            self.data.append(str(start_year + year_add) + '年' + str(num) + '月:')
                        else:
                            self.filename.append(os.getcwd() + excelfilename + str(start_year + year_add) + str(num) + '.xlsx')
                            self.data.append(str(start_year + year_add) + '年' + str(num) + '月:')
                    year_cha = year_cha - 1

                for num in range(1, end_month+1):
                    if(num < 10):
                        self.filename.append(os.getcwd() + excelfilename + str(end_year) + str(0) + str(num) + '.xlsx')
                        self.data.append(str(end_year) + '年' + str(num) + '月:')
                    else:
                        self.filename.append(os.getcwd() + excelfilename + str(end_year) + str(num) + '.xlsx')
                        self.data.append(str(end_year) + '年' + str(num) + '月:')
            else:
                QMessageBox.warning(None, '警告', '请确认好起始年月份是否有误！')
                return
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
       elif (value == 2):
        return self.businessline()


    #勾选不同的指标
    def multistats(self):
        result = 0
        if(self.radioButton_2.isChecked()):
            numindex = self.comboBox_6.currentIndex()
            result = self.switch_index(numindex,None)
        elif(self.radioButton_3.isChecked()):
            numindex = self.comboBox_7.currentIndex()
            result = self.switch_index(numindex,None)
        else:
            QMessageBox.warning(None,'警告','请勾选要计算的指标！')
            result = -1
        return result

    def parentstats(self):
        result = 0
        if (self.radioButton_4.isChecked()):
            numindex = self.comboBox_8.currentIndex()
            result = self.switch_index(numindex,None)
        elif (self.radioButton_5.isChecked()):
            numindex = self.comboBox_9.currentIndex()
            result = self.switch_index(numindex,None)
        else:

            QMessageBox.warning(None, '警告', '请勾选要计算的指标！')
            result = -1
        return result

    def businessline(self):
        result = 0
        if (self.radioButton_6.isChecked()):
            numindex = self.comboBox_11.currentIndex()
            result = self.switch_index(numindex,self.comboBox.currentIndex())
        elif (self.radioButton_7.isChecked()):
            numindex = self.comboBox_10.currentIndex()
            result = self.switch_index(numindex,self.comboBox.currentIndex())
        else:

            QMessageBox.warning(None, '警告', '请勾选要计算的指标！')
            result = -1
        return result

    def switch_index(self,value,value2):
        tabindex = self.tabWidget.currentIndex()
        if(tabindex == 0):
            excelfile1 = read_excelfile(self.filename, '合并资产负债表')
            excelfile2 = read_excelfile(self.filename, '合并损益表')
            fid1 = excelfile1.readmulti()
            fid2 = excelfile2.readmulti()
            if((fid1 == -1)or(fid2 == -1)):
                return -1
            else:
                return self.switch_textindex(excelfile1,excelfile2, value,-1)
        elif(tabindex == 1):
            excelfile1 = read_excelfile(self.filename, '母公司资产负债表')
            excelfile2 = read_excelfile(self.filename, '母公司损益表')
            fid1 = excelfile1.readmulti()
            fid2 = excelfile2.readmulti()
            if((fid1 == -1)or(fid2 == -1)):
                return -1
            else:
                return self.switch_textindex(excelfile1,excelfile2, value,-1)
        elif(tabindex == 2):
            excelfile1 = read_excelfile(self.filename, '考核利润')
            fid1 = excelfile1.readmulti()
            if((fid1 == -1)):
                return -1
            else:
                return self.switch_textindex(excelfile1,None, value,value2)

    def switch_textindex(self,read_excelfile1,read_excelfile2,value,value2):
        if(value2 == -1):
            if (value == 0):
                return read_excelfile1.absolute1('资产总计','资产：')
            elif (value == 1):
                return read_excelfile1.absolute1('负债合计','负债：')
            elif (value == 2):
                return read_excelfile1.absolute1('所有者权益或股东权益合计','净资产：')
            elif (value == 3):
                return read_excelfile2.absolute2('一营业总收入','收入：')
            elif (value == 4):
                return read_excelfile2.absolute2('五净利润净亏损以号填列','净利润：')
        else:
            businessline_name = self.comboBox.currentText()
            aboslute_name = self.comboBox_11.currentText()
            if (value == 0):
                aboslute_name = '净利润'
            return read_excelfile1.absolute3(businessline_name,aboslute_name)


    #负债率计算，返回值为平均负债率
    def loadingrate(self):
        #负债率所用参数都在‘合并资产负载表’
        excelfile = read_excelfile(self.filename, '合并资产负债表')
        excelfile.readmulti()
        return excelfile.loadrate()


