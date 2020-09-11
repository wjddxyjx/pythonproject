import re

import pandas as pd
import numpy as np
import os

from PyQt5.QtWidgets import QMessageBox


class read_excelfile:
    #初始化类
    #类成员定义：
    #var:将要读入的excel文件名组，list类型
    #num:将要读入的excel个数
    #a:excel表的内容将会填充到这里，可以离近为二维矩阵，元素为二维矩阵的list
    #sheetvar: 选择读哪个sheet
    def __init__(self,var,sheetvar):
         self.var = var
         self.num = len(var)
         self.a = []
         self.sheetvar = sheetvar

    #函数功能：将excel的值加入list变量a中，并去除空格及标点符号
    def readmulti(self):

        for num in range(0, self.num):
            try:
                df1 = pd.read_excel(self.var[num], sheet_name = self.sheetvar,header=None)
                self.a.append(np.array(df1))
            except FileNotFoundError:
                QMessageBox.warning(None, '警告', '请检查excel文件是否齐全！')
                return -1
            else:
                for row1 in range(0, self.a[num].shape[0]):
                    for column in range(0, self.a[num].shape[1]):
                        if (isinstance(self.a[num][row1, column], str)):
                            self.a[num][row1, column] = re.sub(r'[^\w\s]', '', self.a[num][row1, column])
                            self.a[num][row1, column] = "".join(self.a[num][row1, column].split())
        return 0

    #计算平均负债率
    def loadrate(self):
        ration = 0.0
        #for循环，计算每个月的负债率
        for num in range(0, self.num):
            itemindex = np.argwhere(self.a[num] == '负债合计')
            b = itemindex[0][1] + 1
            c = itemindex[0][0]
            debt = (self.a[num][c, b])
            itemindex = np.argwhere(self.a[num] == '资产总计')
            b = itemindex[0][1] + 1
            c = itemindex[0][0]
            asset = (self.a[num][c, b])
            ration = ration + debt/asset
        ration = ration/self.num
        return ration

    def absolute1(self,var):
        asset = []
        for num in range(0, self.num):
            itemindex = np.argwhere(self.a[num] == var)
            b = itemindex[0][1] + 1
            c = itemindex[0][0]
            asset.append(var + str(round(self.a[num][c, b],2)))
        return asset

