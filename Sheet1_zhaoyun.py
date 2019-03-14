import xlrd
import math
##"""Sheet1_test 获取表格一的数据列表，其中包括开单日期，商品名称，数量，单价以及总金额"""
class Sheet1_test:
    def __init__(self,Execel_name,Sheet_name,Text_date,Text_time):
        self.list_date = []
        self.list_time = []
        self.shijian=[]
        self.list_test1=[]
        self.data = xlrd.open_workbook(Execel_name)  # 打开文件
        self.sheet_name = self.data.sheet_by_name(Sheet_name)
        with open(Text_date) as read_file:
            for line in read_file:
                self.list_date.append(line.rstrip())
        with open(Text_time) as read_file:
            for line in read_file:
                self.list_time.append(line.rstrip())


        for i in range(len(self.list_time)):
            self.shijian.append(self.list_date[i] + self.list_time[i])


    def get_list1(self):
        zy_nrows = self.sheet_name.nrows  # 行
        zy_ncols = self.sheet_name.ncols  # 列
        for i in range(zy_nrows):
            if i >= 3:
                test1_list = [self.shijian[i - 3], self.sheet_name.cell(i, 5).value, self.sheet_name.cell(i, 6).value,\
                                self.sheet_name.cell(i,7).value, self.sheet_name.cell(i, 8).value]
                self.list_test1.append(test1_list)
        return self.list_test1

##"""Sheet2_test 获取表格二的数据列表，其中包括开单日期，商品名称，数量，单价以及总金额"""
class Sheet2_test(Sheet1_test):
    def get_list2(self):
        zy_nrows = self.sheet_name.nrows  # 行
        zy_ncols = self.sheet_name.ncols  # 列
        for i in range(zy_nrows-1):
            if i >= 5:
                test1_list= [ self.shijian[i-5], self.sheet_name.cell(i, 9).value, self.sheet_name.cell(i, 11).value, \
                               self.sheet_name.cell(i, 12).value, self.sheet_name.cell(i, 13).value]
                self.list_test1.append(test1_list)

        return self.list_test1
##########获取表格没有问题############
class Merge:
    def __init__(self,list_name):
        self.list_product=[]
        self.list_set=[]
        self.list_sheet=list_name.copy()
        for i in range(len(list_name)):
            self.list_product.append(list_name[i][1])
        for i in set(self.list_product):
            if self.list_product.count(i)>=2:
                self.list_set.append(i)
    def merge_product(self):
        for a in self.list_set:
            fisrt_postion=0
            for b in range(self.list_product.count(a)):
                fisrt_postion =self.list_product.index(a,fisrt_postion)
                try:
                    next_postion = self.list_product.index(a, fisrt_postion+1)
                    time_dif2=Time_abs(self.list_sheet[fisrt_postion][0],self.list_sheet[next_postion][0]).time_change()

                except ValueError:
                    break
                if math.fabs(time_dif2)<=12:
                    self.list_sheet[fisrt_postion][2]+=self.list_sheet[next_postion][2]
                    self.list_sheet[fisrt_postion][4]+= self.list_sheet[next_postion][4]
                    del self.list_sheet[next_postion]
                    del self.list_product[next_postion]
                else:
                    fisrt_postion=next_postion

        return self.list_sheet
###整合表格###

class Compare_sheet:
    def __init__(self,list_name1,list_name2):
        self.list_sale1=[]
        self.list_sale2 = []
        self.list_compare=[]
        self.list_sheet1=list_name1.copy()
        self.list_sheet2=list_name2.copy()

        for i in range(len(list_name1)):
            self.list_sale1.append(list_name1[i][1])
        for i in range(len(list_name2)):
            self.list_sale2.append(list_name2[i][1])

    def chayi(self):
        for i in range(len(self.list_sale2)):
            a = self.list_sale1.count(self.list_sale2[i])
            if a == 0:
                zero_sale = [self.list_sheet2[i], '此商品在表1没有，在表二的索引位置为%d' % i]
                self.list_compare.append(zero_sale)
            elif a == 1:
                sale1_postion = self.list_sale1.index(self.list_sale2[i])
                date = Time_abs(self.list_sheet2[i][0],self.list_sheet1[sale1_postion][0]).time_change()

                pay_chayi = self.list_sheet2[i][4] - self.list_sheet1[sale1_postion][4]
                if math.fabs(date)<=12 and pay_chayi != 0:
                    zero_sale = [self.list_sheet2[i], self.list_sheet1[sale1_postion], pay_chayi]
                    self.list_compare.append(zero_sale)
                elif math.fabs(date)>12:
                    zero_sale = [self.list_sheet2[i], '此商品在表1当天没有入库记账，索引位置为%d' % i]
                    self.list_compare.append(zero_sale)
            else:
                first_postion=0

                for ii in range(a):
                    sale1_postion = self.list_sale1.index(self.list_sale2[i],first_postion)
                    date = Time_abs(self.list_sheet2[i][0], self.list_sheet1[sale1_postion][0]).time_change()
                    pay_chayi = self.list_sheet2[i][4] - self.list_sheet1[sale1_postion][4]
                    if math.fabs(date) <=12 and pay_chayi==0:
                        break
                    if math.fabs(date) <=12 and pay_chayi != 0:
                        zero_sale = [self.list_sheet2[i], self.list_sheet1[sale1_postion], pay_chayi]
                        self.list_compare.append(zero_sale)
                        break
                    first_postion=sale1_postion
                    if ii==a:
                        zero_sale = [self.list_sheet2[i], '此商品在表1当天没有入库记账，索引位置为%d' % i]
                        self.list_compare.append(zero_sale)
        return self.list_compare

class Time_abs():
    def __init__(self,time1,time2):
        Time_unit = ['年', '月', '日', '时']
        self.list1 = []
        self.list2 = []
        self.time_dif=[]
        for i in Time_unit:
            a,b = time1.split(i)
            self.list1.append(float(a))
            time1=b
        for i in Time_unit:
            a, b = time2.split(i)
            self.list2.append(float(a))
            time2 = b

        for i in range(len(self.list1)):
            a=float(self.list1[i])-float(self.list2[i])
            self.time_dif.append(a)
    def time_change(self):
        jingzhi=[12,31,24]
        for i in range(len(self.time_dif)-1):
            self.time_dif[i+1]+=self.time_dif[i]*jingzhi[i]
        return self.time_dif[-1]

path='E:\Python_learning\python1\zhaoyun'
sheet1_zhaoyun=Sheet1_test(path+'\华工(1)测试_赵云.xlsx','1.1-1.15明细',path+'\测试日期.txt',path+'\测试时间.txt')
list1=sheet1_zhaoyun.get_list1()
sheet2_zhaoyun=Sheet2_test(path+'\华工1月核对明细(1)测试_赵云.xls','报表',path+'\核对明细日期.txt',path+'\核对明细时间.txt')
list2=sheet2_zhaoyun.get_list2()
list_merge_sheet2=Merge(list2).merge_product()
list_merge_sheet1=Merge(list1).merge_product()
list_compare=Compare_sheet(list_merge_sheet1,list_merge_sheet2)
sheet3=list_compare.chayi()