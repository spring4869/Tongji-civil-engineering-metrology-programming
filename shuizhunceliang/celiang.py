import sys
import xlrd

file = 'data.xls'
f = open("res.txt","w+")#做笔记
sys.stdout = f

worksheet = xlrd.open_workbook(file).sheet_by_name('Table 1')
list_1 = []
list_2 = []
list_3 = []
list_4 = []
list_5 = []
list_6 = []
list_7 = []
list_8 = []
list_9 = []
list_10 = []
list_11 = []
list_12 = [0]
list_13 = []
list_14 = []
list_15 = []
list_16 = []
list_17 = []
list_18 = []
list_19 = []
list_20 = []
j = 0
sum_18 = 0.0
sum_20 = 0.0
ha = 42.002
for i in range(3,92,4):
    cell_value1 = worksheet.cell_value(i, 2)
    list_1.append(cell_value1)
    cell_value2 = worksheet.cell_value(i+1, 2)
    list_2.append(cell_value2)
    cell_value3 = worksheet.cell_value(i,7)
    list_3.append(cell_value3)
    cell_value4 = worksheet.cell_value(i,4)
    list_4.append(cell_value4)
    cell_value5 = worksheet.cell_value(i+1,4)
    list_5.append(cell_value5)
    cell_value6 = worksheet.cell_value(i+1,7)
    list_6.append(cell_value6)
    cell_value7 = worksheet.cell_value(i+1,8)
    list_7.append(cell_value7)
    cell_value8 = worksheet.cell_value(i,8)
    list_8.append(cell_value8)
    cell_value9 = abs(list_1[-1] - list_2[-1])/10
    list_9.append(cell_value9)
    cell_value10 = abs(list_4[-1] - list_5[-1])/10
    list_10.append(cell_value10)
    cell_value11 = list_9[-1] - list_10[-1]
    if(abs(cell_value11) < 5):
        {
            list_11.append(round(cell_value11,1))
        }
    cell_value12 = list_12[-1] + list_11[-1]
    list_12.append(round(cell_value12,1))
    cell_value13 = list_6[-1] - list_7[-1] + 4787
    list_13.append(cell_value13)
    cell_value14 = list_3[-1] - list_8[-1] + 4787
    list_14.append(cell_value14)
    cell_value15 = list_3[-1] - list_6[-1]
    list_15.append(cell_value15)
    cell_value16 = list_8[-1] - list_7[-1]
    list_16.append(cell_value16)
    cell_value17 = list_15[-1] - list_16[-1]
    list_17.append(cell_value17)
    cell_value18 = ((list_15[-1] + list_16[-1])/2)/1000
    list_18.append(round(cell_value18,4))
    cell_value20 = list_9[-1] + list_10[-1]
    list_20.append(cell_value20)

'''cell_value2 = worksheet.cell_value(3, 4)
cell_value3 = worksheet.cell_value(3, 5)
cell_value4 = worksheet.cell_value(3, 7)'''
del list_12[0]

for i in range(0,23,1):
    if(abs(list_9[i]) > 80.0 or abs(list_10[i]) > 80.0):
        print("不符合四等水准测量视线长度要求！")
        exit(-1)
    if(abs(list_11[i]) > 5):
        print("前后视距差超限！")
        exit(-1)
    if(abs(list_12[i]) > 10):
        print("视距累积差超限！")
        exit(-1)
    if(abs(list_17[i]) > 5):
        print("红黑面高差之差超限！")
        exit(-1)
    sum_18 += list_18[i]
    sum_20 += list_20[i]


print("四等水准测量视线长度符合要求！")
print("前后视距差符合精度要求！")
print("视距累积差符合精度要求！")
print("红黑面高差之差符合精度要求！")
print("")

print("水准路线全长为：%.1f"%sum_20)
print("高差闭合差为：%.4f"%sum_18)#精度
print("")

if(abs(sum_18) > 20 * pow(sum_20,0.5)):
    print("环线闭合差不符合精度标准！")
    exit(-1)

h = ha + list_18[0] + list_18[1]
print("B点改正前的高程为：%.3f"%h)
h += list_18[2] + list_18[3] + list_18[4]
print("C点改正前的高程为：%.3f"%h)
h += list_18[5] + list_18[6] + list_18[7] + list_18[8]
print("D点改正前的高程为：%.3f"%h)
h += list_18[9] + list_18[10] + list_18[11] + list_18[12]
print("E点改正前的高程为：%.3f"%h)
h += list_18[13] + list_18[14]
print("F点改正前的高程为：%.3f"%h)
h += list_18[15] + list_18[16] + list_18[17] + list_18[18] + list_18[19]
print("G点改正前的高程为：%.3f"%h)

average = -sum_18/sum_20
for i in range(0,23,1):
     cell_value19 = list_18[i] + average * list_20[i]#语法错误
     list_19.append(cell_value19)

print("")
h = ha + list_19[0] + list_19[1]
print("B点改正后的高程为：%.3f"%h)
h += list_19[2] + list_19[3] + list_19[4]
print("C点改正后的高程为：%.3f"%h)
h += list_19[5] + list_19[6] + list_19[7] + list_19[8]
print("D点改正后的高程为：%.3f"%h)
h += list_19[9] + list_19[10] + list_19[11] + list_19[12]
print("E点改正后的高程为：%.3f"%h)
h += list_19[13] + list_19[14]
print("F点改正后的高程为：%.3f"%h)
h += list_19[15] + list_19[16] + list_19[17] + list_19[18] + list_19[19]
print("G点改正后的高程为：%.3f"%h)

print("")
print("备注行列数据如下：")
print(list_1)
print(list_2)
print(list_3)
print(list_4)
print(list_5)
print(list_6)
print(list_7)
print(list_8)
print(list_9)
print(list_10)
print(list_11)
print(list_12)
print(list_13)
print(list_14)
print(list_15)
print(list_16)
print(list_17)
print(list_18)

