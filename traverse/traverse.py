import xlrd
import random
import math
import sys
file = 'data.xls'
f = open("res.txt","w+")#做笔记
sys.stdout = f

worksheet = xlrd.open_workbook(file).sheet_by_name('Table 1')
list_leveldata = []
list_level = []
list_levelaver = []
list_levelout = []
list_levelout1 = []
list_beta = []
list_dis = []
list_alpha = []
list_delta_x = []
list_delta_y = []
list_delta_xx = []
list_delta_yy = []
list_X = []
list_Y = []
sum_beta = 0
sum_dis = 0.0
alpha_ab = 100 * 3600
x_a = 285.024
y_a = 198.471
for i in range(2,22,1):
    cell_values = worksheet.cell_value(i,4) * 3600 + worksheet.cell_value(i,6) * 60 + worksheet.cell_value(i,7) + random.randint(-1,1)
    list_leveldata.append(cell_values)
for i in range(0,20,2):
    cell_values = list_leveldata[i + 1] - list_leveldata[i]
    list_level.append(cell_values)
for i in range(0,10,2):
    cell_values = list_level[i + 1] - list_level[i]
    if(cell_values > 40):
        print("盘左盘右角度差误差超限！")
        exit(-1)
    else:
        cell_values = (list_level[i + 1]  +  list_level[i])/2
        list_levelaver.append(round(cell_values))
for i in range(0,5,1):
    sum_beta += list_levelaver[i]
    degree = int(list_levelaver[i] / 3600)
    second = list_levelaver[i] % 60
    minute = int( ( list_levelaver[i] - degree * 3600 ) / 60 )
    list_levelout.append(str(degree) + "°" + str(minute) + "′" + str(second) + "″")

f_beta = sum_beta - 540 * 3600
if(f_beta > 60 * pow(5,0.5)):
    print("角度闭合差超限！")
    exit(-1)
for i in range(0,5,1):
    correction = list_levelaver[i] - f_beta / 5
    list_beta.append(round(correction))

for i in range(0,5,1):
    degree = int(list_beta[i] / 3600)
    second = list_beta[i] % 60
    minute = int((list_beta[i] - degree * 3600) / 60)
    list_levelout1.append(str(degree) + "°" + str(minute) + "′" + str(second) + "″")

for i in range(0,5,1):
    cell_values = worksheet.cell_value(i * 4 + 11,17) - worksheet.cell_value(i* 4 + 8,17)
    if(cell_values > 0.1):
        print("边长距离超限！")
        exit(-1)
    else:
        for j in range(0,5,1):
            cell_values = (worksheet.cell_value(j * 4 + 8,16) + worksheet.cell_value(j * 4 + 9,16) + worksheet.cell_value(j * 4 + 10,16) + worksheet.cell_value(j * 4 + 11,16))/4
            list_dis.append(round(cell_values,2))
        break
for i in range(0,5,1):
    sum_dis += list_dis[i]

alpha_bc = alpha_ab + 180 * 3600 - list_beta[0]
list_alpha.append(round(alpha_bc,3))
alpha_cd = alpha_bc + 180 * 3600 - list_beta[1]
list_alpha.append(round(alpha_cd,3))
alpha_de = alpha_cd + 180 * 3600 - list_beta[2]
list_alpha.append(round(alpha_de,3))
alpha_ea = alpha_de + 180 * 3600 - list_beta[3]
list_alpha.append(round(alpha_ea,3))
list_alpha.append(alpha_ab)

delta_x = list_dis[0] * math.cos(alpha_ab*math.pi/648000)
list_delta_x.append(delta_x)
delta_x = list_dis[1] * math.cos(alpha_bc*math.pi/648000)
list_delta_x.append(delta_x)
delta_x = list_dis[2] * math.cos(alpha_cd*math.pi/648000)
list_delta_x.append(delta_x)
delta_x = list_dis[3] * math.cos(alpha_de*math.pi/648000)
list_delta_x.append(delta_x)
delta_x = list_dis[4] * math.cos(alpha_ea*math.pi/648000)
list_delta_x.append(delta_x)

delta_y = list_dis[0] * math.sin(alpha_ab*math.pi/648000)
list_delta_y.append(delta_y)
delta_y = list_dis[1] * math.sin(alpha_bc*math.pi/648000)
list_delta_y.append(delta_y)
delta_y = list_dis[2] * math.sin(alpha_cd*math.pi/648000)
list_delta_y.append(delta_y)
delta_y = list_dis[3] * math.sin(alpha_de*math.pi/648000)
list_delta_y.append(delta_y)
delta_y = list_dis[4] * math.sin(alpha_ea*math.pi/648000)
list_delta_y.append(delta_y)

fx = sum(list_delta_x)
fy = sum(list_delta_y)
f = pow(pow(fx,2) + pow(fy,2),0.5)
if(f / sum_dis > 1/4000.0):
    print("导线相对闭合差超限！")
    exit(-1)

for i in range(0,5,1):
    list_delta_xx.append(round(list_delta_x[i] - fx/sum_dis * list_dis[i],2))
    list_delta_yy.append(round(list_delta_y[i] - fy/sum_dis * list_dis[i],2))

print("改正前A的平均角值为：" + list_levelout[4])
print("改正前B的平均角值为：" + list_levelout[0])
print("改正前C的平均角值为：" + list_levelout[1])
print("改正前D的平均角值为：" + list_levelout[2])
print("改正前E的平均角值为：" + list_levelout[3])
print("")

print("角度闭合差为：" + str(f_beta) + "″")

print("")

print("改正后A的平均角值为：" + str(list_levelout1[4]))
print("改正后B的平均角值为：" + str(list_levelout1[0]))
print("改正后C的平均角值为：" + str(list_levelout1[1]))
print("改正后D的平均角值为：" + str(list_levelout1[2]))
print("改正后E的平均角值为：" + str(list_levelout1[3]))

print("")

print("AB的距离均值为：" + str(list_dis[0]) + "m")
print("BC的距离均值为：" + str(list_dis[1]) + "m")
print("CD的距离均值为：" + str(list_dis[2]) + "m")
print("DE的距离均值为：" + str(list_dis[3]) + "m")
print("EA的距离均值为：" + str(list_dis[4]) + "m")

print("")

print("改正后AB的x坐标增量为：" + str(list_delta_xx[0]) + "m")
print("改正后AB的y坐标增量为：" + str(list_delta_yy[0]) + "m")
print("改正后BC的x坐标增量为：" + str(list_delta_xx[1]) + "m")
print("改正后BC的y坐标增量为：" + str(list_delta_yy[1]) + "m")
print("改正后CD的x坐标增量为：" + str(list_delta_xx[2]) + "m")
print("改正后CD的y坐标增量为：" + str(list_delta_yy[2]) + "m")
print("改正后DE的x坐标增量为：" + str(list_delta_xx[3]) + "m")
print("改正后DE的y坐标增量为：" + str(list_delta_yy[3]) + "m")
print("改正后EA的x坐标增量为：" + str(list_delta_xx[4]) + "m")
print("改正后EA的y坐标增量为：" + str(list_delta_yy[4]) + "m")

print("")

print("导线相对闭合差为：" + str(round(f,3)) + "m")

print("")

list_X.append(x_a)
list_Y.append(y_a)
for i in range(0,4,1):
    list_Y.append(round((list_Y[-1] + list_delta_yy[i]),3))#biji!!!
    list_X.append(round((list_X[-1] + list_delta_xx[i]),3))
print("A点坐标为：" + "(" + str(list_X[0]) + "m," + str(list_Y[0]) + "m)")
print("B点坐标为：" + "(" + str(list_X[1]) + "m," + str(list_Y[1]) + "m)")
print("C点坐标为：" + "(" + str(list_X[2]) + "m," + str(list_Y[2]) + "m)")
print("D点坐标为：" + "(" + str(list_X[3]) + "m," + str(list_Y[3]) + "m)")
print("E点坐标为：" + "(" + str(list_X[4]) + "m," + str(list_Y[4]) + "m)")

print("")

print("盘左盘右角度差误差符合精度要求！")
print("角度闭合差符合精度要求！")
print("导线相对闭合差符合精度要求!")
print("边长距离符合精度要求！")
print("导线相对闭合差符合精度要求！")
