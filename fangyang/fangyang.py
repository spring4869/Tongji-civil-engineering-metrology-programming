import sys
import xlrd
import math

file = 'data.xls'
f = open("res.txt","w+")#做笔记
sys.stdout = f

worksheet = xlrd.open_workbook(file).sheet_by_name('Sheet1')
list_a = []
list_b = []
list_p = []
list_dis = []
list_angle = []
cell_value_a = worksheet.cell_value(1,1)
list_a.append(cell_value_a)
cell_value_a = worksheet.cell_value(1,2)
list_a.append(cell_value_a)

cell_value_b = worksheet.cell_value(2,1)
list_b.append(cell_value_b)
cell_value_b = worksheet.cell_value(2,2)
list_b.append(cell_value_b)

for i in range(3,7,1):
    cell_value_c = worksheet.cell_value(i,1)
    list_p.append(cell_value_c)
    cell_value_c = worksheet.cell_value(i,2)
    list_p.append(cell_value_c)

angle_ab = math.atan((list_b[1]-list_a[1])/(list_b[0]-list_a[0])) * 180/math.pi
degree = int(angle_ab)
minute = (angle_ab - degree) * 60
second = round((minute - int(minute)) * 60)
minute = int((angle_ab - degree) * 60)

for i in range(0,7,2):
    angle_ap = math.atan((list_p[i+1] - list_a[1]) / (list_p[i] - list_a[0])) * 180 / math.pi
    delta_ap_x = pow(list_a[0] - list_p[i], 2)
    delta_ap_y = pow(list_a[1] - list_p[i+1], 2)
    dis_ap = pow(delta_ap_x + delta_ap_y, 0.5)
    if(angle_ap < 0 ):
        angle_ap += 360
    list_dis.append(round(dis_ap,3))
    angle_bap = angle_ap - angle_ab
    degree = int(angle_bap)
    minute = (angle_bap - degree) * 60
    second = round((minute - int(minute)) * 60)
    minute = int((angle_bap - degree) * 60)
    list_angle.append(str(degree) + "°" + str(minute) + "′" + str(second) + "″")
#以AB为轴，A为基点，逆时针为β的正方向
for i in range(0,4,1):
    print("AP"+ str(i) + "的距离S为:" + str(list_dis[i]) + "米")
    print("BAP" + str(i) + "角度β为:" + str(list_angle[i]))
print("注：以AB为轴，A为基点，逆时针为β的正方向")