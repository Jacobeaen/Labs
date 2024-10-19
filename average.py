import openpyxl
import matplotlib.pyplot as plt


wb = openpyxl.load_workbook("table.xlsx")
sheet = wb.active




# №1
lst_x_1 = []
for i in range(2, 57):
    number = f"A{i}"
    cell = sheet[number]
    lst_x_1.append(10-cell.value)
lst_y_1 = []
for i in range(2, 57):
    number = f"B{i}"
    cell = sheet[number]
    lst_y_1.append(cell.value - 1.03)

lst_Y1 = []
lst_X1 = []
for i in range(len(lst_y_1)-1):
    a = int((lst_y_1[i+1] - lst_y_1[i]) * 100)
    b = int((lst_x_1[i+1] - lst_x_1[i]) * 100)
    for g in range(b):
        lst_Y1.append(lst_y_1[i] + (g * a)/(b * 100))
        lst_X1.append(round(lst_x_1[i] + g/100, 2))

# №2
lst_x_2 = []
for i in range(57, 117):
    number = f"A{i}"
    cell = sheet[number]
    lst_x_2.append(10-cell.value)
lst_y_2 = []
for i in range(57, 117):
    number = f"B{i}"
    cell = sheet[number]
    lst_y_2.append(cell.value - 1.03)


lst_Y2 = []
lst_X2 = []
for i in range(len(lst_y_2)-1):
    a = int((lst_y_2[i+1] - lst_y_2[i]) * 100)
    b = int((lst_x_2[i+1] - lst_x_2[i]) * 100)
    for g in range(b):
        lst_Y2.append(lst_y_2[i] + (g * a)/(b * 100))
        lst_X2.append(round(lst_x_2[i] + g/100, 2))

# №3
lst_x_3 = []
for i in range(117, 149):
    number = f"A{i}"
    cell = sheet[number]
    lst_x_3.append(10-cell.value)
lst_y_3 = []
for i in range(117, 149):
    number = f"B{i}"
    cell = sheet[number]
    lst_y_3.append(cell.value - 1.03)


lst_Y3 = []
lst_X3 = []
for i in range(len(lst_y_3)-1):
    a = int((lst_y_3[i+1] - lst_y_3[i]) * 100)
    b = int((lst_x_3[i+1] - lst_x_3[i]) * 100)
    for g in range(b):
        lst_Y3.append(lst_y_3[i] + (g * a)/(b * 100))
        lst_X3.append(round(lst_x_3[i] + g/100, 2))


# №4
lst_x_4 = []
for i in range(149, 180):
    number = f"A{i}"
    cell = sheet[number]
    lst_x_4.append(10-cell.value)
lst_y_4 = []
for i in range(149, 180):
    number = f"B{i}"
    cell = sheet[number]
    lst_y_4.append(cell.value - 1.03)


lst_Y4 = []
lst_X4 = []
for i in range(len(lst_y_4)-1):
    a = int((lst_y_4[i+1] - lst_y_4[i]) * 100)
    b = int((lst_x_4[i+1] - lst_x_4[i]) * 100)
    for g in range(b):
        lst_Y4.append(lst_y_4[i] + (g * a)/(b * 100))
        lst_X4.append(round(lst_x_4[i] + g/100, 2))


lst_y_a = []
lst_x_a = []

for i in range(1000):
    if lst_X1.count(i/100) > 0 and lst_X2.count(i/100) > 0 and lst_X3.count(i/100) > 0 and lst_X4.count(i/100) > 0:
        lst_y_a.append((lst_Y1[lst_X1.index(i/100)] + lst_Y2[lst_X2.index(i/100)] + lst_Y3[lst_X3.index(i/100)] + lst_Y4[lst_X4.index(i/100)])/4)
        lst_x_a.append(lst_X1[lst_X1.index(i/100)])
        a = 1
    # elif lst_X1.count(i/100) == 1 and lst_X2.count(i/100) == 1 and lst_X3.count(i/100) == 1 and lst_X4.count(i/100) == 0:
    #     lst_y_a.append((lst_Y1[lst_X1.index(i/100)] + lst_Y2[lst_X2.index(i/100)] + lst_Y3[lst_X3.index(i/100)])/3)
    #     lst_x_a.append(lst_X1[lst_X1.index(i/100)])
    # elif lst_X1.count(i/100) > 0 and lst_X2.count(i/100) > 0 and lst_X3.count(i/100) == 0 and lst_X4.count(i/100) > 0:
    #     lst_y_a.append((lst_Y1[lst_X1.index(i/100)] + lst_Y2[lst_X2.index(i/100)] + lst_Y4[lst_X4.index(i/100)])/3)
    #     lst_x_a.append(lst_X1[lst_X1.index(i/100)])
    # elif lst_X1.count(i/100) > 0 and lst_X2.count(i/100) == 0 and lst_X3.count(i/100) > 0 and lst_X4.count(i/100) > 0:
    #     lst_y_a.append((lst_Y1[lst_X1.index(i/100)] + lst_Y3[lst_X3.index(i/100)] + lst_Y4[lst_X4.index(i/100)])/3)
    #     lst_x_a.append(lst_X1[lst_X1.index(i/100)])
    # elif lst_X1.count(i/100) == 0 and lst_X2.count(i/100) > 0 and lst_X3.count(i/100) > 0 and lst_X4.count(i/100) > 0:
    #     lst_y_a.append((lst_Y2[lst_X2.index(i/100)] + lst_Y3[lst_X3.index(i/100)] + lst_Y4[lst_X4.index(i/100)])/3)
    #     lst_x_a.append(lst_X2[lst_X2.index(i/100)])

lst_x = []
for x1, x2, x3, x4 in zip(lst_x_1, lst_x_2, lst_x_3, lst_x_4):
    x = (x1 + x2 + x3 + x4) / 4
    lst_x.append(x)

lst_y = []
for y1, y2, y3, y4 in zip(lst_y_1, lst_y_2, lst_y_3, lst_y_4):
    y = (y1 + y2 + y3 + y4) / 4
    lst_y.append(y)

#lst_x =list(map(lambda x: sum(x) / 4, zip(lst_x_1, lst_x_2, lst_x_3, lst_x_4)))
#lst_y = list(map(lambda y: sum(y) / 4, zip(lst_y_1, lst_y_2, lst_y_3, lst_x_4)))


# plt.xlim(0, 10)
# plt.ylim(0, 1)
# plt.grid()

# plt.title("Характеристическая кривая")
# plt.xlabel("x, mm")
# plt.ylabel("Fr, N")

#plt.plot(lst_x, lst_y, color="black", label="Среднее из массивов")
# plt.plot(lst_x_1, lst_y_1, color="red", label="FN1")
# plt.plot(lst_x_2, lst_y_2, color="blue", label="FN2")
# plt.plot(lst_x_3, lst_y_3, color="orange", label="FN3")
# plt.plot(lst_x_4, lst_y_4, color="green", label="FN4")

# plt.plot(lst_X1, lst_Y1, color="brown", label="FN1_a")
# plt.plot(lst_X2, lst_Y2, color="brown", label="FN2_a")
# plt.plot(lst_X3, lst_Y3, color="brown", label="FN3_a")
# plt.plot(lst_X4, lst_Y4, color="brown", label="FN4_a")
# plt.plot(lst_x_a, lst_y_a, color="black", label="Fr a")

# plt.legend(loc = "lower left")

# plt.show()

print(lst_y_a[lst_x_a.index(4.75)])