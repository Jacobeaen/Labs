import openpyxl
import matplotlib.pyplot as plt


wb = openpyxl.load_workbook("table.xlsx")
sheet = wb.active

# №1
lst_x_1 = []
for i in range(2, 57):
    number = f"A{i}"
    cell = sheet[number]
    lst_x_1.append(cell.value)
lst_y_1 = []
for i in range(2, 57):
    number = f"B{i}"
    cell = sheet[number]
    lst_y_1.append(cell.value - 1.3)

# №2
lst_x_2 = []
for i in range(57, 117):
    number = f"A{i}"
    cell = sheet[number]
    lst_x_2.append(cell.value)
lst_y_2 = []
for i in range(57, 117):
    number = f"B{i}"
    cell = sheet[number]
    lst_y_2.append(cell.value - 1.3)

# №3
lst_x_3 = []
for i in range(117, 149):
    number = f"A{i}"
    cell = sheet[number]
    lst_x_3.append(cell.value)
lst_y_3 = []
for i in range(117, 149):
    number = f"B{i}"
    cell = sheet[number]
    lst_y_3.append(cell.value - 1.3)

# №4
lst_x_4 = []
for i in range(149, 180):
    number = f"A{i}"
    cell = sheet[number]
    lst_x_4.append(cell.value)
lst_y_4 = []
for i in range(149, 180):
    number = f"B{i}"
    cell = sheet[number]
    lst_y_4.append(cell.value - 1.3)


lst_x =list(map(lambda x: sum(x) / 4, zip(lst_x_1, lst_x_2, lst_x_3, lst_x_4)))
lst_y = list(map(lambda y: sum(y) / 4, zip(lst_y_1, lst_y_2, lst_y_3, lst_x_4)))


plt.xlim(0, 10)
plt.ylim(0, 2.5)
plt.grid()

plt.title("Характеристическая кривая")
plt.xlabel("x, mm")
plt.ylabel("Fr, N")

plt.plot(lst_x, lst_y, color="red", label="Среднее из массивов")

plt.legend(loc = "upper left")

plt.show()
