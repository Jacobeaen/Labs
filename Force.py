import openpyxl
import matplotlib.pyplot as plt
import numpy as np


wb = openpyxl.load_workbook("table.xlsx")
sheet = wb.active

lst_x_1 = []
for i in range(30, 39):
    number = f"I{i}"
    cell = sheet[number]
    lst_x_1.append(cell.value)
lst_y_1 = []
for i in range(30, 39):
    number = f"G{i}"
    cell = sheet[number]
    lst_y_1.append(cell.value)

plt.xlim(0, 2.5)
plt.ylim(0, 2.5)
plt.grid()

plt.title("Зависимость прогиба от силы")
plt.xlabel("F, N")
plt.ylabel("λ, mm")

coefficients = np.polyfit(lst_x_1, lst_y_1, 1)
m, c = coefficients
y_fit = []
for i in range(len(lst_x_1)):
    y_fit.append(m * lst_x_1[i] + c)
plt.plot(lst_x_1, y_fit,)

plt.scatter(lst_x_1, lst_y_1)
plt.errorbar(lst_x_1, lst_y_1, xerr=0.02, yerr=0.03)
plt.show()

