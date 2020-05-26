import math
import random
import statistics

import pandas as pd
import scipy.stats
import xlsxwriter

print('Welcome!')
print('This algorithm will calculate a new order point using Monte Carlo simulation method.')
Num = int(input('Please type the number of simulations to perform(100-10,000): ' + '\n'))
file_name = 'C:\App\File.xlsx'
df = pd.read_excel(file_name, sheet_name=0)

List_Current_Order_Level = []
List_New_Order_Level = []
List_Demands = []
for i in df.index:
    Catalog_Number = df['Catalog_Number'][i]
    Description = df['Description'][i]
    Price = df['Price'][i]
    Current_Order_Level = df['Current_Order_Level'][i]
    Stock = df['Stock'][i]
    Confidence = df['Confidence'][i]
    Lead_Time_Days = df['Lead_Time_Days'][i]
    Standard_deviation_Lead_Time = df['Standard_deviation_Lead_Time'][i]
    for j in range(1, 48):
        List_Demands.append(df[j][i])


    def reorder_point(avg, ltd, conf, sdv, slt):
        a = avg * ltd / 30
        b = scipy.stats.norm.ppf(conf)
        c = math.sqrt(pow(sdv, 2) * ltd / 30 + pow(avg, 2) * pow(slt, 2))
        ol = math.ceil(a + b * c)
        return ol


    def conversion(ltd):
        if ltd in range(0, 31):
            return 1
        elif ltd in range(31, 61):
            return 2
        elif ltd in range(61, 91):
            return 3
        elif ltd in range(91, 121):
            return 4
        elif ltd in range(121, 151):
            return 5
        elif ltd in range(151, 181):
            return 6
        elif ltd in range(181, 211):
            return 7
        elif ltd in range(211, 241):
            return 8
        else:
            return 9


    def one_step_forward(arr, n, order):
        for i in range(n, -1, -1):
            arr[i] = arr[i - 1]
            i -= 1
        arr[0] = order
        return arr


    def fill(ary, ol, stock):
        sum_of_array = sum(ary)
        orders = ol - (sum_of_array + stock)
        if orders > 0:
            return orders
        else:
            return 0


    def monte_carlo(list1, number, stock, order_level, ltm):
        ltm = ltm - 1
        x = 0
        list2 = []
        for i in range(0, ltm + 1):
            list2.append(0)
        availability = 0
        total_stock = stock
        while x < number:
            demand = random.choice(list1)
            if stock >= 0:
                availability += 1
                stock -= demand
                orders = fill(list2, order_level, stock)
            else:
                stock = 0
                orders = fill(list2, order_level, stock)
            stock += list2[ltm]
            total_stock += stock
            one_step_forward(list2, ltm, orders)
            x += 1
        overall_availability = availability / number
        return overall_availability


    def result(check, confidence, order_level):
        if check < confidence:
            while check < confidence:
                order_level += 1
                check = monte_carlo(List_Demands, Num, Stock, order_level, Lead_Time_Months)
            List_New_Order_Level.append(order_level)
        else:
            List_New_Order_Level.append(order_level)


    List_Current_Order_Level.append(Current_Order_Level)
    Avg = sum(List_Demands) / len(List_Demands)
    Standard_deviation = statistics.stdev(List_Demands)
    Lead_Time_Months = conversion(Lead_Time_Days)
    Order_Level = reorder_point(Avg, Lead_Time_Days, Confidence, Standard_deviation, Standard_deviation_Lead_Time)
    #Num = 1000

    Check = monte_carlo(List_Demands, Num, Stock, Order_Level, Lead_Time_Months)

    result(Check, Confidence, Order_Level)

    List_Demands = []
    """ reset """

workbook = xlsxwriter.Workbook('C:\App\Result.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, 'Current_Order_Level')
worksheet.write(0, 1, 'New_Order_Level')

row = 1
column = 0
for item in List_Current_Order_Level:
    worksheet.write(row, column, item)
    row += 1

row = 1
column = 1
for item in List_New_Order_Level:
    worksheet.write(row, column, item)
    row += 1

workbook.close()

print('Done!')
print('A new Excel file was created on your computer')
print('File Name: Result.xlsx')
print('You can access it in the following path: "C:\App"')
input()