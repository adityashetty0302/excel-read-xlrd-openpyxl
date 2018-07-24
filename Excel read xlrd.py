import os
import xlrd

if __name__ == '__main__':

    os.chdir("C:\\Users\\Aditya.Shetty\\Desktop")
    # print(os.listdir('.'))

    wb = xlrd.open_workbook('t1.xls')
    # print(wb.sheet_names())

    sheet = wb.sheet_by_name('Sheet1')
    print(sheet.name)
    columndata = sheet.col_values(0)
    print(len(columndata))
    rowdata = sheet.row_values(0)
    print(len(rowdata), "\n")

    for i in range(0, len(columndata)):
        with open("Outputxlrd.txt", "a") as text_file:
            for j in range(0, len(rowdata)):
                print(sheet.cell(i, j).value, "\t")
                text_file.write(str(sheet.cell(i, j).value))
                text_file.write("\t")
            print("\n")
            text_file.write("\n")
