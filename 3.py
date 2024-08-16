from openpyxl import *
import os
import psutil
import function
from tkinter import messagebox

while True:
    n = input('Enter the dimension of the square matrix: ')
    if n.isdigit() and n != '0' and n != '1':
        n = int(n)
        break
    else:
        print('The number is entered incorrectly!')
i, fg = 1, False
directory = os.getcwd()
matr = []
name_file = input('Enter the file name: ')
while not fg:
    num = matr
    function.close_excel()
    book = Workbook()
    book.remove(book['Sheet'])
    worksheet = book.create_sheet('Data entry')
    worksheet.append([x if x != 0 else '' for x in range(n+1)])
    for j in range(n):
        mas = [j+1]
        mas.extend([y for i in range(len(num)) for y in num[j] if i == j and len(num) != 0])
        worksheet.append(mas)
    function.close_excel()
    name_file = function.saving_excel(book, name_file, f'({i})')
    book.close()
    os.startfile(directory + '/' + name_file + f'({i}).xlsx')
    while 'EXCEL.EXE' in (m.name() for m in psutil.process_iter()):
        pass
    book = load_workbook(name_file + f'({i}).xlsx')
    function.removing_sheets(book, 'Data entry')
    num = function.reading_data(book, n)
    num, fg = function.data_checking(num)
    i += 1
    if not fg:
        error = set(j + 1 for j in range(len(num)) for y in num[j] if y is None)
        if len(error) > 1:
            messagebox.showerror(title='Filling ERROR', message=f'The error is in the following lines: {str(error)[1:-1]}')
        else:
            messagebox.showerror(title='Filling ERROR', message=f'The error is in the next line: {str(error)[1:-1]}')
    matr = num
matrix = []
for i in range(len(matr)):
    line = []
    for j in range(len(matr)):
        line.append(function.transformation(matr[i][j], i, j, matr))
    matrix.append(line)
book = Workbook()
book.remove(book['Sheet'])
function.data_output(book, 'Data entry', 'The original matrix', matr, n)
function.data_output(book, 'Data output', 'The transformed matrix', matrix, n)
name_file = function.saving_excel(book, '3.xlsx', '')
os.startfile(directory + '/3.xlsx')
