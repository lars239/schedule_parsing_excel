import xlrd
import csv

'''
Функция определяет относиться ли предмет к чётной 
или нечётной неделе. На входе строка,
чётный предмет, нечётный, ощий
'''
def chek_parity(line:str) -> str:
    obj = line.strip(' ')
    if '/' == obj[0]:
        return 'чёт'
    if '/' == obj[len(obj) - 1]:
        return 'нечёт'  

    return 'общ'


db_set = [] # список содержит номер пары, группа, название предмета, имя преподавателя, день недели, определитель чётности
lst_weekday = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
xls_file = xlrd.open_workbook('./bd_semestr.xlsx')
faculty_name_lst = xls_file.sheet_names()

for sheet_num in range(len(faculty_name_lst)):
    sheet = xls_file.sheet_by_index(sheet_num)
    faculty_name = faculty_name_lst[sheet_num]
    for col_num in range(0, sheet.ncols): # преклюение по колонкам
        weekday = 0
        num_couple = 0 
        
        for record in range(1 , len(sheet.col(col_num)), 2): # переключение по строкам          
            group = sheet.col(col_num)[0].value.strip(' ').replace(' ', '')
            ticher = sheet.col(col_num)[record+1].value.strip(' ').replace(' ', '')
            subject = sheet.col(col_num)[record].value.strip(' ').replace(' ', '')
            num_couple += 1
            type_cell = sheet.col(col_num)[record].ctype
            if type_cell == 1:
                if ticher[0] == '/' and subject[len(subject) - 1] == '/':
                    db_set.append([str(num_couple), faculty_name, group, subject, lst_weekday[weekday], 'нечёт'])
                    db_set.append([str(num_couple), faculty_name, group, ticher, lst_weekday[weekday], 'чёт'])
                else:
                    db_set.append([str(num_couple), faculty_name, group, subject + ' ' + ticher, lst_weekday[weekday], chek_parity(subject + ' ' + ticher)])
            if num_couple > 5:
                num_couple = 0
                weekday += 1

with open('eggs.csv', 'w', newline='') as csvfile:
    spamwriter = csv.writer(csvfile, delimiter=',',
                            quotechar=' ', quoting=csv.QUOTE_MINIMAL)
    for i in db_set:
        spamwriter.writerow(i) 