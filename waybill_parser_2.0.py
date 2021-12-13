"""
waybill.xls - file that contains only master's current positions to buy AND to create 10ml
"""


import xlrd
import xlwt

wb = xlrd.open_workbook('waybill.xls', formatting_info=True)
sheet = wb.sheet_by_index(0)  # Возвращает лист книги по индексу


crude_list_of_rows = []   # список [номер заказа системы вида *C, [список всех ароматов данного закакза]]

# start of creating crude_list_of_rows
for rows_number in range(sheet.nrows):   # перебираем все строки от самой первой до их количества (sheet.nrows)
    current_row = sheet.row_values(rows_number)
    # print(current_row, type(current_row[0]), type(current_row[1]))

    crude_list_of_rows.append([current_row[0], str(current_row[1]).split(';')])
del crude_list_of_rows[0]  # удаляем шапку таблицы, которая занеслась в список
# end of creating crude_list_of_rows

#

cleaned_list_of_rows = []

# start of creating cleaned_list_of_rows - список [номер заказа системы вида *C, один из ароматов этого заказа]
for orders in range(len(crude_list_of_rows)):
    for container in range(len(crude_list_of_rows[orders][1])):
        cleaned_list_of_rows.append([crude_list_of_rows[orders][0], crude_list_of_rows[orders][1][container]])

for cur in range(len(cleaned_list_of_rows)):
    current_name = cleaned_list_of_rows[cur][1]  # название текущего аромата
    if current_name[0] == ' ':    # удаляем первый пробел
        current_name = current_name[1:]

    # print(current_name, '~before other editings')

    indx_rspace = str(current_name).rfind(' ')
    indx_mspace = str(current_name[0:indx_rspace]).rfind(' ')
    current_name_amount = int(current_name[indx_mspace+1:indx_rspace])   # неявно вычисляем количество

    indx_dash = str(current_name).find('-')
    current_name = current_name[0:indx_dash-1]   # чистим название

    cleaned_list_of_rows[cur] = [cleaned_list_of_rows[cur][0], current_name, current_name_amount]
    # print(cleaned_list_of_rows[cur])
# the end: cleaned_list_of_rows - список [номер заказа системы вида *C, аромат из этого заказа, количество]

# start

# for cur in range(len(cleaned_list_of_rows)):
#     print(cleaned_list_of_rows[cur])
final_list_of_rows = []

for cur in range(len(cleaned_list_of_rows)):
    cleaned_list_of_rows[cur].append('NT')   #  Not Taken - не внесли в окончательный лист уникальных ароматов

for cur in range(len(cleaned_list_of_rows)):
    search = cleaned_list_of_rows[cur]
    if search[3] == 'T':
        continue

    cleaned_list_of_rows[cur][3] = 'T'

    for ifdupl in range(cur+1, len(cleaned_list_of_rows)):
        if cleaned_list_of_rows[ifdupl][1] == search[1] and cleaned_list_of_rows[ifdupl][3] == 'NT':
            search[2] += cleaned_list_of_rows[ifdupl][2]
            search[0] = search[0] + ', ' + cleaned_list_of_rows[ifdupl][0]
            cleaned_list_of_rows[ifdupl][3] = 'T'

    final_list_of_rows.append(search)
# final_list_of_rows.sort()
# end


# start of sorting final list
for i in range(len(final_list_of_rows)):
    if final_list_of_rows[i][1].find('Скидочная карта') == 0:
        del final_list_of_rows[i]
        break

for cur in range(0, len(final_list_of_rows)):
    search = final_list_of_rows[cur]
    if search[1].find('10 ') >= 0 or search[1].find('"') >= 0 or search[1].find('10ml') >= 0 or \
            search[1].find('Подарочный карандаш') >= 0 or search[1].find('Пробник любой') >= 0 or \
            search[1].find('Пробник аромата (любой)') >= 0:
        final_list_of_rows[cur][3] = 'NT'   # не берём позицию как закупную, рассматриваем для СБ


# end of sorting final list (separating 10ml to the RIGHT and FULLml to the LEFT)


font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.colour_index = 0   # black
font0.bold = False  # жирный шрифт: да/нет

style0 = xlwt.XFStyle()  # стиль usual
style0.font = font0

font1 = xlwt.Font()
font1.name = 'Times New Roman'
font1.colour_index = 0   # black
font1.bold = True  # жирный шрифт: да/нет

style1 = xlwt.XFStyle()  # стиль для названий
style1.font = font1


workbook = xlwt.Workbook()
worksheet_s = workbook.add_sheet('SADOVOD')
# worksheet_r = workbook.add_sheet('RASPIVS')

worksheet_s.write(0, 0, 'Номер заказа', style1)
worksheet_s.write(0, 1, 'Наименование', style1)
worksheet_s.write(0, 2, 'Количество', style1)

line_left = 1
line_right = 1
for cur in range(len(final_list_of_rows)):
    if final_list_of_rows[cur][3] == 'T':
        for column in range(0, len(final_list_of_rows[cur])-1):
           worksheet_s.write(line_left, column, final_list_of_rows[cur][column], style0)
        line_left += 1
    elif final_list_of_rows[cur][3] == 'NT':
        for column in range(0, len(final_list_of_rows[cur])-1):
           worksheet_s.write(line_right, column+5, final_list_of_rows[cur][column], style0)
           # print(final_list_of_rows[cur][column], '\n')

        line_right += 1


workbook.save('waybill_parsed.xls')

# workbook.close()


