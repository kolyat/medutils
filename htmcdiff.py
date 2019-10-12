# Copyright (c) 2019 Kirill 'Kolyat' Kiselnikov
# This file is the part of medutils, released under modified MIT license
# See the file LICENSE included in this distribution

import openpyxl


def find_difference(primary_wb_obj, secondary_wb_name):
    # Prepare data
    primary_data = primary_wb_obj[secondary_wb_name]
    secondary_wb = openpyxl.load_workbook(f'{secondary_wb_name}.xlsx')
    secondary_data = secondary_wb['Лист2']
    # Find difference
    primary_set = set()
    for row in primary_data.iter_rows(values_only=True):
        primary_set.add(row[4])
    secondary_set = set()
    for row in secondary_data.iter_rows(values_only=True):
        secondary_set.add(row[1])
    set_diff = sorted(secondary_set - primary_set)
    total = len(set_diff)
    # Output
    # print(f'\n{secondary_wb_name} (всего: {total})')
    # print('===================')
    # print(*set_diff, sep='\n')
    sheet = secondary_wb.create_sheet(title=secondary_wb_name)
    for i in range(total):
        sheet.cell(column=1, row=i+1, value=set_diff[i])
    secondary_wb.save(f'{secondary_wb_name}.xlsx')


if __name__ == '__main__':
    mis_wb = openpyxl.load_workbook('МИС.xlsx')
    find_difference(mis_wb, 'ВМП ФЕД')
    find_difference(mis_wb, 'ВМП ОМС')
