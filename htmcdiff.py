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
    # Output
    print(f'\n{secondary_wb_name} (всего: {len(set_diff)})')
    print('===================')
    print(*set_diff, sep='\n')


if __name__ == '__main__':
    mis_wb = openpyxl.load_workbook('МИС.xlsx')
    find_difference(mis_wb, 'ВМП ФЕД')
    find_difference(mis_wb, 'ВМП ОМС')
