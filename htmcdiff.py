# Copyright (c) 2019 Kirill 'Kolyat' Kiselnikov
# This file is the part of medutils, released under modified MIT license
# See the file LICENSE included in this distribution

import openpyxl


def vmp_oms_difference():
    # Prepare data
    _mis = openpyxl.load_workbook('МИС.xlsx')
    _vmp_oms = openpyxl.load_workbook('ВМП-ОМС.xlsx')
    vmp_oms = _vmp_oms['Лист2']
    mis_vmp_oms = _mis['ВМП-ОМС']
    # Find difference
    vmp_oms_set = set([r[1] for r in vmp_oms.iter_rows(values_only=True)])
    mis_vmp_oms_set = set([r[4] for r in mis_vmp_oms.iter_rows(values_only=True)])
    vmp_oms_diff = sorted(vmp_oms_set - mis_vmp_oms_set)
    print('\nВМП ОМС (всего: {})'.format(len(vmp_oms_diff)))
    print('===================')
    for e in vmp_oms_diff:
        print(e)


def vmp_fed_difference():
    # Prepare data
    _mis = openpyxl.load_workbook('МИС.xlsx')
    _vmp_fed = openpyxl.load_workbook('ВМП-ФЕД.xlsx')
    vmp_fed = _vmp_fed['Лист2']
    mis_vmp_fed = _mis['ВМП-ФЕД']
    # Find difference
    vmp_fed_set = set([r[1] for r in vmp_fed.iter_rows(values_only=True)])
    mis_vmp_fed_set = set([r[4] for r in mis_vmp_fed.iter_rows(values_only=True)])
    vmp_fed_diff = sorted(vmp_fed_set - mis_vmp_fed_set)
    print('\nВМП ФЕД (всего: {})'.format(len(vmp_fed_diff)))
    print('===================')
    for e in vmp_fed_diff:
        print(e)


if __name__ == '__main__':
    vmp_oms_difference()
    vmp_fed_difference()
