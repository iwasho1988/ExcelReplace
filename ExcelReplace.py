#!/usr/bin/env python3
# coding: utf-8

import openpyxl as opx
import json
import os
import sys

REPLACE_TXT = os.path.dirname(os.path.abspath(__file__)) + "/replace.txt"


def read_text(args):
    replace_list_dict = {}
    with open(REPLACE_TXT,mode="r",encoding='utf-8') as f:
        for rows in f.read().split():
            row = rows.split(":")
            replace_list_dict[row[0]] = row[1]
    return replace_list_dict

def load_data(ws,replace_list_dict):
    replace_target_dict = {}
    for row in ws:
        for cell in row:
            if cell.value in replace_list_dict.keys():
                replace_target_dict[cell.coordinate] = cell.value
    return replace_target_dict

def replace_cell(ws,replace_list_dict,replace_target_dict):
    for num,val in replace_target_dict.items():
        ws[num].value = replace_list_dict[val]
    return "OK"

def main():
    try:
        args = sys.argv
        EXCEL_FILE = args[1]
        EXCEL_SHEET = args[2]
        print(EXCEL_FILE,EXCEL_SHEET)

        replace_list_dict = read_text(args)
        wb = opx.load_workbook(EXCEL_FILE)
        ws = wb[EXCEL_SHEET]
        replace_target_dict = load_data(ws,replace_list_dict)
        if not replace_target_dict:
            return "Nothing"
        result = replace_cell(ws,replace_list_dict,replace_target_dict)
        wb.save(EXCEL_FILE)
        return "success!!"
    except Exception as e:
        return str(e)

if __name__ == '__main__':
    print(main())
