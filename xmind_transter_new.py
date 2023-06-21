import openpyxl
from openpyxl.styles import Alignment
from xmindparser import xmind_to_dict
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

"""适用于禅道的用例导入，xmind形式编写用例后，自动生成禅道导入格式的用例"""
def resolve_path(d, lists, title):
    title = title.strip()
    if len(title) == 0:
        concat_title = d['title'].strip()
    else:
        if 'makers' in d.keys():
            level = ''
            if 'priority-2' in d['makers']:
                level = 'P1'
            elif 'priority-1' in d['makers']:
                level = 'P0'
            elif 'priority-3' in d['makers']:
                level = 'P2'
            concat_title = title + '\t' + d['title'].strip() + '\t' + level
        else:
            concat_title = title + '\t' + d['title'].strip()
    if 'topics' not in d:
        lists.append(concat_title)
    else:
        for sub_d in d['topics']:
            resolve_path(sub_d, lists, concat_title)


def xmind_cat(lst):
    # wb = openpyxl.load_workbook(excelname)
    wb = openpyxl.Workbook()
    sheetname = wb.sheetnames
    sheet = wb[sheetname[0]]
    title_list = ['所属模块', '用例标题', '优先级', '前置条件', '步骤', '预期', '用例类型']
    for i in range(1, len(title_list) + 1):
        sheet.cell(row=1, column=i).value = title_list[i - 1]
    index = 1
    for h in range(len(lst)):
        lists = []
        resolve_path(lst[h], lists, '')
        for j in range(len(lists)):
            lists[j] = lists[j].split('\t')
            start_column = 1
            if len(lists[j]) >= 3:
                module_details = '-'.join(lists[j][1:4])
                sheet.cell(row=j + index + 1, column=1).value = lists[j][0]
                sheet.cell(row=j + index + 1, column=2).value = module_details
                if len(lists[j]) > 4:
                    sheet.cell(row=j + index + 1, column=3).value = lists[j][4]
                sheet.cell(row=j + index + 1, column=7).value = "功能测试"
        if j == len(lists) - 1:
            index += len(lists)
    now = datetime.now()
    date_time = now.strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{date_time}_case.xlsx"
    wb.save(filename)


def select_file():
    filepath = filedialog.askopenfilename()
    return filepath


def maintest():
    file_name = select_file()
    out = xmind_to_dict(file_name)
    xmind_cat(out[0]['topic']['topics'])


if __name__ == '__main__':
    maintest()