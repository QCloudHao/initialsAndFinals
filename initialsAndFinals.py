#!/usr/bin/env python3
# -*- coding: utf-8 -*-
__author__ = 'qyh'
__date__ = '2019/3/15 16:09'

import xlrd
import xlwt
from xlutils.copy import copy

filename = "D:\毕设\项目文件\录入数据库文件\常用汉字信息.xls"
tone_dic = {'ā': 'a', 'á': 'a', 'ǎ': 'a', 'à': 'a',
            'ō': 'o', 'ó': 'o', 'ǒ': 'o', 'ò': 'o',
            'ē': 'e', 'é': 'e', 'ě': 'e', 'è': 'e',
            'ī': 'i', 'í': 'i', 'ǐ': 'i', 'ì': 'i',
            'ū': 'u', 'ú': 'u', 'ǔ': 'u', 'ù': 'u',
            'ǖ': 'ü', 'ǘ': 'ü', 'ǚ': 'ü', 'ǜ': 'ü'}
initials = ['b', 'p', 'm', 'f', 'd', 't', 'n', 'l', 'g', 'k', 'h', 'j', 'q', 'x',
            'zh', 'ch', 'sh', 'r', 'z', 'c', 's', 'y', 'w']
special_initials = ['j', 'q', 'x', 'y']
special_finals = ['u', 'ue', 'uan', 'un']

# 去除音调
def remove_tone(pinyin):
    for i in range(0, len(pinyin)):
        if pinyin[i] in tone_dic.keys():
            tone = pinyin[i]
            no_tone = pinyin.replace(tone, tone_dic[tone])
            return no_tone
    return pinyin

# u转化为ü
def u_to_v(initial, final):
    true_final = final
    if initial in special_initials and final in special_finals:
        true_final = final.replace('u', 'ü ')
    return true_final


def get_initials_finals(pinyin):
    length = len(pinyin)
    # 长度为1，肯定只有一个韵母
    if length == 1:
        return "", pinyin
    else:
        if pinyin[:2] in initials:
            initial = pinyin[:2]
            final = pinyin[2:]
        elif pinyin[0] in initials:
            initial = pinyin[0]
            final = pinyin[1:]
        else:
            initial = ""
            final = pinyin
        final = u_to_v(initial, final)
        return initial, final


def main():
    rb = xlrd.open_workbook(filename)
    sheet = rb.sheet_by_name("Sheet1")
    wb = copy(rb)
    write_sheet = wb.get_sheet(0)
    nrows = sheet.nrows

    for i in range(1, nrows):
        data = sheet.cell_value(i, 2)
        data_lis = data.split("&")
        length = len(data_lis)
        if length == 1:
            pinyin = remove_tone(data)
            initial, final = get_initials_finals(pinyin)
            # 写入到第四第五列
            write_sheet.write(i, 3, initial)
            write_sheet.write(i, 4, final)
        else:
            # 多音字可能有重复的声母或韵母，去重
            initials_set = set()
            finals_set = set()
            for j in range(0, length):
                pinyin = remove_tone(data_lis[j])
                initial, final = get_initials_finals(pinyin)
                if initial != "":
                    initials_set.add(initial)
                finals_set.add(final)
            initials_str = "&".join(initials_set)
            finals_str = "&".join(finals_set)
            # print(initials_str, finals_str)
            write_sheet.write(i, 3, initials_str)
            write_sheet.write(i, 4, finals_str)
    wb.save("D:\毕设\项目文件\录入数据库文件\常用汉字信息.xls")


if __name__ == '__main__':
    main()
