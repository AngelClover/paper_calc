#!/usr/bin/env python
# -*- coding: utf-8 -*-
########################################################################
# 
# 
########################################################################
 
"""
File: A.py
Author: AngelClover(AngelClover@aliyun.com)
Date: 2016/03/14 21:48:14
"""
#!/usr/bin/python
#coding=utf-8
import xlrd
from pyExcelerator import *
FILE_IN_NAME = "in.xlsx"
START_FLAG = u'编号'
ID_FLAG = u'淘宝ID'
SUM_FLAG = u'总数'

FILE_OUT_NAME = "out.xls"
OUT_THING_FLAG = u'物品'
OUT_SUM_FLAG = u'SUM'
OUT_NAME_FLAG = U'ID'

def read():
    bk = xlrd.open_workbook(FILE_IN_NAME)
    try:
        sh = bk.sheet_by_name(bk._sheet_names[0])
    except:
        print "no sheet in %s named Sheet1" % FILE_IN_NAME
    n = sh.nrows
    m = sh.ncols
    #print "n:%d m:%d" % (n,m)
    cols = dict()
    for i in range(0, m):
        cols[sh.cell_value(0, i)] = i
        if sh.cell_value(0, i) == START_FLAG:
            start_pos = i
            break
    #print "start_pos:%d" % start_pos
    #print cols

    for i in range(1, n):
        start = cols[START_FLAG]
        name = sh.cell_value(i, cols[ID_FLAG])
        tot = 0
        for j in range(start, m):
            t = sh.cell_value(i, j)
            if t:
                li = t.split('*')
                num = 1
                if len(li) == 2:
                    num = int(li[1])
                tot += num
                add(name, li[0], num)
        if tot != sh.cell_value(i, cols[SUM_FLAG]):
            print "line:%d counts:%d not equal to %d" % (i, tot, sh.cell_value(i, cols[SUM_FLAG]))

papers = dict()
papers_num = dict()
def add(name, thing, num):
    #print "add (%s, %s, %d)" % (name, thing, num)
    if papers.has_key(thing) == False:
        papers[thing] = dict()
        papers_num[thing] = 0
    addpaper(papers[thing], name, num)
    papers_num[thing] += num

def addpaper(dic, name, num):
    if dic.has_key(name) == False:
        dic[name] = 0
    dic[name] += num

def write():
    w = Workbook()
    ws = w.add_sheet('Sheet1')
    ws.write(0, 0, OUT_THING_FLAG)
    ws.write(0, 1, OUT_SUM_FLAG)
    ws.write(0, 2, OUT_NAME_FLAG)

    i = 1
    tot = 0
    for k,v in papers.items():
        #print k,v
        ws.write(i, 0, k)
        ws.write(i, 1, papers_num[k])
        tot += papers_num[k]
        j = 2
        for kk,vv in v.items():
            #print kk,vv
            st = kk
            if vv != 1:
                st += ('*%d' % vv)
            ws.write(i, j, st)
            j += 1
        i += 1
    ws.write(i, 1, tot)
    w.save(FILE_OUT_NAME)

import os
os.chdir(os.path.split(os.path.realpath(__file__))[0])
read()
#print papers
#print papers_num
write()
print '%s -> %s' % (FILE_IN_NAME, FILE_OUT_NAME)
    
