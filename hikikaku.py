#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import os.path
import xlwt
import codecs
import re

def retrieveFiles(directly):
    newFiles = []
    csaFiles = []

    for root, dirs, files in os.walk(directly):
        for f in files:
            path = os.path.join(root,f)
            if f.endswith('.kif'):
                newFiles += [path]
                # print "kif:%s" % path
            elif f.endswith('.csa'):
                csaFiles += [path]
                # print "csa:%s" % path

    removeList = []
    csaNewFiles = []
    for x in csaFiles:
        root, ext = os.path.splitext(x)
        kiffile = root + '.kif'

        if not os.path.isfile(kiffile):
            print 'remove: %s' % kiffile 
            removeList += [x]
        else:
            csaNewFiles += [x]

    # print "newFiles = %d" % len(newFiles)
    # print "csaFiles = %d" % len(csaFiles)
    # print "csaNewFiles = %d" % len(csaNewFiles)
    # print "removeList = %d" % len(removeList)

    return csaNewFiles, newFiles

def calcSenkei(csa,kif):
    date = re.compile(u'開始日時：([0-9]+)/([0-9]+)/([0-9]+) ([0-9]+):([0-9]+):([0-9]+)\r\n')
    sente = re.compile(u'先手：(.+)\r\n')
    gote = re.compile(u'後手：(.+)\r\n')
    senkei = re.compile(u'戦型：(.+)\r\n')
    made   = re.compile(u'まで([0-9]+)手で([先後]手)の(入玉勝ち|勝ち)\r\n')
    summary = re.compile(u"'summary:([a-z_ ]+):([^:]+)\ (lose|win|draw):([^:]+)\ (win|lose|draw)")

    reason_map = {u'illegal move':u'反則負け',u'max_moves':u'最大手数',u'oute_sennichite':u'王手千日手',u'sennichite':u'千日手',u'time up':u'時間切れ負け',u'oute_kaihimore':u'王手回避もれ',u'uchifuzume':u'打ち歩詰め'}
    num = 1
    senkei_table = {u'合計':0}

    print (u','.join([u'年',u'月',u'日',u'時',u'分',u'秒',u'先手',u'後手',u'戦型',u'手数',u'勝者(先後)',u'勝者',u'敗者',u'結果',u'棋譜ファイル',u'リンク'])+'\r').encode('cp932')
    
    record = []

    for k in kif:
        num += 1
        senkei_str = u'戦型データなし'
        root, ext = os.path.splitext(k)
        csafile = root + '.csa'
        print 'analyze:' + str(num) + ':' + csafile
        kf = codecs.open(k, 'rU', 'cp932')
        try:
            li = kf.readlines()
        except:
            print csafile + ' is ignored!!!!!!!!!!!!!!!!!!!!!!!!!!!'
            num -= 1
            kf.close()
            continue

        kf.close()
        #print li
        kifu = []

        for el in li:
            m = date.match(el)
            if m:
                #print u'日時='.encode('utf-8') + repr(m.group(1,2,3,4,5,6))
                kifu = list(m.group(1,2,3,4,5,6))

            m = senkei.match(el)
            if m:
                #print u'戦型='.encode('utf-8') + m.group(1).encode('utf-8')
                senkei_str = m.group(1)[:]
                senkei_table[u'合計'] += 1
                ge = senkei_table.get(senkei_str,u'えらー')
                if ge == u'えらー':
                    senkei_table[senkei_str] = 1
                else:
                    senkei_table[senkei_str] += 1

            m = sente.match(el)
            if m:
                #print u'先手='.encode('utf-8') + m.group(1).encode('utf-8')
                sente_str = m.group(1)[:]
                
            m = gote.match(el)
            if m:
                #print u'後手='.encode('utf-8') + m.group(1).encode('utf-8')
                gote_str = m.group(1)[:]
        
        if senkei_str == u'戦型データなし':
            senkei_table[u'合計'] += 1
            ge = senkei_table.get(senkei_str,u'えらー')
            if ge == u'えらー':
                senkei_table[senkei_str] = 1
            else:
                senkei_table[senkei_str] += 1

        shouhai = False
        kifu += [sente_str]
        kifu += [gote_str]
        kifu += [senkei_str]

        if li[len(li)-1].startswith(u'まで'):
            m = made.match(li[len(li)-1])
            if m:
                #print u'まで'.encode('utf-8') + m.group(1).encode('utf-8') + u'手,'.encode('utf-8')+m.group(2).encode('utf-8')+u'：投了'.encode('utf-8')
                kifu += [m.group(1)+u'手']
                if m.group(3) == u'入玉勝ち':
                    kachi = u'入玉勝ち'
                else:
                    kachi = u'投了'

                if m.group(2) == u'先手':
                    kifu += [u'先手']
                    kifu += [sente_str]
                    kifu += [gote_str]
                    kifu += [kachi]
                else:
                    kifu += [u'後手']
                    kifu += [gote_str]
                    kifu += [sente_str]
                    kifu += [kachi]
                shouhai = True
            else:
                print u'xxxxx'.encode('cp932')
                print li[len(li)-1].encode('cp932')+'\r'

        else:
            cf = codecs.open(csafile, 'rU', 'cp932')
            li2 = cf.readlines()
            cf.close()

            kifu += [u'？手']

            for em in li2:
                if em.startswith("'summary:"):
                    m = summary.match(em)
                    if m:
                        # print em
                        # print 'CSA '+m.group(1).encode('utf-8')
                        if m.group(3) == u'draw':
                            #print u'まで'.encode('utf-8') + u'?手'.encode('utf-8') + u'引き分け'.encode('utf-8') + u'：'.encode('utf-8') + (reason_map[m.group(1)]).encode('utf-8')
                            kifu += [u'引き分け']
                            kifu += [u'']
                            kifu += [u'']
                            kifu += [reason_map[m.group(1)]]
                            shouhai = True
                        elif m.group(3) == u'win':
                            #print u'まで'.encode('utf-8') + u'?手'.encode('utf-8') + u'先手'.encode('utf-8') + u'：'.encode('utf-8') + (reason_map[m.group(1)]).encode('utf-8')
                            kifu += [u'先手']
                            kifu += [sente_str]
                            kifu += [gote_str]
                            kifu += [reason_map[m.group(1)]]
                            shouhai = True
                        elif m.group(5) == u'win':
                            #print u'まで'.encode('utf-8') + u'?手'.encode('utf-8') + u'後手'.encode('utf-8') + u'：'.encode('utf-8') + (reason_map[m.group(1)]).encode('utf-8')
                            kifu += [u'後手']
                            kifu += [gote_str]
                            kifu += [sente_str]
                            kifu += [reason_map[m.group(1)]]
                            shouhai = True
                    else:
                        pass

            if shouhai == False:
                # print u'まで?手勝敗不明：不明'.encode('utf-8')               
                kifu += [u'不明']
                kifu += [u'']
                kifu += [u'']
                kifu += [u'不明']

        kifu += [csafile[2:]]
        kifu += [u'=HYPERLINK(O'+ (u'%d' % num) + ')']


        kifu_t = u','.join(kifu)
        # print kifu_t.encode('utf-8')
        # print 'analyze:' + str(num) + ':' + csafile
        record += [kifu[:]]

    return record, senkei_table

def writeExcel(record, dest_filename):
    workbook = xlwt.Workbook() 
    sheet = workbook.add_sheet("SENKEI")

    row = 0
    col = 0
    for x in [u'年',u'月',u'日',u'時',u'分',u'秒',u'先手',u'後手',u'戦型',u'手数',u'勝者(先後)',u'勝者',u'敗者',u'結果',u'棋譜ファイル',u'リンク']:
        sheet.write(row, col, x) # A-P
        col += 1

    row = 1
    for re in record:
        for i in range(0,15):
            sheet.write(row, i, re[i]) # A-O
        sheet.write(row, 15, xlwt.Formula('HYPERLINK(O'+str(row+1)+')'))
        print row
        row += 1

    workbook.save(dest_filename)

from openpyxl import Workbook

def writeExcelXML(record, dest_filename):
    wb = Workbook()
    
    ws1 = wb.active
    ws1.title = u"戦型一覧表"

    ws1.append([u'年',u'月',u'日',u'時',u'分',u'秒',u'先手',u'後手',u'戦型',u'手数',u'勝者(先後)',u'勝者',u'敗者',u'結果',u'棋譜ファイル',u'リンク'])

    for row in range(2, len(record)+2):
        ws1.append(record[row-2])
        print row

    wb.save(filename = dest_filename)
    print "file=%s: %d records are saved(^^)." % (dest_filename,len(record))

import sys

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print 'usage: python hikikaku.py sheetname'
    else:
        print sys.argv[1]
        csa, kif = retrieveFiles(".")
        record,senkei = calcSenkei(csa,kif)
        #writeExcel(record,"foobar.xls")
        #writeExcelXML(record,"HikikakuKun2013.xlsx")
        writeExcelXML(record,sys.argv[1])
        for x in senkei.keys():
            print "%s: %d" % (x.encode('utf-8'),senkei[x])

