#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import os.path
import codecs
import re
from openpyxl import Workbook
from openpyxl.cell import get_column_letter
import datetime

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
        broken = False
        for el in li:
            if el.startswith(u'開'):
                m = date.match(el)
                if m:
                    #print u'日時='.encode('utf-8') + repr(m.group(1,2,3,4,5,6))
                    kifu = list(m.group(1,2,3,4,5,6))
                else:
                    kifu = list(2014,1,1,0,0,0) # dummy data (broken file...)
                    broken = True

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

        if broken:
            print csafile + ' is ignored(broken file)!!!!!!!!!!!!!!!!!!!!!!!!!!!'
            num -= 1
            continue

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

        kifu += [csafile[2:]]                           # J
        kifu += [u'=HYPERLINK(J'+ (u'%d' % num) + ')']  # K
        kifu += [u'=CELL("filename")']                  # L
        unum = (u'%d' % num)
        kifu += [u'=CONCATENATE(LEFT(L'+unum+u',FIND("★",SUBSTITUTE(L'+unum+u',"/","★",LEN(L'+unum+u')-LEN(SUBSTITUTE(L'+unum+u',"/",""))),1)-1),"/",J'+unum+u')']                   # M
        kifu += [u'=HYPERLINK(RIGHT(M'+unum+u',LEN(M'+unum+u')-1))'] # N

        # kifu_t = u','.join(kifu)
        # print kifu_t.encode('utf-8')
        # print 'analyze:' + str(num) + ':' + csafile
        record += [kifu[:]]

    return record, senkei_table

def writeRow(ws1, rowIdx, reco):
    col = unicode(get_column_letter(1))
    row = unicode(str(rowIdx))
    cellName = u'%s%s' % (col,row)
    try:
        ws1[cellName] = datetime.datetime(int(reco[0]),int(reco[1]),int(reco[2]),int(reco[3]),int(reco[4]),int(reco[5]))
    except:
        print reco[0:5]

    for i in range(6,len(reco)):
        col2 = unicode(get_column_letter(i+1-5))
        cellName = u'%s%s' % (col2,row)
        ws1[cellName] = reco[i]

def writeExcelXML(record, dest_filename):
    wb = Workbook()
    
    ws1 = wb.active
    ws1.title = u"戦型一覧表"

    ws1.append([u'試合開始日時',u'先手',u'後手',u'戦型',u'手数',u'勝者(先後)',u'勝者',u'敗者',u'結果',u'棋譜ファイル',u'リンク(Excel)',u'ファイルパス1',u'ファイルパス2',u'リンク(LibreOffice)'])

    for row in range(2, len(record)+2):
        writeRow(ws1, row, record[row-2])
        print row-1

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
        writeExcelXML(record,sys.argv[1])
        for x in senkei.keys():
            print "%s: %d" % (x.encode('utf-8'),senkei[x])

