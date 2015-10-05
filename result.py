#coding=utf-8
import os
import xlrd
import json
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
from xlutils.copy import copy
import xlrd
import re

def _getOutCell(outSheet, colIndex, rowIndex): 
    """ HACK: Extract the internal xlwt cell representation. """ 
    row = outSheet._Worksheet__rows.get(rowIndex) 
    if not row: return None 
    cell = row._Row__cells.get(colIndex) 
    return cell 

def setOutCell(outSheet, row, col, value): 
    """ Change cell value without changing formatting. """ 
    # HACK to retain cell style. 
    previousCell = _getOutCell(outSheet, col, row) 
    # END HACK, PART I 
    outSheet.write(row, col, value) 
    # HACK, PART II 
    newCell = None
    if previousCell:
        newCell = _getOutCell(outSheet, col, row) 
    if newCell: 
        newCell.xf_idx = previousCell.xf_idx 

if __name__ == '__main__':
    print 'start'
#    print os.getcwd()
    prefix=os.getcwd()+os.path.sep+'images'
    jsonFile = open('./vona.json')
    src = jsonFile.read()
    jsonData=json.loads(src)
    jsonData.sort(key=lambda obj:obj.get('index'))
    jsonFile.close()
    result = xlrd.open_workbook(u'./file/28400-28599.xls',formatting_info=True)
    wb=copy(result)
    s = wb.get_sheet(0)

    for each in jsonData:
        row=each['index']+2
        setOutCell(s,row,5,each['name'].decode('utf-8'))
        if each['catalog'] !=u'Catalog':
            setOutCell(s,row,9,u'无电子书'.decode('utf-8'))
        else:
            setOutCell(s,row,9,u'有且正常显示'.decode('utf-8'))
        if each['pack']:
            setOutCell(s,row,10,u'セット商品-按套出售'.decode('utf-8'))
        else:
            setOutCell(s,row,10,u'セット商品外-单个出售'.decode('utf-8'))
        if each['images'] and each['images'][0].get('path'):
            pass
        else:
            setOutCell(s,row,8,'×-无图片'.decode('utf-8'))

    wb.save('./file/result.xls')

    data = xlrd.open_workbook(os.getcwd()+os.path.sep+'file'+os.path.sep+u'生产商名录.xls')
    table = data.sheets()[0]
    brand_site={}
    brand = table.col_values(2)[2:]
    site = table.col_values(4)[2:]
    for index,each in enumerate(brand):
        if site[index] and site[index]!='':
            s = site[index].split('/')
            try:
                brand_site[each]=s[2]
            except Exception,e:
                print e
                print site[index]
#    print brand_site


    htmlFile = open('./file/items.html','w')
    htmlFile.write('<!Document><html><head><meta charset="utf-8"><title>vona web items</title></head></body>\r\n<ul>')
    for each in jsonData:
        htmlFile.write(str(each['index']+1)+'<li><div style="border:1px dashed #000"><ul>\r\n')
        htmlFile.write('<li>'+each['name']+'</li>\r\n')
        if each['images']:
#            path=os.getcwd()+os.path.sep+'images'+os.path.sep+each['images'][0]['path']
             path = '../images'+os.path.sep+os.path.sep+each['images'][0]['path']
        htmlFile.write('<li><img src="'+path+'"/></li>\r\n')
        pattern = re.compile('.*'+each['brand']+'.*',re.I)     
        s = None
        for e in brand:
            match = pattern.search(e)
            if match:
                s = brand_site.get(e)
                break
        if s:
            #            print each['brand'],e
            patten = re.compile('^www\..*',re.I)
            match = patten.match(s)
            if match:
                s=s[4:]
            htmlFile.write(u'<li><a href="https://search.yahoo.com/search;_ylt=Ak_vFvmZFT62LZ2EMk24mhabvZx4?p='+each['name']+'+site%3A'+s+'&toggle=1&cop=mss&ei=UTF-8&fr=yfp-t-901&fp=1" target="_blan">官方</a></li>\r\n'.decode('utf-8'))

        htmlFile.write(u'<li><a href="https://search.yahoo.com/search;_ylt=AwrBT8OB9gxWQ3IAWMyl87UF;_ylc=X1MDOTU4MTA0NjkEX3IDMgRmcgMEZ3ByaWQDbFBTeGh6WkpUQmVZcm9WdDJ3MHZ3QQRuX3JzbHQDMARuX3N1Z2cDMTAEb3JpZ2luA3NlYXJjaC55YWhvby5jb20EcG9zAzAEcHFzdHIDBHBxc3RybAMEcXN0cmwDNARxdWVyeQNmdWNrBHRfc3RtcAMxNDQzNjkwMTUy?p='+each['name']+'&fr=sfp&fr2=sb-top-search&iscqry=" target="_black" >非官方</a></li>\r\n'.decode('utf-8'))
        htmlFile.write('</ul></div></li>\r\n')
    htmlFile.write('</ul>\r\n</body></html>')
    htmlFile.close()

    print 'finished'
