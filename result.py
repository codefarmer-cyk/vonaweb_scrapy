#coding=utf-8
import os
import xlrd
import json
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
from xlutils.copy import copy

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
    prefix='/home/chenyikui/Desktop/vonaweb/vonaweb/images'
    jsonFile = open('./vona.json')
    src = jsonFile.read()
    jsonData=json.loads(src)
    jsonData.sort(key=lambda obj:obj.get('index'))
    jsonFile.close()
    result = xlrd.open_workbook('./file/2000-2199逸逵.xls',formatting_info=True)
    wb=copy(result)
    s = wb.get_sheet(0)
    #wb = result.sheets()[0]

    row=2
    for each in jsonData:
        if each['images']:
            if os.path.exists(os.path.join(prefix,each['images'][0]['path'])):
                try:
                    os.rename(os.path.join(prefix,each['images'][0]['path']),os.path.join(prefix,'full/'+each['name']+'.jpg'))
                except OSError,e:
                    print e
                    print each['name']
                    os.rename(os.path.join(prefix,each['images'][0]['path']),os.path.join(prefix,'full/'+each['name'].split('/')[0]+'.jpg'))

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

        row+=1
    wb.save('./file/test.xls')
    print 'finished'
