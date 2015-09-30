#coding=utf-8
import os
import xlrd
import json
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
    prefix='/home/chenyikui/Desktop/vonaweb/images/'
    jsonFile = open('/home/chenyikui/Desktop/vonaweb/vonaweb/vona.json')
    src = jsonFile.read()
    jsonData=json.loads(src)
    jsonData.sort(key=lambda obj:obj.get('index'))
    jsonFile.close()
    result = xlrd.open_workbook('/home/chenyikui/Desktop/work/测试.xls',formatting_info=True)
    wb=copy(result)
    s = wb.get_sheet(0)
    #wb = result.sheets()[0]

    row=2
    for each in jsonData:
        if os.path.exists(os.path.join(prefix,each['images'][0]['path'])):
            os.rename(os.path.join(prefix,each['images'][0]['path']),os.path.join(prefix,'full/'+each['name']+'.jpg'))

        #s.write(row,5,each['name'])
        setOutCell(s,row,5,each['name'])
        if each['catalog'] !=u'Catalog':
            setOutCell(s,row,9,u'无电子书')
        else:
            setOutCell(s,row,9,u'有且正常显示')
         #   s.write(row,9,u'无电子书')
        if each['pack']:
            setOutCell(s,row,10,u'セット商品-按套出售')
         #   s.write(row,10,u'セット商品-按套出售')
        else:
         setOutCell(s,row,10,u'セット商品外-单个出售')
         #   s.write(row,10,u'セット商品外-单个出售')
        row+=1
    wb.save('/home/chenyikui/Desktop/work/test.xls')
    print 'finished'
