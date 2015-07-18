# -*- coding: utf-8 -*-
__author__ = 'chenyang91'
import getFile
import sys
#excel基本操作方法类
def getTableByName(fileName,sheetName):
	try:
		excel=getFile.readExcelFile(fileName)
		table=excel.sheet_by_name(sheetName.decode('utf-8'))#处理sheet名字为中文的情况
		return table
	except:
		info=sys.exc_info()
		print 'get excel file failed! errorInfo:'+str(info[1])

def getRowList(fileName,sheetName,rowNumber):
	table=getTableByName(fileName,sheetName)
	try:
		rowList=table.row_values(rowNumber)
		return  rowList
	except:
		info=sys.exc_info()
		print 'no this row! errorInfo:'+str(info[1])

def getColList(fileName,sheetName,colNumber):
	table=getTableByName(fileName,sheetName)
	try:
		colList=table.col_values(colNumber)
		return colList
	except:
		info=sys.exc_info()
		print 'no thi cols! errorInfo:'+str(info[1])

def getTableData(fileName,sheetName):
	table=getTableByName(fileName,sheetName)
	nrows=table.nrows
	data=[]
	for i in range(nrows):
		data.append(table.row_values(i))
	return data

def getCell(fileName,sheetName,row,col):
	table=getTableByName(fileName,sheetName)
	try:
		cell=table.cell(row,col).value
		return cell
	except:
		info=sys.exc_info()
		print 'no this cell errorInfo:'+str(info[1])
