# -*- coding: utf-8 -*-
__author__ = 'chenyang'

import xlrd
def readFile(path):
	exitFile=open(path,'r')
	return exitFile

def writeFile(path):
	newFile=open(path,'w')
	return newFile

def readExcelFile(path):
	excel=xlrd.open_workbook(path)
	return excel
