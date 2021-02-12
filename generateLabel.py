#Version 1.1 Date Feb 12, 2021

import sys
import datetime
import time
from docx import Document
from docx.shared import Length
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd

global cellPointer  #There are four cell per page
global table 
table = None
cellPointer = 0

def generateQuantityString(quantity):
    dozen = int(quantity/12)
    remaining = quantity%12
    resultString = ""
    if dozen > 0:
        resultString = str(dozen) + " โหล"
    
    if remaining > 0:
        resultString = resultString + " " + str(remaining) + " ชิ้น"
    
    return resultString

def writeProductDataToTable(dozenCount, modelName, modelColor, modelSize , modelQuantity):
    global cellPointer
    global table
    number = str(dozenCount)
    model = modelName
    color = modelColor
    size = str(modelSize)
    quantityString = generateQuantityString(modelQuantity)
    
    if(cellPointer % 4 == 0):
        table = document.add_table(rows=2, cols=2)
        cellPointer = 0
    
    settings = list()
    rowIndex = int(cellPointer/2) 
    columnIndex = cellPointer % 2
    
    cells  = table.rows[rowIndex].cells

    p = cells[columnIndex].add_paragraph('(โหลที่)')
    p.add_run(' ' + number).bold = True
    settings.append(p)

    p = cells[columnIndex].add_paragraph('(รุ่น)')
    p.add_run(' ' + model).bold = True
    settings.append(p)

    p = cells[columnIndex].add_paragraph('(สี)')
    p.add_run(' ' + color).bold = True
    settings.append(p)

    p = cells[columnIndex].add_paragraph('(ขนาด)')
    p.add_run(' ' + size).bold = True
    settings.append(p)

    p = cells[columnIndex].add_paragraph('(จำนวน)')
    p.add_run(' ' + quantityString).bold = True
    settings.append(p)       

    for text in settings:
        text.alignment  = WD_ALIGN_PARAGRAPH.CENTER
    
    cellPointer = cellPointer + 1
        
    if(cellPointer == 4):
        document.add_page_break()

try:
	product_data = pd.read_excel("product_input.xlsx")

	try:
		document = Document()
		style = document.styles['Normal']
		font = style.font
		font.name = 'TH Sarabun New'
		font.size = Pt(24)

		columnList = product_data.columns.tolist()
		startColumnSize = 2
		productRowCount = product_data.shape[0]
		productSizeCount = product_data.shape[1] - startColumnSize

		for i in range(productRowCount):
			productModelName = product_data['model'][i]
			productModelColor = product_data['color'][i]
			for j in range(productSizeCount):
				productModelSize = product_data.columns[(startColumnSize ) + j]
				currentProductSizeCount = product_data[productModelSize][i]
				dozenCount = 1
				while (currentProductSizeCount > 0):
					if(currentProductSizeCount > 12 and currentProductSizeCount < 18):
						writeProductDataToTable(dozenCount, productModelName, productModelColor, productModelSize , currentProductSizeCount)
						dozenCount = dozenCount + 1
						currentProductSizeCount = 0
					elif(currentProductSizeCount >= 12):
						writeProductDataToTable(dozenCount, productModelName, productModelColor, productModelSize , 12)
						currentProductSizeCount = currentProductSizeCount - 12
						dozenCount = dozenCount + 1
					elif(currentProductSizeCount < 12):
						writeProductDataToTable(dozenCount, productModelName, productModelColor, productModelSize , currentProductSizeCount)
						currentProductSizeCount = 0
						dozenCount = dozenCount + 1

		timeNow = datetime.datetime.now()
		document.save( str(timeNow.year)+ str(timeNow.month) + str(timeNow.day)+ '_' + str(timeNow.hour) + "-" + str(timeNow.minute)+ "-" + str(timeNow.second)+ '.docx')
	except  Exception as e:
		print("[Error] Please check following error")
		print(e)
		time.sleep(10)
except:
	print("[Error] Cannot open product_input.xlsx")
	time.sleep(10)








