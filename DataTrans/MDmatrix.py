import numpy
import pandas
import xlrd
import xlwt
import csv

##
databook = xlrd.open_workbook("Origin\\Independant.xlsx")
miR_name  = databook.sheet_by_index(0).col_values(0) + databook.sheet_by_index(1).col_values(0)
Drug_name = databook.sheet_by_index(0).col_values(1) + databook.sheet_by_index(1).col_values(1)

miR_name  = numpy.unique(miR_name);  
Drug_name = numpy.unique(Drug_name); 
print(miR_name,"\n\n",Drug_name)
col = len(Drug_name); row = len(miR_name)
print("Drug:",row)
print("miRNA:",col)

##
dic_mir={}; j = 0
for i in miR_name:
    dic_mir[i] = j
    j += 1
print(dic_mir)  

dic_drug={}; j = 0
for i in Drug_name:
    dic_drug[i] = j
    j += 1
print(dic_drug)  

pos = databook.sheet_by_index(0)
neg = databook.sheet_by_index(1)
mdm = numpy.zeros((row,col))

for i in range(0,pos.nrows):
    r = dic_mir[str(pos.cell_value(i,0))]
    d = dic_drug[str(pos.cell_value(i,1))]
    mdm[[r,],[d,]] = 1.0
print(mdm)

for i in range(0,neg.nrows):
    r = dic_mir[str(neg.cell_value(i,0))]
    d = dic_drug[str(neg.cell_value(i,1))]
    mdm[[r,],[d,]] = -1.0

print(list(mdm[...,0]))
print(pandas.DataFrame(mdm))

bdf = {}; j = 0
for i in Drug_name:
    bdf[i] = list(mdm[...,j])
    j += 1


pandas.DataFrame(bdf,index = list(miR_name)).to_excel("mirna_drug.xls")
pandas.DataFrame(miR_name,index=list(numpy.arange(1,row+1,1))).to_excel("miRNA.xls")
pandas.DataFrame(Drug_name,index=list(numpy.arange(1,col+1,1))).to_excel("Drug.xls")
