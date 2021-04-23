import numpy
import pandas
import xlrd
import xlwt

mir_book = xlrd.open_workbook("Origin\\Independant.xlsx")
miR_name = numpy.unique(mir_book.sheet_by_index(0).col_values(0) 
                        + mir_book.sheet_by_index(1).col_values(0))
print(miR_name)

mda_book = xlrd.open_workbook("Origin\\alldata.xlsx").sheet_by_index(0)



mir_disease = [];j=0
for i in range(0,len(miR_name)):
    while miR_name[i] != str(mda_book.cell_value(j,1)):
        j += 1
    mir_disease.append([str(mda_book.cell_value(j,1)),str(mda_book.cell_value(j,2))])
print(mir_disease)

pandas.DataFrame(mir_disease,index=numpy.arange(1,len(miR_name)+1,1)).to_excel("mir_disease.xls")
    
