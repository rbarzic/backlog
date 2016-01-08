#/usr/bin/env python

from __future__ import division  # force division to return float
import csv
import pprint

import locale

FirstLine=1
MonthlyDepreciationColumn = 10
AmountPostedColumn = 4
ResultColumn = 11
DepreciableValueColumn=8
DepreciationPeriodColumn=5

def string2i(my_str):
    my_str2 =my_str.replace(",","")
    "Convert a string to an number"
    if (my_str2 != '') and  (my_str2 is not None):
        locale.setlocale(locale.LC_ALL, '')
        positive = my_str2.translate(None, '()')
        result = locale.atof(positive) if positive == my_str2 else -locale.atof(positive)
        return result
    else:
        return 0


result = []
with open('Depreciations_from_oct15.csv', 'rb') as csvfile:
    Excelreader = csv.reader(csvfile, delimiter=';', quotechar='|')
    line=0
    
    for row in Excelreader:
        print "-I- Reading line ",line,'...'
        row.extend(['']*80)
        if line >= FirstLine:
            
            # MonthlyDepreciation  = string2i(row[MonthlyDepreciationColumn].strip())
            depreciationValue = string2i(row[DepreciableValueColumn].strip())
            depreciationPeriod = string2i(row[DepreciationPeriodColumn].strip())
            monthlyDepreciation = depreciationValue/depreciationPeriod
            rest_prev = string2i(row[AmountPostedColumn].strip())
            col = ResultColumn
            row[MonthlyDepreciationColumn] = monthlyDepreciation
            while True:
                
                rest = rest_prev - monthlyDepreciation
                if rest>0:
                    print "col = %d" % (col)
                    row[col] = monthlyDepreciation
                    col = col + 1
                    rest_prev = rest
                    pass
                else:
                    row[col] = rest_prev
                    break
        result.append(row)
        line += 1
        

pprint.pprint(result)

with open('result_depreciation_oct2015.csv', 'wb') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerows(result)

