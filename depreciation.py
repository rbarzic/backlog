

from __future__ import division # force division to return float
import csv
import pprint

import locale

FirstLine=6
MonthlyDepreciationColumn = 12
AmountPostedColumn = 8
ResultColumn = 16


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
with open('depreciation-oct14.csv', 'rb') as csvfile:
    Excelreader = csv.reader(csvfile, delimiter=';', quotechar='|')
    line=0
    
    for row in Excelreader:
        row.extend(['']*60)
        if line >= FirstLine:
            
            MonthlyDepreciation  = string2i(row[MonthlyDepreciationColumn].strip())
            rest_prev         = string2i(row [AmountPostedColumn]. strip ())
            col = ResultColumn
            
            while True:
                rest = rest_prev - MonthlyDepreciation
                if rest>0:
                    # print "col = %d" % (col)
                    row[col] = MonthlyDepreciation
                    col = col + 1
                    rest_prev = rest
                    pass
                else:
                    row[col] = rest_prev
                    break
        result.append(row)
        line += 1
        

pprint.pprint(result)

with open('result_depreciation.csv', 'wb') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerows(result)

