""" Parse an Excel file"""
from __future__ import division # force division to return float
import csv
import pprint

import locale

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



#http://stackoverflow.com/questions/651794/whats-the-best-way-to-initialize-a-dict-of-dicts-in-python
class AutoVivification(dict):
    """Implementation of perl's autovivification feature."""
    def __getitem__(self, item):
        try:
            return dict.__getitem__(self, item)
        except KeyError:
            value = self[item] = type(self)()
            return value




def compute_revenue_product(revenue,expense):
    if expense>0:
        return revenue
    else:
        return revenue+expense*1.1


def compute_backlog(project,project_data):
    print(">>>>>>>>>>>>>>>>>>>>>>")
    print(">>>>>> " + project)    
    print(">>>>>>>>>>>>>>>>>>>>>>")
    
    pprint.pprint(project_data ['revenue product'])
    revenue_product = project_data ['revenue product']
    remaining_hours =  project_data['Remaining hours']
    print "Remaining hours : " + str(remaining_hours)
    if(remaining_hours ==0):
        project_data['backlog  2014']        =    [0 for x in project_data['Hours 2014']]      
        project_data['backlog  2015']        =    [0 for x in project_data['Hours 2015']]      
        project_data['backlog  2016']        =    [0 for x in project_data['Hours 2016']]      
        project_data['backlog  2017']        =    [0 for x in project_data['Hours 2017']]      
        project_data['backlog  2018']        =    [0 for x in project_data['Hours 2018']]      
        project_data['backlog  2019']        =    [0 for x in project_data['Hours 2019']]      
        project_data['backlog  2020']        =    [0 for x in project_data['Hours 2020']]      
        project_data['backlog  2021']        =    [0 for x in project_data['Hours 2021']]      
        project_data['backlog  over_2021']   =    [0 for x in project_data['Hours over_2021']] 
        
    else:
        project_data['backlog  2014']        =    [((revenue_product* x)/ remaining_hours) for x in project_data['Hours 2014']]      
        project_data['backlog  2015']        =    [((revenue_product* x)/ remaining_hours) for x in project_data['Hours 2015']]      
        project_data['backlog  2016']        =    [((revenue_product* x)/ remaining_hours) for x in project_data['Hours 2016']]      
        project_data['backlog  2017']        =    [((revenue_product* x)/ remaining_hours) for x in project_data['Hours 2017']]      
        project_data['backlog  2018']        =    [((revenue_product* x)/ remaining_hours) for x in project_data['Hours 2018']]      
        project_data['backlog  2019']        =    [((revenue_product* x)/ remaining_hours) for x in project_data['Hours 2019']]      
        project_data['backlog  2020']        =    [((revenue_product* x)/ remaining_hours) for x in project_data['Hours 2020']]      
        project_data['backlog  2021']        =    [((revenue_product* x)/ remaining_hours) for x in project_data['Hours 2021']]      
        project_data['backlog  over_2021']   =    [((revenue_product* x)/ remaining_hours) for x in project_data['Hours over_2021']] 
        


def compute_sum_of_remaining_hours(project_data):
    s =0
    s += sum(project_data['Hours 2014'])
    s += sum(project_data['Hours 2015'])
    s += sum(project_data['Hours 2016'])
    s += sum(project_data['Hours 2017'])
    s += sum(project_data['Hours 2018'])
    s += sum(project_data['Hours 2019'])
    s += sum(project_data['Hours 2020'])
    s += sum(project_data['Hours 2021'])
    s += sum(project_data['Hours over_2021'])
    project_data['Remaining hours'] = s
    
def compute_all(projects):
    "Compute everything..."
    for project, data in projects.iteritems():
        pprint.pprint 
        projects[project]['revenue product'] = compute_revenue_product(string2i(data['Revenue']),string2i(data['Expense']))
        compute_sum_of_remaining_hours(data)
        compute_backlog(project,data)



def print_result(projects):
    rows =[]
    row = ['Name','Maconomy','Consolidate','Revenue','Revenue product',1,2,3,1,2,3,4,5,6,7,8,9,10,11,12,1,2,3,4,5,6,7,8,9,10,11,12,2017,2018,2019,2020,2021,'over2021']
    rows.append(row)
    for project, data in projects.iteritems():
        row =[]
        row.append(project)
        row.append(data['MaconomyID'])
        row.append(data['Consolidate'])
        row.append(data['Revenue'])
        row.append(data['revenue product'])
        
        row.extend(data['backlog  2014'])
        row.extend(data['backlog  2015'])
        row.extend(data['backlog  2016'])
        row.extend(data['backlog  2017'])
        row.extend(data['backlog  2018'])
        row.extend(data['backlog  2019'])
        row.extend(data['backlog  2020'])
        row.extend(data['backlog  2021'])
        row.extend(data['backlog  over_2021'])
        rows.append(row)

    return rows

# Invalid constant name
# pylint: disable-msg=C0103
CommentsColumn =0
StatusColumn = 2
ProductColumn = 4
ProjectNameColumn = 9
DescriptionColumn = 10
MaconomyIDColumn = 8
ConsolidateColumn = 7

RevenueColumn = 12
TotalHourColumn = 13
ExpenseColumn = 15
TotalCostColumn = 16


firstMonth = 'Jan14'
currentMonth = 'Oct14'
lastEntry = 'Over 2021'




firstMonth_idx  = 0
currentMonth_idx  = 0
lastEntry_idx = 0



budgetDescriptionField = 'BUDGET/Planlegging'
forecastDescriptionField = 'FCST/Prognose'
actualValidDescriptionField = 'ACTUAL'
restValidDescriptionField = 'Rest'



previous_project = "an unlikely name"
current_project = ''
line_in_project = 0

projects = AutoVivification()


with open('oct14.csv', 'rb') as csvfile:
# with open('Project-review-Aug14-2-entries.csv', 'rb') as csvfile:
    Excelreader = csv.reader(csvfile, delimiter=';', quotechar='|')
    for row in Excelreader:
        comments = row[CommentsColumn].strip()
        status = row[StatusColumn].strip()
        description = row[DescriptionColumn].strip()
        product = row[ProductColumn].strip()
        if comments == 'Comments':
            # find index for first month and for last year
            for i in range(0, len(row)):
                if row[i] == firstMonth:
                    firstMonth_idx = i
                if row[i] == currentMonth:
                    currentMonth_idx = i
                if row[i] == lastEntry:
                    lastEntry_idx = i
            print "firstMonth_idx : " + str(firstMonth_idx)
            print "currentMonth_idx : " + str(currentMonth_idx)
            print "lastEntry_idx : " + str(lastEntry_idx)


            y2014_idx =  firstMonth_idx
            y2015_idx =  y2014_idx +12
            y2016_idx =  y2015_idx +12
            y2017_idx =  y2016_idx +12
            y2018_idx =  y2017_idx +1
            y2019_idx =  y2018_idx +1
            y2020_idx =  y2019_idx +1
            y2021_idx =  y2020_idx +1
            over2021_idx = y2021_idx +1

            # jan = 0, dec =11
            y2014_length = 12-(currentMonth_idx-firstMonth_idx)
            y2015_length = 12
            y2016_length = 12
            y2017_length = 1
            y2018_length = 1
            y2019_length = 1
            y2020_length = 1
            y2021_length = 1
            over2021_length = 1



        if (description != '') and (product != '') and  (status == "Open"):
            print "Description : " + description
            line_in_project = line_in_project + 1
            if description == budgetDescriptionField:
                line_in_project = 0

            if description == forecastDescriptionField:
                project  = row[ProjectNameColumn].strip()
                maconomyID = row[MaconomyIDColumn].strip()
                consolidate = row[ConsolidateColumn].strip()

                hours_2014 = row[currentMonth_idx:(currentMonth_idx + y2014_length)]
                print "y2014_length" + str(y2014_length)
                hours_2015 = row[y2015_idx:(y2015_idx + y2015_length)]
                hours_2016 = row[y2016_idx:(y2016_idx + y2016_length)]
                hours_2017 = row[y2017_idx:(y2017_idx + y2017_length)]
                hours_2018 = row[y2018_idx:(y2018_idx + y2018_length)]
                hours_2019 = row[y2019_idx:(y2019_idx + y2019_length)]
                hours_2020 = row[y2020_idx:(y2020_idx + y2020_length)]
                hours_2021 = row[y2021_idx:(y2021_idx + y2021_length)]
                hours_over2021 = row[over2021_idx:(over2021_idx + over2021_length)]



            if (description == restValidDescriptionField) and (line_in_project ==3):
                revenue   = row[RevenueColumn].strip()
                totalHour = row[TotalHourColumn].strip()
                expense   = row[ExpenseColumn].strip()
                totalCost = row[TotalCostColumn].strip()

                # this is the last field
                print "============= Project %s ==================" % (project)
                projects[project]['MaconomyID'] = maconomyID
                projects[project]['Product']    = product
                projects[project]['Consolidate']    = consolidate

                projects[project]['Revenue']    = revenue
                projects[project]['TotalHour']  = totalHour
                projects[project]['Expense']    = expense
                projects[project]['TotalCost']  = totalCost
                # projects[project]['row']  = row




                projects[project]['Hours 2014'] = [string2i(x) for x in hours_2014]
                projects[project]['Hours 2015']    = [string2i(x) for x in hours_2015]
                projects[project]['Hours 2016']    = [string2i(x) for x in hours_2016]
                projects[project]['Hours 2017']    = [string2i(x) for x in hours_2017]
                projects[project]['Hours 2018']    = [string2i(x) for x in hours_2018]
                projects[project]['Hours 2019']    = [string2i(x) for x in hours_2019]
                projects[project]['Hours 2020']    = [string2i(x) for x in hours_2020]
                projects[project]['Hours 2021']    = [string2i(x) for x in hours_2021]
                projects[project]['Hours over_2021'] = [string2i(x) for x in hours_over2021]


print "#"*80
print "    Starting calculation"
print "#"*80




compute_all(projects)

pprint.pprint(projects)

rows = print_result(projects)


pprint.pprint(rows)


with open('result.csv', 'wb') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerows(rows)
