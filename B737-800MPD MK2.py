# Importing
import re
import pandas as pd

# Loading data from excel
SnP = pd.read_excel('data\mpdsup.xls', "SYSTEMS AND POWERPLANT MAINTENA", skiprows=6, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "CAT", "TASK", "INTERVAL THRES.", "INTERVAL REPEAT", "ZONE", "ACCESS", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Str = pd.read_excel('data\mpdsup.xls', "STRUCTURAL MAINTENANCE PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "PGM", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Zon = pd.read_excel('data\mpdsup.xls', "ZONAL INSPECTION PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])

# Setting up an empty dataframe for the full table
fullTable = pd.DataFrame()

# Copying the values of the applicability column into a list for quick searching
ApplicabilityApl = SnP['APPLICABILITY APL'].copy()
Index = SnP['Index'].copy()
MPDItem = SnP['MPD ITEM NUMBER'].copy()

# function for checking if the value in this list is a string, to eliminate blank rows
stringCheck = lambda x: [bool(isinstance(value, str)) for value in x]

# Saving output of function to a new column in the the full table
fullTable.insert(column='Index', loc=0, value = Index)
fullTable.insert(column='MPDItem', loc=1, value = MPDItem)
fullTable.insert(column='APPLICABILITY APL', loc=2, value = ApplicabilityApl)
fullTable.insert(column='notBlank' , loc=3, value = stringCheck(ApplicabilityApl))

# Filtering list to only values with a string
truthTable = fullTable.loc[(fullTable["notBlank"] == True)]

# function for searching the applicability function
Checker = lambda x,y: [bool(re.fullmatch(y, value)) for value in x]

Applicabilitytable = ["ALL", "NOTE", "600", "700", "700C", "700IGW", "800", "800BCF", "900", "900ER"]

n=4
for value in Applicalist = Checker(truthTable['APPLICABILITY APL'], value)
    truthTable.insert(column='{}'.format(value), loc=n, value = list)bilitytable:
    
    n += 1

doubletake = truthTable.loc[(truthTable["ALL"] == False) & (truthTable["NOTE"] == False) & (truthTable["600"] == False) & (truthTable["700"] == False) & (truthTable["700C"] == False) & (truthTable["700IGW"] == False) & (truthTable["800"] == False) & (truthTable["800BCF"] == False) & (truthTable["900"] == False) & (truthTable["900ER"] == False)]

test = doubletake['APPLICABILITY APL'].copy()
print(test)

for x in test:
    applicabilityList = str(x).split("\n")
    print(applicabilityList)

doubletake.to_excel("doubletake.xlsx")
#truthTable.to_excel("truthtable.xlsx")