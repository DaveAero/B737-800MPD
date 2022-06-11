# Importing
import re
import pandas as pd

# Loading data from excel
SnP = pd.read_excel('data\mpdsup.xls', "SYSTEMS AND POWERPLANT MAINTENA", skiprows=6, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "CAT", "TASK", "INTERVAL THRES.", "INTERVAL REPEAT", "ZONE", "ACCESS", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Str = pd.read_excel('data\mpdsup.xls', "STRUCTURAL MAINTENANCE PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "PGM", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Zon = pd.read_excel('data\mpdsup.xls', "ZONAL INSPECTION PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])

# function for checking if the value in this list is a string, to eliminate blank rows
stringCheck = lambda x: [bool(isinstance(value, str)) for value in x]

truthTable = pd.DataFrame()
SnP.insert(column = 'BlankTest', loc = int(len(SnP.columns)), value = stringCheck(SnP['APPLICABILITY APL']))

truthTable.to_excel("truthTable.xlsx")

# function for searching the applicability function
Checker = lambda x,y: [bool(re.fullmatch(y, value)) for value in x]

Applicabilitytable = ["ALL", "NOTE", "600", "700", "700C", "700IGW", "800", "800BCF", "900", "900ER"]

# Make a truth table for the unique values and then apply this to the whole SnP table

#    list = Checker(SnP['APPLICABILITY APL'], value)
#    SnP.insert(column='{}'.format(value), loc = int(len(SnP.columns)), value = list)