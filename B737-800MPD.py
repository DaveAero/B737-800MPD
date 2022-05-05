# Importing 
import numpy as np
import pandas as pd

# Entering aircraft type
type = "800"

# Printer function
def printer(df, path):
    df.to_excel(path)

# Add the Applicability column to database
def addApp(df):
        count = int(len(list(df)))
        df.insert(count, "APPLICABILITY", np.nan)

# Loading data from excel
SnP = pd.read_excel('data\mpdsup.xls', "SYSTEMS AND POWERPLANT MAINTENA", skiprows=6, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "CAT", "TASK", "INTERVAL THRES.", "INTERVAL REPEAT", "ZONE", "ACCESS", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Str = pd.read_excel('data\mpdsup.xls', "STRUCTURAL MAINTENANCE PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "PGM", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Zon = pd.read_excel('data\mpdsup.xls', "ZONAL INSPECTION PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])

# Adding columns to the 3 tabs
MPD = [SnP, Str, Zon]
for col in MPD:
    addApp(col)


Applicabilities = Zon["APPLICABILITY APL"].unique()
LenApp = int(len(Applicabilities))
RangeApp = list(range(0 , LenApp))

for i in RangeApp:
    ListApp = str(Applicabilities[i]).split("\n")
    if ("ALL" in ListApp) and ("NOTE" not in ListApp):
        Zon.loc[Zon['APPLICABILITY APL'] == Applicabilities[i],  'APPLICABILITY'] = "Applicable"
        print("{} \n Applicable ALL and not NOTE \n".format(ListApp))

    elif (type in ListApp) and ("NOTE" not in ListApp):
        Zon.loc[Zon['APPLICABILITY APL'] == Applicabilities[i],  'APPLICABILITY'] = "Applicable"
        print("{} \n Applicable type and not NOTE \n".format(ListApp))

    elif ("NOTE" in ListApp) and (type in ListApp):
        Zon.loc[Zon['APPLICABILITY APL'] == Applicabilities[i],  'APPLICABILITY'] = "Note"
        print("{} \n NOTE and type \n".format(ListApp))

    elif ("NOTE" in ListApp) and ("ALL" in ListApp):
        Zon.loc[Zon['APPLICABILITY APL'] == Applicabilities[i],  'APPLICABILITY'] = "Note"
        print("{} \n NOTE and ALL \n".format(ListApp))

    else:
        Zon.loc[Zon['APPLICABILITY APL'] == Applicabilities[i],  'APPLICABILITY'] = "Not Applicable"
        print("{} \n Not Applicable \n".format(ListApp))
    Zon

Description = Zon.loc[Zon['APPLICABILITY'] == "Note", 'TASK DESCRIPTION'].unique()
LenDes = int(len(Description))
RangeDes = list(range(0 , LenDes))

#for x in RangeDes:
#    ListDes = str(Description[x]).split("\n")
#    LenListDes = int(len(ListDes))
#    RangeListDes = list(range(0 , LenListDes))
#    
#    for i in RangeListDes:
#        if "AIRPLANE NOTE:" in ListDes[i]:
#            print(ListDes[i])
#            check = input("Applicable Y/N?")
#            if check == "Y":
#                Zon.loc[Zon['TASK DESCRIPTION'] == Description[x],  'APPLICABILITY'] = "Applicable"
#            elif check == "N":
#                Zon.loc[Zon['TASK DESCRIPTION'] == Description[x],  'APPLICABILITY'] = "Not Applicable"
#            Zon

printer(SnP, "SnP.xlsx")
printer(Str, "Str.xlsx")
printer(Zon, "Zon.xlsx")