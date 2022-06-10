# Importing
import numpy as np
import pandas as pd

# Entering aircraft type
type = "800"

# Loading data from excel
SnP = pd.read_excel('data\mpdsup.xls', "SYSTEMS AND POWERPLANT MAINTENA", skiprows=6, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "CAT", "TASK", "INTERVAL THRES.", "INTERVAL REPEAT", "ZONE", "ACCESS", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Str = pd.read_excel('data\mpdsup.xls', "STRUCTURAL MAINTENANCE PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "PGM", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Zon = pd.read_excel('data\mpdsup.xls', "ZONAL INSPECTION PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])

def Appliability(df, path):
    df.insert(int(len(list(df))), "APPLICABILITY", np.NaN)
    ListApp = df["APPLICABILITY APL"]
    ListAppEng = df["APPLICABILITY ENG"]
    checker = df["APPLICABILITY APL"].isnull()
    App = df["APPLICABILITY"].copy()

    LenApp = int(len(df["APPLICABILITY"]))
    RangeApp = list(range(0 , LenApp))

    for i in RangeApp:
        ListAppCheck = str(ListApp[i]).split("\n")
        ListAppEngCheck = str(ListAppEng[i]).split("\n")

        if checker[i] == True:
            App[i] = np.NaN
        else:
            if ("ALL" in ListAppCheck) and ("NOTE" not in ListAppCheck) and ("ALL" in ListAppEngCheck) and ("NOTE" not in ListAppEngCheck):
                App[i] = "Applicable"
            elif (type in ListAppCheck) and ("NOTE" not in ListAppCheck)and ("ALL" in ListAppEngCheck) and ("NOTE" not in ListAppEngCheck):
                App[i] = "Applicable"
            elif ("ALL" in ListAppCheck) and ("NOTE" in ListAppCheck) and ("ALL" in ListAppEngCheck) and ("NOTE" not in ListAppEngCheck):
                App[i] = "Note"
            elif (type in ListAppCheck) and ("NOTE" in ListAppCheck) and ("ALL" in ListAppEngCheck) and ("NOTE" not in ListAppEngCheck):
                App[i] = "Note"
            elif ("ALL" in ListAppCheck) and ("NOTE" not in ListAppCheck) and ("ALL" in ListAppEngCheck) and ("NOTE" in ListAppEngCheck):
                App[i] = "Note"
            elif (type in ListAppCheck) and ("NOTE" not in ListAppCheck) and ("ALL" in ListAppEngCheck) and ("NOTE" in ListAppEngCheck):
                App[i] = "Note"
            elif (type in ListAppCheck) and ("NOTE" not in ListAppCheck) and ("ALL" in ListAppEngCheck) and ("NOTE" in ListAppEngCheck):
                App[i] = "Note"
            elif ("ALL" not in ListAppCheck) and (type not in ListAppCheck):
                App[i] = "Not Applicable"

    df["APPLICABILITY"] = App       
    df.to_excel(path)

Appliability(SnP, "SnP.xlsx")
Appliability(Str, "Str.xlsx")
Appliability(Zon, "Zon.xlsx")

NoteList = []
def Notegetter(df):
    Description = df['TASK DESCRIPTION'].copy()
    LenDes = int(len(Description))
    RangeDes = list(range(0 , LenDes))

    for x in RangeDes:
        ListDes = str(Description[x]).split("\n")
        RangeListDes = list(range(0 , int(len(ListDes))))

        for i in RangeListDes:
            if "AIRPLANE NOTE:" in ListDes[i]:
                NoteList.append(ListDes[i])
            elif "ENGINE NOTE:" in ListDes[i]:
                NoteList.append(ListDes[i])

Notegetter(SnP)
Notegetter(Str)
Notegetter(Zon)

NoteShort = []

for x in NoteList:
    # check if exists in unique_list or not
    if x not in NoteShort:
        NoteShort.append(x)
        print(x)

print(len(NoteList))
print(len(NoteShort))