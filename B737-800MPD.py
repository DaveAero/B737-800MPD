# Importing
import re
import pandas as pd
import numpy as np

# Loading data from excel
SnP = pd.read_excel('data\mpdsup.xls', "SYSTEMS AND POWERPLANT MAINTENA", skiprows=6, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "CAT", "TASK", "INTERVAL THRES.", "INTERVAL REPEAT", "ZONE", "ACCESS", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Str = pd.read_excel('data\mpdsup.xls', "STRUCTURAL MAINTENANCE PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "PGM", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])
Zon = pd.read_excel('data\mpdsup.xls', "ZONAL INSPECTION PROGRAM", skiprows=5, names=["Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "ZONE", "ACCESS", "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"])

# Setting up Applicability option
Applicabilitytable = ["ALL", "NOTE", "600", "700", "700C", "700IGW", "800", "800BCF", "900", "900ER"]

############################ Functions ############################
# function for searching the applicability function
# Checker will match only if it is a full match
Checker = lambda y, value: [bool(re.fullmatch(y, x)) for x in value]

# Searcher will match if the value is in the provided list
Searcher = lambda y, value: [bool(re.search(y, x, re.IGNORECASE)) for x in value]

# Defining fuction to be used on all 3 MPD tabs
def ApplicabilityChecker(table, path):
    ############################ Aircraft Type Applicability ############################
    # Setting up the truth Table
    # Blank Database
    truthTable = pd.DataFrame()
    # Inserting all the unique values from the Applicability APL column
    truthTable.insert(column = 'APPLICABILITY APL', loc = int(len(truthTable.columns)), value = table['APPLICABILITY APL'].apply(str).unique())

    # Setting up dictionary with applicabilities split
    # Creating key pairs with the unique values of Applicability as the Key and a blank list as the pair
    Applic = dict.fromkeys(table["APPLICABILITY APL"].apply(str).unique(), list([]))
    # For each unique valuse of Applicability / each Key
    for value in table["APPLICABILITY APL"].apply(str).unique():
        # Split the value into each a list sperated by new lines
        split = str(value).split("\n")
        # Attach this list as the pair to its corrisponding key
        Applic["{}".format(value)] = list(split)
    
    # Creating the applicability table
    # Cycling through Each Applicability aircraft type
    for value in Applicabilitytable:
        # Create a blank list
        truthColumn = []
        # Using the applicability dictionary for the line split applicabilities
        for x in Applic:
            #  using checker to find full matches between the Applicability aircraft type And the split list from the key pairs
            theCheck = Checker(value, Applic["{}".format(x)])
            # if statement to return ture / false to any matches
            if True in theCheck:
                truthColumn.append(True)
            else:
                truthColumn.append(False)
        # Appending the resulting list to the Truth Table with the header being the Applicability aircraft type 
        truthTable.insert(column = '{}'.format(value), loc = int(len(truthTable.columns)), value = truthColumn)

    ############################ APPLICABILITY ############################

    # Insert blank Row for Applicability
    truthTable.insert(column = 'Applicability', loc = int(len(truthTable.columns)), value = np.nan)

    # Applying Applicability Rules
    truthTable.loc[truthTable["ALL"] == True, 'Applicability'] = True
    truthTable.loc[truthTable[type] == True, 'Applicability'] = True
    truthTable.loc[truthTable["NOTE"] == True, 'Applicability'] = "NOTE"
    truthTable.loc[(truthTable["ALL"] == False) & (truthTable[type] == False), 'Applicability'] = False

    # Merge Truthtable
    # Merginging the resulting table of all the tests to the full MPD sheet
    # The tables are matched as both contain the APPLICABILITY APL Column
    MPDTruthTable = pd.merge(table, truthTable, on = "APPLICABILITY APL", how = 'inner')

    ############################ Engine Type Applicability ############################
    # Search to check if Note is in the Engine Applicability column
    ENGtest = Searcher('NOTE', MPDTruthTable['APPLICABILITY ENG'].apply(str))
    # Adding Engine Applicability column to the Truth Table
    MPDTruthTable.insert(column = 'ENGINE NOTE', loc = int(len(MPDTruthTable.columns)), value = ENGtest)

    # Runing note checker on the new column
    MPDTruthTable.loc[MPDTruthTable['ENGINE NOTE'] == True, 'Applicability'] = "NOTE"

    ############################ Note Extraction ############################
    # Inserting a blank row for the extracted notes
    MPDTruthTable.insert(column = 'NOTE TEXT', loc = int(len(MPDTruthTable.columns)), value = np.nan)

    # Searching through the table row by row for matches with Airplane note and Engine Note
    for index, row in MPDTruthTable.iterrows():
        # On each row check the Taks Description for a match with Airplane Note
        pattern = re.compile(r'(?:AIRPLANE NOTE:)(?:(?!(INTERVAL NOTE:|SPECIAL NOTE:|SPECIAL NOTE:|ACCESS NOTE:))[A-Za-z0-9\.\s\:\'\;\(\)\-\/\, ])+')
        result = re.search(pattern, str(row['TASK DESCRIPTION']))
        # if the result is not None
        if result != None:
            # Print the extracted text to the new column
            MPDTruthTable.loc[index, 'NOTE TEXT'] = result.group(0)


    # Print to Excel at the end
    MPDTruthTable.to_excel(path)





############################ Main Code ############################
ApplicabilityChecker(SnP, "SnPTruthTable.xlsx")
ApplicabilityChecker(Str, "StrTruthTable.xlsx")
ApplicabilityChecker(Zon, "ZonTruthTable.xlsx")