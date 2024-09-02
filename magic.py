# Importing Libraries
import re
import pandas as pd
import numpy as np

# Function to load data from Excel
def excelReader(file_path, tab, skip, custom_headers):
    excel = pd.read_excel(file_path, sheet_name=tab, skiprows=skip, names=custom_headers)
    return excel


def load_data(file_path):
    custom_headersSnP = [
        "Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "CAT", "TASK", "INTERVAL THRES.",
        "INTERVAL REPEAT", "ZONE", "ACCESS", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS",
        "TASK DESCRIPTION"
    ]
    SnP = excelReader(file_path, "SYSTEMS AND POWERPLANT MAINTENA", 6, custom_headersSnP)

    custom_headersStr = [
        "Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "PGM", "ZONE", "ACCESS",
        "INTERVAL THRES.", "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS",
        "TASK DESCRIPTION"
    ]
    Str = excelReader(file_path, "STRUCTURAL MAINTENANCE PROGRAM", 5, custom_headersStr)

    custom_headersZon = [
        "Revision", "Index", "MPD ITEM NUMBER", "AMM REFERENCE", "ZONE", "ACCESS", "INTERVAL THRES.",
        "INTERVAL REPEAT", "APPLICABILITY APL", "APPLICABILITY ENG", "MAN-HOURS", "TASK DESCRIPTION"
    ]
    Zon = excelReader(file_path, "ZONAL INSPECTION PROGRAM", 5, custom_headersZon)

    return SnP, Str, Zon

# Function for full match checker using regex
def full_match_checker(pattern, values):
    return [bool(re.fullmatch(pattern, value)) for value in values]

# Function for searcher using regex
def searcher(pattern, values):
    return [bool(re.search(pattern, value, re.IGNORECASE)) for value in values]

# Function to create applicability truth table
def create_applicability_truth_table(table, applicability_options):
    truth_table = pd.DataFrame({'APPLICABILITY APL': table['APPLICABILITY APL'].apply(str).unique()})
    applicability_dict = {value: value.split('\n') for value in truth_table['APPLICABILITY APL']}
    
    for aircraft_type in applicability_options:
        truth_column = []
        for value in applicability_dict.values():
            matches = full_match_checker(aircraft_type, value)
            truth_column.append(True if True in matches else False)
        truth_table[aircraft_type] = truth_column
    
    #truth_table['Applicability'] = np.nan
    return truth_table

# Function to apply applicability rules
#def apply_applicability_rules(truth_table, aircraft_type):
#    truth_table.loc[truth_table["ALL"], 'Applicability'] = True
#    truth_table.loc[truth_table[aircraft_type], 'Applicability'] = True
#    truth_table.loc[truth_table["NOTE"], 'Applicability'] = "NOTE"
#    truth_table.loc[(truth_table["ALL"] == False) & (truth_table[aircraft_type] == False), 'Applicability'] = False
#    return truth_table

# Function to merge truth table with main MPD data
def merge_truth_tables(mpd_table, truth_table):
    try:
        merged_table = pd.merge(mpd_table, truth_table, on="APPLICABILITY APL", how='inner')
        return merged_table
    except Exception as e:
        print(f"Error merging tables: {e}")
        return None

# Function to check engine applicability and combine notes
def check_and_combine_notes(mpd_table):
    engine_notes = searcher('NOTE', mpd_table['APPLICABILITY ENG'].apply(str))
    mpd_table['ENGINE NOTE'] = engine_notes
    
    # Combine 'NOTE' and 'ENGINE NOTE' columns into 'NOTE' column using OR logic
    mpd_table['NOTE'] = mpd_table['NOTE'] | mpd_table['ENGINE NOTE']
    mpd_table.drop(columns=['ENGINE NOTE'], inplace=True)  # Drop 'ENGINE NOTE' as it's no longer needed
    return mpd_table

# Function to extract all types of notes from the task description
def extract_notes(mpd_table):
    # Updated regex pattern to include all specified note types
    note_pattern = re.compile(r'(NOTE:|AIRPLANE NOTE:|INTERVAL NOTE:|SPECIAL NOTE:|ACCESS NOTE:|ENGINE NOTE:)([A-Za-z0-9\s\.:;\'\(\)\-\/,]*)', re.IGNORECASE)
    
    def find_all_notes(text):
        matches = note_pattern.findall(str(text))
        if matches:
            return ' '.join([''.join(match) for match in matches])
        return np.nan
    
    mpd_table['NOTE TEXT'] = mpd_table['TASK DESCRIPTION'].apply(find_all_notes)
    return mpd_table

# Main function to handle processing for each MPD category
def process_mpd_category(table, aircraft_type, output_path):
    applicability_options = ["ALL", "NOTE", "600", "700", "700C", "700IGW", "800", "800BCF", "900", "900ER"]
    truth_table = create_applicability_truth_table(table, applicability_options)
    #truth_table = apply_applicability_rules(truth_table, aircraft_type)
    merged_table = merge_truth_tables(table, truth_table)
    if merged_table is not None:
        merged_table = check_and_combine_notes(merged_table)
        merged_table = extract_notes(merged_table)
        merged_table.to_excel(output_path)
        print(f"Processed {output_path}")

# Main script execution
if __name__ == "__main__":
    file_path = 'data/mpdsup.xls'
    SnP, Str, Zon = load_data(file_path)
    process_mpd_category(SnP, "800", "SnPTruthTable.xlsx")
    process_mpd_category(Str, "800", "StrTruthTable.xlsx")
    process_mpd_category(Zon, "800", "ZonTruthTable.xlsx")