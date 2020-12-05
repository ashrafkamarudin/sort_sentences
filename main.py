
# Todo
# [x] Read sentences from "System Requirement MODEL.docx"
# [x] Read variable from "Variable Model.xlsx"
#    [x] get sentences in structures like 
# [x] Sort these sentences into 2 file based on variables in file "Variable Model.xlsx"

from pprint import pprint
import helper
import xlrd
from docx2python import docx2python
import docx
import config

def extractVariableFromSheet(columns, sheet):
    dict = {}

    for key in columns:
        list = helper.extractRowFromSheet(column=key, sheet=sheet)
        dict[columns[key]] = list

    return dict

def seperateSentenceBasedOnVariable(sentences, variables):
    data = { "in": [], "not": [] }

    for sentence in sentences:
        if helper.isVariableInSentence(needle=sentence, haystack=variables):
            data["in"].append(sentence)
        else:
            data["not"].append(sentence)
    return data

# Step 1: Load required files
sheet = xlrd.open_workbook(config.file["sheet"]["path"]).sheet_by_index(0) # variables sheet
document = docx2python(config.file["docx"]["path"]) # docx

# step 2: Extract data from loaded files (sheet)
title = helper.extractColumnFromSheet(row=1, sheet=sheet)
variables = extractVariableFromSheet(columns=title, sheet=sheet)

# Step 3: Reform Loaded List (docx)
newList = helper.reformatToNumberedList(list= helper.setListLowerBound(
    list= document.body[0][0][0], lowerBound = config.file["docx"]["start_from_row"])
)

# Step 4: Perfrom seperation
data = seperateSentenceBasedOnVariable(sentences=newList, variables= variables)

# Step 5: Print Output into docx
helper.writeToDocx(toWrite=data["in"], doc_path= config.output["exist"])
helper.writeToDocx(toWrite=data["not"], doc_path= config.output["not_exist"])

