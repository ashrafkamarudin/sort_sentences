
# Todo
# [x] Read sentences from "System Requirement MODEL.docx"
# [x] Read variable from "Variable Model.xlsx"
#    [x] get sentences in structures like 
# [x] Sort these sentences into 2 file based on variables in file "Variable Model.xlsx"
# [x] Count Occurance of "variables"
# [ ] Count/List Single Word in "System Requirement MODEL.docx"
# [ ] Count/List Double Word (2 word) in "System Requirement MODEL.docx"
# [ ] Draw flow / algorithm

import helper
import xlrd
from docx2python import docx2python
import json
import config

def extractVariableFromSheet(columns, sheet):
    dict = {}

    for key in columns:
        list = helper.extractRowFromSheet(column=key, sheet=sheet)
        dict[columns[key]] = list

    return dict

def seperateSentenceBasedOnVariable(sentences, variables):
    data = { "in": [], "not": [] }
    count_variable = {}
    singleCount = {}
    doubleCount = {}

    for sentence in sentences:
        new_sentence = helper.removeStopWord(stopWordList=config.stopWords,sentence=sentence)

        # single count
        singleCount = helper.countSingleWord(new_sentence, singleCount)

        # doubly count
        doubleCount = helper.countDoubleWord(new_sentence, doubleCount)
        
        found, count_variable = helper.isVariableInSentence(needle=new_sentence, haystack=variables, count_var=count_variable);
        if found:
            data["in"].append(sentence)
        else:
            data["not"].append(sentence)
    
    return data, count_variable, singleCount, doubleCount

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
data, count_variable, singleCount, doubleCount = seperateSentenceBasedOnVariable(sentences=newList, variables= variables)

count_variable = dict(sorted(count_variable.items(), key=lambda item: item[1],reverse=True if config.output["variable_count"]["order"] == "DESC" else False))

variable_count = []
for key in count_variable:
    variable_count.append(f"{key} => {count_variable[key]} ")

single_count = []
for key in singleCount:
    single_count.append(f"{key} => {singleCount[key]} ")

double_count = []
for key in doubleCount:
    double_count.append(f"{key} => {doubleCount[key]} ")

# Step 5: Print Output into docx
helper.writeToDocx(toWrite=data["in"], doc_path= config.output["exist"])
helper.writeToDocx(toWrite=data["not"], doc_path= config.output["not_exist"])
helper.writeToDocx(toWrite=variable_count, doc_path= config.output["variable_count"]["path"])

# WIP
# remove special chars
# restructure code
helper.writeToDocx(toWrite=single_count, doc_path="single.docx")
helper.writeToDocx(toWrite=double_count, doc_path="double.docx")
