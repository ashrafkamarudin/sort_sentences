
# Todo
# [x] Read sentences from "System Requirement MODEL.docx"
# [x] Read variable from "Variable Model.xlsx"
#    [x] get sentences in structures like 
# [x] Sort these sentences into 2 file based on variables in file "Variable Model.xlsx"
# [x] Count Occurance of "variables"
# [x] Count/List Single Word in "System Requirement MODEL.docx"
# [x] Count/List Double Word (2 word) in "System Requirement MODEL.docx"
# [x] Draw flow / algorithm

import helper
import xlrd
from docx2python import docx2python
import json
import config

def seperateSentenceBasedOnVariable(sentences, variables):
    data = { "in": [], "not": [] }

    for sentence in sentences:
        new_sentence = helper.removeStopWord(stopWordList=config.stopWords,sentence=sentence)

        # single count
        singleCount = helper.countSingleWord(new_sentence, {})

        # doubly count
        doubleCount = helper.countDoubleWord(new_sentence, {})
        
        found, variableCount = helper.isVariableInSentence(new_sentence, variables, count_var={});

        data["in"].append(sentence) if found else data["not"].append(sentence)
    
    return data, [variableCount, singleCount, doubleCount]

# Step 1: Load required files
sheet = xlrd.open_workbook(config.file["sheet"]["path"]).sheet_by_index(0) # variables sheet
document = docx2python(config.file["docx"]["path"]) # docx

# step 2: Extract data from loaded files (sheet)
title = helper.extractColumnFromSheet(row=1, sheet=sheet)
variables = helper.extractVariableFromSheet(columns=title, sheet=sheet)
newList = helper.reformatToNumberedList(list= helper.setListLowerBound(
    list= document.body[0][0][0], lowerBound = config.file["docx"]["start_from_row"])
)

# Step 3: Perfrom seperation
data, counts = seperateSentenceBasedOnVariable(sentences=newList, variables= variables)
counts[0] = dict(sorted(counts[0].items(), key=lambda item: item[1],reverse=True if config.output["variable_count"]["order"] == "DESC" else False))

variable_count = helper.transformDictToArrowNotationList(counts[0])
single_count = helper.transformDictToArrowNotationList(counts[1])
double_count = helper.transformDictToArrowNotationList(counts[2])

# Step 4: Print Output into docx
helper.writeToDocx(toWrite=data["in"], doc_path= config.output["exist"])
helper.writeToDocx(toWrite=data["not"], doc_path= config.output["not_exist"])
helper.writeToDocx(toWrite=variable_count, doc_path= config.output["variable_count"]["path"])
helper.writeToDocx(toWrite=single_count, doc_path="single.docx")
helper.writeToDocx(toWrite=double_count, doc_path="double.docx")
