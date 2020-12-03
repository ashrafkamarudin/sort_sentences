
# Todo
# [x] Read sentences from "System Requirement MODEL.docx"
# [x] Read variable from "Variable Model.xlsx"
#    [ ] get sentences in structures like 
# [ ] Sort these sentences into 2 file based on variables in file "Variable Model.xlsx"

from pprint import pprint
import helper
import xlrd
from docx2python import docx2python


sentence_loc = "System Requirement MODEL.docx"
variable_loc = "Variable Model.xlsx"

document = docx2python("System Requirement MODEL.docx")

def extractVariableFromSheet(columns, sheet):
    dict = {}

    for key in columns:
        list = helper.extractRowFromSheet(column=key, sheet=sheet)
        dict[columns[key]] = list

    return dict

def seperateSentenceBasedOnVariable(sentences, variables):
    # dict = {}
    for sentence in sentences:
        print(sentence)
        break
        pprint(helper.inDict(needle=sentence, haystack=variables))
    return

sheet = xlrd.open_workbook(variable_loc).sheet_by_index(0)

title = helper.extractColumnFromSheet(row=1, sheet=sheet)
variables = extractVariableFromSheet(columns=title, sheet=sheet)

seperateSentenceBasedOnVariable(sentences=document.body, variables= variables)

# pprint(variables)
