
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

# variables sheet
sheet = xlrd.open_workbook(config.variables_path).sheet_by_index(0)

title = helper.extractColumnFromSheet(row=1, sheet=sheet)
variables = extractVariableFromSheet(columns=title, sheet=sheet)

# docx
document = docx2python(config.document_path)
data = seperateSentenceBasedOnVariable(sentences=document.body[0][0][0], variables= variables)

helper.writeToDocx(toWrite=data["in"], doc_path= config.output["exist"])
helper.writeToDocx(toWrite=data["not"], doc_path= config.output["not_exist"])