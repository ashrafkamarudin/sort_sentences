import docx
from docx2python import docx2python
import xlrd
import re
import string

alhpabets = "abcdefghijklmnopqrstuvwxyz"

def loadSheet(sheet):
    return xlrd.open_workbook(sheet).sheet_by_index(0)

def loadDocx(doc):
    return docx2python(doc)

def extractColumnFromSheet (row, sheet):
    col_dict = {}
    for i in range(sheet.ncols):
        if not sheet.cell_value(1, i):
            continue

        col_dict[i] = sheet.cell_value(1, i)

    return col_dict

def extractRowFromSheet (column, sheet):
    list = []
    for i in range(sheet.nrows):
        if not sheet.cell_value(i, column):
            continue

        list.append(sheet.cell_value(i, column))
    return list

def extractVariableFromSheet(columns, sheet):
    dict = {}
    for key in columns:
        list = extractRowFromSheet(column=key, sheet=sheet)
        dict[columns[key]] = list

    return dict

def isVariableInSentence (needle, haystack, count):
    found = False
    for key in haystack:
        for variable in haystack[key]:
            if variable.lower() in needle.lower():
                count[variable] = count[variable]+needle.lower().count(variable.lower()) if checkKey(count, variable) else 1
                found = True

    return found, count

def writeToDocx (toWrite, doc_path):
    doc = docx.Document()
    for sentence in toWrite:
        doc.add_paragraph(sentence)

    doc.save(doc_path)

def setListLowerBound (list, lowerBound):
    popidx = 0
    while popidx < (lowerBound - 1):   
        list.pop(popidx)
        popidx+=1

    return list

def removeStopWord (stopWordList, sentence):
    for word in stopWordList:
        sentence = sentence.replace(f' {word} ', ' ')

    return sentence

def reformatToNumberedList (list):
    sidx = 0
    sentences = []
    for index, sentence in enumerate(list):
        if sentence[1] not in alhpabets:
            sentences.append(sentence)
            sidx+=1
            continue

        for index, value in enumerate(alhpabets):
            if f"\t{value})" in sentence:
                sentence = sentence.replace(f"\t{value})", f"\t{sidx}.{index+1})")
                sentences.append(sentence)

    return sentences

def checkKey(dict, key): 
    return True if key in dict.keys() else False

def countSingleWord(haystack, output):
    for current in extractWordFromString(haystack):
        if checkKey(output, current):
            output[current]+=1
            continue
        output[current] = 1

    return output

def countDoubleWord(haystack, output):
    previous = ''
    for current in extractWordFromString(haystack):
        if previous == '':
            previous = current
            continue

        key = f"{previous} {current}"
        previous = current

        if checkKey(output, key):
            output[key]+=1
            continue
        output[key] = 1
        
    return output

def extractWordFromString(haystack):
    return re.sub('['+string.punctuation+']', '', haystack).split()[1:]

def transformDictToArrowNotationList(dict):
    var = []
    for key in dict:
        var.append(f"{key} => {dict[key]} ")

    return var