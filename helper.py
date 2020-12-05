import docx

alhpabets = "abcdefghijklmnopqrstuvwxyz"

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

def isVariableInSentence (needle, haystack):
    found = False

    for key in haystack:
        if found: continue # skip
        for variable in haystack[key]:
            if found: continue # skip
            if variable in needle:
                found = True

    return found

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