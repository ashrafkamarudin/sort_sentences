import helper
import json
import config

def seperateSentenceBasedOnVariable(sentences, variables):
    data = { "in": [], "not": [] }
    variableCount = {}

    for sentence in sentences:
        sentenceWithoutStopWord = helper.removeStopWord(stopWordList=config.stopWords,sentence=sentence)

        singleCount = helper.countSingleWord(sentenceWithoutStopWord, {})
        doubleCount = helper.countDoubleWord(sentenceWithoutStopWord, {})
        
        found, variableCount = helper.isVariableInSentence(sentenceWithoutStopWord, variables, variableCount);
        data["in"].append(sentence) if found else data["not"].append(sentence)
    
    return data, [variableCount, singleCount, doubleCount]

# Step 1: Load required files
sheet = helper.loadSheet(config.file["sheet"]["path"])  # variables sheet
document = helper.loadDocx(config.file["docx"]["path"]) # docx

# step 2`: Extract data from loaded files (sheet)
title = helper.extractColumnFromSheet(row=1, sheet=sheet)
variables = helper.extractVariableFromSheet(columns=title, sheet=sheet)

print(variables)

newList = helper.reformatToNumberedList(
    list= helper.setListLowerBound(
        list= document.body[0][0][0], lowerBound = config.file["docx"]["start_from_row"]
    )
)

# Step 3: Perfrom seperation
data, counts = seperateSentenceBasedOnVariable(sentences=newList, variables= variables)
counts[0] = dict(
    sorted(counts[0].items(), 
    key=lambda item: item[1],
    reverse=True if config.output["variable_count"]["order"] == "DESC" else False)
)

variable_count = helper.transformDictToArrowNotationList(counts[0])
single_count = helper.transformDictToArrowNotationList(counts[1])
double_count = helper.transformDictToArrowNotationList(counts[2])

print (variable_count)

# Step 4: Print Output into docx
helper.writeToDocx(toWrite=data["in"], doc_path= config.output["exist"])
helper.writeToDocx(toWrite=data["not"], doc_path= config.output["not_exist"])
helper.writeToDocx(toWrite=variable_count, doc_path= config.output["variable_count"]["path"])
helper.writeToDocx(toWrite=single_count, doc_path="single.docx")
helper.writeToDocx(toWrite=double_count, doc_path="double.docx")
