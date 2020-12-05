
# -------------------------------------------------
#  Config File for the script
# -------------------------------------------------
# path => path of the file
# start_from_row => Will skip row before the set number. e.g = title
#
file = {
    "docx": {
        "path": "System Requirement MODEL.docx",
        "start_from_row": 2
    },
    "sheet": {
        "path": "Variable Model.xlsx"
    }
}

# Output Config
# must end in .docx
# change right hand side only
output = {
    "exist":"in.docx",
    "not_exist":"not.docx"
}