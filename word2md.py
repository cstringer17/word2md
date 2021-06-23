import docx2txt
import re
import os
import shutil
# extract text
text = docx2txt.process("userguide.docx")
# text = docx2txt.process("userguide.docx", "imagesFromWord") 
pattern = re.compile("[0-9]\.[0-9]")


def deleteOutput(path):
    try:
        shutil.rmtree(path)
    except OSError:
        print ("Deletion of the directory %s failed" % path)
    else:
        print ("Successfully deleted the directory %s" % path)

def makeDir(path):
    try:
        os.mkdir(path)
    except OSError:
        print ("Creation of the directory %s failed" % path)
    else:
        print ("Successfully created the directory %s " % path)

def parseFilename(filename):
    parsed = filename.replace("/", "_")
    parsed = parsed.replace('\t', "")
    parsed = parsed.replace(":","-")
    return parsed

def writemd(filename,content):

    try:
        f = open(parseFilename(filename) + ".md", "x" ,encoding="utf-8")
        f.write(content)
        f.close()
    except FileNotFoundError:
        print("Something went wrong, file not found")   
        print("Failed: " + filename)
    else:
        print("Success: " + filename)

# reset output dir
deleteOutput("output")
makeDir("output")
os.chdir('output')
print("Current working directory: {0}".format(os.getcwdb))

content = ""
currentTitle = "1. Introduction"
for line in text.splitlines():
    if(bool(pattern.search(line)) == 0):
        content += line
    if(bool(pattern.search(line))):
        writemd(currentTitle,content)
        currentTitle = line
        content = ""


    

