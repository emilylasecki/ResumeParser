import pandas as pd
import pdfplumber
import pymupdf
import os
from spire.doc import *
from spire.doc.common import *
import spacy
from bs4 import BeautifulSoup
import re
import openpyxl
import xlsxwriter

path = "Resumes" # path to folder where resumes are
excelFile = "ResumeResults.xlsx" # path to excel file

def main():
    directory = os.fsencode(path)

    if os.path.isfile(excelFile) == False:  # create file if it doesn't exist
        df= pd.DataFrame()  
        df.to_excel(excelFile)

    for file in os.listdir(path):
        filename = os.fsdecode(file)
        if filename.endswith(".pdf"):
            iterateFiles(filename)
        #if word doc, first convert to pdf and delete original file
        if filename.endswith(".docx"):
            document = Document()
            document.LoadFromFile(path+"/" + filename)
            filename = filename[:-5]
            document.SaveToFile(path+"/" + filename + ".pdf", FileFormat.PDF) 
            document.Close()
            os.remove(path+"/" + filename + ".docx")
            iterateFiles(filename + ".pdf")
        if filename.endswith(".doc"):
            document = Document()
            document.LoadFromFile(path+"/" + filename)
            filename = filename[:-4]
            document.SaveToFile(path+"/" + filename + ".pdf", FileFormat.PDF)
            document.Close()
            os.remove(path+"/" + filename + ".doc")
            iterateFiles(filename + ".pdf") 

# anaylze each resume
def iterateFiles(filename):

    lines = []
    j=0

    data = pd.read_excel(excelFile)

    doc = pymupdf.open(path+"/" + filename)
    links = ([page.get_links()for page in doc])

    with pdfplumber.open(path+"/" + filename) as pdf:
        for page in pdf.pages: 
            text = page.extract_text()
            for line in text.split('\n'):
                lines.append(line)
                j = j+1

    s = filename
    name = ''.join([i for i in s if not i.isdigit()])
    name = name[:-12]

    for i, line in enumerate(lines):
        # start saving the content in each resume when match keyword: work experience
        if any(word in line for word in ["Experience ", " Experience ", " Experience: ", "Experience", "EXPERIENCE", " EXPERIENCE ", " EXPERIENCE", " EXPERIENCE: ", "Work History", "WORK HISTORY"]):
            experience = lines[i : j] 
            expstr = '\t'.join(experience)
            s = filename
            # formatting specific to the way handshake pulls resumes
            name = ''.join([i for i in s if not i.isdigit()])  # remove any numbers in filename
            name = name[:-12]  # remove the last 12 character from filename - with handshake this leaves only the resume owners name
            exp= afterCarefulConsideration(expstr)
            newexp= newSoup(exp)
            numberOfExp, expTitle = modifyTheContent2(newexp)
            Pandas(name, expstr, links, data, filename, exp, numberOfExp, expTitle)
            break
        else:  # if resume doesn't parse, (no work experience section found, etc.) print nulls for fields
            s = filename
            name = ''.join([i for i in s if not i.isdigit()])
            expstr = "null"
            exp = "null"
            numberOfExp = "null"
            expTitle = "null"
            Pandas(name, expstr, links, data, filename, exp, numberOfExp, expTitle)

# export content to excel sheet
def Pandas(name, expstr, links, data, filename, exp, numberOfExp, expTitle): 

    dict = {'Name': name, 'FileName': filename, 'Experience': expstr, 'Links': links, 'TypeOfSpeech': exp, 'NumberOfExp': numberOfExp, "ExpNames": expTitle}
    df =pd.DataFrame(dict)  
    first_row = df.head(1)
    appended_df = pd.concat([data, first_row], ignore_index= True)
    appended_df.to_excel(excelFile, engine='xlsxwriter', index= False)

# add parts of speech 
def afterCarefulConsideration(expstr):
    nlp = spacy.load("en_core_web_lg")  # what is this?
    expstr = nlp(expstr)
    dataFrameContent= ''
    for token in expstr:
        dataFrameContent += (token.text + " " + token.lemma_+ " " + token.pos_+ " " +  token.tag_+ " " +  token.dep_+ " " +  token.shape_)
    addNewLineChar(dataFrameContent)
    return dataFrameContent

# indents resume content string when a new line appears in the original resume
def addNewLineChar(dataFrameContent):
    modifiedContent = ''
    i=1
    searchfor = "SPACE"
    for word in dataFrameContent.split():
        if searchfor in word: # split the string in two and enter the \n character
            index = dataFrameContent.find('SPACE')
            # take from the first index to searchfor and append \n
            i=i +1

# genuinely don't know what this does but if i remove it this will break
def newSoup(dataFrameContent): 
    with open("output.html", mode='w', encoding='utf-8') as f:
        modifiedContent = ''
        searchfor = "SPACE"
        soup = BeautifulSoup(''.join(dataFrameContent), features="html.parser")
        for i in soup.prettify().split('SPACE _SP dep'):
            f.write(i)
            f.write("\n")
            modifiedContent += ''.join(i) + "\n"
    return modifiedContent

# search for job experiences: if a line starts with a capital letter and immediately proceeding is a bullet point, add it to a string as an experience
def modifyTheContent2(modifiedContent):
    c=0
    i=0
    k=0
    modifyList= ' '
    modify = modifiedContent
    f= modifiedContent.splitlines()
    for i, lines in enumerate(f):
        modify =lines.lstrip()
        try:
            modify2= f[i+1].lstrip()
        except:
            pass
        try:
            if (modify[0].isupper() == True): #
                k=k+1
                # need to count all types of bullet points: going through and grabbing the most common ones
                #if (modify2[0] == "·") or (modify2[0] == "●") or (modify2[0] =="-") or (modify2[0] == "•") or (modify2[0] =="") or (modify2[0] =="") or (modify2[0]== "○") or (modify2[0]=="–") or (modify2[0]=="▪") or (modify2[0]=="▪") or (modify2[0]=="   ") or (modify2[0]=="") or (modify2[0]=="●") or (modify2[0]=="➢"): #▪ ● ➢ ●
                c=c+1
                # print(modify)
                with open("output.html", mode='r', encoding='utf-8') as f:
                    # for text in string, remove PROPN NNP compound nsubj CCONJ CC PUNCT punct etc. - undo the beautiful soup parts of speech
                    keywords = ["PROPN", "NNP", "compound", "nsubj", "CCONJ", "CC","PUNCT", "punct", "cc", "Xxxxx", "conj", ":", "HYPH", "NOUN", "NNS", "NN", "dobj", "Xxxx", "xxx", "XxxXxxXxx", "S ROOT", "appos", "NUM", "XXX", "Xxx", "XXXX", "XxxXxx", "X", "-LRB-", "-RRB-", "nummod dddd", "ROOT", "nmod", "npadvmod", "SYM", "TDD", "ADJ", "VERB", "amod", "VB", "advcl", "&amp;", "CD", " S ", "dddd", "xdddd", "ADP", "prep", "DT", "pobj", "JJ", "nummod", "ADV RB", "DET", "PART TO", "PART POS", "poss"]
                    content = f.readlines()
                    for keyword in keywords:
                        content[i] = content[i].replace(keyword, "")
                modifyList = modifyList + '\n' + content[i]
        except: 
            pass
    print(c)
    newlist = removeConsecutiveDuplicateWords(modifyList)
    return c, newlist 

# remove characters that interupt duplicate search and remove duplicates
def removeConsecutiveDuplicateWords(s):  # FIXME inprovment: make this not case sensitive and proper line breaks
    for row in s:
        s = s.replace(',', '')
        s = s.replace('(', '')
        s = s.replace(')', '')
        s = s.replace('–', '') 
        s = s.replace('-', '')
        s = s.replace(';', '')
        s = s.replace('|', '')
        s = s.replace('/', '')
    # Use a regular expression to replace consecutive duplicate words 
    # \b denotes word boundaries, \s* matches any whitespace (spaces, tabs, newlines) 
    # \1 refers to the first captured group (the word) 
    return re.sub(r'\b(\w+)\b(\s+\1)+', r'\1', s) 

if __name__ == '__main__':
    main()
