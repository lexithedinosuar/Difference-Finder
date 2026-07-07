"""
Creator: Alexia Piunno
Title: Difference Finder
Purpose: takes two txt files of different verions of the same document and compares all
text within to find items that are different in each file. putting the text back into a word
document with different formatting. text is compaired paragraph by paragraph, then word by word.  
 
input: two txt files
Output: new Docx file with different formatting to represent text comparison
 - black text = text that is the same in both files
 - red text (old word/new word) = text that changes
 - Blue text = additional text (not reviewed because there was nothing to compair to)
 - Green text = number of errors found in the comparison

Next steps:
 - compair sentence by sentence, page by page 
    - straigth from word file?
        - pdf split page by page, each pdf converted to text, iterate through all page files?
 - find matching sentence before comparing text
        -use list indexing methods?
"""
# import Libraries/modules
from wordDocClass import new_document
from paragraphClass import paragraph
from docx import Document
from docx.shared import Pt
import os

#Create the Document for the Differnece report
TitleText= input("heading text? ")
FileName="DifferenceReport"
fullFileName = FileName + ".docx"
diffDoc = new_document(TitleText, FileName )

# variable definition
f1 = paragraph("Sub2.txt")
f2 = paragraph("Sub3.txt")
longest_p_list = 0
shortest_p_list = 0
longest_p_len = 0
ErrorCount = 0

# split test into nested list [[words, in, paragraph, 1.], [words, in paragraph, 2.], ... ,[words, in, paragraph, n.]]
f1_list = f1.splitEverything()
f2_list = f2.splitEverything()


# find Longest list of Paragraphs
if len(f1_list) >= len(f2_list):
    longest_p_len=len(f1_list)
    longest_p_list = f1_list
    shortest_p_list = f2_list
else:
    longest_p_len=len(f2_list)
    longest_p_list = f2_list
    shortest_p_list = f1_list

# iterate through paragraphs
p_index = 0
while p_index < longest_p_len:
    longest_w_list = 0
    shortest_w_list = 0
    w_list_len = 0
    w_index = 0
     
    # prevent indexing error, compare test in paragraphes with index available in both list
    if p_index < len(shortest_p_list):

        f1_paragraph_words = f1_list[p_index]
        f2_paragraph_words = f2_list[p_index]
        
        # start a new paragraph on Difference Report Doc
        diffDoc.newParagraph("")

        # find the longest word list
        if len(f1_paragraph_words) >= len(f2_paragraph_words):
            longest_w_list = f1_paragraph_words
            shortest_w_list = f2_paragraph_words
            w_list_len = len(f1_paragraph_words)
        else:
            longest_w_list = f2_paragraph_words
            shortest_w_list = f1_paragraph_words
            w_list_len = len(f2_paragraph_words)

        # iterate through words
        while w_index < w_list_len:
            # prevent indexing error: compare words with the same index, add all extra words in blue.
            if w_index < len(shortest_w_list):
                w1 = f1_paragraph_words[w_index]
                w2 = f2_paragraph_words[w_index]
                
                #formate matching/different words 
                if w1 == w2:
                    diffDoc.add2Paragraph(w1 + " ")
                else:
                    diffDoc.addRedText(w1 + "/" + w2 + " ")
                    ErrorCount += 1
            else:
                extra_word = longest_w_list[w_index]
                diffDoc.addBlueText(extra_word + " ")
                ErrorCount += 1
        
            w_index += 1
    else: # all paragraphs in the longer list get added in blue without comparison
        diffDoc.newParagraph("")
        diffDoc.addBlueText(longest_p_list[p_index])
        ErrorCount += 1

    p_index += 1

# start a new page
diffDoc.newParagraph("")
diffDoc.addGreenText("Number of Errors: " + str(ErrorCount))
diffDoc.addPageBreak()
print("Number of errors: " + str(ErrorCount))

#Save the document for the difference report
diffDoc.saveFile()
#open the file in word
os.startfile(fullFileName)

print("You have reached the end of your script!")
