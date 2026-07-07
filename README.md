# Difference-Finder
Purpose: takes two txt files of different versions of the same document and compares all text within to find items that are different in each file. putting the text back into a word document with different formatting. text is compared paragraph by paragraph, then word by word. 

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
