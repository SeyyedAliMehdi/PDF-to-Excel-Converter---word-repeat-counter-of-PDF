import re

import PyPDF2
import pandas as pd
import string

myDictionary = dict()
# creating a pdf file object
pdfFileObj = open('sample.pdf', 'rb')

# creating a pdf reader object
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

# printing number of pages in pdf file
print(f"Number of PDFs page: {pdfReader.numPages}")

for i in range(pdfReader.numPages):
    # creating a page object
    pageObj = pdfReader.getPage(i)
    # extracting text from page
    text = pageObj.extractText()

    for char in string.punctuation:
        text = text.replace(char, '')
    text = text.replace('”', "")
    text = text.replace('“', "")

    arr = text.lower().split(" ")
    for j in arr:
        if j not in myDictionary:
            myDictionary[j] = 1
        else:
            myDictionary[j] += 1

# closing the pdf file object
pdfFileObj.close()

print(myDictionary)
# convert into dataframe
df = pd.DataFrame(list(myDictionary.items()), columns=['Word', 'Repeat'])

# convert into excel
df.to_excel("words.xlsx", index=False)
# print("Dictionary converted into excel...")

print("Dictionary successfully converted to Excel.")
