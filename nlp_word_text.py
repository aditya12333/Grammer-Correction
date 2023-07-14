import spacy
import numpy as np
#import pandas as pd
from docx import Document
from spacy import displacy
import en_core_web_lg
#import en_core_web_sm
import re

import textract
import re
text = textract.process("C:/Users/AdityaPandey/Downloads/s1.docx")
text = text.decode('utf-8')
#print(text)
#type(text)
#text = re.sub(",", "", text)
text = text.replace("sr.", "sr")
text = re.sub(f'(?<=\.)\s+(\w)', lambda m: m.group(1).title(), text)
#text = re.sub(".", lambda p: p.group(0).title(), text)
print(text)
type(text)

nlp = spacy.load('en_core_web_lg')
docs = nlp(text)
print(type(docs))
#print(docs.ents)

display = displacy.render(docs, style = "ent", jupyter = True)

names = []
for ent in docs.ents:
   if ent.label_== "PERSON" or ent.label_ in ["GPE", "LOC", "ORG"]:
     names.append(str(ent))
splr = [i.split() for i in names]
#names = [i.title() for i in splr]
res = []

for i in splr:
  for j in i:
    res.append(j.title())
res

#splt = text.split()
splt = re.split('(\W+?)', text)
#print(splt)
#cap = [i.title() for i in splt]
for x in res:
  #print(x)
  #break
  for i,word1 in enumerate(splt):
    t = word1.title()
    if x == t:
      #print(t)
      splt.pop(i)
      splt.insert(i,x)

       #splt = splt.replace(word1, words)
new_txt = "".join(splt)
print(new_txt)

mydoc = Document()
mydoc.add_paragraph(new_txt)
mydoc.save("C:/Users/AdityaPandey/Downloads/img/converted_text.docx")


import win32com.client

#path = "C:\ThePath\OfYourFolder\WithYourDocuments\\" 
# note the \\ at the end of the path name to prevent a SyntaxError

#Create the Application word
Application=win32com.client.gencache.EnsureDispatch("Word.Application")

# Compare documents
Application.CompareDocuments(Application.Documents.Open("C:/Users/AdityaPandey/Downloads/s1.docx"),
                             Application.Documents.Open("C:/Users/AdityaPandey/Downloads/img/converted_text.docx"))

# Save the comparison document as "Comparison.docx"
Application.ActiveDocument.SaveAs (FileName = "C:/Users/AdityaPandey/Downloads/img/Comparison.docx")
# Don't forget to quit your Application
Application.Quit()



