import docx
from docx import Document
from docx.shared import RGBColor
import string
import re


filename = "SRS_ACE_Pump_X00.docx"


def readtxt(filename):
    doc = docx.Document(filename)
    fullText = []
    for paras in doc.paragraphs:
        for run in paras.runs:
            if run.font.color.rgb == RGBColor(255, 000, 000):
                fullText.append(run.text)
    return fullText





fullText = readtxt("SRS_ACE_Pump_X00.docx")

s = ''.join(fullText)
print(s)

#o = s.split(']\w+')
#print(o)


#print str(fullText)[1:-1]
#print(fullText)
#print(fullText[5])

report = Document()
#paragraph = report.add_paragraph(fullText[1])
#paragraph = report.add_paragraph(fullText)
#paragraph = report.add_paragraph(s)
#report.save('report1.docx')


w = (s.replace (']', ']\n'))
paragraph = report.add_paragraph("SRS Ace Pump Document")
paragraph = report.add_paragraph(w)


paragraph = report.add_paragraph("PRS_new_pump.docx")

filename2 = "PRS_new_pump.docx"

fullText2 = readtxt("SRS_ACE_Pump_X00.docx")
b = ''.join(fullText2)
w = (b.replace (']', ']\n'))
print(w)

paragraph = report.add_paragraph(w)

report.save('report1.docx')


