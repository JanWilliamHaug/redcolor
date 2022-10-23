import docx
from docx import Document
from docx.shared import RGBColor

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

#print str(fullText)[1:-1]
#print(fullText)
#print(fullText[5])

report = Document()
#paragraph = report.add_paragraph(fullText[1])
paragraph = report.add_paragraph(fullText)
report.save('report1.docx')

print(fullText)

