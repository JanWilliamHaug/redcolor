import docx
from docx.shared import RGBColor
from docx.shared import RGBColor

filename = "SRS_ACE_Pump_X00.docx"


def readtxt(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.color.rgb == RGBColor(255, 000, 000):
                fullText.append(run.text)
    return fullText


fullText = readtxt("SRS_ACE_Pump_X00.docx")

print(fullText)