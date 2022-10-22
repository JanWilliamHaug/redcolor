import docx
from docx.shared import RGBColor

def readtxt("SRS_ACE_Pump_X01.docx"):
    doc = docx.Document("SRS_ACE_Pump_X01.docx")
    fullText = []
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.color.rgb == RGBColor(255, 000, 000):
                fullText.append(run.text)
    return fullText

fullText = readtxt(/Users/Willi/Desktop7/docx red color/SRS_ACE_Pump_X01.docx)
