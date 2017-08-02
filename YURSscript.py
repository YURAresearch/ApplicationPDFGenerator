# Script to generate PDFs for YURA Applications

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import csv
from operator import itemgetter
from PyPDF2 import PdfFileMerger

data_file='data.csv'
output_directory = 'Output'
output_pdf_file = 'CombinedApps.pdf'

def import_data(data_file):
    applicant_data = sorted(list(csv.reader(open(data_file,"rb"))),key = itemgetter(7,8,9,10,11,12,13,14,15)) #sort by fields
    applicant_data.pop() #remove header entry
    for i in range (len(applicant_data)):
        applicant_data[i][0] = str(i+1)
    return applicant_data

def generatePDFfile(data):
    PDFlist = []
    for i in range (len(data)):
        number=data[i][0]
        name=data[i][1]
        email=data[i][2]
        classyear=data[i][3]
        residentialcollege=data[i][4]
        title=data[i][5]
        abstract=data[i][6]
        #rows 7 to 15 are subject areas
        fields = ""
        for j in range (7,16):
            if data[i][j] != "":
                fields += data[i][j] + ', ';
        if fields != "":
            fields = fields[:-2]
        publicity=data[i][16]
        comments=""
        if data[i][17] != "":
            comments="Comments: "+data[i][17]

        pdf_file_name=output_directory+'/'+number+'-'+fields+'.pdf'
        PDFlist.append(pdf_file_name)
        build_pdf(pdf_file_name, build_flowables(stylesheet(), number, title, abstract, fields, comments))
        print number + " = success"
    return PDFlist

from reportlab.platypus import (
    BaseDocTemplate,
    PageTemplate,
    Frame,
    Paragraph
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.colors import (
    black,
    blue,
    white,
    yellow
)

def stylesheet():
    styles= {
        'default': ParagraphStyle(
            'default',
            fontName='Times-Roman',
            fontSize=10,
            leading=12,
            leftIndent=0,
            rightIndent=0,
            firstLineIndent=0,
            alignment=TA_LEFT,
            spaceBefore=0,
            spaceAfter=0,
            bulletFontName='Times-Roman',
            bulletFontSize=10,
            bulletIndent=0,
            textColor= black,
            backColor=None,
            wordWrap=None,
            borderWidth= 0,
            borderPadding= 0,
            borderColor= None,
            borderRadius= None,
            allowWidows= 1,
            allowOrphans= 0,
            textTransform=None,  # 'uppercase' | 'lowercase' | None
            endDots=None,
            splitLongWords=1,
        ),
    }
    styles['title'] = ParagraphStyle(
        'title',
        parent=styles['default'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=42,
        alignment=TA_CENTER,
        textColor=blue,
    )
    styles['alert'] = ParagraphStyle(
        'alert',
        parent=styles['default'],
        leading=14,
        backColor=yellow,
        borderColor=black,
        borderWidth=1,
        borderPadding=5,
        borderRadius=2,
        spaceBefore=10,
        spaceAfter=10,
    )
    return styles

def build_flowables(stylesheet, number, project_title, abstract, fields, comments):
    return [
        Paragraph(number+": "+project_title, stylesheet['title']),
        Paragraph(abstract, stylesheet['default']),
        Paragraph("Fields: "+fields, stylesheet['alert']),
        Paragraph(comments, stylesheet['default']),
    ]


def build_pdf(filename, flowables):
    doc = BaseDocTemplate(filename)
    doc.addPageTemplates(
        [
            PageTemplate(
                frames=[
                    Frame(
                        doc.leftMargin,
                        doc.bottomMargin,
                        doc.width,
                        doc.height,
                        id=None
                    ),
                ]
            ),
        ]
    )
    doc.build(flowables)

def mergePDFs(pdfList):
    merger = PdfFileMerger()
    for pdf in pdfList:
        merger.append(open(pdf, 'rb'))
    with open(output_pdf_file, 'wb') as fout:
        merger.write(fout)

#actual running script
appdata = import_data(data_file) #get application data
mergePDFs(generatePDFfile(appdata)) #generate PDF outputs and merge together

with open("personaldata.csv", "wb") as f:
    writer = csv.writer(f)
    writer.writerow(['Number','Name','Email','Class Year','Residential College','Title'])
    for i in range(len(appdata)):
        writer.writerow([appdata[i][0],appdata[i][1],appdata[i][2],appdata[i][3],appdata[i][4],appdata[i][5]])
