import pandas as pd
import docx
import numpy as np
import streamlit as st
from PIL import Image
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement, ns
from docxtpl import DocxTemplate

def heading_number_generaator(heading, sub_heading, sub_sub_heading):
    if sub_heading == 0:
        return str(heading)
    elif sub_sub_heading == 0:
        return str(heading)+'.'+str(sub_heading)
    else:
        return str(heading)+'.'+str(sub_heading)+'.'+str(sub_sub_heading)
    
def ahn(hn, level):
    headings = hn.split('.')
    if hn.count('.') == 2:
        heading = int(headings[0])
        sub_heading = int(headings[1])
        sub_sub_heading = int(headings[2])
    elif hn.count('.') == 1:
        heading = int(headings[0])
        sub_heading = int(headings[1])
        sub_sub_heading = 0
    else:
        heading = int(headings[0])
        sub_heading = 0
        sub_sub_heading = 0
    if level == 1:
        heading+=1
        return heading_number_generaator(heading,0,0)
    elif level == 2:
        sub_heading+=1
        return heading_number_generaator(heading,sub_heading,0)
    elif level == 3:
        sub_sub_heading+=1
        return heading_number_generaator(heading,sub_heading,sub_sub_heading)
    else:
        return(hn)

def add_site_observation_to_doc(string,doc,hn):
    rdict = {}
    n=0
    for i in string.split('\n'):
        if len(i)>0:
            if i[0] == '#':
                rdict[f'{n}_h'] = i[2:]
            elif i[0] == '$':
                rdict[f'{n}_s'] = i[2:]
            else:
                rdict[f'{n}_p'] = i
        n+=1
    for k,v in rdict.items():
        if len(v)>0:
            if k[-1] == 'h':
                hn = ahn(hn,2)
                doc.add_heading(f'{hn} {v}',2)
            elif k[-1] == 's':
                hn = ahn(hn,3)
                doc.add_heading(f'{hn} {v}',3)
            else:
                p = doc.add_paragraph()
                p.add_run(v)
    return hn

def table_of_contents(document):
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')  # creates a new element
    fldChar.set(ns.qn('w:fldCharType'), 'begin')  # sets attribute on element
    instrText = OxmlElement('w:instrText')
    instrText.set(ns.qn('xml:space'), 'preserve')  # sets attribute on element
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'   # change 1-3 depending on heading levels you need
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(ns.qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:t')
    fldChar3.text = "Right-click to update field"
    fldChar2.append(fldChar3)
    fldChar4 = OxmlElement('w:fldChar')
    fldChar4.set(ns.qn('w:fldCharType'), 'end')
    r_element = run._r
    r_element.append(fldChar)
    r_element.append(instrText)
    r_element.append(fldChar2)
    r_element.append(fldChar4)
    p_element = paragraph._p
    return document

def create_element(name):
    return OxmlElement(name)

def create_attribute(element, name, value):
    element.set(ns.qn(name), value)

def add_page_number(run):
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')
    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = "PAGE"
    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

def add_table_to_document(df, document):
    if len(df) > 0:
        table = document.add_table(rows=df.shape[0] + 1, cols=df.shape[1])
        table.style = 'Table Grid'
        for j, column in enumerate(df.columns):
            table.cell(0, j).text = column
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                table.cell(i + 1, j).text = str(value)

def make_inspection_document(client_name, client_location, unit_number, report_number, inspection_date, equipment_name, inspection_type, edited_df, result_and_conclusion, site_observation, overall_summary, thickness_details, scanning_details, shellwise_inspection, tower_drawing, shell_plate_pics, detailed_report):
    # Create document
    doc = docx.Document()
    paragraph_format = doc.styles['Normal'].paragraph_format
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(0)
    # add_page_number(doc.sections[0].footer.paragraphs[0].add_run())

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Amasis MT Pro'
    font.size = docx.shared.Pt(11)

    # header
    section = doc.sections[0]
    header = section.header 
    table = header.add_table(1, 3,width=Inches(8))
    hdr_cells = table.rows[0].cells
    # add_page_number(header.paragraphs[0].add_run())
    p = hdr_cells[0].paragraphs[0] 
    format = p.paragraph_format
    format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(r'Logo/ArunTech.jpg', width=Inches(.7))

    p = hdr_cells[1].paragraphs[0] 
    format = p.paragraph_format
    format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f'{client_name} {client_location} {unit_number}')

    p = hdr_cells[2].paragraphs[0] 
    format = p.paragraph_format
    format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(r'Logo/HP.png', width=Inches(1.4))
    table.style = 'Table Grid'

    p = header.add_paragraph()
    p.style = None
    p.add_run(style = None)
    p.add_run(style = None)

    table_0 = doc.add_table(1,2)
    p = table_0.rows[0].cells[0].paragraphs[0]
    format = p.paragraph_format
    # format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f'Report number : {report_number}')
    p = table_0.rows[0].cells[1].paragraphs[0]
    format = p.paragraph_format
    # format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f'Inspection date : {inspection_date}')
    table_0.style = 'Table Grid'

    font.size = docx.shared.Pt(16)
    table_1 = doc.add_table(2,1)
    p = table_1.rows[0].cells[0].paragraphs[0]
    format = p.paragraph_format
    format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f'{equipment_name}')
    p = table_1.rows[1].cells[0].paragraphs[0]
    format = p.paragraph_format
    format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f'{inspection_type}')
    table_1.style = 'Table Grid'
    font.size = docx.shared.Pt(11)

    p = doc.add_paragraph()
    p.style = None
    p.add_run(style = None)
    p.add_run(style = None)

    edited_df.to_csv(rf'Temp\authors.csv')
    authors = pd.read_csv(rf'Temp\authors.csv')
    authors.drop('Unnamed: 0', inplace=True, axis=1)
    add_table_to_document(authors, doc)
    doc.add_page_break()
    header.is_linked_to_previous = True

    section = doc.sections[-1]
    header = section.header 
    table = header.add_table(2, 3,width=Inches(8))
    hdr_cells = table.rows[0].cells
    # add_page_number(header.paragraphs[0].add_run())
    p = hdr_cells[0].paragraphs[0] 
    format = p.paragraph_format
    format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(r'Logo/ArunTech.jpg', width=Inches(.7))

    p = hdr_cells[1].paragraphs[0] 
    format = p.paragraph_format
    format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f'{client_name} {client_location} {unit_number}')

    p = hdr_cells[2].paragraphs[0] 
    format = p.paragraph_format
    format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(r'Logo/HP.png', width=Inches(1.4))
    table.style = 'Table Grid'

    p = header.add_paragraph()
    p.style = None
    p.add_run(style = None)
    p.add_run(style = None)

    doc.add_paragraph('Table of contents')
    doc = table_of_contents(doc)
    doc.add_page_break()
    hn = '1'
    doc.add_heading(f'{hn} Results and conclusion',1)
    doc.add_paragraph(result_and_conclusion)
    hn = ahn(hn,1)
    doc.add_heading(f'{hn} Site observation',1)
    add_site_observation_to_doc(site_observation,doc,hn)

    if overall_summary is not None:
        hn = ahn(hn,1)
        doc.add_heading(f'{hn} Overall Summary',1)
        overall_summary_df = pd.read_csv(overall_summary,encoding='utf-8')
        add_table_to_document(overall_summary_df,doc)

    if thickness_details is not None:
        hn = ahn(hn,1)
        doc.add_heading(f'{hn} Thickness Details',1)
        overall_summary_df = pd.read_csv(thickness_details)
        add_table_to_document(overall_summary_df,doc)
    
    if scanning_details is not None:
        hn = ahn(hn,1)
        doc.add_heading(f'{hn} Scanning Details',1)
        overall_summary_df = pd.read_csv(scanning_details)
        add_table_to_document(overall_summary_df,doc)

    if shellwise_inspection is not None:
        hn = ahn(hn,1)
        doc.add_heading(f'{hn} Shellwise Inspection',1)
        for table in shellwise_inspection:
            overall_summary_df = pd.read_csv(table)
            add_table_to_document(overall_summary_df,doc)
            p = doc.add_paragraph()
            p.style = None
            p.add_run(style = None)
            p.add_run(style = None)

    if tower_drawing is not None:
        hn = ahn(hn,1)
        doc.add_heading(f'{hn} Tower Drawings',1)
        for pic in tower_drawing:
            image = Image.open(pic)
            image.save(rf'Temp\img.png')
            doc.add_picture(rf'Temp\img.png', width = Inches(5))
            p = doc.add_paragraph()
            p.style = None
            p.add_run(style = None)
            p.add_run(style = None)

    if shell_plate_pics is not None:
        hn = ahn(hn,1)
        doc.add_heading(f'{hn} Shellplate Pictures',1)
        for pic in shell_plate_pics:
            image = Image.open(pic)
            image.save(rf'Temp\img.png')
            doc.add_picture(rf'Temp\img.png', width = Inches(5))
            p = doc.add_paragraph()
            p.style = None
            p.add_run(style = None)
            p.add_run(style = None)
    
    if detailed_report is not None:
        hn = ahn(hn,1)
        doc.add_heading(f'{hn} Detailed Report',1)
        for table in detailed_report:
            overall_summary_df = pd.read_csv(table)
            add_table_to_document(overall_summary_df,doc)
            p = doc.add_paragraph()
            p.style = None
            p.add_run(style = None)
            p.add_run(style = None)
    
    return doc
