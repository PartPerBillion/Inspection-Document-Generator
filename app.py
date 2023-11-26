import streamlit as st
import docx
from docx.shared import Inches
# import os
# from PIL import Image
# import numpy as np
import pandas as pd
# import matplotlib.pyplot as plt
from datetime import date


# Create a Streamlit app
st.title(":rainbow[Inspection Report Generator]")


# Get input from the user
st.header(':blue[Info]', divider='violet')
# client
st.subheader(':grey[Client]')#, divider='grey')
c1,c2 = st.columns([4,1])   
client_name = c1.text_input("Client name:")
report_number = c2.number_input('Report number:',min_value = 0)
c3,c4 = st.columns([4,1])
client_location = c3.text_input("Client location:")
unit_number = c4.number_input('Unit number:',min_value = 0)
uploaded_image = st.file_uploader("Upload Client logo:", type=['png','jpeg','jpg'], accept_multiple_files=False)
# inspection
st.subheader(':grey[Inspection]')#, divider='grey')
c5,c6 = st.columns([4,1])
inspection_date = c6.date_input('Inspection date:')
inspection_type = c5.text_input('Inspection type:')
# st.subheader('Equipment')#, divider='grey')
c7,c8 = st.columns([4,1])
equipment_name = c7.text_input('Equipment name:')
tag_number = c8.number_input('Tag number:',min_value = 0)
# Prepared by
st.subheader(':grey[Authors]')#, divider='grey')
df = pd.DataFrame(
    [
       {"Date": date.today(),'Job':"Prepared by", "Designation": 'NDT Technician', "Name": 'Sakthivel', 'Signature':''},
       {"Date": date.today(),'Job': "Reviewed by", "Designation": 'NDT Technician', "Name": 'Kasi', 'Signature':''},
       {"Date": date.today(),'Job':'Approved By', "Designation": 'Managing Director', "Name": "Dharmaraj", 'Signature':''},
   ]
)
edited_df = st.data_editor(df,hide_index=True,use_container_width=True)


# Content
st.header(':blue[Summary]', divider='violet')
st.subheader(':grey[Result and conclusion]')#, divider='grey')
text_list=[]
st.caption("Add '-' before each point")
result_and_conclusion = st.text_area(f'Result and conclusion:','- add point')
st.subheader(':grey[Site observation]')#, divider='grey')
st.caption("Add '#' before headings,  '##' before subheadings,  '-' before each point")
st.text_area(f'Site observation:',"""# Heading
## Sub-heading
- add point""")


# Upload files
st.header(':blue[Upload files]', divider='violet')
st.subheader(':grey[Overall Inspection Summary]')#,divider='blue')
q = st.file_uploader("Upload Overall Inspection Summary file:", type=['csv','xlsx'], accept_multiple_files=False)
st.subheader(':grey[Towershell nominal thickness and height details]')#,divider='blue')
w = st.file_uploader("Upload Towershell nominal thickness and height details file:", type=['csv','xlsx'], accept_multiple_files=False)
st.subheader(':grey[Scanning location and orientation details]')#,divider='blue')
e = st.file_uploader("Upload Scanning location and orientation details file:", type=['csv','xlsx'], accept_multiple_files=False)
st.subheader(':grey[Shellwise inspection summary]')#,divider='blue')
r = st.file_uploader("Upload Shellwise inspection summary files:", type=['csv','xlsx'], accept_multiple_files=True)
st.subheader(':grey[Tower drawings and scanning location]')#,divider='red')
t = st.file_uploader("Upload Tower drawings and scanning location pictures:", type=['png','jpeg','jpg'], accept_multiple_files=True)
st.subheader(':grey[Shell plate pictures]')#,divider='red')
y = st.file_uploader("Upload Shell plate pictures:", type=['png','jpeg','jpg'], accept_multiple_files=True)
st.subheader(':grey[Detailed reports]')#,divider='blue')
u = st.file_uploader("Upload detailed report files:", type=['csv','xlsx'], accept_multiple_files=True)
st.divider()

doc = docx.Document()
section = doc.sections[0]
header = section.header
records = (
    (3, '101', 'Spam'),
    (7, '422', 'Eggs'),
    (4, '631', 'Spam, spam, eggs, and spam')
)
table = header.add_table(1, 3,width=Inches(8))
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Qty'
hdr_cells[1].text = 'Id'
hdr_cells[2].text = 'Desc'
for qty, id, desc in records:
    row_cells = table.add_row().cells
    row_cells[0].text = str(qty)
    row_cells[1].text = id
    row_cells[2].text = desc
table.style = 'Medium List 1'
doc.add_heading(client_name,0)
doc.add_paragraph(result_and_conclusion)
filename = "generated_document.docx"
doc.save(filename)

# Download the document
st.download_button("Generate report" , data=open(filename, "rb").read(), file_name=filename, mime="application/octet-stream")
