import pandas as pd
from datetime import date
import streamlit as st
from docxtpl import DocxTemplate
from docx import Document
import lxml
from make_document import make_inspection_document
from streamlit_gsheets import GSheetsConnection

conn = st.connection("gsheets", type=GSheetsConnection)
data = conn.read(spreadsheet="https://docs.google.com/spreadsheets/d/1DJXGDW3WJbhsHxBAQ6YtZMEQqp62keCe3QPGd7jRZDs/edit?usp=sharing", usecols=[0, 1,2])
ccode = conn.read(spreadsheet="https://docs.google.com/spreadsheets/d/1DJXGDW3WJbhsHxBAQ6YtZMEQqp62keCe3QPGd7jRZDs/edit?usp=sharing", usecols=[0, 1], worksheet='1362423819')

def set_updatefields_true(docx_path):
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    doc = Document(docx_path)
    # add child to doc.settings element
    element_updatefields = lxml.etree.SubElement(
        doc.settings.element, f"{namespace}updateFields"
    )
    element_updatefields.set(f"{namespace}val", "true")
    doc.save(docx_path)## Heading ##

st.title(":rainbow[Inspection Report Generator]")

# Get input from the user
st.header(':blue[Info]', divider='violet')
# client
st.subheader(':grey[Client]')#, divider='grey')
c1,c2 = st.columns([4,1])   
clients = tuple(data['Client'].unique().tolist())
client_name = c1.selectbox('Client name:',clients)
clientcode = ccode[ccode['Client']==client_name]['Code'][0]
client_code = c2.text_input('Client code:',clientcode)
c3,c4 = st.columns([4,1])
clientlocation = tuple(data[data['Client']==client_name]['Location'].unique().tolist())
client_location = c3.selectbox('Client location:',clientlocation)
unitnumber = tuple(data[(data['Client']==client_name)&(data['Location']==client_location)]['Unit'].tolist())
unit_number = c4.selectbox('Unit number:',unitnumber)
client_logo = st.file_uploader("Upload Client logo:", type=['png','jpeg','jpg'], accept_multiple_files=False)
# inspection
st.subheader(':grey[Inspection]')#, divider='grey')
c5,c6 = st.columns([4,1])
equipment_name = c5.text_input('Equipment name:')
tag_number = c6.text_input('Tag number:')
# st.subheader('Equipment')#, divider='grey')
c7,c8 = st.columns([4,1])
inspection_date = c8.date_input('Inspection date:')
inspection_type = c7.text_input('Inspection type:','TOWER INSPECTION BY ROBOTIC CRAWLER')

# Prepared by
st.subheader(':grey[Authors]')#, divider='grey')
df = pd.DataFrame(
    [
       {"Date": date.today(),'Job':"Prepared by", "Designation": 'NDT Technician', "Name": 'Sakthivel'},
       {"Date": date.today(),'Job': "Reviewed by", "Designation": 'NDT Technician', "Name": 'Kasirajan'},
       {"Date": date.today(),'Job':'Approved By', "Designation": 'Managing Director', "Name": "Dharmaraj"},
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
st.caption("Add '#' before headings,  '$' before subheadings,  '-' before each point")
site_observation = st.text_area(f'Site observation:',"""# Heading
$ Sub-heading
- add point""")

# Upload files
st.header(':blue[Upload files]', divider='violet')
st.subheader(':grey[Overall Inspection Summary]')#,divider='blue')
overall_summary = st.file_uploader("Upload Overall Inspection Summary file:", type=['csv'], accept_multiple_files=False)
st.subheader(':grey[Towershell nominal thickness and height details]')#,divider='blue')
thickness_details = st.file_uploader("Upload Towershell nominal thickness and height details file:", type=['csv'], accept_multiple_files=False)
st.subheader(':grey[Scanning location and orientation details]')#,divider='blue')
scanning_details = st.file_uploader("Upload Scanning location and orientation details file:", type=['csv'], accept_multiple_files=False)
st.subheader(':grey[Shellwise inspection summary]')#,divider='blue')
shellwise_inspection = st.file_uploader("Upload Shellwise inspection summary files:", type=['csv'], accept_multiple_files=True)
st.subheader(':grey[Tower drawings and scanning location]')#,divider='red')
tower_drawing = st.file_uploader("Upload Tower drawings and scanning location pictures:", type=['png','jpeg','jpg'], accept_multiple_files=True)
if tower_drawing is not None:
    st.image(tower_drawing,width=233)

st.subheader(':grey[Shell plate pictures]')#,divider='red')
shell_plate_pics = st.file_uploader("Upload Shell plate pictures:", type=['png','jpeg','jpg'], accept_multiple_files=True)
if shell_plate_pics is not None:
    st.image(shell_plate_pics,width=233)

st.subheader(':grey[Detailed reports]')
result = st.number_input('Number of Sections',min_value=1)
detailed_report = {}
for i in range(result):
    section_name = st.text_input(f'Section Title {i+1}:')
    section = st.file_uploader(f"Upload detailed report files {i+1}:", type=['csv'], accept_multiple_files=True)
    detailed_report[section_name] = section
st.divider()

filename = f"document.docx"
doc = make_inspection_document(client_name, client_location, unit_number, client_code, inspection_date, equipment_name, tag_number, inspection_type, edited_df, result_and_conclusion, site_observation, overall_summary, thickness_details, scanning_details, shellwise_inspection, tower_drawing, shell_plate_pics, detailed_report)
doc.save(rf'Temp\{filename}')

set_updatefields_true(rf'Temp\{filename}')
st.download_button("Generate report" , data=open(rf'Temp\{filename}', "rb").read(), file_name=filename, use_container_width = True, mime = "application/octet-stream" )
