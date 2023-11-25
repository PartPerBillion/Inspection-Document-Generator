import streamlit as st
# import docx
# import os
# from PIL import Image
# import numpy as np
import pandas as pd
# import matplotlib.pyplot as plt
from datetime import date


# Create a Streamlit app
st.title("Inspection Report Generator")

# Get input from the user
st.header('Info', divider='rainbow')
# client
st.subheader('Client', divider='green')
c1,c2 = st.columns([4,1])
client_name = c1.text_input("Client name:")
report_number = c2.number_input('Report number:',min_value = 0)
c3,c4 = st.columns([4,1])
client_location = c3.text_input("Client location:")
unit_number = c4.number_input('Unit number:',min_value = 0)
uploaded_image = st.file_uploader("Upload Client logo:", accept_multiple_files=False)
# inspection
st.subheader('Inspection', divider='green')
c5,c6 = st.columns([4,1])
inspection_date = c6.date_input('Inspection date:')
inspection_type = c5.text_input('Inspection type:')
# st.subheader('Equipment', divider='green')
c7,c8 = st.columns([4,1])
equipment_name = c7.text_input('Equipment name:')
tag_number = c8.number_input('Tag number:',min_value = 0)
# Prepared by
st.subheader('Authors', divider='green')
df = pd.DataFrame(
    [
       {"Date": date.today(),'Job':"Prepared by", "Designation": 'NDT Technician', "Name": 'Sakthivel'},
       {"Date": date.today(),'Job': "Reviewed by", "Designation": 'NDT Technician', "Name": 'Kasi'},
       {"Date": date.today(),'Job':'Approved By', "Designation": 'Managing Director', "Name": "Dharmaraj"},
   ]
)
edited_df = st.data_editor(df,hide_index=True,use_container_width=True)
st.divider()

# Content
st.header('Summary', divider='rainbow')
st.subheader('Result and conclusion', divider='green')
text_list=[]
rcfbutton = st.button('Submit Result and conclusion',type='primary')
if not rcfbutton:
    rbutton = st.button('Add result')
    if rbutton:
        point = st.text_input('type here:','- add a point')
        text_list.append(point)
        print(text_list)
else:
    st.write('Completed Result and conclusion')
st.text_area(f'Result and conclusion:')
st.subheader('Site observation', divider='green')
sofbutton = st.button('Submit Site observation',type='primary')
if not sofbutton:
    c9,c10,c11 = st.columns([1,1,1])
    hbutton = c9.button('Add heading')
    shbutton = c10.button('Add sub-heading')
    tbutton = c11.button('Add text')
    if hbutton:
        x = st.text_input('add heading')
else:
    st.write('Completed Site observation')
st.text_area(f'Site observation:')
st.divider()

# Upload files
st.header('Upload files', divider='rainbow')
st.subheader('Overall Inspection Summary',divider='blue')
q = st.file_uploader("Upload Overall Inspection Summary file:", accept_multiple_files=False)
st.subheader('Towershell nominal thickness and height details',divider='blue')
w = st.file_uploader("Upload Towershell nominal thickness and height details file:", accept_multiple_files=False)
st.subheader('Scanning location and orientation details',divider='blue')
e = st.file_uploader("Upload Scanning location and orientation details file:", accept_multiple_files=False)
st.subheader('Shellwise inspection summary',divider='blue')
r = st.file_uploader("Upload Shellwise inspection summary file:", accept_multiple_files=True)
st.subheader('Tower drawings and scanning location',divider='red')
t = st.file_uploader("Upload Tower drawings and scanning location pictures:", accept_multiple_files=True)
st.subheader('Shell plate pictures',divider='red')
y = st.file_uploader("Upload Shell plate pictures:", accept_multiple_files=True)
st.subheader('Detailed reports',divider='blue')
u = st.file_uploader("Upload detailed report files:", accept_multiple_files=True)
st.divider()

# Download the document
st.button('Generate Report')
# st.download_button("Generate report" , data=open(filename, "rb").read(), file_name=filename, mime="application/octet-stream")
