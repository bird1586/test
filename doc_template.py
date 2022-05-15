# -*- coding: utf-8 -*-
"""
Created on Mon May  9 21:37:21 2022

@author: bird1586
"""

from docxtpl import DocxTemplate
import pandas as pd
import requests
import streamlit as st
from docx import Document
from glob import glob 


@st.cache
def get_template():
    url = 'https://github.com/bird1586/test/raw/main/template.docx'
    r = requests.get(url, stream=True)
    open('template.docx', 'wb').write(r.content)
    doc = DocxTemplate("template.docx")
    return doc

    
def format_time(date, time):
    try:
        return "{} {}:00".format(date.strftime('%Y-%m-%d'), 
                                 time.strftime('%H:%M'))
    except :
        return ''
    
def format_date(t):
    try:
        return "{}/{:02d}".format(t.month, t.day)
    except :
        return ''

def parse_info(text):
    try:
        car, phone = text.split('/')
        return car, phone
    except:
        return '', ''

def combine_word_documents(files):
    merged_document = Document()

    for index, file in enumerate(files):
        sub_doc = Document(file)

        # Don't add a page break if you've reached the last file.
        if index < len(files)-1:
           sub_doc.add_page_break()

        for element in sub_doc.element.body:
            merged_document.element.body.append(element)
    merged_document.save('merged.docx')
    
    
doc = get_template()
uploaded_file = st.file_uploader("請上傳EXCEL", type=["xlsx", 'xls'])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, skiprows=[0])
        
    for index, row in df.iterrows():
        context = { 'time' : format_time(row['用車時間'], row.time),
                    'date': format_date(row['用車時間']),
                    'from': row.get('Pickup Address', ''),
                    'to': row.get('DropOff Address', ''),
                    'name': row.get('name', ''),
                    'flight': row.get('Flight\n＊國內', '')
                   }
        doc.render(context)
        doc.save("{}.docx".format(index))
        
    files = glob('*.docx')
    combine_word_documents(files)
    
    st.download_button(
         label="Download {} docx",
         data=open("merged.docx", "rb"),
         file_name='merged.docx',
         mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
         )