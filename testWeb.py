import streamlit as st  

import pandas as pd  

from docx import Document  

from docx.oxml import OxmlElement 

from docx.oxml.ns import qn 

from docx.shared import Pt, RGBColor  

from docx.enum.text import WD_PARAGRAPH_ALIGNMENT  

import base64  

from io import BytesIO 

import os 

# Function to convert the DataFrame to Word document  

 

def mac_num_to_name(mac_num): 

    ## if first two numbers of mac_num are 01, then it is CGS 

    if mac_num[:2] == '05': 

        return 'WPS Government Health Administrators (J-5)' 

    if mac_num[:2] == '06': 

        return 'National Government Services (J-6)' 

    if mac_num[:2] == '08': 

        return 'WPS Government Health Administrators (J-8)' 

    if mac_num[:2] == '15': 

        return 'CGS Administrators (J-15)' 

    if mac_num[:2] == '01': 

        return 'Noridian Healthcare Solutions c/o Cahaba Safeguard Administrators (J-E)' 

    if mac_num[:2] == '02' or mac_num[:2] == '03': 

        return 'Noridian Healthcare Solutions (J-F)' 

    if mac_num[:2] == '04' or mac_num[:2] == '07': 

        return 'Novitas Solutions, Inc. (J-H)' 

    if mac_num[:2] == '10': 

        return 'Palmetto GBA (J-I)' 

    if mac_num[:2] == '13': 

        return 'National Government Services, Inc. (J-K)' 

    if mac_num[:2] == '12': 

        return 'Novitas Solutions, Inc. (J-L)' 

    if mac_num[:2] == '11': 

        return 'Palmetto GBA c/o National Government Services, Inc. (J-M)' 

    if mac_num[:2] == '09': 

        return 'First Coast Service Options, Inc. (J-N)' 

     

 

def format_date(date): 

    year = date[:4] 

    month = date[5:7] 

    day = date[8:10] 

    return f"{month}/{day}/{year}" 

 

def get_issue_content(issue):  

 

    # Function to read content from {issue}.txt file  

    issueformatted = issue.replace(" ", "") 

    filename = f"IssuestoArgs/{issueformatted}.txt"  

    

 

    try:  

 

        with open(filename, 'r') as file:  

 

            content = file.read()  

 

        return content  

 

    except FileNotFoundError:  

 

        return "Issue file not found."  

 

def create_word_document(case_data):  

 

    doc = Document()  

 

    header = doc.add_paragraph('BEFORE THE PROVIDER REIMBURSEMENT REVIEW BOARD') 

 

    header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 

    run = header.runs[0] 

    run.font.size = Pt(9.5) 

    run.font.name = 'Arial' 

 

    p = header._p 

    pPr = p.get_or_add_pPr() 

    pBdr = OxmlElement('w:pBdr')     

    bottom = OxmlElement('w:bottom') 

    bottom.set(qn('w:val'), 'single') 

    bottom.set(qn('w:sz'), '12') 

    bottom.set(qn('w:space'), '1') 

    bottom.set(qn('w:color'), 'auto') 

    pBdr.append(bottom) 

    pPr.append(pBdr) 

 

    case_name = case_data['Case Name'].iloc[0] if 'Case Name' in case_data else 'Case Name not found' 
 
    issue = case_data['Issue'].unique() if 'Issue' in case_data else 'Issue not found'

    if issue == 'Issue not found':
        #split the 'Issue Typ' by comma 
        issue = case_data['Issue Typ'].iloc[0].split(',') if 'Issue Typ' in case_data else 'Issue not found'
        st.write(issue)
        

    provider_numbers = ', '.join(case_data['Provider ID'].unique()) if 'Provider ID' in case_data else 'Provider Numbers not found' 

    provider_num_array = case_data['Provider ID'].unique() if 'Provider ID' in case_data else 'Provider Numbers not found' 

    if len(provider_num_array) > 1: 

        provider_numbers = 'Various' 

    else: 

        pass 

 

    provider_names = ', '.join(case_data['Provider Name'].unique()) if 'Provider Name' in case_data else 'Provider Names not found' 

    provider_name_array = case_data['Provider Name'].unique() if 'Provider Name' in case_data else 'Provider Names not found' 

    if len(provider_name_array) > 1: 

        provider_names = 'Various' 

    else: 

        pass 

 

    case_num = case_data['Case Num'].iloc[0] if 'Case Num' in case_data else 'Case Num not found' 

    mac_num = case_data['MAC'].iloc[0] if 'MAC' in case_data else 'MAC not found' 

    mac_name = mac_num_to_name(mac_num) 

 

    ## get only first 10 characters of the date 

    determination_event_dates = ', '.join([format_date(str(date)[:10]) for date in case_data['Determination Event Date'].unique()]) if 'Determination Event Date' in case_data else 'Determination Event Dates not found' 

    det_event_array = case_data['Determination Event Date'].unique() if 'Determination Event Date' in case_data else 'Determination Event Dates not found' 

    if len(det_event_array) > 1: 

        determination_event_dates = 'Various' 

    else: 

        pass 

 

    ## get only first 10 characters of the date 

    date_of_appeal = format_date(str(case_data['Appeal Date'].iloc[0])[:10]) if 'Appeal Date' in case_data else 'Date of Appeal not found' 

    adj_no = ','.join(case_data['Audit Adj No.'].unique()) if 'Audit Adj No.' in case_data else 'Audit Adj No. not found' 

    if 'Group FYE' in case_data: 

        year = format_date(case_data['Group FYE'].iloc[0]) if 'Group FYE' in case_data else 'FYE not found' 

    else: 

        year = format_date(case_data['FYE'].iloc[0]) if 'FYE' in case_data else 'FYE not found' 

 

 

    table = doc.add_table(rows = 1, cols = 3) 

 

    for cell in table.columns[0].cells: 

        cell.width = Pt(260) 

    for cell in table.columns[1].cells: 

        cell.width = Pt(20) 

    for cell in table.columns[2].cells: 

        cell.width = Pt(260) 

 

     

    cell_left = table.cell(0,0) 

    cell_left.text = f"\nCase Name: {case_name}\n\nProvider Numbers: {provider_numbers}\n\n     Provider Names: {provider_names} \n\n vs. \n\n{mac_name}\n     (Medicare Administrative Contractor)\n\n        and \n\n Federal Specialized Services \n     (Appeals Support Contractor)\n" 

    run = cell_left.paragraphs[0].runs[0] 

    run.font.size = Pt(9.5) 

    run.font.name = 'Arial' 

 

 

    cell_middle = table.cell(0,1) 

    cell_middle.text = ")\n"*25 

    cell_middle.vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 

    cell_middle.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 

    run = cell_middle.paragraphs[0].runs[0] 

    run.font.size = Pt(9.5) 

    run.font.name = 'Arial' 

 

    cell_right = table.cell(0,2) 

    cell_right.text = f"\n\n\n\n\n\n\n\n\n\n\n\nPRRB Case No. {case_num}\n\nFYE: {year[:10]}\n" 

    run = cell_right.paragraphs[0].runs[0] 

    run.font.size = Pt(9.5) 

    run.font.name = 'Arial' 

 

    #line underneath 

 

    line_para = doc.add_paragraph() 

    line_para.add_run() 

    p = line_para._p 

    pPr = p.get_or_add_pPr() 

    pBdr = OxmlElement('w:pBdr') 

    bottom_bdr = OxmlElement('w:bottom') 

    bottom_bdr.set(qn('w:val'), 'single') 

    bottom_bdr.set(qn('w:sz'), '6') 

    bottom_bdr.set(qn('w:space'), '1') 

    bottom_bdr.set(qn('w:color'), 'auto') 

    pBdr.append(bottom_bdr) 

    pPr.append(pBdr) 

 

 

    header = doc.add_paragraph('MEDICARE ADMINISTRATIVE CONTRACTOR’S POSITION PAPER')  

 

     

    header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 

    run = header.runs[0] 

    run.font.size = Pt(9.5) 

    run.font.name = 'Arial' 

 

 

 

    sub = doc.add_paragraph(f"Sumbitted by:\n\n<Name>\n{mac_name}\n<Address>\n<Address line 2>") 

    sub.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 

    run = sub.runs[0] 

    run.font.size = Pt(9.5) 

    run.font.name = 'Arial' 

 

 

    sub = doc.add_paragraph(f"and\n\n<Reviewer Name>\nFederal Specialized Services, LLC\n1701 S. Racine Avenue\nChicago, IL 60608-4058") 

    sub.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 

    run = sub.runs[0] 

    run.font.size = Pt(9.5) 

    run.font.name = 'Arial' 

 

     

 

    ## SECOND PAGE NOW 

    doc.add_page_break() 

     

 

    table1 = doc.add_table(rows = 1, cols = 2) 

 

    for cell in table1.columns[0].cells: 

        cell.width = Pt(500) 

    for cell in table1.columns[1].cells: 

        cell.width = Pt(40) 

     

    cell_left1 = table1.cell(0,0) 

    for cell in table1.columns[0].cells: 

        for paragraph in cell.paragraphs: 

            for run in paragraph.runs: 

                run.font.size = Pt(11) 

                run.font.name = 'Arial' 

                run.font.bold = True 

     

    cell_left1.text = f"TABLE OF CONTENTS" 

    run = cell_left1.paragraphs[0].runs[0] 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

    run.font.bold = True 

 

    cell_right1 = table1.cell(0,1) 

 

    cell_right1.text = f"PAGE" 

    run = cell_right1.paragraphs[0].runs[0] 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

    run.font.bold = True 

 

 

    table = doc.add_table(rows = 1, cols = 2) 

 

    for cell in table.columns[0].cells: 

        cell.width = Pt(500) 

    for cell in table.columns[1].cells: 

        cell.width = Pt(40) 

 

    cell_left = table.cell(0,0) 

 

    cell_left.text = f"\nI. INTRODUCTION\n\nII. ISSUES AND ADJUSTMENTS IN DISPUTE\n\nIII. MAC\'s POSITION" 

    run = cell_left.paragraphs[0].runs[0] 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

 

    cell_right = table.cell(0,1) 

    cell_right.text = "\n1\n\n2\n\n3" 

    run = cell_right.paragraphs[0].runs[0] 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

 

     

    doc.save(f"Case_{case_num}.docx") 

 

    ##NEW PAGE 

    doc.add_page_break() 

 

    header = doc.add_paragraph('I. INTRODUCTION') 

    ##font bold and size color 

    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 

    run = header.runs[0] 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

    run.font.bold = True 

    run.font.color.rgb = RGBColor(0,0,0) 

 

    header = doc.add_paragraph() 

    run = header.add_run() 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

 

    run.text = f"\n\n Case Name: {case_name}\n\nProvider Numbers: {provider_numbers}\n\nLead Contractor: {mac_name}\n\nCalendar Year: {year[-4:]}\n\nPRRB Case Number: {case_num}\n\nDates of Determinations: {determination_event_dates}\n\nDate of Appeal: {date_of_appeal}" 

 

 

    ##NEW PAGE 

    doc.add_page_break() 

 

    header = doc.add_paragraph('II. ISSUES AND ADJUSTMENTS IN DISPUTE') 

 

    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 

    run = header.runs[0] 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

    run.font.bold = True 

    run.font.color.rgb = RGBColor(0,0,0) 

 

    header = doc.add_paragraph() 

    run = header.add_run() 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

     

 

    run.text = f"\n\nIssue(s): {issue[0]}\n\nAdjustment No(s): {adj_no}\n\nApproximate Reimbursement Amount: N/A" 

 

    doc.add_page_break() 

 

    header = doc.add_paragraph('III. MAC\'S POSITION') 

 

    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 

    run = header.runs[0] 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

    run.font.bold = True 

    run.font.color.rgb = RGBColor(0,0,0) 

 

    issue_content = get_issue_content(issue[0]) 

 

    header = doc.add_paragraph(f"Issue 1: {issue_content} \n\n") 

    run = header.add_run() 

    run.font.size = Pt(11) 

    run.font.name = 'Arial' 

 

    i = 1 

    if issue == 'Issue not found':
       pass
    else: 
       while i < len(issue): 

          issue_content = get_issue_content(issue[i]) 

          header = doc.add_paragraph(f"Issue {i+1}: {issue_content} \n\n") 

          run = header.add_run() 

          run.font.size = Pt(11) 

          run.font.name = 'Arial' 

          i += 1 

 

 

 

 

 

    # Save the document to a bytes buffer  

 

    buffer = BytesIO()  

 

    doc.save(buffer)  

 

    return buffer.getvalue()  

 

def string_processing(s): 

    if pd.isnull(s) or s == '': 

        return "Not in the spreadsheet" 

    return str(s).replace('"', '') 

 

def find_case_data(df, case_number): 

    df['Case Num'] = df['Case Num'].map(string_processing) 

    case_data = df[df['Case Num'] == case_number] 

    case_data = case_data.map(string_processing) 

    return case_data 

   

 

# Function to convert binary to download link  

 

def get_download_link(file, filename):  

 

    b64 = base64.b64encode(file).decode()  

 

    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Download file</a>'  

 

    return href  

 

 

 

   

 

# Streamlit UI Components  

 

st.title('Excel Case Finder')  

 

uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])  

 

case_num = st.text_input('Enter Case Number') 

 

create_doc = st.button('Create Document') 

 

 

 

 

   

 

if uploaded_file and case_num and create_doc:  

 

    try:  
        df = pd.read_excel(uploaded_file) 

    except: 

        st.write('Failed to load this file. Make sure it is of type .xlxs or .xls and try again.') 

 

    # Assuming the DataFrame 'df' is now available for processing  

 

     

 

    

     docx_file = create_word_document(find_case_data(df, case_num)) 

    

     st.write('Case not found in the spreadsheet. Please try again with a different case number.') 

 

 

      

 

    st.markdown(get_download_link(docx_file, f'Case_{case_num}.docx'), unsafe_allow_html=True)  

 

if not uploaded_file and create_doc: 

    st.write('Please upload a file') 

 

if not case_num and create_doc: 

    st.write('Please enter a case number') 

 

  

 
