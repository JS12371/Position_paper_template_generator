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
import glob

# Function to convert the DataFrame to Word document  
def mac_num_to_name(mac_num): 
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

def get_issue_content(issue, dest_doc, selected_argument):
    issueformatted = issue.replace(" ", "")
    filename = f"IssuestoArgs/{issueformatted}{selected_argument}.docx"
    if not os.path.exists(filename):
        return f"{issue}"
    try:
        doc1 = Document(filename)
        content = copy_paragraphs(doc1, dest_doc)
        return content
    except Exception as e:
        return f"Error processing issue file: {e}"

def get_possible_arguments(issue):
    issueformatted = issue.replace(" ", "")
    files = glob.glob(f"IssuestoArgs/{issueformatted}*.docx")
    arguments = [os.path.basename(f).replace(f"{issueformatted}", "").replace(".docx", "") for f in files]
    return arguments

def copy_paragraph_format(src_paragraph, dest_paragraph):
    dest_paragraph.alignment = src_paragraph.alignment
    dest_paragraph.style = src_paragraph.style
    
    # Copy paragraph formatting (indentation, spacing)
    if src_paragraph.paragraph_format.left_indent:
        dest_paragraph.paragraph_format.left_indent = src_paragraph.paragraph_format.left_indent
    if src_paragraph.paragraph_format.right_indent:
        dest_paragraph.paragraph_format.right_indent = src_paragraph.paragraph_format.right_indent
    if src_paragraph.paragraph_format.first_line_indent:
        dest_paragraph.paragraph_format.first_line_indent = src_paragraph.paragraph_format.first_line_indent
    if src_paragraph.paragraph_format.space_before:
        dest_paragraph.paragraph_format.space_before = src_paragraph.paragraph_format.space_before
    if src_paragraph.paragraph_format.space_after:
        dest_paragraph.paragraph_format.space_after = src_paragraph.paragraph_format.space_after
    if src_paragraph.paragraph_format.line_spacing:
        dest_paragraph.paragraph_format.line_spacing = src_paragraph.paragraph_format.line_spacing

# Function to copy runs from one paragraph to another
def copy_runs(src_paragraph, dest_paragraph):
    for run in src_paragraph.runs:
        dest_run = dest_paragraph.add_run(run.text)
        dest_run.bold = run.bold
        dest_run.italic = run.italic
        dest_run.underline = run.underline
        dest_run.font.size = run.font.size
        dest_run.font.name = run.font.name
        if run.font.color and run.font.color.rgb:
            dest_run.font.color.rgb = run.font.color.rgb

# Function to copy paragraphs from source to destination
def copy_paragraphs(src, dest):
    for paragraph in src.paragraphs:
        dest_paragraph = dest.add_paragraph()
        copy_paragraph_format(paragraph, dest_paragraph)
        copy_runs(paragraph, dest_paragraph)

def create_word_document(case_data, selected_arguments):  
    doc = Document()  
    header = doc.add_paragraph('BEFORE THE PROVIDER REIMBURSEMENT REVIEW BOARD') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
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

    case_num = case_data['Case Num'].iloc[0] if 'Case Num' in case_data else 'Case Num not found'
    case_name = case_data['Case Name'].iloc[0] if 'Case Name' in case_data else '<input provider name>' 
    issue = case_data['Issue'] if 'Issue' in case_data else ['Issue not found']
    cloneissue = [iss for iss in issue]

    tempissue = [i for i in issue]
    issue = tempissue

    transferred_to_case = case_data['Transferred to Case #'] if 'Transferred to Case #' in case_data else ['transferred to case not found']
    temptransferred_to_case = [i for i in transferred_to_case]
    transferred_to_case = temptransferred_to_case

    group_mode = False
    if case_num.endswith('G') or case_num.endswith('C'):
        group_mode = True

    i = 0
    if len(issue) != 1:
        while i < len(issue):
            if transferred_to_case[i] != 'Not in the spreadsheet':
                issue[i] = f"Transferred to case {transferred_to_case[i]}"
            i += 1

    issue = list(dict.fromkeys(issue))
    if issue[0] == 'Issue not found':
        try:
            issue = case_data['Issue Typ'].iloc[0].split(',') if 'Issue Typ' in case_data else ['Issue not found']
        except:
            pass

    provider_numbers = ', '.join(case_data['Provider ID'].unique()) if 'Provider ID' in case_data else 'Provider Numbers not found' 
    provider_num_array = case_data['Provider ID'].unique() if 'Provider ID' in case_data else 'Provider Numbers not found' 
    if len(provider_num_array) > 1: 
        provider_numbers = "Various"
    else: 
        pass 

    provider_names = ', '.join(case_data['Provider Name'].unique()) if 'Provider Name' in case_data else 'Provider Names not found' 
    provider_name_array = case_data['Provider Name'].unique() if 'Provider Name' in case_data else 'Provider Names not found' 
    if len(provider_name_array) > 1: 
        provider_names = "Various"
    else: 
        pass 

    case_num = case_data['Case Num'].iloc[0] if 'Case Num' in case_data else 'Case Num not found' 
    mac_num = case_data['MAC'].iloc[0] if 'MAC' in case_data else 'MAC not found' 
    mac_name = mac_num_to_name(mac_num) 

    determination_event_dates = ', '.join([format_date(str(date)[:10]) for date in case_data['Determination Event Date'].unique()]) if 'Determination Event Date' in case_data else 'Determination Event Dates not found' 
    det_event_array = case_data['Determination Event Date'].unique() if 'Determination Event Date' in case_data else 'Determination Event Dates not found' 
    if len(det_event_array) > 1: 
        determination_event_dates = 'Various' 
    else: 
        pass 

    if issue[0].startswith('Transfer'):
        issue.remove(issue[0])
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
    cell_left.text = f"\n{case_name}\n\nProvider Numbers: {provider_numbers}\n\n     Provider Names: {provider_names} \n\n vs. \n\n{mac_name}\n     (Medicare Administrative Contractor)\n\n        and \n\n Federal Specialized Services \n     (Appeals Support Contractor)\n" 
    run = cell_left.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    cell_middle = table.cell(0,1) 
    cell_middle.text = ")\n"*16
    cell_middle.vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
    cell_middle.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
    run = cell_middle.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    cell_right = table.cell(0,2) 
    cell_right.text = f"\n\n\n\n\n\nPRRB Case No. {case_num}\n\nFYE: {year[:10]}\n" 
    run = cell_right.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

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
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    sub = doc.add_paragraph(f"Submitted by:\n\n<Name>\n{mac_name}\n<Address>\n<Address line 2>") 
    sub.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = sub.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    sub = doc.add_paragraph(f"and\n\n<Reviewer Name>\nFederal Specialized Services, LLC\n1701 S. Racine Avenue\nChicago, IL 60608-4058") 
    sub.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = sub.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

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
                run.font.name = 'Cambria (Body)' 
                run.font.bold = True 

    cell_left1.text = f"TABLE OF CONTENTS" 
    run = cell_left1.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 

    cell_right1 = table1.cell(0,1) 
    cell_right1.text = f"PAGE" 
    run = cell_right1.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 

    table = doc.add_table(rows = 1, cols = 2) 
    for cell in table.columns[0].cells: 
        cell.width = Pt(500) 
    for cell in table.columns[1].cells: 
        cell.width = Pt(40) 

    cell_left = table.cell(0,0) 
    cell_left.text = f"\nI. INTRODUCTION\n\nII. ISSUES AND ADJUSTMENTS IN DISPUTE\n\nIII. MAC\'s POSITION\n\nIV. CITATION OF PROGRAM LAWS, REGULATIONS, INSTRUCTIONS, AND CASES\n\nV. EXHIBITS" 
    run = cell_left.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    cell_right = table.cell(0,1) 
    cell_right.text = "\n1\n\n2\n\n3\n\n?\n\n?" 
    run = cell_right.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    doc.save(f"Case_{case_num}.docx") 
    doc.add_page_break() 

    header = doc.add_paragraph('I. INTRODUCTION') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 
    run.font.color.rgb = RGBColor(0,0,0) 

    header = doc.add_paragraph() 
    run = header.add_run() 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.text = f"\n\n Case Name: {case_name}\n\nProvider Numbers: {provider_numbers}\n\nLead Contractor: {mac_name}\n\nCalendar Year: {year[-4:]}\n\nPRRB Case Number: {case_num}\n\nDates of Determinations: {determination_event_dates}\n\nDate of Appeal: {date_of_appeal}" 

    doc.add_page_break() 

    header = doc.add_paragraph('II. ISSUES AND ADJUSTMENTS IN DISPUTE') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 
    run.font.color.rgb = RGBColor(0,0,0) 

    header = doc.add_paragraph() 
    run = header.add_run() 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    for i in range(len(issue)):
        if group_mode:
            if len(adj_no) > 1:
                adj_no = "Various"
            header = doc.add_paragraph(f"\nIssue: {issue[i]}\n\nAdjustment No(s): {adj_no}\n\nApproximate Reimbursement Amount: N/A\n")
            run = header.runs[0] 
            run.font.size = Pt(11) 
            run.font.name = 'Cambria (Body)' 
        else:
            if len(adj_no) > 1:
                adj_no = "Various"
            header = doc.add_paragraph(f"\nIssue {i + 1}: {cloneissue[i]}\n\nDisposition: {issue[i]}\n\nAdjustment No(s): {adj_no}\n\nApproximate Reimbursement Amount: N/A\n")
            run = header.runs[0] 
            run.font.size = Pt(11) 
            run.font.name = 'Cambria (Body)'

    doc.add_page_break() 

    header = doc.add_paragraph('III. MAC\'S POSITION') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 
    run.font.color.rgb = RGBColor(0,0,0) 

    i = 0 
    issue = list(filter(None, issue))
    if issue[0] == 'Issue not found':
        pass
    else: 
        if case_num[-1] == 'G' or case_num[-1] == 'C':
            header = doc.add_paragraph("Issue: ")
            run = header.add_run()
            run.font.size = Pt(11)
            run.font.name = 'Cambria (Body)'
            if issue[0].startswith('Transfer'):
                issue_content = issue[1]
            else:
                issue_content = get_issue_content(issue[0], doc, selected_arguments[0])
            header = doc.add_paragraph(f"{issue_content}")
            run = header.add_run()
            run.font.size = Pt(11)
            run.font.name = 'Cambria (Body)'
        else:
            while i < len(issue):
                header = doc.add_paragraph(f"\n\nIssue {i+1}: ")
                run = header.add_run() 
                run.font.size = Pt(11) 
                run.font.name = 'Cambria (Body)' 
                issue_content = get_issue_content(issue[i], doc, selected_arguments[i]) 
                header = doc.add_paragraph(f"{issue_content} \n\n") 
                run = header.add_run() 
                run.font.size = Pt(11) 
                run.font.name = 'Cambria (Body)' 
                i += 1 

    for paragraph in doc.paragraphs:
        if 'None' in paragraph.text:
            paragraph.text = paragraph.text.replace('None', '', 1)

    buffer = BytesIO()  
    doc.save(buffer)  
    return buffer.getvalue()  

def string_processing(s): 
    if pd.isnull(s) or s == '': 
        return "Not in the spreadsheet" 
    return str(s).replace('"', '') 

def find_case_data(df, case_number): 
    case_number = case_number.upper()
    df['Case Num'] = df['Case Num'].map(string_processing) 
    case_data = df[df['Case Num'] == case_number] 
    case_data = case_data.map(string_processing) 
    return case_data 

def get_download_link(file, filename): 
    b64 = base64.b64encode(file).decode() 
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Download file</a>' 
    return href 

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
import glob

# Function to convert the DataFrame to Word document  
def mac_num_to_name(mac_num): 
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

def get_issue_content(issue, dest_doc, selected_argument):
    issueformatted = issue.replace(" ", "")
    filename = f"IssuestoArgs/{issueformatted}{selected_argument}.docx"
    if not os.path.exists(filename):
        return f"{issue}"
    try:
        doc1 = Document(filename)
        content = copy_paragraphs(doc1, dest_doc)
        return content
    except Exception as e:
        return f"Error processing issue file: {e}"

def get_possible_arguments(issue):
    issueformatted = issue.replace(" ", "")
    files = glob.glob(f"IssuestoArgs/{issueformatted}*.docx")
    arguments = [os.path.basename(f).replace(f"{issueformatted}", "").replace(".docx", "") for f in files]
    return arguments

def copy_paragraph_format(src_paragraph, dest_paragraph):
    dest_paragraph.alignment = src_paragraph.alignment
    dest_paragraph.style = src_paragraph.style
    
    # Copy paragraph formatting (indentation, spacing)
    if src_paragraph.paragraph_format.left_indent:
        dest_paragraph.paragraph_format.left_indent = src_paragraph.paragraph_format.left_indent
    if src_paragraph.paragraph_format.right_indent:
        dest_paragraph.paragraph_format.right_indent = src_paragraph.paragraph_format.right_indent
    if src_paragraph.paragraph_format.first_line_indent:
        dest_paragraph.paragraph_format.first_line_indent = src_paragraph.paragraph_format.first_line_indent
    if src_paragraph.paragraph_format.space_before:
        dest_paragraph.paragraph_format.space_before = src_paragraph.paragraph_format.space_before
    if src_paragraph.paragraph_format.space_after:
        dest_paragraph.paragraph_format.space_after = src_paragraph.paragraph_format.space_after
    if src_paragraph.paragraph_format.line_spacing:
        dest_paragraph.paragraph_format.line_spacing = src_paragraph.paragraph_format.line_spacing

# Function to copy runs from one paragraph to another
def copy_runs(src_paragraph, dest_paragraph):
    for run in src_paragraph.runs:
        dest_run = dest_paragraph.add_run(run.text)
        dest_run.bold = run.bold
        dest_run.italic = run.italic
        dest_run.underline = run.underline
        dest_run.font.size = run.font.size
        dest_run.font.name = run.font.name
        if run.font.color and run.font.color.rgb:
            dest_run.font.color.rgb = run.font.color.rgb

# Function to copy paragraphs from source to destination
def copy_paragraphs(src, dest):
    for paragraph in src.paragraphs:
        dest_paragraph = dest.add_paragraph()
        copy_paragraph_format(paragraph, dest_paragraph)
        copy_runs(paragraph, dest_paragraph)

def create_word_document(case_data, selected_arguments):  
    doc = Document()  
    header = doc.add_paragraph('BEFORE THE PROVIDER REIMBURSEMENT REVIEW BOARD') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
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

    case_num = case_data['Case Num'].iloc[0] if 'Case Num' in case_data else 'Case Num not found'
    case_name = case_data['Case Name'].iloc[0] if 'Case Name' in case_data else '<input provider name>' 
    issue = case_data['Issue'] if 'Issue' in case_data else ['Issue not found']
    cloneissue = [iss for iss in issue]

    tempissue = [i for i in issue]
    issue = tempissue

    transferred_to_case = case_data['Transferred to Case #'] if 'Transferred to Case #' in case_data else ['transferred to case not found']
    temptransferred_to_case = [i for i in transferred_to_case]
    transferred_to_case = temptransferred_to_case

    group_mode = False
    if case_num.endswith('G') or case_num.endswith('C'):
        group_mode = True

    i = 0
    if len(issue) != 1:
        while i < len(issue):
            if transferred_to_case[i] != 'Not in the spreadsheet':
                issue[i] = f"Transferred to case {transferred_to_case[i]}"
            i += 1

    issue = list(dict.fromkeys(issue))
    if issue[0] == 'Issue not found':
        try:
            issue = case_data['Issue Typ'].iloc[0].split(',') if 'Issue Typ' in case_data else ['Issue not found']
        except:
            pass

    provider_numbers = ', '.join(case_data['Provider ID'].unique()) if 'Provider ID' in case_data else 'Provider Numbers not found' 
    provider_num_array = case_data['Provider ID'].unique() if 'Provider ID' in case_data else 'Provider Numbers not found' 
    if len(provider_num_array) > 1: 
        provider_numbers = "Various"
    else: 
        pass 

    provider_names = ', '.join(case_data['Provider Name'].unique()) if 'Provider Name' in case_data else 'Provider Names not found' 
    provider_name_array = case_data['Provider Name'].unique() if 'Provider Name' in case_data else 'Provider Names not found' 
    if len(provider_name_array) > 1: 
        provider_names = "Various"
    else: 
        pass 

    case_num = case_data['Case Num'].iloc[0] if 'Case Num' in case_data else 'Case Num not found' 
    mac_num = case_data['MAC'].iloc[0] if 'MAC' in case_data else 'MAC not found' 
    mac_name = mac_num_to_name(mac_num) 

    determination_event_dates = ', '.join([format_date(str(date)[:10]) for date in case_data['Determination Event Date'].unique()]) if 'Determination Event Date' in case_data else 'Determination Event Dates not found' 
    det_event_array = case_data['Determination Event Date'].unique() if 'Determination Event Date' in case_data else 'Determination Event Dates not found' 
    if len(det_event_array) > 1: 
        determination_event_dates = 'Various' 
    else: 
        pass 

    if issue[0].startswith('Transfer'):
        issue.remove(issue[0])
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
    cell_left.text = f"\n{case_name}\n\nProvider Numbers: {provider_numbers}\n\n     Provider Names: {provider_names} \n\n vs. \n\n{mac_name}\n     (Medicare Administrative Contractor)\n\n        and \n\n Federal Specialized Services \n     (Appeals Support Contractor)\n" 
    run = cell_left.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    cell_middle = table.cell(0,1) 
    cell_middle.text = ")\n"*16
    cell_middle.vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
    cell_middle.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
    run = cell_middle.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    cell_right = table.cell(0,2) 
    cell_right.text = f"\n\n\n\n\n\nPRRB Case No. {case_num}\n\nFYE: {year[:10]}\n" 
    run = cell_right.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

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
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    sub = doc.add_paragraph(f"Submitted by:\n\n<Name>\n{mac_name}\n<Address>\n<Address line 2>") 
    sub.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = sub.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    sub = doc.add_paragraph(f"and\n\n<Reviewer Name>\nFederal Specialized Services, LLC\n1701 S. Racine Avenue\nChicago, IL 60608-4058") 
    sub.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = sub.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

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
                run.font.name = 'Cambria (Body)' 
                run.font.bold = True 

    cell_left1.text = f"TABLE OF CONTENTS" 
    run = cell_left1.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 

    cell_right1 = table1.cell(0,1) 
    cell_right1.text = f"PAGE" 
    run = cell_right1.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 

    table = doc.add_table(rows = 1, cols = 2) 
    for cell in table.columns[0].cells: 
        cell.width = Pt(500) 
    for cell in table.columns[1].cells: 
        cell.width = Pt(40) 

    cell_left = table.cell(0,0) 
    cell_left.text = f"\nI. INTRODUCTION\n\nII. ISSUES AND ADJUSTMENTS IN DISPUTE\n\nIII. MAC\'s POSITION\n\nIV. CITATION OF PROGRAM LAWS, REGULATIONS, INSTRUCTIONS, AND CASES\n\nV. EXHIBITS" 
    run = cell_left.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    cell_right = table.cell(0,1) 
    cell_right.text = "\n1\n\n2\n\n3\n\n?\n\n?" 
    run = cell_right.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    doc.save(f"Case_{case_num}.docx") 
    doc.add_page_break() 

    header = doc.add_paragraph('I. INTRODUCTION') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 
    run.font.color.rgb = RGBColor(0,0,0) 

    header = doc.add_paragraph() 
    run = header.add_run() 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.text = f"\n\n Case Name: {case_name}\n\nProvider Numbers: {provider_numbers}\n\nLead Contractor: {mac_name}\n\nCalendar Year: {year[-4:]}\n\nPRRB Case Number: {case_num}\n\nDates of Determinations: {determination_event_dates}\n\nDate of Appeal: {date_of_appeal}" 

    doc.add_page_break() 

    header = doc.add_paragraph('II. ISSUES AND ADJUSTMENTS IN DISPUTE') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 
    run.font.color.rgb = RGBColor(0,0,0) 

    header = doc.add_paragraph() 
    run = header.add_run() 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 

    for i in range(len(issue)):
        if group_mode:
            if len(adj_no) > 1:
                adj_no = "Various"
            header = doc.add_paragraph(f"\nIssue: {issue[i]}\n\nAdjustment No(s): {adj_no}\n\nApproximate Reimbursement Amount: N/A\n")
            run = header.runs[0] 
            run.font.size = Pt(11) 
            run.font.name = 'Cambria (Body)' 
        else:
            if len(adj_no) > 1:
                adj_no = "Various"
            header = doc.add_paragraph(f"\nIssue {i + 1}: {cloneissue[i]}\n\nDisposition: {issue[i]}\n\nAdjustment No(s): {adj_no}\n\nApproximate Reimbursement Amount: N/A\n")
            run = header.runs[0] 
            run.font.size = Pt(11) 
            run.font.name = 'Cambria (Body)'

    doc.add_page_break() 

    header = doc.add_paragraph('III. MAC\'S POSITION') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Cambria (Body)' 
    run.font.bold = True 
    run.font.color.rgb = RGBColor(0,0,0) 

    i = 0 
    issue = list(filter(None, issue))
    if issue[0] == 'Issue not found':
        pass
    else: 
        if case_num[-1] == 'G' or case_num[-1] == 'C':
            header = doc.add_paragraph("Issue: ")
            run = header.add_run()
            run.font.size = Pt(11)
            run.font.name = 'Cambria (Body)'
            if issue[0].startswith('Transfer'):
                issue_content = issue[1]
            else:
                issue_content = get_issue_content(issue[0], doc, selected_arguments[0])
            header = doc.add_paragraph(f"{issue_content}")
            run = header.add_run()
            run.font.size = Pt(11)
            run.font.name = 'Cambria (Body)'
        else:
            while i < len(issue):
                header = doc.add_paragraph(f"\n\nIssue {i+1}: ")
                run = header.add_run() 
                run.font.size = Pt(11) 
                run.font.name = 'Cambria (Body)' 
                issue_content = get_issue_content(issue[i], doc, selected_arguments[i]) 
                header = doc.add_paragraph(f"{issue_content} \n\n") 
                run = header.add_run() 
                run.font.size = Pt(11) 
                run.font.name = 'Cambria (Body)' 
                i += 1 

    for paragraph in doc.paragraphs:
        if 'None' in paragraph.text:
            paragraph.text = paragraph.text.replace('None', '', 1)

    buffer = BytesIO()  
    doc.save(buffer)  
    return buffer.getvalue()  

def string_processing(s): 
    if pd.isnull(s) or s == '': 
        return "Not in the spreadsheet" 
    return str(s).replace('"', '') 

def find_case_data(df, case_number): 
    case_number = case_number.upper()
    df['Case Num'] = df['Case Num'].map(string_processing) 
    case_data = df[df['Case Num'] == case_number] 
    case_data = case_data.map(string_processing) 
    return case_data 

def get_download_link(file, filename): 
    b64 = base64.b64encode(file).decode() 
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Download file</a>' 
    return href 

st.title('Excel Case Finder')  

# Step 1: Upload Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])  

# Maintain the loaded DataFrame
if 'df' not in st.session_state:
    st.session_state.df = None

if uploaded_file and st.session_state.df is None:
    st.session_state.df = pd.read_excel(uploaded_file, engine='openpyxl')
    if not st.session_state.df.empty:
        st.write('File uploaded successfully')
    else:
        st.write('Failed to read the file.')

# Proceed only if the DataFrame is loaded
if st.session_state.df is not None:
    # Step 2: Enter Case Number
    case_num = st.text_input('Enter Case Number') 
    find_case_button = st.button('Find Case') 

    # Maintain the loaded case data
    if 'case_data' not in st.session_state:
        st.session_state.case_data = None

    if case_num and find_case_button:
        st.session_state.case_data = find_case_data(st.session_state.df, case_num)
        st.session_state.selected_arguments = {}

    # Proceed only if the case data is found
    if st.session_state.case_data is not None:
        if not st.session_state.case_data.empty:
            issues = st.session_state.case_data['Issue'].unique()
            for issue in issues:
                arguments = get_possible_arguments(issue)
                if arguments:
                    selected_argument = st.selectbox(f"Select argument for issue '{issue}'", arguments, key=issue, 
                                                     index=arguments.index(st.session_state.selected_arguments.get(issue, arguments[0])))
                    st.session_state.selected_arguments[issue] = selected_argument
                else:
                    st.session_state.selected_arguments[issue] = ""

            # Step 3: Create Document
            create_doc = st.button('Create Document') 
            if create_doc:
                docx_file = create_word_document(st.session_state.case_data, [st.session_state.selected_arguments[issue] for issue in issues])
                st.markdown(get_download_link(docx_file, f'Case_{case_num}.docx'), unsafe_allow_html=True)
                st.session_state.case_data = None
                st.session_state.selected_arguments = {}
                st.session_state.df = None
                st.write("Parameters reset. You can search for a new case.")
        else:
            st.write('Case not found in the spreadsheet. Please try again with a different case number.')

# Reset button
reset_button = st.button('Reset All')
if reset_button:
    st.session_state.df = None
    st.session_state.case_data = None
    st.session_state.selected_arguments = {}
    st.write("Parameters reset. You can search for a new case.")
