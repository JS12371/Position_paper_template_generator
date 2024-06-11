def copy_paragraphs(src, dest):
    in_footnote = False
    footnote_text = []

    for paragraph in src.paragraphs:
        if "FOOTNOTE:" in paragraph.text:
            in_footnote = True
            footnote_text.append(paragraph.text.replace("FOOTNOTE:", "").strip())
        elif "END FOOTNOTE" in paragraph.text:
            in_footnote = False
            # Add the footnote to the document
            footnote = dest.add_paragraph()
            footnote_run = footnote.add_run(" ".join(footnote_text))
            footnote_run.font.size = Pt(11)
            footnote_run.font.name = 'Times New Roman'
            footnote_run.italic = True  # Footnote text is often italicized
            footnote_text = []
        elif in_footnote:
            footnote_text.append(paragraph.text.strip())
        else:
            dest_paragraph = dest.add_paragraph()
            copy_paragraph_format(paragraph, dest_paragraph)
            copy_runs(paragraph, dest_paragraph)

def create_word_document(case_data, selected_arguments):  
    doc = Document()  
    header = doc.add_paragraph('BEFORE THE PROVIDER REIMBURSEMENT REVIEW BOARD') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 
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
    run.font.name = 'Times New Roman' 

    cell_middle = table.cell(0,1) 
    cell_middle.text = ")\n"*16
    cell_middle.vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
    cell_middle.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
    run = cell_middle.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 

    cell_right = table.cell(0,2) 
    cell_right.text = f"\n\n\n\n\n\nPRRB Case No. {case_num}\n\nFYE: {year[:10]}\n" 
    run = cell_right.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 

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
    run.font.name = 'Times New Roman' 

    sub = doc.add_paragraph(f"Submitted by:\n\n<Name>\n{mac_name}\n<Address>\n<Address line 2>") 
    sub.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = sub.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 

    sub = doc.add_paragraph(f"\nand\n\n<Reviewer Name>\nFederal Specialized Services, LLC\n1701 S. Racine Avenue\nChicago, IL 60608-4058") 
    sub.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = sub.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 

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
                run.font.name = 'Times New Roman' 
                run.font.bold = True 

    cell_left1.text = f"TABLE OF CONTENTS" 
    run = cell_left1.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 
    run.font.bold = True 

    cell_right1 = table1.cell(0,1) 
    cell_right1.text = f"PAGE" 
    run = cell_right1.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 
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
    run.font.name = 'Times New Roman' 

    cell_right = table.cell(0,1) 
    cell_right.text = "\n1\n\n2\n\n3\n\n?\n\n?" 
    run = cell_right.paragraphs[0].runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 

    doc.add_page_break() 

    header = doc.add_paragraph('I. INTRODUCTION') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 
    run.font.bold = True 
    run.font.color.rgb = RGBColor(0,0,0) 

    header = doc.add_paragraph() 
    run = header.add_run() 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 
    run.text = f"\n\n Case Name: {case_name}\n\nProvider Numbers: {provider_numbers}\n\nLead Contractor: {mac_name}\n\nCalendar Year: {year[-4:]}\n\nPRRB Case Number: {case_num}\n\nDates of Determinations: {determination_event_dates}\n\nDate of Appeal: {date_of_appeal}" 

    doc.add_page_break() 

    header = doc.add_paragraph('II. ISSUES AND ADJUSTMENTS IN DISPUTE') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 
    run.font.bold = True 
    run.font.color.rgb = RGBColor(0,0,0) 

    header = doc.add_paragraph() 
    run = header.add_run() 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 

    for i in range(len(issue)):
        if group_mode:
            if len(adj_no) > 1:
                adj_no = "Various"
            header = doc.add_paragraph(f"\nIssue: {issue[i]}\n\nAdjustment No(s): {adj_no}\n\nApproximate Reimbursement Amount: N/A\n")
            run = header.runs[0] 
            run.font.size = Pt(11) 
            run.font.name = 'Times New Roman' 
        else:
            if len(adj_no) > 1:
                adj_no = "Various"
            header = doc.add_paragraph(f"\nIssue {i + 1}: {cloneissue[i]}\n\nDisposition: {issue[i]}\n\nAdjustment No(s): {adj_no}\n\nApproximate Reimbursement Amount: N/A\n")
            run = header.runs[0] 
            run.font.size = Pt(11) 
            run.font.name = 'Times New Roman'

    doc.add_page_break() 

    header = doc.add_paragraph('III. MAC\'S POSITION') 
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT 
    run = header.runs[0] 
    run.font.size = Pt(11) 
    run.font.name = 'Times New Roman' 
    run.font.bold = True 
    run.font.color.rgb = RGBColor(0,0,0) 

    exhibits_list = []
    i = 0 
    issue = list(filter(None, issue))
    if issue[0] == 'Issue not found':
        pass
    else: 
        if case_num[-1] == 'G' or case_num[-1] == 'C':
            header = doc.add_paragraph("Issue: ")
            run = header.add_run()
            run.font.size = Pt(11)
            run.font.name = 'Times New Roman'
            if issue[0].startswith('Transfer'):
                issue_content = issue[1]
            else:
                issue_content = get_issue_content(issue[0], doc, selected_arguments[0])
            header = doc.add_paragraph(f"{issue_content}")
            run = header.add_run()
            run.font.size = Pt(11)
            run.font.name = 'Times New Roman'
        else:
            while i < len(issue):
                header = doc.add_paragraph(f"\n\nIssue {i+1}: ")
                run = header.add_run() 
                run.font.size = Pt(11) 
                run.font.name = 'Times New Roman' 
                issue_content = get_issue_content_with_exhibits(issue[i], doc, selected_arguments[i], exhibits_list)
                header = doc.add_paragraph(f"{issue_content} \n\n") 
                run = header.add_run() 
                run.font.size = Pt(11) 
                run.font.name = 'Times New Roman' 
                i += 1 

    # Add exhibits at the end of the document for individual cases
    doc.add_page_break()
    if case_num[-1] == 'G' or case_num[-1] == 'C':
        pass
    else:
        header = doc.add_paragraph('V. EXHIBITS')
        header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        run = header.runs[0]
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)

        for exhibit in exhibits_list:
            exhibit_para = doc.add_paragraph(exhibit)
            exhibit_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            run = exhibit_para.runs[0]
            run.font.size = Pt(11)
            run.font.name = 'Times New Roman'

    for paragraph in doc.paragraphs:
        if 'None' in paragraph.text:
            paragraph.text = paragraph.text.replace('None', '', 1)

    set_font_properties(doc)

    buffer = BytesIO()  
    doc.save(buffer)  
    return buffer.getvalue()
