import fitz  # PyMuPDF
import re
from datetime import datetime
import os

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""

    for page in doc:
        full_text += page.get_text() + "\n"

    return full_text

def extract_determination_date(text):
    # Pattern to match "Month Day, Year" format
    pattern = r'(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}'
    match = re.search(pattern, text)
    if match:
        return match.group(0)
    return None

def check_determination_text(text):
    pattern = r"NOTICE\s+OF\s+QUALITY\s+REPORTING\s+PROGRAM\s+NONCOMPLIANCE\s+DECISION\s+UPHELD"
    return bool(re.search(pattern, text, re.MULTILINE))

def check_iqrp_requirement(text):
    pattern = r"The\s+requirement\s+that\s+facilities\s+participating\s+in\s+the\s+Hospital\s+!QR\s+Program\s+report\s+quality\s+data\s+to\s+CMS\s+is\s+set\s+forth\s+in\s+42\s+Code\s+of\s+Federal\s+Regulations\s+\(CFR\)\s+Part\s+412,\s+Subpart\s+H\."
    pattern2 = r"The\s+requirement\s+that\s+facilities\s+participating\s+in\s+the\s+Hospital\s+IQR\s+Program\s+report\s+quality\s+data\s+to\s+CMS\s+is\s+set\s+forth\s+in\s+42\s+Code\s+of\s+Federal\s+Regulations\s+\(CFR\)\s+Part\s+412,\s+Subpart\s+H\."
    if bool(re.search(pattern, text, re.MULTILINE)):
        return bool(re.search(pattern, text, re.MULTILINE))
    elif bool(re.search(pattern, text, re.MULTILINE)):
        return bool(re.search(pattern2, text, re.MULTILINE))

def extract_provider_name(text):
    # Find the determination text first
    determination_pattern = r"NOTICE\s+OF\s+QUALITY\s+REPORTING\s+PROGRAM\s+NONCOMPLIANCE\s+DECISION\s+UPHELD"
    determination_match = re.search(determination_pattern, text, re.MULTILINE)
    
    if determination_match:
        # Get the text after the determination
        text_after_determination = text[determination_match.end():]
        
        # Look for "for" followed by text until "CMS"
        provider_pattern = r"for\s+(.*?)\s+CMS"
        provider_match = re.search(provider_pattern, text_after_determination, re.IGNORECASE | re.DOTALL)
        
        if provider_match:
            return provider_match.group(1).strip()
    return None

def extract_provider_number(text):
    # Look for "CMS Certification Number" followed by 6 digits
    pattern = r"CMS\s+Certification\s+Number\s+(\d{6})"
    match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
    
    if match:
        # Get the 6 digits and format with a dash
        num = match.group(1)
        return f"{num[:2]}-{num[2:]}"
    return None

def extract_quality_measure(text):
    # Look for the bullet point and capture all text after it until the next bullet point or section break
    pattern = r"•\s*(.*?)(?=\n\s*•|\n\s*[A-Z]|\Z)"
    match = re.search(pattern, text, re.MULTILINE | re.DOTALL)
    
    if match:
        # Clean up the captured text by removing extra whitespace and newlines
        measure_text = match.group(1).strip()
        # Replace multiple newlines with a single space
        measure_text = re.sub(r'\n\s*', ' ', measure_text)
        return measure_text
    return None

def extract_issue_input(text, provider_name):
    # Pattern to match text between "Fiscal Year" and "for"
    pattern = r"Fiscal Year\s+(.*?)\s+for"
    
    # Debug: Print the text we're searching in
    print("\nDEBUG - Text being searched:")
    # Find the position of "Fiscal Year"
    fiscal_year_pos = text.find("Fiscal Year")
    if fiscal_year_pos != -1:
        # Print 100 characters before and after "Fiscal Year"
        start = max(0, fiscal_year_pos - 100)
        end = min(len(text), fiscal_year_pos + 100)
        print("Text around 'Fiscal Year':")
        print(text[start:end])
    
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    
    if match:
        # Debug: Print what was matched
        print("\nDEBUG - Full match:", match.group(0))
        print("DEBUG - Captured group:", match.group(1))
        
        # Clean up the captured text
        issue_input = match.group(1).strip()
        # Replace multiple newlines with a single space
        issue_input = re.sub(r'\n\s*', ' ', issue_input)
        return issue_input
    return None

# Example usage
pdf_file = "/Users/jonahsmith/Desktop/SUMMER2025FSS/QRFinalDet/IQRPModelFinalDetermination.pdf"
text = extract_text_from_pdf(pdf_file)
isDetermination = check_determination_text(text)
isIQRP = check_iqrp_requirement(text)

# Debug printing
print("First 500 characters of text:")
print(text[:500])
print("\nLooking for pattern in text...")

if not isDetermination:
    print("This is not a valid final determination")
    exit()

if isIQRP:
    print("This document contains the IQRP requirement text")
    # Get the directory where the current script is located
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # Construct paths relative to the script location
    facts_path = os.path.join(current_dir, "stratifiers", "IQRP", "Facts.txt")
    issue_path = os.path.join(current_dir, "stratifiers", "IQRP", "Issue.txt")
    
    try:
        with open(facts_path, 'r') as f:
            facts = f.read().strip()
        with open(issue_path, 'r') as f:
            issue = f.read().strip()
    except FileNotFoundError as e:
        print(f"Error: Could not find required files: {e}")
        exit()
else:
    print("This document does not contain the IQRP requirement text")
    facts = None
    issue = None

determination_date = extract_determination_date(text)
provider_name = extract_provider_name(text)
prov_num = extract_provider_number(text)
quality_measure = extract_quality_measure(text)
issue_input = extract_issue_input(text, provider_name)

# Replace <INPUT> with issue_input in the issue text
if issue:
    issue = issue.replace("<INPUT>", issue_input)

print(f"Determination Date: {determination_date}")
print(f"Provider Name: {provider_name}")
print(f"Provider Number: {prov_num}")
print(f"Quality Measure: {quality_measure}")
print(f"Issue Input: {issue_input}")
if isIQRP:
    print(f"Facts: {facts}")
    print(f"Issue: {issue}") 