import streamlit as st
import os
from determinationparsing import (
    extract_text_from_pdf,
    check_determination_text,
    check_iqrp_requirement,
    extract_determination_date,
    extract_provider_name,
    extract_provider_number,
    extract_quality_measure,
    extract_issue_input
)

st.set_page_config(
    page_title="Determination Parser",
    page_icon="📄",
    layout="wide"
)

st.title("Determination Parser")
st.write("Upload a PDF file to parse determination information")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    # Save the uploaded file temporarily
    temp_path = "temp_upload.pdf"
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    
    try:
        # Extract text from PDF
        text = extract_text_from_pdf(temp_path)
        
        # Check if it's a valid determination
        is_determination = check_determination_text(text)
        is_iqrp = check_iqrp_requirement(text)
        
        if not is_determination:
            st.error("This is not a valid final determination")
        else:
            # Create two columns for the layout
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Document Information")
                st.write(f"**Determination Date:** {extract_determination_date(text)}")
                st.write(f"**Provider Name:** {extract_provider_name(text)}")
                st.write(f"**Provider Number:** {extract_provider_number(text)}")
                st.write(f"**Quality Measure:** {extract_quality_measure(text)}")
            
            with col2:
                st.subheader("Issue Information")
                provider_name = extract_provider_name(text)
                issue_input = extract_issue_input(text, provider_name)
                
                if is_iqrp:
                    # Get the directory where the current script is located
                    current_dir = os.path.dirname(os.path.abspath(__file__))
                    facts_path = os.path.join(current_dir, "stratifiers", "IQRP", "Facts.txt")
                    issue_path = os.path.join(current_dir, "stratifiers", "IQRP", "Issue.txt")
                    
                    try:
                        with open(facts_path, 'r') as f:
                            facts = f.read().strip()
                        with open(issue_path, 'r') as f:
                            issue = f.read().strip()
                            
                        # Replace <INPUT> with issue_input in the issue text
                        issue = issue.replace("<INPUT>", issue_input)
                        
                        st.write("**Facts:**")
                        st.text_area("", facts, height=200)
                        st.write("**Issue:**")
                        st.text_area("", issue, height=200)
                    except FileNotFoundError as e:
                        st.error(f"Error: Could not find required files: {e}")
                else:
                    st.warning("This document does not contain the IQRP requirement text")
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
    
    finally:
        # Clean up the temporary file
        if os.path.exists(temp_path):
            os.remove(temp_path) 