# Streamlit representation code with reset button

st.title('Excel Case Finder')

# Step 1: Upload Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

# Maintain the loaded DataFrame
if 'df' not in st.session_state:
    st.session_state.df = None

# Define the necessary columns
necessary_columns = ['Case Num', 'Case Name', 'Issue', 'Provider ID', 'Provider Name', 'MAC', 'Determination Event Date', 'Appeal Date', 'Audit Adj No.', 'Group FYE', 'FYE', 'Transferred to Case #']

if uploaded_file and st.session_state.df is None:
    # Open the file and get the sheet names and columns
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
    sheet_name = xls.sheet_names[0]
    
    # Check which columns are present in the first sheet
    available_columns = xls.parse(sheet_name, nrows=0).columns.tolist()
    
    # Get the intersection of necessary columns and available columns
    columns_to_load = [col for col in necessary_columns if col in available_columns]
    
    # Load only the necessary columns
    st.session_state.df = xls.parse(sheet_name, usecols=columns_to_load)
    st.write('File uploaded successfully')

# Proceed only if the DataFrame is loaded
if st.session_state.df is not None:
    # Step 2: Enter Case Number
    case_num = st.text_input('Enter Case Number')
    find_case_button = st.button('Find Case')
    reset_button = st.button('Reset')

    # Maintain the loaded case data
    if 'case_data' not in st.session_state:
        st.session_state.case_data = None

    if case_num and find_case_button and st.session_state.case_data is None:
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
                    st.session_sta
