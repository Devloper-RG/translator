import streamlit as st
import translator as t
from translator import translator as class_translator
import pandas as pd
from io import BytesIO

# Initialize session state variables
if 'dataframes' not in st.session_state:
    st.session_state.dataframes = {}
if 'translated_df' not in st.session_state:
    st.session_state.translated_df = None
if 'prompt' not in st.session_state:
    st.session_state.prompt = None
if 'language' not in st.session_state:
    st.session_state.language = None
if 'add_on_prompt' not in st.session_state:
    st.session_state.add_on_prompt = None

st.title("Translator")
file = st.file_uploader("Upload Excel",type='xlsx')

if st.toggle("New prompt"):
    st.session_state.prompt = st.text_area("Prompt",placeholder="Type",height=100)

else:
    st.session_state.language = st.selectbox("Select Language to translate In",options=['Hindi','Bengali','Marathi','Tamil','Malyalam'])
    st.session_state.add_on_prompt = st.multiselect("Select add on for prompt",options=['Easy to understand','3-5 word summary in English after each point.'])

class_translator.language = st.session_state.language
class_translator.add_on_prompt = st.session_state.add_on_prompt

if file:
    
    sheets_name = t.read_sheets(file)
    selected_sheet = st.selectbox("Select the sheets",options=sheets_name)

    if selected_sheet:
        cell_list,row_list,column_list = t.range_used(file,selected_sheet)

selection = st.selectbox("Select content to be translated",options=['cell','column','row','sheet','workbook'],index=None,placeholder="Choose from below",)



with st.form("my_form"):

    if selection =='cell':
        selection = st.multiselect("Select cell name",options= cell_list,help= "First Row excluded as its treated as heading or Column Name")
        type = "cell"

    elif selection =='column':
        selection = st.multiselect("Select the column",options= column_list)
        type = "column"

    elif selection == 'row':
        selection = st.multiselect("Select the row",options= row_list,help= "First Row excluded as its treated as heading or Column Name")
        type = "row"

    elif selection == 'sheet':
        type = "sheet"
        
    elif selection == 'workbook':
        type = "workbook"
        pass
    
    preview = st.form_submit_button("Translate and Preview")


if preview:

    if selection:

        if type == "workbook":
            for sheet in sheets_name:
                df,translated_df= t.process(file,sheet,selection,type,st.session_state.prompt)
                new_variable_name = f"{sheet}"
                st.session_state.dataframes[new_variable_name] = translated_df
                st.write(f"Original Excel sheet-: {sheet}")
                st.dataframe(df)
                st.write(f"Translated Excel sheet-: {sheet}")
                st.dataframe(translated_df)

        else:
            df,st.session_state.translated_df= t.process(file,selected_sheet,selection,type,st.session_state.prompt)

        if type == "sheet":
            st.dataframe(df)
            st.dataframe(st.session_state.translated_df)

        elif type!= "workbook":
            st.dataframe(df)

    else:
        st.error("Please Select Part of Excel to be Translated")


c1,c2 =st.columns([1,3])
with c1:
    replace = st.button("Replace Excel")

if replace:
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    if type == "workbook":
        
        for key, value in st.session_state.dataframes.items():
            # Use the key as the sheet name
            value.to_excel(writer, sheet_name=key, index=False)
        writer.close()

        with c2:
            st.download_button("📥 Download Translated Workbook",
                        data=output.getvalue(),
                        file_name=file.name)
    else:
        
        st.session_state.translated_df.to_excel(writer, sheet_name=selected_sheet, index=False)
        writer.close()

        with c2:
            st.download_button("📥 Download Translated Sheet",
                        data=output.getvalue(),
                        file_name=file.name)