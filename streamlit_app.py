
import streamlit as st
import pandas as pd
from openpyxl import load_workbook

st.set_page_config(page_title="Excel Translator", layout="wide")
st.title("📘 מתרגם Excel קובצי בתחום השילוח הבינלאומי")

# העלאת קובץ
uploaded_file = st.file_uploader("להעלות קובץ Excel לתרגום", type=["xlsx"])

# טען את המילון
glossary_df = pd.read_excel("glossary.xlsx")
dictionary = dict(zip(glossary_df["English"], glossary_df["Hebrew"]))

# הצגת טופס להוספת מונחים
st.sidebar.header("הוספת מונחים למילון")
new_eng = st.sidebar.text_input("מונח באנגלית")
new_heb = st.sidebar.text_input("תרגום לעברית")
if st.sidebar.button("➕ הוסף למילון"):
    if new_eng and new_heb:
        glossary_df = glossary_df._append({"English": new_eng, "Hebrew": new_heb}, ignore_index=True)
        glossary_df.to_excel("glossary.xlsx", index=False)
        st.sidebar.success("המונח נוסף בהצלחה!")
    else:
        st.sidebar.warning("יש למלא את שני השדות")

# הצגת טבלת מילון
st.sidebar.subheader("מילון קיים")
st.sidebar.dataframe(glossary_df)

# תרגום קובץ
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    translated_df = df.replace(dictionary)
    st.success("🔄 הקובץ תורגם בהצלחה!")
    st.dataframe(translated_df)

    # הורדת הקובץ
    from io import BytesIO
    output = BytesIO()
    translated_df.to_excel(output, index=False)
    st.download_button(label="📥 הורד קובץ מתורגם", data=output.getvalue(), file_name="translated.xlsx")
