import streamlit as st
import pandas as pd

st.title("Excel Update App")

df=pd.read_excel('files/R917459001X01.xlsx')
st.write(df)

# st.sidebar.header('Options')
# options_form=st.sidebar.form('options_form')
# user_name=options_form.text_input('user_name')
# user_age=options_form.text_input('user_age')
# add_data=options_form.form_submit_button()
# if add_data:
#     st.write(user_name,)