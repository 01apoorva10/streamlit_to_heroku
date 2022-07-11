from tempfile import tempdir
import streamlit as st
import os
import pandas as pd
from openpyxl import load_workbook
import xlwings as xw
# import pythoncom


def upload_file(uploaded_file):
    with open(os.path.join('tempDir', uploaded_file.name), 'wb') as f:
        f.write(uploaded_file.getbuffer())
        print()
        return st.success('saved file:() in tempDir'.format(uploaded_file.name))


def main():
    st.title("XL to XL")
    menu = ['Data', 'About']
    choice = st.sidebar.selectbox('Menu', menu)

    if choice == 'Data':
        st.subheader("Upload files")

        datafile = st.file_uploader(
            "upload xlsx", type=['xlsx'], accept_multiple_files=True)
        print(datafile)
        if datafile is not None:
            for uploaded_file in datafile:
                # bytes_data = uploaded_file.read()
                st.write("filename:", uploaded_file.name)

            # print(datafile.name)
                upload_file(uploaded_file)


    else:
        st.subheader('About')


if __name__ == '__main__':
    main()
