from tempfile import tempdir
import streamlit as st
import os
import pandas as pd
from openpyxl import load_workbook
import xlwings as xw
# import pythoncom


def upload_file(uploadedfile):
    with open(os.path.join('tempDir', uploadedfile.name), 'wb') as f:
        f.write(uploadedfile.getbuffer())
        print()
        return st.success('saved file:() in tempDir'.format(uploadedfile.name))


def transform_uploaded_file(input_file):
    # pythoncom.CoInitialize()
    workbook = load_workbook("tempDir"+"/"+input_file,
                             data_only=True, read_only=False)
    sheet = workbook["SA-6239-ENG"]
    cell_value = sheet["g57"].value
    # print(cell_value)

# ---------------write extracted cell value to text file----------------

    with open("sample.txt", "w", encoding='utf8') as outfile:
        file = outfile.write(cell_value)
        # print(file)


# ----------------------read text file to dataframe------------------------
    dataframe = pd.read_table("sample.txt", header=None, encoding='utf8')
    df = dataframe[(dataframe[0].str.contains(':|-'))]
    df[0] = df[0].str[3:]
    df = df[0].str.split(':|-', 1, expand=True)
    # print(df)

    print("\n__________________\n")

# ------------------load workobook with xlwings module --------------------
    # app = xw.App(visible=False)
    workbook = xw.Book("tempDir"+"/"+input_file)
    sheet = workbook.sheets["SA-6239-ENG"]

# first column of dataframe
    sheet.range('G15').options(index=False, header=False).value = df[0]
    sheet.range('Z15').options(
        index=False, header=False).value = df[1]  # second column
    workbook.save()


def main():
    st.title("excel transformation")
    menu = ['Data', 'About']
    choice = st.sidebar.selectbox('Menu', menu)

    if choice == 'Data':
        st.subheader("Upload files")

        datafile = st.file_uploader("upload xlsx", type=['xlsx'])
        if datafile is not None:
            file_details = {'FileName': datafile.name,
                            'FileType': datafile.type}
            print(datafile.name)
            # df=pd.read_excel(datafile)
            # st.dataframe(df)
            upload_file(datafile)
            transform_uploaded_file(datafile.name)

    else:
        st.subheader('About')


if __name__ == '__main__':
    main()
