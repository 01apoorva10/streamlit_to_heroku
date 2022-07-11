# from tempfile import tempdir
import streamlit as st
import os
import pandas as pd
from openpyxl import load_workbook
import xlwings as xw
import pythoncom
import regex as re
import os


def upload_single_file(uploaded_file, dir_path):

    for f in os.listdir(dir_path):
        print("the remaining files are", f)
        os.remove(os.path.join(
            "D:/PROJECTS/ELLIPSONIC/ceone automation/streamlit/xl_to_xl/singleFile", f))

    with open(os.path.join('singleFile', uploaded_file.name), 'wb') as f:

        f.write(uploaded_file.getbuffer())
        print()
        return st.success('saved file : {} in Directory'.format(uploaded_file.name))


def upload_multiple_file(uploaded_file, multiple_dir_path):
    # for f in os.listdir(multiple_dir_path):
    #     print("the remaining files are", f)
    #     os.remove(os.path.join(
    #         "D:/PROJECTS/ELLIPSONIC/ceone automation/streamlit/xl_to_xl/multipleFiles", f))
    with open(os.path.join('multipleFiles', uploaded_file.name), 'wb') as f:
        f.write(uploaded_file.getbuffer())
        print()
        return st.success('saved file : {} in {} Directory'.format(uploaded_file.name, multiple_dir_path))


def transform_uploaded_file(dir_path_for_all):
    pythoncom.CoInitialize()

    excel_file_list = []
    for path in os.listdir(dir_path_for_all):
        # check if current path is a file
        if os.path.isfile(os.path.join(dir_path_for_all, path)):
            excel_file_list.append(path)
    print(excel_file_list)

    for eachExcel_file in excel_file_list:
        workbook = load_workbook(dir_path_for_all+"/"+eachExcel_file,
                                 data_only=False, read_only=False)
        sheet = workbook["SA-6239-ENG"]
        cell_value = sheet["g57"].value
    # print(cell_value)

# ---------------write extracted cell value to text file----------------

        with open("sample.txt", "w+") as outfile:
            outfile.write(cell_value)

# ----------------------read text file to dataframe------------------------
        dataframe = pd.read_table(
            r"sample.txt", encoding="latin1", header=None)
        df = dataframe[(dataframe[0].str.contains(':|-'))]
        df[0] = df[0].str[3:]
        df = df[0].str.split(':|-', 1, expand=True)
        print(df)

        print("\n________\n")

# ------------------load workobook with xlwings module --------------------
        app = xw.App(visible=False)
        workbook = xw.Book(dir_path_for_all+"/"+eachExcel_file)
        sheet = workbook.sheets['SA-6239-ENG']

# first column of dataframe
        sheet.range('G15').options(index=False, header=False).value = df[0]
        sheet.range('Z15').options(
            index=False, header=False).value = df[1]  # second column
        workbook.save(dir_path_for_all+"/"+eachExcel_file)
        workbook.close()


def download_single_file(directory_path):
    excel_file = os.listdir(directory_path)
    for filename in excel_file:
        with open(directory_path+"/"+filename, 'rb') as my_file:
            st.download_button(label='Download', data=my_file, file_name=filename,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def main():
    dir_path = "singleFile"
    multiple_dir_path = "multipleFiles"
    st.title("Excel files transformation")
    menu = ['About', 'single excel file trasformation',
            'multiple excel trasformation']
    choice = st.sidebar.selectbox('Menu', menu)

    if choice == 'single excel file trasformation':
        st.header("Upload single excel files")

        datafile = st.file_uploader(
            "upload xlsx", type=['xlsx'])
        print(datafile)
        if datafile is not None:
            st.write("filename:", datafile.name)

            # function to upload and transform
            upload_single_file(datafile, dir_path)

        transform_uploaded_file(dir_path)
        download_single_file(dir_path)

    elif choice == 'multiple excel trasformation':
        st.header("Upload multiple files")

        datafile = st.file_uploader(
            "upload xlsx", type=['xlsx'], accept_multiple_files=True)
        print("uploaded files are \n____________________________________\n", datafile)
        if datafile is not None:
            for uploaded_file in datafile:
                # bytes_data = uploaded_file.read()
                st.write("filename:", uploaded_file.name)

                upload_multiple_file(uploaded_file, multiple_dir_path)
                transform_uploaded_file(multiple_dir_path)
    else:
        st.subheader('About')


if __name__ == '__main__':
    main()
