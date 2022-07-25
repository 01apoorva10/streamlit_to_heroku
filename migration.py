# from tempfile import tempdir
import streamlit as st
import os
import pandas as pd
from openpyxl import load_workbook
import xlwings as xw
import pythoncom
import regex as re
import os
from streamlit_option_menu import option_menu
from PIL import Image
import zipfile


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
    
    os.remove(os.path.join(
            "D:/PROJECTS/ELLIPSONIC/ceone automation/streamlit/xl_to_xl/multipleFiles/", f))
    with open(os.path.join('multipleFiles', uploaded_file.name), 'wb') as f:
        f.write(uploaded_file.getbuffer())
        print()
        return st.success('saved file : {} in Directory'.format(uploaded_file.name))


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

# ------------------ for zipping multiple files ----------------------


def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))

zipf = zipfile.ZipFile('Zipped_file.zip', 'w', zipfile.ZIP_DEFLATED)
zipdir('./multipleFiles', zipf)
zipf.close()


def main():

    with st.sidebar.container():
        image = Image.open("D:/PROJECTS/STREAMLIT/CEONE/xl_to_xl/logo22.png")
        st.image(image, use_column_width=True)


    dir_path = "singleFile"
    multiple_dir_path = "multipleFiles"
    #
    with st.sidebar:

        app_mode = option_menu("NAVIGATION", ["DATA MIGRATION", "MTO", "TEMPLATE AUTOMATION"],
                               icons=['file-earmark-binary', 'file-earmark-spreadsheet',
                                      'file-spreadsheet-fill'],
                               menu_icon="list", default_index=0,
                               styles={
            "container": {"padding": "5!important", "background-color": "#f0ff6"},
            "icon": {"color": "white", "font-size": "28px"},
            "nav-link": {"font-size": "16px", "text-align": "left", "margin": "0px", "--hover-color": "#eeee"},
            "nav-link-selected": {"background-color": "#2C845"},
        }
        )
    if app_mode == "DATA MIGRATION":
        st.title("Data Migration")
        menu = ['About', 'single excel file trasformation',
                'multiple excel trasformation']
        choice = st.selectbox('Menu', menu)
        st.write("")
        st.write("")

        if choice == 'single excel file trasformation':
            st.header("Upload single excel files")

            datafile = st.file_uploader(
                "upload .xlsx files only", type=['xlsx'])
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
                "upload .xlsx files only ", type=['xlsx'], accept_multiple_files=True)
            print(datafile)
            if datafile is not None:
                for uploaded_file in datafile:
                    # bytes_data = uploaded_file.read()
                    st.write("filename:", uploaded_file.name)

                    upload_multiple_file(uploaded_file, multiple_dir_path)
                transform_uploaded_file(multiple_dir_path)

        else:
            st.subheader('About')
            st.text('Yet to be given by ceone')
    elif app_mode == "MTO":
        st.title("MTO operation")
        menu = ['About', 'single excel file trasformation',
                'multiple excel trasformation']
        choice = st.selectbox('Menu', menu)
        st.write("")
        st.write("")

        if choice == 'single excel file trasformation':
            st.header("Upload single excel files")

            datafile = st.file_uploader(
                "upload .xlsx files only", type=['xlsx'])
            print(datafile)


if __name__ == '__main__':
    main()
