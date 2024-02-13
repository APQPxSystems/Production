# Production Counter Web App
# Kent Jym Katigbak -- Staff
# Systems Engineering

# -------------------------------------

# Import Libraries
import streamlit as st
import pandas as pd
from datetime import datetime, date
import xlsxwriter
from io import BytesIO




# -------------------------------------

# App Configurations
st.set_page_config(page_title="Apps by Systems Engineering",
                   page_icon="🤠",
                   layout="wide")

# Sidebar Inputs

with st.sidebar:
    date = st.date_input("Select Date:", format="MM/DD/YYYY")
    line_no = st.text_input("Line no.:")
    shift = st.selectbox("Select Shift:", ["DS", "NS"])
    plan = st.text_input("Plan:")
    actual = st.text_input("Actual Produced:")
# -------------------------------------

# App Title and Description
st.markdown("<h1 style='text-align: center;'>PRODUCTION COUNTER APP</h1>", unsafe_allow_html=True)
st.write("------------------------------------------")


# Header --- Line, Shift, and Date Info
headercol1, headercol2, headercol3 = st.columns(3)

with headercol1:
    st.subheader(f"Line {line_no}")
with headercol2:
    st.subheader(shift)
with headercol3:
    st.subheader(date)

st.write("------------------------------------------")

















# -------------------------------------
# Write Excel File

def generate_excel_file():
    # Create a BytesIO buffer to write the Excel file to
    excel_buffer = BytesIO()

    # Create a workbook and add a worksheet
    workbook = xlsxwriter.Workbook(excel_buffer, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    # Write data to the worksheet
    worksheet.write(0,0, "Date:")
    worksheet.write(0,1, "Line:")
    worksheet.write(0,2, "Shift:")

    # Close the workbook
    workbook.close()

    return excel_buffer

def main():
    st.title("Download Excel File")

    # Generate the Excel file
    excel_buffer = generate_excel_file()

    # Create a download button
    st.download_button(
        label="Download Excel",
        data=excel_buffer.getvalue(),
        file_name="PD Counter.xlsx",
        key="download_button"
    )

if __name__ == '__main__':
    main()

# -------------------------------------
