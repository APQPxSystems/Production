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
                   page_icon="📁",
                   layout="wide")

# App Title and Description
st.markdown("<h1 style='text-align: center;'>PRODUCTION COUNTER APP</h1>", unsafe_allow_html=True)
st.write("------------------------------------------")

# Fill Up Form
st.markdown("<h5 style='text-align: center; background-color:DodgerBlue;'>Line and Output Details</h5>", unsafe_allow_html=True)

# Line / Shift / Date Column
st.markdown("<h3 style='text-align: center;'>LINE / SHIFT / DATE</h3>", unsafe_allow_html=True)
lsdcol1, lsdcol2, lsdcol3 = st.columns(3)
with lsdcol1:
    line_no = st.text_input("Line no.:")
with lsdcol2:
    shift = st.selectbox("Select Shift:", ["DS", "NS"])
with lsdcol3:
    date = st.date_input("Select Date:", format="MM/DD/YYYY")
st.write("------------------------------------------")

# Target Yield and PPM
st.markdown("<h3 style='text-align: center;'>TARGET YIELD AND PPM</h3>", unsafe_allow_html=True)
yieldcol, ppmcol = st.columns(2)
with yieldcol:
    yield_target = st.text_input("Target Yield:")
with ppmcol:
    ppm_target= st.text_input("Target PPM")
st.write("------------------------------------------")

# Plan --- Target / Actual Column
st.markdown("<h3 style='text-align: center;'>PLAN</h3>", unsafe_allow_html=True)
plancol1, plancol2 = st.columns(2)
with plancol1:
    plan_target = st.text_input("Plan Target:")
with plancol2:
    plan_actual = st.text_input("Plan Actual:")
st.write("------------------------------------------")

# Acctg Efficiency --- Target / Actual Column
st.markdown("<h3 style='text-align: center;'>ACCTG EFFICIENCY</h3>", unsafe_allow_html=True)
acctgcol1, acctgcol2 = st.columns(2)
with acctgcol1:
    acctg_target = st.text_input("Acctg Target:")
with acctgcol2:
    acctg_actual = st.text_input("Acctg Actual:")
st.write("------------------------------------------")

# Hourly Output --- Target / Actual Column
st.markdown("<h3 style='text-align: center;'>HOURLY OUTPUT</h3>", unsafe_allow_html=True)
hourlycol1, hourlycol2 = st.columns(2)
with hourlycol1:
    hourly_target = st.text_input("Hourly Target:")
with hourlycol2:
    hourly_actual = st.text_input("Hourly Actual:")
st.write("------------------------------------------")

# -------------------------------------

st.markdown("<h5 style='text-align: center; background-color:DodgerBlue;'>Manpower Details</h5>", unsafe_allow_html=True)

# PD Manpower
st.markdown("<h3 style='text-align: center;'>PD MANPOWER</h3>", unsafe_allow_html=True)
pdmpcol1, pdmpcol2 = st.columns(2)
with pdmpcol1:
    pdmp_plan = st.text_input("PD MP Plan")
with pdmpcol2:
    pdmp_actual = st.text_input("PD MP Actual")
st.write("------------------------------------------")

# QA Manpower
st.markdown("<h3 style='text-align: center;'>QA MANPOWER</h3>", unsafe_allow_html=True)
qampcol1, qampcol2 = st.columns(2)
with qampcol1:
    qamp_plan = st.text_input("QA MP Plan")
with qampcol2:
    qamp_actual = st.text_input("QA MP Actual")
st.write("------------------------------------------")

# -------------------------------------

st.markdown("<h5 style='text-align: center; background-color:DodgerBlue;'>Inspection Output</h5>", unsafe_allow_html=True)

# Dimension Inspection
st.markdown("<h3 style='text-align: center;'>DIMENSION INSPECTION</h3>", unsafe_allow_html=True)
dimensioncol1, dimensioncol2 = st.columns(2)
with dimensioncol1:
    dimension_good = st.text_input("Count of Dimension Good:")
with dimensioncol2:
    dimension_ng = st.text_input("Count of Dimension NG:")
st.write("------------------------------------------")

# ECT Inspection
st.markdown("<h3 style='text-align: center;'>ECT INSPECTION</h3>", unsafe_allow_html=True)
ectcol1, ectcol2 = st.columns(2)
with ectcol1:
    ect_good = st.text_input("Count of ECT Good:")
with ectcol2:
    ect_ng = st.text_input("Count of ECT NG:")
st.write("------------------------------------------")

# Clamp Checking Inspection
st.markdown("<h3 style='text-align: center;'>CLAMP CHECKING</h3>", unsafe_allow_html=True)
clampcol1, clampcol2 = st.columns(2)
with clampcol1:
    clamp_good = st.text_input("Count of Clamp Checking Good:")
with clampcol2:
    clamp_ng = st.text_input("Count of Clamp Checking NG:")
st.write("------------------------------------------")

# Appearance Inspection
st.markdown("<h3 style='text-align: center;'>APPEARANCE INSPECTION</h3>", unsafe_allow_html=True)
appearancecol1, appearancecol2 = st.columns(2)
with appearancecol1:
    appearance_good = st.text_input("Count of Appearance Good:")
with appearancecol2:
    appearance_ng = st.text_input("Count of Appearance NG:")
st.write("------------------------------------------")

# QA Inspection
st.markdown("<h3 style='text-align: center;'>QA INSPECTION</h3>", unsafe_allow_html=True)
qainspectioncol1, qainspectioncol2 = st.columns(2)
with qainspectioncol1:
    qainspection_good = st.text_input("Count of QA Inspection Good:")
with qainspectioncol2:
    qainspection_ng = st.text_input("Count of QA Inspection NG:")
st.write("------------------------------------------")

# -------------------------------------














# -------------------------------------
# Write Excel File
st.write("------------------------------------------")
def generate_excel_file():
    # Create a BytesIO buffer to write the Excel file to
    excel_buffer = BytesIO()

    # Create a workbook and add a worksheet
    workbook = xlsxwriter.Workbook(excel_buffer, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    # Formats
    dateformat1 = workbook.add_format({"num_format": "mm/dd/yyyy"})

    # Write data to the worksheet
    worksheet.write(0,0, "Date:")
    worksheet.write(0,1, date, dateformat1)
    worksheet.write(1,0, "Line:")
    worksheet.write(1,1, line_no)
    worksheet.write(2,0, "Shift:")
    worksheet.write(2,1, shift)

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
