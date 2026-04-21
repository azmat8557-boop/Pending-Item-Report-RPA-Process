*** Settings ***
Library      pending_report.py

*** Variables ***
${DATA_FILE}    C:/Users/DELL XPS/OneDrive/Desktop/Inventory Mangement project/dashbaord & test data/dashboards/test data/Pending_Items_Report.xlsb
${DASHBOARD_FILE}    C:/Users/DELL XPS/OneDrive/Desktop/Inventory Mangement project/dashbaord & test data/dashboards/Pending Items Report - Formula Sheet.xlsx

*** Tasks ***
Generate Pending Item Report
    Process Pending Report        ${DATA_FILE}    ${DASHBOARD_FILE}
