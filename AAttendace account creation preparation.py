########################################################
## This script is used for creating an Excel file
## which can be imported in Academy Attendance to
## create student user accounts
##
## 1. Ask LFB for Excel files (see Werkdocumentatie)
## 2. Move the Excel files to the same folder as this .py file
## 3. Run Script for each Excel file
## 4. Upload Output file in Academy Attendance (see Werkdocumentatie)
## 

import pandas as pd
from pathlib import Path
from openpyxl import Workbook

Cohort = "2024"  # al aangepast voor 24/25 op 08-08-24 door Laura

# Lees het Excel-bestand in
student_data = pd.read_excel("AAttendace account creation preparation\Bachelorpsy2324.xlsx", skiprows=1) # Eerste rij overslaan want in rij 2 staan de headers pas
student_data = student_data[["Voornaam", "Tussenvoegsel", "Achtrnm", "ID", "E-mail (Pr)"]] # Kies welke kolommen je wil bewaren en in welke volgorde

# Voeg een kolom 'Cohort' toe
student_data['Cohort'] = Cohort

# Hernoem de kolommen zodat ze de naam hebben die AA verwacht
student_data.columns = ['Voornaam', 'Tussenvoegsel', 'Achternaam', 'Studentnummer', 'E-mail', 'Cohort']

# Schrijf het dataframe weg naar een Excel-bestand
name = Path("AAttendace account creation preparation\\aattendance_nieuwe_studenten_cohort_"+Cohort+".xlsx")
with pd.ExcelWriter(name, engine='openpyxl') as writer:
    student_data.to_excel(writer, index=False, sheet_name='Sheet1')