'''
    Script for preparing ELinks data for the COBISS-SR database
    Author:
    Mitar Zečević
    National Library of Serbia
    mitar.zecevic@nb.rs / mitarzecevic648@gmail.com
'''

import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# Returns a list of url's from the backup reference file
def listaBStrana():
    with open('bStrane.txt', 'r', encoding='utf-8') as bStrane:
        content = bStrane.read().split('\n')
        urlLista = []
        for item in content:
            urlLista.append(item.split(',')[0])
        return urlLista

# Writes an .xlsx file for the backup Elinks (B strane and zvučne knjige, they remain the same)
def ispisiBStrane(datum):
    with open('bStrane.txt', 'r', encoding='utf-8') as bStrane:
        df = pd.read_csv(bStrane, header=None, names=['URL', 'COBISS.SR-ID'])
        df.insert(0, 'Redni broj', range(1, len(df) + 1))
        df["Napomena"] = ""
        df = df[["Redni broj", "COBISS.SR-ID", "URL", "Napomena"]]
        
        df.to_excel(f'Elinks_{datum}_b.xlsx', index=False, engine='openpyxl')

# Formats the output excel files
def adjustColumnWidths(fileName, sheetName):
    wb = load_workbook(fileName)
    ws = wb.active

    # Iterate through columns to adjust their widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Get column letter
        for cell in column:
            if cell.value:  # Avoid NoneType errors for empty cells
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = max_length + 2  # Add some padding
        ws.column_dimensions[column_letter].width = adjusted_width

    # Change the name of the worksheet
    ws.title = sheetName
    # Save the workbook with adjusted column widths
    wb.save(fileName)

# ========================================================================

izvestaj = 'procitajIzvestaj.txt'
datum = datetime.now().strftime("%d_%m_%Y")

listaB = listaBStrana()

with open(izvestaj, 'r', encoding='utf-8') as inFile:
    # Clean the input file
    content = inFile.read()
    content = content.strip()
    content = content.split('\n')

    # Prepare the filtered dictionary, removing duplicates or short or empty ID rows
    recnik = {}
    for item in content:
        if item.endswith('|'):
            pass
        else:
            url, cobissID = item.split('|')
            # Remove the rows present in the backup file and short IDs
            if len(cobissID) > 4 and url not in listaB:
                recnik[url] = cobissID
    
    # Prepare the dataframe for writing with all the needed columns
    df = pd.DataFrame(list(recnik.items()), columns=['URL', 'COBISS.SR-ID'])
    df['COBISS.SR-ID'] = pd.to_numeric(df['COBISS.SR-ID'])
    df.insert(0, 'Redni broj', range(1, len(df) + 1))
    df["Napomena"] = ""
    df = df[["Redni broj", "COBISS.SR-ID", "URL", "Napomena"]]
    df.to_excel(f'Elinks_{datum}.xlsx', index=False)

# Write the backup file and format outputs
ispisiBStrane(datum)
adjustColumnWidths(f'Elinks_{datum}.xlsx', 'Elinks glavno')
adjustColumnWidths(f'Elinks_{datum}_b.xlsx', 'B strane i zv knjige')
