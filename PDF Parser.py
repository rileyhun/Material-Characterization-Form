import re
import pandas as pd
from tkinter import Tk
from tkinter import filedialog
from subprocess import PIPE, Popen

root = Tk()
files = filedialog.askopenfilenames(parent=root, title="Select the MSDS's that you want to pull data from")
filez = list(files)

for file in filez:
    extracted_text = Popen([r"C:\Users\RHun\xpdfbin-win-3.04\bin64\pdftotext.exe", file, "-"], stdout=PIPE).communicate()[0]
    extracted_text = extracted_text.decode('ISO-8859-1')

    numbers = re.findall(r'\d+-\d+-\d+', extracted_text)

    df = pd.read_excel('SafeTec ALL SDS.xlsx', encoding='ISO-8859-1')

    listofCAS = list(set(df['CAS Number'].tolist()))
    ChemicalTable = df[['CAS Number', 'Chemical Component']].drop_duplicates('CAS Number').sort_values('CAS Number')

    myCAS = list(set([ele for ele in numbers if ele in listofCAS]))



    ChemicalTable.set_index('CAS Number', inplace=True)
    ChemicalTable['Chemical Component'] = ChemicalTable['Chemical Component'].str.replace('\xa0', ' ')

    Components = []
    for i in myCAS:
        Components.append(ChemicalTable.loc[i]['Chemical Component'])


    print(list(zip(myCAS, Components)))