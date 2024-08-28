import pandas as pd



# Pfade zu den Excel-Dateien
datei1 = r'C:\Users\Pharmosan\Documents\Projekte\Hera\FehlenderNanobrick.xlsx'
datei2 = r'C:\Users\Pharmosan\Documents\Projekte\Hera\Nano_Liste.xls'

# Daten aus der ersten Datei laden
df1 = pd.read_excel(datei1, sheet_name='Tabelle1', engine = 'openpyxl')

# Daten aus der zweiten Datei laden
df2 = pd.read_excel(datei2, sheet_name='zuNano_01', engine = 'xlrd')
df2_sheet2 = pd.read_excel(datei2, sheet_name='zuNano_02', engine = 'xlrd')
df2_sheet3 = pd.read_excel(datei2, sheet_name='zuNano_03', engine = 'xlrd')
df2_sheet4 = pd.read_excel(datei2, sheet_name='zuNano_04', engine = 'xlrd')

def suche_und_einfuegen(postleitzahl, df_sheets, index):
    for df in df_sheets:
        match_row = df[df.iloc[:, 1] == postleitzahl]
        if not match_row.empty:
            # Wert in der zehnten Spalte (Spalte J, index 9) in derselben Zeile
            wert_aus_zweiter_datei = match_row.iloc[0, 9]
            # Diesen Wert in die sechste Spalte (Spalte F, index 5) der ersten Datei einfügen
            df1.at[index, df1.columns[5]] = wert_aus_zweiter_datei
            return True  # Treffer gefunden, daher Rückgabe True
    return False  # Kein Treffer gefunden

# Durch die Zeilen der ersten Datei iterieren, ab der zweiten Zeile
for index, row in df1.iterrows():
    if index == 0:  # Überspringe die erste Zeile (index 0)
        continue

    # Postleitzahl in der zweiten Spalte (Spalte B, index 1) der ersten Datei
    postleitzahl = row[2]  # row[1] bezieht sich auf Spalte B

    # Suche in den vier Sheets
    sheets = [df2, df2_sheet2, df2_sheet3, df2_sheet4]
    treffer = suche_und_einfuegen(postleitzahl, sheets, index)
# Die aktualisierte erste Datei speichern
df1.to_excel('aktualisierte_datei1.xlsx', index=False)

print("Der Abgleich wurde abgeschlossen und die aktualisierte Datei wurde als 'aktualisierte_datei1.xlsx' gespeichert.")
