import pandas as pd
from os.path import splitext

d = {"G": "3.9", "A": "5.5", "V": "7.0",
     "L": "8.5", "I": "8.5", "F": "9.7",
     "Y": "10.4", "W": "10.9", "S": "6.1",
     "T": "6.1", "C": "6.4", "M": "10.3",
     "N": "7.5", "Q": "9.0", "D": "6.5",
     "E": "8.0", "K": "11.3", "R": "11.0",
     "H": "8.5", "P": "6.2"
     }


# Questa funzione prende in input l'excel da ripulire e genera un file Excel pulito
def estrapola_da_excel(input_file):
    sheet1 = pd.read_excel(input_file, sheet_name="Foglio1")
    sheet2 = pd.read_excel(input_file, sheet_name="Foglio2")
    sheet3 = pd.read_excel(input_file, sheet_name="Foglio3")
    lista_sheets = [sheet1, sheet2, sheet3]

    lista1 = []
    lista2 = []
    lista3 = []
    lista_liste = [lista1, lista2, lista3]


    cont_sheet = 0
    for sheet in lista_sheets:
        for elem in sheet['MUTAZIONE']: 
            if type(elem) is str:
                if not elem[0].isnumeric():
                    if len(elem) <= 6: 
                        elem = elem.strip() 
                        aminoacido = elem[0]
                        locazione = elem[1:len(elem) - 1]
                        mutazione = elem[len(elem) - 1]
                        lista_liste[cont_sheet].append([aminoacido, locazione, mutazione])
                        
        cont_sheet += 1

    file_name, file_extension = splitext(input_file)
    final_filename = file_name + "_elaborato" + file_extension

    df1 = pd.DataFrame(lista_liste[0], columns=['Aminoacido', 'Locazione', 'Mutazione'])
    df2 = pd.DataFrame(lista_liste[1], columns=['Aminoacido', 'Locazione', 'Mutazione'])
    df3 = pd.DataFrame(lista_liste[2], columns=['Aminoacido', 'Locazione', 'Mutazione'])

    with pd.ExcelWriter(final_filename) as writer:
        df1.to_excel(writer, sheet_name='Foglio1', index=False)
        df2.to_excel(writer, sheet_name='Foglio2', index=False)
        df3.to_excel(writer, sheet_name='Foglio3', index=False)

    print("File '" + final_filename + "' scritto con successo!")
    return final_filename


# Questa funzione prende in input l'Excel pulito e lo scrive su un altro Excel con dati aggiuntivi
def genera_excel(input_file):
    sheet1 = pd.read_excel(input_file, sheet_name="Foglio1")
    sheet2 = pd.read_excel(input_file, sheet_name="Foglio2")
    sheet3 = pd.read_excel(input_file, sheet_name="Foglio3")
    lista_sheets = [sheet1, sheet2, sheet3]

    lista1 = []
    lista2 = []
    lista3 = []
    lista_liste = [lista1, lista2, lista3]

    cont = 0
    cont_sheet = 0
    for sheet in lista_sheets:
        for row in sheet.itertuples():  # Faster than iterows() row["Aminoacido"]
            print(cont, row.Aminoacido, row.Locazione, row.Mutazione)
            line_out = row.Aminoacido + str(row.Locazione) + row.Mutazione
            lista_liste[cont_sheet].append([line_out, d[row.Aminoacido], d[row.Mutazione]])
            cont += 1

        cont_sheet += 1

    file_name, file_extension = splitext(input_file)
    final_filename = file_name + "_finale" + file_extension

    # CAMBIA QUI I NOMI DELLE COLONNE DELL'EXCEL FINALE
    df1 = pd.DataFrame(lista_liste[0], columns=["Col1", "Col2", "Col3"])
    df2 = pd.DataFrame(lista_liste[1], columns=["Col1", "Col2", "Col3"])
    df3 = pd.DataFrame(lista_liste[2], columns=["Col1", "Col2", "Col3"])

    with pd.ExcelWriter(final_filename) as writer:
        df1.to_excel(writer, sheet_name='Foglio1', index=False)
        df2.to_excel(writer, sheet_name='Foglio2', index=False)
        df3.to_excel(writer, sheet_name='Foglio3', index=False)

    print("File " + final_filename + " scritto con successo!")


if __name__ == '__main__':
    nome_file_elaborato = estrapola_da_excel(
        "Pieroni_originale.xlsx") 
    genera_excel(nome_file_elaborato) 
