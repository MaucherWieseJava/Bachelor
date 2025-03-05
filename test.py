import pandas as pd


# Funktion zum Verarbeiten der CSV-Datei und Erstellen der Berechnungen
def process_csv(input_file, edrk_file_output, nzg_file_output):
    # Eingabe-CSV-Datei lesen
    df = pd.read_csv(input_file, delimiter=";")  # Semikolon als Trennzeichen, anpassen falls anders

    # Neues DataFrame für #NZG erstellen
    products = df['N'].unique()  # Eindeutige Versicherungsprodukte (Spalte N)
    months = df['O'].unique()  # Eindeutige Monate (Spalte O)

    # #NZG DataFrame initialisieren
    nzg_data = pd.DataFrame(columns=['Produkt'] + list(months))  # Spalten: 'Produkt' + Monate
    nzg_data['Produkt'] = products

    # Berechnungen nach ZÄHLENWENNS-Logik
    for product in products:
        for month in months:
            # Filtere Zeilen nach Produkt und Monat
            count = df[(df['N'] == product) & (df['O'] == month)].shape[0]
            # Setze den berechneten Wert im #NZG DataFrame
            nzg_data.loc[nzg_data['Produkt'] == product, month] = count

    # CSV-Dateien speichern
    df.to_csv(edrk_file_output, index=False, sep=";")  # Originaldaten speichern
    nzg_data.to_csv(nzg_file_output, index=False, sep=";")  # Ergebnisse speichern

    print(f"Dateien erfolgreich erstellt: {edrk_file_output}, {nzg_file_output}")


# Hauptprogramm
if __name__ == '__main__':
    # Definition der Datei-Pfade
    input_file = 'ED-RK.csv'  # Eingabedatei im CSV-Format
    edrk_output = 'ED-RK-Processed.csv'  # Ausgabe der Originaldaten
    nzg_output = 'NZG-Processed.csv'  # Ausgabe der berechneten Werte

    # Funktion ausführen
    process_csv(input_file, edrk_output, nzg_output)
