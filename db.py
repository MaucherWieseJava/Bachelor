import pyodbc
import csv
import os

# Verbindung zur MSSQL-Datenbank mit Windows Authentication
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=your_server_name;"
    "DATABASE=your_database_name;"
    "Trusted_Connection=yes;"
)

cursor = conn.cursor()

# SQL-Abfrage - hier ggf. anpassen, um nicht zu viele Daten auf einmal zu laden
query = "SELECT * FROM your_table"
cursor.execute(query)

# Spaltennamen abrufen
columns = [column[0] for column in cursor.description]

# Desktop-Pfad abrufen
desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
csv_file = os.path.join(desktop_path, "exported_data.csv")

# **BATCH-Größe festlegen (z. B. 1000 Zeilen pro Durchlauf)**
batch_size = 1000

# CSV-Datei schreiben
with open(csv_file, mode="w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file, delimiter=";", quotechar='"', quoting=csv.QUOTE_MINIMAL)

    # Spaltennamen in die erste Zeile der CSV-Datei schreiben
    writer.writerow(columns)

    # Daten blockweise abrufen und schreiben
    while True:
        rows = cursor.fetchmany(batch_size)  # **Holt max. 1000 Zeilen auf einmal**
        if not rows:
            break  # **Falls keine Daten mehr da sind, abbrechen**

        writer.writerows(rows)  # **Daten in die CSV-Datei schreiben**

# Verbindung zur Datenbank schließen
conn.close()

print(f"✅ Export erfolgreich! Datei gespeichert unter: {csv_file}")
