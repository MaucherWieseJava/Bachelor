import pyodbc
import pandas as pd


class DatabaseToExcelExporter:
    def __init__(self, db_connection_string, output_file):
        """
        Initialisierung des Exporters mit Verbindungsdetails zur Datenbank und der Ausgabedatei.
        """
        self.db_connection_string = db_connection_string
        self.output_file = output_file

    def export_to_excel(self, queries_with_sheets):
        """
        Führt die SQL-Abfragen aus und schreibt die Ergebnisse direkt in eine Excel-Datei.

        :param queries_with_sheets: Ein Dictionary, das SQL-Abfragen den gewünschten Sheet-Namen zuordnet.
        """
        try:
            # Verbindung zur Datenbank herstellen
            conn = pyodbc.connect(self.db_connection_string)

            # Excel-Writer für mehrere Sheets
            with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
                for sheet_name, query in queries_with_sheets.items():
                    print(f"🔄 Exportiere Daten für Sheet: {sheet_name} ...")

                    # SQL-Abfrage ausführen und Daten abrufen
                    df = pd.read_sql_query(query, conn)

                    # Daten in das entsprechende Sheet der Excel-Datei schreiben
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Verbindung zur Datenbank schließen
            conn.close()
            print(f"✅ Export erfolgreich! Datei gespeichert unter: {self.output_file}")

        except Exception as e:
            print(f"❌ Fehler beim Exportieren: {e}")


# Hauptprogramm
if __name__ == '__main__':
    # Verbindungseinstellungen für die Datenbank (z. B. SQL Server)
    db_connection_string = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=your_server_name;"
        "DATABASE=your_database_name;"
        "Trusted_Connection=yes;"
    )

    # Ziel-Excel-Datei
    output_file = "Database_Export.xlsx"

    # SQL-Abfragen mit den entsprechenden Sheet-Namen
    queries_with_sheets = {
        "Tabelle1": "SELECT * FROM your_table1",  # Daten für Sheet "Tabelle1"
        "Tabelle2": "SELECT * FROM your_table2",  # Daten für Sheet "Tabelle2"
        # Füge hier weitere Queries und Sheets hinzu
    }

    # Exporter-Klasse initialisieren und den Export starten
    exporter = DatabaseToExcelExporter(db_connection_string, output_file)
    exporter.export_to_excel(queries_with_sheets)
