import pyodbc
import pandas as pd
import os


class ExcelExporterWithSummary:
    def __init__(self, db_connection_string):
        self.db_connection_string = db_connection_string

    def export_table_with_summary(self, table_name):
        try:
            # Pfad für die zu speichernde Datei festlegen (Desktop)
            desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
            output_file = os.path.join(desktop_path, f"{table_name}_Export.xlsx")

            # Datenbankverbindung herstellen
            print("🔄 Verbinde mit der Datenbank...")
            conn = pyodbc.connect(self.db_connection_string)

            # Alle Daten aus der Tabelle abrufen
            query = f"SELECT * FROM {table_name}"
            print(f"🔍 Lade vollständige Daten aus der Tabelle '{table_name}'...")
            df = pd.read_sql_query(query, conn)

            # Sicherstellen, dass wichtige Spalten vorhanden sind
            if 'Customer_Name' not in df.columns or 'RKMDAT' not in df.columns:
                print("❌ Fehler: Die Tabelle muss die Spalten 'Customer_Name' und 'RKMDAT' enthalten.")
                return

            # Erstellung der Zusammenfassung
            print("📊 Erstelle Zusammenfassungstabellen basierend auf 'RKMDAT' und 'Customer_Name'...")
            customer_names = df['Customer_Name'].unique()  # Eindeutige Kunden ermitteln
            rkmdat_values = df['RKMDAT'].unique()  # Eindeutige RKMDAT-Werte ermitteln

            # Leeres DataFrame für die Zusammenfassung
            summary_df = pd.DataFrame(columns=['RKMDAT'] + list(customer_names))

            # Zusammenfassung aufbauen
            for rkmdat in rkmdat_values:
                row = {'RKMDAT': rkmdat}  # Start mit der RKMDAT-Wert
                for customer in customer_names:
                    # Anzahl der entsprechenden Zeilen zählen
                    count = len(df[(df['RKMDAT'] == rkmdat) & (df['Customer_Name'] == customer)])
                    row[customer] = count  # Zähler einfügen
                summary_df = pd.concat([summary_df, pd.DataFrame([row])], ignore_index=True)

            # Excel-Datei mit zwei Sheets speichern
            print("💾 Speichere Daten in die Excel-Datei...")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Sheet1: Alle Daten
                df.to_excel(writer, index=False, sheet_name="Sheet1")

                # Sheet2: Zusammenfassung
                summary_df.to_excel(writer, index=False, sheet_name="Sheet2")

            # Verbindung schließen
            conn.close()
            print(f"✅ Export erfolgreich! Datei gespeichert unter: {output_file}")

        except Exception as e:
            print(f"❌ Fehler beim Exportieren: {e}")


# Hauptprogramm
if __name__ == '__main__':
    # Verbindungseinstellungen für die Datenbank
    db_connection_string = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=your_server_name;"  # Anpassen
        "DATABASE=your_database_name;"  # Anpassen
        "Trusted_Connection=yes;"  # Windows-Authentifizierung
    )

    # Name der Tabelle
    table_name = "your_table_name"  # Anpassen

    # Export starten
    exporter = ExcelExporterWithSummary(db_connection_string)
    exporter.export_table_with_summary(table_name)
