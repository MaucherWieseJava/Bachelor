import pyodbc
import pandas as pd
import os


class ExcelExporterWithSummary:
    def __init__(self, db_connection_string):
        self.db_connection_string = db_connection_string

    def export_table_with_summary(self, table_name):
        try:
            # Definiere den Pfad für die zu speichernde Datei (Desktop)
            desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
            output_file = os.path.join(desktop_path, f"{table_name}_Export.xlsx")

            # Verbinde mit der Datenbank
            print("🔄 Verbinde mit der Datenbank...")
            conn = pyodbc.connect(self.db_connection_string)

            # Lade alle Daten aus der Tabelle
            query = f"SELECT * FROM [dbo].[Tabelle1$]"
            print(f"🔍 Lade Daten aus der Tabelle '{table_name}'...")
            df = pd.read_sql_query(query, conn)

            # Sicherstellen, dass die erforderlichen Spalten in der Tabelle vorhanden sind
            required_columns = ['Customer Name', 'RKMDAT', 'DELLAT', 'Deletion Type', 'Amount']
            for col in required_columns:
                if col not in df.columns:
                    print(f"❌ Fehler: Spalte '{col}' fehlt in der Tabelle.")
                    return

            # Eindeutige Kunden, Spaltenwerte und Beträge abrufen
            customer_names = df['Customer Name'].unique()
            rkmdat_values = df['RKMDAT'].unique()
            amount_values = df['Amount'].unique()

            # Generiere Sheet2 (Zusammenfassung RKMDAT und Customer Name)
            print("📊 Erstelle Zusammenfassungen für RKMDAT...")
            summary_df = self.create_summary_with_amounts(df, 'RKMDAT', customer_names, rkmdat_values, amount_values)

            # Generiere Sheet3 (Deletion Type 1, 2, 5 für RKMDAT)
            print("📊 Erstelle gefilterte Daten für Deletion Type (1, 2, 5)...")
            filtered_summary_df_3 = self.create_filtered_summary_with_amounts(
                df, 'RKMDAT', [1, 2, 5], customer_names, rkmdat_values, amount_values
            )

            # Generiere Sheet4 (Deletion Type 3, 4, 6, 7 für RKMDAT)
            print("📊 Erstelle gefilterte Daten für Deletion Type (3, 4, 6, 7)...")
            filtered_summary_df_4 = self.create_filtered_summary_with_amounts(
                df, 'RKMDAT', [3, 4, 6, 7], customer_names, rkmdat_values, amount_values
            )

            # Generiere Sheet5 (Deletion Type 1, 2, 5 für DELLAT)
            print("📊 Erstelle gefilterte Daten für Deletion Type (1, 2, 5) mit 'DELLAT'...")
            dellat_values = df['DELLAT'].unique()
            filtered_summary_df_5 = self.create_filtered_summary_with_amounts(
                df, 'DELLAT', [1, 2, 5], customer_names, dellat_values, amount_values
            )

            # Speichere die Daten in die Excel-Datei
            print("💾 Speichere Daten in die Excel-Datei...")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Sheet1: Alle Daten
                df.to_excel(writer, index=False, sheet_name="Gesamtdaten")

                # Sheet2: Zusammenfassung (RKMDAT)
                summary_df.to_excel(writer, index=False, sheet_name="#NZG")

                # Sheet3: Gefilterte Zusammenfassung (Deletion Type 1, 2, 5 für RKMDAT)
                filtered_summary_df_3.to_excel(writer, index=False, sheet_name="Widerrufe")

                # Sheet4: Gefilterte Zusammenfassung (Deletion Type 3, 4, 6, 7 für RKMDAT)
                filtered_summary_df_4.to_excel(writer, index=False, sheet_name="#Kündigungen")

                # Sheet5: Gefilterte Zusammenfassung (Deletion Type 1, 2, 5 für DELLAT)
                filtered_summary_df_5.to_excel(writer, index=False, sheet_name="#widfürRainer")

            conn.close()
            print(f"✅ Export erfolgreich! Datei gespeichert unter: {output_file}")

        except Exception as e:
            print(f"❌ Fehler beim Export: {e}")

    def create_summary_with_amounts(self, df, column, customer_names, unique_values, amount_values):
        # Erstellt eine Zusammenfassung, die alle 'Amount'-Werte als separate Spalten enthält
        cols = [column] + [f"{customer} - {amount}" for customer in customer_names for amount in amount_values]
        summary_df = pd.DataFrame(columns=cols)

        for value in unique_values:
            row = {column: value}
            for customer in customer_names:
                for amount in amount_values:
                    count = len(df[(df[column] == value) &
                                   (df['Customer Name'] == customer) &
                                   (df['Amount'] == amount)])
                    row[f"{customer} - {amount}"] = count
            summary_df = pd.concat([summary_df, pd.DataFrame([row])], ignore_index=True)

        return summary_df.sort_values(by=column, ascending=True)

    def create_filtered_summary_with_amounts(self, df, column, deletion_types, customer_names, unique_values,
                                             amount_values):
        # Filtert die Daten nach Deletion Types und erstellt die Zusammenfassung
        filtered_df = df[df['Deletion Type'].isin(deletion_types)]
        return self.create_summary_with_amounts(filtered_df, column, customer_names, unique_values, amount_values)


# Hauptprogramm
if __name__ == '__main__':
    db_connection_string = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=your_server_name;"  # Anpassen
        "DATABASE=your_database_name;"  # Anpassen
        "Trusted_Connection=yes;"
    )

    table_name = "EXSB"  # Anpassen

    exporter = ExcelExporterWithSummary(db_connection_string)
    exporter.export_table_with_summary(table_name)
