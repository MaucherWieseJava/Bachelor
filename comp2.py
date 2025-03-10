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
            special_customer = 'FO-SCL'  # Nur Customer FO-SCL wird unterteilt

            # Generiere Sheet2 (Zusammenfassung RKMDAT und Customer Name)
            print("📊 Erstelle Zusammenfassungen für RKMDAT...")
            summary_df = self.create_summary_with_special_handling(
                df, 'RKMDAT', customer_names, rkmdat_values, amount_values, special_customer
            )

            # Generiere Sheet3 (Deletion Type 1, 2, 5 für RKMDAT)
            print("📊 Erstelle gefilterte Daten für Deletion Type (1, 2, 5)...")
            filtered_summary_df_3 = self.create_filtered_summary_with_special_handling(
                df, 'RKMDAT', [1, 2, 5], customer_names, rkmdat_values, amount_values, special_customer
            )

            # Berechne die Widerrufsquote und füge sie als separates Sheet hinzu
            print("🔢 Berechne Widerrufsquote und erstelle neues Worksheet...")
            widerrufsquote_df = self.calculate_widerrufsquote(summary_df, filtered_summary_df_3)

            # Generiere Sheet6 (CPO_NZG basierend auf Sheet2)
            print("📂 Erstelle Worksheet für CPO_NZG...")
            cpo_nzg_df = self.create_cpo_nzg(summary_df)

            # Generiere Sheet4 (Deletion Type 3, 4, 6, 7 für RKMDAT)
            print("📊 Erstelle gefilterte Daten für Deletion Type (3, 4, 6, 7)...")
            filtered_summary_df_4 = self.create_filtered_summary_with_special_handling(
                df, 'RKMDAT', [3, 4, 6, 7], customer_names, rkmdat_values, amount_values, special_customer
            )

            # Generiere Sheet5 (Deletion Type 1, 2, 5 für DELLAT)
            print("📊 Erstelle gefilterte Daten für Deletion Type (1, 2, 5) mit 'DELLAT'...")
            dellat_values = df['DELLAT'].unique()
            filtered_summary_df_5 = self.create_filtered_summary_with_special_handling(
                df, 'DELLAT', [1, 2, 5], customer_names, dellat_values, amount_values, special_customer
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

                # Sheet6: Widerrufsquote
                widerrufsquote_df.to_excel(writer, index=False, sheet_name="#Widerrufsquote")

                # Sheet7: CPO_NZG
                cpo_nzg_df.to_excel(writer, index=False, sheet_name="#CPO_NZG")

            conn.close()
            print(f"✅ Export erfolgreich! Datei gespeichert unter: {output_file}")

        except Exception as e:
            print(f"❌ Fehler beim Export: {e}")

    def create_summary_with_special_handling(self, df, column, customer_names, unique_values, amount_values,
                                             special_customer):
        # Erstellt eine Zusammenfassung: Für FO-SCL werden Amount-Spalten erstellt, für andere normal
        columns = [column]
        for customer in customer_names:
            if customer == special_customer:
                columns += [f"{customer} - {amount}" for amount in amount_values]
            else:
                columns.append(customer)

        summary_df = pd.DataFrame(columns=columns)

        for value in unique_values:
            row = {column: value}
            for customer in customer_names:
                if customer == special_customer:
                    for amount in amount_values:
                        count = len(df[(df[column] == value) &
                                       (df['Customer Name'] == customer) &
                                       (df['Amount'] == amount)])
                        row[f"{customer} - {amount}"] = count
                else:
                    count = len(df[(df[column] == value) & (df['Customer Name'] == customer)])
                    row[customer] = count
            summary_df = pd.concat([summary_df, pd.DataFrame([row])], ignore_index=True)

        return summary_df.sort_values(by=column, ascending=True)

    def calculate_widerrufsquote(self, summary_df, widerrufe_df):
        """
        Berechnet die Widerrufsquote:
        Widerrufsquote (%) = (Widerrufe / #NZG) * 100
        """
        if summary_df.shape != widerrufe_df.shape:
            raise ValueError("Die Tabellen für #NZG und Widerrufe müssen die gleiche Struktur haben.")

        widerrufsquote_df = widerrufe_df.copy()

        for col in widerrufsquote_df.columns[1:]:  # Ignoriere die erste Spalte (Datumswerte)
            widerrufsquote_df[col] = widerrufe_df[col] / summary_df[col].replace(0, pd.NA) * 100
            widerrufsquote_df[col] = widerrufsquote_df[col].fillna(0)  # Ersetze Division durch 0 mit 0

        widerrufsquote_df.columns = widerrufe_df.columns
        return widerrufsquote_df

    def create_cpo_nzg(self, nzg_df):
        """
        Erstellt das Sheet: #CPO_NZG
        Multipliziert Werte in jeder Zelle mit 59,9 oder 49,9, basierend auf dem Wert von RKMDAT (Datumswert in Spalte A).
        """
        cpo_nzg_df = nzg_df.copy()

        for index, row in cpo_nzg_df.iterrows():
            if index == 0:  # Überspringe Überschriften
                continue

            rkmdat = row[0]
            try:
                rkmdat = int(rkmdat)
                factor = 59.9 if rkmdat > 202206 else 49.9
            except ValueError:
                continue

            for col in cpo_nzg_df.columns[1:]:
                if pd.notna(row[col]) and isinstance(row[col], (int, float)):
                    cpo_nzg_df.at[index, col] = row[col] * factor

        return cpo_nzg_df


# Hauptprogramm
if __name__ == '__main__':
    db_connection_string = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=your_server_name;"  # Anpassen
        "DATABASE=your_database_name;"  # Anpassen
        "Trusted_Connection=yes;"
    )

    table_name = "your_table_name"  # Anpassen

    exporter = ExcelExporterWithSummary(db_connection_string)
    exporter.export_table_with_summary(table_name)
