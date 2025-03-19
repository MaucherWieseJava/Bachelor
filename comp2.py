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
                df, 'DELLAT', [3, 4, 6, 7], customer_names, rkmdat_values, amount_values, special_customer
            )

            # Generiere Sheet5 (Deletion Type 1, 2, 5 für DELLAT)
            print("📊 Erstelle gefilterte Daten für Deletion Type (1, 2, 5) mit 'DELLAT'...")
            dellat_values = df['DELLAT'].unique()
            filtered_summary_df_5 = self.create_filtered_summary_with_special_handling(
                df, 'DELLAT', [1, 2, 5], customer_names, dellat_values, amount_values, special_customer
            )

            # Erstelle Worksheet für CPO_WID
            print("📂 Erstelle Worksheet für CPO_WID...")
            cpo_wid_df = self.create_cpo_wid(filtered_summary_df_5)

            print("📂 Erstelle Worksheet UFC_NZG...")
            ufc_nzg_df = self.create_ufc_nzg(df)

            print("📂 Erstelle Worksheet UFC_WID...")
            ufc_wid_df = self.create_ufc_wid(df)

            print("📂 Erstelle Worksheet UFC_KÜN...")
            ufc_kün_df = self.create_ufc_kün(df)

            # Speichere die Daten in die Excel-Datei
            print("💾 Speichere Daten in die Excel-Datei...")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Sheet1: Alle Daten
                df.to_excel(writer, index=False, sheet_name="Gesamtdaten")

                # Sheet2: Zusammenfassung (RKMDAT)
                summary_df.to_excel(writer, index=False, sheet_name="#NZG")

                # Sheet3: Gefilterte Zusammenfassung (Deletion Type 1, 2, 5 für RKMDAT)
                filtered_summary_df_3.to_excel(writer, index=False, sheet_name="Widerrufe")

                # Sheet4: Gefilterte Zusammenfassung (Deletion Type 3, 4, 6, 7 für DELLAT)
                filtered_summary_df_4.to_excel(writer, index=False, sheet_name="#Kündigungen")

                # Sheet5: Gefilterte Zusammenfassung (Deletion Type 1, 2, 5 für DELLAT)
                filtered_summary_df_5.to_excel(writer, index=False, sheet_name="#widfürRainer")

                # Sheet6: Widerrufsquote
                widerrufsquote_df.to_excel(writer, index=False, sheet_name="#Widerrufsquote")

                # Sheet7: CPO_NZG
                cpo_nzg_df.to_excel(writer, index=False, sheet_name="#CPO_NZG")

                # Sheet8: CPO_WID
                cpo_wid_df.to_excel(writer, index=False, sheet_name="#CPO_WID")

                ufc_nzg_df.to_excel(writer, index=True, sheet_name="UFC_NZG")

                ufc_wid_df.to_excel(writer, index=True, sheet_name="UFC_WID")

                ufc_kün_df.to_excel(writer, index=True, sheet_name="UFC_KÜN")

            conn.close()
            print(f"✅ Export erfolgreich! Datei gespeichert unter: {output_file}")

        except Exception as e:
            print(f"❌ Fehler beim Export: {e}")

    def create_summary_with_special_handling(self, df, column, customer_names, unique_values, amount_values,
                                             special_customer):
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

    def create_filtered_summary_with_special_handling(self, df, column, deletion_types, customer_names, unique_values,
                                                      amount_values, special_customer):
        filtered_df = df[df['Deletion Type'].isin(deletion_types)]
        return self.create_summary_with_special_handling(
            filtered_df, column, customer_names, unique_values, amount_values, special_customer
        )

    def calculate_widerrufsquote(self, summary_df, widerrufe_df):
        if summary_df.shape != widerrufe_df.shape:
            raise ValueError("Die Tabellen für #NZG und Widerrufe müssen die gleiche Struktur haben.")

        widerrufsquote_df = widerrufe_df.copy()

        for col in widerrufsquote_df.columns[1:]:
            widerrufsquote_df[col] = widerrufe_df[col] / summary_df[col].replace(0, pd.NA) * 100
            widerrufsquote_df[col] = widerrufsquote_df[col].fillna(0)

        widerrufsquote_df.columns = widerrufe_df.columns
        return widerrufsquote_df

    def create_cpo_nzg(self, nzg_df):
        cpo_nzg_df = nzg_df.copy()

        for index, row in cpo_nzg_df.iterrows():
            if index == 0 and not isinstance(row[0], (int, float)):
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

    def create_cpo_wid(self, filtered_summary_df_5):
        if not isinstance(filtered_summary_df_5, pd.DataFrame) or filtered_summary_df_5.empty:
            raise ValueError("Fehler: filtered_summary_df_5 ist leer oder ungültig.")

        cpo_wid_df = filtered_summary_df_5.copy()

        for index, row in cpo_wid_df.iterrows():
            if index == 0 and not isinstance(row[0], (int, float)):
                continue

            dellat = row[0]
            try:
                dellat = int(dellat)
                factor1 = 59.9 if dellat > 202206 else 49.9
            except ValueError:
                continue

            for col in cpo_wid_df.columns[1:]:
                if pd.notna(row[col]) and isinstance(row[col], (int, float)):
                    cpo_wid_df.at[index, col] = row[col] * factor1

        return cpo_wid_df

    def create_ufc_nzg(self, df):
        # Definierte Variablen
        special_customer = "FO-SCL"
        amounts = [9.9, 14.9, 19.9, 24.9, 29.9]
        rkmdat_threshold = 202206
        ufc_values = {
            9.9: 107.9949,
            14.9: 162.5167,
            19.9: 217.0588,
            24.9: 271.6009,
            29.9: 326.1226
        }

        # Filtere Daten für Customer Name = FO-SCL
        filtered_df = df[df['Customer Name'] == special_customer]

        # Generiere eine neue DataFrame für das Worksheet
        summary_data = []
        grouped = filtered_df.groupby(['RKMDAT', 'Amount'])

        # Gruppiere nach rkmdat und den Amount-Werten
        for (rkmdat, amount), group in grouped:
            if amount in amounts:
                count = len(group)  # Anzahl der Einträge
                summary_data.append({'RKMDAT': rkmdat, 'Amount': amount, 'Count': count})

        # Umwandeln in DataFrame für Pivot-Tabelle
        summary_df = pd.DataFrame(summary_data)

        # Pivot: Zeilen sind 'RKMDAT', Spalten sind 'Amount', Werte sind 'Count'
        pivot_df = summary_df.pivot(index='RKMDAT', columns='Amount', values='Count').fillna(0)
        pivot_df = pivot_df.sort_index()

        # Umbenennen der Spaltennamen mit "FO-SCL-" Präfix
        pivot_df.columns = [f"{special_customer}-{amount}" for amount in pivot_df.columns]

        # Multiplikation für RKMDAT > 202206
        for amount, ufc_value in ufc_values.items():
            column_name = f"{special_customer}-{amount}"
            if column_name in pivot_df.columns:
                # Werte setzen: 0 für rkmdat ≤ 202206, Multiplikation für rkmdat > 202206
                pivot_df[column_name] = pivot_df.apply(
                    lambda row: row[column_name] * ufc_value if row.name > rkmdat_threshold else 0,
                    axis=1
                )

        # Rückgabe der Pivot-Tabelle
        return pivot_df

    def create_ufc_wid(self, df):
        # Definierte Variablen
        special_customer = "FO-SCL"
        amounts = [9.9, 14.9, 19.9, 24.9, 29.9]
        dellat_threshold = 202206
        ufc_values = {
            9.9: 107.9949,
            14.9: 162.5167,
            19.9: 217.0588,
            24.9: 271.6009,
            29.9: 326.1226
        }

        # Filtere Daten für Customer Name = FO-SCL und Deletion Type in [1, 2, 5], sowie gleiche Werte bei RKMDAT und DELLAT
        filtered_df = df[
            (df['Customer Name'] == special_customer) &
            (df['Deletion Type'].isin([1, 2, 5])) &
            (df['RKMDAT'] == df['DELLAT'])  # Nur Datensätze mit gleichen Werten in RKMDAT und DELLAT
            ]

        # Generiere eine neue DataFrame für das Worksheet
        summary_data = []
        grouped = filtered_df.groupby(['DELLAT', 'Amount'])

        # Gruppiere nach dellat und den Amount-Werten
        for (dellat, amount), group in grouped:
            if amount in amounts:
                count = len(group)  # Anzahl der Einträge
                summary_data.append({'DELLAT': dellat, 'Amount': amount, 'Count': count})

        # Umwandeln in DataFrame für Pivot-Tabelle
        summary_df = pd.DataFrame(summary_data)

        # Pivot: Zeilen sind 'DELLAT', Spalten sind 'Amount', Werte sind 'Count'
        pivot_df = summary_df.pivot(index='DELLAT', columns='Amount', values='Count').fillna(0)
        pivot_df = pivot_df.sort_index()

        # Umbenennen der Spaltennamen mit "FO-SCL-" Präfix
        pivot_df.columns = [f"{special_customer}-{amount}" for amount in pivot_df.columns]

        # Multiplikation für DELLAT > 202206
        for amount, ufc_value in ufc_values.items():
            column_name = f"{special_customer}-{amount}"
            if column_name in pivot_df.columns:
                # Werte setzen: 0 für dellat ≤ 202206, Multiplikation für dellat > 202206
                pivot_df[column_name] = pivot_df.apply(
                    lambda row: row[column_name] * ufc_value if row.name > dellat_threshold else 0,
                    axis=1
                )

        # Rückgabe der Pivot-Tabelle
        return pivot_df

    def create_ufc_kün(self, df):
        """
        Erstellt die UFC-KÜN-Zusammenfassung basierend auf FirstDueDate und Last Due Date.
        Wendet die Berechnung gemäß der Excel-Funktion an:
        WENN(H4<=L4; DATEDIF(H4; L4; "m") + 1; 0).
        """
        # Definierte Variablen
        special_customer = "FO-SCL"
        amounts = [9.9, 14.9, 19.9, 24.9, 29.9]
        ufc_values = {
            9.9: 107.9949,
            14.9: 162.5167,
            19.9: 21,
            24.9: 271.6009,
            29.9: 326.1226
        }

        # Filtere relevante Daten (FO-SCL und spezifische Deletion Types)
        filtered_df = df[
            (df['Customer Name'] == special_customer) &
            (df['Deletion Type'].isin([3, 4, 6, 7])) &  # Nur spezifische Deletion Types
            (df['RKMDAT'] > 202206)  # Bedingung für RKMDAT
            ].copy()

        # Konvertiere die relevanten Spalten in Datum
        filtered_df['FirstDueDate'] = pd.to_datetime(filtered_df['FirstDueDate'], errors='coerce')
        filtered_df['Last Due Date'] = pd.to_datetime(filtered_df['Last Due Date'], errors='coerce')

        # Hilfsfunktion: Berechnung der Differenz in vollen Monaten (DATEDIF + 1)
        def calculate_month_difference(first_due_date, last_due_date):
            if first_due_date <= last_due_date:
                # Berechne die Differenz in Jahren und Monaten
                years_diff = last_due_date.year - first_due_date.year
                months_diff = last_due_date.month - first_due_date.month
                total_months = years_diff * 12 + months_diff

                # Prüfe, ob der letzte Tag kleiner ist als der erste, und korrigiere
                if last_due_date.day < first_due_date.day:
                    total_months -= 1

                # +1 zur Differenz hinzufügen (gemäß Excel-Formel)
                return total_months + 1
            else:
                # Rückgabe 0, wenn FirstDueDate > Last Due Date
                return 0

        # Berechne die Monate mithilfe der Funktion und füge sie als neue Spalte hinzu
        filtered_df['Months Difference'] = filtered_df.apply(
            lambda row: calculate_month_difference(row['FirstDueDate'], row['Last Due Date']),
            axis=1
        )

        # Filtere nur Datensätze, bei denen die Differenz < 24 ist
        filtered_df = filtered_df[filtered_df['Months Difference'] < 24]

        # Berechnung der Beträge für jeden Betrag (Amount)
        summary_data = []
        grouped = filtered_df.groupby(['DELLAT', 'Amount'])

        for (dellat, amount), group in grouped:
            if amount in amounts:
                total = 0
                for _, row in group.iterrows():
                    # Berechnung des Wertes basierend auf der Monatsdifferenz
                    months = row['Months Difference']
                    ufc_value = ufc_values[amount]
                    value = (1 - (months / 24)) * ufc_value
                    total += value

                summary_data.append({'DELLAT': dellat, f"{special_customer}-{amount}": total})

        # Umwandeln in DataFrame
        summary_df = pd.DataFrame(summary_data)

        # Pivot: Zeilen sind 'DELLAT', Spalten sind dynamisch benannte Beträge (FO-SCL-{Amount})
        pivot_df = summary_df.pivot_table(index='DELLAT', aggfunc='sum').fillna(0)

        pivot_df = pivot_df.sort_index()  # Sortiere die Tabelle nach DELLAT

        # Rückgabe der berechneten Pivot-Tabelle
        return pivot_df


# Hauptprogramm
if __name__ == '__main__':
    db_connection_string = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=your_server_name;"
        "DATABASE=your_database_name;"
        "Trusted_Connection=yes;"
    )

    table_name = "your_table_name"

    exporter = ExcelExporterWithSummary(db_connection_string)
    exporter.export_table_with_summary(table_name)
