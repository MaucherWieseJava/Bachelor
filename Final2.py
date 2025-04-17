# =====================================================================
# AUTOR: @Adrian Stötzler

import pandas as pd
import os

class ExcelExporterWithSummary:
    def __init__(self, excel_file_path):
        self.excel_file_path = excel_file_path

    def export_table_with_summary(self):
        try:
            desktop_path = os.path.join(os.environ["HOME"], "Desktop")
            output_file = os.path.join(desktop_path, "EXSB_Export.xlsx")

            # Lade die Excel-Daten
            print(f"🔄 Lade Daten aus der Excel-Datei: {self.excel_file_path}...")
            df = pd.read_excel(self.excel_file_path, engine='openpyxl')

            # Sicherstellen, dass die erforderlichen Spalten in der Tabelle vorhanden sind
            required_columns = ['Product Code', 'Kampagne', 'RKMDAT', 'DELLAT', 'Deletion Type', 'Amount']
            for col in required_columns:
                if col not in df.columns:
                    print(f"❌ Fehler: Spalte '{col}' fehlt in der Tabelle.")
                    return

            # Eindeutige Kunden, Spaltenwerte und Beträge abrufen
            customer_names = df['Kampagne'].unique()
            rkmdat_values = df['RKMDAT'].unique()
            amount_values = df['Amount'].unique()
            special_customer = 'FO-SCL'  # Nur Customer FO-SCL wird unterteilt

            print("📊 Erstelle Zusammenfassungen für RKMDAT...")
            summary_df = self.create_summary_with_special_handling(
                df, 'RKMDAT', customer_names, rkmdat_values, amount_values, special_customer
            )

            print("📊 Erstelle gefilterte Daten für Deletion Type (1, 2, 5)...")
            filtered_summary_df_3 = self.create_filtered_summary_with_special_handling(
                df, 'RKMDAT', [1, 2, 5], customer_names, rkmdat_values, amount_values, special_customer
            )

            print("🔢 Berechne Widerrufsquote und erstelle neues Worksheet...")
            widerrufsquote_df = self.calculate_widerrufsquote(summary_df, filtered_summary_df_3)

            print("📂 Erstelle Worksheet für CPO_NZG...")
            cpo_nzg_df = self.create_cpo_nzg(summary_df, df)  # Hinzufügen von df als Argument

            print("📊 Erstelle gefilterte Daten für Deletion Type (3, 4, 6, 7)...")
            filtered_summary_df_4 = self.create_filtered_summary_with_special_handling(
                df, 'DELLAT', [3, 4, 6, 7], customer_names, rkmdat_values, amount_values, special_customer
            )


            print("📊 Erstelle gefilterte Daten für Deletion Type (1, 2, 5) mit 'DELLAT'...")
            dellat_values = df['DELLAT'].unique()
            filtered_summary_df_5 = self.create_filtered_summary_with_special_handling(
                df, 'DELLAT', [1, 2, 5], customer_names, dellat_values, amount_values, special_customer
            )

            print("📂 Erstelle Worksheet für CPO_WID...")
            cpo_wid_df = self.create_cpo_wid(filtered_summary_df_5, df)  # Hinzufügen von df als Argument

            if cpo_wid_df is None:
                print("❌ Fehler: create_cpo_wid hat None zurückgegeben.")
                return

            print("📂 Erstelle Worksheet für CPO_NZG (3783)...")
            try:
                cpo_nzg_df_3783 = self.create_cpo_nzg_3783(summary_df, df)
            except Exception as e:
                print(f"⚠️ Fehler bei CPO_NZG (3783): {e}")
                cpo_nzg_df_3783 = pd.DataFrame()

            print("📂 Erstelle Worksheet für CPO_WID (3783)...")
            try:
                cpo_wid_df_3783 = self.create_cpo_wid_3783(filtered_summary_df_5, df)
            except Exception as e:
                print(f"⚠️ Fehler bei CPO_WID (3783): {e}")
                cpo_wid_df_3783 = pd.DataFrame()


            print("📂 Erstelle Worksheet UFC_NZG...")
            ufc_nzg_df = self.create_ufc_nzg(df)
            if ufc_nzg_df is None:
                print("❌ Fehler: create_ufc_nzg hat None zurückgegeben.")
                return

            print("📂 Erstelle Worksheet UFC_WID...")
            ufc_wid_df = self.create_ufc_wid(df)
            if ufc_wid_df is None:
                print("❌ Fehler: create_ufc_wid hat None zurückgegeben.")
                return

            print("📂 Erstelle Worksheet UFC_KÜN...")
            ufc_kün_df = self.create_ufc_kün(df)
            if ufc_kün_df is None:
                print("❌ Fehler: create_ufc_kün hat None zurückgegeben.")
                return

            print("📂 Erstelle Worksheet Result V2...")
            result_v2_df = self.create_result_v2(cpo_nzg_df, cpo_wid_df, filtered_summary_df_5, ufc_nzg_df, ufc_wid_df, ufc_kün_df, cpo_nzg_df_3783, cpo_wid_df_3783)
            if result_v2_df is None:
                print("❌ Fehler: create_result_v2 hat None zurückgegeben.")
                return


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

                if not cpo_nzg_df_3783.empty:
                    cpo_nzg_df_3783.to_excel(writer, index=False, sheet_name="#CPO_NZG_3783")
                if not cpo_wid_df_3783.empty:
                    cpo_wid_df_3783.to_excel(writer, index=False, sheet_name="#CPO_WID_3783")

                ufc_nzg_df.to_excel(writer, index=True, sheet_name="UFC_NZG")

                ufc_wid_df.to_excel(writer, index=True, sheet_name="UFC_WID")

                ufc_kün_df.to_excel(writer, index=True, sheet_name="UFC_KÜN")

                result_v2_df.to_excel(writer, index=False, sheet_name="Result V2")

            print(f"✅ Export erfolgreich! Datei gespeichert unter: {output_file}")

        except Exception as e:
            print(f"❌ Fehler beim Export: {e}")
            import traceback
            traceback.print_exc()

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
                                       (df['Kampagne'] == customer) &
                                       (df['Amount'] == amount)])
                        row[f"{customer} - {amount}"] = count
                else:
                    count = len(df[(df[column] == value) & (df['Kampagne'] == customer)])
                    row[customer] = count
            summary_df = pd.concat([summary_df, pd.DataFrame([row])], ignore_index=True)

        # Überprüfe und vereinheitliche die Datentypen in der Spalte für die Sortierung
        try:
            # Versuche zuerst, alle Werte in Zahlen umzuwandeln
            summary_df[column] = pd.to_numeric(summary_df[column], errors='raise')
        except:
            # Wenn das nicht funktioniert, wandle alles in Strings um
            summary_df[column] = summary_df[column].astype(str)

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

    def create_cpo_nzg(self, nzg_df, df):
        """
        Erstellt das Worksheet CPO_NZG für Produkt 3402
        """
        # Filtere nur Daten mit Product Code 3402
        filtered_df = df[df['Product Code'] == 3402]

        # Erstelle ein neues nzg_df basierend auf dem gefilterten DataFrame
        filtered_nzg_df = self.create_summary_with_special_handling(
            filtered_df, 'RKMDAT', filtered_df['Kampagne'].unique(),
            filtered_df['RKMDAT'].unique(), filtered_df['Amount'].unique(), 'FO-SCL'
        )

        cpo_nzg_df = filtered_nzg_df.copy()

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

    def create_cpo_nzg_3783(self, nzg_df, df):
        """
        Erstellt das Worksheet CPO_NZG für Produkt 3783 mit robuster Handhabung falls nur FO-SCL vorhanden
        """
        # Filtere nur Daten mit Product Code 3783
        filtered_df = df[df['Product Code'] == 3783]

        # Abbrechen, wenn keine Daten für 3783 vorhanden
        if filtered_df.empty:
            print("⚠️ Keine Daten für Produkt 3783 gefunden")
            return pd.DataFrame()

        # Vorhandene Kampagnen ermitteln
        available_campaigns = filtered_df['Kampagne'].unique()
        print(f"📊 Verfügbare Kampagnen für Produkt 3783: {', '.join(available_campaigns)}")

        # Erstelle ein neues nzg_df basierend auf dem gefilterten DataFrame
        filtered_nzg_df = self.create_summary_with_special_handling(
            filtered_df, 'RKMDAT', available_campaigns,
            filtered_df['RKMDAT'].unique(), filtered_df['Amount'].unique(), 'FO-SCL'
        )

        cpo_nzg_df = filtered_nzg_df.copy()

        for index, row in cpo_nzg_df.iterrows():
            if index == 0 and not isinstance(row[0], (int, float)):
                continue

            rkmdat = row[0]
            try:
                rkmdat = int(rkmdat)
                factor = 59.9 if rkmdat > 202206 else 49.9
            except ValueError:
                continue

            # Alle vorhandenen Spalten multiplizieren (ohne RKMDAT)
            for col in cpo_nzg_df.columns[1:]:
                if pd.notna(row[col]) and isinstance(row[col], (int, float)):
                    cpo_nzg_df.at[index, col] = row[col] * factor

        return cpo_nzg_df

    def create_cpo_wid(self, filtered_summary_df_5, df):
        """
        Erstellt das Worksheet CPO_WID für Produkt 3402
        """
        # Filtere nur Daten mit Product Code 3402
        filtered_df = df[df['Product Code'] == 3402]

        # Erzeuge eine neue filtered_summary_df_5 für die gefilterten Daten
        filtered_summary = self.create_filtered_summary_with_special_handling(
            filtered_df, 'DELLAT', [1, 2, 5], filtered_df['Kampagne'].unique(),
            filtered_df['DELLAT'].unique(), filtered_df['Amount'].unique(), 'FO-SCL'
        )

        if not isinstance(filtered_summary, pd.DataFrame) or filtered_summary.empty:
            raise ValueError("Fehler: filtered_summary für Produkt 3402 ist leer oder ungültig.")

        print("🔄 Starte Erstellung von cpo_wid_df für Produkt 3402...")

        cpo_wid_df = filtered_summary.copy()
        # Rest des Codes bleibt gleich
        if 'DELLAT' not in cpo_wid_df.columns:
            print("❌ Fehler: Spalte 'DELLAT' fehlt in cpo_wid_df.")
            return pd.DataFrame()
        else:
            print("✅ Spalte 'DELLAT' gefunden in cpo_wid_df.")

        # Hier weitermachen mit dem bestehenden Code...
        # [Rest der Funktion identisch zur ursprünglichen]

        try:
            min_dellat = int(cpo_wid_df['DELLAT'].min())
            max_dellat = int(cpo_wid_df['DELLAT'].max())
            dellat_range = pd.date_range(start=f"{min_dellat // 100}-{min_dellat % 100:02d}",
                                         end=f"{max_dellat // 100}-{max_dellat % 100:02d}",
                                         freq='MS').strftime('%Y%m').astype(int)
        except Exception as e:
            print(f"❌ Fehler beim Generieren der DELLAT-Werte: {e}")
            return pd.DataFrame()

        for dellat in dellat_range:
            if dellat not in cpo_wid_df['DELLAT'].values:
                cpo_wid_df = pd.concat([cpo_wid_df, pd.DataFrame({'DELLAT': [dellat]})], ignore_index=True)
                print(f"➕ Fehlender DELLAT-Wert hinzugefügt: {dellat}")

        cpo_wid_df = cpo_wid_df.sort_values(by='DELLAT').reset_index(drop=True)

        for index, row in cpo_wid_df.iterrows():
            if index == 0 and not isinstance(row.iloc[0], (int, float)):
                print(f"⚠️ Überspringe erste Zeile, da der Wert kein int oder float ist: {row.iloc[0]}")
                continue

            dellat = row.iloc[0]
            try:
                dellat = int(dellat)
                factor1 = 59.9 if dellat > 202206 else 49.9
                print(f"🔢 DELLAT Wert: {dellat}, Faktor: {factor1}")
            except ValueError as e:
                print(f"❌ Fehler beim Konvertieren von DELLAT: {e}")
                continue

            for col in cpo_wid_df.columns[1:]:
                if pd.notna(row[col]) and isinstance(row[col], (int, float)):
                    original_value = row[col]
                    cpo_wid_df.at[index, col] = row[col] * factor1
                    print(f"🔄 Aktualisiere Wert in Spalte '{col}' von {original_value} zu {cpo_wid_df.at[index, col]}")

        try:
            cpo_wid_df['DELLAT2'] = cpo_wid_df['DELLAT'].copy()
            print("✅ Spalte 'DELLAT2' erfolgreich hinzugefügt.")
        except Exception as e:
            print(f"❌ Fehler beim Hinzufügen der Spalte 'DELLAT2': {e}")

        # Neue Spalte 'FO' = alles zusammenaddiert 'FO-S2S', 'FO-ITM' und 'FO-OTM'
        try:
            cpo_wid_df['FO'] = cpo_wid_df[
                ['FO-S2S', 'FO-ITM', 'FO-OTM', 'FO-SCL - 19.9', 'FO-SCL - 9.9', 'FO-SCL - 14.9', 'FO-SCL - 29.9',
                 'FO-SCL - 24.9']].sum(axis=1)
            print("✅ Spalte 'FO' erfolgreich hinzugefügt.")
        except Exception as e:
            print(f"❌ Fehler beim Hinzufügen der Spalte 'FO': {e}")

        # Neue Spalte 'JD' = WErte aus 'JD-OTM' und 'JD-ITM' zusammenaddieren
        try:
            cpo_wid_df['JD'] = cpo_wid_df[['JD-OTM', 'JD-ITM']].sum(axis=1)
            print("✅ Spalte 'JD' erfolgreich hinzugefügt.")
        except Exception as e:
            print(f"❌ Fehler beim Hinzufügen der Spalte 'JD': {e}")

        try:
            cpo_wid_df_sorted = cpo_wid_df.sort_values(by='DELLAT')
            print("✅ DataFrame erfolgreich nach 'DELLAT' sortiert.")
        except Exception as e:
            print(f"❌ Fehler beim Sortieren des DataFrames nach 'DELLAT': {e}")

        return cpo_wid_df_sorted

    def create_cpo_wid_3783(self, filtered_summary_df_5, df):
        """
        Erstellt das Worksheet CPO_WID für Produkt 3783 mit robuster Handhabung falls nur FO-SCL vorhanden
        """
        # Filtere nur Daten mit Product Code 3783
        filtered_df = df[df['Product Code'] == 3783]

        # Abbrechen, wenn keine Daten für 3783 vorhanden
        if filtered_df.empty:
            print("⚠️ Keine Daten für Produkt 3783 gefunden")
            return pd.DataFrame()

        # Vorhandene Kampagnen ermitteln
        available_campaigns = filtered_df['Kampagne'].unique()
        print(f"📊 Verfügbare Kampagnen für Produkt 3783: {', '.join(available_campaigns)}")

        # Erzeuge eine neue filtered_summary_df_5 für die gefilterten Daten
        try:
            filtered_summary = self.create_filtered_summary_with_special_handling(
                filtered_df, 'DELLAT', [1, 2, 5], available_campaigns,
                filtered_df['DELLAT'].unique(), filtered_df['Amount'].unique(), 'FO-SCL'
            )

            if filtered_summary.empty:
                print("⚠️ Keine Widerrufsdaten für 3783 gefunden")
                return pd.DataFrame()

            cpo_wid_df = filtered_summary.copy()

            if 'DELLAT' not in cpo_wid_df.columns:
                print("❌ Fehler: Spalte 'DELLAT' fehlt in cpo_wid_df für 3783.")
                return pd.DataFrame()

            try:
                min_dellat = int(cpo_wid_df['DELLAT'].min())
                max_dellat = int(cpo_wid_df['DELLAT'].max())
                dellat_range = pd.date_range(start=f"{min_dellat // 100}-{min_dellat % 100:02d}",
                                             end=f"{max_dellat // 100}-{max_dellat % 100:02d}",
                                             freq='MS').strftime('%Y%m').astype(int)
            except Exception as e:
                print(f"❌ Fehler beim Generieren der DELLAT-Werte für 3783: {e}")
                return pd.DataFrame()

            for dellat in dellat_range:
                if dellat not in cpo_wid_df['DELLAT'].values:
                    cpo_wid_df = pd.concat([cpo_wid_df, pd.DataFrame({'DELLAT': [dellat]})], ignore_index=True)

            cpo_wid_df = cpo_wid_df.sort_values(by='DELLAT').reset_index(drop=True)

            for index, row in cpo_wid_df.iterrows():
                if index == 0 and not isinstance(row.iloc[0], (int, float)):
                    continue

                dellat = row.iloc[0]
                try:
                    dellat = int(dellat)
                    factor = 59.9 if dellat > 202206 else 49.9
                except ValueError:
                    continue

                for col in cpo_wid_df.columns[1:]:
                    if pd.notna(row[col]) and isinstance(row[col], (int, float)):
                        cpo_wid_df.at[index, col] = row[col] * factor

            # Spalte DELLAT2 hinzufügen
            cpo_wid_df['DELLAT2'] = cpo_wid_df['DELLAT'].copy()

            # Spalte 'FO' dynamisch erstellen basierend auf vorhandenen FO-Spalten
            fo_columns = [col for col in cpo_wid_df.columns if col.startswith('FO-')]
            if fo_columns:
                cpo_wid_df['FO'] = cpo_wid_df[fo_columns].sum(axis=1)
            else:
                cpo_wid_df['FO'] = 0

            # Spalte 'JD' nur erstellen, wenn entsprechende Spalten existieren
            jd_columns = [col for col in cpo_wid_df.columns if col.startswith('JD-')]
            if jd_columns:
                cpo_wid_df['JD'] = cpo_wid_df[jd_columns].sum(axis=1)
            else:
                cpo_wid_df['JD'] = 0

            return cpo_wid_df.sort_values(by='DELLAT')

        except Exception as e:
            print(f"❌ Fehler bei der Erstellung der gefilterten Zusammenfassung für 3783: {e}")
            return pd.DataFrame()

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

        # Kopie erstellen und beide Spalten in numerische Werte umwandeln
        df_copy = df.copy()
        try:
            df_copy['RKMDAT'] = pd.to_numeric(df_copy['RKMDAT'], errors='coerce')
            df_copy['Amount'] = pd.to_numeric(df_copy['Amount'], errors='coerce')
            print(f"🔢 Amount-Datentyp nach Konvertierung: {df_copy['Amount'].dtype}")
        except Exception as e:
            print(f"⚠️ Warnung beim Konvertieren: {e}")

        # Filtere Daten für Kampagne = FO-SCL und RKMDAT > 202206
        filtered_df = df_copy[
            (df_copy['Kampagne'] == special_customer) &
            (df_copy['RKMDAT'] > rkmdat_threshold)
            ]

        print(f"📊 UFC-NZG: {len(filtered_df)} Datensätze nach Filterung gefunden")

        # Erstelle statische Indexliste für alle Monate
        all_months = sorted(filtered_df['RKMDAT'].unique())
        if not all_months:
            print("⚠️ Keine Monate in den Daten gefunden")
            return pd.DataFrame(index=[], columns=[f"{special_customer}-{amount}" for amount in amounts])

        print(f"🗓️ Gefundene Monate: {all_months}")

        # Manuelles Gruppieren und Zählen
        result_data = {}
        for month in all_months:
            result_data[month] = {}
            for amount in amounts:
                # Zähle Einträge für diese Kombination
                count = len(filtered_df[(filtered_df['RKMDAT'] == month) &
                                        (filtered_df['Amount'] == amount)])
                result_data[month][f"{special_customer}-{amount}"] = count * ufc_values[amount]
                print(f"📊 Monat {month}, Amount {amount}: {count} Einträge")

        # In DataFrame umwandeln
        result_df = pd.DataFrame.from_dict(result_data, orient='index')

        # Stelle sicher, dass alle Spalten vorhanden sind
        for amount in amounts:
            col_name = f"{special_customer}-{amount}"
            if col_name not in result_df.columns:
                result_df[col_name] = 0

        # Fülle NaN-Werte mit 0
        result_df = result_df.fillna(0)

        print(f"✅ UFC-NZG: DataFrame mit {len(result_df)} Zeilen und {len(result_df.columns)} Spalten erstellt")

        return result_df

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

        # Kopie erstellen und numerische Umwandlung
        df_copy = df.copy()
        try:
            df_copy['DELLAT'] = pd.to_numeric(df_copy['DELLAT'], errors='coerce')
            df_copy['Amount'] = pd.to_numeric(df_copy['Amount'], errors='coerce')
            df_copy['Deletion Type'] = pd.to_numeric(df_copy['Deletion Type'], errors='coerce')
        except Exception as e:
            print(f"⚠️ Warnung beim Konvertieren: {e}")

        # Filtere relevante Daten - keine Produktcode-Filterung mehr!
        filtered_df = df_copy[
            (df_copy['Kampagne'] == special_customer) &
            (df_copy['Deletion Type'].isin([1, 2, 5])) &
            (df_copy['DELLAT'].notna())
            ]

        print(f"📊 UFC-WID: {len(filtered_df)} Datensätze nach Filterung gefunden")

        # Manuelles Erstellen der Ergebnistabelle
        all_months = sorted(filtered_df['DELLAT'].unique())
        if not all_months:
            print("⚠️ Keine DELLAT-Werte in den Daten gefunden")
            return pd.DataFrame(index=[], columns=[f"{special_customer}-{amount}" for amount in amounts])

        print(f"🗓️ Gefundene DELLAT-Werte: {all_months}")

        # Manuelles Gruppieren und Zählen
        result_data = {}
        for month in all_months:
            result_data[month] = {}
            for amount in amounts:
                # Zähle Einträge für diese Kombination
                count = len(filtered_df[(filtered_df['DELLAT'] == month) &
                                        (filtered_df['Amount'] == amount)])
                result_data[month][f"{special_customer}-{amount}"] = count * ufc_values[amount]

        # In DataFrame umwandeln
        result_df = pd.DataFrame.from_dict(result_data, orient='index')

        # Stelle sicher, dass alle Spalten vorhanden sind
        for amount in amounts:
            col_name = f"{special_customer}-{amount}"
            if col_name not in result_df.columns:
                result_df[col_name] = 0

        # Fülle NaN-Werte mit 0
        result_df = result_df.fillna(0)

        print(f"✅ UFC-WID: DataFrame mit {len(result_df)} Zeilen und {len(result_df.columns)} Spalten erstellt")

        return result_df

    def create_ufc_kün(self, df):
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

        # Kopie erstellen und numerische Umwandlung
        df_copy = df.copy()
        try:
            df_copy['RKMDAT'] = pd.to_numeric(df_copy['RKMDAT'], errors='coerce')
            df_copy['DELLAT'] = pd.to_numeric(df_copy['DELLAT'], errors='coerce')
            df_copy['Amount'] = pd.to_numeric(df_copy['Amount'], errors='coerce')
            df_copy['Deletion Type'] = pd.to_numeric(df_copy['Deletion Type'], errors='coerce')
            print(
                f"🔢 Datentypen nach Konvertierung: DELLAT={df_copy['DELLAT'].dtype}, Amount={df_copy['Amount'].dtype}")
        except Exception as e:
            print(f"⚠️ Warnung beim Konvertieren: {e}")

        # Filtere relevante Daten
        filtered_df = df_copy[
            (df_copy['Kampagne'] == special_customer) &
            (df_copy['Deletion Type'].isin([3, 4, 6, 7])) &
            (df_copy['RKMDAT'] > rkmdat_threshold) &
            (df_copy['DELLAT'].notna())
            ].copy()

        print(f"📊 UFC-KÜN: {len(filtered_df)} Datensätze nach Filterung gefunden")

        # Wenn keine Daten gefunden wurden, leeren DataFrame zurückgeben
        if filtered_df.empty:
            print("⚠️ Keine UFC-KÜN-Daten gefunden")
            return pd.DataFrame(index=[], columns=[f"{special_customer}-{amount}" for amount in amounts])

        # Konvertiere die relevanten Spalten in Datum
        try:
            filtered_df['FirstDueDate'] = pd.to_datetime(filtered_df['FirstDueDate'], errors='coerce')
            filtered_df['Last Due Date'] = pd.to_datetime(filtered_df['Last Due Date'], errors='coerce')
            print(f"📅 FirstDueDate und Last Due Date in Datumsformat konvertiert")
        except Exception as e:
            print(f"⚠️ Warnung beim Konvertieren von Datumsfeldern: {e}")
            return pd.DataFrame(index=[], columns=[f"{special_customer}-{amount}" for amount in amounts])

        # Hilfsfunktion: Berechnung der Differenz in vollen Monaten (DATEDIF + 1)
        def calculate_month_difference(first_due_date, last_due_date):
            if pd.isna(first_due_date) or pd.isna(last_due_date):
                return 0

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
                return 0

        # Berechne die Monate mithilfe der Funktion
        filtered_df['Months Difference'] = filtered_df.apply(
            lambda row: calculate_month_difference(row['FirstDueDate'], row['Last Due Date']),
            axis=1
        )

        # Filtere nur Datensätze, bei denen die Differenz < 24 ist
        filtered_df = filtered_df[filtered_df['Months Difference'] < 24]
        print(f"📊 UFC-KÜN: {len(filtered_df)} Datensätze nach Monatsdifferenzfilterung")

        # Wenn keine passenden Daten gefunden wurden, leeren DataFrame zurückgeben
        if filtered_df.empty:
            print("⚠️ Keine UFC-KÜN-Daten mit gültiger Monatsdifferenz gefunden")
            return pd.DataFrame(index=[], columns=[f"{special_customer}-{amount}" for amount in amounts])

        # Alle Monate identifizieren
        all_months = sorted(filtered_df['DELLAT'].unique())
        print(f"🗓️ Gefundene DELLAT-Monate für KÜN: {all_months}")

        # Manuelles Gruppieren und Werte berechnen
        result_data = {}
        for month in all_months:
            result_data[month] = {}
            for amount in amounts:
                # Filtere Datensätze für diesen Monat und Betrag
                month_amount_df = filtered_df[(filtered_df['DELLAT'] == month) & (filtered_df['Amount'] == amount)]

                total_value = 0
                for _, row in month_amount_df.iterrows():
                    # Berechnung des Wertes basierend auf der Monatsdifferenz
                    months = row['Months Difference']
                    ufc_value = ufc_values.get(amount, 0)
                    value = (1 - (months / 24)) * ufc_value
                    total_value += value

                result_data[month][f"{special_customer}-{amount}"] = total_value

        # In DataFrame umwandeln
        result_df = pd.DataFrame.from_dict(result_data, orient='index')

        # Stelle sicher, dass alle Spalten vorhanden sind
        for amount in amounts:
            col_name = f"{special_customer}-{amount}"
            if col_name not in result_df.columns:
                result_df[col_name] = 0

        # Fülle NaN-Werte mit 0
        result_df = result_df.fillna(0)

        print(f"✅ UFC-KÜN: DataFrame mit {len(result_df)} Zeilen und {len(result_df.columns)} Spalten erstellt")

        return result_df

    def create_result_v2(self, cpo_nzg_df, cpo_wid_df, filtered_summary_df_5, ufc_nzg_df, ufc_wid_df, ufc_kün_df,
                         cpo_nzg_df_3783=None, cpo_wid_df_3783=None):
        """
        Erstellt das Gesamtübersicht-Sheet 'Result V2' mit robuster Datumsverarbeitung
        """

        # Sichere Extraktion von Datumswerten
        def get_dates(df, column_name=None, is_index=False):
            dates = []
            if df is None or df.empty:
                return dates

            try:
                if is_index:
                    # Behandle Index-Werte
                    for idx in df.index:
                        try:
                            if pd.notna(idx):
                                dates.append(int(idx))
                        except (ValueError, TypeError):
                            pass
                else:
                    # Behandle Spaltenwerte
                    if column_name in df.columns:
                        for val in df[column_name]:
                            try:
                                if pd.notna(val):
                                    dates.append(int(val))
                            except (ValueError, TypeError):
                                pass
                return dates
            except Exception as e:
                print(f"⚠️ Fehler bei Datumsextraktion: {e}")
                return dates

        # Debugausgaben
        print(f"🔍 UFC_NZG Index: {list(ufc_nzg_df.index)[:5]}{'...' if len(ufc_nzg_df.index) > 5 else ''}")
        print(f"🔍 UFC_WID Index: {list(ufc_wid_df.index)[:5]}{'...' if len(ufc_wid_df.index) > 5 else ''}")
        print(f"🔍 UFC_KÜN Index: {list(ufc_kün_df.index)[:5]}{'...' if len(ufc_kün_df.index) > 5 else ''}")

        # Sammle alle Datumsangaben
        all_dates = []

        # CPO Datumsangaben hinzufügen
        all_dates.extend(get_dates(cpo_nzg_df, 'RKMDAT'))
        all_dates.extend(get_dates(cpo_wid_df, 'DELLAT'))
        all_dates.extend(get_dates(filtered_summary_df_5, 'DELLAT'))

        # UFC Datumsangaben hinzufügen (Index)
        all_dates.extend(get_dates(ufc_nzg_df, is_index=True))
        all_dates.extend(get_dates(ufc_wid_df, is_index=True))
        all_dates.extend(get_dates(ufc_kün_df, is_index=True))


        # Falls 3783 Daten vorhanden sind, füge deren Datumsangaben hinzu
        if cpo_nzg_df_3783 is not None and not cpo_nzg_df_3783.empty:
            all_dates = sorted(set(all_dates).union(cpo_nzg_df_3783['RKMDAT']))
        if cpo_wid_df_3783 is not None and not cpo_wid_df_3783.empty:
            all_dates = sorted(set(all_dates).union(cpo_wid_df_3783['DELLAT']))

        # Initialisiere eine Ergebnis-Datenstruktur
        result_data = []

        for date in all_dates:
            # Initialisiere eine leere Zeile für das Datum
            row = {'Datum': date}

            # FO-NZG-CPO (Produkt 3402) (Werte der Spalten FO-ITM, FO-OTM, FO-S2S summieren)
            nzg_cpo = cpo_nzg_df.loc[
                (cpo_nzg_df['RKMDAT'] == date),
                ['FO-ITM', 'FO-OTM', 'FO-S2S', 'FO-SCL - 19.9', 'FO-SCL - 9.9', 'FO-SCL - 14.9', 'FO-SCL - 29.9',
                 'FO-SCL - 24.9']
            ].sum().sum() if date in get_dates(cpo_nzg_df, 'RKMDAT') else 0
            row['FO-NZG-CPO'] = nzg_cpo


            # FO-CPO-Widerruf (Produkt 3402): Summiere alle positiven Werte aus CPO_WID und setze positives Vorzeichen auf negativ
            fo_cpo_wid = cpo_wid_df.loc[
                (cpo_wid_df['DELLAT'] == date),
                'FO'
            ].sum() if date in cpo_wid_df['DELLAT'].values else 0
            row['FO-CPO-Widerruf'] = fo_cpo_wid

            row['RE Call Center'] = ''
            row['IST - SOLL'] = ''

            # FO-NZG-CPO (Produkt 3783)
            # FO-NZG-CPO (Produkt 3783) - robust mit verfügbaren Spalten umgehen
            nzg_cpo_3783 = 0
            if cpo_nzg_df_3783 is not None and not cpo_nzg_df_3783.empty:
                # Nur vorhandene FO-Spalten verwenden
                fo_columns = [col for col in cpo_nzg_df_3783.columns if col.startswith('FO-')]
                if fo_columns and date in cpo_nzg_df_3783['RKMDAT'].values:
                    nzg_cpo_3783 = cpo_nzg_df_3783.loc[cpo_nzg_df_3783['RKMDAT'] == date, fo_columns].sum().sum()
            row['FO-NZG-CPO 3783'] = nzg_cpo_3783

            # FO-CPO-Widerruf (Produkt 3783) - robust prüfen ob 'FO' vorhanden ist
            fo_cpo_wid_3783 = 0
            if cpo_wid_df_3783 is not None and not cpo_wid_df_3783.empty:
                if 'FO' in cpo_wid_df_3783.columns and date in cpo_wid_df_3783['DELLAT'].values:
                    fo_cpo_wid_3783 = cpo_wid_df_3783.loc[cpo_wid_df_3783['DELLAT'] == date, 'FO'].sum()
            row['FO-CPO-Widerruf 3783'] = fo_cpo_wid_3783

            row['RE Call Center 3783'] = ''
            row['IST - SOLL 3783'] = ''

            # JD-NZG-CPO (Werte der Spalten JD-OTM und JD-ITM summieren)
            jd_nzg_cpo = cpo_nzg_df.loc[
                (cpo_nzg_df['RKMDAT'] == date),
                ['JD-OTM', 'JD-ITM']
            ].sum().sum() if date in cpo_nzg_df['RKMDAT'].values else 0
            row['JD-NZG-CPO'] = jd_nzg_cpo

            # JD-CPO-Widerruf: Summiere alle positiven Werte aus CPO_WID und setze positives Vorzeichen auf negativ
            jd_cpo_wid = cpo_wid_df.loc[
                (cpo_wid_df['DELLAT'] == date),
                'JD'
            ].sum() if date in cpo_wid_df['DELLAT'].values else 0
            row['JD-CPO-Widerruf'] = jd_cpo_wid

            # F und G werden leer initialisiert
            row['RE Call Center JD'] = ''
            row['IST - SOLL JD'] = ''

            # FO-NZG-UFC (Werte aus UFC_NZG summieren für das Datum)
            ufc_nzg = ufc_nzg_df.loc[date].sum().sum() if date in ufc_nzg_df.index else 0
            row['FO-NZG-UFC'] = ufc_nzg

            # FO-Widerruf-UFC (Werte aus UFC_WID summieren und negativ machen)
            ufc_wid = ufc_wid_df.loc[date].sum().sum() * -1 if date in ufc_wid_df.index else 0
            row['FO-Widerruf-UFC'] = ufc_wid

            # FO-CB-UFC (Werte aus UFC_KÜN summieren und negativ machen)
            ufc_cb = ufc_kün_df.loc[date].sum().sum() * -1 if date in ufc_kün_df.index else 0
            row['FO-CB-UFC'] = ufc_cb

            row['RE Call Center UFC'] = ''
            row['IST - SOLL UFC'] = ''

            ts_nzg_cpo = cpo_nzg_df.loc[
                (cpo_nzg_df['RKMDAT'] == date),
                ['TS-OTM']
            ].sum().sum() if date in cpo_nzg_df['RKMDAT'].values else 0
            row['TS-NZG-CPO'] = ts_nzg_cpo

            ts_cpo_wid = cpo_wid_df.loc[
                (cpo_wid_df['DELLAT'] == date),
                'TS-OTM'
            ].sum() if date in cpo_wid_df['DELLAT'].values else 0
            row['TS-CPO-Widerruf'] = ts_cpo_wid

            row['RE Call Center TS'] = ''
            row['IST - SOLL TS'] = ''

            m3_nzg_cpo = cpo_nzg_df.loc[
                (cpo_nzg_df['RKMDAT'] == date),
                ['M3-OTM']
            ].sum().sum() if date in cpo_nzg_df['RKMDAT'].values else 0
            row['M3-NZG-CPO'] = m3_nzg_cpo

            m3_cpo_wid = cpo_wid_df.loc[
                (cpo_wid_df['DELLAT'] == date),
                'M3-OTM'
            ].sum() if date in cpo_wid_df['DELLAT'].values else 0
            row['M3-CPO-Widerruf'] = m3_cpo_wid

            row['RE Call Center M3'] = ''
            row['IST - SOLL M3'] = ''

            # Zeile zur Ergebnisliste hinzufügen
            result_data.append(row)

        # Konvertiere die Ergebnisliste in einen DataFrame
        result_v2_df = pd.DataFrame(result_data).fillna(0)

        # Sortiere nach Datum
        result_v2_df = result_v2_df.sort_values(by='Datum')

        return result_v2_df

# Hauptprogramm
if __name__ == '__main__':
    excel_file_path = r"/Users/adrianstotzler/Desktop/Training.xlsx"
    exporter = ExcelExporterWithSummary(excel_file_path)
    exporter.export_table_with_summary()