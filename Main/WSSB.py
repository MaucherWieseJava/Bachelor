# =====================================================================
# AUTOR: @Adrian Stötzler
# WSSB Abrechnung
# =====================================================================

import pandas as pd
import os

class ExcelExporterWithSummary:
    def __init__(self, excel_file_path):
        self.excel_file_path = excel_file_path

    def export_table_with_summary(self):
        try:
            desktop_path = os.path.join(os.environ["HOME"], "Desktop")
            output_file = os.path.join(desktop_path, "WSSB_Abrechnung.xlsx")
            print("🔄 Lade Excel-Datei...")
            df = pd.read_excel(self.excel_file_path, engine='openpyxl')

            required_columns = ['ID', 'Kampagne', 'Deletion Type', 'RKMDAT', 'DELLAT', 'Amount']
            for col in required_columns:
                if col not in df.columns:
                    print(f"❌ Fehler: Die Spalte '{col}' fehlt in der Excel-Datei.")
                    return
            
            campaigns = df['Kampagne'].unique()
            rkmdat_values = df['RKMDAT'].unique()
            dellat_values = df['DELLAT'].unique()
            amounts = df['Amount'].unique()

            print("📊 Erstelle Zusammenfassungstabelle für Neuzugänge...")
            neuzugaenge_df = self.create_summary_table(df, 'RKMDAT', campaigns)

            print("📊 Erstelle Zusammenfassungstabelle für Widerrufe...")
            widerrufe_df = self.create_filtered_summary_table(df, 'DELLAT', [1, 2, 5], campaigns)

            print("📊 Erstelle Zusammenfassungstabelle für Kündigungen...")
            kuendigungen_df = self.create_filtered_summary_table(df, 'DELLAT', [3, 4, 6], campaigns)
            
            print("📊 Erstelle CPO_NZG Tabelle...")
            cpo_nzg_df = self.create_cpo_nzg(df, neuzugaenge_df, campaigns)
            
            print("📊 Erstelle CPO_WID Tabelle...")
            cpo_wid_df = self.create_cpo_wid(df, widerrufe_df, campaigns)
            
            print("📊 Erstelle CPO_KUN Tabelle...")
            cpo_kun_df = self.create_cpo_kun(df, campaigns)
            
            print("📊 Erstelle Result V2 Tabelle...")
            result_v2_df = self.create_result_v2(cpo_nzg_df, cpo_wid_df, cpo_kun_df, campaigns)

            # Speichere die Daten in die Excel-Datei
            print("💾 Speichere Daten in die Excel-Datei...")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Sheet1: Alle Daten
                df.to_excel(writer, index=False, sheet_name="Gesamtdaten")
                
                # Sheet2: Neuzugänge
                neuzugaenge_df.to_excel(writer, index=False, sheet_name="Neuzugänge")
                
                # Sheet3: Widerrufe
                widerrufe_df.to_excel(writer, index=False, sheet_name="Widerrufe")
                
                # Sheet4: Kündigungen
                kuendigungen_df.to_excel(writer, index=False, sheet_name="Kündigungen")
                
                # Sheet5: CPO_NZG
                cpo_nzg_df.to_excel(writer, index=False, sheet_name="CPO_NZG")
                
                # Sheet6: CPO_WID
                cpo_wid_df.to_excel(writer, index=False, sheet_name="CPO_WID")
                
                # Sheet7: CPO_KUN
                cpo_kun_df.to_excel(writer, index=False, sheet_name="CPO_KUN")
                
                # Sheet8: Result V2
                result_v2_df.to_excel(writer, index=False, sheet_name="Result V2")

            print(f"✅ Export erfolgreich! Datei gespeichert unter: {output_file}")

        except Exception as e:
            print(f"❌ Fehler beim Export: {e}")
            import traceback
            traceback.print_exc()

    def create_summary_table(self, df, date_column, campaigns):
        """
        Erstellt eine Zusammenfassungstabelle mit Anzahl der Einträge pro Datum und Kampagne
        """
        # Erstelle einen leeren DataFrame mit den benötigten Spalten
        columns = [date_column] + list(campaigns)
        summary_df = pd.DataFrame(columns=columns)
        
        # Extrahiere alle eindeutigen Datumsangaben
        unique_dates = sorted(df[date_column].dropna().unique())
        
        # Für jedes Datum zähle die Einträge pro Kampagne
        for date in unique_dates:
            row = {date_column: date}
            for campaign in campaigns:
                count = len(df[(df[date_column] == date) & (df['Kampagne'] == campaign)])
                row[campaign] = count
            summary_df = pd.concat([summary_df, pd.DataFrame([row])], ignore_index=True)
        
        # Konvertiere Datumsangaben wenn möglich in numerische Werte für bessere Sortierung
        try:
            summary_df[date_column] = pd.to_numeric(summary_df[date_column], errors='raise')
        except:
            # Wenn die Konvertierung fehlschlägt, als String belassen
            summary_df[date_column] = summary_df[date_column].astype(str)
        
        return summary_df.sort_values(by=date_column, ascending=True)

    def create_filtered_summary_table(self, df, date_column, deletion_types, campaigns):
        """
        Erstellt eine gefilterte Zusammenfassungstabelle basierend auf Deletion Types
        """
        # Filtere den DataFrame nach den angegebenen Deletion Types
        filtered_df = df[df['Deletion Type'].isin(deletion_types)]
        
        # Erstelle die Zusammenfassungstabelle aus dem gefilterten DataFrame
        return self.create_summary_table(filtered_df, date_column, campaigns)
        
    def create_cpo_nzg(self, df, neuzugaenge_df, campaigns):
        """
        Erstellt eine Tabelle mit den CPO_NZG Werten:
        27,90 + 53% von drei Jahresprämien für jeden Vertrag
        """
        # Kopie der Neuzugänge-Tabelle erstellen mit Struktur
        cpo_nzg_df = pd.DataFrame(columns=['RKMDAT'] + list(campaigns))
        
        # Für jede Zeile in der Neuzugänge-Tabelle
        for _, row in neuzugaenge_df.iterrows():
            rkmdat = row['RKMDAT']
            new_row = {'RKMDAT': rkmdat}
            
            # Für jede Kampagne
            for campaign in campaigns:
                # Anzahl der Verträge für diese Kombination
                if campaign in row:
                    contract_count = row[campaign]
                    if contract_count > 0:
                        # Finde alle Verträge mit diesem RKMDAT und dieser Kampagne
                        contracts = df[(df['RKMDAT'] == rkmdat) & (df['Kampagne'] == campaign)]
                        
                        # Berechne den Gesamtwert für alle Verträge dieser Kombination
                        total_value = 0
                        for _, contract in contracts.iterrows():
                            amount = contract['Amount']
                            # Berechnung: 27.90 + 53% von drei Jahresprämien
                            value = 27.90 + (0.53 * (3 * amount))
                            total_value += value
                            
                        new_row[campaign] = total_value
                    else:
                        new_row[campaign] = 0
                else:
                    new_row[campaign] = 0
                    
            # Füge die Zeile zum Ergebnis-DataFrame hinzu
            cpo_nzg_df = pd.concat([cpo_nzg_df, pd.DataFrame([new_row])], ignore_index=True)
        
        # Sortiere nach RKMDAT
        try:
            cpo_nzg_df['RKMDAT'] = pd.to_numeric(cpo_nzg_df['RKMDAT'], errors='raise')
        except:
            cpo_nzg_df['RKMDAT'] = cpo_nzg_df['RKMDAT'].astype(str)
            
        return cpo_nzg_df.sort_values(by='RKMDAT', ascending=True)
    
    def create_cpo_wid(self, df, widerrufe_df, campaigns):
        """
        Erstellt eine Tabelle mit den CPO_WID Werten für Verträge mit Deletion Type 1, 2, 5:
        27,90 + 53% von drei Jahresprämien für jeden Vertrag (negativ)
        """
        # Erstelle einen leeren DataFrame für das Ergebnis
        cpo_wid_df = pd.DataFrame(columns=['DELLAT'] + list(campaigns))
        
        # Filtere nur die Widerrufe (Deletion Type 1, 2, 5)
        widerruf_df = df[df['Deletion Type'].isin([1, 2, 5])]
        
        # Für jede Zeile in der Widerrufe-Tabelle
        for _, row in widerrufe_df.iterrows():
            dellat = row['DELLAT']
            
            # Überprüfen, ob diese DELLAT bereits im Ergebnis-DataFrame vorhanden ist
            existing_row = cpo_wid_df[cpo_wid_df['DELLAT'] == dellat]
            
            if len(existing_row) == 0:
                # Wenn nicht, erstelle eine neue Zeile
                new_row = {'DELLAT': dellat}
                for campaign in campaigns:
                    new_row[campaign] = 0
                cpo_wid_df = pd.concat([cpo_wid_df, pd.DataFrame([new_row])], ignore_index=True)
        
        # Für jede DELLAT und Kampagne die Werte berechnen
        for dellat in cpo_wid_df['DELLAT'].unique():
            for campaign in campaigns:
                # Finde alle Widerrufe mit dieser DELLAT und dieser Kampagne
                contracts = widerruf_df[(widerruf_df['DELLAT'] == dellat) & 
                                        (widerruf_df['Kampagne'] == campaign)]
                
                # Berechne den Gesamtwert für alle Widerrufe
                total_value = 0
                for _, contract in contracts.iterrows():
                    amount = contract['Amount']
                    # Berechnung: 27.90 + 53% von drei Jahresprämien
                    value = 27.90 + (0.53 * (3 * amount))
                    total_value += value
                
                # Widerrufswerte müssen negativ sein (Geld wird zurückgezahlt)
                if total_value > 0:
                    total_value = -total_value
                    
                # Aktualisiere den Wert im Ergebnis-DataFrame
                cpo_wid_df.loc[cpo_wid_df['DELLAT'] == dellat, campaign] = total_value
        
        # Sortiere nach DELLAT
        try:
            cpo_wid_df['DELLAT'] = pd.to_numeric(cpo_wid_df['DELLAT'], errors='raise')
        except:
            cpo_wid_df['DELLAT'] = cpo_wid_df['DELLAT'].astype(str)
        
        return cpo_wid_df.sort_values(by='DELLAT', ascending=True)
    
    def calculate_month_difference(self, rkmdat, dellat):
        """
        Berechnet den Unterschied in Monaten zwischen RKMDAT und DELLAT
        Format: YYYYMM (z.B. 202105 für Mai 2021)
        """
        try:
            # Konvertiere zu Integers
            rkmdat_int = int(rkmdat)
            dellat_int = int(dellat)
            
            # Extrahiere Jahr und Monat
            rkmdat_year = rkmdat_int // 100
            rkmdat_month = rkmdat_int % 100
            
            dellat_year = dellat_int // 100
            dellat_month = dellat_int % 100
            
            # Berechne die Differenz in Monaten
            month_diff = (dellat_year - rkmdat_year) * 12 + (dellat_month - rkmdat_month)
            
            return month_diff
        except:
            return -1  # Fehlerfall, ungültige Werte
            
    def create_cpo_kun(self, df, campaigns):
        """
        Erstellt eine Tabelle mit den CPO_KUN Werten für Verträge mit Deletion Type 3, 4, 6:
        Berechnung basierend auf der Laufzeit zwischen RKMDAT und DELLAT
        """
        # Filtere nur die Kündigungen (Deletion Type 3, 4, 6)
        kuendigungen_df = df[df['Deletion Type'].isin([3, 4, 6])].copy()
        
        # Erstelle einen leeren DataFrame für das Ergebnis
        cpo_kun_df = pd.DataFrame(columns=['DELLAT'] + list(campaigns))
        
        # Sammle alle DELLAT-Werte für Kündigungen
        all_dellat = sorted(kuendigungen_df['DELLAT'].unique())
        
        # Initialisiere den Ergebnis-DataFrame mit allen DELLAT-Werten
        for dellat in all_dellat:
            new_row = {'DELLAT': dellat}
            for campaign in campaigns:
                new_row[campaign] = 0
            cpo_kun_df = pd.concat([cpo_kun_df, pd.DataFrame([new_row])], ignore_index=True)
        
        # Für jeden Vertrag mit Kündigung
        for _, contract in kuendigungen_df.iterrows():
            rkmdat = contract['RKMDAT']
            dellat = contract['DELLAT']
            campaign = contract['Kampagne']
            amount = contract['Amount']
            
            # Berechne den Monatsunterschied
            month_diff = self.calculate_month_difference(rkmdat, dellat)
            
            # Ignoriere Fälle, wo DELLAT < RKMDAT oder month_diff > 36
            if month_diff < 1 or month_diff > 36:
                continue
                
            # Berechne den Gesamtbetrag der Provision
            total_provision = 27.90 + (0.53 * (3 * amount))
            
            # Lineare Rückzahlung basierend auf der Laufzeit
            # Je länger der Vertrag lief, desto weniger muss zurückgezahlt werden
            remaining_months = 36 - month_diff
            refund_percentage = remaining_months / 36
            refund_amount = total_provision * refund_percentage
            
            # Werte müssen negativ sein (Geld wird zurückgezahlt)
            refund_amount = -refund_amount
            
            # Füge den Wert zum entsprechenden DELLAT und Kampagne hinzu
            idx = cpo_kun_df[cpo_kun_df['DELLAT'] == dellat].index
            if len(idx) > 0:
                current_value = cpo_kun_df.at[idx[0], campaign]
                cpo_kun_df.at[idx[0], campaign] = current_value + refund_amount
        
        # Sortiere nach DELLAT
        try:
            cpo_kun_df['DELLAT'] = pd.to_numeric(cpo_kun_df['DELLAT'], errors='raise')
        except:
            cpo_kun_df['DELLAT'] = cpo_kun_df['DELLAT'].astype(str)
        
        return cpo_kun_df.sort_values(by='DELLAT', ascending=True)
    
    def create_result_v2(self, cpo_nzg_df, cpo_wid_df, cpo_kun_df, campaigns):
        """
        Erstellt eine Zusammenfassungstabelle (Result V2) mit Werten nach Call Center gruppiert
        """
        # Identifiziere alle Call Center (erste 2 Buchstaben jeder Kampagne)
        # Sicherstellen, dass alle Kampagnen Strings sind
        call_centers = set()
        for campaign in campaigns:
            try:
                # Konvertiere zu String falls nötig
                campaign_str = str(campaign)
                # Prüfe ob der String lang genug ist
                if len(campaign_str) >= 2:
                    call_centers.add(campaign_str[:2])
                else:
                    print(f"⚠️ Warnung: Kampagne '{campaign_str}' ist zu kurz für Präfix-Extraktion")
            except Exception as e:
                print(f"⚠️ Fehler bei der Verarbeitung von Kampagne {campaign}: {e}")
        
        call_centers = sorted(call_centers)
        print(f"🏢 Erkannte Call Center: {', '.join(call_centers)}")
        
        # Sammle alle Datumsangaben (RKMDAT und DELLAT)
        all_dates = set()
        
        if not cpo_nzg_df.empty:
            all_dates.update(cpo_nzg_df['RKMDAT'])
        
        if not cpo_wid_df.empty:
            all_dates.update(cpo_wid_df['DELLAT'])
            
        if not cpo_kun_df.empty:
            all_dates.update(cpo_kun_df['DELLAT'])
            
        all_dates = sorted(all_dates)
        
        # Definiere die Spalten für den Result V2 DataFrame - FIX: Mache Spaltennamen einzigartig
        columns = ['Datum']
        for cc in call_centers:
            columns.extend([f"{cc}-NZG", f"{cc}-WID", f"{cc}-KUN", f"RE Call Center {cc}", f"IST - SOLL {cc}"])
        
        # Erstelle einen leeren DataFrame für das Ergebnis
        result_v2_df = pd.DataFrame(columns=columns)
        
        # Für jedes Datum eine Zeile erstellen
        for date in all_dates:
            row = {'Datum': date}
            
            # Für jedes Call Center
            for cc in call_centers:
                # Finde alle Kampagnen, die zu diesem Call Center gehören
                cc_campaigns = []
                for campaign in campaigns:
                    campaign_str = str(campaign)
                    if len(campaign_str) >= 2 and campaign_str.startswith(cc):
                        cc_campaigns.append(campaign)
                
                # Berechne NZG (Neuzugänge) für dieses Call Center
                nzg_value = 0
                if not cpo_nzg_df.empty:
                    nzg_rows = cpo_nzg_df[cpo_nzg_df['RKMDAT'] == date]
                    if not nzg_rows.empty:
                        for campaign in cc_campaigns:
                            if campaign in nzg_rows.columns:
                                nzg_value += nzg_rows[campaign].sum()
                row[f"{cc}-NZG"] = nzg_value
                
                # Berechne WID (Widerrufe) für dieses Call Center
                wid_value = 0
                if not cpo_wid_df.empty:
                    wid_rows = cpo_wid_df[cpo_wid_df['DELLAT'] == date]
                    if not wid_rows.empty:
                        for campaign in cc_campaigns:
                            if campaign in wid_rows.columns:
                                wid_value += wid_rows[campaign].sum()
                row[f"{cc}-WID"] = wid_value
                
                # Berechne KUN (Kündigungen) für dieses Call Center
                kun_value = 0
                if not cpo_kun_df.empty:
                    kun_rows = cpo_kun_df[cpo_kun_df['DELLAT'] == date]
                    if not kun_rows.empty:
                        for campaign in cc_campaigns:
                            if campaign in kun_rows.columns:
                                kun_value += kun_rows[campaign].sum()
                row[f"{cc}-KUN"] = kun_value
                
                # FIX: Verwende eindeutige Spaltennamen für jedes Call Center
                row[f"RE Call Center {cc}"] = ""
                row[f"IST - SOLL {cc}"] = ""
            
            # Füge die Zeile zum Ergebnis-DataFrame hinzu
            result_v2_df = pd.concat([result_v2_df, pd.DataFrame([row])], ignore_index=True)
        
        # Sortiere nach Datum
        try:
            result_v2_df['Datum'] = pd.to_numeric(result_v2_df['Datum'], errors='raise')
        except:
            result_v2_df['Datum'] = result_v2_df['Datum'].astype(str)
        
        return result_v2_df.sort_values(by='Datum', ascending=True)

# Hauptprogramm
if __name__ == '__main__':
    # Direkte Pfadangabe zur Eingabedatei
    excel_file_path = r"/Users/adrianstotzler/Desktop/WSSB.xlsx"  # Hier können Sie den gewünschten Pfad eingeben
    exporter = ExcelExporterWithSummary(excel_file_path)
    exporter.export_table_with_summary()

