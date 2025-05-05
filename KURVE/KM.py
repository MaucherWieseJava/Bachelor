import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from lifelines import KaplanMeierFitter
from pathlib import Path


class KaplanMeierAnalysisTool:
    """
    Klasse zur Erstellung von Kaplan-Meier-Kurven und zur Analyse der Deletion-Types
    """

    def __init__(self, file_path=None):
        """
        Initialisiert das Tool mit dem angegebenen Dateipfad

        Parameter:
            file_path (str): Pfad zur Excel-Datei mit Versicherungsdaten
        """
        self.file_path = file_path or os.path.join(os.environ["HOME"], "Desktop", "Training.xlsx")
        self.df = None
        self.active_count = 0
        self.total_count = 0
        self.active_percent = 0

        # Output-Verzeichnis erstellen
        self.output_dir = Path("output")
        self.output_dir.mkdir(exist_ok=True)

        # Konstanten
        self.MAX_DURATION = 1200  # Maximale Beobachtungsdauer in Tagen

        print("=" * 80)
        print("KAPLAN-MEIER-ANALYSE UND DELETION-TYPE-VERTEILUNG")
        print("=" * 80)

    def load_data(self):
        """
        Lädt Daten aus der Excel-Datei
        """
        try:
            print(f"\nLade Daten aus: {self.file_path}")
            self.df = pd.read_excel(self.file_path)
            print(f"Daten erfolgreich geladen: {len(self.df)} Zeilen, {len(self.df.columns)} Spalten")

            # "Deletion Type" zu Integer konvertieren
            if 'Deletion Type' in self.df.columns:
                self.df['Deletion Type'] = pd.to_numeric(self.df['Deletion Type'], errors='coerce').fillna(0).astype(
                    int)

                # NEUE AUSFÜHRLICHE ANALYSE DES DELETION TYPE
                deletion_type_zero = (self.df['Deletion Type'] == 0).sum()
                deletion_type_nonzero = (self.df['Deletion Type'] > 0).sum()
                total_records = len(self.df)

                print("\n" + "=" * 60)
                print("DETAILANALYSE DER DATENSÄTZE:")
                print(f"Gesamtanzahl Datensätze: {total_records}")
                print(
                    f"Datensätze mit Deletion Type = 0: {deletion_type_zero} ({deletion_type_zero / total_records * 100:.1f}%)")
                print(
                    f"Datensätze mit Deletion Type > 0: {deletion_type_nonzero} ({deletion_type_nonzero / total_records * 100:.1f}%)")
                print("=" * 60)

                # Verteilung der Löschungstypen anzeigen
                deletion_counts = self.df['Deletion Type'].value_counts()
                print("\nVerteilung der einzelnen Deletion Types:")
                for dt, count in deletion_counts.items():
                    print(f"  Typ {dt}: {count} Datensätze ({count / len(self.df) * 100:.1f}%)")

            return True

        except FileNotFoundError:
            print(f"Fehler: Die Datei '{self.file_path}' wurde nicht gefunden.")
            return False
        except pd.errors.EmptyDataError:
            print("Fehler: Die Datei enthält keine Daten.")
            return False
        except Exception as e:
            print(f"Ein unerwarteter Fehler ist aufgetreten: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def preprocess_data_for_survival_analysis(self):
        """
        Bereitet die Daten für die Kaplan-Meier-Analyse vor
        """
        if self.df is None:
            print("Keine Daten geladen!")
            return None

        print("\nBereite Daten für Überlebenszeitanalyse vor...")

        # Kopie erstellen
        df_survival = self.df.copy()

        # Datum-Spalten konvertieren
        date_columns = ['Start Insurance', 'Deletion allowed at', 'End Insurance']

        for col in date_columns:
            if col in df_survival.columns:
                df_survival[col] = pd.to_datetime(df_survival[col], errors='coerce')
                print(f"  ✓ '{col}' zu Datum konvertiert")
            else:
                print(f"  ⚠��� Spalte '{col}' nicht gefunden - erforderlich für Überlebenszeitanalyse")
                if col == 'Deletion allowed at':
                    print("     Versuche alternative Spalte 'End Insurance'...")
                    if 'End Insurance' in df_survival.columns:
                        df_survival['Deletion allowed at'] = pd.to_datetime(df_survival['End Insurance'],
                                                                            errors='coerce')
                        print("     'End Insurance' als Alternative verwendet")
                    else:
                        print("     Keine geeignete Alternative gefunden")
                        return None

        # Zielvariable und Überlebensdauer berechnen
        if 'Start Insurance' in df_survival.columns and 'Deletion allowed at' in df_survival.columns:
            # Event (1 = Kündigung, 0 = aktiv/zensiert) - WICHTIG: Nur Deletion Type > 0 ist ein Event!
            df_survival['event'] = (df_survival['Deletion Type'] > 0).astype(int)

            # Zählung der aktiven und kündigten Verträge
            self.total_count = len(df_survival)
            self.active_count = (df_survival['Deletion Type'] == 0).sum()
            self.canceled_count = (df_survival['Deletion Type'] > 0).sum()

            self.active_percent = (self.active_count / self.total_count) * 100
            self.canceled_percent = (self.canceled_count / self.total_count) * 100

            print(f"Gesamtanzahl Verträge: {self.total_count}")
            print(f"Aktive Verträge: {self.active_count} ({self.active_percent:.1f}%)")
            print(f"Gekündigte Verträge: {self.canceled_count} ({self.canceled_percent:.1f}%)")
            print(f"→ HINWEIS: Die Überlebenskurve sollte theoretisch nicht unter {self.active_percent:.1f}% fallen")

            # Überlebensdauer korrekt berechnen:
            # - Für aktive Verträge: Start bis End Insurance (zensiert)
            # - Für gekündigte Verträge: Start bis Deletion allowed at (Event)
            df_survival['duration'] = np.where(
                df_survival['event'] == 1,
                (df_survival['Deletion allowed at'] - df_survival['Start Insurance']).dt.days,
                (df_survival['End Insurance'] - df_survival['Start Insurance']).dt.days
            )

            # Negative oder fehlende Werte korrigieren
            df_survival['duration'] = df_survival['duration'].fillna(0).clip(lower=0)

            # Begrenzung auf MAX_DURATION (ohne Event-Status zu ändern!)
            df_survival.loc[df_survival['duration'] > self.MAX_DURATION, 'duration'] = self.MAX_DURATION

            print("\nANALYSE DER BERECHNETEN ÜBERLEBENSDAUERN:")
            print(
                f"Durchschnittliche Dauer für aktive Verträge: {df_survival.loc[df_survival['event'] == 0, 'duration'].mean():.1f} Tage")
            print(
                f"Durchschnittliche Dauer für gekündigte Verträge: {df_survival.loc[df_survival['event'] == 1, 'duration'].mean():.1f} Tage")
            print(f"Minimum Dauer: {df_survival['duration'].min():.1f} Tage")
            print(f"Maximum Dauer: {df_survival['duration'].max():.1f} Tage")
            print(f"Median Dauer: {df_survival['duration'].median():.1f} Tage")

            # Verteilung der Dauern in Intervallen
            duration_bins = [0, 30, 90, 180, 365, 730, 1200, float('inf')]
            duration_labels = ['0-30', '31-90', '91-180', '181-365', '366-730', '731-1200', '>1200']
            df_survival['duration_group'] = pd.cut(df_survival['duration'], bins=duration_bins, labels=duration_labels)
            duration_dist = df_survival['duration_group'].value_counts().sort_index()
            print("\nVerteilung der Überlebensdauern:")
            for group, count in duration_dist.items():
                print(f"  {group} Tage: {count} Datensätze ({count / len(df_survival) * 100:.1f}%)")

            # Analyse der Verträge bei 1200 Tagen
            contracts_at_max = len(df_survival[df_survival['duration'] >= self.MAX_DURATION])
            active_at_max = len(
                df_survival[(df_survival['duration'] >= self.MAX_DURATION) & (df_survival['event'] == 0)])
            print(f"  → Bei {self.MAX_DURATION} Tagen: {contracts_at_max} Verträge insgesamt")
            print(f"  → Bei {self.MAX_DURATION} Tagen: {active_at_max} aktive Verträge (Deletion Type = 0)")
            print(
                f"  → Bei {self.MAX_DURATION} Tagen: {active_at_max / contracts_at_max * 100:.1f}% der Verträge sind noch aktiv")

            # Bei Amount Spalte: Kategorien erstellen für stratifizierte Analyse
            if 'Amount' in df_survival.columns:
                try:
                    # Amount als numerisch konvertieren
                    df_survival['Amount'] = pd.to_numeric(df_survival['Amount'], errors='coerce')

                    # Quantile für Amount berechnen
                    q33 = df_survival['Amount'].quantile(0.33)
                    q66 = df_survival['Amount'].quantile(0.66)

                    # Amount-Kategorien erstellen
                    df_survival['Amount_Category'] = pd.cut(
                        df_survival['Amount'],
                        bins=[0, q33, q66, float('inf')],
                        labels=['Niedrig', 'Mittel', 'Hoch']
                    )
                    print(f"  ✓ 'Amount' kategorisiert (Niedrig: ≤{q33:.2f}, Mittel: ≤{q66:.2f}, Hoch: >{q66:.2f})")
                except Exception as e:
                    print(f"  ⚠️ Fehler bei der Kategorisierung von 'Amount': {e}")

            return df_survival
        else:
            print("Erforderliche Spalten für Überlebenszeitanalyse nicht gefunden")
            return None

    def visualize_deletion_type_pie_chart(self):
        """
        Erstellt ein Kreisdiagramm mit aktiven Verträgen vs. Widerrufen vs. anderen Kündigungen
        """
        if self.df is None or 'Deletion Type' not in self.df.columns:
            print("Keine geeigneten Daten für Deletion Type-Analyse gefunden")
            return

        print("\nErstelle Deletion Type Kreisdiagramm...")

        # Anzahl für jede Kategorie ermitteln
        active_count = (self.df['Deletion Type'] == 0).sum()
        withdrawal_count = self.df['Deletion Type'].isin([1, 2, 5]).sum()
        other_count = len(self.df) - active_count - withdrawal_count

        # Daten für das Kreisdiagramm vorbereiten
        pie_data = {
            'Aktive Verträge': active_count,
            'Widerruf': withdrawal_count
        }

        # Andere Typen nur hinzufügen, wenn vorhanden
        if other_count > 0:
            pie_data['Andere Kündigungen'] = other_count

        # Kreisdiagramm erstellen
        plt.figure(figsize=(10, 7))

        # Farbpalette für besseren Kontrast
        colors = ['#4ecdc4', '#ff6b6b', '#f9c74f']

        # Labels und Prozente vorbereiten
        labels = []
        for key, value in pie_data.items():
            percentage = value / len(self.df) * 100
            labels.append(f'{key}: {value} ({percentage:.1f}%)')

        # Kreisdiagramm zeichnen
        plt.pie(pie_data.values(), labels=labels, colors=colors, autopct='%1.1f%%',
                startangle=90, shadow=True, explode=[0.05] * len(pie_data))

        plt.axis('equal')  # Kreis statt Ellipse
        plt.title('Verteilung der Verträge', fontsize=14)
        plt.tight_layout()

        output_path = self.output_dir / 'deletion_type_distribution.png'
        plt.savefig(output_path, dpi=300)
        plt.close()

        print(f"✅ Kreisdiagramm gespeichert als '{output_path}'")

    def perform_stratified_kaplan_meier_analysis(self, stratify_column, max_categories=6):
        """
        Führt eine stratifizierte Kaplan-Meier-Analyse durch
        """
        survival_data = self.preprocess_data_for_survival_analysis()

        if survival_data is None or 'duration' not in survival_data.columns or 'event' not in survival_data.columns:
            print(f"Keine geeigneten Daten für stratifizierte Analyse nach '{stratify_column}'")
            return

        if stratify_column not in survival_data.columns:
            print(f"Spalte '{stratify_column}' nicht in Daten gefunden")
            return

        print(f"\nFühre stratifizierte Kaplan-Meier-Analyse nach '{stratify_column}' durch...")

        # WICHTIG: Erstelle eine tiefe Kopie und wende die gleiche Korrektur wie bei der Hauptkurve an
        km_data = survival_data.copy(deep=True)
        km_data.loc[km_data['event'] == 0, 'duration'] = self.MAX_DURATION

        # Bei Amount die kategorisierte Version verwenden
        if stratify_column == 'Amount' and 'Amount_Category' in km_data.columns:
            stratify_column = 'Amount_Category'
            print("  → Verwende Amount-Kategorien statt numerischer Werte")

        # Häufigsten Werte identifizieren
        value_counts = km_data[stratify_column].value_counts().head(max_categories)

        # Plot vorbereiten
        plt.figure(figsize=(12, 8))
        palette = sns.color_palette("tab10", len(value_counts))

        # Für jeden Wert eine separate Kurve zeichnen
        for i, (value, count) in enumerate(value_counts.items()):
            mask = km_data[stratify_column] == value
            if mask.sum() >= 5:  # Mindestens 5 Datenpunkte
                subgroup_data = km_data[mask]

                # Kaplan-Meier-Fitter initialisieren
                kmf = KaplanMeierFitter()

                # Anpassen und Plotten
                kmf.fit(
                    durations=subgroup_data['duration'],
                    event_observed=subgroup_data['event'],
                    label=f'{stratify_column}={value} (n={mask.sum()})'
                )
                kmf.plot(ci_show=False, color=palette[i])

                # Kündigungsrate anzeigen
                try:
                    survival_at_250 = kmf.survival_function_at_times(250).iloc[0]
                    print(f"  → {stratify_column}={value}: Bei 250 Tagen noch {survival_at_250 * 100:.1f}% aktiv")
                except Exception:
                    pass
            else:
                print(f"  ⚠️ Zu wenig Daten für {stratify_column}={value}, überspringe")

        # Formatierung und Speicherung wie bisher
        plt.axhline(y=self.active_percent / 100, color='red', linestyle='--', alpha=0.3)
        plt.title(f'Kaplan-Meier Überlebenskurven nach {stratify_column}', fontsize=14)
        plt.xlabel('Tage seit Vertragsbeginn', fontsize=12)
        plt.ylabel('Überlebenswahrscheinlichkeit', fontsize=12)
        plt.grid(True, alpha=0.3)
        plt.legend(loc='best', frameon=True, framealpha=0.8)
        plt.xlim(0, self.MAX_DURATION)
        plt.ylim(max(0.45, self.active_percent / 100 * 0.9), 1.05)
        plt.tight_layout()

        output_path = self.output_dir / f'kaplan_meier_by_{stratify_column}.png'
        plt.savefig(output_path, dpi=300)
        plt.close()

        print(f"✅ Stratifizierte Kaplan-Meier-Kurve gespeichert als '{output_path}'")

    def validate_km_computation(self, survival_data):
        """Manuelle Berechnung der Kaplan-Meier-Kurve zur Validierung"""
        print("\nMANUELLE VALIDIERUNG DER KAPLAN-MEIER-SCHÄTZUNG:")

        # Sortiere Daten nach Dauer
        sorted_data = survival_data.sort_values('duration')

        # Gruppiere nach einzigartigen Zeitpunkten
        times = sorted_data['duration'].unique()

        # Für einige repräsentative Zeitpunkte
        check_times = [90, 180, 365, 730]

        for t in check_times:
            # Fälle bis zu diesem Zeitpunkt
            cases_until_t = sorted_data[sorted_data['duration'] <= t]

            # Anzahl der Events bis zu diesem Zeitpunkt
            events_until_t = cases_until_t['event'].sum()

            # Anzahl der Fälle, die mindestens bis zu diesem Zeitpunkt beobachtet wurden
            total_at_risk = len(survival_data[survival_data['duration'] >= t])

            print(f"Zeit {t}: {events_until_t} Events, {total_at_risk} Fälle im Risiko")

    def run_analysis(self):
        """
        Führt die vollständige Analyse durch
        """
        print("\nStarte Analyse...")

        # Daten laden
        if not self.load_data():
            return False

        # Kreisdiagramm zur Verteilung der Deletion Types
        self.visualize_deletion_type_pie_chart()

        # Kaplan-Meier-Analyse durchführen
        self.perform_kaplan_meier_analysis()

        # Zusätzliche stratifizierte Analysen
        if 'Amount' in self.df.columns:
            self.perform_stratified_kaplan_meier_analysis('Amount')

        if 'Kampagne' in self.df.columns:
            self.perform_stratified_kaplan_meier_analysis('Kampagne')

        print("\n" + "=" * 80)
        print("ANALYSE ABGESCHLOSSEN")
        print("=" * 80)

        return True


def main():
    # Dateipfad kann angepasst werden oder None für Standardpfad
    file_path = None  # z.B. "C:/Daten/meine_daten.xlsx"

    # Analyse-Tool initialisieren und ausführen
    analysis_tool = KaplanMeierAnalysisTool(file_path)
    analysis_tool.run_analysis()


if __name__ == "__main__":
    main()