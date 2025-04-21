import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.ticker import PercentFormatter
import matplotlib.patches as mpatches


class CallCenterAnalyse:
    def __init__(self):
        # Kampagnen-Daten: Format "Kampagne": [Gesamtverträge, Stornierte Verträge]
        self.kampagne_data = {
            "FO-S2S": [788, 159],
            "FO-ITM": [15781, 7441],
            "FO-OTM": [28391, 11396],
            "FO-SCL": [33487, 17050],
            "JD-ITM": [474, 269],
            "JD-OTM": [84, 59],
            "M3-OTM": [6778, 2758],
            "TS-OTM": [108, 51]
        }

        # Amount-Daten: Format "Amount": [Gesamtverträge, Stornierte Verträge]
        self.amount_data = {
            "14,9": [16111, 7437],
            "19,9": [40662, 19891],
            "24,9": [1556, 674],
            "29,9": [3014, 1471],
            "9,9": [24545, 9707]
        }

        # Output-Ordner erstellen
        self.output_folder = os.path.join(os.environ["HOME"], "Desktop", "Call_Center_Analyse")
        os.makedirs(self.output_folder, exist_ok=True)

        # Theme für Plots
        sns.set_theme(style="whitegrid")
        plt.rcParams.update({'font.size': 12})

    def erstelle_dataframes(self):
        """Erstellt DataFrames aus den gegebenen Daten"""
        # DataFrame für Kampagnen-Daten
        kampagne_records = []
        for kampagne, values in self.kampagne_data.items():
            call_center, vertriebsweg = kampagne.split('-')
            kampagne_records.append({
                'Kampagne': kampagne,
                'Call_Center': call_center,
                'Vertriebsweg': vertriebsweg,
                'Gesamtverträge': values[0],
                'Stornierte_Verträge': values[1],
                'Stornoquote': round(values[1] / values[0] * 100, 2)
            })
        self.df_kampagne = pd.DataFrame(kampagne_records)

        # DataFrame für Amount-Daten
        amount_records = []
        for amount, values in self.amount_data.items():
            amount_records.append({
                'Amount': amount,
                'Gesamtverträge': values[0],
                'Stornierte_Verträge': values[1],
                'Stornoquote': round(values[1] / values[0] * 100, 2)
            })
        self.df_amount = pd.DataFrame(amount_records)

        return self.df_kampagne, self.df_amount

    def analyse_call_center(self):
        """Analysiert Daten nach Call-Center"""
        df_cc = self.df_kampagne.groupby('Call_Center').agg({
            'Gesamtverträge': 'sum',
            'Stornierte_Verträge': 'sum'
        }).reset_index()

        df_cc['Stornoquote'] = round(df_cc['Stornierte_Verträge'] / df_cc['Gesamtverträge'] * 100, 2)

        return df_cc

    def analyse_vertriebsweg(self):
        """Analysiert Daten nach Vertriebsweg"""
        df_vw = self.df_kampagne.groupby('Vertriebsweg').agg({
            'Gesamtverträge': 'sum',
            'Stornierte_Verträge': 'sum'
        }).reset_index()

        df_vw['Stornoquote'] = round(df_vw['Stornierte_Verträge'] / df_vw['Gesamtverträge'] * 100, 2)

        return df_vw

    def visualisiere_call_center(self, df_cc):
        """Visualisiert Call-Center Daten"""
        # 1. Balkendiagramm der Stornoquoten
        plt.figure(figsize=(12, 7))
        ax = sns.barplot(
            x='Call_Center',
            y='Stornoquote',
            data=df_cc,
            palette='viridis'
        )

        ax.set_title('Stornoquoten nach Call Center', fontsize=16)
        ax.set_xlabel('Call Center', fontsize=14)
        ax.set_ylabel('Stornoquote (%)', fontsize=14)

        # Formatierung als Prozent
        ax.yaxis.set_major_formatter(PercentFormatter())

        # Werte anzeigen
        for i, v in enumerate(df_cc['Stornoquote']):
            ax.text(i, v + 1, f'{v}%', ha='center')

        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'call_center_Stornoquote.png'))
        plt.close()

        # 2. Gegenüberstellung Gesamtverträge vs. stornierte Verträge
        plt.figure(figsize=(14, 8))
        df_melted = df_cc.melt(
            id_vars='Call_Center',
            value_vars=['Gesamtverträge', 'Stornierte_Verträge'],
            var_name='Typ',
            value_name='Anzahl'
        )

        ax = sns.barplot(
            x='Call_Center',
            y='Anzahl',
            hue='Typ',
            data=df_melted,
            palette=['#3498db', '#e74c3c']
        )

        ax.set_title('Gesamtverträge vs. Stornierte Verträge nach Call Center', fontsize=16)
        ax.set_xlabel('Call Center', fontsize=14)
        ax.set_ylabel('Anzahl', fontsize=14)

        # Werte anzeigen
        for i, p in enumerate(ax.patches):
            height = p.get_height()
            ax.text(p.get_x() + p.get_width() / 2., height + 100, f'{int(height)}',
                    ha='center', fontsize=10)

        plt.legend(title='')
        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'call_center_vergleich.png'))
        plt.close()

    def visualisiere_vertriebsweg(self, df_vw):
        """Visualisiert Vertriebsweg Daten"""
        # 1. Balkendiagramm der Stornoquoten
        plt.figure(figsize=(12, 7))
        ax = sns.barplot(
            x='Vertriebsweg',
            y='Stornoquote',
            data=df_vw,
            palette='mako'
        )

        ax.set_title('Stornoquoten nach Vertriebsweg', fontsize=16)
        ax.set_xlabel('Vertriebsweg', fontsize=14)
        ax.set_ylabel('Stornoquote (%)', fontsize=14)

        # Formatierung als Prozent
        ax.yaxis.set_major_formatter(PercentFormatter())

        # Werte anzeigen
        for i, v in enumerate(df_vw['Stornoquote']):
            ax.text(i, v + 1, f'{v}%', ha='center')

        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'vertriebsweg_Stornoquote.png'))
        plt.close()

        # 2. Donut-Chart für Vertriebswege nach Gesamtverträgen
        plt.figure(figsize=(10, 10))

        # Donut-Chart erstellen
        colors = sns.color_palette('mako', len(df_vw))
        plt.pie(
            df_vw['Gesamtverträge'],
            labels=df_vw['Vertriebsweg'],
            autopct='%1.1f%%',
            startangle=90,
            colors=colors,
            wedgeprops={'edgecolor': 'white', 'linewidth': 2}
        )

        # Kreis in der Mitte für Donut-Effekt
        centre_circle = plt.Circle((0, 0), 0.6, fc='white')
        plt.gcf().gca().add_artist(centre_circle)

        plt.title('Marktanteil der Vertriebswege (Gesamtverträge)', fontsize=16)
        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'vertriebsweg_marktanteil.png'))
        plt.close()

    def visualisiere_kampagne(self):
        """Visualisiert Kampagnen-Daten"""
        # 1. Balkendiagramm der Stornoquoten nach Kampagne
        plt.figure(figsize=(14, 8))

        ax = sns.barplot(
            x='Kampagne',
            y='Stornoquote',
            data=self.df_kampagne.sort_values('Stornoquote'),
            palette='plasma'
        )

        ax.set_title('Stornoquoten nach Kampagne', fontsize=16)
        ax.set_xlabel('Kampagne', fontsize=14)
        ax.set_ylabel('Stornoquote (%)', fontsize=14)

        # Formatierung als Prozent
        ax.yaxis.set_major_formatter(PercentFormatter())

        # Werte anzeigen
        for i, v in enumerate(ax.patches):
            width, height = v.get_width(), v.get_height()
            x, y = v.get_xy()
            ax.text(x + width / 2, height + 1,
                    f'{self.df_kampagne.sort_values("Stornoquote")["Stornoquote"].iloc[i]:.2f}%',
                    ha='center')

        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'kampagne_Stornoquote.png'))
        plt.close()

        # 2. Heatmap mit Call-Center und Vertriebsweg
        # Pivot-Tabelle erstellen
        pivot_df = pd.pivot_table(
            self.df_kampagne,
            values='Stornoquote',
            index='Call_Center',
            columns='Vertriebsweg'
        )

        plt.figure(figsize=(12, 8))
        sns.heatmap(
            pivot_df,
            annot=True,
            fmt='.2f',
            cmap='YlOrRd',
            cbar_kws={'label': 'Stornoquote (%)'},
            linewidths=1
        )

        plt.title('Stornoquote nach Call-Center und Vertriebsweg', fontsize=16)
        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'heatmap_cc_vertriebsweg.png'))
        plt.close()

        # 3. Bubble Chart für Kampagnen (x=Gesamtverträge, y=Stornoquote, size=Stornierte)
        plt.figure(figsize=(14, 10))

        # Farben nach Call-Center
        call_centers = self.df_kampagne['Call_Center'].unique()
        color_map = dict(zip(call_centers, sns.color_palette('Set1', len(call_centers))))

        for idx, row in self.df_kampagne.iterrows():
            plt.scatter(
                row['Gesamtverträge'],
                row['Stornoquote'],
                s=row['Stornierte_Verträge'] / 30,
                color=color_map[row['Call_Center']],
                alpha=0.7,
                edgecolors='white',
                linewidth=2
            )
            plt.text(
                row['Gesamtverträge'] + 200,
                row['Stornoquote'],
                row['Kampagne'],
                fontsize=11
            )

        # Legende für Call-Center
        patches = [mpatches.Patch(color=color, label=cc) for cc, color in color_map.items()]
        plt.legend(handles=patches, title='Call Center', loc='upper right')

        plt.title('Kampagnen-Übersicht: Gesamtverträge vs. Stornoquote', fontsize=16)
        plt.xlabel('Gesamtverträge', fontsize=14)
        plt.ylabel('Stornoquote (%)', fontsize=14)
        plt.grid(True, alpha=0.3)

        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'kampagne_bubble_chart.png'))
        plt.close()

    def visualisiere_amounts(self):
        """Visualisiert Amount-Daten"""
        # 1. Balkendiagramm der Stornoquoten nach Amount
        plt.figure(figsize=(12, 7))

        # Sortieren nach Amount (als Float)
        self.df_amount['Amount_Num'] = self.df_amount['Amount'].str.replace(',', '.').astype(float)
        sorted_df = self.df_amount.sort_values('Amount_Num')

        ax = sns.barplot(
            x='Amount',
            y='Stornoquote',
            data=sorted_df,
            palette='crest',
            order=sorted_df['Amount']
        )

        ax.set_title('Stornoquoten nach Amount', fontsize=16)
        ax.set_xlabel('Amount', fontsize=14)
        ax.set_ylabel('Stornoquote (%)', fontsize=14)

        # Formatierung als Prozent
        ax.yaxis.set_major_formatter(PercentFormatter())

        # Werte anzeigen
        for i, v in enumerate(sorted_df['Stornoquote']):
            ax.text(i, v + 1, f'{v}%', ha='center')

        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'amount_Stornoquote.png'))
        plt.close()

        # 2. Liniendiagramm für die Relation zwischen Amount und Stornoquote
        plt.figure(figsize=(12, 7))

        plt.plot(
            sorted_df['Amount_Num'],
            sorted_df['Stornoquote'],
            marker='o',
            linewidth=2,
            markersize=10,
            color='#1abc9c'
        )

        # Trendlinie
        z = np.polyfit(sorted_df['Amount_Num'], sorted_df['Stornoquote'], 1)
        p = np.poly1d(z)
        plt.plot(
            sorted_df['Amount_Num'],
            p(sorted_df['Amount_Num']),
            linestyle='--',
            color='#e74c3c'
        )

        # R-Quadrat berechnen
        y_mean = np.mean(sorted_df['Stornoquote'])
        ss_tot = sum((sorted_df['Stornoquote'] - y_mean) ** 2)
        ss_res = sum((sorted_df['Stornoquote'] - p(sorted_df['Amount_Num'])) ** 2)
        r_squared = 1 - (ss_res / ss_tot)

        plt.title(f'Relation zwischen Amount und Stornoquote (R² = {r_squared:.2f})', fontsize=16)
        plt.xlabel('Amount (€)', fontsize=14)
        plt.ylabel('Stornoquote (%)', fontsize=14)
        plt.grid(True, alpha=0.3)

        # Werte anzeigen
        for i, row in sorted_df.iterrows():
            plt.text(
                row['Amount_Num'] + 0.2,
                row['Stornoquote'] + 0.5,
                f"{row['Amount']}€: {row['Stornoquote']}%",
                fontsize=10
            )

        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'amount_relation.png'))
        plt.close()

        # 3. Balkendiagramm für Volumenverhältnis
        plt.figure(figsize=(14, 8))

        # Sortieren nach Amount
        sorted_df = self.df_amount.sort_values('Amount_Num')

        # Plotbereich aufteilen
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(20, 8))

        # Erster Plot: Gesamtverträge
        sns.barplot(
            x='Amount',
            y='Gesamtverträge',
            data=sorted_df,
            palette='Blues_d',
            order=sorted_df['Amount'],
            ax=ax1
        )

        ax1.set_title('Gesamtverträge nach Amount', fontsize=16)
        ax1.set_xlabel('Amount', fontsize=14)
        ax1.set_ylabel('Anzahl', fontsize=14)

        # Werte anzeigen
        for i, v in enumerate(sorted_df['Gesamtverträge']):
            ax1.text(i, v + 500, f'{int(v)}', ha='center')

        # Zweiter Plot: Stornierte Verträge
        sns.barplot(
            x='Amount',
            y='Stornierte_Verträge',
            data=sorted_df,
            palette='Reds_d',
            order=sorted_df['Amount'],
            ax=ax2
        )

        ax2.set_title('Stornierte Verträge nach Amount', fontsize=16)
        ax2.set_xlabel('Amount', fontsize=14)
        ax2.set_ylabel('Anzahl', fontsize=14)

        # Werte anzeigen
        for i, v in enumerate(sorted_df['Stornierte_Verträge']):
            ax2.text(i, v + 300, f'{int(v)}', ha='center')

        plt.tight_layout()
        plt.savefig(os.path.join(self.output_folder, 'amount_volumen.png'))
        plt.close()

    def erstelle_analyse_bericht(self):
        """Erstellt einen Bericht mit allen Analysen"""
        # Ausgabepfad
        report_path = os.path.join(self.output_folder, 'Analysebericht.txt')

        with open(report_path, 'w') as f:
            f.write("=== ANALYSEBERICHT: CALL-CENTER, VERTRIEBSWEGE UND KAMPAGNEN ===\n\n")

            # 1. Übersicht
            f.write("1. GESAMTÜBERBLICK\n")
            f.write("-----------------\n")
            gesamt = self.df_kampagne['Gesamtverträge'].sum()
            storniert = self.df_kampagne['Stornierte_Verträge'].sum()
            Stornoquote = round(storniert / gesamt * 100, 2)
            f.write(f"Gesamtanzahl Verträge: {gesamt}\n")
            f.write(f"Stornierte Verträge: {storniert}\n")
            f.write(f"GesamtStornoquote: {Stornoquote}%\n\n")

            # 2. Call-Center Analyse
            df_cc = self.analyse_call_center()
            f.write("2. CALL-CENTER ANALYSE\n")
            f.write("---------------------\n")
            for idx, row in df_cc.iterrows():
                f.write(f"Call-Center: {row['Call_Center']}\n")
                f.write(f"  Gesamtverträge: {row['Gesamtverträge']}\n")
                f.write(f"  Stornierte Verträge: {row['Stornierte_Verträge']}\n")
                f.write(f"  Stornoquote: {row['Stornoquote']}%\n\n")

            # 3. Vertriebsweg Analyse
            df_vw = self.analyse_vertriebsweg()
            f.write("3. VERTRIEBSWEG ANALYSE\n")
            f.write("-----------------------\n")
            for idx, row in df_vw.iterrows():
                f.write(f"Vertriebsweg: {row['Vertriebsweg']}\n")
                f.write(f"  Gesamtverträge: {row['Gesamtverträge']}\n")
                f.write(f"  Stornierte Verträge: {row['Stornierte_Verträge']}\n")
                f.write(f"  Stornoquote: {row['Stornoquote']}%\n\n")

            # 4. Kampagnen Analyse
            f.write("4. KAMPAGNEN ANALYSE\n")
            f.write("-------------------\n")
            for idx, row in self.df_kampagne.sort_values('Stornoquote', ascending=False).iterrows():
                f.write(f"Kampagne: {row['Kampagne']}\n")
                f.write(f"  Call-Center: {row['Call_Center']}\n")
                f.write(f"  Vertriebsweg: {row['Vertriebsweg']}\n")
                f.write(f"  Gesamtverträge: {row['Gesamtverträge']}\n")
                f.write(f"  Stornierte Verträge: {row['Stornierte_Verträge']}\n")
                f.write(f"  Stornoquote: {row['Stornoquote']}%\n\n")

            # 5. Amount Analyse
            f.write("5. AMOUNT ANALYSE\n")
            f.write("----------------\n")
            for idx, row in self.df_amount.sort_values('Amount_Num').iterrows():
                f.write(f"Amount: {row['Amount']}€\n")
                f.write(f"  Gesamtverträge: {row['Gesamtverträge']}\n")
                f.write(f"  Stornierte Verträge: {row['Stornierte_Verträge']}\n")
                f.write(f"  Stornoquote: {row['Stornoquote']}%\n\n")

        print(f"Analysebericht wurde erstellt: {report_path}")

    def run_analysis(self):
        """Führt die komplette Analyse und Visualisierung durch"""
        print("Starting analysis...")

        # Dataframes erstellen
        self.erstelle_dataframes()

        # Analyse nach Call-Center und Vertriebsweg
        df_cc = self.analyse_call_center()
        df_vw = self.analyse_vertriebsweg()

        # Visualisierungen erstellen
        self.visualisiere_call_center(df_cc)
        self.visualisiere_vertriebsweg(df_vw)
        self.visualisiere_kampagne()
        self.visualisiere_amounts()

        # Analysebericht erstellen
        self.erstelle_analyse_bericht()

        print(f"Analyse abgeschlossen. Alle Ergebnisse wurden im Ordner gespeichert: {self.output_folder}")


if __name__ == "__main__":
    analyzer = CallCenterAnalyse()
    analyzer.run_analysis()