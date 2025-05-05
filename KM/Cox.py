import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import confusion_matrix, accuracy_score
from lifelines import CoxPHFitter
import datetime
import warnings

from KM.KM import numerical_features

warnings.filterwarnings('ignore')


class StornoprognoseModel:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.numeric_features = ['Amount', 'contract_duration', 'days_from_start']
        self.categorical_features = ['Country_Region Code', 'Product Code', 'Kampagne']
        self.cox_model = None
        print(f"\n=== INITIALISIERE STORNOPROGNOSE-MODELL ===\nDateipfad: {self.file_path}")
        print(f"  → Numerische Features: {self.numeric_features}")
        print(f"  → Kategorische Features: {self.categorical_features}")

    def load_data(self):
        """
        Lädt und bereitet Daten aus der Excel-Datei vor
        """
        try:
            print(f"\nLade Daten aus: {self.file_path}")
            self.df = pd.read_excel(self.file_path)
            print(f"Daten erfolgreich geladen: {len(self.df)} Zeilen, {len(self.df.columns)} Spalten")

            # "Deletion Type" zu Integer konvertieren
            if 'Deletion Type' in self.df.columns:
                self.df['Deletion Type'] = pd.to_numeric(self.df['Deletion Type'], errors='coerce').fillna(0).astype(
                    int)
                print(f"'Deletion Type' zu Integer konvertiert")

            # Verteilung der Löschungstypen anzeigen
            if 'Deletion Type' in self.df.columns:
                deletion_counts = self.df['Deletion Type'].value_counts()
                print("\nVerteilung der Löschungstypen:")
                for dt, count in deletion_counts.items():
                    print(f"  Typ {dt}: {count} Datensätze ({count / len(self.df) * 100:.1f}%)")

            return True
        except Exception as e:
            print(f"Fehler beim Laden der Daten: {e}")
            return False

    def preprocess_data(self):
        """
        Bereitet die Daten für das Cox-Modell auf
        """
        print("\nBereite Daten für Stornoprognose vor...")
        df_prep = self.df.copy()

        # Datumsfelder konvertieren
        date_columns = ['Start Insurance', 'End Insurance']
        for col in date_columns:
            if col in df_prep.columns:
                df_prep[col] = pd.to_datetime(df_prep[col], errors='coerce')

        # Feature-Engineering
        today = pd.Timestamp.now()

        # Vertragsdauer
        if 'Start Insurance' in df_prep.columns and 'End Insurance' in df_prep.columns:
            mask = df_prep['Start Insurance'].notna() & df_prep['End Insurance'].notna()
            df_prep.loc[mask, 'contract_duration'] = (df_prep.loc[mask, 'End Insurance'] -
                                                      df_prep.loc[mask, 'Start Insurance']).dt.days
            df_prep['contract_duration'] = df_prep['contract_duration'].fillna(365).clip(lower=1)

        # Zeit seit Vertragsbeginn
        if 'Start Insurance' in df_prep.columns:
            mask = df_prep['Start Insurance'].notna()
            df_prep.loc[mask, 'days_from_start'] = (today - df_prep.loc[mask, 'Start Insurance']).dt.days
            df_prep['days_from_start'] = df_prep['days_from_start'].fillna(180).clip(lower=0)

        # Zielvariable: Storno (1) oder nicht (0)
        df_prep['event'] = (df_prep['Deletion Type'] != 0).astype(int)
        print(f"  → Zielvariable 'event' erstellt: {df_prep['event'].value_counts().to_dict()}")

        # Beobachtungszeit für Cox-Modell
        df_prep['observation_time'] = df_prep['days_from_start']

        # Daten in Training und Test aufteilen (80/20 statt 70/30)
        train_df, test_df = train_test_split(df_prep, test_size=0.2, random_state=42, stratify=df_prep['event'])
        print(
            f"  → Daten aufgeteilt: {len(train_df)} Training ({len(train_df) / len(df_prep):.1%}), {len(test_df)} Test ({len(test_df) / len(df_prep):.1%})")

        self.train_df = train_df
        self.test_df = test_df

        return True

    def prepare_model_features(self, df):
        """
        Verbesserte Feature-Vorbereitung mit besonderem Fokus auf wichtige Features
        """
        model_df = df.copy()

        print(f"Verfügbare Features: {model_df.columns.tolist()}")

        # Numerische Features behandeln
        for col in self.numeric_features:
            if col in model_df.columns:
                print(f"Verarbeite numerisches Feature: {col}")

                # Besondere Behandlung für Amount-Spalte
                if col == 'Amount':
                    # Detaillierte Konvertierung für Amount
                    if isinstance(model_df[col].iloc[0], str):
                        model_df[col] = model_df[col].astype(str).str.replace(',', '.').str.replace(' ', '')

                    numeric_values = pd.to_numeric(model_df[col], errors='coerce')
                    print(
                        f"  → Amount-Statistiken: Min={numeric_values.min()}, Median={numeric_values.median()}, Max={numeric_values.max()}")

                    # Median der gültigen Werte für Ersetzung verwenden
                    median_val = numeric_values.median()
                    model_df[col] = numeric_values.fillna(median_val)
                else:
                    # Für andere numerische Spalten
                    model_df[col] = pd.to_numeric(model_df[col], errors='coerce')
                    if not model_df[col].dropna().empty:
                        median_val = model_df[col].dropna().median()
                        print(
                            f"  → {col}-Statistiken: Min={model_df[col].min()}, Median={median_val}, Max={model_df[col].max()}")
                        model_df[col] = model_df[col].fillna(median_val)
                    else:
                        print(f"  → Warnung: {col} enthält keine gültigen Werte!")
                        model_df[col] = 0
            else:
                print(f"  → Warnung: Feature {col} nicht in Datensatz gefunden!")
                model_df[col] = 0

        # Kategorische Features kodieren
        for col in self.categorical_features:
            if col in model_df.columns:
                print(f"Verarbeite kategorisches Feature: {col}")
                # One-Hot-Encoding mit verbesserten Fallback-Werten
                model_df[col] = model_df[col].fillna('unknown').astype(str)
                unique_values = model_df[col].nunique()
                print(f"  → {col} hat {unique_values} eindeutige Werte")

                # One-Hot-Encoding
                dummies = pd.get_dummies(model_df[col], prefix=col, drop_first=True)
                model_df = pd.concat([model_df, dummies], axis=1)
            else:
                print(f"  → Warnung: Kategorisches Feature {col} nicht gefunden!")

        return model_df

    def train_cox_model(self):
        """
        Trainiert das Cox-Modell mit den vorbereiteten Daten
        """
        print("\nTrainiere Cox-Modell für Stornoprognose...")

        # Features vorbereiten
        model_df = self.prepare_model_features(self.train_df)

        # Feature-Spalten und ihre Transformationen speichern
        self.feature_names = []

        # Numerische Features
        for col in self.numeric_features:
            if col in model_df.columns:
                self.feature_names.append(col)

        # One-Hot-kodierte kategorische Features
        for col in self.categorical_features:
            dummy_cols = [c for c in model_df.columns if c.startswith(f"{col}_")]
            self.feature_names.extend(dummy_cols)

        print(f"  → Training mit {len(self.feature_names)} Features")

        # Cox-Modell trainieren
        try:
            # Daten für das Cox-Modell vorbereiten
            cph_df = model_df[self.feature_names + ['observation_time', 'event']].copy()

            # Numerische Features skalieren und Transformer speichern
            self.scalers = {}
            for col in [f for f in self.numeric_features if f in cph_df.columns]:
                scaler = StandardScaler()
                cph_df[col] = scaler.fit_transform(cph_df[[col]])
                self.scalers[col] = scaler

            # Cox-Modell trainieren
            cph = CoxPHFitter(penalizer=0.1)
            cph.fit(cph_df, duration_col='observation_time', event_col='event')

            self.cox_model = cph
            print("  ✓ Cox-Modell erfolgreich trainiert!")
            return True
        except Exception as e:
            print(f"  ✗ Fehler beim Training des Cox-Modells: {e}")
            return False

    def predict(self, time_horizon=180):
        """
        Erstellt Stornoprognosen für die Testdaten
        """
        print(f"\nErstelle Stornoprognosen für {len(self.test_df)} Datensätze...")

        # Features für Testdaten vorbereiten
        test_features = self.prepare_model_features(self.test_df)

        # Sicherstellen, dass alle Trainingsfeatures vorhanden sind
        for feature in self.feature_names:
            if feature not in test_features.columns:
                print(f"  → Fehlendes Feature '{feature}' wird ergänzt")
                test_features[feature] = 0

        # Features auf die im Training verwendeten beschränken
        test_features = test_features[self.feature_names].copy()

        # Gleiche Skalierung wie beim Training anwenden
        for col, scaler in self.scalers.items():
            if col in test_features.columns:
                test_features[col] = scaler.transform(test_features[[col]])

        # Ergebnis-DataFrame initialisieren
        result_df = self.test_df.copy()

        try:
            # Überlebenswahrscheinlichkeit vorhersagen
            survival_prob = self.cox_model.predict_survival_function(test_features, times=[time_horizon])
            result_df['survival_probability'] = np.squeeze(np.array(survival_prob))
            result_df['cancellation_probability'] = 1 - result_df['survival_probability']

            # Analyse der Wahrscheinlichkeitsverteilung für optimalen Schwellenwert
            probs = result_df['cancellation_probability']
            print(
                f"  → Stornowahrscheinlichkeiten: Min={probs.min():.4f}, Median={probs.median():.4f}, Max={probs.max():.4f}")

            # Ziel: Balance zwischen Precision und Recall
            y_true = result_df['event']

            # Verschiedene Schwellenwerte testen
            thresholds = [np.percentile(probs, p) for p in [50, 60, 70, 75, 80, 90]]
            best_f1 = 0
            best_threshold = thresholds[0]

            print("\nOptimiere Schwellenwert...")
            for thresh in thresholds:
                y_pred = (probs > thresh).astype(int)

                # Elemente der Konfusionsmatrix berechnen
                tn = sum((y_true == 0) & (y_pred == 0))
                fp = sum((y_true == 0) & (y_pred == 1))
                fn = sum((y_true == 1) & (y_pred == 0))
                tp = sum((y_true == 1) & (y_pred == 1))

                # Metriken
                prec = tp / (tp + fp) if (tp + fp) > 0 else 0
                rec = tp / (tp + fn) if (tp + fn) > 0 else 0
                f1 = 2 * prec * rec / (prec + rec) if (prec + rec) > 0 else 0

                print(f"  → Schwelle {thresh:.4f}: F1={f1:.4f}, Precision={prec:.4f}, Recall={rec:.4f}")

                if f1 > best_f1:
                    best_f1 = f1
                    best_threshold = thresh

            print(f"  → Bester Schwellenwert: {best_threshold:.4f} (F1={best_f1:.4f})")
            threshold = best_threshold

            result_df['predicted_cancellation'] = (result_df['cancellation_probability'] > threshold).astype(int)

            # Vorhersagestatistik
            pos_count = result_df['predicted_cancellation'].sum()
            print(f"  → {pos_count} Stornos vorhergesagt ({pos_count / len(result_df):.1%})")

            self.predictions = result_df
            return result_df
        except Exception as e:
            print(f"  ��� Fehler bei der Vorhersage: {e}")
            import traceback
            traceback.print_exc()
            return None

    def create_confusion_matrix(self):
        """
        Erstellt eine detaillierte Konfusionsmatrix und Leistungsmetriken
        """
        print("\n=== KONFUSIONSMATRIX FÜR STORNOPROGNOSE ===")

        # Sicherstellen, dass Vorhersagen vorhanden sind
        if not hasattr(self, 'predictions'):
            print("  ✗ Keine Vorhersagen gefunden! Bitte zuerst predict() ausführen.")
            return None

        # Tatsächliche Werte und Vorhersagen extrahieren
        y_true = self.predictions['event'].astype(int)
        y_pred = self.predictions['predicted_cancellation'].astype(int)

        # Elemente der Konfusionsmatrix manuell berechnen
        tn = sum((y_true == 0) & (y_pred == 0))
        fp = sum((y_true == 0) & (y_pred == 1))
        fn = sum((y_true == 1) & (y_pred == 0))
        tp = sum((y_true == 1) & (y_pred == 1))

        # Matrix erstellen
        cm = np.array([[tn, fp], [fn, tp]])

        # Metriken berechnen und ausgeben (wie bisher)
        accuracy = (tp + tn) / (tp + tn + fp + fn)
        precision = tp / (tp + fp) if (tp + fp) > 0 else 0
        recall = tp / (tp + fn) if (tp + fn) > 0 else 0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall > 0 else 0

        # Textausgabe wie gehabt...
        print("\n" + "-" * 70)
        print("|                   | Vorhergesagt: Kein Storno | Vorhergesagt: Storno |")
        print("|-------------------+-------------------------+---------------------|")
        print(f"| Tatsächlich: Kein |       {tn:8d} (TN)       |     {fp:8d} (FP)     |")
        print(f"| Tatsächlich: Storno |       {fn:8d} (FN)       |     {tp:8d} (TP)     |")
        print("-" * 70)

        print("\nLEISTUNGSMETRIKEN:")
        print(f"  → Accuracy:  {accuracy:.4f}")
        print(f"  → Precision: {precision:.4f} (Korrekte Stornos / Alle vorhergesagten Stornos)")
        print(f"  → Recall:    {recall:.4f} (Korrekte Stornos / Alle tatsächlichen Stornos)")
        print(f"  → F1-Score:  {f1:.4f}")

        # KORRIGIERTER TEIL: Visualisierung ohne doppelte Zahlen
        plt.figure(figsize=(10, 8))

        # 1. Leere Heatmap ohne jegliche Annotationen erstellen
        ax = sns.heatmap(cm, annot=False, fmt='', cbar=True, cmap='Blues',
                         xticklabels=['Kein Storno', 'Storno'],
                         yticklabels=['Kein Storno', 'Storno'])

        # 2. Manuell genau EINE Annotation pro Zelle hinzufügen
        for i in range(2):
            for j in range(2):
                text_color = "white" if cm[i, j] > cm.max() / 2 else "black"
                ax.text(j + 0.5, i + 0.5, str(cm[i, j]),
                        ha="center", va="center", fontsize=16,
                        fontweight='bold', color=text_color)

        plt.title('Konfusionsmatrix: Cox-Modell Stornoprognose', fontsize=16, fontweight='bold')
        plt.xlabel('Vorhergesagte Klasse', fontsize=14)
        plt.ylabel('Tatsächliche Klasse', fontsize=14)
        plt.tight_layout()
        plt.savefig("konfusionsmatrix_stornoprognose.png", dpi=300)
        plt.close()

        print(f"\n→ Konfusionsmatrix gespeichert als 'konfusionsmatrix_stornoprognose.png'")

        return {
            'confusion_matrix': cm,
            'accuracy': accuracy,
            'precision': precision,
            'recall': recall,
            'f1': f1
        }


# Hauptfunktion zum Ausführen der Stornoprognose
def run_stornoprognose(file_path):
    """
    Führt die komplette Stornoprognose mit Cox-Modell durch
    """
    print("\n" + "=" * 80)
    print("STORNOPROGNOSE MIT COX-MODELL")
    print("=" * 80)

    try:
        # Modell initialisieren
        model = StornoprognoseModel(file_path)

        # Daten laden und verarbeiten
        if model.load_data() and model.preprocess_data():
            # Cox-Modell trainieren
            if model.train_cox_model():
                # Stornoprognose erstellen
                model.predict(time_horizon=180)

                # Konfusionsmatrix erstellen und Leistung bewerten
                metrics = model.create_confusion_matrix()

                print("\n" + "=" * 50)
                print("ERGEBNIS DER STORNOPROGNOSE:")
                print(f"  → Accuracy:  {metrics['accuracy']:.4f}")
                print(f"  → Precision: {metrics['precision']:.4f}")
                print(f"  → Recall:    {metrics['recall']:.4f}")
                print(f"  → F1-Score:  {metrics['f1']:.4f}")
                print("=" * 50)
            else:
                print("⚠️ Cox-Modell konnte nicht erstellt werden")
        else:
            print("⚠️ Fehler beim Laden oder Verarbeiten der Daten")

    except Exception as e:
        print(f"⚠️ Fehler bei der Stornoprognose: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    import os

    file_path = os.path.join(os.environ["HOME"], "Desktop", "Training.xlsx")
    run_stornoprognose(file_path)