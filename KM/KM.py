# =====================================================================
# AUTOR: @Adrian Stötzler
# TITEL: Kaplan-Meier basiertes Tool zur Ermittlung der Stornowahrscheinlichkeit
# BESCHREIBUNG: Dieses Skript analysiert Versicherungsdaten mit der Kaplan-Meier-Methode,
# um Stornovorhersagen zu treffen. Es verwendet Überlebenszeitanalyse, um die
# Wahrscheinlichkeit einer Kündigung zu prognostizieren.
# =====================================================================

# ===== BIBLIOTHEKEN IMPORTIEREN =====
import os  # Für Datei- und Pfadoperationen
import pandas as pd  # Für Datenverarbeitung und -analyse
import numpy as np  # Für numerische Berechnungen
from sklearn.model_selection import train_test_split  # Für Datenaufteilung
from sklearn.preprocessing import StandardScaler  # Für Datennormalisierung
from sklearn.metrics import classification_report, confusion_matrix, accuracy_score, roc_auc_score, roc_curve
import matplotlib.pyplot as plt  # Für Visualisierungen
import seaborn as sns  # Für erweiterte Visualisierungen
from lifelines import KaplanMeierFitter, CoxPHFitter  # Für Überlebenszeitanalyse
import datetime  # Für Datumsfunktionen

# Warnungen unterdrücken (optional)
import warnings

warnings.filterwarnings('ignore')

print("=" * 80)
print("KAPLAN-MEIER STORNOPROGNOSE-TOOL")
print("=" * 80)

# ===== DATEIPFAD DEFINIEREN =====
file_path = os.path.join(os.environ["HOME"], "Desktop", "Training.xlsx")
print(f"Versuche Daten zu laden von: {file_path}")


# ===== DATUMS-VERARBEITUNGSFUNKTION =====
def process_date_columns(df, date_columns):
    """
    Konvertiert Datumsspalten in datetime-Format und behandelt fehlende Werte

    Parameter:
        df (pandas.DataFrame): DataFrame mit den zu konvertierenden Spalten
        date_columns (list): Liste mit den Spaltennamen für Datumskonvertierung

    Rückgabe:
        pandas.DataFrame: DataFrame mit konvertierten Datumsspalten
    """
    print(f"Konvertiere {len(date_columns)} Datumsspalten...")
    for col in date_columns:
        if col in df.columns:
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                print(f"  ✓ '{col}' zu Datum konvertiert")
            except Exception as e:
                print(f"  ✗ Fehler bei Konvertierung von '{col}': {e}")
    return df


# ===== FUNKTION FÜR FEATURE-ENGINEERING =====
# ===== FUNKTION FÜR FEATURE-ENGINEERING =====
def create_survival_features(df):
    """
    Erstellt Features für die Überlebenszeitanalyse ohne Kündigungsinformationen zu verwenden
    """
    print("\nErstelle Features für Überlebenszeitanalyse...")

    # Kopie erstellen, um Original nicht zu verändern
    df_survival = df.copy()

    # Zielvariable: Wurde der Vertrag gekündigt? (1=ja, 0=nein)
    df_survival['event'] = (df_survival['Deletion Type'] != 0).astype(int)
    print(f"  → Zielvariable 'event' erstellt: {df_survival['event'].value_counts().to_dict()}")

    # Versicherungsdauer berechnen (wenn vorhanden)
    if 'Start Insurance' in df_survival.columns and 'End Insurance' in df_survival.columns:
        df_survival['duration'] = (df_survival['End Insurance'] - df_survival['Start Insurance']).dt.days
        df_survival['duration'] = df_survival['duration'].fillna(-1).clip(lower=0)
        print(f"  → Feature 'duration' erstellt (Versicherungsdauer in Tagen)")

    # Beobachtungszeitraum (bis heute)
    today = datetime.datetime.now()
    print(f"  → Aktuelles Datum für Berechnung: {today.strftime('%d.%m.%Y')}")

    # Zeit von Start bis heute für alle Verträge
    mask_with_start = df_survival['Start Insurance'].notna()
    df_survival.loc[mask_with_start, 'observation_time'] = (today - df_survival.loc[mask_with_start, 'Start Insurance']).dt.days
    print(f"  → Für {mask_with_start.sum()} Verträge: Beobachtungszeit = Zeit bis heute")

    # Für gekündigte Verträge: Zeit von Start bis zur Kündigung (nur für Trainingsdaten)
    # Dieses Feature wird NUR für das Training verwendet, nicht für Vorhersagen
    if 'Meldedatum' in df_survival.columns:
        # Meldedatum als neutralere Variable verwenden
        mask_with_meldedatum = (df_survival['event'] == 1) & (df_survival['Start Insurance'].notna()) & (df_survival['Meldedatum'].notna())
        if mask_with_meldedatum.sum() > 0:
            df_survival.loc[mask_with_meldedatum, 'time_until_event'] = (
                    df_survival.loc[mask_with_meldedatum, 'Meldedatum'] -
                    df_survival.loc[mask_with_meldedatum, 'Start Insurance']
            ).dt.days
            print(f"  → Für {mask_with_meldedatum.sum()} gekündigte Verträge: Zeit bis zum Meldedatum berechnet (nur für Training)")

    # Fehlende Werte auffüllen und negative Werte vermeiden
    missing_before = df_survival['observation_time'].isna().sum()
    df_survival['observation_time'] = df_survival['observation_time'].fillna(-1).clip(lower=0)
    print(f"  → {missing_before} Datensätze hatten keine Beobachtungszeit und wurden auf 0 gesetzt")

    # WICHTIG: Entfernen der kündigungsbezogenen Spalten für Vorhersagen
    leakage_columns = ['Deletion allowed at', 'Promised Deletion Date', 'Last Due Date']
    for col in leakage_columns:
        if col in df_survival.columns:
            df_survival.drop(col, axis=1, inplace=True)
            print(f"  → '{col}' entfernt um Data Leakage zu vermeiden")

    # Verträge mit gültigen Beobachtungszeiten auswählen
    valid_data = df_survival[df_survival['observation_time'] >= 0].copy()
    print(f"  → {len(valid_data)} von {len(df_survival)} Datensätzen haben gültige Beobachtungszeiten")

    return valid_data


# ===== FUNCTION FÜR KAPLAN-MEIER-ANALYSE =====
def perform_kaplan_meier_analysis(df, stratify_columns=None):
    """
    Führt Kaplan-Meier-Analyse durch und visualisiert die Überlebenskurven

    Parameter:
        df (pandas.DataFrame): DataFrame mit event und observation_time Spalten
        stratify_columns (list): Spalten, nach denen stratifiziert werden soll

    Rückgabe:
        dict: Dictionary mit KM-Kurven für verschiedene Gruppen


    """

    if stratify_columns is None:
        stratify_columns = []

    print("\nFühre Kaplan-Meier-Analyse durch...")
    km_curves = {}

    # Basis-Analyse für gesamten Datensatz
    kmf = KaplanMeierFitter()
    kmf.fit(durations=df['observation_time'], event_observed=df['event'], label='Gesamt')
    km_curves['overall'] = kmf

    # Visualisierung der Gesamtkurve
    plt.figure(figsize=(12, 6))
    kmf.plot_survival_function()
    plt.title('Kaplan-Meier Überlebenskurve: Wahrscheinlichkeit des Vertragsbestands')
    plt.xlabel('Zeit (Tage)')
    plt.ylabel('Wahrscheinlichkeit ohne Kündigung')
    plt.grid(True)
    plt.savefig("kaplan_meier_overall.png")
    plt.close()
    print(f"  → Gesamte Überlebenskurve erstellt und gespeichert")

    # Stratifizierte Analyse für ausgewählte Features
    for col in stratify_columns:
        if col not in df.columns:
            continue

        print(f"  → Stratifiziere nach '{col}'...")
        # Kategorische Spalten in Gruppen aufteilen

        # Für kontinuierliche Variablen: Quantile verwenden
        if df[col].nunique() > 10 and pd.api.types.is_numeric_dtype(df[col]):
            df['temp_group'] = pd.qcut(df[col], q=4, duplicates='drop', labels=False)
            groups = df['temp_group'].dropna().unique()
            group_col = 'temp_group'
        else:
            # Für kategorische Variablen: Top-N-Kategorien verwenden
            top_categories = df[col].value_counts().head(5).index
            df['temp_group'] = df[col].apply(lambda x: x if x in top_categories else 'Other')
            groups = df['temp_group'].dropna().unique()
            group_col = 'temp_group'

        # Plot erstellen
        plt.figure(figsize=(12, 6))
        km_dict = {}

        for group in groups:
            mask = df[group_col] == group
            if mask.sum() > 30:  # Nur Gruppen mit genügend Daten
                kmf = KaplanMeierFitter()
                label = f"{col}={group}"
                kmf.fit(durations=df.loc[mask, 'observation_time'],
                        event_observed=df.loc[mask, 'event'],
                        label=label)
                kmf.plot_survival_function()
                km_dict[group] = kmf

        plt.title(f'Kaplan-Meier Kurven stratifiziert nach {col}')
        plt.xlabel('Zeit (Tage)')
        plt.ylabel('Wahrscheinlichkeit ohne Kündigung')
        plt.grid(True)
        plt.legend()
        plt.savefig(f"kaplan_meier_{col}.png")
        plt.close()

        km_curves[col] = km_dict

        # Bereinigen
        if 'temp_group' in df.columns:
            df.drop('temp_group', axis=1, inplace=True)

    return km_curves


def build_cox_model(df, features):
    """
    Erstellt ein robustes Cox Proportional Hazards Modell ohne Kündigungsinformationen
    """
    print("\nTrainiere Cox Proportional Hazards Modell...")

    # Entferne alle potenziellen Leakage-Spalten aus den Features
    leakage_features = ['days_until_deletion_allowed', 'Deletion allowed at',
                        'Promised Deletion Date', 'Last Due Date']
    safe_features = [f for f in features if
                     f not in leakage_features and not any(leak in f for leak in leakage_features)]

    print(f"  → {len(features) - len(safe_features)} potenzielle Leakage-Features entfernt")

    # Feature-Vorbereitung
    model_features = []
    for col in safe_features:
        if col in df.columns and col not in ['observation_time', 'event']:
            if pd.api.types.is_numeric_dtype(df[col]):
                model_features.append(col)
            else:
                print(f"  → Überspringe nicht-numerische Spalte: {col}")

    # Nur mit numerischen Features fortfahren
    if len(model_features) < 2:
        print("  ✗ Zu wenige numerische Features für Cox-Modell")
        return None

    # Daten für Cox-Modell vorbereiten
    model_df = df[model_features + ['observation_time', 'event']].copy()

    # Multikollinearität erkennen und entfernen
    from sklearn.feature_selection import VarianceThreshold

    # 1. Features mit nahezu konstanten Werten entfernen
    selector = VarianceThreshold(threshold=0.01)
    X_numeric = model_df[model_features]
    selector.fit(X_numeric)
    selected_features = [model_features[i] for i in range(len(model_features)) if selector.get_support()[i]]

    if len(selected_features) == 0:
        print("  ✗ Keine Features nach Varianzfilterung übrig")
        return None

    print(f"  → {len(selected_features)}/{len(model_features)} Features nach Varianzfilterung behalten")

    # 2. Korrelationsanalyse
    corr_matrix = X_numeric[selected_features].corr().abs()
    upper = corr_matrix.where(np.triu(np.ones(corr_matrix.shape), k=1).astype(bool))
    to_drop = [column for column in upper.columns if any(upper[column] > 0.95)]

    selected_features = [f for f in selected_features if f not in to_drop]
    if len(to_drop) > 0:
        print(f"  → {len(to_drop)} stark korrelierende Features entfernt")

    # Standardisierung der Features
    from sklearn.preprocessing import StandardScaler
    scaler = StandardScaler()
    model_df[selected_features] = scaler.fit_transform(model_df[selected_features])

    # Cox-Modell mit Regularisierung trainieren
    print(f"  → Training mit {len(selected_features)} Features: {', '.join(selected_features)}")
    cph = CoxPHFitter(penalizer=0.1)  # Regularisierung hinzufügen

    try:
        cph.fit(model_df[selected_features + ['observation_time', 'event']],
                duration_col='observation_time',
                event_col='event',
                robust=True,  # Robuste Standardfehler
                step_size=0.5)  # Kleinere Schrittweite für bessere Konvergenz

        # Modellzusammenfassung
        print("\nCox-Modell erfolgreich trainiert!")
        print(cph.summary[['coef', 'exp(coef)', 'p']])

        return cph
    except Exception as e:
        print(f"  ✗ Fehler beim Training des Cox-Modells: {e}")

        # Alternative mit noch mehr Regularisierung versuchen
        try:
            print("  → Versuche mit stärkerer Regularisierung (penalizer=0.5)...")
            cph = CoxPHFitter(penalizer=0.5)
            cph.fit(model_df[selected_features + ['observation_time', 'event']],
                    duration_col='observation_time',
                    event_col='event')
            print("  ✓ Cox-Modell mit stärkerer Regularisierung erfolgreich trainiert")
            return cph
        except:
            print("  ✗ Cox-Modelltraining weiterhin nicht erfolgreich")
            return None

def predict_cancellation_kaplan_meier(km_curves, df_test, stratify_columns, time_horizon=1000):
    """
    Vorhersage der Kündigungswahrscheinlichkeiten ohne Kündigungsinformationen
    """
    print(f"\nErstelle Kaplan-Meier-basierte Stornoprognose für Zeithorizont: {time_horizon} Tage...")

    # Entferne alle potenziellen Leakage-Spalten aus den Eingabedaten
    result_df = df_test.copy()
    leakage_columns = ['Deletion allowed at', 'Promised Deletion Date', 'Last Due Date',
                       'days_until_deletion_allowed']

    for col in leakage_columns:
        if col in result_df.columns:
            result_df.drop(col, axis=1, inplace=True)
            print(f"  → '{col}' aus Vorhersagedaten entfernt (Data Leakage)")

    # Auch Spalten entfernen, die diese Namen als Teil haben
    leakage_patterns = ['deletion', 'kündig', 'cancel']
    for col in result_df.columns:
        if any(pattern in col.lower() for pattern in leakage_patterns):
            if col not in ['event']:  # event nicht entfernen
                result_df.drop(col, axis=1, inplace=True)
                print(f"  → '{col}' könnte Leakage enthalten und wurde entfernt")

    result_df['cancellation_probability'] = None

    # Verwende allgemeine Überlebensfunktion für alle Datensätze
    overall_curve = km_curves['overall']
    timeline = overall_curve.timeline
    closest_time = min(timeline, key=lambda x: abs(x - time_horizon))

    # Standardwert für alle Einträge
    default_surv_prob = float(overall_curve.survival_function_.loc[closest_time])
    result_df['cancellation_probability'] = 1 - default_surv_prob

    # Stratifizierte Vorhersagen (nur mit sicheren Spalten)
    safe_stratify = [col for col in stratify_columns if col not in leakage_columns and
                     not any(pattern in col.lower() for pattern in leakage_patterns)]

    print(f"  → Verwende {len(safe_stratify)}/{len(stratify_columns)} sichere Stratifizierungsspalten")

    for col in safe_stratify:
        if col in km_curves:
            print(f"  → Verwende Stratifizierung nach '{col}' für präzisere Vorhersagen")
            for group, curve in km_curves[col].items():
                mask = df_test[col] == group
                if mask.sum() > 0 and closest_time in curve.survival_function_.index:
                    surv_prob = float(curve.survival_function_.loc[closest_time])
                    result_df.loc[mask, 'cancellation_probability'] = 1 - surv_prob

    # WICHTIG: Fester Schwellenwert (0.15) für bessere Kündigungserkennung
    threshold = 0.15  # Niedrigerer Schwellenwert als 0.3
    print(f"  → Verwende festen niedrigen Schwellenwert: {threshold:.4f} für Kaplan-Meier-Klassifikation")
    result_df['predicted_cancellation'] = (result_df['cancellation_probability'] > threshold).astype(int)

    # Ausgabe der Verteilung
    pred_rate = result_df['predicted_cancellation'].mean()
    print(f"  → Stornoprognose für {len(result_df)} Datensätze erstellt")
    print(f"  → Durchschnittliche Stornowahrscheinlichkeit: {result_df['cancellation_probability'].mean():.4f}")
    print(f"  → Anteil vorhergesagter Kündigungen: {pred_rate:.4f} ({int(pred_rate * len(result_df))} von {len(result_df)})")

    return result_df


def predict_cancellation_cox(cox_model, df, kaplan_meier_curves=None, time_horizon=1000):
    """
    Sagt Kündigungswahrscheinlichkeit basierend auf Cox-Modell voraus ohne Kündigungsinformationen
    """
    print(f"\nErstelle Cox-basierte Stornoprognose für Zeithorizont: {time_horizon} Tage...")

    # Entferne alle potenziellen Leakage-Spalten aus den Eingabedaten
    result_df = df.copy()
    leakage_columns = ['Deletion allowed at', 'Promised Deletion Date', 'Last Due Date',
                       'days_until_deletion_allowed']

    for col in leakage_columns:
        if col in result_df.columns:
            result_df.drop(col, axis=1, inplace=True)
            print(f"  → '{col}' aus Vorhersagedaten entfernt (Data Leakage)")

    # Auch Spalten entfernen, die diese Namen als Teil haben
    leakage_patterns = ['deletion', 'kündig', 'cancel']
    for col in result_df.columns:
        if any(pattern in col.lower() for pattern in leakage_patterns):
            if col not in ['event']:  # event nicht entfernen
                result_df.drop(col, axis=1, inplace=True)
                print(f"  → '{col}' könnte Leakage enthalten und wurde entfernt")

    if cox_model is not None:
        # Vorhersage der Überlebenswahrscheinlichkeit für den angegebenen Zeithorizont
        try:
            # Survivalfunktion für jede Person vorhergesagt
            survival_prob = cox_model.predict_survival_function(result_df, times=[time_horizon])
            result_df['survival_probability'] = np.squeeze(np.array(survival_prob))
            result_df['cancellation_probability'] = 1 - result_df['survival_probability']

            print(f"  ✓ Cox-Modell-Vorhersage für {len(result_df)} Datensätze erfolgreich")
        except Exception as e:
            print(f"  ✗ Fehler bei Cox-Modell-Vorhersage: {e}")
            if kaplan_meier_curves is not None:
                print("  → Fallback auf Kaplan-Meier-Vorhersage")
                return predict_cancellation_kaplan_meier(kaplan_meier_curves, df, [], time_horizon)
            else:
                print("  ✗ Keine Fallback-Option verfügbar")
                result_df['cancellation_probability'] = 0.5
    else:
        # Wenn kein Cox-Modell verfügbar, verwende Kaplan-Meier als Fallback
        if kaplan_meier_curves is not None:
            print("  → Kein Cox-Modell verfügbar, verwende Kaplan-Meier")
            return predict_cancellation_kaplan_meier(kaplan_meier_curves, df, [], time_horizon)
        else:
            print("  ✗ Kein Modell verfügbar für Vorhersage")
            result_df['cancellation_probability'] = 0.5

    # WICHTIG: Fester Schwellenwert (0.15) für bessere Kündigungserkennung
    threshold = 0.15  # Niedrigerer Schwellenwert als 0.3
    print(f"  → Verwende festen niedrigen Schwellenwert: {threshold:.4f} für Cox-Klassifikation")
    result_df['predicted_cancellation'] = (result_df['cancellation_probability'] > threshold).astype(int)

    # Ausgabe der Verteilung
    pred_rate = result_df['predicted_cancellation'].mean()
    print(f"  → Durchschnittliche Stornowahrscheinlichkeit: {result_df['cancellation_probability'].mean():.4f}")
    print(f"  → Anteil vorhergesagter Kündigungen: {pred_rate:.4f} ({int(pred_rate * len(result_df))} von {len(result_df)})")

    return result_df


def evaluate_model(df):
    """
    Bewertet die Vorhersagequalität mit korrekter Konfusionsmatrix
    """
    print("\nBewerte Modellqualität...")

    # Sicherstellen, dass die benötigten Spalten vorhanden sind
    if 'event' not in df.columns or 'predicted_cancellation' not in df.columns:
        print("  ⚠️ FEHLER: Benötigte Spalten nicht gefunden")
        return {'accuracy': 0, 'auc': 0}

    # Tatsächliche Kündigung und Vorhersage als Integer-Arrays
    y_true = df['event'].astype(int).values
    y_pred = df['predicted_cancellation'].astype(int).values

    # Verteilung der Klassen für Debugging ausgeben
    true_positive_count = sum(1 for t, p in zip(y_true, y_pred) if t == 1 and p == 1)
    true_negative_count = sum(1 for t, p in zip(y_true, y_pred) if t == 0 and p == 0)
    false_positive_count = sum(1 for t, p in zip(y_true, y_pred) if t == 0 and p == 1)
    false_negative_count = sum(1 for t, p in zip(y_true, y_pred) if t == 1 and p == 0)

    print(f"  → Echte Kündigungen: {sum(y_true)} von {len(y_true)} ({sum(y_true) / len(y_true):.2%})")
    print(f"  → Vorhergesagte Kündigungen: {sum(y_pred)} von {len(y_pred)} ({sum(y_pred) / len(y_pred):.2%})")

    # Konfusionsmatrix manuell berechnen und darstellen
    manual_cm = np.array([
        [true_negative_count, false_positive_count],
        [false_negative_count, true_positive_count]
    ])

    print("\nKonfusionsmatrix (manuell berechnet):")
    print(f"[[TN={true_negative_count}, FP={false_positive_count}],")
    print(f" [FN={false_negative_count}, TP={true_positive_count}]]")

    # Auch die scikit-learn Konfusionsmatrix berechnen
    cm = confusion_matrix(y_true, y_pred)
    print("\nKonfusionsmatrix (sklearn):")
    print(cm)

    # Accuracy und AUC berechnen
    acc = accuracy_score(y_true, y_pred)
    print(f"  → Accuracy: {acc:.4f}")

    # Konfusionsmatrix visualisieren mit klaren Labels
    plt.figure(figsize=(10, 8))
    sns.heatmap(manual_cm, annot=True, fmt='d', cmap='Blues',
                xticklabels=['Kein Storno', 'Storno'],
                yticklabels=['Kein Storno', 'Storno'])
    plt.title('Konfusionsmatrix: Stornoprognose', fontsize=16)
    plt.xlabel('Vorhergesagte Klasse', fontsize=14)
    plt.ylabel('Tatsächliche Klasse', fontsize=14)

    # Zellinhalte deutlich anzeigen
    for i in range(2):
        for j in range(2):
            plt.text(j + 0.5, i + 0.5, str(manual_cm[i, j]),
                     ha="center", va="center", fontsize=14,
                     color="white" if i == j else "black")

    plt.tight_layout()
    plt.savefig("confusion_matrix_allgemein.png", dpi=300)
    plt.close()

    return {'accuracy': acc}

# ===== FUNKTION FÜR FRÜH-STORNO-ANALYSE =====
def analyze_early_cancellations(df, days_threshold=30):
    """
    Analysiert und visualisiert frühe Kündigungen innerhalb eines definierten Zeitraums

    Parameter:
        df (pandas.DataFrame): DataFrame mit observation_time und event
        days_threshold (int): Schwellenwert in Tagen für frühe Kündigungen
    """
    print(f"\nAnalysiere frühe Kündigungen (≤ {days_threshold} Tage)...")

    # Frühe Kündigungen identifizieren
    early_cancellations = df[(df['event'] == 1) & (df['observation_time'] <= days_threshold)]
    all_cancellations = df[df['event'] == 1]

    # Statistiken berechnen
    early_count = len(early_cancellations)
    all_count = len(all_cancellations)
    total_count = len(df)

    early_pct_total = early_count / total_count * 100
    early_pct_cancellations = early_count / all_count * 100 if all_count > 0 else 0

    # Ergebnisse ausgeben
    print(
        f"  → {early_count} von {total_count} Verträgen wurden innerhalb der ersten {days_threshold} Tage gekündigt ({early_pct_total:.2f}%)")
    print(f"  → {early_pct_cancellations:.2f}% aller Kündigungen erfolgen in den ersten {days_threshold} Tagen")

    # Histogramm der Kündigungszeiten erstellen
    plt.figure(figsize=(12, 6))
    max_days = min(1000, df[df['event'] == 1]['observation_time'].max())
    bins = [0, 30, 60, 90, 180, 365, 1000, max_days]

    # Gruppiere Kündigungen nach Zeiträumen
    cancellation_times = df[df['event'] == 1]['observation_time']
    hist, edges = np.histogram(cancellation_times, bins=bins)

    # Bereinige Beschriftungen für die Darstellung
    labels = []
    for i in range(len(bins) - 1):
        labels.append(f"{bins[i]}-{bins[i + 1]} Tage")

    # Erstelle Balkendiagramm
    plt.bar(labels, hist, color='salmon')
    plt.title('Verteilung der Kündigungszeitpunkte')
    plt.xlabel('Zeitraum nach Versicherungsbeginn')
    plt.ylabel('Anzahl der Kündigungen')
    plt.xticks(rotation=45)
    plt.grid(axis='y', alpha=0.3)
    plt.tight_layout()
    plt.savefig("early_cancellations_histogram.png")
    plt.close()

    # Erstelle auch Kreisdiagramm zur Veranschaulichung des Anteils
    plt.figure(figsize=(10, 6))
    plt.pie([early_count, all_count - early_count],
            labels=[f'≤ {days_threshold} Tage ({early_pct_cancellations:.1f}%)',
                    f'> {days_threshold} Tage ({100 - early_pct_cancellations:.1f}%)'],
            colors=['salmon', 'lightblue'], autopct='%1.1f%%', startangle=90)
    plt.title(f'Anteil früher Kündigungen (≤ {days_threshold} Tage) an allen Kündigungen')
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig("early_cancellations_pie.png")
    plt.close()

    return {'early_count': early_count, 'early_pct_total': early_pct_total,
            'early_pct_cancellations': early_pct_cancellations}


def evaluate_kaplan_meier_predictions(df_predictions):
    """
    Bewertet die Kaplan-Meier-Vorhersagen mit korrekter Konfusionsmatrix
    """
    print("\nBewerte Kaplan-Meier-Modellqualität...")

    # Sicherstellen, dass die benötigten Spalten vorhanden sind
    if 'event' not in df_predictions.columns or 'predicted_cancellation' not in df_predictions.columns:
        print("  ⚠️ FEHLER: Benötigte Spalten nicht gefunden")
        return {'accuracy': 0, 'auc': 0}

    # Tatsächliche Kündigung über 'event' Spalte
    y_true = df_predictions['event'].astype(int)
    y_pred = df_predictions['predicted_cancellation'].astype(int)
    y_prob = df_predictions['cancellation_probability']

    # Verteilung der Vorhersagen und echten Werte prüfen
    print(f"  → Echte Kündigungen: {y_true.sum()} von {len(y_true)} ({y_true.mean():.2%})")
    print(f"  → Vorhergesagte Kündigungen: {y_pred.sum()} von {len(y_pred)} ({y_pred.mean():.2%})")

    # Konfusionsmatrix korrekt berechnen
    cm = confusion_matrix(y_true, y_pred)

    # Explizite Anzeige der Matrixwerte für Debugging
    print("\nKonfusionsmatrix (Reihenfolge: [[TN, FP], [FN, TP]]):")
    print(cm)

    # Für besseres Verständnis auch die einzelnen Zellen ausgeben
    tn, fp, fn, tp = cm.ravel()
    print(f"  → True Negatives (kein Storno korrekt): {tn}")
    print(f"  → False Positives (fälschlicherweise Storno): {fp}")
    print(f"  → False Negatives (verpasste Stornos): {fn}")
    print(f"  → True Positives (Storno korrekt): {tp}")

    # Klassifikationsbericht
    print("\nKaplan-Meier Klassifikationsbericht:")
    print(classification_report(y_true, y_pred))

    # Accuracy und AUC berechnen
    acc = accuracy_score(y_true, y_pred)
    try:
        auc = roc_auc_score(y_true, y_prob)
        print(f"  → KM-Accuracy: {acc:.4f}, KM-AUC: {auc:.4f}")
    except Exception as e:
        print(f"  ✗ AUC-Berechnung fehlgeschlagen: {e}")
        auc = None

    # Konfusionsmatrix visualisieren mit korrekten Labels
    plt.figure(figsize=(8, 6))
    sns.heatmap(cm, annot=True, fmt='d', cmap='YlGnBu',
                xticklabels=['Kein Storno', 'Storno'],
                yticklabels=['Kein Storno', 'Storno'])
    plt.title('Konfusionsmatrix: Kaplan-Meier-Methode')
    plt.xlabel('Vorhergesagte Klasse')
    plt.ylabel('Tatsächliche Klasse')
    plt.tight_layout()
    plt.savefig("confusion_matrix_kaplan_meier_specific.png")
    plt.close()

    return {'accuracy': acc, 'auc': auc}


def evaluate_cox_predictions(df_predictions):
    """
    Bewertet die Cox-Modell-Vorhersagen mit korrekter Konfusionsmatrix
    """
    print("\nBewerte Cox-Modellqualität...")

    # Sicherstellen, dass die benötigten Spalten vorhanden sind
    if 'event' not in df_predictions.columns or 'predicted_cancellation' not in df_predictions.columns:
        print("  ⚠️ FEHLER: Benötigte Spalten nicht gefunden")
        return {'accuracy': 0, 'auc': 0}

    # Tatsächliche Kündigung über 'event' Spalte
    y_true = df_predictions['event'].astype(int)
    y_pred = df_predictions['predicted_cancellation'].astype(int)
    y_prob = df_predictions['cancellation_probability']

    # Verteilung der Vorhersagen und echten Werte prüfen
    print(f"  → Echte Kündigungen: {y_true.sum()} von {len(y_true)} ({y_true.mean():.2%})")
    print(f"  → Vorhergesagte Kündigungen: {y_pred.sum()} von {len(y_pred)} ({y_pred.mean():.2%})")

    # Konfusionsmatrix korrekt berechnen
    cm = confusion_matrix(y_true, y_pred)

    # Explizite Anzeige der Matrixwerte für Debugging
    print("\nKonfusionsmatrix (Reihenfolge: [[TN, FP], [FN, TP]]):")
    print(cm)

    # Für besseres Verständnis auch die einzelnen Zellen ausgeben
    tn, fp, fn, tp = cm.ravel()
    print(f"  → True Negatives (kein Storno korrekt): {tn}")
    print(f"  → False Positives (fälschlicherweise Storno): {fp}")
    print(f"  → False Negatives (verpasste Stornos): {fn}")
    print(f"  → True Positives (Storno korrekt): {tp}")

    # Klassifikationsbericht
    print("\nCox-Modell Klassifikationsbericht:")
    print(classification_report(y_true, y_pred))

    # Accuracy und AUC berechnen
    acc = accuracy_score(y_true, y_pred)
    try:
        auc = roc_auc_score(y_true, y_prob)
        print(f"  → Cox-Accuracy: {acc:.4f}, Cox-AUC: {auc:.4f}")
    except Exception as e:
        print(f"  ✗ AUC-Berechnung fehlgeschlagen: {e}")
        auc = None

    # Konfusionsmatrix visualisieren mit korrekten Labels
    plt.figure(figsize=(8, 6))
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues',
                xticklabels=['Kein Storno', 'Storno'],
                yticklabels=['Kein Storno', 'Storno'])
    plt.title('Konfusionsmatrix: Cox-Modell')
    plt.xlabel('Vorhergesagte Klasse')
    plt.ylabel('Tatsächliche Klasse')
    plt.tight_layout()
    plt.savefig("confusion_matrix_cox_model.png")
    plt.close()

    return {'accuracy': acc, 'auc': auc}

# ===== HAUPTPROGRAMM =====
try:
    print("\nBeginne Datenverarbeitung...")

    # Excel-Datei laden
    df = pd.read_excel(file_path)
    print(f"Datei erfolgreich geladen: {len(df)} Zeilen, {len(df.columns)} Spalten")
    total_records_loaded = len(df)

    # "Deletion Type" zu Integer konvertieren
    if 'Deletion Type' in df.columns:
        df['Deletion Type'] = pd.to_numeric(df['Deletion Type'], errors='coerce').fillna(0).astype(int)
        print(f"'Deletion Type' zu Integer konvertiert")

        # Verteilung der Löschungstypen
        deletion_counts = df['Deletion Type'].value_counts()
        print("\nVerteilung der Löschungstypen:")
        for dt, count in deletion_counts.items():
            print(f"  Typ {dt}: {count} Datensätze ({count / len(df) * 100:.1f}%)")

    # Datumsspalten konvertieren
    date_columns = ['Start Insurance', 'End Insurance', 'FirstDueDate', 'Deletion allowed at',
                    'Promised Deletion Date', 'Last Due Date', 'Meldedatum']
    df = process_date_columns(df, date_columns)

    # Kategorische Spalten vorbereiten
    categorical_columns = ['Country_Region Code', 'Product Code', 'Kampagne']
    for col in categorical_columns:
        if col in df.columns:
            # Häufigste Werte beibehalten, seltene zu "Other" zusammenfassen
            top_values = df[col].value_counts().head(10).index
            df[col] = df[col].apply(lambda x: x if x in top_values else 'Other')
            print(f"'{col}' kategorisiert: Top-10-Werte beibehalten, Rest als 'Other'")

    # Amount-Spalte behandeln
    if 'Amount' in df.columns:
        df['Amount'] = pd.to_numeric(df['Amount'].astype(str).str.replace(',', '.'), errors='coerce')
        print("'Amount'-Spalte zu numerischen Werten konvertiert")

    # Features für Überlebenszeitanalyse erstellen
    df_survival = create_survival_features(df)

    # Selektiere Features für das Modell
    numerical_features = ['Amount', 'duration', 'observation_time']
    numerical_features = [f for f in numerical_features if f in df_survival.columns and f != 'observation_time']

    # Bereinige numerische Features
    for col in numerical_features:
        if df_survival[col].isna().any():
            median_val = df_survival[col].median()
            df_survival[col] = df_survival[col].fillna(median_val)

    # Train-Test-Split
    print("\nTeile Daten in Trainings- und Testdaten (80/20-Split)...")
    df_train, df_test = train_test_split(df_survival, test_size=0.2, random_state=42, stratify=df_survival['event'])
    print(f"  → Trainingsdaten: {len(df_train)} Datensätze")
    print(f"  → Testdaten: {len(df_test)} Datensätze")

    # Train-Test-Verteilung prüfen
    print(f"  → Trainings-Events: {df_train['event'].mean() * 100:.1f}% Kündigungen")
    print(f"  → Test-Events: {df_test['event'].mean() * 100:.1f}% Kündigungen")

    # Kaplan-Meier-Analyse durchführen
    stratify_columns = ['Country_Region Code', 'Product Code', 'Amount', 'Kampagne']
    km_curves = perform_kaplan_meier_analysis(df_train, stratify_columns)

    early_stats = analyze_early_cancellations(df_train, days_threshold=30)

    # Features für Cox-Modell vorbereiten
    model_features = categorical_columns + numerical_features + ['Deletion Type']
    model_features = [f for f in model_features if f in df_train.columns]

    # Daten für das Modelltraining vorbereiten
    print("\nBereite Daten für Cox-Modell vor...")
    df_train_model = df_train.copy()
    df_test_model = df_test.copy()

    # Kategorische Features als One-Hot-Kodierung
    for col in categorical_columns:
        if col in df_train_model.columns:
            # One-Hot-Kodierung für kategorische Variablen
            dummies_train = pd.get_dummies(df_train_model[col], prefix=col, drop_first=True)
            dummies_test = pd.get_dummies(df_test_model[col], prefix=col, drop_first=True)

            # Sicherstellen, dass Test die gleichen Spalten hat
            for dummy_col in dummies_train.columns:
                if dummy_col not in dummies_test.columns:
                    dummies_test[dummy_col] = 0

            # Nur gemeinsame Spalten behalten
            common_cols = [c for c in dummies_train.columns if c in dummies_test.columns]
            df_train_model = pd.concat([df_train_model, dummies_train[common_cols]], axis=1)
            df_test_model = pd.concat([df_test_model, dummies_test[common_cols]], axis=1)

    # Cox-Modell trainieren
    # Cox-Modell trainieren
    # Cox-Modell trainieren
    cox_model = build_cox_model(df_train_model, list(df_train_model.columns))

    # 1. KAPLAN-MEIER-VORHERSAGEN (unabhängig vom Cox-Modell)
    print("\nErstelle Kaplan-Meier-Vorhersagen für separate Evaluation...")
    km_predictions = predict_cancellation_kaplan_meier(km_curves, df_test, stratify_columns, time_horizon=180)

    # Kaplan-Meier Modell separat evaluieren
    km_metrics = evaluate_kaplan_meier_predictions(km_predictions)
    print("\nKAPLAN-MEIER MODELLLEISTUNG:")
    print(f"KM-Accuracy: {km_metrics['accuracy']:.4f}")
    if km_metrics['auc'] is not None:
        print(f"KM-AUC: {km_metrics['auc']:.4f}")

    # Kaplan-Meier-Ergebnisse speichern
    km_predictions.to_csv("kaplan_meier_predictions.csv", index=False)
    print("Kaplan-Meier-Ergebnisse in 'kaplan_meier_predictions.csv' gespeichert")

    # 2. COX-MODELL-VORHERSAGEN (wenn verfügbar)
    if cox_model is not None:
        print("\nErstelle Cox-Modell-Vorhersagen...")
        cox_predictions = predict_cancellation_cox(cox_model, df_test_model, km_curves, time_horizon=180)

        # Cox-Modell mit der neuen Funktion evaluieren
        cox_metrics = evaluate_cox_predictions(cox_predictions)
        print("\nCOX-MODELL LEISTUNG:")
        print(f"Cox-Accuracy: {cox_metrics['accuracy']:.4f}")
        if cox_metrics['auc'] is not None:
            print(f"Cox-AUC: {cox_metrics['auc']:.4f}")

        # Cox-Ergebnisse speichern
        cox_predictions.to_csv("cox_model_predictions.csv", index=False)
        print("Cox-Modell-Ergebnisse in 'cox_model_predictions.csv' gespeichert")

        # Feature-Wichtigkeit für Cox-Modell ausgeben
        print("\nFeature-Wichtigkeit (Cox-Modell):")
        cox_summary = cox_model.summary[['coef', 'exp(coef)', 'p']].sort_values(by='p')
        print(cox_summary.head(10))  # Top-10 signifikante Features
    else:
        print("\nKein Cox-Modell verfügbar - nur Kaplan-Meier-Ergebnisse wurden erzeugt")

    # Zusammenfassung ausgeben
    print("\n" + "-" * 60)
    print(f"MODELLVERGLEICH:")
    print(f"Kaplan-Meier-Accuracy: {km_metrics['accuracy']:.4f}")
    if cox_model is not None:
        print(f"Cox-Modell-Accuracy: {cox_metrics['accuracy']:.4f}")
        print(f"Accuracy-Differenz: {abs(km_metrics['accuracy'] - cox_metrics['accuracy']):.4f}")
    print("-" * 60)

    print("\n" + "-" * 60)
    print(f"DATENSATZ-ZUSAMMENFASSUNG:")
    print(f"Geladene Datensätze gesamt: {total_records_loaded}")
    print(f"Für Überlebenszeitanalyse verwendete Datensätze: {len(df_survival)}")
    print(f"Davon für Training verwendet: {len(df_train)} ({len(df_train) / len(df_survival) * 100:.1f}%)")
    print(f"Davon für Test verwendet: {len(df_test)} ({len(df_test) / len(df_survival) * 100:.1f}%)")
    print(
        f"Aktive Verträge: {(df_survival['event'] == 0).sum()} ({(df_survival['event'] == 0).sum() / len(df_survival) * 100:.1f}%)")
    print(
        f"Gekündigte Verträge: {(df_survival['event'] == 1).sum()} ({(df_survival['event'] == 1).sum() / len(df_survival) * 100:.1f}%)")
    print("-" * 60)

except FileNotFoundError:
    print(f"Fehler: Die Datei '{file_path}' wurde nicht gefunden.")
except pd.errors.EmptyDataError:
    print(f"Fehler: Die Datei '{file_path}' ist leer.")
except Exception as e:
    print(f"Ein unerwarteter Fehler ist aufgetreten: {str(e)}")
    import traceback

    traceback.print_exc()


print("\n" + "=" * 80)
print("ANALYSE ABGESCHLOSSEN")
print("=" * 80)


