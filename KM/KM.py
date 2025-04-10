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
def create_survival_features(df):
    """
    Erstellt spezielle Features für die Überlebenszeitanalyse

    Parameter:
        df (pandas.DataFrame): DataFrame mit den Versicherungsdaten

    Rückgabe:
        pandas.DataFrame: DataFrame mit zusätzlichen Features für Überlebenszeitanalyse
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

    # Beobachtungszeitraum (bis heute oder bis Kündigung)
    today = datetime.datetime.now()

    # Verträge ohne Kündigung: Zeit von Start bis heute
    mask_active = (df_survival['event'] == 0) & (df_survival['Start Insurance'].notna())
    df_survival.loc[mask_active, 'observation_time'] = (today - df_survival.loc[mask_active, 'Start Insurance']).dt.days

    # Gekündigte Verträge: Zeit von Start bis Kündigung
    for end_col in ['End Insurance', 'Last Due Date']:
        if end_col in df_survival.columns:
            mask_cancelled = (df_survival['event'] == 1) & (df_survival['Start Insurance'].notna()) & (
                df_survival[end_col].notna())
            df_survival.loc[mask_cancelled, 'observation_time'] = (df_survival.loc[mask_cancelled, end_col] -
                                                                   df_survival.loc[
                                                                       mask_cancelled, 'Start Insurance']).dt.days
            if mask_cancelled.sum() > 0:
                print(f"  → Feature 'observation_time' erstellt mit {end_col} für {mask_cancelled.sum()} Datensätze")
                break

    # Fehlende Werte auffüllen und negative Werte vermeiden
    df_survival['observation_time'] = df_survival['observation_time'].fillna(-1).clip(lower=0)

    # Verträge ohne Startdatum aussortieren (falls nötig)
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


# ===== FUNKTION FÜR COX-MODELL =====
def build_cox_model(df, features):
    """
    Erstellt ein Cox Proportional Hazards Modell

    Parameter:
        df (pandas.DataFrame): DataFrame mit event und observation_time Spalten
        features (list): Liste mit Features für das Modell

    Rückgabe:
        CoxPHFitter: Trainiertes Cox-Modell
    """
    print("\nTrainiere Cox Proportional Hazards Modell...")

    # Nur relevante Spalten auswählen
    model_df = df[features + ['observation_time', 'event']].copy()

    # Fehlende Werte behandeln
    for col in features:
        if model_df[col].isna().any():
            model_df[col] = model_df[col].fillna(model_df[col].median())

    # NaN-Werte entfernen (falls noch vorhanden)
    model_df = model_df.dropna()
    print(f"  → {len(model_df)} Datensätze für Cox-Modelltraining verwendet")

    # Cox-Modell trainieren
    cph = CoxPHFitter()
    try:
        cph.fit(model_df, duration_col='observation_time', event_col='event')

        # Modellzusammenfassung
        print("\nCox-Modell Summary:")
        print(cph.summary[['coef', 'exp(coef)', 'p']])

        # Visualisierung der Koeffizienten
        plt.figure(figsize=(12, 8))
        cph.plot()
        plt.title('Cox-Modell: Koeffizienten und Konfidenzintervalle')
        plt.tight_layout()
        plt.savefig("cox_model_coefficients.png")
        plt.close()

        return cph
    except Exception as e:
        print(f"  ✗ Fehler beim Training des Cox-Modells: {e}")
        return None


# ===== FUNKTION FÜR STORNO-VORHERSAGE =====
def predict_cancellation(model, df_test, features, time_horizon=180):
    """
    Sagt Kündigungswahrscheinlichkeit für einen bestimmten Zeitraum voraus

    Parameter:
        model: Trainiertes Cox-Modell
        df_test (pandas.DataFrame): Test-DataFrame
        features (list): Features für die Vorhersage
        time_horizon (int): Zeithorizont für Prognose in Tagen

    Rückgabe:
        pandas.DataFrame: DataFrame mit Kündigungswahrscheinlichkeiten
    """
    print(f"\nErstelle Stornoprognose für Zeithorizont: {time_horizon} Tage...")

    # Nur relevante Spalten auswählen
    pred_df = df_test[features].copy()

    # Fehlende Werte behandeln
    for col in features:
        if pred_df[col].isna().any():
            pred_df[col] = pred_df[col].fillna(pred_df[col].median())

    # Wahrscheinlichkeiten vorhersagen
    try:
        # Überlebenswahrscheinlichkeit vorhersagen
        survival_prob = model.predict_survival_function(pred_df)

        # Zeithorizont auswählen (oder nächstgelegenen Zeitpunkt)
        available_times = survival_prob.index
        closest_time = min(available_times, key=lambda x: abs(x - time_horizon))

        # Kündigungswahrscheinlichkeit = 1 - Überlebenswahrscheinlichkeit
        cancellation_prob = 1 - survival_prob.loc[closest_time].values

        # Ergebnis-DataFrame erstellen
        result_df = df_test.copy()
        result_df['cancellation_probability'] = cancellation_prob

        # Binäre Vorhersage mit optimiertem Threshold
        threshold = 0.5  # Standardwert, kann optimiert werden
        result_df['predicted_cancellation'] = (result_df['cancellation_probability'] > threshold).astype(int)

        print(f"  → Stornoprognose für {len(result_df)} Datensätze erstellt")
        return result_df

    except Exception as e:
        print(f"  ✗ Fehler bei der Vorhersage: {e}")
        return df_test


# ===== FUNKTION FÜR MODELLBEWERTUNG =====
def evaluate_model(df):
    """
    Bewertet die Vorhersagequalität des Modells

    Parameter:
        df (pandas.DataFrame): DataFrame mit tatsächlichen und vorhergesagten Werten

    Rückgabe:
        dict: Dictionary mit Performance-Metriken
    """
    print("\nBewerte Modellqualität...")

    # Tatsächliche Kündigung (1 wenn Deletion Type != 0)
    y_true = (df['Deletion Type'] != 0).astype(int)
    y_pred = df['predicted_cancellation'].astype(int)
    y_prob = df['cancellation_probability']

    # Klasssifikationsbericht
    print("\nKlassifikationsbericht:")
    print(classification_report(y_true, y_pred))

    # Konfusionsmatrix erstellen
    cm = confusion_matrix(y_true, y_pred)

    # Accuracy und AUC berechnen
    acc = accuracy_score(y_true, y_pred)
    try:
        auc = roc_auc_score(y_true, y_prob)
        print(f"Accuracy: {acc:.4f}, AUC: {auc:.4f}")
    except:
        print(f"Accuracy: {acc:.4f}, AUC nicht berechenbar")
        auc = None

    # Konfusionsmatrix visualisieren
    plt.figure(figsize=(8, 6))
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues',
                xticklabels=['Kein Storno', 'Storno'],
                yticklabels=['Kein Storno', 'Storno'])
    plt.title('Konfusionsmatrix: Stornoprognose')
    plt.xlabel('Vorhergesagte Klasse')
    plt.ylabel('Tatsächliche Klasse')
    plt.tight_layout()
    plt.savefig("confusion_matrix_kaplan_meier.png")
    plt.close()

    # ROC-Kurve darstellen (falls AUC berechenbar)
    if auc is not None:
        plt.figure(figsize=(8, 6))
        fpr, tpr, _ = roc_curve(y_true, y_prob)
        plt.plot(fpr, tpr, label=f'AUC = {auc:.4f}')
        plt.plot([0, 1], [0, 1], 'k--')  # Diagonallinie
        plt.xlabel('False Positive Rate')
        plt.ylabel('True Positive Rate')
        plt.title('ROC-Kurve: Stornoprognose')
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        plt.savefig("roc_curve_kaplan_meier.png")
        plt.close()

    # Verteilung der Wahrscheinlichkeiten nach tatsächlichem Ergebnis
    plt.figure(figsize=(10, 6))

    # Getrennte Darstellung für gekündigte und nicht gekündigte Verträge
    df_cancelled = df[y_true == 1].copy()
    df_active = df[y_true == 0].copy()

    sns.kdeplot(df_active['cancellation_probability'], label='Aktive Verträge', color='green', fill=True)
    sns.kdeplot(df_cancelled['cancellation_probability'], label='Gekündigte Verträge', color='red', fill=True)

    plt.axvline(0.5, color='black', linestyle='--', label='Threshold (0.5)')
    plt.title('Verteilung der Stornowahrscheinlichkeiten')
    plt.xlabel('Vorhergesagte Kündigungswahrscheinlichkeit')
    plt.ylabel('Dichte')
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.savefig("probability_distribution_kaplan_meier.png")
    plt.close()

    return {'accuracy': acc, 'auc': auc}


# ===== HAUPTPROGRAMM =====
try:
    print("\nBeginne Datenverarbeitung...")

    # Excel-Datei laden
    df = pd.read_excel(file_path)
    print(f"Datei erfolgreich geladen: {len(df)} Zeilen, {len(df.columns)} Spalten")

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
    stratify_columns = ['Country_Region Code', 'Product Code']
    km_curves = perform_kaplan_meier_analysis(df_train, stratify_columns)

    # Features für Cox-Modell vorbereiten
    model_features = categorical_columns + numerical_features
    model_features = [f for f in model_features if f in df_train.columns]

    # Dummy-Variablen für kategorische Features erstellen
    df_train_model = pd.get_dummies(df_train, columns=categorical_columns, drop_first=True)
    df_test_model = pd.get_dummies(df_test, columns=categorical_columns, drop_first=True)

    # Sicherstellen, dass Test-Set die gleichen Spalten hat
    for col in df_train_model.columns:
        if col not in df_test_model.columns and col not in ['observation_time', 'event']:
            df_test_model[col] = 0

    # Cox-Modell trainieren
    model_features = [col for col in df_train_model.columns
                      if col not in ['observation_time', 'event', 'Deletion Type']]
    cox_model = build_cox_model(df_train_model, model_features)

    if cox_model is not None:
        # Vorhersagen auf Testdaten
        df_predictions = predict_cancellation(cox_model, df_test_model, model_features)

        # Modell bewerten
        metrics = evaluate_model(df_predictions)

        # Ergebnisse speichern
        df_predictions.to_csv("kaplan_meier_predictions.csv", index=False)
        print("\nErgebnisse in 'kaplan_meier_predictions.csv' gespeichert")

        # Modell speichern
        import joblib

        joblib.dump(cox_model, 'cox_model.pkl')
        print("Cox-Modell in 'cox_model.pkl' gespeichert")

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

