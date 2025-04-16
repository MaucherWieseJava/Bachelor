# =====================================================================
# AUTOR: @Adrian Stötzler
# TITEL: ML-basierte Stornoprognosetools
# BESCHREIBUNG: Dieses Skript implementiert verschiedene ML-Algorithmen für die
# Stornoprognose und vergleicht deren Performance mit Kaplan-Meier und Cox.
# =====================================================================

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
from sklearn.metrics import (classification_report, confusion_matrix, accuracy_score,
                             roc_auc_score, roc_curve, precision_recall_curve, average_precision_score)

# ML-Algorithmen
from sklearn.ensemble import GradientBoostingClassifier, RandomForestClassifier
from sklearn.linear_model import LogisticRegression
import xgboost as xgb
import lightgbm as lgbm

# Warnungen unterdrücken
import warnings

warnings.filterwarnings('ignore')


class MLStornoPredictionTool:
    """
    Klasse zur Implementierung verschiedener ML-Algorithmen für die Stornoprognose
    und Vergleich mit Kaplan-Meier und Cox-Modellen
    """

    def __init__(self, file_path=None):
        """
        Initialisiert das ML-Tool mit dem angegebenen Dateipfad

        Parameter:
            file_path (str): Pfad zur Excel-Datei mit Versicherungsdaten
        """
        self.file_path = file_path or os.path.join(os.environ["HOME"], "Desktop", "Training.xlsx")
        self.models = {}
        self.results = {}
        self.df = None
        self.df_processed = None
        self.X_train = None
        self.X_test = None
        self.y_train = None
        self.y_test = None

        print("=" * 80)
        print("ML-BASIERTE STORNOPROGNOSE")
        print("=" * 80)

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

        except FileNotFoundError:
            print(f"Fehler: Die Datei '{self.file_path}' wurde nicht gefunden.")
            return False
        except Exception as e:
            print(f"Ein unerwarteter Fehler ist aufgetreten: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def preprocess_data(self):
        """
        Verarbeitet die Daten für ML-Modelle (Feature-Engineering, Kodierung, etc.)
        """
        print("\nBereite Daten für ML-Modelle vor...")

        try:
            # Kopie erstellen
            df_prep = self.df.copy()

            # Datum-Spalten konvertieren
            date_columns = ['Start Insurance', 'End Insurance', 'FirstDueDate', 'Deletion allowed at',
                            'Promised Deletion Date', 'Last Due Date', 'Meldedatum']

            for col in date_columns:
                if col in df_prep.columns:
                    df_prep[col] = pd.to_datetime(df_prep[col], errors='coerce')
                    print(f"  ✓ '{col}' zu Datum konvertiert")

            # Feature-Engineering für Datumsspalten
            if 'Start Insurance' in df_prep.columns:
                today = pd.Timestamp.now()

                # Vertragsdauer berechnen
                if 'End Insurance' in df_prep.columns:
                    df_prep['contract_duration'] = (df_prep['End Insurance'] - df_prep['Start Insurance']).dt.days
                    df_prep['contract_duration'] = df_prep['contract_duration'].fillna(-1).clip(lower=0)

                # Zeit seit Vertragsbeginn
                df_prep['days_from_start'] = (today - df_prep['Start Insurance']).dt.days
                df_prep['days_from_start'] = df_prep['days_from_start'].fillna(-1).clip(lower=0)

                # Zeit bis Vertrag erlaubt gekündigt werden kann
                if 'Deletion allowed at' in df_prep.columns:
                    df_prep['days_until_deletion_allowed'] = (
                                df_prep['Deletion allowed at'] - df_prep['Start Insurance']).dt.days
                    df_prep['days_until_deletion_allowed'] = df_prep['days_until_deletion_allowed'].fillna(-1)

            # "Amount" Spalte behandeln
            if 'Amount' in df_prep.columns:
                df_prep['Amount'] = pd.to_numeric(df_prep['Amount'].astype(str).str.replace(',', '.'), errors='coerce')
                df_prep['Amount'] = df_prep['Amount'].fillna(df_prep['Amount'].median())

            # Zielvariable erstellen: Kündigung (1) oder nicht (0)
            df_prep['target'] = (df_prep['Deletion Type'] != 0).astype(int)
            print(f"  → Zielvariable 'target' erstellt: {df_prep['target'].value_counts().to_dict()}")


            # Dummy-Variablen für kategoriale Features erstellen
            categorical_columns = ['Country_Region Code', 'Product Code', 'Kampagne']
            for col in categorical_columns:  # Änderung hier: Verwendung der lokalen Variable
                if col in df_prep.columns:
                    df_prep[col] = df_prep[col].astype(str)
                    print(f"  → '{col}' zu einheitlichem String-Format konvertiert")

            # Features für das Modell auswählen
            selected_features = []

            # Numerische Features
            numeric_cols = ['Amount', 'contract_duration', 'days_from_start']
            selected_features.extend([col for col in numeric_cols if col in df_prep.columns])

            # Kategoriale Features
            cat_cols = [col for col in categorical_columns if col in df_prep.columns]
            selected_features.extend(cat_cols)

            # Features und Zielvariable extrahieren
            X = df_prep[selected_features].copy()
            y = df_prep['target'].copy()

            # In Trainings- und Testdaten aufteilen (Stratifizieren nach Zielvariable)
            X_train, X_test, y_train, y_test = train_test_split(
                X, y, test_size=0.2, random_state=42, stratify=y
            )

            # Speichern für spätere Verwendung
            self.X_train = X_train
            self.X_test = X_test
            self.y_train = y_train
            self.y_test = y_test

            # Informationen über Split ausgeben
            print(f"  → Daten in {len(X_train)} Trainings- und {len(X_test)} Testdatensätze aufgeteilt")
            print(f"  → Trainings-Events: {y_train.mean() * 100:.1f}% Kündigungen")
            print(f"  → Test-Events: {y_test.mean() * 100:.1f}% Kündigungen")
            print(f"  → {len(selected_features)} Features für ML-Modelle ausgewählt: {', '.join(selected_features)}")

            # Feature-Sets für Modelltraining speichern
            self.numeric_features = [col for col in numeric_cols if col in df_prep.columns]
            self.categorical_features = cat_cols

            return True

        except Exception as e:
            print(f"Fehler bei der Datenvorverarbeitung: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def build_preprocessor(self):
        """
        Erstellt einen Präprozessor für ML-Pipeline (Behandlung von kategorischen und numerischen Features)
        """
        # Numerische Features: Fehlende Werte auffüllen und skalieren
        numeric_transformer = Pipeline(steps=[
            ('imputer', SimpleImputer(strategy='median')),
            ('scaler', StandardScaler())
        ])

        # Kategoriale Features: Fehlende Werte auffüllen und One-Hot-Encoding
        categorical_transformer = Pipeline(steps=[
            ('imputer', SimpleImputer(strategy='constant', fill_value='missing')),
            ('onehot', OneHotEncoder(handle_unknown='ignore', sparse_output=False))
        ])

        # Kombination der Transformationen
        preprocessor = ColumnTransformer(
            transformers=[
                ('num', numeric_transformer, self.numeric_features),
                ('cat', categorical_transformer, self.categorical_features)
            ])

        return preprocessor

    def train_gradient_boosting(self):
        """
        Trainiert ein Gradient Boosting-Modell für die Stornoprognose
        """
        print("\n" + "-" * 60)
        print("GRADIENT BOOSTING MODELLTRAINING")

        # Präprozessor erstellen
        preprocessor = self.build_preprocessor()

        # Gradient Boosting-Modell definieren
        gb_model = GradientBoostingClassifier(
            n_estimators=100,
            learning_rate=0.1,
            max_depth=3,
            random_state=42
        )

        # Pipeline erstellen
        pipeline = Pipeline(steps=[
            ('preprocessor', preprocessor),
            ('classifier', gb_model)
        ])

        # Modell trainieren
        print("Training des Gradient Boosting-Modells...")
        pipeline.fit(self.X_train, self.y_train)

        # Modell für spätere Verwendung speichern
        self.models['gradient_boosting'] = pipeline

        # Prognosen erstellen
        y_pred = pipeline.predict(self.X_test)
        y_prob = pipeline.predict_proba(self.X_test)[:, 1]

        # Ergebnisse evaluieren und speichern
        results = self.evaluate_model(y_pred, y_prob, "Gradient Boosting")
        self.results['gradient_boosting'] = results

        return results

    def train_random_forest(self):
        """
        Trainiert ein Random Forest-Modell für die Stornoprognose
        """
        print("\n" + "-" * 60)
        print("RANDOM FOREST MODELLTRAINING")

        # Präprozessor erstellen
        preprocessor = self.build_preprocessor()

        # Random Forest-Modell definieren
        rf_model = RandomForestClassifier(
            n_estimators=100,
            max_depth=5,
            random_state=42
        )

        # Pipeline erstellen
        pipeline = Pipeline(steps=[
            ('preprocessor', preprocessor),
            ('classifier', rf_model)
        ])

        # Modell trainieren
        print("Training des Random Forest-Modells...")
        pipeline.fit(self.X_train, self.y_train)

        # Modell für spätere Verwendung speichern
        self.models['random_forest'] = pipeline

        # Prognosen erstellen
        y_pred = pipeline.predict(self.X_test)
        y_prob = pipeline.predict_proba(self.X_test)[:, 1]

        # Ergebnisse evaluieren und speichern
        results = self.evaluate_model(y_pred, y_prob, "Random Forest")
        self.results['random_forest'] = results

        # Feature-Wichtigkeit
        if hasattr(pipeline['classifier'], 'feature_importances_'):
            feature_names = []
            if preprocessor.transformers_[0][2]:  # Numerische Features
                feature_names.extend(preprocessor.transformers_[0][2])
            if preprocessor.transformers_[1][2]:  # Kategoriale Features
                one_hot_encoder = pipeline['preprocessor'].transformers_[1][1]['onehot']
                cat_features = []
                for i, col in enumerate(preprocessor.transformers_[1][2]):
                    categories = one_hot_encoder.categories_[i]
                    cat_features.extend([f"{col}_{cat}" for cat in categories])
                feature_names.extend(cat_features)

            # Top 10 wichtigsten Features anzeigen (falls vorhanden)
            if len(feature_names) > 0:
                try:
                    importances = pipeline['classifier'].feature_importances_
                    indices = np.argsort(importances)[-10:]  # Top 10
                    plt.figure(figsize=(10, 6))
                    plt.title('Random Forest: Top 10 Feature-Wichtigkeit')
                    plt.barh(range(len(indices)), importances[indices], align='center')
                    plt.yticks(range(len(indices)),
                               [feature_names[i] if i < len(feature_names) else f"Feature {i}" for i in indices])
                    plt.tight_layout()
                    plt.savefig("random_forest_feature_importance.png")
                    plt.close()
                    print("  → Feature-Wichtigkeit für Random Forest gespeichert")
                except Exception as e:
                    print(f"  ✗ Fehler bei Feature-Wichtigkeit-Visualisierung: {e}")

        return results

    def train_logistic_regression(self):
        """
        Trainiert ein logistisches Regressionsmodell für die Stornoprognose
        """
        print("\n" + "-" * 60)
        print("LOGISTISCHE REGRESSION MODELLTRAINING")

        # Präprozessor erstellen
        preprocessor = self.build_preprocessor()

        # Logistische Regression definieren
        lr_model = LogisticRegression(
            C=1.0,
            class_weight='balanced',
            max_iter=1000,
            random_state=42
        )

        # Pipeline erstellen
        pipeline = Pipeline(steps=[
            ('preprocessor', preprocessor),
            ('classifier', lr_model)
        ])

        # Modell trainieren
        print("Training des logistischen Regressionsmodells...")
        pipeline.fit(self.X_train, self.y_train)

        # Modell für spätere Verwendung speichern
        self.models['logistic_regression'] = pipeline

        # Prognosen erstellen
        y_pred = pipeline.predict(self.X_test)
        y_prob = pipeline.predict_proba(self.X_test)[:, 1]

        # Ergebnisse evaluieren und speichern
        results = self.evaluate_model(y_pred, y_prob, "Logistische Regression")
        self.results['logistic_regression'] = results

        return results

    def train_xgboost(self):
        """
        Trainiert ein XGBoost-Modell für die Stornoprognose
        """
        print("\n" + "-" * 60)
        print("XGBOOST MODELLTRAINING")

        # Präprozessor erstellen
        preprocessor = self.build_preprocessor()

        # XGBoost-Modell definieren
        xgb_model = xgb.XGBClassifier(
            n_estimators=100,
            learning_rate=0.1,
            max_depth=4,
            random_state=42,
            use_label_encoder=False,
            eval_metric='logloss'
        )

        # Pipeline erstellen
        pipeline = Pipeline(steps=[
            ('preprocessor', preprocessor),
            ('classifier', xgb_model)
        ])

        # Modell trainieren
        print("Training des XGBoost-Modells...")
        pipeline.fit(self.X_train, self.y_train)

        # Modell für spätere Verwendung speichern
        self.models['xgboost'] = pipeline

        # Prognosen erstellen
        y_pred = pipeline.predict(self.X_test)
        y_prob = pipeline.predict_proba(self.X_test)[:, 1]

        # Ergebnisse evaluieren und speichern
        results = self.evaluate_model(y_pred, y_prob, "XGBoost")
        self.results['xgboost'] = results

        return results

    def train_lightgbm(self):
        """
        Trainiert ein LightGBM-Modell für die Stornoprognose
        """
        print("\n" + "-" * 60)
        print("LIGHTGBM MODELLTRAINING")

        # Präprozessor erstellen
        preprocessor = self.build_preprocessor()

        # LightGBM-Modell definieren
        lgbm_model = lgbm.LGBMClassifier(
            n_estimators=100,
            learning_rate=0.1,
            max_depth=5,
            num_leaves=31,
            random_state=42
        )

        # Pipeline erstellen
        pipeline = Pipeline(steps=[
            ('preprocessor', preprocessor),
            ('classifier', lgbm_model)
        ])

        # Modell trainieren
        print("Training des LightGBM-Modells...")
        pipeline.fit(self.X_train, self.y_train)

        # Modell für spätere Verwendung speichern
        self.models['lightgbm'] = pipeline

        # Prognosen erstellen
        y_pred = pipeline.predict(self.X_test)
        y_prob = pipeline.predict_proba(self.X_test)[:, 1]

        # Ergebnisse evaluieren und speichern
        results = self.evaluate_model(y_pred, y_prob, "LightGBM")
        self.results['lightgbm'] = results

        return results

    def evaluate_model(self, y_pred, y_prob, model_name):
        """
        Evaluiert die Modellperformance und erstellt Visualisierungen

        Parameter:
            y_pred (array): Vorhergesagte Klassen
            y_prob (array): Vorhergesagte Wahrscheinlichkeiten
            model_name (str): Name des Modells für Ausgaben/Dateinamen

        Rückgabe:
            dict: Dictionary mit Performance-Metriken
        """
        print(f"\n{model_name} MODELLLEISTUNG:")

        # Performance-Metriken berechnen
        acc = accuracy_score(self.y_test, y_pred)
        try:
            auc = roc_auc_score(self.y_test, y_prob)
            ap = average_precision_score(self.y_test, y_prob)
        except:
            auc = None
            ap = None

        # Klassifikationsbericht erstellen
        print("\nKlassifikationsbericht:")
        print(classification_report(self.y_test, y_pred))

        # Konfusionsmatrix erstellen
        cm = confusion_matrix(self.y_test, y_pred)

        # Accuracy und AUC ausgeben
        print(f"Accuracy: {acc:.4f}")
        if auc is not None:
            print(f"AUC: {auc:.4f}")
        if ap is not None:
            print(f"Average Precision: {ap:.4f}")

        # Konfusionsmatrix visualisieren
        plt.figure(figsize=(8, 6))
        sns.heatmap(cm, annot=True, fmt='d', cmap='Blues',
                    xticklabels=['Kein Storno', 'Storno'],
                    yticklabels=['Kein Storno', 'Storno'])
        plt.title(f'Konfusionsmatrix: Stornoprognose mit {model_name}')
        plt.xlabel('Vorhergesagte Klasse')
        plt.ylabel('Tatsächliche Klasse')
        plt.tight_layout()
        plt.savefig(f"confusion_matrix_{model_name.lower().replace(' ', '_')}.png")
        plt.close()

        # ROC-Kurve darstellen (falls AUC berechenbar)
        if auc is not None:
            plt.figure(figsize=(8, 6))
            fpr, tpr, _ = roc_curve(self.y_test, y_prob)
            plt.plot(fpr, tpr, label=f'AUC = {auc:.4f}')
            plt.plot([0, 1], [0, 1], 'k--')  # Diagonallinie
            plt.xlabel('False Positive Rate')
            plt.ylabel('True Positive Rate')
            plt.title(f'ROC-Kurve: {model_name}')
            plt.legend()
            plt.grid(True)
            plt.tight_layout()
            plt.savefig(f"roc_curve_{model_name.lower().replace(' ', '_')}.png")
            plt.close()

        # Precision-Recall-Kurve
        if ap is not None:
            plt.figure(figsize=(8, 6))
            precision, recall, _ = precision_recall_curve(self.y_test, y_prob)
            plt.plot(recall, precision, label=f'AP = {ap:.4f}')
            plt.xlabel('Recall')
            plt.ylabel('Precision')
            plt.title(f'Precision-Recall-Kurve: {model_name}')
            plt.legend()
            plt.grid(True)
            plt.tight_layout()
            plt.savefig(f"pr_curve_{model_name.lower().replace(' ', '_')}.png")
            plt.close()

        # Verteilung der Wahrscheinlichkeiten nach tatsächlichem Ergebnis
        plt.figure(figsize=(10, 6))

        # Getrennte Darstellung für gekündigte und nicht gekündigte Verträge
        df_results = pd.DataFrame({
            'y_true': self.y_test,
            'y_prob': y_prob
        })

        df_cancelled = df_results[df_results['y_true'] == 1].copy()
        df_active = df_results[df_results['y_true'] == 0].copy()

        sns.kdeplot(df_active['y_prob'], label='Aktive Verträge', color='green', fill=True)
        sns.kdeplot(df_cancelled['y_prob'], label='Gekündigte Verträge', color='red', fill=True)

        plt.axvline(0.5, color='black', linestyle='--', label='Threshold (0.5)')
        plt.title(f'Verteilung der Stornowahrscheinlichkeiten ({model_name})')
        plt.xlabel('Vorhergesagte Kündigungswahrscheinlichkeit')
        plt.ylabel('Dichte')
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        plt.savefig(f"probability_distribution_{model_name.lower().replace(' ', '_')}.png")
        plt.close()

        # Ergebnisstatistiken zurückgeben
        return {
            'accuracy': acc,
            'auc': auc,
            'ap': ap,
            'confusion_matrix': cm
        }

    def compare_models(self):
        """
        Vergleicht die Performance aller trainierten Modelle
        """
        if not self.results:
            print("Keine Modelle zum Vergleichen verfügbar")
            return

        print("\n" + "=" * 60)
        print("MODELLVERGLEICH")
        print("=" * 60)

        # Performance-Tabelle erstellen
        comparison_data = []
        for model_name, results in self.results.items():
            row = {
                'Modell': model_name,
                'Accuracy': results.get('accuracy', 0),
                'AUC': results.get('auc', 0),
                'AP': results.get('ap', 0)
            }
            comparison_data.append(row)

        # DataFrame erstellen und nach Accuracy sortieren
        comparison_df = pd.DataFrame(comparison_data).sort_values(by='Accuracy', ascending=False)
        print(comparison_df.to_string(index=False))

        # Accuracy-Vergleichsdiagramm
        plt.figure(figsize=(12, 6))
        ax = sns.barplot(x='Modell', y='Accuracy', data=comparison_df)
        plt.title('Accuracy-Vergleich der verschiedenen ML-Modelle')
        plt.ylabel('Accuracy')
        plt.xlabel('Modell')
        plt.ylim(0.5, 1.0)  # Skala bei 0.5 beginnen für bessere Visualisierung
        plt.xticks(rotation=45)

        # Werte über den Balken anzeigen
        for i, v in enumerate(comparison_df['Accuracy']):
            ax.text(i, v + 0.01, f"{v:.4f}", ha='center')

        plt.tight_layout()
        plt.savefig("model_comparison_accuracy.png")
        plt.close()

        # AUC-Vergleichsdiagramm
        plt.figure(figsize=(12, 6))
        ax = sns.barplot(x='Modell', y='AUC', data=comparison_df)
        plt.title('AUC-Vergleich der verschiedenen ML-Modelle')
        plt.ylabel('AUC')
        plt.xlabel('Modell')
        plt.ylim(0.5, 1.0)  # Skala bei 0.5 beginnen für bessere Visualisierung
        plt.xticks(rotation=45)

        # Werte über den Balken anzeigen
        for i, v in enumerate(comparison_df['AUC']):
            ax.text(i, v + 0.01, f"{v:.4f}", ha='center')

        plt.tight_layout()
        plt.savefig("model_comparison_auc.png")
        plt.close()

        # Zusammenfassung
        best_model = comparison_df.iloc[0]['Modell']
        best_accuracy = comparison_df.iloc[0]['Accuracy']

        print("\n" + "-" * 60)
        print(f"BESTE PERFORMANCE: {best_model} mit Accuracy von {best_accuracy:.4f}")
        print("-" * 60)

    def train_neural_network(self):
        """
        Trainiert ein neuronales Netz für die Stornoprognose
        """
        print("\n" + "-" * 60)
        print("DEEP LEARNING MODELLTRAINING")

        # Präprozessor erstellen
        preprocessor = self.build_preprocessor()

        # Daten vorbereiten
        X_train_processed = preprocessor.fit_transform(self.X_train)
        X_test_processed = preprocessor.transform(self.X_test)

        # Keras-Modell definieren
        from tensorflow import keras

        # Feature-Dimension bestimmen
        input_dim = X_train_processed.shape[1]

        model = keras.Sequential([
            keras.layers.Dense(64, activation='relu', input_dim=input_dim),
            keras.layers.Dropout(0.2),
            keras.layers.Dense(32, activation='relu'),
            keras.layers.Dropout(0.2),
            keras.layers.Dense(1, activation='sigmoid')
        ])

        # Modell kompilieren
        model.compile(
            optimizer='adam',
            loss='binary_crossentropy',
            metrics=['accuracy', keras.metrics.AUC()]
        )

        # Early Stopping zur Vermeidung von Overfitting
        early_stopping = keras.callbacks.EarlyStopping(
            monitor='val_loss', patience=10, restore_best_weights=True
        )

        # Modell trainieren
        print("Training des neuronalen Netzes...")
        history = model.fit(
            X_train_processed, self.y_train,
            epochs=100,
            batch_size=32,
            validation_split=0.2,
            callbacks=[early_stopping],
            verbose=1
        )

        # Modell und Präprozessor speichern
        self.models['neural_network'] = {
            'preprocessor': preprocessor,
            'model': model
        }

        # Prognosen erstellen
        y_pred_prob = model.predict(X_test_processed)
        y_pred = (y_pred_prob > 0.5).astype(int).flatten()

        # Ergebnisse evaluieren
        results = self.evaluate_model(y_pred, y_pred_prob.flatten(), "Neural Network")
        self.results['neural_network'] = results

        # Lernkurve plotten
        plt.figure(figsize=(12, 4))
        plt.subplot(1, 2, 1)
        plt.plot(history.history['loss'], label='Training Loss')
        plt.plot(history.history['val_loss'], label='Validation Loss')
        plt.title('Lernkurve: Loss')
        plt.xlabel('Epoch')
        plt.ylabel('Loss')
        plt.legend()

        plt.subplot(1, 2, 2)
        plt.plot(history.history['accuracy'], label='Training Accuracy')
        plt.plot(history.history['val_accuracy'], label='Validation Accuracy')
        plt.title('Lernkurve: Accuracy')
        plt.xlabel('Epoch')
        plt.ylabel('Accuracy')
        plt.legend()

        plt.tight_layout()
        plt.savefig("neural_network_learning_curve.png")
        plt.close()

        return results

    def train_lstm_model(self):
        """Trainiert ein LSTM-Modell für zeitliche Muster in Stornodaten"""
        from tensorflow import keras

        # Daten vorbereiten (mit Zeitfenstern)
        preprocessor = self.build_preprocessor()
        X_train_processed = preprocessor.fit_transform(self.X_train)
        X_test_processed = preprocessor.transform(self.X_test)

        # Reshapen für LSTM [samples, timesteps, features]
        # Annahme: wir simulieren timesteps durch Gruppierung von Features
        feature_count = X_train_processed.shape[1]
        timesteps = 4  # Künstliche Zeitfenster
        features_per_step = feature_count // timesteps

        X_train_lstm = X_train_processed[:, :features_per_step * timesteps].reshape(
            (X_train_processed.shape[0], timesteps, features_per_step))
        X_test_lstm = X_test_processed[:, :features_per_step * timesteps].reshape(
            (X_test_processed.shape[0], timesteps, features_per_step))

        # LSTM-Modell bauen
        model = keras.Sequential([
            keras.layers.LSTM(64, return_sequences=True, input_shape=(timesteps, features_per_step)),
            keras.layers.LSTM(32),
            keras.layers.Dense(16, activation='relu'),
            keras.layers.Dense(1, activation='sigmoid')
        ])

        model.compile(optimizer='adam',
                      loss='binary_crossentropy',
                      metrics=['accuracy', keras.metrics.AUC()])

        # Training
        history = model.fit(
            X_train_lstm, self.y_train,
            epochs=50,
            batch_size=32,
            validation_split=0.2,
            callbacks=[keras.callbacks.EarlyStopping(patience=10)]
        )

        # Evaluierung
        y_pred_prob = model.predict(X_test_lstm)
        y_pred = (y_pred_prob > 0.5).astype(int).flatten()

        # Speichern und auswerten
        self.models['lstm'] = {'preprocessor': preprocessor, 'model': model}
        results = self.evaluate_model(y_pred, y_pred_prob.flatten(), "LSTM")
        self.results['lstm'] = results

        return results

    def train_transformer_model(self):
        """Trainiert einen Transformer für komplexe Muster in Stornodaten"""
        from tensorflow import keras
        import tensorflow as tf

        # Daten vorbereiten
        preprocessor = self.build_preprocessor()
        X_train_processed = preprocessor.fit_transform(self.X_train)
        X_test_processed = preprocessor.transform(self.X_test)

        # Parameter
        embed_dim = 32
        num_heads = 4
        ff_dim = 64
        input_dim = X_train_processed.shape[1]

        # Transformer Layer
        def transformer_encoder(inputs, head_size, num_heads, ff_dim, dropout=0):
            x = keras.layers.LayerNormalization(epsilon=1e-6)(inputs)
            x = keras.layers.MultiHeadAttention(
                key_dim=head_size, num_heads=num_heads, dropout=dropout)(x, x)
            x = keras.layers.Dropout(dropout)(x)
            res = x + inputs

            x = keras.layers.LayerNormalization(epsilon=1e-6)(res)
            x = keras.layers.Dense(ff_dim, activation="relu")(x)
            x = keras.layers.Dropout(dropout)(x)
            x = keras.layers.Dense(inputs.shape[-1])(x)
            return x + res

        # Modell bauen
        inputs = keras.Input(shape=(input_dim,))
        x = keras.layers.Reshape((input_dim, 1))(inputs)
        x = keras.layers.Dense(embed_dim)(x)
        x = transformer_encoder(x, embed_dim // num_heads, num_heads, ff_dim)
        x = keras.layers.GlobalAveragePooling1D()(x)
        x = keras.layers.Dropout(0.2)(x)
        x = keras.layers.Dense(20, activation="relu")(x)
        outputs = keras.layers.Dense(1, activation="sigmoid")(x)
        model = keras.Model(inputs=inputs, outputs=outputs)

        model.compile(optimizer="adam", loss="binary_crossentropy", metrics=["accuracy"])

        # Training
        history = model.fit(
            X_train_processed, self.y_train,
            epochs=30,
            batch_size=32,
            validation_split=0.2,
            callbacks=[keras.callbacks.EarlyStopping(patience=5)]
        )

        # Evaluierung
        y_pred_prob = model.predict(X_test_processed)
        y_pred = (y_pred_prob > 0.5).astype(int).flatten()

        # Speichern und auswerten
        self.models['transformer'] = {'preprocessor': preprocessor, 'model': model}
        results = self.evaluate_model(y_pred, y_pred_prob.flatten(), "Transformer")
        self.results['transformer'] = results

        return results

    def run_full_analysis(self):
        """
        Führt die komplette Analyse mit allen implementierten Modellen durch
        """
        try:
            print("Starte die vollständige ML-Analyse...")

            # Daten laden
            if not self.load_data():
                return False

            # Daten vorverarbeiten
            if not self.preprocess_data():
                return False

            # Modelle trainieren und evaluieren
            self.train_gradient_boosting()
            self.train_random_forest()
            self.train_logistic_regression()

            try:
                self.train_xgboost()
            except ImportError:
                print("XGBoost nicht verfügbar - überspringe XGBoost-Modell")

            try:
                self.train_lightgbm()
            except ImportError:
                print("LightGBM nicht verfügbar - überspringe LightGBM-Modell")

            try:
                self.train_neural_network()
            except ImportError:
                print("TensorFlow nicht verfügbar - überspringe neuronales Netzwerk")

            try:
                self.train_neural_network()
                self.train_lstm_model()
                self.train_transformer_model()
            except ImportError as e:
                print(f"Deep Learning-Modelle konnten nicht trainiert werden: {e}")
                print("Bitte installiere TensorFlow mit: pip install tensorflow")


            # Modelle vergleichen
            self.compare_models()

            # Zusammenfassung
            total_records = len(self.df) if self.df is not None else 0
            print("\n" + "-" * 60)
            print(f"DATENSATZ-ZUSAMMENFASSUNG:")
            print(f"Geladene Datensätze gesamt: {total_records}")
            print(f"Trainset-Größe: {len(self.X_train)} ({len(self.X_train) / total_records * 100:.1f}%)")
            print(f"Testset-Größe: {len(self.X_test)} ({len(self.X_test) / total_records * 100:.1f}%)")
            print(f"Kündigungsrate im Trainingsset: {self.y_train.mean() * 100:.1f}%")
            print(f"Kündigungsrate im Testset: {self.y_test.mean() * 100:.1f}%")
            print("-" * 60)

            print("\nML-Analyse erfolgreich abgeschlossen!")
            return True

        except Exception as e:
            print(f"Ein unerwarteter Fehler ist aufgetreten: {str(e)}")
            import traceback
            traceback.print_exc()
            return False


def main():
            """
            Hauptfunktion zum Ausführen der ML-basierten Stornoprognose
            """
            # Dateipfad kann als Parameter übergeben oder der Standard verwendet werden
            file_path = None  # Default-Pfad wird in der Klasse verwendet

            # ML-Tool initialisieren und ausführen
            ml_tool = MLStornoPredictionTool(file_path)
            ml_tool.run_full_analysis()

            print("\n" + "=" * 80)
            print("ML-ANALYSE ABGESCHLOSSEN")
            print("=" * 80)

if __name__ == "__main__":
    main()