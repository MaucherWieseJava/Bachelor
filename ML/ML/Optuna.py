# =====================================================================
# AUTOR: @Adrian Stötzler
# TITEL: Optuna-basierte Hyperparameter-Optimierung für ML-Modelle
# BESCHREIBUNG: Dieses Skript optimiert Gradient Boosting und LSTM
# für die Stornoprognose mittels Optuna
# =====================================================================

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
from sklearn.metrics import (classification_report, confusion_matrix, accuracy_score,
                             roc_auc_score, roc_curve, precision_recall_curve,
                             average_precision_score, precision_score, recall_score, f1_score)

# ML-Algorithmen
from sklearn.ensemble import GradientBoostingClassifier
import xgboost as xgb
import lightgbm as lgbm
import tensorflow as tf
from tensorflow import keras

# Optuna für Hyperparameter-Optimierung
import optuna
from optuna.visualization import plot_optimization_history, plot_param_importances
from optuna.pruners import MedianPruner
from optuna.samplers import TPESampler

# Warnungen unterdrücken
import warnings

warnings.filterwarnings('ignore')


class OptunaMLOptimizer:
    """
    Klasse zur Optimierung von Gradient Boosting und LSTM-Modellen
    mit Hilfe von Optuna für die Stornoprognose
    """

    def __init__(self, file_path=None):
        """
        Initialisiert das ML-Optimierungstool

        Parameter:
            file_path (str): Pfad zur Excel-Datei mit Versicherungsdaten
        """
        self.file_path = file_path or os.path.join(os.environ["HOME"], "Desktop", "Training2.xlsx")
        self.models = {}
        self.results = {}
        self.df = None
        self.df_processed = None
        self.X_train = None
        self.X_test = None
        self.y_train = None
        self.y_test = None
        self.best_params_gb = None
        self.best_params_lstm = None
        self.numeric_features = None
        self.categorical_features = None

        print("=" * 80)
        print("OPTUNA-OPTIMIERTE STORNOPROGNOSE")
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

            # WICHTIG: Entferne Features, die Stornoinformationen direkt offenlegen
            leakage_columns = ['Deletion Type', 'Deletion allowed at', 'Promised Deletion Date', 'Last Due Date',
                               'Meldedatum']
            for col in leakage_columns:
                if col in df_prep.columns:
                    print(f"  ⚠ Entferne '{col}' als Feature, da es Storno-Informationen offenlegt")
                    if col != 'Deletion Type':  # Deletion Type behalten wir für die Zielvariable
                        df_prep.drop(col, axis=1, inplace=True)

            # Datum-Spalten konvertieren
            date_columns = ['Start Insurance', 'End Insurance', 'FirstDueDate', 'Birthday']
            for col in date_columns:
                if col in df_prep.columns:
                    df_prep[col] = pd.to_datetime(df_prep[col], errors='coerce')
                    print(f"  ✓ '{col}' zu Datum konvertiert")

            # Feature-Engineering für Geburtsdatum
            if 'Birthday' in df_prep.columns:
                # Alter berechnen
                today = pd.Timestamp.now()
                df_prep['age'] = (today - df_prep['Birthday']).dt.days / 365.25
                df_prep['age'] = df_prep['age'].fillna(-1).clip(lower=0)
                print(f"  → 'age'-Feature aus 'Birthday' erstellt")

                # Geburtsmonat als kategoriales Feature
                df_prep['birth_month'] = df_prep['Birthday'].dt.month
                df_prep['birth_month'] = df_prep['birth_month'].fillna(-1).astype(int)
                print(f"  → 'birth_month'-Feature aus 'Birthday' erstellt")

            # Vertragsdauer berechnen (ohne Bezug auf Stornierung)
            if 'Start Insurance' in df_prep.columns and 'End Insurance' in df_prep.columns:
                df_prep['contract_duration'] = (df_prep['End Insurance'] - df_prep['Start Insurance']).dt.days
                df_prep['contract_duration'] = df_prep['contract_duration'].fillna(-1).clip(lower=0)
                print(f"  → 'contract_duration'-Feature erstellt")

            # Zeit seit Vertragsbeginn (ohne Bezug auf Stornierung)
            if 'Start Insurance' in df_prep.columns:
                today = pd.Timestamp.now()
                df_prep['days_from_start'] = (today - df_prep['Start Insurance']).dt.days
                df_prep['days_from_start'] = df_prep['days_from_start'].fillna(-1).clip(lower=0)
                print(f"  → 'days_from_start'-Feature erstellt")

            # "Amount" Spalte behandeln
            if 'Amount' in df_prep.columns:
                df_prep['Amount'] = pd.to_numeric(df_prep['Amount'].astype(str).str.replace(',', '.'), errors='coerce')
                df_prep['Amount'] = df_prep['Amount'].fillna(df_prep['Amount'].median())
                print(f"  → 'Amount'-Feature bereinigt")

            # Zielvariable erstellen: Kündigung (1) oder nicht (0)
            if 'Deletion Type' in df_prep.columns:
                df_prep['target'] = (df_prep['Deletion Type'] != 0).astype(int)
                print(f"  → Zielvariable 'target' erstellt: {df_prep['target'].value_counts().to_dict()}")

            # "Amount" Spalte behandeln
            if 'Amount' in df_prep.columns:
                df_prep['Amount'] = pd.to_numeric(df_prep['Amount'].astype(str).str.replace(',', '.'), errors='coerce')
                df_prep['Amount'] = df_prep['Amount'].fillna(df_prep['Amount'].median())

            # Zielvariable erstellen: Kündigung (1) oder nicht (0)
            df_prep['target'] = (df_prep['Deletion Type'] != 0).astype(int)
            print(f"  → Zielvariable 'target' erstellt: {df_prep['target'].value_counts().to_dict()}")

            # Dummy-Variablen für kategoriale Features erstellen
            categorical_columns = ['Country_Region Code', 'Product Code', 'Kampagne']
            # Geburtsmonat als kategorisches Feature hinzufügen, falls vorhanden
            if 'birth_month' in df_prep.columns:
                categorical_columns.append('birth_month')

            for col in categorical_columns:
                if col in df_prep.columns:
                    df_prep[col] = df_prep[col].astype(str)
                    print(f"  → '{col}' zu einheitlichem String-Format konvertiert")

            # Features für das Modell auswählen
            selected_features = []

            # Numerische Features
            numeric_cols = ['Amount', 'contract_duration', 'days_from_start', 'days_until_deletion_allowed']
            # Alter als numerisches Feature hinzufügen, falls vorhanden
            if 'age' in df_prep.columns:
                numeric_cols.append('age')

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
        cm, cm_norm, metrics_dict = self.visualize_perfect_confusion_matrix(
            self.y_test, y_pred, model_name)

        # Aktualisiere die Ergebnisse mit den erweiterten Metriken
        results = {
            'accuracy': acc,
            'auc': auc,
            'ap': ap,
            'confusion_matrix': cm
        }
        results.update(metrics_dict)

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

        return results

    def visualize_perfect_confusion_matrix(self, y_true, y_pred, model_name="Modell"):
        """
        Erstellt eine Konfusionsmatrix und eine perfekte Konfusionsmatrix mit 100% Genauigkeit
        """
        # Konfusionsmatrix berechnen
        cm = confusion_matrix(y_true, y_pred)

        # Perfekte Konfusionsmatrix (100% Genauigkeit) erstellen
        perfect_cm = np.zeros_like(cm)
        for i in range(len(perfect_cm)):
            perfect_cm[i, i] = np.sum(y_true == i)  # Summe der tatsächlichen Klassen auf der Diagonale

        # Zusätzliche Metriken
        precision = precision_score(y_true, y_pred)
        recall = recall_score(y_true, y_pred)
        f1 = f1_score(y_true, y_pred)
        accuracy = (cm[0, 0] + cm[1, 1]) / cm.sum()

        # Visualisierung
        fig, ax = plt.subplots(1, 2, figsize=(16, 7))

        # Absolute Werte der tatsächlichen Konfusionsmatrix
        sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', ax=ax[0],
                    xticklabels=['Kein Storno', 'Storno'],
                    yticklabels=['Kein Storno', 'Storno'])
        ax[0].set_title(f'Reale Konfusionsmatrix: {model_name}')
        ax[0].set_xlabel('Vorhergesagte Klasse')
        ax[0].set_ylabel('Tatsächliche Klasse')

        # Perfekte Konfusionsmatrix
        sns.heatmap(perfect_cm, annot=True, fmt='d', cmap='Blues', ax=ax[1],
                    xticklabels=['Kein Storno', 'Storno'],
                    yticklabels=['Kein Storno', 'Storno'])
        ax[1].set_title(f'Perfekte Konfusionsmatrix (100% Genauigkeit)')
        ax[1].set_xlabel('Vorhergesagte Klasse')

        # Metriken als Text hinzufügen
        plt.figtext(0.5, 0.01,
                    f'Accuracy: {accuracy:.2%} | Precision: {precision:.2%} | Recall: {recall:.2%} | F1-Score: {f1:.2%}',
                    ha='center', fontsize=12, bbox={'facecolor': 'lightblue', 'alpha': 0.5, 'pad': 5})

        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        plt.savefig(f"perfect_confusion_matrix_{model_name.lower().replace(' ', '_')}.png", dpi=300)
        plt.close()

        return cm, perfect_cm, {'accuracy': accuracy, 'precision': precision, 'recall': recall, 'f1': f1}

    def optimize_gradient_boosting_with_optuna(self, n_trials=50):
        """
        Optimiert Gradient Boosting Hyperparameter mit Optuna

        Parameter:
            n_trials (int): Anzahl der Optuna-Optimierungsdurchläufe

        Rückgabe:
            dict: Optimierte Hyperparameter und Modellergebnisse
        """
        print("\n" + "=" * 60)
        print(f"GRADIENT BOOSTING HYPERPARAMETER-OPTIMIERUNG MIT OPTUNA")
        print(f"Durchführung von {n_trials} Optimierungsversuchen...")
        print("=" * 60)

        # Präprozessor erstellen
        preprocessor = self.build_preprocessor()

        # X_train und X_test transformieren (für konsistente Transformationen über alle Trials hinweg)
        X_train_preprocessed = preprocessor.fit_transform(self.X_train)
        X_test_preprocessed = preprocessor.transform(self.X_test)

        # Optuna-Studie erstellen
        study_name = "gradient_boosting_optimization"
        study = optuna.create_study(
            study_name=study_name,
            direction="maximize",  # AUC maximieren
            pruner=MedianPruner(n_warmup_steps=5),
            sampler=TPESampler(seed=42)
        )

        # Zielfunktion für Optuna definieren
        def objective(trial):
            # Hyperparameter festlegen
            param = {
                'n_estimators': trial.suggest_int('n_estimators', 50, 500),
                'learning_rate': trial.suggest_float('learning_rate', 0.01, 0.3, log=True),
                'max_depth': trial.suggest_int('max_depth', 3, 10),
                'min_samples_split': trial.suggest_int('min_samples_split', 2, 20),
                'min_samples_leaf': trial.suggest_int('min_samples_leaf', 1, 10),
                'subsample': trial.suggest_float('subsample', 0.5, 1.0),
                'max_features': trial.suggest_categorical('max_features', ['sqrt', 'log2', None]),
                'random_state': 42
            }

            # Modell mit den aktuellen Hyperparametern erstellen
            model = GradientBoostingClassifier(**param)

        try:
            # Cross-Validation durchführen
            cv_scores = cross_val_score(
                model, X_train_preprocessed, self.y_train,
                cv=5, scoring='roc_auc', n_jobs=-1
            )

            # Mittleren AUC-Score zurückgeben
            return cv_scores.mean()
        except Exception as e:
            # Bei perfekter Trennung kann AUC nicht berechnet werden
            print(f"Warnung bei Trial: {str(e)}")
            return 0.0  # Niedrigsten Wert zurückgeben


        # Optimierung durchführen
        study.optimize(objective, n_trials=n_trials, show_progress_bar=True)

        # Beste Hyperparameter ausgeben
        print("\n" + "-" * 60)
        print("BESTE HYPERPARAMETER FÜR GRADIENT BOOSTING:")
        print("-" * 60)
        best_params = study.best_params
        for param, value in best_params.items():
            print(f"{param}: {value}")
        print(f"Bester AUC-Score (CV): {study.best_value:.4f}")

        # Optimierungsverlauf visualisieren
        fig = plot_optimization_history(study)
        fig.write_image("gb_optimization_history.png")

        try:
            # Cross-Validation durchführen
            cv_scores = cross_val_score(
                model, X_train_preprocessed, self.y_train,
                cv=5, scoring='roc_auc', n_jobs=-1
            )

            # Mittleren AUC-Score zurückgeben
            return cv_scores.mean()
        except Exception as e:
            # Bei perfekter Trennung kann AUC nicht berechnet werden
            print(f"Warnung bei Trial: {str(e)}")
            return 0.0  # Niedrigsten Wert zurückgeben

        # Bestes Modell trainieren und evaluieren
        print("\n" + "-" * 60)
        print("TRAINING DES OPTIMIERTEN GRADIENT BOOSTING-MODELLS")
        print("-" * 60)
        best_model = GradientBoostingClassifier(**best_params)
        best_model.fit(X_train_preprocessed, self.y_train)

        # Vorhersagen mit dem besten Modell
        y_pred = best_model.predict(X_test_preprocessed)
        y_prob = best_model.predict_proba(X_test_preprocessed)[:, 1]

        # Modell evaluieren
        results = self.evaluate_model(y_pred, y_prob, "Optimiertes Gradient Boosting")

        # Modell und Ergebnisse speichern
        self.models['optimized_gradient_boosting'] = {
            'preprocessor': preprocessor,
            'model': best_model,
            'study': study
        }

        self.results['optimized_gradient_boosting'] = results
        self.best_params_gb = best_params

        # Feature-Wichtigkeit visualisieren
        feature_importances = best_model.feature_importances_

        plt.figure(figsize=(12, 8))
        indices = np.argsort(feature_importances)[-15:]  # Top 15 Features
        plt.barh(range(len(indices)), feature_importances[indices])
        plt.yticks(range(len(indices)), [f"Feature {i}" for i in indices])
        plt.title('Optimiertes Gradient Boosting: Feature-Wichtigkeit')
        plt.tight_layout()
        plt.savefig("optimized_gb_feature_importance.png")
        plt.close()

        return {
            'best_params': best_params,
            'best_score': study.best_value,
            'results': results
        }

    def optimize_lstm_with_optuna(self, n_trials=30):
        """
        Optimiert LSTM Hyperparameter mit Optuna für Storno-Vorhersage

        Parameter:
            n_trials (int): Anzahl der Optuna-Optimierungsdurchläufe

        Rückgabe:
            dict: Optimierte Hyperparameter und Modellergebnisse
        """
        print("\n" + "=" * 60)
        print(f"LSTM HYPERPARAMETER-OPTIMIERUNG MIT OPTUNA")
        print(f"Durchführung von {n_trials} Optimierungsversuchen...")
        print("=" * 60)

        # Präprozessor erstellen und Daten transformieren
        preprocessor = self.build_preprocessor()
        X_train_processed = preprocessor.fit_transform(self.X_train)
        X_test_processed = preprocessor.transform(self.X_test)

        # Feature-Dimension
        input_dim = X_train_processed.shape[1]

        # Optuna-Studie erstellen
        study_name = "lstm_optimization"
        study = optuna.create_study(
            study_name=study_name,
            direction="maximize",  # AUC maximieren
            pruner=MedianPruner(n_warmup_steps=5),
            sampler=TPESampler(seed=42)
        )

        # Zielfunktion für Optuna definieren
        def objective(trial):
            # Hyperparameter definieren
            lstm_units_1 = trial.suggest_int('lstm_units_1', 32, 256)
            lstm_units_2 = trial.suggest_int('lstm_units_2', 16, 128)
            dense_units = trial.suggest_int('dense_units', 16, 128)
            dropout_rate_1 = trial.suggest_float('dropout_rate_1', 0.1, 0.5)
            dropout_rate_2 = trial.suggest_float('dropout_rate_2', 0.1, 0.5)
            learning_rate = trial.suggest_float('learning_rate', 1e-4, 1e-2, log=True)
            batch_size = trial.suggest_categorical('batch_size', [16, 32, 64, 128])

            # Temporäre Daten für LSTM-Form transformieren
            # Für LSTM benötigen wir ein 3D-Format: [Samples, Timesteps, Features]
            # Wir teilen Features künstlich in Timesteps auf
            timesteps = trial.suggest_int('timesteps', 2, 8)
            features_per_step = input_dim // timesteps

            # Wenn die Features nicht gleichmäßig aufgeteilt werden können, passen wir an
            if features_per_step * timesteps < input_dim:
                features_per_step += 1

            # Auffüllen, falls nötig
            if features_per_step * timesteps > input_dim:
                padding = features_per_step * timesteps - input_dim
                X_train_padded = np.pad(X_train_processed, ((0, 0), (0, padding)), mode='constant')
                X_test_padded = np.pad(X_test_processed, ((0, 0), (0, padding)), mode='constant')
            else:
                X_train_padded = X_train_processed
                X_test_padded = X_test_processed

            X_train_lstm = X_train_padded.reshape(-1, timesteps, features_per_step)

            # LSTM-Modell bauen
            tf.keras.backend.clear_session()  # Session zurücksetzen

            # Klassengewichte berechnen (für unbalancierte Daten)
            from sklearn.utils.class_weight import compute_class_weight
            class_weights = compute_class_weight('balanced', classes=np.unique(self.y_train), y=self.y_train)
            class_weight_dict = {i: w for i, w in enumerate(class_weights)}

            # Modellarchitektur
            model = keras.Sequential()

            # Erste LSTM-Schicht (mit Return-Sequences für gestapelte LSTM)
            model.add(keras.layers.LSTM(
                lstm_units_1,
                input_shape=(timesteps, features_per_step),
                return_sequences=True
            ))
            model.add(keras.layers.Dropout(dropout_rate_1))
            model.add(keras.layers.BatchNormalization())

            # Zweite LSTM-Schicht
            model.add(keras.layers.LSTM(
                lstm_units_2,
                return_sequences=False
            ))
            model.add(keras.layers.Dropout(dropout_rate_2))
            model.add(keras.layers.BatchNormalization())

            # Dense-Schichten
            model.add(keras.layers.Dense(dense_units, activation='relu'))
            model.add(keras.layers.Dropout(dropout_rate_1))

            # Ausgabeschicht
            model.add(keras.layers.Dense(1, activation='sigmoid'))

            # Modell kompilieren
            model.compile(
                optimizer=keras.optimizers.Adam(learning_rate=learning_rate),
                loss='binary_crossentropy',
                metrics=['accuracy', keras.metrics.AUC()]
            )

            # Callbacks für frühes Stoppen
            early_stopping = keras.callbacks.EarlyStopping(
                monitor='val_loss',
                patience=10,
                restore_best_weights=True
            )

            # Modell trainieren
            history = model.fit(
                X_train_lstm, self.y_train,
                epochs=100,  # Wir verwenden Early Stopping
                batch_size=batch_size,
                validation_split=0.2,
                class_weight=class_weight_dict,
                callbacks=[early_stopping],
                verbose=0
            )

            # Validation AUC aus dem letzten Epoch
            val_auc_key = None
            for key in history.history.keys():
                if 'auc' in key.lower() and 'val' in key.lower():
                    val_auc_key = key
                    break

            if val_auc_key:
                val_auc = history.history[val_auc_key][-1]
            else:
                # Fallback: Accuracy verwenden
                val_auc = history.history['val_accuracy'][-1]

            return val_auc

        # Optimierung durchführen
        study.optimize(objective, n_trials=n_trials, show_progress_bar=True)

        # Beste Hyperparameter ausgeben
        print("\n" + "-" * 60)
        print("BESTE HYPERPARAMETER FÜR LSTM:")
        print("-" * 60)
        best_params = study.best_params
        for param, value in best_params.items():
            print(f"{param}: {value}")
        print(f"Bester Score (Validation AUC): {study.best_value:.4f}")

        # Optimierungsverlauf visualisieren
        fig = plot_optimization_history(study)
        fig.write_image("lstm_optimization_history.png")

        fig = plot_param_importances(study)
        fig.write_image("lstm_param_importances.png")
        print("✓ Optimierungsverlauf und Parameter-Wichtigkeiten gespeichert")

        # Bestes LSTM-Modell trainieren und evaluieren
        print("\n" + "-" * 60)
        print("TRAINING DES OPTIMIERTEN LSTM-MODELLS")
        print("-" * 60)

        # Besten Hyperparameter extrahieren
        lstm_units_1 = best_params['lstm_units_1']
        lstm_units_2 = best_params['lstm_units_2']
        dense_units = best_params['dense_units']
        dropout_rate_1 = best_params['dropout_rate_1']
        dropout_rate_2 = best_params['dropout_rate_2']
        learning_rate = best_params['learning_rate']
        batch_size = best_params['batch_size']
        timesteps = best_params['timesteps']

        # Feature-Dimension
        input_dim = X_train_processed.shape[1]
        features_per_step = input_dim // timesteps
        if features_per_step * timesteps < input_dim:
            features_per_step += 1

        # Daten für LSTM aufbereiten
        if features_per_step * timesteps > input_dim:
            padding = features_per_step * timesteps - input_dim
            X_train_padded = np.pad(X_train_processed, ((0, 0), (0, padding)), mode='constant')
            X_test_padded = np.pad(X_test_processed, ((0, 0), (0, padding)), mode='constant')
        else:
            X_train_padded = X_train_processed
            X_test_padded = X_test_processed

        X_train_lstm = X_train_padded.reshape(-1, timesteps, features_per_step)
        X_test_lstm = X_test_padded.reshape(-1, timesteps, features_per_step)

        # Klassengewichte
        from sklearn.utils.class_weight import compute_class_weight
        class_weights = compute_class_weight('balanced', classes=np.unique(self.y_train), y=self.y_train)
        class_weight_dict = {i: w for i, w in enumerate(class_weights)}

        # LSTM-Modell bauen
        tf.keras.backend.clear_session()
        final_model = keras.Sequential()

        # Erste LSTM-Schicht
        final_model.add(keras.layers.LSTM(
            lstm_units_1,
            input_shape=(timesteps, features_per_step),
            return_sequences=True
        ))
        final_model.add(keras.layers.Dropout(dropout_rate_1))
        final_model.add(keras.layers.BatchNormalization())

        # Zweite LSTM-Schicht
        final_model.add(keras.layers.LSTM(
            lstm_units_2,
            return_sequences=False
        ))
        final_model.add(keras.layers.Dropout(dropout_rate_2))
        final_model.add(keras.layers.BatchNormalization())

        # Dense-Schichten
        final_model.add(keras.layers.Dense(dense_units, activation='relu'))
        final_model.add(keras.layers.Dropout(dropout_rate_1))

        # Ausgabeschicht
        final_model.add(keras.layers.Dense(1, activation='sigmoid'))

        # Modell kompilieren
        final_model.compile(
            optimizer=keras.optimizers.Adam(learning_rate=learning_rate),
            loss='binary_crossentropy',
            metrics=['accuracy', keras.metrics.AUC()]
        )

        # Callbacks
        early_stopping = keras.callbacks.EarlyStopping(
            monitor='val_loss',
            patience=15,
            restore_best_weights=True
        )

        model_checkpoint = keras.callbacks.ModelCheckpoint(
            'best_lstm_model.h5',
            monitor='val_loss',
            save_best_only=True,
            verbose=1
        )

        # Modell trainieren
        history = final_model.fit(
            X_train_lstm, self.y_train,
            epochs=150,
            batch_size=batch_size,
            validation_split=0.2,
            class_weight=class_weight_dict,
            callbacks=[early_stopping, model_checkpoint],
            verbose=1
        )

        # Modell evaluieren
        y_prob = final_model.predict(X_test_lstm)
        y_pred = (y_prob > 0.5).astype(int).flatten()

        # Ergebnisse evaluieren
        results = self.evaluate_model(y_pred, y_prob, "Optimiertes LSTM")

        # Modell und Ergebnisse speichern
        self.models['optimized_lstm'] = {
            'preprocessor': preprocessor,
            'model': final_model,
            'study': study,
            'timesteps': timesteps,
            'features_per_step': features_per_step
        }

        self.results['optimized_lstm'] = results
        self.best_params_lstm = best_params

        # Visualisierung des Trainingsverlaufs
        plt.figure(figsize=(12, 5))

        plt.subplot(1, 2, 1)
        plt.plot(history.history['loss'], label='Train Loss')
        plt.plot(history.history['val_loss'], label='Validation Loss')
        plt.title('LSTM Trainingsverlauf - Loss')
        plt.xlabel('Epoch')
        plt.ylabel('Loss')
        plt.legend()
        plt.grid(True)

        plt.subplot(1, 2, 2)
        plt.plot(history.history['accuracy'], label='Train Accuracy')
        plt.plot(history.history['val_accuracy'], label='Validation Accuracy')
        plt.title('LSTM Trainingsverlauf - Accuracy')
        plt.xlabel('Epoch')
        plt.ylabel('Accuracy')
        plt.legend()
        plt.grid(True)

        plt.tight_layout()
        plt.savefig("optimized_lstm_training_history.png")
        plt.close()

        return {
            'best_params': best_params,
            'best_score': study.best_value,
            'results': results
        }

    def compare_optimized_models(self):
        """
        Vergleicht die Leistung der optimierten Modelle (Gradient Boosting und LSTM)
        """
        if 'optimized_gradient_boosting' not in self.results or 'optimized_lstm' not in self.results:
            print("\nFehler: Optimierte Modelle müssen zuerst trainiert werden.")
            return False

        print("\n" + "=" * 60)
        print("VERGLEICH DER OPTIMIERTEN MODELLE")
        print("=" * 60)

        # Metriken für beide Modelle sammeln
        gb_results = self.results['optimized_gradient_boosting']
        lstm_results = self.results['optimized_lstm']

        # Tabelle für Metriken erstellen
        metrics = ['accuracy', 'auc', 'precision', 'recall', 'f1']
        models = ['Gradient Boosting', 'LSTM']

        plt.figure(figsize=(12, 7))

        # Barplot für jede Metrik
        x = np.arange(len(metrics))
        width = 0.35

        gb_values = [gb_results.get(m, 0) for m in metrics]
        lstm_values = [lstm_results.get(m, 0) for m in metrics]

        plt.bar(x - width / 2, gb_values, width, label=models[0])
        plt.bar(x + width / 2, lstm_values, width, label=models[1])

        plt.xlabel('Metriken')
        plt.ylabel('Wert')
        plt.title('Vergleich optimierter Modelle')
        plt.xticks(x, metrics)
        plt.legend()
        plt.grid(True, axis='y')

        for i, v in enumerate(gb_values):
            plt.text(i - width / 2, v + 0.01, f'{v:.3f}', ha='center')

        for i, v in enumerate(lstm_values):
            plt.text(i + width / 2, v + 0.01, f'{v:.3f}', ha='center')

        plt.tight_layout()
        plt.savefig("model_comparison.png", dpi=300)
        plt.close()

        # Detaillierte Ausgabe
        print("\nMetrik-Vergleich:")
        print("-" * 40)
        print(f"{'Metrik':<12} {'Gradient Boosting':<18} {'LSTM':<18}")
        print("-" * 40)
        for m in metrics:
            gb_val = gb_results.get(m, "N/A")
            lstm_val = lstm_results.get(m, "N/A")

            if isinstance(gb_val, float):
                gb_str = f"{gb_val:.4f}"
            else:
                gb_str = str(gb_val)

            if isinstance(lstm_val, float):
                lstm_str = f"{lstm_val:.4f}"
            else:
                lstm_str = str(lstm_val)

            print(f"{m:<12} {gb_str:<18} {lstm_str:<18}")

        return True

    def run(self):
        """
        Führt die vollständige Analyse mit Optuna-Optimierung durch
        """
        print("\n" + "=" * 60)
        print("STARTE OPTUNA-OPTIMIERTE ML-ANALYSE")
        print("=" * 60)

        # Daten laden
        if not self.load_data():
            return False

        # Daten vorverarbeiten
        if not self.preprocess_data():
            return False

        # Gradient Boosting optimieren
        self.optimize_gradient_boosting_with_optuna(n_trials=50)

        # LSTM optimieren
        self.optimize_lstm_with_optuna(n_trials=30)

        # Optimierte Modelle vergleichen
        self.compare_optimized_models()

        print("\n" + "=" * 60)
        print("OPTUNA-OPTIMIERUNG ABGESCHLOSSEN")
        print("=" * 60)

        return True

def main():
    """
    Hauptfunktion zum Ausführen der Optuna-basierten ML-Optimierung
    """
    # Dateipfad kann als Parameter übergeben oder der Standard verwendet werden
    file_path = None  # Default-Pfad wird in der Klasse verwendet

    # Optuna-Tool initialisieren und ausführen
    optimizer = OptunaMLOptimizer(file_path)
    optimizer.run()

    print("\n" + "=" * 60)
    print("OPTUNA-OPTIMIERUNG ABGESCHLOSSEN")
    print("=" * 60)

if __name__ == "__main__":
    main()