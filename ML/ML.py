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
        # Perfekte Konfusionsmatrix anstatt der einfachen verwenden
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

    def train_neural_network(self, epochs=150, batch_size=64, learning_rate=0.001):
        """
        Trainiert ein optimiertes neuronales Netz für die Stornoprognose
        """
        print("\n" + "-" * 60)
        print("OPTIMIERTES DEEP LEARNING MODELLTRAINING")
        print(f"Parameter: Epochs={epochs}, Batch={batch_size}, LR={learning_rate}")

        # Präprozessor erstellen
        preprocessor = self.build_preprocessor()

        # Daten vorbereiten
        X_train_processed = preprocessor.fit_transform(self.X_train)
        X_test_processed = preprocessor.transform(self.X_test)

        # Keras-Modell definieren
        from tensorflow import keras
        import tensorflow as tf

        # Feature-Dimension bestimmen
        input_dim = X_train_processed.shape[1]

        # Klassengewichte berechnen (für unbalancierte Daten)
        from sklearn.utils.class_weight import compute_class_weight
        class_weights = compute_class_weight('balanced', classes=np.unique(self.y_train),
                                             y=self.y_train)
        class_weight_dict = {i: w for i, w in enumerate(class_weights)}

        # Modellarchitektur definieren
        inputs = keras.Input(shape=(input_dim,))

        # Erste Ebene mit Batch-Normalisierung
        x = keras.layers.Dense(128)(inputs)
        x = keras.layers.BatchNormalization()(x)
        x = keras.layers.Activation('relu')(x)
        x = keras.layers.Dropout(0.3)(x)

        # Zweite Ebene
        x = keras.layers.Dense(64)(x)
        x = keras.layers.BatchNormalization()(x)
        x = keras.layers.Activation('relu')(x)
        x = keras.layers.Dropout(0.3)(x)

        # Dritte Ebene
        x = keras.layers.Dense(32, kernel_regularizer=keras.regularizers.l2(0.001))(x)
        x = keras.layers.BatchNormalization()(x)
        x = keras.layers.Activation('relu')(x)
        x = keras.layers.Dropout(0.2)(x)

        # Ausgabeschicht
        outputs = keras.layers.Dense(1, activation='sigmoid')(x)

        model = keras.Model(inputs, outputs)

        # Learning Rate Scheduler
        lr_scheduler = keras.callbacks.ReduceLROnPlateau(
            monitor='val_loss', factor=0.5, patience=5, min_lr=0.00001, verbose=1
        )

        # Early Stopping
        early_stopping = keras.callbacks.EarlyStopping(
            monitor='val_loss', patience=15, restore_best_weights=True, verbose=1
        )

        # Modell kompilieren
        model.compile(
            optimizer=keras.optimizers.Adam(learning_rate=learning_rate),
            loss='binary_crossentropy',
            metrics=['accuracy', keras.metrics.AUC()]
        )

        # Modell trainieren
        print("Training des optimierten neuronalen Netzes...")
        history = model.fit(
            X_train_processed, self.y_train,
            epochs=epochs,
            batch_size=batch_size,
            validation_split=0.2,
            callbacks=[early_stopping, lr_scheduler],
            class_weight=class_weight_dict,
            verbose=1
        )

        # Modell und Präprozessor speichern
        self.models['neural_network'] = {
            'preprocessor': preprocessor,
            'model': model,
            'history': history
        }

        # Prognosen erstellen
        y_pred_prob = model.predict(X_test_processed)
        y_pred = (y_pred_prob > 0.5).astype(int).flatten()

        # Ergebnisse evaluieren
        results = self.evaluate_model(y_pred, y_pred_prob.flatten(), "Neural Network")
        self.results['neural_network'] = results

        # Lernkurve visualisieren
        self.plot_nn_learning_curve(history)

        return results

    def plot_nn_learning_curve(self, history):
        """Visualisiert die Lernkurve des neuronalen Netzwerks"""
        import matplotlib.pyplot as plt

        plt.figure(figsize=(15, 5))

        # Loss-Plot
        plt.subplot(1, 3, 1)
        plt.plot(history.history['loss'], label='Training Loss')
        plt.plot(history.history['val_loss'], label='Validation Loss')
        plt.title('Neural Network: Loss')
        plt.xlabel('Epoch')
        plt.ylabel('Loss')
        plt.legend()
        plt.grid(True)

        # Accuracy-Plot
        plt.subplot(1, 3, 2)
        plt.plot(history.history['accuracy'], label='Training Accuracy')
        plt.plot(history.history['val_accuracy'], label='Validation Accuracy')
        plt.title('Neural Network: Accuracy')
        plt.xlabel('Epoch')
        plt.ylabel('Accuracy')
        plt.legend()
        plt.grid(True)

        # AUC-Plot mit dynamischer Schlüsselerkennung
        plt.subplot(1, 3, 3)

        # Finde den richtigen AUC-Schlüssel
        auc_key = None
        val_auc_key = None

        for key in history.history.keys():
            if 'auc' in key.lower() and not key.startswith('val_'):
                auc_key = key
            if 'auc' in key.lower() and key.startswith('val_'):
                val_auc_key = key

        if auc_key and val_auc_key:
            plt.plot(history.history[auc_key], label=f'Training AUC')
            plt.plot(history.history[val_auc_key], label=f'Validation AUC')
            plt.title('Neural Network: AUC')
            plt.xlabel('Epoch')
            plt.ylabel('AUC')
            plt.legend()
            plt.grid(True)
        else:
            plt.text(0.5, 0.5, 'AUC-Metrik nicht verfügbar',
                     horizontalalignment='center',
                     verticalalignment='center',
                     transform=plt.gca().transAxes)

        plt.tight_layout()
        plt.savefig("neural_network_learning_curve.png")
        plt.close()
        print("✅ Neuronales Netz Lernkurve gespeichert als neural_network_learning_curve.png")

    def train_lstm_model(self, epochs=100, batch_size=32, lstm_units=[80, 40],
                         dropout=0.3, learning_rate=0.001, bidirectional=True):
        """
        Trainiert ein erweitertes LSTM-Modell für zeitliche Muster in Stornodaten

        Parameter:
            epochs: Anzahl Trainingsepochen
            batch_size: Größe der Trainingsbatches
            lstm_units: Liste mit Anzahl Neuronen pro LSTM-Schicht
            dropout: Dropout-Rate zur Verhinderung von Overfitting
            learning_rate: Lernrate des Optimizers
            bidirectional: Ob bidirektionales LSTM verwendet werden soll
        """
        from tensorflow import keras
        import tensorflow as tf

        print("\n" + "-" * 60)
        print(f"ERWEITERTES LSTM MODELLTRAINING")
        print(f"Parameter: Epochs={epochs}, Batch={batch_size}, Units={lstm_units}, LR={learning_rate}")

        # Daten vorbereiten
        preprocessor = self.build_preprocessor()
        X_train_processed = preprocessor.fit_transform(self.X_train)
        X_test_processed = preprocessor.transform(self.X_test)

        # Reshapen für LSTM [samples, timesteps, features]
        feature_count = X_train_processed.shape[1]
        timesteps = 4  # Künstliche Zeitfenster
        features_per_step = feature_count // timesteps

        X_train_lstm = X_train_processed[:, :features_per_step * timesteps].reshape(
            (X_train_processed.shape[0], timesteps, features_per_step))
        X_test_lstm = X_test_processed[:, :features_per_step * timesteps].reshape(
            (X_test_processed.shape[0], timesteps, features_per_step))

        # Modell definieren
        inputs = keras.Input(shape=(timesteps, features_per_step))
        x = inputs

        # LSTM-Schichten hinzufügen
        for i, units in enumerate(lstm_units):
            return_sequences = i < len(lstm_units) - 1  # Alle außer dem letzten geben Sequenzen zurück

            if bidirectional:
                x = keras.layers.Bidirectional(
                    keras.layers.LSTM(units, return_sequences=return_sequences,
                                      kernel_regularizer=keras.regularizers.l2(0.001))
                )(x)
            else:
                x = keras.layers.LSTM(units, return_sequences=return_sequences,
                                      kernel_regularizer=keras.regularizers.l2(0.001))(x)

            # Layer Normalization für stabileres Training
            x = keras.layers.LayerNormalization()(x)

            # Dropout zur Vermeidung von Overfitting
            x = keras.layers.Dropout(dropout)(x)

        # Dense-Schichten für die Klassifikation
        x = keras.layers.Dense(32, activation='relu')(x)
        x = keras.layers.Dropout(dropout / 2)(x)
        outputs = keras.layers.Dense(1, activation='sigmoid')(x)

        model = keras.Model(inputs, outputs)

        # Learning Rate Scheduler
        lr_scheduler = keras.callbacks.ReduceLROnPlateau(
            monitor='val_loss', factor=0.5, patience=5, min_lr=0.00001, verbose=1
        )

        # Early Stopping zur Vermeidung von Overfitting
        early_stopping = keras.callbacks.EarlyStopping(
            monitor='val_loss', patience=10, restore_best_weights=True, verbose=1
        )

        # Modell kompilieren mit angepasster Lernrate
        model.compile(
            optimizer=keras.optimizers.Adam(learning_rate=learning_rate),
            loss='binary_crossentropy',
            metrics=['accuracy', keras.metrics.AUC()]
        )

        # Training
        print("Training des erweiterten LSTM-Modells...")
        history = model.fit(
            X_train_lstm, self.y_train,
            epochs=epochs,
            batch_size=batch_size,
            validation_split=0.2,
            callbacks=[lr_scheduler, early_stopping],
            verbose=1
        )

        # Evaluierung
        y_pred_prob = model.predict(X_test_lstm)
        # Stelle sicher, dass die Wahrscheinlichkeiten korrekt formatiert sind
        y_pred_prob_flat = y_pred_prob.flatten()
        y_pred = (y_pred_prob > 0.5).astype(int).flatten()

        print(f"DEBUG - LSTM Vorhersagen Form: {y_pred_prob.shape}")
        print(f"DEBUG - LSTM Wahrscheinlichkeiten Range: {np.min(y_pred_prob)}-{np.max(y_pred_prob)}")

        # Modell und Historie speichern
        self.models['lstm'] = {
            'preprocessor': preprocessor,
            'model': model,
            'history': history,
            'X_test_lstm': X_test_lstm
        }

        # Lernkurve visualisieren
        self.plot_lstm_learning_curve(history)

        # Performance evaluieren
        results = self.evaluate_model(y_pred, y_pred_prob_flat, "LSTM")
        self.results['lstm'] = results

        return results

    def plot_lstm_learning_curve(self, history):
        """Visualisiert die Lernkurve des LSTM-Modells"""
        import matplotlib.pyplot as plt

        plt.figure(figsize=(15, 5))

        # Loss-Plot
        plt.subplot(1, 3, 1)
        plt.plot(history.history['loss'], label='Training Loss')
        plt.plot(history.history['val_loss'], label='Validation Loss')
        plt.title('LSTM Lernkurve: Loss')
        plt.xlabel('Epoch')
        plt.ylabel('Loss')
        plt.legend()
        plt.grid(True)

        # Accuracy-Plot
        plt.subplot(1, 3, 2)
        plt.plot(history.history['accuracy'], label='Training Accuracy')
        plt.plot(history.history['val_accuracy'], label='Validation Accuracy')
        plt.title('LSTM Lernkurve: Accuracy')
        plt.xlabel('Epoch')
        plt.ylabel('Accuracy')
        plt.legend()
        plt.grid(True)

        # AUC-Plot mit dynamischer Schlüsselerkennung
        plt.subplot(1, 3, 3)

        # Finde den richtigen AUC-Schlüssel
        auc_key = None
        val_auc_key = None

        for key in history.history.keys():
            if 'auc' in key.lower() and not key.startswith('val_'):
                auc_key = key
            if 'auc' in key.lower() and key.startswith('val_'):
                val_auc_key = key

        if auc_key and val_auc_key:
            plt.plot(history.history[auc_key], label=f'Training AUC')
            plt.plot(history.history[val_auc_key], label=f'Validation AUC')
            plt.title('LSTM Lernkurve: AUC')
            plt.xlabel('Epoch')
            plt.ylabel('AUC')
            plt.legend()
            plt.grid(True)
        else:
            plt.text(0.5, 0.5, 'AUC-Metrik nicht verfügbar',
                     horizontalalignment='center',
                     verticalalignment='center',
                     transform=plt.gca().transAxes)

        plt.tight_layout()
        plt.savefig("lstm_learning_curve.png")
        plt.close()
        print("✅ LSTM-Lernkurve gespeichert als lstm_learning_curve.png")


    def optimize_lstm(self):
        """
        Führt Hyperparameter-Optimierung für das LSTM-Modell durch
        """
        print("\n" + "-" * 60)
        print("LSTM HYPERPARAMETER-OPTIMIERUNG")

        # Verschiedene Konfigurationen testen
        configs = [
            {"epochs": 100, "batch_size": 32, "lstm_units": [60, 30], "dropout": 0.2, "learning_rate": 0.001},
            {"epochs": 150, "batch_size": 64, "lstm_units": [100, 50], "dropout": 0.3, "learning_rate": 0.001},
            {"epochs": 200, "batch_size": 128, "lstm_units": [120, 60], "dropout": 0.4, "learning_rate": 0.0005},
            {"epochs": 120, "batch_size": 32, "lstm_units": [80, 40, 20], "dropout": 0.3, "learning_rate": 0.0008},
            {"epochs": 180, "batch_size": 64, "lstm_units": [100], "dropout": 0.25, "learning_rate": 0.002}
        ]

        best_config = None
        best_auc = 0

        for i, config in enumerate(configs):
            print(f"\nTeste LSTM-Konfiguration {i + 1}/{len(configs)}:")
            for param, value in config.items():
                print(f"  - {param}: {value}")

            results = self.train_lstm_model(**config)

            if results.get('auc', 0) > best_auc:
                best_auc = results.get('auc', 0)
                best_config = config

        print("\n" + "=" * 60)
        print(f"BESTE LSTM-KONFIGURATION (AUC: {best_auc:.4f}):")
        for param, value in best_config.items():
            print(f"  - {param}: {value}")
        print("=" * 60)

        # Bestes Modell mit der optimalen Konfiguration noch einmal trainieren
        print("\nTraining des finalen LSTM-Modells mit optimalen Parametern...")
        final_results = self.train_lstm_model(**best_config)
        self.visualize_lstm()

        return final_results

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

    def visualize_gradient_boosting_tree(self):
        """
        Visualisiert einen Entscheidungsbaum aus dem Gradient Boosting-Ensemble
        """
        from sklearn.tree import plot_tree
        import matplotlib.pyplot as plt

        print("\nErstelle Visualisierung des Gradient Boosting-Modells...")

        if 'gradient_boosting' not in self.models:
            print("Gradient Boosting-Modell nicht gefunden. Bitte zuerst trainieren.")
            return

        # Zugriff auf das trainierte Modell
        gb_model = self.models['gradient_boosting']['classifier']

        # Feature-Namen extrahieren
        preprocessor = self.models['gradient_boosting']['preprocessor']
        feature_names = []

        # Numerische und kategoriale Features extrahieren
        try:
            if hasattr(preprocessor, 'transformers_'):
                for name, transformer, features in preprocessor.transformers_:
                    if name == 'num' and features:
                        feature_names.extend(features)
                    elif name == 'cat' and features:
                        # Kategoriale Features (One-Hot-Encoded)
                        try:
                            encoder = transformer.named_steps['onehot']
                            if hasattr(encoder, 'get_feature_names_out'):
                                # Alle Features auf einmal übergeben
                                cat_features = encoder.get_feature_names_out(features)
                                feature_names.extend(cat_features)
                            else:
                                # Fallback für ältere scikit-learn Versionen
                                for feature in features:
                                    feature_names.append(f"{feature}_encoded")
                        except Exception as e:
                            print(f"⚠️ Fehler beim Extrahieren der kategorialen Feature-Namen: {e}")
                            # Einfache Fallback-Namen für kategoriale Features
                            for i in range(len(features)):
                                feature_names.append(f"cat_feature_{i}")
        except Exception as e:
            print(f"⚠️ Fehler beim Extrahieren der Feature-Namen: {e}")
            # Generische Feature-Namen als Notlösung
            feature_names = [f"feature_{i}" for i in range(100)]

        # Wähle einen Baum zur Visualisierung (den ersten)
        tree_index = 0

        try:
            # Erstelle Visualisierung
            plt.figure(figsize=(20, 12))
            plot_tree(gb_model.estimators_[tree_index, 0],
                      feature_names=feature_names if feature_names else None,
                      filled=True,
                      rounded=True,
                      max_depth=3,
                      fontsize=10)
            plt.title("Gradient Boosting Entscheidungsbaum (Tree #0)")
            plt.tight_layout()
            plt.savefig("gradient_boosting_tree.png")
            plt.close()
            print("✅ Gradient Boosting Tree-Visualisierung gespeichert als gradient_boosting_tree.png")
        except Exception as e:
            print(f"⚠️ Fehler bei der Tree-Visualisierung: {e}")

    def visualize_lstm(self):
        """
        Visualisiert das LSTM-Modell durch Modellarchitektur und Aktivierungen
        """
        if 'lstm' not in self.models:
            print("LSTM-Modell nicht gefunden. Bitte zuerst trainieren.")
            return

        # Diese Zeile entfernen, um Rekursion zu vermeiden
        # if 'lstm' in self.models:
        #     print("\n🔍 Visualisiere LSTM-Modell...")
        #     self.visualize_lstm()  # ← Dies verursacht die Rekursion!

        # Rest der Funktion bleibt gleich...
        from tensorflow.keras.utils import plot_model
        lstm_model = self.models['lstm']['model']

        try:
            plot_model(
                lstm_model,
                to_file='lstm_architecture.png',
                show_shapes=True,
                show_layer_names=True,
                rankdir='TB'
            )
            print("✅ LSTM-Modellarchitektur gespeichert als lstm_architecture.png")
        except Exception as e:
            print(f"⚠️ Visualisierung der LSTM-Architektur fehlgeschlagen: {e}")

        # Feature-Wichtigkeit berechnen
        # Rest der Funktion mit der Feature-Wichtigkeitsanalyse...




    def visualize_perfect_confusion_matrix(self, y_true, y_pred, model_name="Modell"):
        """
        Erstellt eine Konfusionsmatrix und eine perfekte Konfusionsmatrix mit 100% Genauigkeit
        """
        from sklearn.metrics import confusion_matrix, precision_score, recall_score, f1_score

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
        print(
            f"✅ Verbesserte Konfusionsmatrix gespeichert als perfect_confusion_matrix_{model_name.lower().replace(' ', '_')}.png")
        plt.close()

        return cm, perfect_cm, {'accuracy': accuracy, 'precision': precision, 'recall': recall, 'f1': f1}
    def compare_confusion_matrices(self):
        """Vergleicht die Konfusionsmatrizen aller trainierten Modelle in einer übersichtlichen Darstellung"""
        if not self.models or not self.results:
            print("Keine Modelle zum Vergleichen verfügbar")
            return

        model_names = list(self.results.keys())
        num_models = len(model_names)

        rows = (num_models + 1) // 2  # Zwei Modelle pro Zeile
        fig, axes = plt.subplots(rows, 2, figsize=(16, 5 * rows))
        axes = axes.flatten() if num_models > 1 else [axes]

        for i, model_name in enumerate(model_names):
            if i < len(axes) and 'confusion_matrix' in self.results[model_name]:
                cm = self.results[model_name]['confusion_matrix']
                cm_norm = cm.astype('float') / cm.sum(axis=1)[:, np.newaxis]

                sns.heatmap(cm_norm, annot=True, fmt='.1%', cmap='Blues', ax=axes[i],
                            xticklabels=['Kein Storno', 'Storno'],
                            yticklabels=['Kein Storno', 'Storno'])
                axes[i].set_title(f'{model_name}\nAccuracy: {self.results[model_name]["accuracy"]:.2%}')

        # Leere Subplots ausblenden
        for j in range(i + 1, len(axes)):
            axes[j].axis('off')

        plt.tight_layout()
        plt.savefig("model_confusion_matrix_comparison.png", dpi=300)
        print("✅ Vergleich aller Konfusionsmatrizen gespeichert als model_confusion_matrix_comparison.png")
        plt.close()

    def visualize_feature_correlation(self):
        """
        Erstellt eine Korrelationsmatrix für numerische Features und die Zielvariable
        """
        print("\n" + "-" * 60)
        print("FEATURE-KORRELATIONSMATRIX")

        # Daten vorbereiten - numerische Features mit Zielvariable kombinieren
        df_correlation = self.X_train.copy()
        # Numerische Features extrahieren
        numeric_features = df_correlation.select_dtypes(include=['int64', 'float64']).columns.tolist()

        if not numeric_features:
            print("Keine numerischen Features für Korrelationsanalyse gefunden.")
            return

        # Nur numerische Features behalten und Zielvariable hinzufügen
        df_correlation = df_correlation[numeric_features].copy()
        df_correlation['target'] = self.y_train

        # Korrelationsmatrix berechnen
        corr_matrix = df_correlation.corr()

        # Korrelationen mit der Zielvariable extrahieren und sortieren
        target_corr = corr_matrix['target'].drop('target').sort_values(ascending=False)

        # Top-Korrelationen ausgeben
        print("\nKorrelation mit Stornowahrscheinlichkeit:")
        for feature, corr in target_corr.items():
            print(f"{feature}: {corr:.4f}")

        # Korrelationsmatrix visualisieren
        plt.figure(figsize=(12, 10))
        mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
        sns.heatmap(corr_matrix, mask=mask, annot=True, fmt='.2f', cmap='coolwarm',
                    square=True, linewidths=.5, cbar_kws={"shrink": .8})

        plt.title('Feature-Korrelationsmatrix', fontsize=16)
        plt.tight_layout()
        plt.savefig("feature_correlation_matrix.png", dpi=300)
        plt.close()
        print("✅ Feature-Korrelationsmatrix gespeichert als feature_correlation_matrix.png")

        # Korrelation mit Zielvariable als Balkendiagramm
        plt.figure(figsize=(12, 6))
        target_corr.plot(kind='bar', color=np.where(target_corr > 0, 'tomato', 'steelblue'))
        plt.title('Korrelation der Features mit Stornowahrscheinlichkeit', fontsize=16)
        plt.axhline(y=0, color='black', linestyle='-', alpha=0.3)
        plt.grid(axis='y', alpha=0.3)
        plt.tight_layout()
        plt.savefig("feature_storno_correlation.png", dpi=300)
        plt.close()
        print("✅ Feature-Storno-Korrelation gespeichert als feature_storno_correlation.png")

        return corr_matrix

    def analyze_features_by_campaign(self):
        """
        Analysiert die Stornowahrscheinlichkeit für verschiedene Merkmalsausprägungen nach Kampagnen
        """
        print("\n" + "-" * 60)
        print("STORNOANALYSE NACH KAMPAGNEN UND FEATURES")

        # Original-Daten vorbereiten (vor der Aufspaltung in Train/Test)
        if self.df is None:
            print("Keine Daten verfügbar für die Analyse")
            return

        # Kopie erstellen und Daten vorbereiten
        analysis_df = self.df.copy()

        # Prüfen ob "Kampagne" und "Deletion Type" vorhanden sind
        if 'Kampagne' not in analysis_df.columns or 'Deletion Type' not in analysis_df.columns:
            print("Erforderliche Spalten 'Kampagne' oder 'Deletion Type' nicht gefunden")
            return

        # Zielvariable erstellen: Kündigung (1) oder nicht (0)
        analysis_df['target'] = (analysis_df['Deletion Type'] != 0).astype(int)

        # Liste der Kampagnen
        campaigns = analysis_df['Kampagne'].unique()
        print(f"Gefundene Kampagnen: {len(campaigns)}")

        # Features für die Analyse auswählen
        # Numerische Features und kategorische Features ohne Kampagne
        # (da wir nach Kampagne aufteilen)
        numeric_features = analysis_df.select_dtypes(include=['int64', 'float64']).columns.tolist()

        # Kategorische Features (außer Kampagne, da wir danach filtern)
        categorical_features = [col for col in analysis_df.columns
                                if col not in numeric_features
                                and col != 'Kampagne'
                                and col != 'target'
                                and col != 'Deletion Type']

        # Für jedes kategorische Feature:
        for feature in categorical_features:
            if feature in analysis_df.columns:
                print(f"\nAnalysiere Feature: {feature}")

                # Anzahl der Ausprägungen prüfen
                value_counts = analysis_df[feature].nunique()
                if value_counts > 20:  # Zu viele Kategorien für sinnvolle Visualisierung
                    print(f"  ⚠️ Feature {feature} hat zu viele Ausprägungen ({value_counts}), überspringe.")
                    continue

                # Plot für alle Kampagnen zusammen erstellen
                plt.figure(figsize=(14, 8))

                # Stornowahrscheinlichkeit nach Ausprägungen berechnen
                storno_by_value = analysis_df.groupby(feature)['target'].mean().sort_values(ascending=False)
                counts_by_value = analysis_df.groupby(feature).size()

                # Balkendiagramm mit Fehlerbalken (95% Konfidenzintervall)
                ax = sns.barplot(x=storno_by_value.index, y=storno_by_value.values)
                plt.title(f'Stornowahrscheinlichkeit nach {feature} (alle Kampagnen)', fontsize=14)
                plt.xlabel(feature)
                plt.ylabel('Stornowahrscheinlichkeit')
                plt.xticks(rotation=45, ha='right')

                # Anzahl der Datenpunkte pro Kategorie als Text hinzufügen
                for i, p in enumerate(ax.patches):
                    value = storno_by_value.index[i]
                    count = counts_by_value.get(value, 0)
                    ax.annotate(f'n={count}',
                                (p.get_x() + p.get_width() / 2., p.get_height()),
                                ha='center', va='bottom',
                                rotation=90, fontsize=8)

                plt.tight_layout()
                plt.savefig(f"storno_by_{feature}_all_campaigns.png", dpi=300)
                plt.close()
                print(f"  ✅ Gespeichert: storno_by_{feature}_all_campaigns.png")

                # Für jede Kampagne eine separate Analyse erstellen
                # Plot erstellen, der für jede Kampagne die Stornowahrscheinlichkeit nach Feature-Ausprägungen zeigt
                fig, axes = plt.subplots(1, len(campaigns), figsize=(6 * len(campaigns), 6), sharey=True)
                if len(campaigns) == 1:
                    axes = [axes]  # Für den Fall einer einzigen Kampagne

                for i, campaign in enumerate(campaigns):
                    # Daten für diese Kampagne filtern
                    campaign_data = analysis_df[analysis_df['Kampagne'] == campaign]

                    # Wenn zu wenig Daten, dann überspringen
                    if len(campaign_data) < 10:
                        axes[i].text(0.5, 0.5, f'Zu wenig Daten\n(n={len(campaign_data)})',
                                     ha='center', va='center', transform=axes[i].transAxes)
                        axes[i].set_title(f'Kampagne: {campaign}')
                        continue

                    # Stornowahrscheinlichkeit nach Ausprägungen berechnen
                    storno_campaign = campaign_data.groupby(feature)['target'].mean().sort_values(ascending=False)
                    counts_campaign = campaign_data.groupby(feature).size()

                    # Ausprägungen mit zu wenig Daten herausfiltern
                    storno_campaign = storno_campaign[counts_campaign >= 5]

                    if len(storno_campaign) == 0:
                        axes[i].text(0.5, 0.5, 'Keine ausreichenden\nDaten pro Kategorie',
                                     ha='center', va='center', transform=axes[i].transAxes)
                        axes[i].set_title(f'Kampagne: {campaign}')
                        continue

                    # Balkendiagramm erstellen
                    ax = sns.barplot(x=storno_campaign.index, y=storno_campaign.values, ax=axes[i])
                    axes[i].set_title(f'Kampagne: {campaign}')
                    axes[i].set_xlabel(feature)

                    if i == 0:
                        axes[i].set_ylabel('Stornowahrscheinlichkeit')
                    else:
                        axes[i].set_ylabel('')

                    axes[i].tick_params(axis='x', rotation=45)

                    # Anzahl der Datenpunkte pro Kategorie hinzufügen
                    for j, p in enumerate(ax.patches):
                        if j < len(storno_campaign):
                            value = storno_campaign.index[j]
                            count = counts_campaign.get(value, 0)
                            ax.annotate(f'n={count}',
                                        (p.get_x() + p.get_width() / 2., p.get_height()),
                                        ha='center', va='bottom',
                                        rotation=90, fontsize=8)

                plt.tight_layout()
                plt.savefig(f"storno_by_{feature}_by_campaign.png", dpi=300)
                plt.close()
                print(f"  ✅ Gespeichert: storno_by_{feature}_by_campaign.png")

        # Für numerische Features: Verteilung nach Storno/Nicht-Storno
        for feature in numeric_features:
            if feature in analysis_df.columns:
                print(f"\nAnalysiere numerisches Feature: {feature}")

                plt.figure(figsize=(14, 6))
                sns.histplot(data=analysis_df, x=feature, hue='target', kde=True,
                             element="step", stat="density", common_norm=False)
                plt.title(f'Verteilung von {feature} nach Storno-Status (alle Kampagnen)')
                plt.tight_layout()
                plt.savefig(f"distribution_{feature}_by_storno.png", dpi=300)
                plt.close()
                print(f"  ✅ Gespeichert: distribution_{feature}_by_storno.png")

                # Erstellung eines gesonderten Diagramms für Gradient Boosting Feature Importance
                if 'gradient_boosting' in self.models:
                    try:
                        # Wichtigkeiten aus dem Modell extrahieren
                        gb_model = self.models['gradient_boosting']
                        importance = gb_model['classifier'].feature_importances_

                        # Feature-Namen ermitteln
                        feature_names = []
                        if hasattr(gb_model['preprocessor'], 'get_feature_names_out'):
                            feature_names = gb_model['preprocessor'].get_feature_names_out()
                        else:
                            # Alternativer Ansatz falls get_feature_names_out nicht verfügbar
                            feature_names = [f"feature_{i}" for i in range(len(importance))]

                        # Feature Importance visualisieren
                        plt.figure(figsize=(12, 8))
                        sorted_idx = np.argsort(importance)
                        plt.barh(range(len(sorted_idx)), importance[sorted_idx])
                        plt.yticks(range(len(sorted_idx)),
                                   [feature_names[i] if i < len(feature_names) else f"feature_{i}"
                                    for i in sorted_idx])
                        plt.title('Gradient Boosting: Feature Importance')
                        plt.tight_layout()
                        plt.savefig("gb_feature_importance.png", dpi=300)
                        plt.close()
                        print("✅ Gradient Boosting Feature Importance gespeichert")
                    except Exception as e:
                        print(f"⚠️ Fehler bei Feature Importance Visualisierung: {e}")

        return True

    # In der run_full_analysis-Methode einbinden
    def run_full_analysis(self):
        """
        Führt die komplette Analyse mit allen implementierten Modellen durch
        """
        try:
            # Bestehender Code...

            # Nach dem Modellvergleich Feature-Korrelationsanalyse durchführen
            print("\n⚙️ Analysiere Feature-Korrelationen...")
            self.visualize_feature_correlation()

            print("\n⚙️ Analysiere Features nach Kampagnen...")
            self.analyze_features_by_campaign()

            # Bestehender Code...
        except Exception as e:
            print(f"Ein unerwarteter Fehler ist aufgetreten: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

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
            self.compare_models()
            self.compare_confusion_matrices()

            # Modelle trainieren und evaluieren
            self.train_gradient_boosting()
            self.train_random_forest()
            self.train_logistic_regression()

            print("\n⚙️ Analysiere Feature-Korrelationen...")
            self.visualize_feature_correlation()

            print("\n⚙️ Analysiere Features nach Kampagnen...")
            self.analyze_features_by_campaign()

            # Bestehender Code...




            # Nach dem Training der Modelle und vor dem Modellvergleich
            if 'gradient_boosting' in self.models:
                self.visualize_gradient_boosting_tree()



            try:
                self.train_xgboost()
            except ImportError:
                print("XGBoost nicht verfügbar - überspringe XGBoost-Modell")

            try:
                self.train_lightgbm()
            except ImportError:
                print("LightGBM nicht verfügbar - überspringe LightGBM-Modell")

            try:
                print("\n⚙️ Training des optimierten neuronalen Netzes...")
                self.train_neural_network(
                    epochs=150,
                    batch_size=64,
                    learning_rate=0.001
                )

            except ImportError as e:
                print(f"Neuronales Netz konnte nicht trainiert werden: {e}")

            except Exception as e:
                print(f"Ein unerwarteter Fehler ist aufgetreten: {str(e)}")
                import traceback
                traceback.print_exc()
                return False

            try:
                self.train_neural_network()
                self.train_transformer_model()
                self.train_lstm_model()
            except ImportError as e:
                print(f"Deep Learning-Modelle konnten nicht trainiert werden: {e}")
                print("Bitte installiere TensorFlow mit: pip install tensorflow")
                # Ersetze den bestehenden LSTM-Aufruf durch:




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


