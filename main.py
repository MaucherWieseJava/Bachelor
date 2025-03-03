import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd


# Funktion zum Laden der CSV-Datei und Anzeigen in der Tabelle
def load_csv():
    filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if filename:
        df = pd.read_csv(filename, delimiter=";")  # Trennzeichen auf Semikolon setzen  # CSV-Datei einlesen
        display_table(df)


# Funktion zur Anzeige der Tabelle in Tkinter mit Treeview
def display_table(df):
    # Lösche bestehende Einträge, falls schon Daten geladen wurden
    for widget in frame.winfo_children():
        widget.destroy()

    # Treeview-Widget für die Tabelle erstellen
    tree = ttk.Treeview(frame, style="Custom.Treeview")
    tree["columns"] = list(df.columns)  # Spaltennamen setzen
    tree["show"] = "headings"  # Keine Indexspalte anzeigen

    # Scrollbars hinzufügen
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    # Spaltenüberschriften hinzufügen
    for col in df.columns:
        tree.heading(col, text=col)  # Spaltenüberschrift
        tree.column(col, width=100, anchor="center")  # Breite setzen

    # Datenzeilen hinzufügen mit alternierenden Farben für das Rastermuster
    for i, row in df.iterrows():
        tag = "even" if i % 2 == 0 else "odd"
        tree.insert("", "end", values=list(row), tags=(tag,))

    # Tags für Farben definieren
    tree.tag_configure("even", background="#E8E8E8")  # Hellgrau
    tree.tag_configure("odd", background="white")  # Weiß

    # Grid-ähnliche Struktur durch Trennlinien
    style = ttk.Style()
    style.configure("Custom.Treeview", rowheight=25, borderwidth=1, relief="solid")

    # Platzierung der Elemente
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    # Grid für das Frame anpassen
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)


# Tkinter Hauptfenster erstellen
root = tk.Tk()
root.title("CSV Tabellenanzeige mit Tkinter")
root.geometry("800x400")

# Button zum Laden der CSV-Datei
btn = tk.Button(root, text="CSV laden", command=load_csv)
btn.pack(pady=10)

# Frame für die Tabelle
frame = tk.Frame(root)
frame.pack(expand=True, fill="both")

# Tkinter Hauptloop starten
root.mainloop()
