import base64
import io
import os
import random
import shutil
import string
import csv
import time
import logging
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.backends.backend_svg  # Importiert das Backend explizit (für PyInstaller wichtig!)
import configparser
import locale
import urllib.parse
from datetime import datetime
from reportlab.graphics import renderPDF
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, KeepTogether, PageBreak, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from svglib.svglib import svg2rlg
from lscontrolling_logo import lscontrolling_logo
from version import program_version

# zum Debug mit print alles ausdrucken
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# Locale auf Deutsch setzen (für Formatierung von Zahlen)
locale.setlocale(locale.LC_ALL, 'de_DE.UTF-8')


# --- Definition von Funktionen ----------------------------------------------------------------------------------------

# Klasse für Konfigurationswerte (inkl. Festlegen der Defaultwerte)
class LSControllingConfig:
    def __init__(self, config_file=None):
        self.config = configparser.ConfigParser()

        # Defaultwerte festlegen
        self.defaults = {
            'csv_stammdaten': 'input/WPS_PSP_STAMMDATEN_V1.csv',
            'check_stammdaten': True,
            'header_stammdaten': 3,
            'csv_budget': 'input/WFI_001_FC_BUDGET_V1.csv',
            'check_budget': True,
            'header_budget': 4,
            'csv_obligo': 'input/WFI_001_FC_OBLIGOS_V1.csv',
            'check_obligo': True,
            'header_obligo': 4,
            'csv_kst': 'input/WPSM_004_KSD.csv',
            'check_kst': True,
            'header_kst': 4,
            'liste_pa_aufteilung': [68, 69, 90, 91, 92, 99],
            'liste_pa_keine_aufteilung': [70, 94],
            'csv_detailplot': 'input/PSP_PLOT.csv',
            'rm_beendet': True,
            'rm_current_year': True,
            'prt_raw': False,
            'obfuscated': False
        }

        if config_file and os.path.exists(config_file):
            self.config.read(config_file)
            if not self.config.has_section('lscontrolling'):
                raise ValueError(f"Die Sektion 'lscontrolling' fehlt in der Konfigurationsdatei {config_file}.")

    def __getitem__(self, key):
        section = "lscontrolling"
        if self.config.has_option(section, key):
            value = self.config.get(section, key)
            # Konvertiere Boolean-Werte korrekt
            if value.lower() in ['true', 'false']:
                return self.config.getboolean(section, key)
            # Konvertiere Listenwerte korrekt (z.B. für liste_pa_aufteilung)
            elif "liste" in key:
                if ',' in value:
                    return [int(x.strip()) for x in value.split(',')]
                else:
                    return [int(value.strip())]
            # Konvertiere Integer-Werte korrekt
            elif value.isdigit():
                return int(value)
            return value

        else:
            return self.defaults.get(key)


# Klasse für zufällige, temporäre Verzeichnisse
class RandomTemp:
    def __init__(self, base_dir='.'):
        self.temp_dir = None
        random_suffix = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
        tmp_dir_name = f'tmp_{random_suffix}'
        tmp_dir_path = os.path.join(base_dir, tmp_dir_name)
        os.makedirs(tmp_dir_path)
        self.temp_dir = tmp_dir_path

    def delete_temp_dir(self):
        if self.temp_dir:
            shutil.rmtree(self.temp_dir)


# Zeitmessung und Infotext
class LogContext:
    def __init__(self, message):
        self.message = message
        self.start_time = None

    def __enter__(self):
        self.start_time = time.time()
        print(f"{self.message}...", end='', flush=True)

    def __exit__(self, exc_type, exc_val, exc_tb):
        end_time = time.time()
        elapsed_time = end_time - self.start_time
        print(f" OK ({elapsed_time:.2f} s)")


# SAP CSV Dateien grob prüfen, ob die richtigen Header vorhanden sind
def check_sap_csv_content(file_path, csv_type):
    # Mapping der erwarteten Werte
    expected_values = {
        'stammdaten': "Stammdaten HHP",
        'budget': "Budget",
        'obligo': "Obligos",
        'kst': "Kontostand"
    }

    if csv_type not in expected_values:
        raise Exception(f"Programmierfehler: {csv_type} nicht bekannt in Funktion check_sap_csv_content!")

    with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile, delimiter=';')  # Semikolon als Trennzeichen
        first_line = next(reader)  # Liest die erste Zeile
        tp = first_line[1]
        expected_value = expected_values[csv_type]
        if tp == expected_value:
            return True
        else:
            raise Exception(f"{file_path} enthält nicht den erwarteten Inhalt '{expected_value}' sondern '{tp}'")


# Funktion zum Laden der CSV-Datei mit dynamischem Header
def load_csv_with_dynamic_header(file_path, header_row, dtype_map=None):
    try:
        return pd.read_csv(file_path, sep=';', skiprows=header_row, header=None, dtype=dtype_map,
                           decimal=',', thousands='.')
    except pd.errors.EmptyDataError:
        print(f"Die Datei {file_path} enthält keine Datenzeilen. Bitte prüfen. Im Falle von nicht vorhandenen Obligos "
              f"bitte mit einem existierenden PSP-Element und Festlegungen von 0 Euro auffüllen.")
        exit(1)
    except FileNotFoundError:
        print(f"Für das Programm müssen bestimmte CSV Dateien vorhanden sein.\nBitte prüfen Sie, dass die Datei "
              f"{file_path} im korrekten Unterordner vorliegt und nutzbar ist!")
        exit(1)


# Funktion zum Schreiben von CSV Daten
def write_csv(df, file_path):
    try:
        df.to_csv(file_path, sep=";", decimal=',', encoding='latin1', index=False)
    except PermissionError:
        print(f"Die Datei {file_path} kann nicht geschrieben werden. Bitte prüfen Sie ob sie nicht noch geöffnet ist!")


# Funktion zur Selektion gewisser Spalten die einen Eintrag enthalten
def cont(df, column, select, regex=True):
    return df[df[column].str.contains(select, regex=regex)].reset_index(drop=True)


# Funktion zur Selektion gewisser Spalten die einen Eintrag NICHT enthalten
def not_cont(df, column, select, regex=True):
    return df[~df[column].astype(str).str.contains(select, regex=regex)].reset_index(drop=True)


def rem_current_year(df):
    # Finde das letzte Jahr im Datensatz
    last_year = df['Jahr'].max()

    # Entferne alle Einträge des letzten Jahres
    return df[df['Jahr'] != last_year]


# laufende Projekte nach cutoff ignorieren
def laufende_projekte_ignorieren(df, cutoff):
    # Filtern der Zeilen, deren Projektende nach Cutoff liegt
    filtered_df = df[df['Projektende'] <= cutoff]
    return filtered_df


# nur laufende Projekte
def nur_laufende_projekte(df, cutoff):
    # Filtern der Zeilen, deren Projektende nach Cutoff liegt
    filtered_df = df[df['Projektende'] > cutoff]
    return filtered_df


# Nur Sammelkonten filtern
def nur_sammelkonten(df):
    return cont(df, "Geldgeber", "^999$|^1$", True)


# Sammelkonten rausfiltern
def keine_sammelkonten(df):
    return not_cont(df, "Geldgeber", "^999$|^1$", True)


# involvierte IKZ aus dem Datensatz extrahieren, wenn mehr als eine dann, Exception werfen
def get_ikz(df):
    df_ikz = df['PSP'].str[5:11]
    grouped = df_ikz.value_counts()

    # Überprüfe die Anzahl der Gruppen
    if len(grouped) > 1:
        raise Exception(f"Mehr als eine IKZ im Datensatz gefunden! Bitte prüfen! {list(df.columns)}")
    else:
        return grouped.index[0]


# Anzahl der Jahre im Datensatz zählen. Wenn nur ein Jahr enthalten ist, dann soll eine Exception geworfen werden
def check_jahr(df):
    # Überprüfe die Anzahl der Gruppen
    grouped = df['Jahr'].value_counts()
    if len(grouped) <= 1:
        raise Exception(f"Zu wenig Jahre im Datensatz gefunden. Bitte prüfen! {list(df.columns)}")


# Datenimport aus SAP CSV Tabellen
def import_sap_csv(config: LSControllingConfig):
    # Festlegung der Datentypen (abweichend von standard)
    cv_stammdaten = 'str'
    cv_budget = {0: 'str', 2: 'str', 7: 'float', 8: 'float', 9: 'float'}
    cv_obligo = {0: 'str', 3: 'str', 4: 'str', 7: 'float'}
    cv_kst = {0: 'str', 1: 'str', 2: 'str', 3: 'float', 4: 'float', 5: 'float', 6: 'float', 7: 'float'}

    daten = ['stammdaten', 'budget', 'obligo', 'kst']

    # Prüfen, ob die Header in den CSV-Dateien geprüft werden sollen
    for d in daten:
        if config[f'check_{d}']:
            check_sap_csv_content(config[f'csv_{d}'], d)

    # CSV-Dateien laden
    df_stammdaten = load_csv_with_dynamic_header(config['csv_stammdaten'], config['header_stammdaten'], cv_stammdaten)
    df_budget = load_csv_with_dynamic_header(config['csv_budget'], config['header_budget'], cv_budget)
    df_obligo = load_csv_with_dynamic_header(config['csv_obligo'], config['header_obligo'], cv_obligo)
    df_kst = load_csv_with_dynamic_header(config['csv_kst'], config['header_kst'], cv_kst)

    # Daten vorab bereinigen (alle Zeilen löschen, die ein Ergebnis oder Gesamtergebnis sind)
    df_budget = not_cont(df_budget, 6, 'Ergebnis')
    df_kst = not_cont(df_kst, 2, 'Ergebnis')
    df_kst = df_kst[df_kst[0] != 'Gesamtergebnis']

    # Auswahl und Sortierung von relevanten Spalten der einzelnen Tabellen und NaN mit 0 ersetzen
    df_stammdaten_relevant = df_stammdaten[[3, 4, 2, 7, 10]]
    df_budget_relevant = df_budget[[0, 1, 2, 7, 8, 9]].fillna(0)
    df_obligo_relevant = df_obligo[[3, 4, 0, 7]].fillna(0)
    df_kst_relevant = df_kst[[0, 1, 2, 3, 4, 5, 6, 7]].fillna(0)

    # neue Spaltenüberschriften nach dem Filtern setzen
    df_stammdaten_relevant.columns = ['PSP', 'PSPName', 'Status', 'Projektende', 'Geldgeber']
    df_budget_relevant.columns = ['PSP', 'PSPName', 'Jahr', 'Budgetrest aus Vorjahr', 'Originalbudget',
                                  'Sonstige Zuweisungen']
    df_obligo_relevant.columns = ['PSP', 'PSPName', 'Jahr', 'Festlegungen']
    df_kst_relevant.columns = ['PSP', 'PSPName', 'Jahr', 'Einnahmen ILA', 'Einnahmen-Ist',
                               'Eigen- und Industrieanteile', 'Ausgaben-Ist', 'Kontostand Jahr']

    # IKZ der Datensätze prüfen (es müssen alle aus allen Datensätzen gleich sein)
    ikz_stammdaten = get_ikz(df_stammdaten_relevant)
    ikz_budget = get_ikz(df_budget_relevant)
    ikz_obligo = get_ikz(df_obligo_relevant)
    ikz_kst = get_ikz(df_kst_relevant)

    if not (ikz_stammdaten == ikz_budget == ikz_obligo == ikz_kst):
        raise Exception("Die IKZ-Werte der vier Input Dateien stimmen nicht überein!")
    ikz = ikz_stammdaten

    # Bereinigen und Gruppieren nach Ergebnissen pro Jahr
    df_budget_relevant = df_budget_relevant.groupby(['PSP', 'PSPName', 'Jahr']).sum(numeric_only=True).reset_index()
    df_obligo_relevant = df_obligo_relevant.groupby(['PSP', 'PSPName', 'Jahr']).sum(numeric_only=True).reset_index()
    df_kst_relevant = df_kst_relevant.groupby(['PSP', 'PSPName', 'Jahr']).sum(numeric_only=True).reset_index()

    # Prüfen, ob mehr als ein Jahr enthalten ist
    check_jahr(df_budget_relevant)
    check_jahr(df_kst_relevant)

    # Stammdaten mergen
    df_budget_merged = pd.merge(df_budget_relevant,
                                df_stammdaten_relevant,
                                on=['PSP', 'PSPName'])[['PSP', 'PSPName', 'Status', 'Projektende', 'Geldgeber', 'Jahr',
                                                        'Budgetrest aus Vorjahr', 'Originalbudget',
                                                        'Sonstige Zuweisungen']]
    df_obligo_merged = pd.merge(df_obligo_relevant,
                                df_stammdaten_relevant,
                                on=['PSP', 'PSPName'])[['PSP', 'PSPName', 'Status', 'Projektende', 'Geldgeber', 'Jahr',
                                                        'Festlegungen']]
    df_kst_merged = pd.merge(df_kst_relevant,
                             df_stammdaten_relevant,
                             on=['PSP', 'PSPName'])[['PSP', 'PSPName', 'Status', 'Projektende', 'Geldgeber', 'Jahr',
                                                     'Einnahmen ILA', 'Einnahmen-Ist', 'Eigen- und Industrieanteile',
                                                     'Ausgaben-Ist', 'Kontostand Jahr']]

    # Drittmittelkontostand, Budget und Obligo mergen
    df_budget_obligo_merged = pd.merge(df_budget_merged,
                                       df_obligo_merged,
                                       how='outer',
                                       on=['PSP', 'PSPName', 'Status', 'Projektende', 'Geldgeber', 'Jahr']
                                       ).fillna(0)
    df_budget_kst_merged = pd.merge(df_budget_obligo_merged,
                                    df_kst_merged,
                                    how='outer',
                                    on=['PSP', 'PSPName', 'Status', 'Projektende', 'Geldgeber', 'Jahr']
                                    )
    df_budget_kst_merged['PA'] = df_budget_kst_merged['PSP'].str[3:5]
    df_budget_kst_merged['Projektende'] = pd.to_datetime(df_budget_kst_merged['Projektende'], format='%d.%m.%Y')

    # Berechne die kumulative Summe nur für PSP-Elemente mit Einträgen in "Kontostand Jahr"
    def calculate_cumsum(group):
        if group['Kontostand Jahr'].notna().any():
            group['End Kontostand DM'] = group['Kontostand Jahr'].fillna(0).cumsum()
        else:
            group['End Kontostand DM'] = np.nan
        return group

    df_budget_kst_merged = df_budget_kst_merged.groupby('PSP').apply(calculate_cumsum).reset_index(drop=True)

    # Budgetrest aus Vorjahr ein Jahr nach vorne schieben, Obligo abziehen und im Jahr davor als End-Kontostand angeben
    def shift_kontostand(group):
        group['End Kontostand Budget'] = group['Budgetrest aus Vorjahr'].shift(-1) - group['Festlegungen']
        group.at[group.index[-1], 'End Kontostand Budget'] = 0
        return group

    df_budget_kst_merged = df_budget_kst_merged.groupby('PSP').apply(shift_kontostand).reset_index(drop=True)

    # Erzeuge die neue Spalte 'Kontostand' nach dem Prinzip: wenn es ein Kontostand aus Drittmitteln gibt, nimm den,
    # ansonsten den Kontostand, der aus dem Budget erzeugt wurde
    def choose_kontostand(group):
        if group['End Kontostand DM'].notna().any():
            group['Kontostand'] = group['End Kontostand DM']
        else:
            group['Kontostand'] = group['End Kontostand Budget']
        return group

    df_budget_kst_merged = df_budget_kst_merged.groupby('PSP').apply(choose_kontostand).reset_index(drop=True)

    # Auswahl der Ausdrucke im Detailbericht
    prt = ['PSP', 'PSPName', 'PA', 'Status', 'Geldgeber', 'Projektende', 'Jahr', 'Budgetrest aus Vorjahr',
           'Originalbudget', 'Sonstige Zuweisungen', 'Festlegungen', 'End Kontostand Budget', 'Einnahmen-Ist',
           'Einnahmen ILA', 'Eigen- und Industrieanteile', 'Ausgaben-Ist', 'Kontostand Jahr', 'End Kontostand DM',
           'Kontostand']

    # finales Dataframe für Ausgabe filtern
    df_budget_kst_merged = (df_budget_kst_merged[prt].
                            sort_values(by=['PA', 'Projektende', 'PSP', 'Jahr', 'Status'],
                                        ascending=[True, True, True, True, True]))

    # beendete Projekte entfernen
    if config['rm_beendet']:
        df_budget_kst_merged = not_cont(df_budget_kst_merged, 'Status', 'beendet')

    # aktuelles Jahr entfernen wenn verlangt
    if config['rm_current_year']:
        df_budget_kst_merged = rem_current_year(df_budget_kst_merged)

    # Rohdaten schreiben, wenn gewünscht
    if config['prt_raw']:
        write_csv(df_budget_merged, ikz + '_Budget.csv')
        write_csv(df_kst_merged, ikz + '_Drittmittelkontostand.csv')
        write_csv(df_budget_kst_merged, ikz + '_Kombi_Budget_Drittmittelkontostand.csv')

    # --- Datensatz verfremden für Testzwecke, wenn "obfuscated"-Flag gesetzt
    if config['obfuscated']:
        # Funktion zum Hinzufügen von Rauschen
        def add_noise_to_numbers(df):
            for column in df.select_dtypes(include=[np.number]).columns:
                noise = np.random.uniform(-0.25, 0.25, df[column].shape)
                df[column] = df[column] * (1 + noise)
            return df

        # Funktion zum Verschleiern der PSP Elemente
        def obfuscate_psp(df):
            df['PSP'] = df['PSP'].apply(lambda x: x[:5] + '000000' + x[11:])
            df['PSPName'] = df['PSPName'].apply(lambda x: 'x' * len(x))
            return df

        # Verschleierten Datensatz zurückliefern
        return "000000", obfuscate_psp(add_noise_to_numbers(df_budget_kst_merged))
    else:
        # nicht-verfremdeten Datensatz zurückliefern
        return ikz, df_budget_kst_merged


# Daten zum Detailplot extrahieren
def import_detail_plot(df, fn_detailplot, lst):
    if os.path.exists(fn_detailplot):
        cv_detailplot = {0: 'str'}
        df_detailplot = load_csv_with_dynamic_header(fn_detailplot, 1, cv_detailplot)
        for index, row in df_detailplot.iterrows():
            result = df[df['PSP'] == row[0]]
            if not result.empty:
                res = [df, result['PSP'].iloc[0], f"{result['PSPName'].iloc[0]} ({result['PSP'].iloc[0]})", False]
                lst.append(res)
            else:
                print(f"Warnung: PSP {row[0]} in der Datei {fn_detailplot} ignoriert, da es nicht im SAP-Auszug ist.")


# Alle Projektarten im Datensatz zu der Auswertung hinzufügen
def import_pa_sap(df, lst):
    pa = df.groupby(['PA']).last().reset_index()[['PA']]
    for index, row in pa.iterrows():
        res = [df, PABericht.pa_pattern(row['PA']), f"Projektart {row['PA']}"]
        lst.append(res)


# Matplotlib Diagramm erstellen
def plot_pa(df, filename, title):
    if df.empty:
        return False

    # Setze das Jahr als Index
    df = df.set_index('Jahr')

    # Numerische Werte für den Index (Jahr) für die Regression
    x = np.arange(len(df))
    y = df['Kontostand']

    # Diagramm erstellen
    plt.figure(figsize=(10, 5))

    # Liniendiagramm für den Kontostand
    plt.plot(df.index, df['Kontostand'], marker='o', color='#00549F', label='Kontostand')

    # Fläche unter der Linie füllen
    plt.fill_between(df.index, df['Kontostand'], color='#00549F', alpha=0.25)

    # Lineare Regression als Gerade hinzufügen (gestrichelt)
    if len(df) > 1:
        reg_coeff = np.polyfit(x, y, deg=1)
        poly_eqn = np.poly1d(reg_coeff)
        y_pred = poly_eqn(x)
        slope = locale.format_string('%.2f €/J', reg_coeff[0], grouping=True)  # Steigung der Regressionsgeraden
        plt.plot(df.index, y_pred, '--', color='#F6A800', lw=2.0, label=f'Lin. Regression ({slope})')

    # Mittelwert-Linie (punkt-gestrichelt)
    mean_value = y.mean()
    mean_value_form = locale.format_string('%.2f €', mean_value, grouping=True)
    plt.plot([df.index[0], df.index[-1]], [mean_value, mean_value], '-.', color='#BDCD00', lw=2.0,
             label=f'Mittelwert ({mean_value_form})')

    if title:
        plt.title(title)
    plt.xlabel("Jahr")
    plt.ylabel("Kontostand [€]")
    plt.legend()
    plt.savefig(filename)
    plt.close()

    return True


# Funktion zum Aggregieren der Daten nach Projekten
def agg_proj(df):
    # Daten nach Projekten gruppieren und letzten Wert für die Kontostände nehmen
    df1 = df.groupby(['PSP', 'PSPName', 'Status', 'Projektende', 'Geldgeber']).last().reset_index()
    return df1[['PSP', 'PSPName', 'PA', 'Status', 'Projektende', 'Geldgeber', 'End Kontostand Budget',
                'End Kontostand DM', 'Kontostand']].sort_values(by=['PA', 'Projektende', 'PSP', 'Status'],
                                                                ascending=[True, True, True, True])


# Text Reports schreiben
class TXTReport:
    def __init__(self, filename):
        self.txt = open(filename, 'w')

    def append(self, text):
        self.txt.write(text)

    # Überschrift schreiben
    def append_title(self, title):
        self.txt.write("-" * len(title) + f"\n{title}\n" + "-" * len(title) + "\n\n")

    def signature_lines(self, ikz):
        self.append("Ich habe diesen Bericht gesehen und zur Kenntnis genommen. Bei eventuellen Unklarheiten habe ich "
                    "mich vor Unterschrift mit dem Dekanat abgestimmt.\n\n\n\n")
        self.append("______________________________________________________\n")
        self.append(f"Datum und Unterschrift der Leitung der IKZ {ikz}\n")

    def finalize(self):
        self.txt.close()


# PDF Reports basierend auf reportlab schreiben
class PDFReport:
    def __init__(self, filename, ikz):
        self.pdf = SimpleDocTemplate(filename=filename, pagesize=A4, leftMargin=2 * cm,
                                     rightMargin=2 * cm, topMargin=3 * cm, bottomMargin=2 * cm)
        self.ikz = ikz
        self.styles = getSampleStyleSheet()
        self.pdf_elements = list()
        self.rwth_tab_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#00549F")),  # HKS 44 - 100 % (Header)
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Weißer Text für den Header
            ('ALIGN', (0, 1), (-1, -1), 'RIGHT'),  # Rechtsbündige Ausrichtung aller Werte außer Header
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#C7DDF2")),
            # Einheitliche Hintergrundfarbe ab der zweiten Zeile
            ('GRID', (0, 0), (-1, -1), 1, colors.black)  # Schwarze Gitterlinien
        ]

    # Funktion für den Header auf jeder Seite
    def lscontrolling_brand(self, canvas, doc):
        # Seitenbreite und Ränder
        page_width, page_height = A4
        margin_top = 1 * cm  # von der Oberkante
        margin_left = doc.leftMargin  # Verwendung der vom Dokument definierten linken Marge
        margin_right = doc.rightMargin  # Verwendung der vom Dokument definierten rechten Marge

        # SVG-Logo laden
        drawing = svg2rlg(io.BytesIO(base64.b64decode(lscontrolling_logo)))

        # Berechnung für Skalierung des SVG
        logo_width = 6 * cm  # maximale Breite des Logos
        scale_factor = logo_width / drawing.width
        logo_height = drawing.height * scale_factor

        # Logo rechtsbündig platzieren
        logo_x = page_width - margin_right - logo_width
        logo_y = page_height - margin_top - logo_height

        # Logo rendern (skalieren und platzieren)
        canvas.saveState()  # Zustand speichern
        canvas.translate(logo_x, logo_y)  # Positionierung
        canvas.scale(scale_factor, scale_factor)  # Skalierung
        renderPDF.draw(drawing, canvas, 0, 0)
        canvas.restoreState()  # Zustand wiederherstellen

        # Text linksbündig platzieren
        canvas.setFont("Helvetica-Bold", 18)
        text = f"Controlling der IKZ {self.ikz}"

        # Der Text wird an der Unterkante des Logos ausgerichtet
        text_x = margin_left + 5
        text_y = logo_y  # Position des Textes auf der Höhe der Unterkante des Logos

        # Text vom linken Rand platzieren
        canvas.drawString(text_x, text_y, text)

        # Datum in den Footer stellen
        footer_text = ("Bericht erzeugt am: " + datetime.now().strftime("%d.%m.%Y") +
                       " | Programmversion " + program_version)
        footer_x = doc.leftMargin + 5
        footer_y = doc.bottomMargin - 1 * cm

        canvas.setFont("Helvetica", 10)
        canvas.drawString(footer_x, footer_y, footer_text)

    # Funktion zum Einfärben der Beträge der Zusammenfassungstabelle
    @staticmethod
    def zusammenfassungstabelle_farbe(table_data):
        # Initialisiere eine leere Liste für die Stile
        style = []
        total_sum = 0

        positive_green = 0
        negative_red = 0
        negative_orange = 0

        # Konvertiert einen Euro-Betrag von String in eine Zahl zurück
        def parse_euro_amount(amount_str: str) -> float:
            try:
                return float(amount_str.replace('.', '').replace(',', '.').replace('€', '').strip())
            except ValueError:
                return 0

        # Durchlaufe alle Zellen in der Tabelle und prüfe ihren Inhalt
        for row_idx, row in enumerate(table_data[:-1]):  # Die letzte Zeile (Summe) wird später behandelt.
            projektart, bemerkung, kontostand_str = row

            # Konvertiere den Kontostand-String in eine Zahl
            kontostand = parse_euro_amount(kontostand_str)

            if isinstance(kontostand, float):  # Prüfe, ob es sich um eine Zahl handelt
                if 'vor' in bemerkung.lower() or 'alle' in bemerkung.lower():
                    if kontostand < 0:
                        style.append(('TEXTCOLOR', (2, row_idx), (2, row_idx), colors.red))
                        negative_red += kontostand
                    elif kontostand > 0:
                        style.append(('TEXTCOLOR', (2, row_idx), (2, row_idx), colors.green))
                        positive_green += kontostand

                elif 'nach' in bemerkung.lower():
                    if kontostand < 0:
                        style.append(('TEXTCOLOR', (2, row_idx), (2, row_idx), colors.orange))
                        negative_orange += kontostand
                    elif kontostand > 0:
                        style.append(('TEXTCOLOR', (2, row_idx), (2, row_idx), colors.green))
                        positive_green += kontostand

                total_sum += kontostand

        # Einfärben der finalen Summe
        if total_sum > 0:
            style.append(('TEXTCOLOR', (2, len(table_data) - 1), (2, len(table_data) - 1), colors.green))
        else:
            if total_sum - negative_orange >= 0:
                style.append(('TEXTCOLOR', (2, len(table_data) - 1), (2, len(table_data) - 1), colors.orange))
            else:
                style.append(('TEXTCOLOR', (2, len(table_data) - 1), (2, len(table_data) - 1), colors.red))

        return style

    def signature_lines(self, ikz):
        p1 = Paragraph(
            "Ich habe diesen Bericht gesehen und zur Kenntnis genommen. Bei eventuellen Unklarheiten habe ich "
            "mich vor Unterschrift mit dem Dekanat abgestimmt.", self.styles['Normal'])
        p1.spaceAfter = 1.8 * cm
        p2 = Paragraph("______________________________________________________", self.styles['Normal'])
        p3 = Paragraph(f"Datum und Unterschrift der Leitung der IKZ {ikz}", self.styles['Normal'])
        self.append(KeepTogether([p1, p2, p3]))

    def append(self, element):
        self.pdf_elements.append(element)

    def append_title(self, title):
        self.pdf_elements.append(Paragraph(title, self.styles['Heading1']))

    def append_title2(self, title):
        self.pdf_elements.append(Paragraph(title, self.styles['Heading2']))

    def finalize(self):
        self.pdf.build(self.pdf_elements, onFirstPage=self.lscontrolling_brand, onLaterPages=self.lscontrolling_brand)


class PABericht:
    def __init__(self, txt=None, pdf=None, tmp=None):
        self.txt = txt
        self.pdf = pdf
        self.tmp = tmp
        self.summary = pd.DataFrame(columns=['Projektart', 'Bemerkung', 'Kontostand'])

    # Funktion zum Erzeugen einer Suchmaske basieren auf der Projektart
    @staticmethod
    def pa_pattern(pa):
        return fr'^\d{{3}}{pa}\d{{10}}$'

    # Schreiben der Projektarten oder PSP-Elemente in die Berichte
    def pa_auflistung(self, df, pattern, title, sum_up=True):
        # Datensatz nach gewünschter PA filtern
        if pattern:
            df = cont(df, 'PSP', pattern)

        # Gruppierung nur für das Jahr erzeugen
        gdf = df.groupby('Jahr').sum(numeric_only=True).reset_index()

        # wenn was übrig ist, dann alles Schreiben und Daten für Zusammenfassung berechnen
        if len(gdf):
            dc = dict()
            tit = title.split('|')
            dc['Projektart'] = tit[0]
            if len(tit) > 1:
                dc['Bemerkung'] = tit[1]
            else:
                dc['Bemerkung'] = ''
            dc['Kontostand'] = round(gdf.iloc[-1]['Kontostand'], 2)

            # Ergebnis für Zusammenfassung in Instanz zwischenspeichern
            if sum_up:
                self.summary.loc[len(self.summary)] = dc

            # Ergebnisse im Detail schreiben
            if self.txt:
                self.txt.append_title(title)
                cdf = gdf.copy(deep=True)
                for col in cdf.select_dtypes(include=['number']):
                    cdf[col] = cdf[col].apply(lambda x: locale.format_string('%.2f €', x, grouping=True))
                self.txt.append(cdf.to_string(index=False, float_format=lambda x: f'{x:.2f}') + "\n\n")

            if self.pdf:
                p_tit = Paragraph(title, self.pdf.styles['Heading3'])

                img = Paragraph("")  # leeren Abschnitt als Bildersatz nutzen.

                # Plot für PDF Bericht vorbereiten (wenn ein temporäres Verzeichnis gegeben wurde)
                if self.tmp:
                    fn = self.tmp.temp_dir + f"/{urllib.parse.quote(pattern, safe='')}.svg"
                    if not plot_pa(gdf[['Jahr', 'Kontostand']], fn, title):
                        return False

                    # Plot in PDF integrieren
                    page_width, page_height = A4
                    image_width = page_width * 0.85
                    matplotlib.use('svg')  # Stellt sicher, dass das SVG-Backend verwendet wird
                    img = svg2rlg(fn)
                    scale_factor = image_width / img.width
                    img.width = image_width
                    img.height *= scale_factor
                    img.scale(scale_factor, scale_factor)

                # Tabelle Kontostand
                df_pdf = gdf[['Jahr', 'Kontostand']].copy(deep=True)
                df_pdf['Kontostand'] = df_pdf['Kontostand'].apply(
                    lambda x: locale.format_string('%.2f €', x, grouping=True))
                table_data = [df_pdf.columns.tolist()] + df_pdf.values.tolist()
                tab = Table(table_data)
                tab.setStyle(self.pdf.rwth_tab_style)
                tab.spaceAfter = 1 * cm
                tab.spaceBefore = 1 * cm

                # Mittelwert
                mean_value = locale.format_string('%.2f €', gdf['Kontostand'].mean(), grouping=True)
                mean_style = self.pdf.styles["BodyText"]
                mean_style.alignment = TA_CENTER
                mean_style.fontName = "Helvetica-BoldOblique"  # Fett und Kursiv
                mean = Paragraph(f"Mittelwert: {mean_value}", mean_style)

                # Alles in den Bericht packen
                self.pdf.append(KeepTogether([p_tit, img]))  # Tabelle und Mittelwert im Bericht weglassen
                # self.pdf.append(KeepTogether([p_tit, img, tab, mean]))

    # Schreiben der Details in den Textbericht (im PDF ist das nicht integriert, da es zu detailliert ist)
    def detail(self, df, title):
        if self.txt:
            self.txt.append(f"\n")
            self.txt.append_title(title)
            for col in df.select_dtypes(include=['number']):
                df[col] = df[col].apply(lambda x: locale.format_string('%.2f €', x, grouping=True
                                                                       ) if not pd.isna(x) else 'k.A.')
            self.txt.append(df.to_string(index=False, float_format=lambda x: f'{x:.2f}') + "\n\n")

    # Schreiben der Zusammenfassung
    def zusammenfassung(self, title):
        # Summe der Zusammenfassung berechnen
        sums = self.summary.sum(numeric_only=True)
        sums[self.summary.columns[0]] = 'Summe'
        sums[self.summary.columns[1]] = ''
        summary = pd.concat([self.summary, sums.to_frame().T], ignore_index=True)
        summary['Kontostand'] = summary['Kontostand'].apply(
            lambda x: locale.format_string('%.2f €', x, grouping=True))

        if self.txt:
            self.txt.append(f"\n")
            self.txt.append_title(title)
            self.txt.append(summary.to_string(index=False))
            self.txt.append("\n\n")

        if self.pdf:
            table_data = [summary.columns.tolist()] + summary.values.tolist()
            tab = Table(table_data)
            tab.setStyle(self.pdf.rwth_tab_style)
            number_styles = PDFReport.zusammenfassungstabelle_farbe(table_data)
            tab.setStyle(TableStyle(number_styles))
            tab.spaceBefore = 1 * cm
            tab.spaceAfter = 1.5 * cm
            par = Paragraph("Eigene Anmerkungen:", self.pdf.styles['Heading3'])
            par.spaceAfter = 8 * cm
            self.pdf.append(PageBreak())
            self.pdf.append(KeepTogether([
                Paragraph(title, self.pdf.styles['Heading2']),
                tab, par
            ]))
