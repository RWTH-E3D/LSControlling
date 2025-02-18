from datetime import datetime, timedelta
from funktionen import PABericht, RandomTemp, TXTReport, PDFReport, agg_proj, write_csv, import_sap_csv, \
    import_detail_plot, laufende_projekte_ignorieren, nur_sammelkonten, keine_sammelkonten, \
    nur_laufende_projekte, LSControllingConfig, LogContext

if __name__ == "__main__":
    try:
        # --- Datenimport und Preprocessing ----------------------------------------------------------------------------

        # Prüfen, ob eine Config Datei gegeben wurde; wenn ja, dann Werte aus Config nutzen, ansonsten Default Werte
        cfg = LSControllingConfig('config.ini')

        # Daten aus SAP importieren
        with LogContext("Datenimport und -bereinigung"):
            ikz, df_ikz, rep_dates = import_sap_csv(cfg)

        # Jahresspanne der Daten ermitteln
        min_jahr = df_ikz['Jahr'].min()
        max_jahr = df_ikz['Jahr'].max()

        # --- Datenfilter anwenden -------------------------------------------------------------------------------------

        with LogContext("Datenfilterung"):
            cut1 = datetime(int(max_jahr), 6, 30)  # 30. Juni des letzten Jahres als Cutoff nutzen
            cut2 = cut1 + timedelta(days=1)
            df_ikz_sk = nur_sammelkonten(df_ikz)  # nur Sammelkonten
            df_ikz_ek_alle = keine_sammelkonten(df_ikz)  # alle Einzelkonten aber keine Sammelkonten
            df_ikz_ek_abgelaufen = laufende_projekte_ignorieren(df_ikz_ek_alle, cut1)  # abgelaufene E.konten vor cutoff
            df_ikz_ek_laufend = nur_laufende_projekte(df_ikz_ek_alle, cut1)  # nur laufende E.konten nach cutoff

        # --- Datenauswertung ------------------------------------------------------------------------------------------

        # zufälliges temporäres Verzeichnis erzeugen
        tmp = RandomTemp()

        with LogContext(f"Berichtsinstanzen für IKZ {ikz} erzeugen"):
            # Text-Bericht Instanz erzeugen
            txt = TXTReport(f"{ikz}_Bericht.txt")
            txt.append_title(f"Finanzübersicht {min_jahr} - {max_jahr} für die IKZ {ikz}")

            # PDF-Bericht Instanz erzeugen
            pdf = PDFReport(f"{ikz}_Bericht.pdf", ikz)
            pdf.append_title(f"Finanzübersicht {min_jahr} - {max_jahr} für die IKZ {ikz}")

            # Instanz für Berichtsinhalt erzeugen
            bericht = PABericht(txt=txt, pdf=pdf, tmp=tmp)

            # Ermitteln der einzelnen Sub-Positionen für den Bericht und Schreiben der Ergebnisse in eine Datei
            txt.append("Kontostände nach Projektart und Jahr in Euro\n\n")
            pdf.append_title2("Kontostände nach Projektart und Jahr in Euro")

        with LogContext("Festlegen der relevanten Projektarten"):
            # Definieren der relevanten Projektarten für den Bericht
            pa_rel = []
            for pa in cfg['liste_pa_aufteilung']:
                pa_rel.append([df_ikz_sk, bericht.pa_pattern(pa),
                               f"Projektart {pa} | Sammelkonten (alle)", True])
                pa_rel.append([df_ikz_ek_abgelaufen, bericht.pa_pattern(pa),
                               f"Projektart {pa} | Einzelkonten (Projektende vor {cut1.strftime('%d.%m.%y')})", True])
                pa_rel.append([df_ikz_ek_laufend, bericht.pa_pattern(pa),
                               f"Projektart {pa} | Einzelkonten (Projektende nach {cut2.strftime('%d.%m.%y')})", True])
            for pa in cfg['liste_pa_keine_aufteilung']:
                pa_rel.append([df_ikz, bericht.pa_pattern(pa), f"Projektart {pa} | Alle Konten", True])

            # Prüfen, ob ein Detailplot integriert werden soll, wenn ja, pa_rel erweitern
            import_detail_plot(df_ikz, cfg['csv_detailplot'], pa_rel)

        # Erzeugen der Berichtsdaten für die relevanten Projektarten und Projekte
        for pa in pa_rel:
            with LogContext(f"Erzeugung Sub-Bericht {pa[2]}"):
                bericht.pa_auflistung(*pa)

        # Zusammenfassung schreiben
        with LogContext(f"Erzeugung der Zusammenfassung für IKZ {ikz}"):
            bericht.zusammenfassung(f"Zusammenfassung für IKZ {ikz} (Stand 31.12.{max_jahr})")

        # Details nach Projekt in CSV und Textbericht schreiben
        with LogContext("Erzeugung der Projektdetailansichten"):
            ap = agg_proj(df_ikz)
            write_csv(ap, ikz + '_Projektansicht.csv')
            bericht.detail(ap, f"Details nach Projekt für IKZ {ikz} (Stand 31.12.{max_jahr})")

        # Berichtsdateien finalisieren schließen
        with LogContext("Finalisieren des Berichtes"):
            txt.signature_lines(ikz)
            txt.berichts_info(rep_dates)
            txt.finalize()
            pdf.signature_lines(ikz)
            pdf.berichts_info(rep_dates)
            pdf.finalize()

            # Temporäres Verzeichnis löschen
            tmp.delete_temp_dir()

    except Exception as e:
        print(f"FEHLER: {e}\n")
        input("Bitte eine beliebige Taste drücken zum Beenden.")
