# Berichtsheft-Generator

Ein kleines Python-Projekt aus meinem Ausbildungsalltag: Das Skript bereitet den wöchentlichen Ausbildungsnachweis automatisch vor und verschickt ihn per E-Mail, statt dass die Vorlage jede Woche von Hand ausgefüllt wird.

## Was es macht

- Ermittelt aus dem aktuellen Datum die Kalenderwoche, Wochenanfang und -ende, das Ausbildungsjahr und die fortlaufende Berichtsnummer
- Erkennt anhand einer hinterlegten Wochenliste, ob Berufsschule ansteht, und passt die Stundenangaben entsprechend an
- Schlägt eine Aufteilung der Tagesstunden vor, die vor der Abgabe noch inhaltlich ergänzt wird
- Setzt alle Werte in die Platzhalter einer Word-Vorlage ein (`python-docx`)
- Verschickt das fertige Dokument als Anhang per SMTP an die eingetragenen Empfänger

Läuft bei mir als wöchentlicher Cronjob.

## Setup

```bash
pip install python-docx
```

Zugangsdaten kommen aus den Umgebungsvariablen `MAIL_USER` und `MAIL_PASSWORD` (Gmail-App-Passwort). Empfängerliste, Berufsschulwochen und Dateipfade werden oben im Skript gesetzt.

```bash
python berichtsheft.py
```

## Hinweis

Ein privates Hilfsprojekt, kein fertiges Produkt. Es tut zuverlässig das, wofür ich es gebaut habe, an der Struktur gibt es aber noch einiges zu verbessern.
