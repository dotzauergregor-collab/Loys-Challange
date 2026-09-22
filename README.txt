LOYS Aktien-Challenge - Komplettpaket mit Excel-Anbindung
========================================================

Titel
-----
4 Aktien, 1 Jahr - Wer wird der Top-Performer?

Spielmechanik
-------------
1. Jeder Teilnehmer trägt drei unterschiedliche Aktien ein, bei denen er für das
   kommende Jahr eine positive Kursentwicklung erwartet.
2. Zusätzlich wird eine vierte, ebenfalls unterschiedliche Aktie eingetragen,
   bei der der Teilnehmer im kommenden Jahr fallende Kurse erwartet.
3. Aktienname und Börsenkürzel/Ticker werden frei eingegeben. Es gibt keine
   vorgegebene Auswahlliste.
4. Die Magnificent Seven sind ausgeschlossen: Apple, Microsoft, NVIDIA,
   Alphabet/Google, Amazon, Meta Platforms und Tesla.
5. Die Mag-7-Sperre wird sowohl im Browser als auch serverseitig anhand von
   Aktienname und Ticker geprüft.
6. Alle vier Aktien müssen unterschiedlich sein. Der Server prüft Namen und
   Ticker auf doppelte Einträge.
7. Nach der verbindlichen Abgabe werden Name, E-Mail, alle vier Aktiennamen und
   alle vier Ticker direkt in Excel gespeichert.
8. Doppelte E-Mail-Adressen werden serverseitig blockiert.
9. Auf der Website gibt es kein Live-Leaderboard.

Mobile Weboberfläche
--------------------
Die Website ist mobile-first aufgebaut:
- einspaltige Ansicht auf Smartphones
- große Touch-Flächen und mindestens 48 px hohe Eingabefelder
- pro Aktie eine eigene Karte mit Aktienname + Ticker
- große, gut erreichbare Absende-Schaltfläche
- keine horizontalen Tabellen oder Dropdowns
- Desktop-Ansicht passt sich ab größeren Bildschirmbreiten automatisch an

Excel-Arbeitsblätter
---------------------
- Teilnahmen: Teilnehmerdaten mit 3 positiven Ideen + 1 Negativ-These
- Konfiguration: Challenge-Parameter und Spielregeln
- Ausschlüsse: dokumentierte Mag-7-Sperrliste
- Marktdaten: leere Vorlage für eine spätere Performance-Auswertung
- Auswertung: interne Excel-Übersicht; kein Web-Leaderboard

Start unter Windows
-------------------
1. Python installieren, falls noch nicht vorhanden.
2. ZIP-Datei entpacken.
3. start_windows.bat doppelklicken.
4. Im Browser http://127.0.0.1:8000 öffnen.

Start unter macOS / Linux
-------------------------
1. Terminal im entpackten Ordner öffnen.
2. chmod +x start_mac_linux.sh
3. ./start_mac_linux.sh
4. Im Browser http://127.0.0.1:8000 öffnen.

Alternativ manuell
------------------
python -m pip install -r requirements.txt
python server.py

Wo landen die Daten?
---------------------
Alle bestätigten Teilnahmen werden in folgender Datei gespeichert:

data/LOYS_Aktien_Challenge.xlsx

WICHTIG: Die Excel-Datei während laufender Einsendungen möglichst nicht in
Excel geöffnet halten, da Windows die Datei sperren kann.

Technischer Hinweis
-------------------
Dieser Prototyp schreibt direkt in eine XLSX-Datei auf dem Server. Für einen
öffentlichen Wettbewerb mit vielen gleichzeitigen Zugriffen ist eine Datenbank
als Primärspeicher robuster; Excel kann dann automatisch als Reporting-/Export-
Datei erzeugt werden.

Vor Livegang sollten Datenschutz, Teilnahmebedingungen, Hosting, Backups,
SSL/HTTPS sowie Spam- und Bot-Schutz finalisiert werden.

Live-Deployment
---------------
Diese Version ist zusätzlich für Render vorbereitet. Siehe DEPLOY_RENDER.md.
Im Cloudbetrieb wird HOST/PORT automatisch über Umgebungsvariablen gesetzt und die
Excel-Datei auf einem persistenten Datenträger unter /var/data gespeichert.
