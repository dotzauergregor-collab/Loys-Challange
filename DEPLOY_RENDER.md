# LOYS Aktien-Challenge live auf Render stellen

## Warum Render?
Die Challenge hat einen Python-Server und schreibt Einsendungen in eine Excel-Datei. Deshalb reicht reines Static Hosting nicht. Dieses Paket ist so vorbereitet, dass die Excel-Datei auf einem persistenten Render-Datenträger unter `/var/data` gespeichert wird.

## 1. GitHub-Repository anlegen
1. Auf GitHub ein neues privates Repository anlegen, z. B. `loys-aktien-challenge`.
2. Den **Inhalt dieses Ordners** in das Repository hochladen. `render.yaml` muss im Hauptverzeichnis des Repositorys liegen.
3. Änderungen committen.

## 2. Render einrichten
1. Bei https://render.com anmelden.
2. `New` -> `Blueprint` auswählen.
3. GitHub verbinden und das Repository auswählen.
4. Render erkennt `render.yaml` automatisch.
5. Prüfen, dass die Region **Frankfurt** ist.
6. Deployment starten.

Das Blueprint legt an:
- einen Python-Webservice
- Build: `pip install -r requirements.txt`
- Start: `python server.py`
- einen persistenten Datenträger `/var/data`
- die Excel-Datei wird beim ersten Start automatisch dorthin kopiert

Nach erfolgreichem Deployment erhältst du eine URL ähnlich:
`https://loys-aktien-challenge.onrender.com`

## 3. Eigene Domain empfehlen
Für den QR-Code ist eine dauerhafte Domain besser, zum Beispiel:
`https://challenge.loys.de`

In Render:
1. Webservice öffnen
2. `Settings` -> `Custom Domains`
3. `challenge.loys.de` hinzufügen
4. Den von Render angezeigten DNS-Eintrag bei eurem Domain-/DNS-Anbieter setzen
5. Domain in Render verifizieren

Erst danach den finalen QR-Code erzeugen. So bleibt der QR-Code auch dann gleich, wenn der Hosting-Anbieter später gewechselt wird.

## 4. QR-Code
Der QR-Code enthält ausschließlich die finale URL, beispielsweise:
`https://challenge.loys.de`

Den QR-Code anschließend mit Smartphone-Kameras (iPhone + Android) testen.

## Wo liegen die Einsendungen?
Im Livebetrieb auf Render:
`/var/data/LOYS_Aktien_Challenge.xlsx`

Wichtig: Nur der Inhalt unter `/var/data` ist dauerhaft gespeichert.

## Excel-Datei herunterladen
Die Datei liegt auf dem persistenten Server-Datenträger. Für einen komfortablen laufenden Download sollte in einer nächsten Version ein passwortgeschützter Admin-Download ergänzt werden. Alternativ kann die Datei über Render-Shell/SSH abgerufen werden.

## Vor öffentlichem Einsatz
- Datenschutzerklärung und Teilnahmebedingungen finalisieren
- Bot-/Spam-Schutz ergänzen
- Backup-/Download-Prozess für die Excel-Datei definieren
- bei sehr vielen gleichzeitigen Einsendungen besser Datenbank als Primärspeicher nutzen und Excel daraus exportieren
