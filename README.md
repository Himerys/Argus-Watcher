# Argus – Active Directory Export Tool (LDAP Edition)

Grafisches Tool (WPF) **und** Kommandozeilen-Tool zum Auslesen und Exportieren von
Active-Directory-Benutzerobjekten nach **Excel** und/oder **CSV** – über **reines LDAP**
(Port 389/636/3268/3269).

Argus benötigt bewusst **kein ADWS** (Port 9389) und **kein RSAT / ActiveDirectory-Modul**.
Es läuft damit auch in Umgebungen, in denen ADWS blockiert ist oder RSAT nicht installiert
werden darf. Der CSV-Export hat **keinerlei Abhängigkeiten**; nur für den Excel-Export wird
das PowerShell-Modul [`ImportExcel`](https://github.com/dfinke/ImportExcel) benötigt.

## Funktionsumfang

| Bereich | Funktionen |
|---|---|
| **Quelle** | Alle Benutzer · Mitglieder von Gruppe(n) (rekursiv, inkl. primärer Gruppe) · Liste spezifischer Benutzer (sAMAccountName/UPN/Mail/CN) mit Abgleichbericht nicht gefundener IDs |
| **Filter** | Dynamischer Filter-Builder mit UND/ODER-Verknüpfung, typgerechten Operatoren (Datum: „innerhalb/älter als N Tage", Status: wahr/falsch), Live-LDAP-Filter-Vorschau, RAW-Zusatzfilter für Power-User |
| **Spalten** | Attribut-Katalog mit Suchfeld · beliebige zusätzliche LDAP-Attribute (mit Schema-Prüfung) · abgeleitete Felder (Aktiviert, Passwort läuft nie ab, SID, GUID) · True/False-Spalten für Gruppenmitgliedschaften |
| **Vorschau** | Ergebnis-Vorschau im DataGrid **vor** dem Export (sortierbar, kopierbar) |
| **Export** | Excel (formatierte Tabelle, Info-Blatt mit Audit-Metadaten) · CSV (UTF-8 mit BOM, Trennzeichen wählbar) · Zeitstempel im Dateinamen · Überschreib-Schutz |
| **Verbindung** | LDAPS (636/3269) · Signing+Sealing auf 389 · Global Catalog (3268/3269, Forest-weit) · alternative Anmeldedaten · Verbindungstest · OU-Picker (Baumansicht) · Zeitlimit |
| **Betrieb** | Profile (JSON) speichern/laden · Headless-Modus (`-NoGui`) für die Aufgabenplanung · Protokolldatei (Audit-Trail) · Abbruch jederzeit · letzte Einstellungen werden gemerkt |

## Voraussetzungen

- Windows mit **PowerShell 5.1** oder **PowerShell 7**
- Domänenzugriff per LDAP (389) bzw. LDAPS (636) – **kein** ADWS, **kein** RSAT
- Nur für Excel-Export: Modul `ImportExcel` (`Install-Module ImportExcel -Scope CurrentUser`);
  das GUI bietet die Installation an und zeigt den Modulstatus. CSV funktioniert immer.

## Schnellstart (GUI)

```powershell
.\Argus.ps1
```

Empfohlener Ablauf entlang der Tabs:

1. **Verbindung** – DC/Port/Anmeldung wählen (leer = DC-Locator), optional Base-DN über den
   OU-Picker, dann „Verbindung testen".
2. **Quelle & Filter** – Benutzerquelle wählen, Bedingungen hinzufügen; die
   LDAP-Filter-Vorschau zeigt live den erzeugten Filter.
3. **Spalten** – Export-Felder wählen, optional Gruppen-Spalten.
4. **Suchen (F5)** – Ergebnis erscheint als Vorschau; erst danach **Exportieren**.

Für eine Verknüpfung ohne Konsolenfenster:

```
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Tools\Argus.ps1"
```

## Headless-Modus (Aufgabenplanung / Automatisierung)

Im GUI unter **Profil → Speichern…** ein Profil erzeugen, dann:

```powershell
# Profilgesteuert (empfohlen)
.\Argus.ps1 -NoGui -ProfilePath C:\Argus\Profile\hr-export.json -Force

# Ad hoc, ohne Profil
.\Argus.ps1 -NoGui -Mode group -Group "VPN-Benutzer" -Format csv `
            -Server dc01.firma.local -UseLdaps -OutputPath C:\Export

# Mit alternativen Anmeldedaten
.\Argus.ps1 -NoGui -ProfilePath .\export.json -Credential (Get-Credential) -Force
```

Explizit gesetzte Parameter überschreiben das Profil. **Exit-Codes:** `0` = Export
erstellt · `2` = keine Treffer (keine Datei) · `1` = Fehler. Vollständige Referenz:
`Get-Help .\Argus.ps1 -Full`

## Profile

Profile sind JSON-Dateien mit allen Einstellungen (Verbindung, Quelle, Filter, Spalten,
Export). **Kennwörter werden nie gespeichert.** Das GUI merkt sich außerdem die letzte
Sitzung unter `%APPDATA%\Argus\last-settings.json`.

## Protokollierung

Jeder Lauf protokolliert Filter, Suchbasis, Treffer, Dauer und Zieldateien – im GUI-Protokoll
und (abschaltbar) in `%APPDATA%\Argus\Logs\Argus_<Datum>.log` (`-LogPath` überschreibt).
Der Excel-Export enthält zusätzlich ein **Info-Blatt** mit denselben Audit-Metadaten
(wer, wann, von wo, welcher Filter).

## Sicherheit

- **Transport:** Auf Port 389 wird Signing+Sealing ausgehandelt (kompatibel mit der
  DC-Härtung „LDAP signing required"); LDAPS über 636/3269.
- **Injektionsschutz:** Zellen, die mit `=` beginnen, werden beim Export neutralisiert
  (Excel und CSV). Der optionale „strenge CSV-Schutz" neutralisiert zusätzlich führende
  `+`, `-`, `@` – Standard ist er aus, damit Telefonnummern (`+49 …`) lesbar bleiben.
- **Datentreue:** Excel-Export ohne Zahlen-Autokonvertierung – führende Nullen
  (PLZ `01067`, EmployeeID) bleiben erhalten.
- **Zugriffs-Gate:** Optional per `-RequiredGroup "DOMAIN\Gruppe"` (oder Konstante
  `$script:DefaultRequiredGroup` im Skript) auf eine AD-Gruppe beschränken.
- **Datenschutz:** Das Tool exportiert personenbezogene Daten. Export-Ziel bewusst wählen
  (Standard: Dokumente-Ordner, nicht `C:\Temp`), Dateien nach Gebrauch löschen.
- **Signierung:** Vor produktivem Rollout das Skript mit einem Code-Signing-Zertifikat
  signieren (`Set-AuthenticodeSignature`), damit es unter `AllSigned` läuft.

## Bekannte Grenzen

- **Letzte Anmeldung** basiert auf `lastLogonTimestamp` – das Attribut wird nur alle
  9–14 Tage repliziert und ist bewusst „ungefähr". Für exakte Werte müsste jeder DC
  einzeln abgefragt werden (nicht Ziel dieses Tools).
- **Konstruierte Attribute** (z. B. `msDS-UserPasswordExpiryTimeComputed`) liefert eine
  LDAP-Suche nicht; entsprechende Spalten bleiben leer.
- **„Konto läuft ab":** leere Zelle bedeutet „läuft nie ab / nicht gesetzt".
- **Constrained Language Mode:** Die WPF-Oberfläche benötigt Full Language Mode. In
  CLM-/WDAC-Umgebungen das signierte Skript entsprechend freigeben.
- Sehr große Exporte (>100k Benutzer) werden komplett im Speicher gehalten; ab ~20 000
  Zeilen wird die Excel-Spaltenbreiten-Automatik aus Performance-Gründen deaktiviert.

## Fehlerbehebung

| Symptom | Ursache / Lösung |
|---|---|
| „Der Server ist nicht verbunden" / Timeout | DC/Port prüfen, „Verbindung testen" nutzen; Firewall 389/636 |
| Suche scheitert sofort mit „Es ist ein Fehler bei der Ausführung aufgetreten" (operations error, 0x80072020) | Sitzung hat keine nutzbaren Domänen-Anmeldedaten (PowerShell-Remoting/Double-Hop, lokales Konto, Rechner nicht in der Domäne) – RootDSE ist anonym lesbar, daher „gelingt" der Bind scheinbar. Abhilfe: „Alternative Anmeldedaten" bzw. `-Credential` verwenden und den DC als FQDN (nicht IP) angeben. Seit v2.0.1 deckt „Verbindung testen" dies per Probesuche auf |
| Excel-Export scheitert, CSV geht | `ImportExcel` fehlt → Button im Ausgabe-Bereich oder Softwareverteilung |
| „Datei ist gesperrt" | Zieldatei ist in Excel geöffnet → schließen |
| Gruppenname „mehrdeutig" | Vollständigen DN der Gruppe angeben (Meldung listet die Kandidaten) |
| Benutzer fehlt trotz Gruppenmitgliedschaft | Nur primäre Gruppe? Wird seit v2 mitgeprüft – Version prüfen |
| Umlaute im CSV kaputt | Datei mit Excel öffnen (BOM wird gesetzt); Fremd-Tools auf UTF-8 stellen |

## Tests

Die reine Logik (LDAP-Escaping, Klausel-Builder, FILETIME-Behandlung, CSV-Writer,
Profil-Roundtrip) ist ohne Domäne testbar:

```powershell
powershell -File .\tests\Argus.Tests.ps1   # oder pwsh, läuft auch unter Linux
```

## Änderungen in v2.0.1

- „Verbindung testen" führt zusätzlich eine Probesuche (1 Objekt, Basis-DN) aus.
  Damit fallen Verbindungen ohne nutzbare Domänen-Anmeldedaten (Double-Hop/Remoting,
  lokales Konto) bereits beim Test mit einer verständlichen Meldung auf – statt erst
  bei der eigentlichen Suche mit einem generischen „operations error" (0x80072020).

## Änderungen in v2.0.0 (gegenüber dem ursprünglichen Skript)

- Neue Oberfläche: Tabs, Ergebnis-Vorschau, Menü, Verbindungstest, OU-Picker,
  Live-Filter-Vorschau, Abbruch-Button, Tastenkürzel, Fenster-/Sitzungs-Merken
- Neue Fälle: LDAPS/GC/alternative Anmeldedaten, Headless-Modus, Profile, CSV-Export,
  Datums-/Status-Filter, ODER-Verknüpfung, Gruppen-Spalten inkl. primärer Gruppe,
  Abgleich nicht gefundener Benutzer, Protokolldatei
- Korrektheit: RFC-4515-Escaping auch für Gruppen-DNs, primäre-Gruppe-Lücke geschlossen,
  Mehrfachwerte nicht mehr abgeschnitten, memberOf-Ranged-Retrieval (>1500 Gruppen),
  mehrdeutige Gruppennamen werden gemeldet statt geraten, Ressourcen-Freigabe auf allen
  Fehlerpfaden, Excel-Altdaten-Reste beim Überschreiben beseitigt, Formel-Injektion und
  Verlust führender Nullen im Export behoben

---
**Owner:** Semih Danyeri · **Kontakt:** Microsoft Teams → „Semih Danyeri"
