# Xray Excel-Testausfuehrung

Eigenstaendiges Paket fuer den Excel-Export von Xray-Testausfuehrungen.

## Start

```powershell
.\start-excel-export.bat -IssueKeys ANW-123
.\Export-XrayTestausfuehrungen.ps1 -FilterUrl 'https://jira.example.invalid/issues/?jql=project%20%3D%20ANW'
```

Als Eingabe funktionieren Testausfuehrungs-IDs/-URLs, Testplan-IDs/-URLs und Jira-Filter-URLs. Bei Testplaenen werden die angehaengten Testausfuehrungen exportiert. Excel muss auf diesem Rechner installiert sein.

Nach einem neuen Clone die lokale Konfiguration und CSV-Vorlage anlegen:

```powershell
Copy-Item .\Xray-Export-Konfiguration.example.json .\Xray-Export-Konfiguration.json
Copy-Item .\SharePoint-Dateiliste.example.csv .\SharePoint-Dateiliste.csv
```

Die Jira-Adresse und Proxy-Einstellungen stehen in `Xray-Export-Konfiguration.json`; die CSV wird mit den freigegebenen SharePoint-Dateinamen und Links befuellt. Das Token kann beim ersten Start verdeckt eingegeben oder einmalig benutzergeschuetzt gespeichert werden:

```powershell
.\Export-XrayTestausfuehrungen.ps1 -SaveToken
```

Die SharePoint-Zuordnung verwendet `SharePoint-Dateiliste.csv`; `SharePoint-Dateiliste.example.csv` enthaelt nur die Spaltenkoepfe. Bestaetigte projektspezifische Statuslisten koennen mit `-StatusConfigPath .\Xray-Export-Status.json` angegeben werden. Excel-Dateien entstehen standardmaessig unter `Exporte\`.

Die echte Jira-Konfiguration, Statusliste und SharePoint-Dateiliste sind lokal und werden von Git ignoriert. Vor dem Commit pruefen, dass weder interne SharePoint-Links noch Jira-Serverdaten enthalten sind.

## SharePoint-Dateiliste aktualisieren

1. In Edge oder Chrome anmelden und den zu durchsuchenden SharePoint-Ordner
   oeffnen.
2. Mit `F12` die Entwicklertools oeffnen, unter **Sources > Snippets** ein
	neues Snippet anlegen und den Inhalt von
	`SharePoint-Dateiliste-aktualisieren.js` einfuegen.
3. Das Snippet speichern und mit `Ctrl+Enter` ausfuehren. Es durchsucht den
	in der SharePoint-URL ausgewaehlten Ordner samt Unterordnern und
	OData-Folgeseiten. Falls die URL keinen Ordner enthaelt, nach Aufforderung
	dessen server-relativen Pfad eingeben.
4. Den eingeblendeten Button **SharePoint-CSV speichern** anklicken. Beim
	ersten Mal im Dateidialog diesen Excel-Skriptordner auswaehlen und als
	`SharePoint-Dateiliste.csv` speichern beziehungsweise die vorhandene Datei
	auswaehlen. Es wird direkt in diese Datei geschrieben, nicht in den
	Browser-Downloads-Ordner. Der Browser merkt sich den zuletzt gewaehlten
	Speicherort fuer weitere Aktualisierungen.

Die Aktualisierung muss auf der angemeldeten SharePoint-Site ausgefuehrt
werden, da die REST-Aufrufe die Browser-Sitzung verwenden. Das Snippet leitet
Site und Host aus dem aktuellen SharePoint-Kontext ab und enthaelt keine
fest codierten Tenant- oder Projektpfade.
