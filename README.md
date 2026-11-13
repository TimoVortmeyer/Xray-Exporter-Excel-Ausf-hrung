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
