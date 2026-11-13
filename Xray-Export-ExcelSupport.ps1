function Get-ExportTokenPath {
    param([string]$BaseUrl)
    $Base = [uri]$BaseUrl
    if ($Base.Scheme -ne 'https' -or -not $Base.Host -or $Base.UserInfo) { throw 'Fuer die Tokenablage ist eine gueltige HTTPS-Jira-Adresse ohne Zugangsdaten erforderlich.' }
    if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { throw 'LOCALAPPDATA ist fuer die Tokenablage nicht verfuegbar.' }
    $Origin = $Base.GetLeftPart([UriPartial]::Authority).ToLowerInvariant()
    $Hash = [Security.Cryptography.SHA256]::Create()
    try { $Name = [BitConverter]::ToString($Hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($Origin))).Replace('-', '').ToLowerInvariant() }
    finally { $Hash.Dispose() }
    return Join-Path (Join-Path $env:LOCALAPPDATA 'Xray-Testexport') ($Name + '.token')
}

function Get-ExportToken {
    param([string]$BaseUrl)
    $Path = Get-ExportTokenPath $BaseUrl
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return Read-Host 'Jira Personal Access Token (nur fuer den Export)' -AsSecureString
    }
    try { return ConvertTo-SecureString ([IO.File]::ReadAllText($Path)) }
    catch { throw 'Das gespeicherte Jira-Token kann fuer dieses Windows-Konto nicht gelesen werden.' }
}

function Save-ExportToken {
    param([string]$BaseUrl)
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Die benutzergeschuetzte Tokenablage ist nur unter Windows verfuegbar.' }
    $Path = Get-ExportTokenPath $BaseUrl
    if (Test-Path -LiteralPath $Path) { throw 'Fuer diesen Jira-Server ist bereits ein Token gespeichert. Zum Ersetzen zuerst die vorhandene Tokendatei entfernen.' }
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path)) | Out-Null
    $Secret = Read-Host 'Jira Personal Access Token (wird benutzergeschuetzt gespeichert)' -AsSecureString
    try {
        if ($Secret.Length -eq 0) { throw 'Ein leeres Token kann nicht gespeichert werden.' }
        $Protected = ConvertFrom-SecureString $Secret
        $Stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $Writer = [IO.StreamWriter]::new($Stream, [Text.Encoding]::UTF8)
            try { $Writer.Write($Protected) } finally { $Writer.Dispose() }
        }
        finally { $Stream.Dispose() }
        Write-Host 'Jira-Token fuer diesen Windows-Benutzer und Server gespeichert.'
    }
    finally { $Secret.Dispose() }
}

function Invoke-ExportRequest {
    param([string]$Path, [switch]$Binary)
    $Target = [uri]::new([uri]($script:ExportBaseUrl + '/'), $Path)
    $Base = [uri]$script:ExportBaseUrl
    if ($Target.Scheme -ne 'https' -or $Target.Authority -ne $Base.Authority -or $Target.UserInfo) {
        throw 'Download ausserhalb des HTTPS-Jira-Servers verweigert.'
    }
    $Parameters = @{
        Uri = $Target.AbsoluteUri
        Headers = $script:ExportHeaders
        Method = 'Get'
        MaximumRedirection = 0
        ErrorAction = 'Stop'
    }
    $Proxy = $script:ExportConfig.Proxy
    if ($Proxy.Mode -eq 'Explicit') {
        $Parameters.Proxy = [string]$Proxy.Url
        $Parameters.ProxyUseDefaultCredentials = [bool]$Proxy.UseDefaultCredentials
    }
    elseif (-not $Proxy.Mode -or $Proxy.Mode -eq 'System') {
        $SystemProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        if (-not $SystemProxy.IsBypassed($Target)) {
            $Parameters.Proxy = $SystemProxy.GetProxy($Target).AbsoluteUri
            $Parameters.ProxyUseDefaultCredentials = ($Proxy.UseDefaultCredentials -ne $false)
        }
        elseif ($PSVersionTable.PSVersion.Major -ge 7) { $Parameters.NoProxy = $true }
    }
    else { throw 'Unbekannter Proxy-Modus.' }
    try {
        if ($Binary) {
            $Response = Invoke-WebRequest @Parameters -UseBasicParsing
            $Mime = ([string]$Response.Headers['Content-Type']).Split(';')[0].Trim()
            if ($Mime -notin @('image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/bmp')) {
                throw 'Nicht unterstuetztes Bildformat oder keine Bildantwort.'
            }
            if ($Response.Content -isnot [byte[]]) { throw 'Keine binaere Bildantwort.' }
            return "data:$Mime;base64,$([Convert]::ToBase64String($Response.Content))"
        }
        $Response = Invoke-RestMethod @Parameters
        Write-Output -NoEnumerate $Response
    }
    catch {
        $StatusCode = 'unbekannt'
        $ResponseBody = ''
        $HttpResponse = $_.Exception.Response
        if ($HttpResponse) {
            try {
                $StatusCode = [string][int]$HttpResponse.StatusCode
                if ($HttpResponse -is [System.Net.HttpWebResponse]) {
                    $ResponseStream = $HttpResponse.GetResponseStream()
                    if ($ResponseStream) {
                        $Reader = [System.IO.StreamReader]::new($ResponseStream)
                        try { $ResponseBody = $Reader.ReadToEnd() } finally { $Reader.Dispose() }
                    }
                }
                elseif ($HttpResponse.Content) {
                    $ResponseBody = $HttpResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                }
            }
            catch {}
        }
        if (-not $ResponseBody) { $ResponseBody = [string]$_.ErrorDetails.Message }
        $ResponseBody = ($ResponseBody -replace '\s+', ' ').Trim()
        if ($ResponseBody.Length -gt 500) { $ResponseBody = $ResponseBody.Substring(0, 500) + '...' }
        $Details = if ($ResponseBody) { " Antwort: $ResponseBody." } else { '' }
        throw "Jira-Abfrage fehlgeschlagen (HTTP $StatusCode): $($Target.AbsolutePath).$Details Keine Datei erstellt."
    }
}

function Get-ExportStatuses {
    param([object]$Items, [string]$Kind)
    $Names = @($Items | ForEach-Object {
        if ($_ -is [string]) { $_ } elseif ($_.name) { [string]$_.name }
        else { throw "Unbekanntes Format der $Kind-Statusliste." }
    })
    $Names = @($Names | Where-Object { $_ -ine 'SKIPPED' })
    if (-not $Names.Count -or @($Names | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count) {
        throw "Leere $Kind-Statusliste."
    }
    if (@($Names | Select-Object -Unique).Count -ne $Names.Count) { throw "Doppelte $Kind-Statuswerte." }
    return ,$Names
}

function ConvertTo-ExportFileName {
    param([string]$Summary, [string]$Fallback)
    $Name = [regex]::Replace($Summary, '[<>:"/\\|?*\x00-\x1f]', '_').Trim().TrimEnd([char[]]' .')
    if (-not $Name) { $Name = $Fallback }
    if ($Name.Length -gt 120) { $Name = $Name.Substring(0, 120).TrimEnd([char[]]' .') }
    if ([char]::IsHighSurrogate($Name[$Name.Length - 1])) { $Name = $Name.Substring(0, $Name.Length - 1) }
    if ($Name -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { $Name = '_' + $Name }
    return $Name
}

function Get-ExportFilterSelection {
    param([string]$Url)
    $FilterUri = $null
    $BaseUri = [uri]$script:ExportBaseUrl
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$FilterUri) -or
        $FilterUri.Scheme -ne 'https' -or $FilterUri.Authority -ne $BaseUri.Authority -or $FilterUri.UserInfo) {
        throw 'Die Filter-URL muss zum konfigurierten HTTPS-Jira-Server gehoeren.'
    }
    Add-Type -AssemblyName System.Web
    $Query = [System.Web.HttpUtility]::ParseQueryString($FilterUri.Query)
    if (@($Query.GetValues('jql')).Count -gt 1 -or @($Query.GetValues('filter')).Count -gt 1) {
        throw 'Die Filter-URL enthaelt mehrdeutige jql- oder filter-Parameter.'
    }
    $Jql = [string]$Query['jql']
    if (-not [string]::IsNullOrWhiteSpace($Jql)) {
        return @{ Jql = $Jql; FilterId = $null }
    }
    $FilterId = [string]$Query['filter']
    if ($FilterId -match '^[1-9][0-9]*$') { return @{ Jql = $null; FilterId = $FilterId } }
    throw 'Die Filter-URL muss eine JQL-Abfrage (?jql=...) oder eine gespeicherte Filter-ID (?filter=12345) enthalten.'
}

function Get-ExportFilterKeys {
    param([hashtable]$Selection)
    $Jql = [string]$Selection.Jql
    if ($Selection.FilterId) {
        $Filter = Invoke-ExportRequest ('/rest/api/2/filter/' + $Selection.FilterId)
        $Jql = [string]$Filter.jql
    }
    if ([string]::IsNullOrWhiteSpace($Jql)) { throw 'Der Jira-Filter liefert keine JQL-Abfrage.' }
    $EncodedJql = [uri]::EscapeDataString($Jql)
    $Keys = [System.Collections.Generic.List[string]]::new()
    $Seen = @{}
    $StartAt = 0
    $Total = $null
    for ($Page = 0; $Page -lt 1000; $Page++) {
        $Response = Invoke-ExportRequest "/rest/api/2/search?jql=$EncodedJql&fields=key&startAt=$StartAt&maxResults=100"
        if ($null -eq $Response.PSObject.Properties['issues'] -or
            $null -eq $Response.PSObject.Properties['total'] -or
            [string]$Response.total -notmatch '^\d+$') {
            throw 'Die Jira-Suche liefert keine gueltigen issues/total-Paginierungsdaten.'
        }
        if ($null -eq $Total) { $Total = [long]$Response.total }
        elseif ($Total -ne [long]$Response.total) { throw 'Die Trefferzahl hat sich waehrend des Exports geaendert. Bitte erneut starten.' }
        if ($null -ne $Response.PSObject.Properties['startAt'] -and [long]$Response.startAt -ne $StartAt) {
            throw 'Die Jira-Suche liefert eine unerwartete Ergebnisseite.'
        }
        $Items = @($Response.issues | Where-Object { $null -ne $_ })
        foreach ($Item in $Items) {
            $Key = [string]$Item.key
            if ($Key -notmatch '^[A-Z][A-Z0-9_]*-[1-9][0-9]*$') { throw 'Die Jira-Suche liefert einen ungueltigen Vorgangsschluessel.' }
            if ($Seen.ContainsKey($Key)) { throw 'Die Jira-Suche liefert wiederholte Treffer. Pagination nicht verlaesslich.' }
            $Seen[$Key] = $true
            $Keys.Add($Key)
        }
        if ($Keys.Count -gt $Total) { throw 'Die Jira-Suche liefert mehr Treffer als angekuendigt.' }
        if ($Keys.Count -eq $Total) {
            if (-not $Keys.Count) { throw 'Der Jira-Filter enthaelt keine sichtbaren Vorgaenge.' }
            Write-Host "Jira-Filter: $($Keys.Count) Vorgaenge gefunden."
            foreach ($Key in $Keys) { $Key }
            return
        }
        if (-not $Items.Count) { throw 'Die Jira-Suche endet vor Erreichen aller Treffer.' }
        $StartAt += $Items.Count
    }
    throw 'Maximale Seitenanzahl der Jira-Suche erreicht. Bitte den Filter einschraenken.'
}
