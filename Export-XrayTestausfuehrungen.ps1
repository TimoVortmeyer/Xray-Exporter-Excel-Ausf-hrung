#requires -Version 5.1
[CmdletBinding()]
param(
    [string[]]$IssueKeys,
    [string]$FilterUrl,
    [string]$ConfigPath = '',
    [string]$OutputPath,
    [string]$SharePointListPath = '',
    [string]$StatusConfigPath,
    [switch]$SaveToken,
    [switch]$LibraryOnly
)

$ErrorActionPreference = 'Stop'
$RunLibraryOnly = $LibraryOnly
$RunIssueKeys = @($IssueKeys | Where-Object { $null -ne $_ })
$RunFilterUrl = $FilterUrl
$RunConfigPath = $ConfigPath
$RunOutputPath = $OutputPath
$RunStatusConfigPath = $StatusConfigPath
$RunSaveToken = $SaveToken
. (Join-Path $PSScriptRoot 'Xray-Export-ExcelSupport.ps1')
$IssueKeys = $RunIssueKeys
$FilterUrl = $RunFilterUrl
$ConfigPath = $RunConfigPath
$OutputPath = $RunOutputPath
$StatusConfigPath = $RunStatusConfigPath
$SaveToken = $RunSaveToken
$LibraryOnly = $RunLibraryOnly

function Get-TestRunIssueKey {
    param([string]$Value)
    $Trimmed = $Value.Trim()
    if ($Trimmed -match '^[A-Z][A-Z0-9_]*-[1-9][0-9]*$') { return $Trimmed.ToUpperInvariant() }
    $Uri = $null
    if ([uri]::TryCreate($Trimmed, [UriKind]::Absolute, [ref]$Uri)) {
        $Base = [uri]$script:ExportBaseUrl
        if ($Uri.Scheme -ne 'https' -or $Uri.Authority -ne $Base.Authority -or $Uri.UserInfo) {
            throw 'Die Vorgangs-URL muss zum konfigurierten HTTPS-Jira-Server gehoeren.'
        }
        if ($Uri.AbsolutePath -match '(?i)/browse/(?<key>[A-Z][A-Z0-9_]*-[1-9][0-9]*)/?$') {
            return $Matches.key.ToUpperInvariant()
        }
    }
    throw "Ungueltige Vorgangs-ID oder Jira-Vorgangs-URL: $Value"
}

function Get-TestRunIssue {
    param([string]$Key)
    $Issue = Invoke-ExportRequest "/rest/api/2/issue/${Key}?fields=summary,issuetype,project"
    if ([string]$Issue.key -cne $Key -or -not $Issue.fields.issuetype.name -or -not $Issue.fields.project.key) {
        throw "Jira hat fuer $Key keine gueltigen Vorgangsdaten geliefert."
    }
    return $Issue
}

function Get-TestRunReferences {
    param([object]$Response, [string]$Context)
    $Items = @($Response | Where-Object { $null -ne $_ })
    $Keys = [System.Collections.Generic.List[string]]::new()
    foreach ($Item in $Items) {
        $Key = if ($Item -is [string]) { [string]$Item }
        elseif ($Item.key) { [string]$Item.key }
        elseif ($Item.issueKey) { [string]$Item.issueKey }
        elseif ($Item.testExecKey) { [string]$Item.testExecKey }
        else { '' }
        if ($Key -notmatch '^[A-Z][A-Z0-9_]*-[1-9][0-9]*$') { throw "Ungueltiger Vorgang in der Xray-Antwort fuer $Context." }
        if ($Keys -notcontains $Key) { $Keys.Add($Key) }
    }
    return ,@($Keys)
}

function Get-TestRunExecutionTests {
    param([string]$ExecutionKey)
    $PageSize = 200
    $PageNumber = 1
    $TestKeys = [System.Collections.Generic.List[string]]::new()
    do {
        $Path = "$script:ExportApiPath/testexec/$ExecutionKey/test?limit=$PageSize&page=$PageNumber"
        $Page = Get-TestRunReferences (Invoke-ExportRequest $Path) "$ExecutionKey (Seite $PageNumber)"
        foreach ($TestKey in $Page) { $TestKeys.Add($TestKey) }
        $PageNumber++
    } while ($Page.Count -eq $PageSize)
    return ,$TestKeys.ToArray()
}

function Get-TestRunExecutions {
    param([string[]]$InputKeys)
    $Executions = [System.Collections.Generic.List[object]]::new()
    $Seen = @{}
    foreach ($Key in $InputKeys) {
        $Issue = Get-TestRunIssue $Key
        $Type = [string]$Issue.fields.issuetype.name
        if ($Type -ieq 'Test Plan') {
            $Related = Get-TestRunReferences (Invoke-ExportRequest "$script:ExportApiPath/testplan/$Key/testexecs") $Key
            foreach ($ExecutionKey in $Related) {
                $Execution = Get-TestRunIssue $ExecutionKey
                if ([string]$Execution.fields.issuetype.name -ine 'Test Execution') {
                    throw "$ExecutionKey ist keine Xray-Testausfuehrung (Issue-Typ: $($Execution.fields.issuetype.name))."
                }
                if (-not $Seen.ContainsKey($ExecutionKey)) {
                    $Seen[$ExecutionKey] = $true
                    $Executions.Add($Execution)
                }
            }
        }
        elseif ($Type -ieq 'Test Execution') {
            if (-not $Seen.ContainsKey($Key)) {
                $Seen[$Key] = $true
                $Executions.Add($Issue)
            }
        }
        else { throw "$Key ist weder Testplan noch Testausfuehrung (Issue-Typ: $Type)." }
    }
    if (-not $Executions.Count) { throw 'Es wurden keine Testausfuehrungen gefunden.' }
    return ,$Executions.ToArray()
}

function Get-TestRunSharePointLink {
    param([object]$Comments, [object[]]$Files, [string]$TestKey)
    $Candidates = [System.Collections.Generic.List[string]]::new()
    foreach ($Comment in @($Comments.comments)) {
        $Match = [regex]::Match([string]$Comment.body, '(?im)Originalworddokument\s*:\s*(?<name>[^\r\n]+)')
        if ($Match.Success) { $Candidates.Add($Match.Groups['name'].Value.Trim().Trim('"', "'", '`').Replace('\_', '_')) }
    }
    if (-not $Candidates.Count) { return '' }
    $MatchedFiles = [System.Collections.Generic.List[object]]::new()
    foreach ($Candidate in $Candidates) {
        $Name = [IO.Path]::GetFileName($Candidate)
        $Stem = $Name -replace '(?i)\.(docx?|docm)$', ''
        $Found = @($Files | Where-Object {
            $FileName = [IO.Path]::GetFileName([string]$_.Dateiname)
            $FileName -ieq $Name -or [IO.Path]::GetFileNameWithoutExtension($FileName) -ieq $Stem
        })
        foreach ($File in $Found) {
            if (-not @($MatchedFiles | Where-Object { $_.Link -ceq $File.Link }).Count) { $MatchedFiles.Add($File) }
        }
    }
    $Unique = @($MatchedFiles | Group-Object Link | ForEach-Object { $_.Group[0] })
    if ($Unique.Count -gt 1) {
        $Unique = @($Unique | Where-Object {
            $CandidateUri = $null
            if (-not [uri]::TryCreate([string]$_.Link, [UriKind]::Absolute, [ref]$CandidateUri)) { return $true }
            $Segments = $CandidateUri.AbsolutePath -split '/'
            for ($Index = 0; $Index -lt $Segments.Count - 1; $Index++) {
                if ([uri]::UnescapeDataString($Segments[$Index]) -ieq 'MUW_OBSOLET') { return $false }
            }
            return $true
        })
    }
    if ($Unique.Count -gt 1) {
        Write-Warning "Mehrere SharePoint-Dateien passen zu den Kommentaren von $TestKey. Statt eines Links wird die Testfallbeschreibung verwendet."
        return ''
    }
    if (-not $Unique.Count) {
        Write-Warning "Kein SharePoint-Link fuer $TestKey gefunden (Kommentarname: $($Candidates -join ', '))."
        return ''
    }
    $Link = [string]$Unique[0].Link
    $LinkUri = $null
    if (-not [uri]::TryCreate($Link, [UriKind]::Absolute, [ref]$LinkUri) -or $LinkUri.Scheme -ne 'https') {
        throw "Ungueltiger SharePoint-Link fuer $TestKey."
    }
    return $Link
}

function Get-TestRunTone {
    param([string]$Status)
    switch -Regex ($Status.ToUpperInvariant() -replace '[^A-Z0-9]', '') {
        '^TODO$' { 'neutral'; break }
        '^EXECUTING$' { 'active'; break }
        '^PASS$' { 'success'; break }
        '^PASSRESTRICTED$' { 'restricted'; break }
        '^FAIL$' { 'failure'; break }
        '^(ABORTED|IRRELEVANT)$' { 'aborted'; break }
        default { 'neutral' }
    }
}

function ConvertTo-TestRunRgb {
    param([string]$Hex)
    Add-Type -AssemblyName System.Drawing
    $Color = [Drawing.ColorTranslator]::FromHtml($Hex)
    return ($Color.R + ($Color.G * 256) + ($Color.B * 65536))
}

function Add-TestRunValidation {
    param([object]$Range, [int]$MaximumLength)
    $Range.Validation.Add(6, 1, 8, [string]$MaximumLength)
    $Range.Validation.IgnoreBlank = $true
    $Range.Validation.ErrorTitle = 'Eingabe zu lang'
    $Range.Validation.ErrorMessage = "Maximal $MaximumLength Zeichen."
    $Range.Validation.ShowError = $true
}

function New-TestRunWorkbook {
    param([object]$Excel, [object]$Execution, [string[]]$Statuses, [object[]]$Files, [string]$Destination, [string[]]$Tests, [hashtable]$ProgressState)
    $Workbook = $Excel.Workbooks.Add()
    $Sheet = $Workbook.Worksheets.Item(1)
    $Sheet.Name = 'Testergebnis'
    $Headers = @(('Schl{0}ssel' -f [char]0x00fc), 'Testfallbezeichnung', 'Beschreibung / Link Word-Testdokumentation', 'Tester', 'Status', 'Anmerkung')
    for ($Column = 0; $Column -lt $Headers.Count; $Column++) { $Sheet.Cells.Item(1, $Column + 1).Value2 = $Headers[$Column] }
    $Sheet.Cells.Item(1, 7).Value2 = 'Statusindex'
    $Sheet.Columns.Item(7).Hidden = $true
    $Sheet.Columns.Item(1).ColumnWidth = 16
    $Sheet.Columns.Item(2).ColumnWidth = 48
    $Sheet.Columns.Item(3).ColumnWidth = 50
    $Sheet.Columns.Item(4).ColumnWidth = 22
    $Sheet.Columns.Item(5).ColumnWidth = 30
    $Sheet.Columns.Item(6).ColumnWidth = 52
    $Header = $Sheet.Range('A1:F1')
    $Header.Font.Bold = $true
    $Header.Interior.Color = ConvertTo-TestRunRgb '#0050A4'
    $Header.Font.Color = ConvertTo-TestRunRgb '#FFFFFF'
    $Header.WrapText = $true
    $Header.RowHeight = 32
    $Sheet.Application.ActiveWindow.SplitRow = 1
    $Sheet.Application.ActiveWindow.FreezePanes = $true
    if (-not $Tests.Count) { throw "$($Execution.key) enthaelt keine Testfaelle." }
    $TodoIndex = 0
    for ($StatusIndex = 0; $StatusIndex -lt $Statuses.Count; $StatusIndex++) {
        if ($Statuses[$StatusIndex] -ieq 'TODO') { $TodoIndex = $StatusIndex + 1; break }
    }
    $Row = 2
    if (-not $ProgressState.Stopwatch.IsRunning) { $ProgressState.Stopwatch.Start() }
    foreach ($TestKey in $Tests) {
        $Issue = Invoke-ExportRequest "/rest/api/2/issue/${TestKey}?fields=summary,issuetype,comment,description"
        if ([string]$Issue.fields.issuetype.name -ine 'Test') { throw "$TestKey ist kein Xray-Testfall (Issue-Typ: $($Issue.fields.issuetype.name))." }
        $Sheet.Cells.Item($Row, 1).Value2 = $TestKey
        $Sheet.Cells.Item($Row, 2).Value2 = [string]$Issue.fields.summary
        $Link = Get-TestRunSharePointLink $Issue.fields.comment $Files $TestKey
        if ($Link) {
            $ScreenTip = 'SharePoint-Testfalldokument {0}ffnen' -f [char]0x00f6
            $LinkText = 'Dokument {0}ffnen' -f [char]0x00f6
            $Sheet.Hyperlinks.Add($Sheet.Cells.Item($Row, 3), $Link, '', $ScreenTip, $LinkText) | Out-Null
        }
        else { $Sheet.Cells.Item($Row, 3).Value2 = [string]$Issue.fields.description }
        $Sheet.Cells.Item($Row, 4).Locked = $false
        $Sheet.Cells.Item($Row, 6).Locked = $false
        $Sheet.Cells.Item($Row, 7).Locked = $false
        $Sheet.Cells.Item($Row, 7).Value2 = [string]$TodoIndex
        $RowHeight = [Math]::Max(48, 17 * $Statuses.Count + 12)
        $Sheet.Rows.Item($Row).RowHeight = $RowHeight
        $Sheet.Cells.Item($Row, 1).NumberFormat = '@'
        $Sheet.Cells.Item($Row, 4).WrapText = $true
        $Sheet.Cells.Item($Row, 3).WrapText = $true
        $Sheet.Cells.Item($Row, 6).WrapText = $true
        $Sheet.Rows.Item($Row).AutoFit()
        if ($Sheet.Rows.Item($Row).RowHeight -lt $RowHeight) { $Sheet.Rows.Item($Row).RowHeight = $RowHeight }
        $Sheet.Range("A${Row}:F${Row}").VerticalAlignment = -4160
        $Sheet.Range("A${Row}:F${Row}").Interior.Color = ConvertTo-TestRunRgb '#F5F5F6'
        $Top = [double]$Sheet.Cells.Item($Row, 5).Top + 1
        $Left = [double]$Sheet.Cells.Item($Row, 5).Left + 2
        $Height = [Math]::Max(17, ($Sheet.Rows.Item($Row).RowHeight - 12) / $Statuses.Count)
        $Group = $Sheet.Shapes.AddFormControl(4, $Left, $Top, [double]$Sheet.Cells.Item($Row, 5).Width - 4, [double]$Sheet.Rows.Item($Row).RowHeight - 2)
        $Group.TextFrame.Characters().Text = ''
        $Group.Line.Visible = 0
        $Group.Fill.Visible = 0
        for ($StatusIndex = 0; $StatusIndex -lt $Statuses.Count; $StatusIndex++) {
            $Option = $Sheet.Shapes.AddFormControl(7, $Left + 2, $Top + 5 + ($StatusIndex * $Height), [double]$Sheet.Cells.Item($Row, 5).Width - 8, $Height)
            $Option.TextFrame.Characters().Text = $Statuses[$StatusIndex]
            $Option.Fill.Visible = 0
            $Option.ControlFormat.LinkedCell = $Sheet.Cells.Item($Row, 7).Address($false, $false)
            $Option.ControlFormat.Value = -4146
            $Tone = Get-TestRunTone $Statuses[$StatusIndex]
            $Fill = switch ($Tone) {
                'active' { '#FBFAF0' }
                'success' { '#F4F7F0' }
                'restricted' { '#F7FCE8' }
                'failure' { '#F9EFEE' }
                'aborted' { '#E8E8E8' }
                default { '#F5F5F6' }
            }
            $Condition = $Sheet.Range("A${Row}:F${Row}").FormatConditions.Add(2, 0, "=`$G`$$Row=$($StatusIndex + 1)")
            $Condition.Interior.Color = ConvertTo-TestRunRgb $Fill
        }
        $Sheet.Cells.Item($Row, 7).Value2 = [string]$TodoIndex
        $Sheet.Cells.Item($Row, 3).HorizontalAlignment = -4108
        Add-TestRunValidation $Sheet.Cells.Item($Row, 4) 100
        Add-TestRunValidation $Sheet.Cells.Item($Row, 6) 1000
        $Row++
        $ProgressState.Completed++
        $PercentComplete = [Math]::Min(99, [int][Math]::Floor(100 * $ProgressState.Completed / $ProgressState.Total))
        $AverageSeconds = $ProgressState.Stopwatch.Elapsed.TotalSeconds / $ProgressState.Completed
        $RemainingSeconds = $AverageSeconds * ($ProgressState.Total - $ProgressState.Completed)
        $ExpectedEndTime = (Get-Date).AddSeconds($RemainingSeconds).ToString('HH:mm:ss')
        $ProgressStatus = '{0}/{1} Testfaelle; voraussichtliches Ende: {2}' -f $ProgressState.Completed, $ProgressState.Total, $ExpectedEndTime
        Write-Progress -Activity 'Xray-Export' -Status $ProgressStatus -PercentComplete $PercentComplete
    }
    $Sheet.Range("A1:F$($Row - 1)").Borders.LineStyle = 1
    $Sheet.Protect('', $false, $true, $true, $true)
    $Workbook.SaveAs($Destination, 51)
    $Workbook.Close($true)
}

if ($LibraryOnly) { return }
if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $ConfigPath = Join-Path $PSScriptRoot 'Xray-Export-Konfiguration.json' }
if ([string]::IsNullOrWhiteSpace($SharePointListPath)) { $SharePointListPath = Join-Path $PSScriptRoot 'SharePoint-Dateiliste.csv' }
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { throw 'Export-Konfiguration fehlt.' }
$script:ExportConfig = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$script:ExportBaseUrl = ([string]$script:ExportConfig.JiraBaseUrl).TrimEnd('/')
if ([uri]$script:ExportBaseUrl -isnot [uri] -or ([uri]$script:ExportBaseUrl).Scheme -ne 'https') { throw 'JiraBaseUrl muss HTTPS verwenden.' }
if ($SaveToken) { Save-ExportToken $script:ExportBaseUrl; return }
if (-not (Test-Path -LiteralPath $SharePointListPath -PathType Leaf)) { throw 'SharePoint-Dateiliste.csv wurde nicht gefunden.' }
if ($StatusConfigPath -and -not (Test-Path -LiteralPath $StatusConfigPath -PathType Leaf)) { throw 'Die angegebene Statusdatei wurde nicht gefunden.' }
$script:ExportApiPath = ([string]$script:ExportConfig.Xray.ApiPath).TrimEnd('/')
if (-not $script:ExportApiPath) { $script:ExportApiPath = '/rest/raven/1.0/api' }
if (-not $IssueKeys -and -not $FilterUrl) { $IssueKeys = @((Read-Host 'Testausfuehrungs-/Testplan-ID oder Jira-Filter-URL')) }
if (-not $FilterUrl -and $IssueKeys.Count -eq 1 -and $IssueKeys[0].Trim() -match '^https?://') {
    $InputUri = [uri]$IssueKeys[0].Trim()
    if ($InputUri.Query -match '(?i)(?:^|[?&])(jql|filter)=') { $FilterUrl = $IssueKeys[0].Trim(); $IssueKeys = @() }
    else { $IssueKeys = @((Get-TestRunIssueKey $IssueKeys[0])) }
}
$InputKeys = @($IssueKeys | ForEach-Object { Get-TestRunIssueKey $_ } | Select-Object -Unique)
$FilterSelection = $null
if ($FilterUrl) { $FilterSelection = Get-ExportFilterSelection $FilterUrl.Trim() }
if (-not $InputKeys.Count -and -not $FilterSelection) { throw 'Bitte mindestens einen Testplan, eine Testausfuehrung oder einen Jira-Filter angeben.' }
$Secret = Get-ExportToken $script:ExportBaseUrl
$Pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secret)
$Excel = $null
$Workbooks = $null
try {
    $script:ExportHeaders = @{ Authorization = 'Bearer ' + [Runtime.InteropServices.Marshal]::PtrToStringBSTR($Pointer) }
    if ($FilterSelection) { $InputKeys += @(Get-ExportFilterKeys $FilterSelection) }
    $InputKeys = @($InputKeys | Select-Object -Unique)
    $Executions = Get-TestRunExecutions $InputKeys
    $ExecutionTests = @{}
    $TotalTests = 0
    for ($ExecutionIndex = 0; $ExecutionIndex -lt $Executions.Count; $ExecutionIndex++) {
        $Execution = $Executions[$ExecutionIndex]
        $ExecutionKey = [string]$Execution.key
        $PercentComplete = [int][Math]::Floor(100 * $ExecutionIndex / $Executions.Count)
        $ProgressStatus = 'Ermittle Testfaelle fuer Ausfuehrung {0} von {1}' -f ($ExecutionIndex + 1), $Executions.Count
        Write-Progress -Activity 'Xray-Export' -Status $ProgressStatus -PercentComplete $PercentComplete
        $Tests = Get-TestRunExecutionTests $ExecutionKey
        if (-not $Tests.Count) { throw "$ExecutionKey enthaelt keine Testfaelle." }
        $ExecutionTests[$ExecutionKey] = @($Tests)
        $TotalTests += $Tests.Count
    }
    $ProgressState = @{
        Total = $TotalTests
        Completed = 0
        Stopwatch = [Diagnostics.Stopwatch]::new()
    }
    $ConfiguredStatuses = $null
    if ($StatusConfigPath) { $ConfiguredStatuses = Get-Content -LiteralPath $StatusConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json }
    $LiveStatuses = $null
    $Files = @(Import-Csv -LiteralPath $SharePointListPath -Delimiter ';' -Encoding UTF8)
    if (-not $OutputPath) { $OutputPath = Join-Path $PSScriptRoot ('Exporte\Testausfuehrungen-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
    [IO.Directory]::CreateDirectory($OutputPath) | Out-Null
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbooks = $Excel.Workbooks
    foreach ($Execution in $Executions) {
        $ExecutionKey = [string]$Execution.key
        $ProjectKey = [string]$Execution.fields.project.key
        $ProjectStatusConfig = if ($ConfiguredStatuses) { $ConfiguredStatuses.PSObject.Properties[$ProjectKey] } else { $null }
        if ($ProjectStatusConfig) {
            if ($ProjectStatusConfig.Value.confirmed -ne $true) { throw "Statusliste fuer $ProjectKey ist nicht bestaetigt." }
            $Statuses = Get-ExportStatuses $ProjectStatusConfig.Value.test 'Projekt-Testlauf'
        }
        else {
            if ($null -eq $LiveStatuses) {
                $StatusResponse = Invoke-ExportRequest "$script:ExportApiPath/settings/teststatuses"
                $LiveStatuses = Get-ExportStatuses $StatusResponse 'Xray-Testlauf'
            }
            $Statuses = @($LiveStatuses)
        }
        if ($Statuses -notcontains 'TODO') { throw "Statusliste fuer $ProjectKey enthaelt keinen TODO-Status." }
        $Name = ConvertTo-ExportFileName ([string]$Execution.fields.summary) $ExecutionKey
        $Destination = Join-Path $OutputPath ($ExecutionKey + ' - ' + $Name + '.xlsx')
        Write-Host "Erstelle $Destination ..."
        try { New-TestRunWorkbook $Excel $Execution $Statuses $Files $Destination $ExecutionTests[$ExecutionKey] $ProgressState }
        catch [Runtime.InteropServices.COMException] {
            throw "Excel hat die Automatisierung abgelehnt. Bitte Excel-Dialoge schliessen und den Export erneut starten. Details: $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity 'Xray-Export' -Status 'Export abgeschlossen' -PercentComplete 100
    Write-Host "Fertig: $($Executions.Count) Excel-Datei(en) in $OutputPath"
}
finally {
    Write-Progress -Activity 'Xray-Export' -Completed
    try { if ($Excel) { $Excel.Quit() } } catch {}
    try { if ($Workbooks) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Workbooks) } } catch {}
    try { if ($Excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) } } catch {}
    if ($Pointer) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Pointer) }
    if ($Secret) { $Secret.Dispose() }
    $script:ExportHeaders = $null
}