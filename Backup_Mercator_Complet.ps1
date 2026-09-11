<#
.SYNOPSIS
    Sauvegarde complete de Mercator via API REST.
.DESCRIPTION
    Exporte les ressources Mercator en JSON, CSV et HTML, puis telecharge les
    rapports Word dans le dossier WORD. Cree un index HTML, une synthese CSV,
    un controle JSON, une archive ZIP et applique la retention.
.NOTES
    Compatible Windows PowerShell 5.1 et PowerShell 7.
#>
#Requires -Version 5.1
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# CONFIGURATION
$ApiBase = '#Lien_MERCATOR'
$CredPath = '#emplacement_fichier_cred\mercator_cred.xml'
$BackupRoot = '#emplacement_destination_de_la_sauvegarde'
$TimeoutSec = 300
$ApiDelayMilliseconds = 500
$ReportDelaySeconds = 5
$MaxRetryAttempts = 5
$InitialRetryDelaySeconds = 10
$MaximumRetryDelaySeconds = 120
$JsonDepth = 100
$RetentionDays = 30
$CreateZip = $true
$RemoveFolderAfterZip = $false

# L endpoint permissions a ete retire. L endpoint backups est inclus.
$Endpoints = @(
    'activities','actors','admin-users','annuaires','application-blocks',
    'application-flows','application-modules','application-services','applications',
    'backups','bays','buildings','certificates','clusters','containers',
    'data-processings','databases','documents','domains','entities',
    'external-connected-entities','forest-ads','gateways','information','lans',
    'logical-flows','logical-servers','macro-processuses','mans','network-switches',
    'networks','operations','peripherals','phones','physical-links','physical-routers',
    'physical-security-devices','physical-servers','physical-switches','processes',
    'queries','relations','routers','security-controls','security-devices','sites',
    'storage-devices','subnetworks','tasks','vlans','wans','wifi-terminals',
    'workstations','zone-admins',
)

$WordReports = @(
    [pscustomobject]@{ Name='cartography'; Endpoint='report/cartography'; FileName='cartography.docx' },
    [pscustomobject]@{ Name='entities'; Endpoint='report/entities'; FileName='entities.docx' },
    [pscustomobject]@{ Name='applicationsByBlocks'; Endpoint='report/applicationsByBlocks'; FileName='applicationsByBlocks.docx' },
    [pscustomobject]@{ Name='directory'; Endpoint='report/directory'; FileName='directory.docx' },
    [pscustomobject]@{ Name='logicalServers'; Endpoint='report/logicalServers'; FileName='logicalServers.docx' },
    [pscustomobject]@{ Name='securityNeeds'; Endpoint='report/securityNeeds'; FileName='securityNeeds.docx' },
    [pscustomobject]@{ Name='logicalServerConfigs'; Endpoint='report/logicalServerConfigs'; FileName='logicalServerConfigs.docx' },
    [pscustomobject]@{ Name='externalAccess'; Endpoint='report/externalAccess'; FileName='externalAccess.docx' },
    [pscustomobject]@{ Name='physicalInventory'; Endpoint='report/physicalInventory'; FileName='physicalInventory.docx' },
    [pscustomobject]@{ Name='vlans'; Endpoint='report/vlans'; FileName='vlans.docx' },
    [pscustomobject]@{ Name='workstations'; Endpoint='report/workstations'; FileName='workstations.docx' },
    [pscustomobject]@{ Name='cve'; Endpoint='report/cve'; FileName='cve.docx' },
    [pscustomobject]@{ Name='activityList'; Endpoint='report/activityList'; FileName='activityList.docx' },
    [pscustomobject]@{ Name='activityReport'; Endpoint='report/activityReport'; FileName='activityReport.docx' },
    [pscustomobject]@{ Name='impacts'; Endpoint='report/impacts'; FileName='impacts.docx' },
    [pscustomobject]@{ Name='rto'; Endpoint='report/rto'; FileName='rto.docx' }
)

$StartTime = Get-Date
$DateStamp = $StartTime.ToString('yyyy-MM-dd_HH-mm-ss')
$ExportPath = Join-Path $BackupRoot $DateStamp
$JsonPath = Join-Path $ExportPath 'JSON'
$CsvPath = Join-Path $ExportPath 'CSV'
$HtmlPath = Join-Path $ExportPath 'HTML'
$WordPath = Join-Path $ExportPath 'WORD'
$LogFile = Join-Path $ExportPath 'BackupMercator.log'
$IndexFile = Join-Path $ExportPath 'index.html'
$SummaryCsvFile = Join-Path $ExportPath 'synthese.csv'
$ZipFile = Join-Path $BackupRoot ("Mercator_{0}.zip" -f $DateStamp)
foreach ($Directory in @($BackupRoot,$ExportPath,$JsonPath,$CsvPath,$HtmlPath,$WordPath)) {
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        $null = New-Item -ItemType Directory -Path $Directory -Force
    }
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][scriptblock]$Operation,
        [Parameter(Mandatory=$true)][string]$OperationName,
        [int]$MaxAttempts = $MaxRetryAttempts,
        [int]$InitialDelaySeconds = $InitialRetryDelaySeconds,
        [int]$MaximumDelaySeconds = $MaximumRetryDelaySeconds
    )
    for ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
        try { return & $Operation }
        catch {
            $StatusCode = Get-HttpStatusCode -ErrorRecord $_
            if ($StatusCode -ne 429 -or $Attempt -ge $MaxAttempts) { throw }
            $RetryAfterSeconds = $null
            try {
                $Response = $_.Exception.Response
                if ($null -ne $Response -and $null -ne $Response.Headers) {
                    $RetryAfterValue = [string]$Response.Headers['Retry-After']
                    $ParsedSeconds = 0
                    if ([int]::TryParse($RetryAfterValue,[ref]$ParsedSeconds)) { $RetryAfterSeconds = $ParsedSeconds }
                    elseif (-not [string]::IsNullOrWhiteSpace($RetryAfterValue)) {
                        $RetryDate = [datetime]::MinValue
                        if ([datetime]::TryParse($RetryAfterValue,[ref]$RetryDate)) {
                            $RetryAfterSeconds = [math]::Ceiling(($RetryDate.ToUniversalTime()-(Get-Date).ToUniversalTime()).TotalSeconds)
                        }
                    }
                }
            } catch { $RetryAfterSeconds = $null }
            if ($null -eq $RetryAfterSeconds -or $RetryAfterSeconds -lt 1) {
                $Delay = $InitialDelaySeconds * [math]::Pow(2,$Attempt-1)
                $RetryAfterSeconds = [int][math]::Min($MaximumDelaySeconds,$Delay+(Get-Random -Minimum 0 -Maximum 4))
            } else { $RetryAfterSeconds = [int][math]::Min($MaximumDelaySeconds,$RetryAfterSeconds) }
            Write-Log ("HTTP 429 pour {0}. Tentative {1}/{2}. Nouvelle tentative dans {3} seconde(s)." -f $OperationName,$Attempt,$MaxAttempts,$RetryAfterSeconds) 'WARNING'
            Start-Sleep -Seconds $RetryAfterSeconds
        }
    }
}

function Write-Log {
    param([Parameter(Mandatory=$true)][string]$Message,
          [ValidateSet('INFO','WARNING','ERROR')][string]$Level='INFO')
    $Line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Host $Line
    Add-Content -LiteralPath $LogFile -Value $Line -Encoding UTF8
}

function Get-HttpStatusCode {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    try { if ($null -ne $ErrorRecord.Exception.Response) { return [int]$ErrorRecord.Exception.Response.StatusCode } } catch {}
    return $null
}

function Get-HttpErrorBody {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    if ($null -ne $ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.ErrorDetails.Message)) {
        return [string]$ErrorRecord.ErrorDetails.Message
    }
    try {
        $Response = $ErrorRecord.Exception.Response
        if ($null -eq $Response) { return '' }
        if ($Response.PSObject.Methods.Name -contains 'GetResponseStream') {
            $Stream = $Response.GetResponseStream()
            $Reader = New-Object System.IO.StreamReader($Stream)
            try { return $Reader.ReadToEnd() } finally { $Reader.Dispose(); $Stream.Dispose() }
        }
    } catch {}
    return ''
}

function Get-EndpointItems {
    param([AllowNull()]$ApiResult)
    if ($null -eq $ApiResult) { return @() }
    if ($ApiResult.PSObject.Properties.Name -contains 'data' -and $null -ne $ApiResult.data) { return @($ApiResult.data) }
    return @($ApiResult)
}

function ConvertTo-FlatObject {
    param([AllowNull()]$InputObject)
    if ($null -eq $InputObject) { return [pscustomobject]@{} }
    $Flat = [ordered]@{}
    foreach ($Property in $InputObject.PSObject.Properties) {
        $Value = $Property.Value
        if ($null -eq $Value -or $Value -is [string] -or $Value -is [ValueType]) {
            $Flat[$Property.Name] = $Value
        } else {
            $Flat[$Property.Name] = ($Value | ConvertTo-Json -Depth $JsonDepth -Compress)
        }
    }
    return [pscustomobject]$Flat
}

function New-OfflineHtmlPage {
    param([string]$Endpoint,[array]$Items,[string]$OutputFile,[datetime]$ExportDate)
    $Rows = New-Object System.Text.StringBuilder
    if ($Items.Count -eq 0) {
        $Columns = @('Information')
        $null = $Rows.AppendLine('<tr><td>Aucun element retourne par API.</td></tr>')
    } else {
        $FlatItems = @(foreach ($Item in $Items) { ConvertTo-FlatObject $Item })
        $Columns = @($FlatItems | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
        foreach ($Item in $FlatItems) {
            $null = $Rows.AppendLine('<tr>')
            foreach ($Column in $Columns) {
                $P = $Item.PSObject.Properties[$Column]
                $V = if ($null -eq $P -or $null -eq $P.Value) { '' } else { [System.Net.WebUtility]::HtmlEncode([string]$P.Value) }
                $null = $Rows.AppendLine("<td>$V</td>")
            }
            $null = $Rows.AppendLine('</tr>')
        }
    }
    $Headers = ($Columns | ForEach-Object { '<th>{0}</th>' -f [System.Net.WebUtility]::HtmlEncode([string]$_) }) -join ''
    $Title = [System.Net.WebUtility]::HtmlEncode($Endpoint)
    $Html = @"
<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mercator - $Title</title><style>body{font-family:Segoe UI,Arial;margin:0;background:#f4f6f8;color:#202020}header{background:#003b5c;color:white;padding:22px 30px}main{padding:22px 30px}input{padding:10px;width:380px;max-width:90%;margin-bottom:15px}table{border-collapse:collapse;width:100%;background:white}th{background:#e7eef3;text-align:left;position:sticky;top:0}th,td{padding:9px;border:1px solid #dbe2e6;vertical-align:top}a{color:#0067a3}.hidden{display:none}</style></head>
<body><header><h1>Mercator : $Title</h1><p>Export du $($ExportDate.ToString('dd/MM/yyyy HH:mm:ss'))</p></header><main><p><a href="../index.html">Retour index</a> | <a href="../JSON/$Endpoint.json">JSON</a> | <a href="../CSV/$Endpoint.csv">CSV</a></p><input id="q" placeholder="Rechercher..." oninput="f()"><div style="overflow:auto"><table id="t"><thead><tr>$Headers</tr></thead><tbody>$($Rows.ToString())</tbody></table></div></main><script>function f(){let q=document.getElementById('q').value.toLowerCase();document.querySelectorAll('#t tbody tr').forEach(r=>r.classList.toggle('hidden',!r.textContent.toLowerCase().includes(q)));}</script></body></html>
"@
    [IO.File]::WriteAllText($OutputFile,$Html,(New-Object Text.UTF8Encoding($true)))
}

function Get-ErrorText {
    param([System.Management.Automation.ErrorRecord]$Record)
    $Text = $Record.Exception.Message
    $Code = Get-HttpStatusCode $Record
    $Body = Get-HttpErrorBody $Record
    if ($null -ne $Code) { $Text = "HTTP $Code : $Text" }
    if (-not [string]::IsNullOrWhiteSpace($Body)) { $Text += " | $Body" }
    return $Text
}

Write-Log '=== Debut sauvegarde Mercator ==='
$Results = New-Object 'System.Collections.Generic.List[object]'
$OverallFailure = $false

try {
    if (-not (Test-Path -LiteralPath $CredPath -PathType Leaf)) { throw "Identifiants introuvables : $CredPath" }
    $Credential = Import-Clixml -LiteralPath $CredPath
    $LoginBody = @{ login=$Credential.UserName; password=$Credential.GetNetworkCredential().Password }
    $Auth = Invoke-RestMethod -Uri "$ApiBase/login" -Method Post -Body $LoginBody -ContentType 'application/x-www-form-urlencoded' -TimeoutSec $TimeoutSec
    $Token = [string]$Auth.access_token
    if ([string]::IsNullOrWhiteSpace($Token)) { throw 'Aucun access_token retourne.' }
    $Headers = @{ Authorization="Bearer $Token"; Accept='application/json' }
    Write-Log 'Authentification reussie'
} catch {
    Write-Log ("Echec authentification : {0}" -f (Get-ErrorText $_)) 'ERROR'
    exit 1
} finally { $LoginBody=$null; $Credential=$null }

foreach ($Endpoint in $Endpoints) {
    $Begin = Get-Date
    try {
        Write-Log "Export API $Endpoint"
        $Data = Invoke-WithRetry -OperationName "API $Endpoint" -Operation {
            Invoke-RestMethod -Uri "$ApiBase/$Endpoint" -Method Get -Headers $Headers -TimeoutSec $TimeoutSec -ErrorAction Stop
        }
        $JsonFile = Join-Path $JsonPath "$Endpoint.json"
        $CsvFile = Join-Path $CsvPath "$Endpoint.csv"
        $HtmlFile = Join-Path $HtmlPath "$Endpoint.html"
        [IO.File]::WriteAllText($JsonFile,($Data | ConvertTo-Json -Depth $JsonDepth),(New-Object Text.UTF8Encoding($true)))
        $Items = @(Get-EndpointItems $Data | Where-Object { $null -ne $_ })
        if ($Items.Count) { @(foreach ($Item in $Items) { ConvertTo-FlatObject $Item }) | Export-Csv $CsvFile -Delimiter ';' -NoTypeInformation -Encoding UTF8 }
        else { Set-Content $CsvFile 'Information;Aucun element retourne par API' -Encoding UTF8 }
        New-OfflineHtmlPage $Endpoint $Items $HtmlFile (Get-Date)
        $Duration=[math]::Round(((Get-Date)-$Begin).TotalSeconds,2)
        $Results.Add([pscustomobject]@{Endpoint=$Endpoint;Status='OK';Count=$Items.Count;Duration=$Duration;Json="JSON\$Endpoint.json";Csv="CSV\$Endpoint.csv";Html="HTML\$Endpoint.html";Word='';Error=''})
        Write-Log "OK : $Endpoint | $($Items.Count) element(s)"
    } catch {
        $OverallFailure=$true; $ErrorText=Get-ErrorText $_
        $Results.Add([pscustomobject]@{Endpoint=$Endpoint;Status='ERREUR';Count=0;Duration=[math]::Round(((Get-Date)-$Begin).TotalSeconds,2);Json='';Csv='';Html='';Word='';Error=$ErrorText})
        Write-Log "Erreur $Endpoint : $ErrorText" 'ERROR'
    }
    finally {
        Start-Sleep -Milliseconds $ApiDelayMilliseconds
    }
}

# RAPPORTS WORD
foreach ($Report in $WordReports) {
    Start-Sleep -Seconds $ReportDelaySeconds
    $Begin = Get-Date
    $WordFile = Join-Path $WordPath $Report.FileName
    try {
        Write-Log "Export rapport $($Report.Endpoint)"
        $ReportHeaders = @{ Authorization="Bearer $Token"; Accept='application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
        $Params = @{ Uri="$ApiBase/$($Report.Endpoint)"; Method='GET'; Headers=$ReportHeaders; OutFile=$WordFile; TimeoutSec=$TimeoutSec; ErrorAction='Stop' }
        if ($PSVersionTable.PSVersion.Major -le 5) { $Params.UseBasicParsing = $true }
        Invoke-WithRetry -OperationName "rapport $($Report.Name)" -Operation {
            Invoke-WebRequest @Params | Out-Null
        } | Out-Null
        if (-not (Test-Path $WordFile -PathType Leaf)) { throw 'Fichier non cree.' }
        $Bytes = [IO.File]::ReadAllBytes($WordFile)
        if ($Bytes.Length -lt 4 -or $Bytes[0] -ne 0x50 -or $Bytes[1] -ne 0x4B) {
            $Preview = [Text.Encoding]::UTF8.GetString($Bytes,0,[math]::Min($Bytes.Length,1000))
            throw "La reponse n'est pas un DOCX (signature PK absente). Reponse : $Preview"
        }
        $Duration=[math]::Round(((Get-Date)-$Begin).TotalSeconds,2)
        $Results.Add([pscustomobject]@{Endpoint=$Report.Endpoint;Status='OK';Count=1;Duration=$Duration;Json='';Csv='';Html='';Word="WORD\$($Report.FileName)";Error=''})
        Write-Log "OK : rapport $($Report.Name) | $([math]::Round($Bytes.Length/1MB,2)) Mo"
    } catch {
        $OverallFailure=$true; $ErrorText=Get-ErrorText $_
        if (Test-Path $WordFile) { Remove-Item $WordFile -Force -ErrorAction SilentlyContinue }
        $Results.Add([pscustomobject]@{Endpoint=$Report.Endpoint;Status='ERREUR';Count=0;Duration=[math]::Round(((Get-Date)-$Begin).TotalSeconds,2);Json='';Csv='';Html='';Word='';Error=$ErrorText})
        Write-Log "Erreur rapport $($Report.Name) : $ErrorText" 'ERROR'
    }
}

$Token=$null; $Headers=$null
$EndTime=Get-Date
$Results | Export-Csv $SummaryCsvFile -Delimiter ';' -NoTypeInformation -Encoding UTF8

# INDEX HTML
$Rows = foreach ($R in $Results) {
    $Links=@()
    foreach ($P in @('Html','Json','Csv','Word')) { if (-not [string]::IsNullOrWhiteSpace([string]$R.$P)) { $Href=([string]$R.$P).Replace('\','/'); $Links += "<a href='$Href'>$P</a>" } }
    '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f [Net.WebUtility]::HtmlEncode([string]$R.Endpoint),$R.Status,$R.Count,($Links -join ' | '),[Net.WebUtility]::HtmlEncode([string]$R.Error)
}
$Index=@"
<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>Mercator - Index de crise</title><style>body{font-family:Segoe UI,Arial;margin:0;background:#eef2f4}header{background:#003b5c;color:#fff;padding:25px 35px}main{padding:25px 35px}table{border-collapse:collapse;width:100%;background:#fff}th,td{padding:10px;border:1px solid #d7e0e5;text-align:left}th{background:#dfe9ef}a{color:#0067a3}</style></head><body><header><h1>Index de crise Mercator</h1><p>Export du $($EndTime.ToString('dd/MM/yyyy HH:mm:ss'))</p></header><main><table><thead><tr><th>Endpoint</th><th>Etat</th><th>Elements</th><th>Fichiers</th><th>Observation</th></tr></thead><tbody>$($Rows -join "`n")</tbody></table></main></body></html>
"@
[IO.File]::WriteAllText($IndexFile,$Index,(New-Object Text.UTF8Encoding($true)))

$Success=@($Results|Where-Object Status -eq 'OK').Count
$Failed=@($Results|Where-Object Status -ne 'OK').Count
$Control=[ordered]@{exportDate=$EndTime.ToString('o');apiBase=$ApiBase;computerName=$env:COMPUTERNAME;successfulExports=$Success;failedExports=$Failed;durationSeconds=[math]::Round(($EndTime-$StartTime).TotalSeconds,2);status=$(if($Failed-eq 0){'SUCCESS'}else{'PARTIAL'})}
[IO.File]::WriteAllText((Join-Path $ExportPath 'controle.json'),($Control|ConvertTo-Json),(New-Object Text.UTF8Encoding($true)))

$ZipCreated=$false
if ($CreateZip) {
    try {
        if (Test-Path $ZipFile) { Remove-Item $ZipFile -Force }
        Compress-Archive -Path "$ExportPath\*" -DestinationPath $ZipFile -CompressionLevel Optimal -Force
        $ZipCreated=$true; Write-Log "Archive creee : $ZipFile"
    } catch { $OverallFailure=$true; Write-Log "Erreur ZIP : $($_.Exception.Message)" 'ERROR' }
}
if ($RetentionDays -gt 0) {
    $Limit=(Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem $BackupRoot -Filter 'Mercator_*.zip' -File -ErrorAction SilentlyContinue | Where-Object LastWriteTime -lt $Limit | Remove-Item -Force
    Get-ChildItem $BackupRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.FullName-ne$ExportPath -and $_.LastWriteTime-lt$Limit -and $_.Name-match '^\d{4}-\d{2}-\d{2}_' } | Remove-Item -Recurse -Force
}
if ($RemoveFolderAfterZip -and $ZipCreated) { Remove-Item $ExportPath -Recurse -Force }
Write-Log "Resultat : $Success succes, $Failed echec(s)"
Write-Log '=== Fin sauvegarde Mercator ==='
if ($OverallFailure) { exit 2 }
exit 0
