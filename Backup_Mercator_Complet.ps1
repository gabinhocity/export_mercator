<#
.SYNOPSIS
    Sauvegarde Mercator 2026.08.28 via API REST.
.DESCRIPTION
    Exporte toutes les ressources GET documentees de l API en JSON, CSV et HTML.
    Exporte tous les rapports documentes et detecte automatiquement leur format
    DOCX ou XLSX. Cree un index de crise,
    un index historique, une synthese CSV, un controle JSON et une archive ZIP.
#>
#Requires -Version 5.1
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# CONFIGURATION
$ApiBase = '#https://#Lien_MERCATOR/api'
$CredPath = '#C:\emplacement_fichier_cred\mercator_cred.xm\mercator_cred.xml'
$BackupRoot = '#emplacement_destination_de_la_sauvegarde'
$TimeoutSec = 300
$JsonDepth = 100
$RetentionDays = 30
$CreateZip = $true
$RemoveFolderAfterZip = $false
$ApiDelayMilliseconds = 500
$ReportDelaySeconds = 5
$MaxRetryAttempts = 5
$InitialRetryDelaySeconds = 10
$MaximumRetryDelaySeconds = 120

$Endpoints = @(
    # Ecosysteme
    'entities','relations',

    # Metier du systeme d information
    'macro-processuses','processes','activities','operations','tasks','actors','information',

    # Applications
    'application-blocks','applications','application-services','application-modules',
    'databases','application-flows',

    # Administration
    'zone-admins','annuaires','forest-ads','domains','admin-users',

    # Infrastructure logique
    'networks','subnetworks','gateways','external-connected-entities',
    'network-switches','routers','security-devices','dhcp-servers','dnsservers',
    'clusters','logical-servers','backups','logical-flows','containers','certificates','vlans',

    # Infrastructure physique
    'sites','buildings','bays','zones','physical-servers','workstations',
    'storage-devices','peripherals','phones','physical-switches','physical-routers',
    'wifi-terminals','physical-security-devices','physical-links','wans','mans','lans',

    # Configuration et administration de Mercator
    'users','roles','cartographers','permissions','documents',

    # Requetes et ressources presentes sur l instance
    'queries','security-controls'
)


# Les endpoints zones, users, roles, cartographers et permissions peuvent retourner HTTP 403 selon le role API.
# L endpoint perimeters peut retourner HTTP 404 si cette route n existe pas dans la version installee.
# Tous les rapports documentes sont testes. Le format DOCX/XLSX est detecte apres telechargement.
$Reports = @(
   #[pscustomobject]@{ Name='cartography';          Endpoint='report/cartography' },
    [pscustomobject]@{ Name='entities';             Endpoint='report/entities' },
    [pscustomobject]@{ Name='applicationsByBlocks'; Endpoint='report/applicationsByBlocks' },
    [pscustomobject]@{ Name='directory';            Endpoint='report/directory' },
    [pscustomobject]@{ Name='logicalServers';       Endpoint='report/logicalServers' },
    [pscustomobject]@{ Name='securityNeeds';        Endpoint='report/securityNeeds' },
    [pscustomobject]@{ Name='logicalServerConfigs'; Endpoint='report/logicalServerConfigs' },
    [pscustomobject]@{ Name='externalAccess';       Endpoint='report/externalAccess' },
    [pscustomobject]@{ Name='physicalInventory';    Endpoint='report/physicalInventory' },
    [pscustomobject]@{ Name='vlans';                Endpoint='report/vlans' },
    [pscustomobject]@{ Name='workstations';         Endpoint='report/workstations' },
    [pscustomobject]@{ Name='cve';                  Endpoint='report/cve' },
    [pscustomobject]@{ Name='activityList';         Endpoint='report/activityList' },
    [pscustomobject]@{ Name='activityReport';       Endpoint='report/activityReport' },
    [pscustomobject]@{ Name='impacts';              Endpoint='report/impacts' },
    [pscustomobject]@{ Name='rto';                  Endpoint='report/rto' }
)

# DOSSIERS
$StartTime = Get-Date
$DateStamp = $StartTime.ToString('yyyy-MM-dd_HH-mm-ss')
$ExportPath = Join-Path $BackupRoot $DateStamp
$JsonPath = Join-Path $ExportPath 'JSON'
$CsvPath = Join-Path $ExportPath 'CSV'
$HtmlPath = Join-Path $ExportPath 'HTML'
$ReportPath = Join-Path $ExportPath 'RAPPORTS'
$ErrorPath = Join-Path $ExportPath 'ERRORS'
$LogFile = Join-Path $ExportPath 'BackupMercator.log'
$IndexFile = Join-Path $ExportPath 'index.html'
$SummaryCsvFile = Join-Path $ExportPath 'synthese.csv'
$ControlFile = Join-Path $ExportPath 'controle.json'
$ZipFile = Join-Path $BackupRoot (('Mercator_{0}.zip') -f $DateStamp)
$GlobalIndexFile = Join-Path $BackupRoot 'index.html'
foreach ($Directory in @($BackupRoot,$ExportPath,$JsonPath,$CsvPath,$HtmlPath,$ReportPath,$ErrorPath)) {
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        $null = New-Item -ItemType Directory -Path $Directory -Force
    }
}

function Write-Log {
    param([Parameter(Mandatory=$true)][string]$Message,[ValidateSet('INFO','WARNING','ERROR')][string]$Level='INFO')
    $Line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Host $Line
    try { Add-Content -LiteralPath $LogFile -Value $Line -Encoding UTF8 } catch { Write-Warning $_.Exception.Message }
}
function Write-Utf8File {
    param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)][AllowEmptyString()][string]$Content)
    [IO.File]::WriteAllText($Path,$Content,(New-Object Text.UTF8Encoding($true)))
}
function ConvertTo-HtmlText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return [Net.WebUtility]::HtmlEncode([string]$Value)
}
function Get-HttpStatusCode {
    param([Parameter(Mandatory=$true)][Management.Automation.ErrorRecord]$ErrorRecord)
    try { if ($null -ne $ErrorRecord.Exception.Response.StatusCode) { return [int]$ErrorRecord.Exception.Response.StatusCode } } catch {}
    return $null
}
function Get-HttpErrorBody {
    param([Parameter(Mandatory=$true)][Management.Automation.ErrorRecord]$ErrorRecord)
    if ($null -ne $ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.ErrorDetails.Message)) { return [string]$ErrorRecord.ErrorDetails.Message }
    try {
        $Response=$ErrorRecord.Exception.Response
        if ($null -eq $Response) { return '' }
        if ($Response.PSObject.Methods.Name -contains 'GetResponseStream') {
            $Stream=$Response.GetResponseStream(); if ($null -eq $Stream) { return '' }
            $Reader=New-Object IO.StreamReader($Stream)
            try { return $Reader.ReadToEnd() } finally { $Reader.Dispose(); $Stream.Dispose() }
        }
        if ($Response.PSObject.Properties.Name -contains 'Content') { return [string]$Response.Content }
    } catch {}
    return ''
}
function Get-ErrorText {
    param([Parameter(Mandatory=$true)][Management.Automation.ErrorRecord]$Record)
    $Text=$Record.Exception.Message; $Code=Get-HttpStatusCode $Record; $Body=Get-HttpErrorBody $Record
    if ($null -ne $Code) { $Text="HTTP $Code : $Text" }
    if (-not [string]::IsNullOrWhiteSpace($Body)) { $Text+=" | $Body" }
    return $Text
}
function Invoke-WithRetry {
    param([Parameter(Mandatory=$true)][scriptblock]$Operation,[Parameter(Mandatory=$true)][string]$OperationName)
    for ($Attempt=1; $Attempt -le $MaxRetryAttempts; $Attempt++) {
        try { return & $Operation }
        catch {
            $Code=Get-HttpStatusCode $_
            $Transient=($null -eq $Code)-or($Code -eq 429)-or($Code -ge 500)
            if (-not $Transient -or $Attempt -ge $MaxRetryAttempts) { throw }
            $Delay=[int][Math]::Min($MaximumRetryDelaySeconds,($InitialRetryDelaySeconds*[Math]::Pow(2,$Attempt-1))+(Get-Random -Minimum 0 -Maximum 4))
            Write-Log "Erreur transitoire pour $OperationName (HTTP $Code). Reprise dans ${Delay}s." 'WARNING'
            Start-Sleep -Seconds $Delay
        }
    }
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
    $Flat=[ordered]@{}
    foreach ($Property in $InputObject.PSObject.Properties) {
        $Value=$Property.Value
        if ($null -eq $Value -or $Value -is [string] -or $Value -is [ValueType]) { $Flat[$Property.Name]=$Value }
        else { $Flat[$Property.Name]=($Value|ConvertTo-Json -Depth $JsonDepth -Compress) }
    }
    return [pscustomobject]$Flat
}
function Get-OfficePackageType {
    <# Detecte le type reel du paquet Office telecharge. #>
    param([Parameter(Mandatory=$true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{IsValid=$false;Extension='';MimeType='';Reason='Fichier absent.'}
    }

    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $Archive=[IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $Names=@($Archive.Entries | ForEach-Object { $_.FullName })
            if ($Names -notcontains '[Content_Types].xml' -or $Names -notcontains '_rels/.rels') {
                return [pscustomobject]@{IsValid=$false;Extension='';MimeType='';Reason='Paquet Open XML incomplet.'}
            }

            if ($Names -contains 'word/document.xml') {
                $MainEntry='word/document.xml'
                $Extension='.docx'
                $MimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }
            elseif ($Names -contains 'xl/workbook.xml') {
                $MainEntry='xl/workbook.xml'
                $Extension='.xlsx'
                $MimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }
            else {
                return [pscustomobject]@{IsValid=$false;Extension='';MimeType='';Reason='Ni word/document.xml ni xl/workbook.xml dans le paquet.'}
            }

            $Entry=$Archive.GetEntry($MainEntry)
            $Stream=$Entry.Open()
            $Reader=New-Object IO.StreamReader($Stream)
            try {
                $Xml=New-Object Xml.XmlDocument
                $Xml.LoadXml($Reader.ReadToEnd())
            }
            finally {
                $Reader.Dispose()
                $Stream.Dispose()
            }
        }
        finally { $Archive.Dispose() }

        return [pscustomobject]@{IsValid=$true;Extension=$Extension;MimeType=$MimeType;Reason=''}
    }
    catch {
        return [pscustomobject]@{IsValid=$false;Extension='';MimeType='';Reason=$_.Exception.Message}
    }
}

function Get-BaseCss {
@'
:root{--ivo:#4F9B9F;--ivo-light:#BFDDDF;--dark:#595959;--ok:#ACE3D1;--ko:#EFB5B9;--bg:#f5f8f8;font-family:"Segoe UI",Arial,sans-serif}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:#303030}header{background:var(--ivo);color:#fff;padding:22px 30px;border-bottom:6px solid var(--ivo-light)}header h1{margin:0 0 6px}main{padding:22px 30px}.toolbar{display:flex;flex-wrap:wrap;gap:12px;align-items:center;margin-bottom:18px}.search{flex:1 1 300px;padding:11px;border:1px solid #aab5bd;border-radius:6px}.table-wrapper{overflow:auto;background:#fff;border:1px solid #d6dde2;border-radius:8px}table{border-collapse:collapse;width:100%;font-size:13px}th{background:var(--dark);color:#fff;text-align:left;padding:10px}td{padding:9px 10px;border-bottom:1px solid #e2e7ea;vertical-align:top;word-break:break-word}tbody tr:nth-child(even){background:#f5fafa}a{color:#356f72;font-weight:700;text-decoration:none;margin-right:12px}.status-ok,.status-error{display:inline-block;padding:4px 9px;border-radius:12px;font-weight:700}.status-ok{background:var(--ok);color:#174f3d}.status-error{background:var(--ko);color:#6d232b}.hidden{display:none}footer{margin-top:18px;color:#63747e;font-size:12px}@media(max-width:760px){header,main{padding:15px}.table-wrapper{border:0;background:transparent}table,thead,tbody,th,td,tr{display:block}thead{position:absolute;left:-9999px}tbody tr{background:#fff;border:1px solid #d6dde2;border-radius:8px;margin-bottom:10px}td{position:relative;padding-left:43%;min-height:35px}td:before{position:absolute;left:10px;width:38%;font-weight:700;color:#356f72;content:attr(data-label)}}
'@
}
function New-OfflineHtmlPage {
    param([string]$Endpoint,[array]$Items,[string]$OutputFile,[datetime]$ExportDate)
    $Rows=New-Object Text.StringBuilder
    if ($Items.Count -eq 0) { $Columns=@('Information');$null=$Rows.AppendLine('<tr><td data-label="Information">Aucun element retourne par API.</td></tr>') }
    else {
        $FlatItems=@(foreach($Item in $Items){ConvertTo-FlatObject $Item})
        $Columns=@($FlatItems|ForEach-Object{$_.PSObject.Properties.Name}|Select-Object -Unique)
        foreach($Item in $FlatItems){$null=$Rows.AppendLine('<tr>');foreach($Column in $Columns){$P=$Item.PSObject.Properties[$Column];$V=if($null -eq $P -or $null -eq $P.Value){''}else{ConvertTo-HtmlText $P.Value};$EncodedColumn=ConvertTo-HtmlText $Column;$null=$Rows.AppendLine('<td data-label="'+$EncodedColumn+'">'+$V+'</td>')};$null=$Rows.AppendLine('</tr>')}
    }
    $Headers=@($Columns|ForEach-Object{'<th>'+(ConvertTo-HtmlText $_)+'</th>'}) -join ''
    $Title=ConvertTo-HtmlText $Endpoint
    $Html=@"
<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Mercator - $Title</title><style>$(Get-BaseCss)</style></head><body><header><h1>Mercator : $Title</h1><p>Export du $($ExportDate.ToString('dd/MM/yyyy HH:mm:ss'))</p></header><main><div class="toolbar"><a href="../index.html">Index de crise</a><a href="../../index.html">Historique</a><a href="../JSON/$Endpoint.json">JSON</a><a href="../CSV/$Endpoint.csv">CSV</a><input id="q" class="search" placeholder="Rechercher..." oninput="f()"></div><div class="table-wrapper"><table id="t"><thead><tr>$Headers</tr></thead><tbody>$($Rows.ToString())</tbody></table></div><footer>Export INOVALYS hors ligne.</footer></main><script>function f(){var q=document.getElementById('q').value.toLowerCase();document.querySelectorAll('#t tbody tr').forEach(function(r){r.classList.toggle('hidden',!r.textContent.toLowerCase().includes(q));});}</script></body></html>
"@
    Write-Utf8File $OutputFile $Html
}
function New-CrisisIndex {
    param([array]$Results,[string]$OutputFile,[datetime]$StartDate,[datetime]$EndDate)
    $Rows=New-Object Text.StringBuilder
    foreach($R in $Results){
        # Tableau simple volontairement utilise pour eviter ArgumentException sur List[string].
        $Links=@()
        foreach($Property in @('Html','Json','Csv','Report')){
            $Value=''
            if($R.PSObject.Properties.Name -contains $Property){$Value=[string]$R.$Property}
            if(-not[string]::IsNullOrWhiteSpace($Value)){$Href=$Value.Replace('\','/');$Links+=('<a href="'+(ConvertTo-HtmlText $Href)+'">'+$Property+'</a>')}
        }
        $Class=if($R.Status -eq 'OK'){'status-ok'}else{'status-error'}
        $EndpointHtml=ConvertTo-HtmlText $R.Endpoint;$StatusHtml=ConvertTo-HtmlText $R.Status;$ErrorHtml=ConvertTo-HtmlText $R.Error;$LinksHtml=$Links -join ' ';$null=$Rows.AppendLine('<tr><td data-label="Endpoint">'+$EndpointHtml+'</td><td data-label="Etat"><span class="'+$Class+'">'+$StatusHtml+'</span></td><td data-label="Elements">'+[string]$R.Count+'</td><td data-label="Fichiers">'+$LinksHtml+'</td><td data-label="Observation">'+$ErrorHtml+'</td></tr>')
    }
    $Success=@($Results|Where-Object{$_.Status -eq 'OK'}).Count;$Failed=$Results.Count-$Success
    $Html=@"
<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Mercator - Index de crise</title><style>$(Get-BaseCss)</style></head><body><header><h1>Index de crise Mercator</h1><p>Export du $($EndDate.ToString('dd/MM/yyyy HH:mm:ss'))</p></header><main><div class="toolbar"><input id="q" class="search" placeholder="Rechercher..." oninput="f()"><a href="../index.html">Historique global</a><a href="synthese.csv">Synthese CSV</a></div><div class="table-wrapper"><table id="t"><thead><tr><th>Endpoint</th><th>Etat</th><th>Elements</th><th>Fichiers</th><th>Observation</th></tr></thead><tbody>$($Rows.ToString())</tbody></table></div><footer>$Success succes, $Failed erreur(s). Debut $($StartDate.ToString('dd/MM/yyyy HH:mm:ss')).</footer></main><script>function f(){var q=document.getElementById('q').value.toLowerCase();document.querySelectorAll('#t tbody tr').forEach(function(r){r.classList.toggle('hidden',!r.textContent.toLowerCase().includes(q));});}</script></body></html>
"@
    Write-Utf8File $OutputFile $Html
}
function New-GlobalIndex {
    param([string]$Root,[string]$OutputFile)
    $Rows=New-Object Text.StringBuilder
    Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue|Where-Object{$_.Name -match '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}$'}|Sort-Object Name -Descending|ForEach-Object{if(Test-Path (Join-Path $_.FullName 'index.html')){$FolderName=ConvertTo-HtmlText $_.Name;$null=$Rows.AppendLine('<tr><td>'+$FolderName+'</td><td><a href="'+$FolderName+'/index.html">Index de crise</a></td></tr>')}}
    $Html=@"
<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>Mercator - Historique</title><style>$(Get-BaseCss)</style></head><body><header><h1>Historique des sauvegardes Mercator</h1></header><main><div class="table-wrapper"><table><thead><tr><th>Sauvegarde</th><th>Acces</th></tr></thead><tbody>$($Rows.ToString())</tbody></table></div></main></body></html>
"@
    Write-Utf8File $OutputFile $Html
}

# EXECUTION
Write-Log '=== Debut sauvegarde Mercator ==='
$Results=New-Object 'System.Collections.Generic.List[object]'
$OverallFailure=$false
try{
    if(-not(Test-Path -LiteralPath $CredPath -PathType Leaf)){throw "Identifiants introuvables : $CredPath"}
    $Credential=Import-Clixml -LiteralPath $CredPath
    $LoginBody=@{login=$Credential.UserName;password=$Credential.GetNetworkCredential().Password}
    $Auth=Invoke-RestMethod -Uri "$ApiBase/login" -Method Post -Body $LoginBody -ContentType 'application/x-www-form-urlencoded' -TimeoutSec $TimeoutSec -ErrorAction Stop
    $Token=[string]$Auth.access_token
    if([string]::IsNullOrWhiteSpace($Token)){throw 'Aucun access_token retourne.'}
    $Headers=@{Authorization="Bearer $Token";Accept='application/json'}
    Write-Log 'Authentification reussie'
}catch{Write-Log ("Echec authentification : {0}" -f (Get-ErrorText $_)) 'ERROR';exit 1}finally{$LoginBody=$null;$Credential=$null}

foreach($Endpoint in $Endpoints){
    $Begin=Get-Date
    try{
        Write-Log "Export API $Endpoint"
        $Data=Invoke-WithRetry -OperationName "API $Endpoint" -Operation {Invoke-RestMethod -Uri "$ApiBase/$Endpoint" -Method Get -Headers $Headers -TimeoutSec $TimeoutSec -ErrorAction Stop}
        $JsonFile=Join-Path $JsonPath "$Endpoint.json";$CsvFile=Join-Path $CsvPath "$Endpoint.csv";$HtmlFile=Join-Path $HtmlPath "$Endpoint.html"
        Write-Utf8File $JsonFile ($Data|ConvertTo-Json -Depth $JsonDepth)
        $Items=@(Get-EndpointItems $Data|Where-Object{$null -ne $_})
        if($Items.Count -gt 0){@(foreach($Item in $Items){ConvertTo-FlatObject $Item})|Export-Csv -LiteralPath $CsvFile -Delimiter ';' -NoTypeInformation -Encoding UTF8}else{Write-Utf8File $CsvFile "Information`r`nAucun element retourne par API"}
        New-OfflineHtmlPage $Endpoint $Items $HtmlFile (Get-Date)
        $Results.Add([pscustomobject]@{Endpoint=$Endpoint;Status='OK';Count=$Items.Count;Duration=[Math]::Round(((Get-Date)-$Begin).TotalSeconds,2);Json="JSON\$Endpoint.json";Csv="CSV\$Endpoint.csv";Html="HTML\$Endpoint.html";Report='';Error=''})
    }catch{$OverallFailure=$true;$ErrorText=Get-ErrorText $_;$Results.Add([pscustomobject]@{Endpoint=$Endpoint;Status='ERREUR';Count=0;Duration=[Math]::Round(((Get-Date)-$Begin).TotalSeconds,2);Json='';Csv='';Html='';Report='';Error=$ErrorText});Write-Log "Erreur $Endpoint : $ErrorText" 'ERROR'}finally{Start-Sleep -Milliseconds $ApiDelayMilliseconds}
}

foreach($Report in $Reports){
    Start-Sleep -Seconds $ReportDelaySeconds
    $Begin=Get-Date
    $TempFile=Join-Path $ReportPath ($Report.Name + '.download')
    Remove-Item -LiteralPath $TempFile -Force -ErrorAction SilentlyContinue

    try{
        Write-Log "Export rapport $($Report.Endpoint)"
        $ReportHeaders=@{
            Authorization="Bearer $Token"
            Accept='application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/octet-stream'
        }
        $Params=@{
            Uri="$ApiBase/$($Report.Endpoint)"
            Method='GET'
            Headers=$ReportHeaders
            OutFile=$TempFile
            TimeoutSec=$TimeoutSec
            ErrorAction='Stop'
        }
        if($PSVersionTable.PSVersion.Major -le 5){$Params.UseBasicParsing=$true}

        Invoke-WithRetry -OperationName "rapport $($Report.Name)" -Operation {
            Invoke-WebRequest @Params | Out-Null
        } | Out-Null

        $Package=Get-OfficePackageType -Path $TempFile
        if(-not $Package.IsValid){
            $DiagnosticFile=Join-Path $ErrorPath ($Report.Name + '.error.txt')
            $Preview=''
            try{
                $Bytes=[IO.File]::ReadAllBytes($TempFile)
                $Length=[Math]::Min($Bytes.Length,4000)
                if($Length -gt 0){$Preview=[Text.Encoding]::UTF8.GetString($Bytes,0,$Length)}
            }catch{}
            Write-Utf8File -Path $DiagnosticFile -Content ("Raison : " + $Package.Reason + "`r`n`r`nApercu :`r`n" + $Preview)
            throw "Rapport Office invalide : $($Package.Reason). Diagnostic : $DiagnosticFile"
        }

        $FinalFileName=$Report.Name + $Package.Extension
        $FinalFile=Join-Path $ReportPath $FinalFileName
        Remove-Item -LiteralPath $FinalFile -Force -ErrorAction SilentlyContinue
        Move-Item -LiteralPath $TempFile -Destination $FinalFile -Force

        $Results.Add([pscustomobject]@{
            Endpoint=$Report.Endpoint
            Status='OK'
            Count=1
            Duration=[Math]::Round(((Get-Date)-$Begin).TotalSeconds,2)
            Json=''
            Csv=''
            Html=''
            Report="RAPPORTS\$FinalFileName"
            Error=''
        })
        Write-Log "OK : rapport $($Report.Name) | format $($Package.Extension)"
    }
    catch{
        $OverallFailure=$true
        $ErrorText=Get-ErrorText $_
        Remove-Item -LiteralPath $TempFile -Force -ErrorAction SilentlyContinue
        $Results.Add([pscustomobject]@{
            Endpoint=$Report.Endpoint
            Status='ERREUR'
            Count=0
            Duration=[Math]::Round(((Get-Date)-$Begin).TotalSeconds,2)
            Json=''
            Csv=''
            Html=''
            Report=''
            Error=$ErrorText
        })
        Write-Log "Erreur rapport $($Report.Name) : $ErrorText" 'ERROR'
    }
}

$EndTime=Get-Date;$Token=$null;$Headers=$null
$Results|Export-Csv -LiteralPath $SummaryCsvFile -Delimiter ';' -NoTypeInformation -Encoding UTF8
# Conversion explicite en tableau standard pour eviter ArgumentException.
$ResultsArray=@($Results|ForEach-Object{$_})
New-CrisisIndex -Results $ResultsArray -OutputFile $IndexFile -StartDate $StartTime -EndDate $EndTime
if(-not(Test-Path $IndexFile -PathType Leaf)){throw "Index non cree : $IndexFile"}

$Success=@($ResultsArray|Where-Object{$_.Status -eq 'OK'}).Count;$Failed=$ResultsArray.Count-$Success;$TotalObjects=0
foreach($R in @($ResultsArray|Where-Object{$_.Status -eq 'OK'})){$TotalObjects+=[int]$R.Count}
$Control=[ordered]@{exportDate=$EndTime.ToString('o');mercatorVersion='2026.08.28';apiBase=$ApiBase;successfulExports=$Success;failedExports=$Failed;totalObjects=$TotalObjects;durationSeconds=[Math]::Round(($EndTime-$StartTime).TotalSeconds,2);status=$(if($Failed -eq 0){'SUCCESS'}else{'PARTIAL'})}
Write-Utf8File $ControlFile ($Control|ConvertTo-Json -Depth 10)
New-GlobalIndex $BackupRoot $GlobalIndexFile

$ZipCreated=$false
if($CreateZip){try{if(Test-Path $ZipFile){Remove-Item $ZipFile -Force};Compress-Archive -Path (Join-Path $ExportPath '*') -DestinationPath $ZipFile -CompressionLevel Optimal -Force -ErrorAction Stop;$ZipCreated=$true;Write-Log "Archive creee : $ZipFile"}catch{$OverallFailure=$true;Write-Log "Erreur ZIP : $($_.Exception.Message)" 'ERROR'}}
if($RetentionDays -gt 0){$Limit=(Get-Date).AddDays(-$RetentionDays);Get-ChildItem $BackupRoot -Filter 'Mercator_*.zip' -File -ErrorAction SilentlyContinue|Where-Object{$_.LastWriteTime -lt $Limit}|Remove-Item -Force -ErrorAction Continue;Get-ChildItem $BackupRoot -Directory -ErrorAction SilentlyContinue|Where-Object{$_.FullName -ne $ExportPath -and $_.LastWriteTime -lt $Limit -and $_.Name -match '^\d{4}-\d{2}-\d{2}_'}|Remove-Item -Recurse -Force -ErrorAction Continue}
if($RemoveFolderAfterZip -and $ZipCreated){Remove-Item $ExportPath -Recurse -Force}
Write-Log "Resultat : $Success succes, $Failed echec(s)"
Write-Log "Index : $IndexFile"
Write-Log '=== Fin sauvegarde Mercator ==='
if($OverallFailure){exit 2}
exit 0
