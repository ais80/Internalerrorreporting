# ==============================================================================
# Sync-NearMissMaster.ps1
# UHBFT Secondary Care Pharmacy Near Miss Reporting & Learning System
# Master Consolidation Engine (Phase 3)
# Consolidates atomic events & import queue into Excel, CSV, JS & Telemetry
# ==============================================================================
[CmdletBinding()]
param(
    [string]$DataDir = "",
    [string]$OutXlsx = "",
    [string]$OutCsv  = "",
    [string]$OutJs   = "",
    [string]$OutTelemetry = ""
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = $PSScriptRoot }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }

if (-not $DataDir) { $DataDir = Join-Path $scriptDir "data" }
if (-not $OutXlsx) { $OutXlsx = Join-Path $scriptDir "Near_Miss_Master_Log.xlsx" }
if (-not $OutCsv)  { $OutCsv  = Join-Path $scriptDir "Near_Miss_Master_Log.csv" }
if (-not $OutJs)   { $OutJs   = Join-Path $scriptDir "nearmiss_shared_data.js" }
if (-not $OutTelemetry) { $OutTelemetry = Join-Path $scriptDir "nearmiss_telemetry.json" }

$DataDir = [System.IO.Path]::GetFullPath($DataDir)
$OutXlsx = [System.IO.Path]::GetFullPath($OutXlsx)
$OutCsv  = [System.IO.Path]::GetFullPath($OutCsv)
$OutJs   = [System.IO.Path]::GetFullPath($OutJs)
$OutTelemetry = [System.IO.Path]::GetFullPath($OutTelemetry)

$EventsDir = Join-Path $DataDir "events"
$ImportQueueDir = Join-Path $DataDir "_import_queue"
$ProcessedQueueDir = Join-Path $ImportQueueDir "_processed"
$BackupsDir = Join-Path $DataDir "_backups"

foreach ($d in @($DataDir, $EventsDir, $ImportQueueDir, $ProcessedQueueDir, $BackupsDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " UHBFT Pharmacy Near Miss - Consolidating Master Datasets                       " -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " Data Directory    : $DataDir"
Write-Host " Target Excel      : $OutXlsx"
Write-Host " Target CSV        : $OutCsv"
Write-Host " Target Shared JS  : $OutJs"
Write-Host " Target Telemetry  : $OutTelemetry"
Write-Host ""

$u8NoBom = [System.Text.UTF8Encoding]::new($false)
$u8WithBom = [System.Text.UTF8Encoding]::new($true)

# 1. Process Import Queue Drops (e.g. offline queue exports)
$queueFiles = Get-ChildItem -Path $ImportQueueDir -Filter "*.json" -File -ErrorAction SilentlyContinue | Where-Object { $_.DirectoryName -notlike "*_processed*" }
if ($queueFiles.Count -gt 0) {
    Write-Host "Found $($queueFiles.Count) file(s) in _import_queue. Ingesting into events store..." -ForegroundColor Yellow
    foreach ($qf in $queueFiles) {
        try {
            $raw = [System.IO.File]::ReadAllText($qf.FullName, $u8NoBom)
            $parsed = ConvertFrom-Json $raw -ErrorAction Stop
            
            # If array of events
            $eventsToIngest = @()
            if ($parsed -is [System.Collections.IEnumerable] -and $parsed -isnot [string]) {
                $eventsToIngest = $parsed
            } else {
                $eventsToIngest = @($parsed)
            }

            foreach ($ev in $eventsToIngest) {
                if ($ev.eventId) {
                    $mFolder = if ($ev.submittedAtIso -and $ev.submittedAtIso.Length -ge 7) { $ev.submittedAtIso.Substring(0, 7) } else { (Get-Date).ToString("yyyy-MM") }
                    $destFolder = Join-Path $EventsDir $mFolder
                    if (-not (Test-Path $destFolder)) { New-Item -ItemType Directory -Path $destFolder -Force | Out-Null }
                    $destPath = Join-Path $destFolder "$($ev.eventId).json"
                    $jsonStr = ConvertTo-Json $ev -Depth 10
                    [System.IO.File]::WriteAllText($destPath, $jsonStr, $u8NoBom)
                }
            }

            Move-Item -Path $qf.FullName -Destination (Join-Path $ProcessedQueueDir $qf.Name) -Force
            Write-Host "  Ingested and archived $($qf.Name)" -ForegroundColor DarkGreen
        } catch {
            Write-Warning "Failed to ingest queue file $($qf.Name): $_"
        }
    }
}

# 2. Gather All Atomic Event Files (Deduplicating by eventId)
$eventFiles = Get-ChildItem -Path $EventsDir -Filter "NM-*.json" -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.DirectoryName -notlike "*_undone*" }
Write-Host "Discovered $($eventFiles.Count) atomic event file(s) across event partitions." -ForegroundColor Green

$eventMap = [System.Collections.Generic.Dictionary[string, object]]::new()
foreach ($ef in $eventFiles) {
    try {
        $raw = [System.IO.File]::ReadAllText($ef.FullName, $u8NoBom)
        $ev = ConvertFrom-Json $raw -ErrorAction Stop
        if ($ev.eventId) {
            $eventMap[$ev.eventId] = $ev
        }
    } catch {
        Write-Warning "Error reading event file $($ef.FullName): $_"
    }
}

# Sort chronologically by submittedAtIso
$events = [System.Collections.Generic.List[object]]::new()
foreach ($key in $eventMap.Keys) {
    $events.Add($eventMap[$key])
}
$events.Sort({
    param($a, $b)
    $da = if ($a.submittedAtIso) { $a.submittedAtIso } else { "" }
    $db = if ($b.submittedAtIso) { $b.submittedAtIso } else { "" }
    return $da.CompareTo($db)
})

Write-Host "Total unique near-miss events to consolidate: $($events.Count)" -ForegroundColor Cyan

# 3. Create Pre-Sync Snapshot Backup if existing files exist
if (Test-Path $OutXlsx) {
    $bkTime = Get-Date -Format "yyyyMMdd-HHmmss"
    $bkName = "Near_Miss_Master_Log_$bkTime.xlsx"
    Copy-Item -Path $OutXlsx -Destination (Join-Path $BackupsDir $bkName) -Force
}

# 4. Standard Column Definitions
$cols = @(
    'EventID', 'SubmittedAt', 'Site', 'SiteOther', 'NearMissType',
    'IsSatelliteUnit', 'SatelliteUnitName', 'CompletedBy', 'CompletedOn', 'TimeOfDay',
    'ErrorCategories',
    'ErrorDrugName', 'ErrorDrugCMM', 'ErrorDrugVTMID', 'ErrorDrugVTMName',
    'ErrorDrugVMPID', 'ErrorDrugVMPName', 'ErrorDrugAMPID', 'ErrorDrugAMPName',
    'ErrorDrugAMPPID', 'ErrorDrugGTIN14', 'ErrorDrugBatch', 'ErrorDrugExpiry',
    'ErrorDrugIsExpired', 'ErrorDrugIsNTI',
    'IntendedDrugName', 'IntendedDrugCMM', 'IntendedDrugVTMID', 'IntendedDrugVMPID',
    'IntendedDrugAMPID', 'IntendedDrugIsNTI',
    'QuantityDispensed', 'QuantityPrescribed', 'ContributingFactors', 'LearningNote'
)

function Get-EventRowValues($ev) {
    $cats = if ($ev.errorCategories) { ($ev.errorCategories -join "; ") } else { "" }
    $factors = if ($ev.contributingFactors) { ($ev.contributingFactors -join "; ") } else { "" }

    $ed = $ev.errorDrug
    $id = $ev.intendedDrug

    return @{
        'EventID'             = [string]$ev.eventId
        'SubmittedAt'         = [string]$ev.submittedAtIso
        'Site'                = [string]$ev.site
        'SiteOther'           = [string]$ev.siteOther
        'NearMissType'        = [string]$ev.nearMissType
        'IsSatelliteUnit'     = [string]$ev.isSatelliteUnit
        'SatelliteUnitName'   = [string]$ev.satelliteUnitName
        'CompletedBy'         = [string]$ev.completedBy
        'CompletedOn'         = [string]$ev.completedOn
        'TimeOfDay'           = [string]$ev.timeOfDay
        'ErrorCategories'     = $cats
        
        'ErrorDrugName'       = if ($ed) { [string]$ed.displayName } else { "" }
        'ErrorDrugCMM'        = if ($ed) { [string]$ed.cmmCode } else { "" }
        'ErrorDrugVTMID'      = if ($ed) { [string]$ed.vtmId } else { "" }
        'ErrorDrugVTMName'    = if ($ed) { [string]$ed.vtmName } else { "" }
        'ErrorDrugVMPID'      = if ($ed) { [string]$ed.vmpId } else { "" }
        'ErrorDrugVMPName'    = if ($ed) { [string]$ed.vmpName } else { "" }
        'ErrorDrugAMPID'      = if ($ed) { [string]$ed.ampId } else { "" }
        'ErrorDrugAMPName'    = if ($ed) { [string]$ed.ampName } else { "" }
        'ErrorDrugAMPPID'     = if ($ed) { [string]$ed.amppId } else { "" }
        'ErrorDrugGTIN14'     = if ($ed) { [string]$ed.scannedGtin14 } else { "" }
        'ErrorDrugBatch'      = if ($ed) { [string]$ed.batchNumber } else { "" }
        'ErrorDrugExpiry'     = if ($ed) { [string]$ed.expiryDate } else { "" }
        'ErrorDrugIsExpired'  = if ($ed -and $ed.isExpiredAtScan) { "Yes" } else { "No" }
        'ErrorDrugIsNTI'      = if ($ed -and $ed.isNti) { "Yes" } else { "No" }
        
        'IntendedDrugName'    = if ($id) { [string]$id.displayName } else { "" }
        'IntendedDrugCMM'     = if ($id) { [string]$id.cmmCode } else { "" }
        'IntendedDrugVTMID'   = if ($id) { [string]$id.vtmId } else { "" }
        'IntendedDrugVMPID'   = if ($id) { [string]$id.vmpId } else { "" }
        'IntendedDrugAMPID'   = if ($id) { [string]$id.ampId } else { "" }
        'IntendedDrugIsNTI'   = if ($id -and $id.isNti) { "Yes" } else { "No" }
        
        'QuantityDispensed'   = [string]$ev.quantityDispensed
        'QuantityPrescribed'  = [string]$ev.quantityPrescribed
        'ContributingFactors' = $factors
        'LearningNote'        = [string]$ev.learningNote
    }
}

function Get-ColLetter([int]$i) {
    $dividend = $i + 1; $colName = ''
    while ($dividend -gt 0) {
        $modulo = ($dividend - 1) % 26
        $colName = [char](65 + $modulo) + $colName
        $dividend = [int](($dividend - $modulo) / 26)
    }
    return $colName
}

function Esc([string]$s) {
    if (-not $s) { return '' }
    return [System.Security.SecurityElement]::Escape(([regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')))
}

# 5. Write OpenXML Excel (.xlsx) with 100% String Preservation (inlineStr)
Write-Host "Compiling OpenXML Master Excel Workbook..." -ForegroundColor Yellow
$tempDir = Join-Path $env:TEMP ([Guid]::NewGuid().ToString())
$xlDir = Join-Path $tempDir 'xl'; $wsDir = Join-Path $xlDir 'worksheets'
$relsDir = Join-Path $tempDir '_rels'; $xlRelsDir = Join-Path $xlDir '_rels'
foreach ($d in @($wsDir, $relsDir, $xlRelsDir)) { [void](New-Item -ItemType Directory -Path $d -Force) }
$hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n"

[System.IO.File]::WriteAllText((Join-Path $tempDir '[Content_Types].xml'), $hdr +
'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>', $u8NoBom)

[System.IO.File]::WriteAllText((Join-Path $relsDir '.rels'), $hdr +
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', $u8NoBom)

[System.IO.File]::WriteAllText((Join-Path $xlDir 'workbook.xml'), $hdr +
'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Near Miss Master Log" sheetId="1" r:id="rId1"/></sheets></workbook>', $u8NoBom)

[System.IO.File]::WriteAllText((Join-Path $xlRelsDir 'workbook.xml.rels'), $hdr +
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>', $u8NoBom)

$wsPath = Join-Path $wsDir 'sheet1.xml'
$sw = [System.IO.StreamWriter]::new($wsPath, $false, $u8NoBom)
$sw.WriteLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
$sw.WriteLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>')

# Header Row
$sw.Write('<row r="1">')
for ($c = 0; $c -lt $cols.Count; $c++) {
    $sw.Write("<c r=`"$(Get-ColLetter $c)1`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(Esc $cols[$c])</t></is></c>")
}
$sw.WriteLine('</row>')

# Data Rows
$rowNum = 2
foreach ($ev in $events) {
    $rowVals = Get-EventRowValues $ev
    $sw.Write("<row r=`"$rowNum`">")
    for ($c = 0; $c -lt $cols.Count; $c++) {
        $colKey = $cols[$c]
        $val = $rowVals[$colKey]
        if (-not [string]::IsNullOrEmpty($val)) {
            $sw.Write("<c r=`"$(Get-ColLetter $c)$rowNum`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(Esc $val)</t></is></c>")
        }
    }
    $sw.WriteLine('</row>')
    $rowNum++
}

$sw.WriteLine('</sheetData></worksheet>')
$sw.Close()

if (Test-Path $OutXlsx) { Remove-Item -Path $OutXlsx -Force }
$zipArchive = [System.IO.Compression.ZipFile]::Open($OutXlsx, [System.IO.Compression.ZipArchiveMode]::Create)
$allFiles = Get-ChildItem -Path $tempDir -Recurse -File
foreach ($f in $allFiles) {
    $relPath = $f.FullName.Substring($tempDir.Length).TrimStart('\', '/')
    $zipEntryName = $relPath -replace '\\', '/'
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zipArchive, $f.FullName, $zipEntryName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
}
$zipArchive.Dispose()
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "  Written: $OutXlsx ($([math]::Round((Get-Item -LiteralPath $OutXlsx).Length / 1KB, 1)) KB)" -ForegroundColor Green

# 6. Write CSV (UTF-8 with BOM)
Write-Host "Writing CSV Export (UTF-8 BOM)..." -ForegroundColor Yellow
$csvWriter = [System.IO.StreamWriter]::new($OutCsv, $false, $u8WithBom)

# CSV Header
$csvWriter.WriteLine(($cols | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join ',')

# CSV Data Rows
foreach ($ev in $events) {
    $rowVals = Get-EventRowValues $ev
    $rowFields = @()
    foreach ($c in $cols) {
        $v = $rowVals[$c]
        $safe = if ($v) { '"' + ($v -replace '"', '""') + '"' } else { '""' }
        $rowFields += $safe
    }
    $csvWriter.WriteLine(($rowFields -join ','))
}
$csvWriter.Close()
Write-Host "  Written: $OutCsv ($([math]::Round((Get-Item -LiteralPath $OutCsv).Length / 1KB, 1)) KB)" -ForegroundColor Green

# 7. Write Shared Data JS for Analytics Dashboard
Write-Host "Writing nearmiss_shared_data.js for Browser Analytics..." -ForegroundColor Yellow
$jsonFull = if ($events.Count -gt 0) { ConvertTo-Json $events -Depth 10 } else { "[]" }
$jsContent = "/* UHBFT Pharmacy Near Miss Master Shared Data - Auto-generated */`nwindow.NEARMISS_SHARED_DATA = $jsonFull;`n"
[System.IO.File]::WriteAllText($OutJs, $jsContent, $u8NoBom)
Write-Host "  Written: $OutJs ($([math]::Round((Get-Item -LiteralPath $OutJs).Length / 1KB, 1)) KB)" -ForegroundColor Green
Write-Host "  Written: $OutJs ($([math]::Round((Get-Item $OutJs).Length / 1KB, 1)) KB)" -ForegroundColor Green

# 8. Compute Factual Governance Telemetry
Write-Host "Computing Governance & Learning Telemetry..." -ForegroundColor Yellow
$bySite = @{}
$byCategory = @{}
$byType = @{ "Dispensing" = 0; "Clinical" = 0 }
$ntiCount = 0
$expiredCount = 0
$lasaPairs = @{}

foreach ($ev in $events) {
    $site = if ($ev.site) { $ev.site } else { "Unknown" }
    if (-not $bySite.ContainsKey($site)) { $bySite[$site] = 0 }
    $bySite[$site]++

    $type = if ($ev.nearMissType) { $ev.nearMissType } else { "Dispensing" }
    if ($byType.ContainsKey($type)) { $byType[$type]++ }

    if ($ev.errorDrug) {
        if ($ev.errorDrug.isNti) { $ntiCount++ }
        if ($ev.errorDrug.isExpiredAtScan) { $expiredCount++ }

        if ($ev.intendedDrug -and $ev.intendedDrug.displayName) {
            $pairKey = "$($ev.errorDrug.displayName) -> $($ev.intendedDrug.displayName)"
            if (-not $lasaPairs.ContainsKey($pairKey)) { $lasaPairs[$pairKey] = 0 }
            $lasaPairs[$pairKey]++
        }
    }

    if ($ev.errorCategories) {
        foreach ($cat in $ev.errorCategories) {
            if (-not $byCategory.ContainsKey($cat)) { $byCategory[$cat] = 0 }
            $byCategory[$cat]++
        }
    }
}

# Top 10 LASA Pairs
$topLasa = @()
$sortedLasa = $lasaPairs.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10
foreach ($p in $sortedLasa) {
    $topLasa += @{ pair = $p.Key; count = $p.Value }
}

$telemetry = @{
    lastConsolidatedAt = (Get-Date).ToString("o")
    totalNearMisses    = $events.Count
    typeBreakdown      = $byType
    siteBreakdown      = $bySite
    categoryBreakdown  = $byCategory
    ntiIncidentCount   = $ntiCount
    expiredScansCount  = $expiredCount
    topLasaPairs       = $topLasa
}

$telemetryJson = ConvertTo-Json $telemetry -Depth 10
[System.IO.File]::WriteAllText($OutTelemetry, $telemetryJson, $u8NoBom)
Write-Host "  Written: $OutTelemetry" -ForegroundColor Green

Write-Host ""
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " Consolidation Complete: $($events.Count) Near Miss Event(s) Synchronized      " -ForegroundColor Green
Write-Host "================================================================================" -ForegroundColor Cyan
