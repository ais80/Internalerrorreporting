# ==============================================================================
# nearmiss-server.ps1
# UHBFT Secondary Care Pharmacy Near Miss Reporting & Learning System
# Lightweight Local/Network Bridge Micro-Server (Phase 3)
# ==============================================================================
[CmdletBinding()]
param(
    [int]$Port = 8088,
    [string]$DataDir = ""
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = $PSScriptRoot }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }

if (-not $DataDir) { $DataDir = Join-Path $scriptDir "data" }
$DataDir = [System.IO.Path]::GetFullPath($DataDir)
$EventsDir = Join-Path $DataDir "events"
$ImportQueueDir = Join-Path $DataDir "_import_queue"

# Ensure directories exist
foreach ($d in @($DataDir, $EventsDir, $ImportQueueDir, (Join-Path $EventsDir "_undone"))) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
}

$u8 = [System.Text.UTF8Encoding]::new($false)

# Initialize HttpListener
$listener = New-Object System.Net.HttpListener
$prefixes = [System.Collections.Generic.List[string]]::new()
$prefixes.Add("http://localhost:$Port/")
$prefixes.Add("http://127.0.0.1:$Port/")

# Auto-bind local network IPs for mobile phones / tablets on the same Wi-Fi
try {
    $hostName = [System.Net.Dns]::GetHostName()
    $prefixes.Add("http://${hostName}:$Port/")
    $localIps = [System.Net.Dns]::GetHostAddresses($hostName) | Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork }
    foreach ($ip in $localIps) {
        $prefixes.Add("http://${ip}:$Port/")
    }
} catch {
    Write-Warning "Could not enumerate local IP addresses: $_"
}

foreach ($p in $prefixes) {
    try {
        $listener.Prefixes.Add($p)
    } catch {
        Write-Warning "Could not bind prefix ${p}: $_"
    }
}

try {
    $listener.Start()
} catch {
    Write-Error "Failed to start HttpListener on port $Port. Is another instance already running?"
    exit 1
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " UHBFT Pharmacy Near Miss Bridge Server Running                                " -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " Local URL   : http://localhost:$Port/Near-Miss-Logger.html" -ForegroundColor Green
foreach ($p in $listener.Prefixes) {
    if ($p -notlike "*localhost*" -and $p -notlike "*127.0.0.1*") {
        Write-Host " Mobile URL  : $($p)Near-Miss-Logger.html" -ForegroundColor Green
    }
}
Write-Host " Data Dir    : $DataDir"
Write-Host " Events Dir  : $EventsDir"
Write-Host " Press Ctrl+C in this console to stop the server." -ForegroundColor Yellow
Write-Host "================================================================================" -ForegroundColor Cyan

function Send-Response {
    param(
        [System.Net.HttpListenerResponse]$Response,
        [int]$StatusCode = 200,
        [string]$ContentType = "application/json; charset=utf-8",
        [byte[]]$Bytes = $null
    )

    $Response.StatusCode = $StatusCode
    $Response.ContentType = $ContentType
    $Response.AddHeader("Access-Control-Allow-Origin", "*")
    $Response.AddHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
    $Response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Requested-With")
    
    if ($Bytes -ne $null) {
        $Response.ContentLength64 = $Bytes.Length
        $Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
    } else {
        $Response.ContentLength64 = 0
    }
    $Response.OutputStream.Close()
}

function Get-MimeType([string]$ext) {
    switch ($ext.ToLower()) {
        ".html" { return "text/html; charset=utf-8" }
        ".htm"  { return "text/html; charset=utf-8" }
        ".js"   { return "application/javascript; charset=utf-8" }
        ".json" { return "application/json; charset=utf-8" }
        ".css"  { return "text/css; charset=utf-8" }
        ".png"  { return "image/png" }
        ".jpg"  { return "image/jpeg" }
        ".ico"  { return "image/x-icon" }
        ".svg"  { return "image/svg+xml" }
        default { return "application/octet-stream" }
    }
}

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        $httpMethod = $request.HttpMethod.ToUpper()
        $rawPath = $request.Url.AbsolutePath

        # Handle CORS Preflight
        if ($httpMethod -eq "OPTIONS") {
            Send-Response -Response $response -StatusCode 200
            continue
        }

        # API: POST /api/nearmiss
        if ($httpMethod -eq "POST" -and $rawPath -eq "/api/nearmiss") {
            try {
                $reader = New-Object System.IO.StreamReader($request.InputStream, $u8)
                $body = $reader.ReadToEnd()
                $reader.Close()

                if ([string]::IsNullOrWhiteSpace($body)) {
                    $errBytes = $u8.GetBytes('{"error": "Empty body"}')
                    Send-Response -Response $response -StatusCode 400 -Bytes $errBytes
                    continue
                }

                $json = ConvertFrom-Json $body -ErrorAction Stop

                # Extract or generate eventId
                $eventId = $json.eventId
                if ([string]::IsNullOrWhiteSpace($eventId)) {
                    $ts = Get-Date -Format "yyyyMMdd-HHmmss"
                    $site = if ($json.site) { ($json.site -replace '\s+', '').Substring(0, [math]::Min(4, $json.site.Length)).ToUpper() } else { "BENCH" }
                    $rand = -join ((65..90) + (48..57) | Get-Random -Count 4 | ForEach-Object { [char]$_ })
                    $eventId = "NM-$ts-$site-$rand"
                    $json | Add-Member -MemberType NoteProperty -Name "eventId" -Value $eventId -Force
                }

                # Partition by Month YYYY-MM
                $monthFolder = (Get-Date).ToString("yyyy-MM")
                $targetDir = Join-Path $EventsDir $monthFolder
                if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }

                $targetFile = Join-Path $targetDir "$eventId.json"
                $formattedJson = ConvertTo-Json $json -Depth 10
                [System.IO.File]::WriteAllText($targetFile, $formattedJson, $u8)

                $type = if ($json.nearMissType) { $json.nearMissType } else { "Event" }
                $site = if ($json.site) { $json.site } else { "Site" }
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [POST 200] Saved Near Miss $eventId ($site - $type)" -ForegroundColor Green

                $resBytes = $u8.GetBytes((ConvertTo-Json @{ success = $true; eventId = $eventId; path = $targetFile }))
                Send-Response -Response $response -StatusCode 200 -Bytes $resBytes
            } catch {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [POST 500] Error: $_" -ForegroundColor Red
                $errBytes = $u8.GetBytes((ConvertTo-Json @{ error = $_.Exception.Message }))
                Send-Response -Response $response -StatusCode 500 -Bytes $errBytes
            }
            continue
        }

        # API: DELETE /api/nearmiss/{eventId} (5-second Undo)
        if ($httpMethod -eq "DELETE" -and $rawPath.StartsWith("/api/nearmiss/")) {
            $eventId = $rawPath.Substring("/api/nearmiss/".Length).Trim()
            $found = $false
            if ($eventId) {
                # Look for event file in data\events\*\<eventId>.json
                $matches = Get-ChildItem -Path $EventsDir -Filter "$eventId.json" -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.DirectoryName -notlike "*_undone*" }
                if ($matches.Count -gt 0) {
                    $undoneDir = Join-Path $EventsDir "_undone"
                    foreach ($m in $matches) {
                        Move-Item -Path $m.FullName -Destination (Join-Path $undoneDir $m.Name) -Force
                        $found = $true
                    }
                }
            }

            if ($found) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [DELETE 200] Undone Near Miss $eventId (moved to _undone)" -ForegroundColor Yellow
                $resBytes = $u8.GetBytes((ConvertTo-Json @{ success = $true; undone = $true; eventId = $eventId }))
                Send-Response -Response $response -StatusCode 200 -Bytes $resBytes
            } else {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [DELETE 404] Event $eventId not found" -ForegroundColor DarkGray
                $resBytes = $u8.GetBytes((ConvertTo-Json @{ success = $false; message = "Event not found" }))
                Send-Response -Response $response -StatusCode 404 -Bytes $resBytes
            }
            continue
        }

        # API: GET /api/status
        if ($httpMethod -eq "GET" -and $rawPath -eq "/api/status") {
            $totalCount = (Get-ChildItem -Path $EventsDir -Filter "NM-*.json" -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.DirectoryName -notlike "*_undone*" }).Count
            $statusObj = @{
                status = "running"
                timestamp = (Get-Date).ToString("o")
                port = $Port
                totalEventsRecorded = $totalCount
            }
            $resBytes = $u8.GetBytes((ConvertTo-Json $statusObj))
            Send-Response -Response $response -StatusCode 200 -Bytes $resBytes
            continue
        }

        # Static File Serving
        $localRel = $rawPath.TrimStart('/')
        if ([string]::IsNullOrWhiteSpace($localRel)) {
            $localRel = "Near-Miss-Logger.html"
        }

        $localFilePath = Join-Path $scriptDir $localRel
        if (-not (Test-Path $localFilePath -PathType Leaf)) {
            # Try parent directory if running from NearMiss subfolder
            $parentFilePath = Join-Path (Split-Path -Parent $scriptDir) $localRel
            if (Test-Path $parentFilePath -PathType Leaf) {
                $localFilePath = $parentFilePath
            }
        }

        if (Test-Path $localFilePath -PathType Leaf) {
            $ext = [System.IO.Path]::GetExtension($localFilePath)
            $mime = Get-MimeType $ext
            $fileBytes = [System.IO.File]::ReadAllBytes($localFilePath)
            Send-Response -Response $response -StatusCode 200 -ContentType $mime -Bytes $fileBytes
        } else {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [GET 404] Not Found: $rawPath" -ForegroundColor DarkGray
            $errBytes = $u8.GetBytes('{"error": "File Not Found"}')
            Send-Response -Response $response -StatusCode 404 -Bytes $errBytes
        }
    }
} finally {
    if ($listener.IsListening) {
        $listener.Stop()
    }
    $listener.Close()
    Write-Host "`nServer stopped cleanly." -ForegroundColor Yellow
}

