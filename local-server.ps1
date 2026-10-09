[CmdletBinding()]
param([switch]$NoBrowser)

$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
$serverPort = 8787
$serverUrl = "http://127.0.0.1:$serverPort/"
$utf8 = New-Object Text.UTF8Encoding($false)
$rootHasher = [Security.Cryptography.SHA256]::Create()
try {
    $rootDigest = [BitConverter]::ToString($rootHasher.ComputeHash($utf8.GetBytes($projectRoot.ToLowerInvariant()))).Replace('-', '').ToLowerInvariant()
} finally {
    $rootHasher.Dispose()
}
$serverIdentity = "NijinanaSimulator/1:$rootDigest"
$publicFiles = @{
    '/index.html' = 'text/html; charset=utf-8'
    '/deck/index.html' = 'text/html; charset=utf-8'
    '/app-navigation.js' = 'text/javascript; charset=utf-8'
    '/app.js' = 'text/javascript; charset=utf-8'
    '/deck-storage.js' = 'text/javascript; charset=utf-8'
    '/json-export.js' = 'text/javascript; charset=utf-8'
    '/deck-txt.js' = 'text/javascript; charset=utf-8'
    '/deck-manager.js' = 'text/javascript; charset=utf-8'
    '/app.css' = 'text/css; charset=utf-8'
    '/deck-manager.css' = 'text/css; charset=utf-8'
    '/default-cards.csv' = 'text/csv; charset=utf-8'
}

function Open-SimulatorBrowser {
    if (-not $NoBrowser) {
        try { Start-Process -FilePath $serverUrl | Out-Null }
        catch { Write-Host "ブラウザを開けませんでした。次のURLを開いてください: $serverUrl" }
    }
}

function Test-ExistingSimulator {
    try {
        $request = [Net.HttpWebRequest]::Create("${serverUrl}__nijinana_server")
        $request.Method = 'GET'
        $request.Proxy = $null
        $request.AllowAutoRedirect = $false
        $request.Timeout = 3000
        $request.ReadWriteTimeout = 3000
        $response = $request.GetResponse()
        try {
            return ($response.StatusCode -eq [Net.HttpStatusCode]::OK -and $response.Headers['X-Nijinana-Server'] -ceq $serverIdentity)
        } finally { $response.Dispose() }
    } catch { return $false }
}

function Send-HttpResponse {
    param($Stream, [int]$Status, [string]$Reason, [string]$ContentType, [byte[]]$Body, [bool]$HeadOnly = $false, [string]$ExtraHeaders = '')
    $header = "HTTP/1.1 $Status $Reason`r`nContent-Type: $ContentType`r`nContent-Length: $($Body.Length)`r`nConnection: close`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`nReferrer-Policy: no-referrer`r`n$ExtraHeaders`r`n"
    $headerBytes = [Text.Encoding]::ASCII.GetBytes($header)
    $Stream.Write($headerBytes, 0, $headerBytes.Length)
    if (-not $HeadOnly -and $Body.Length -gt 0) { $Stream.Write($Body, 0, $Body.Length) }
    $Stream.Flush()
}

function Send-HttpError {
    param($Stream, [int]$Status, [string]$Reason, [bool]$HeadOnly = $false, [string]$ExtraHeaders = '')
    Send-HttpResponse $Stream $Status $Reason 'text/plain; charset=utf-8' $utf8.GetBytes("$Status $Reason`n") $HeadOnly $ExtraHeaders
}

function Get-PublicFilePath {
    param([string]$RequestPath)
    if (-not $publicFiles.ContainsKey($RequestPath)) { return $null }
    $parts = $RequestPath.TrimStart('/').Split('/')
    $candidate = $projectRoot
    foreach ($part in $parts) {
        $candidate = [IO.Path]::Combine($candidate, $part)
        $entry = Get-Item -LiteralPath $candidate -Force -ErrorAction SilentlyContinue
        if (-not $entry -or ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint)) { return $null }
    }
    $candidate = [IO.Path]::GetFullPath($candidate)
    $prefix = $projectRoot + [IO.Path]::DirectorySeparatorChar
    if (-not $candidate.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -or -not [IO.File]::Exists($candidate)) { return $null }
    return $candidate
}

function Serve-Client {
    param($Client)
    $Client.ReceiveTimeout = 3000
    $Client.SendTimeout = 3000
    $stream = $Client.GetStream()
    $stream.ReadTimeout = 3000
    $stream.WriteTimeout = 3000
    $headerBuffer = New-Object IO.MemoryStream
    try {
        $ending = 0
        $headerClock = [Diagnostics.Stopwatch]::StartNew()
        $requestComplete = $false
        while ($headerBuffer.Length -lt 16384) {
            $remainingMs = 3000 - $headerClock.ElapsedMilliseconds
            if ($remainingMs -le 0) { Send-HttpError $stream 408 'Request Timeout'; return }
            $stream.ReadTimeout = [Math]::Max(1, [int]$remainingMs)
            $nextByte = $stream.ReadByte()
            if ($nextByte -lt 0) { return }
            $headerBuffer.WriteByte([byte]$nextByte)
            $ending = (($ending -shl 8) -bor $nextByte) -band 0xFFFFFFFFL
            if ($ending -eq 0x0D0A0D0A) { $requestComplete = $true; break }
        }
        if (-not $requestComplete) { Send-HttpError $stream 431 'Request Header Fields Too Large'; return }
        $requestText = [Text.Encoding]::ASCII.GetString($headerBuffer.ToArray())
        $lines = $requestText.Split(@("`r`n"), [StringSplitOptions]::None)
        $requestLine = [regex]::Match($lines[0], '^([A-Z]+) ([^ ]{1,2048}) HTTP/1\.[01]$')
        if (-not $requestLine.Success) { Send-HttpError $stream 400 'Bad Request'; return }
        $method = $requestLine.Groups[1].Value
        $headOnly = $method -ceq 'HEAD'
        if ($method -cne 'GET' -and -not $headOnly) { Send-HttpError $stream 405 'Method Not Allowed' $false "Allow: GET, HEAD`r`n"; return }
        $hosts = @($lines | Where-Object { $_ -match '^Host:' })
        if ($hosts.Count -ne 1 -or $hosts[0] -notmatch "^Host:\s*(127\.0\.0\.1|localhost)(:$serverPort)?\s*$") { Send-HttpError $stream 400 'Bad Request' $headOnly; return }
        $rawTarget = $requestLine.Groups[2].Value
        if (-not $rawTarget.StartsWith('/') -or $rawTarget.Contains('#')) { Send-HttpError $stream 400 'Bad Request' $headOnly; return }
        $rawPath = ($rawTarget -split '\?', 2)[0]
        try { $requestPath = [Uri]::UnescapeDataString($rawPath) }
        catch { Send-HttpError $stream 400 'Bad Request' $headOnly; return }
        if ($requestPath -match '[\\\x00-\x20\x7F]' -or $requestPath.Contains('//') -or $requestPath.Contains('%') -or $requestPath -match '(^|/)\.{1,2}(/|$)') { Send-HttpError $stream 400 'Bad Request' $headOnly; return }
        if ($requestPath -ceq '/__nijinana_server') {
            Send-HttpResponse $stream 200 'OK' 'text/plain; charset=utf-8' $utf8.GetBytes('NijinanaSimulator local server') $headOnly "X-Nijinana-Server: $serverIdentity`r`n"
            return
        }
        if ($requestPath -ceq '/') { $requestPath = '/index.html' }
        elseif ($requestPath -ceq '/deck' -or $requestPath -ceq '/deck/') { $requestPath = '/deck/index.html' }
        $filePath = Get-PublicFilePath $requestPath
        if (-not $filePath) { Send-HttpError $stream 404 'Not Found' $headOnly; return }
        $fileInfo = Get-Item -LiteralPath $filePath
        if ($fileInfo.Length -gt 16777216) { Send-HttpError $stream 413 'Content Too Large' $headOnly; return }
        $body = [IO.File]::ReadAllBytes($filePath)
        Send-HttpResponse $stream 200 'OK' $publicFiles[$requestPath] $body $headOnly
    } finally { $headerBuffer.Dispose() }
}

$listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, $serverPort)
try { $listener.Start(128) }
catch {
    if (Test-ExistingSimulator) {
        Write-Host "NijinanaSimulator は既に起動しています: $serverUrl"
        Open-SimulatorBrowser
        exit 0
    }
    Write-Host "起動できませんでした。ポート $serverPort を他のアプリが使用しているか、通信が制限されています。"
    Write-Host 'このウィンドウを閉じ、ポートを使用しているアプリを終了してからもう一度起動してください。'
    exit 1
}

Write-Host "NijinanaSimulator: $serverUrl"
Write-Host 'このウィンドウを開いたまま使用してください。終了するには Ctrl+C を押すか、このウィンドウを閉じてください。'
Open-SimulatorBrowser
try {
    while ($true) {
        if (-not $listener.Pending()) { Start-Sleep -Milliseconds 50; continue }
        $client = $listener.AcceptTcpClient()
        try { Serve-Client $client }
        catch {
            # An incomplete or disconnected request must not stop the local server.
        } finally { $client.Dispose() }
    }
} finally { $listener.Stop() }