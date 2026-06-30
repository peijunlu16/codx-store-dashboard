$ErrorActionPreference = "Stop"

$noOpen = $env:DASHBOARD_NO_OPEN -eq "1"

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = 8765
$htmlName = "门店运营看板_成品.html"
$listener = $null
for ($tryPort = 8765; $tryPort -le 8785; $tryPort++) {
  try {
    $candidate = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Parse("127.0.0.1"), $tryPort)
    $candidate.Start()
    $listener = $candidate
    $port = $tryPort
    break
  } catch {
    if ($tryPort -eq 8785) { throw }
  }
}
$url = "http://127.0.0.1:$port/$([uri]::EscapeDataString($htmlName))"

function Get-ContentType([string]$path) {
  switch ([System.IO.Path]::GetExtension($path).ToLowerInvariant()) {
    ".html" { "text/html; charset=utf-8" }
    ".htm"  { "text/html; charset=utf-8" }
    ".js"   { "application/javascript; charset=utf-8" }
    ".css"  { "text/css; charset=utf-8" }
    ".json" { "application/json; charset=utf-8" }
    ".xlsx" { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
    ".xls"  { "application/vnd.ms-excel" }
    ".csv"  { "text/csv; charset=utf-8" }
    ".ico"  { "image/x-icon" }
    ".png"  { "image/png" }
    ".jpg"  { "image/jpeg" }
    ".jpeg" { "image/jpeg" }
    ".svg"  { "image/svg+xml" }
    default { "application/octet-stream" }
  }
}

function Send-Response($stream, [int]$status, [string]$statusText, [byte[]]$body, [string]$contentType, [bool]$headOnly, [string]$lastModified = "", [long]$contentLength = -1) {
  $responseLength = if ($contentLength -ge 0) { $contentLength } else { $body.Length }
  $header = "HTTP/1.1 $status $statusText`r`n" +
    "Content-Type: $contentType`r`n" +
    "Content-Length: $responseLength`r`n" +
    $(if ($lastModified) { "Last-Modified: $lastModified`r`n" } else { "" }) +
    "Cache-Control: no-store`r`n" +
    "Access-Control-Allow-Origin: *`r`n" +
    "Connection: close`r`n`r`n"
  $headerBytes = [System.Text.Encoding]::UTF8.GetBytes($header)
  $stream.Write($headerBytes, 0, $headerBytes.Length)
  if (-not $headOnly -and $body.Length -gt 0) {
    $stream.Write($body, 0, $body.Length)
  }
}

function Read-SharedFileBytes([string]$path) {
  $lastError = $null
  for ($attempt = 1; $attempt -le 5; $attempt++) {
    try {
      $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
      $file = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
      try {
        $memory = [System.IO.MemoryStream]::new()
        $file.CopyTo($memory)
        return $memory.ToArray()
      } finally {
        if ($memory) { $memory.Dispose() }
        $file.Dispose()
      }
    } catch {
      $lastError = $_
      Start-Sleep -Milliseconds 500
    }
  }
  throw $lastError
}

function Resolve-RequestPath([string]$rawPath) {
  if ([string]::IsNullOrWhiteSpace($rawPath) -or $rawPath -eq "/") {
    $rawPath = "/" + [uri]::EscapeDataString($htmlName)
  }
  $pathOnly = $rawPath.Split("?")[0]
  $rawName = $pathOnly.TrimStart("/")
  $decoded = [uri]::UnescapeDataString($pathOnly).TrimStart("/")
  $decoded = $decoded -replace "/", "\"
  $full = [System.IO.Path]::GetFullPath((Join-Path $here $decoded))
  $root = [System.IO.Path]::GetFullPath($here)
  if ($full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $full -PathType Leaf)) {
    return $full
  }

  # Fallback for Chinese filenames under older Windows PowerShell encodings.
  $decodedLeaf = [System.IO.Path]::GetFileName($decoded)
  $matched = Get-ChildItem -LiteralPath $here -File | Where-Object {
    $_.Name -eq $decodedLeaf -or [uri]::EscapeDataString($_.Name) -eq $rawName
  } | Select-Object -First 1
  if ($matched) { return $matched.FullName }
  return $null
}

Write-Host "正在启动门店运营看板自动更新服务..."
Write-Host "目录: $here"
Write-Host "地址: $url"
Write-Host "保持此窗口打开；修改同目录的销售/库存 Excel 后，看板会在约 30 秒内自动刷新。"
Write-Host "关闭此窗口即可停止服务。"

if (-not $noOpen) {
  Start-Process $url
}

try {
  while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
      $stream = $client.GetStream()
      $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::ASCII, $false, 1024, $true)
      $requestLine = $reader.ReadLine()
      if (-not $requestLine) { continue }

      while (($line = $reader.ReadLine()) -ne $null -and $line -ne "") {}

      $parts = $requestLine.Split(" ")
      $method = $parts[0]
      $requestPath = if ($parts.Length -gt 1) { $parts[1] } else { "/" }
      $headOnly = $method -eq "HEAD"

      if ($method -ne "GET" -and $method -ne "HEAD") {
        Send-Response $stream 405 "Method Not Allowed" ([System.Text.Encoding]::UTF8.GetBytes("Method Not Allowed")) "text/plain; charset=utf-8" $headOnly
        continue
      }

      $filePath = Resolve-RequestPath $requestPath
      if (-not $filePath -or -not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
        Send-Response $stream 404 "Not Found" ([System.Text.Encoding]::UTF8.GetBytes("File Not Found")) "text/plain; charset=utf-8" $headOnly
        continue
      }

      $lastModified = [System.IO.File]::GetLastWriteTimeUtc($filePath).ToString("R", [System.Globalization.CultureInfo]::InvariantCulture)
      $fileLength = (Get-Item -LiteralPath $filePath).Length
      $bytes = if ($headOnly) { [byte[]]::new(0) } else { Read-SharedFileBytes $filePath }
      Send-Response $stream 200 "OK" $bytes (Get-ContentType $filePath) $headOnly $lastModified $fileLength
    } catch {
      try {
        Send-Response $stream 500 "Server Error" ([System.Text.Encoding]::UTF8.GetBytes($_.Exception.Message)) "text/plain; charset=utf-8" $false
      } catch {}
    } finally {
      $client.Close()
    }
  }
} finally {
  $listener.Stop()
}
