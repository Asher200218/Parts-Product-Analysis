# Parts Analysis - local launcher (Windows, pure PowerShell, no installs needed).
# Serves the folder this script/exe lives in and opens the default browser.
# Works both as: powershell -File launch.ps1   AND   compiled Parts Analysis.exe
# Close this window to stop the server.

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Resolve our own folder (script dir when run as .ps1, exe dir when compiled by ps2exe)
if ($PSScriptRoot) {
    $base = $PSScriptRoot
} else {
    $base = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}

# Pick a free port
function Test-FreePort([int]$p) {
    try {
        $l = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $p)
        $l.Start(); $l.Stop(); return $true
    } catch { return $false }
}
$port = 4173
foreach ($p in 4173, 4174, 4175, 4180, 8099, 8123) {
    if (Test-FreePort $p) { $port = $p; break }
}
$prefix = "http://localhost:$port/"

# Bind to localhost only (no admin needed; nothing exposed to the network)
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)
try {
    $listener.Start()
} catch {
    Write-Host "无法启动本地服务器 / Could not start local server: $($_.Exception.Message)"
    Read-Host "按回车关闭 / Press Enter to close"
    exit 1
}

$mime = @{
    '.html' = 'text/html; charset=utf-8'; '.htm' = 'text/html; charset=utf-8'
    '.js'   = 'text/javascript; charset=utf-8'; '.mjs' = 'text/javascript; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'; '.json' = 'application/json; charset=utf-8'
    '.svg'  = 'image/svg+xml'; '.ico' = 'image/x-icon'; '.png' = 'image/png'
    '.jpg'  = 'image/jpeg'; '.jpeg' = 'image/jpeg'; '.gif' = 'image/gif'
    '.woff' = 'font/woff'; '.woff2' = 'font/woff2'; '.map' = 'application/json'
    '.xlsx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
}

# Open the browser once the server is listening
Start-Process $prefix | Out-Null

Write-Host "------------------------------------------------------------"
Write-Host "  Parts Analysis is running:  $prefix"
Write-Host "  零件分析录入网页正在运行。关闭此窗口即可停止。"
Write-Host "  Close this window to stop the server."
Write-Host "------------------------------------------------------------"

$baseFull = [System.IO.Path]::GetFullPath($base)
while ($listener.IsListening) {
    try {
        $ctx = $listener.GetContext()
    } catch { break }
    $req = $ctx.Request
    $res = $ctx.Response
    try {
        $rel = [System.Uri]::UnescapeDataString($req.Url.AbsolutePath).TrimStart('/')
        if ([string]::IsNullOrWhiteSpace($rel)) { $rel = 'index.html' }
        $rel = $rel -replace '/', '\'
        $full = [System.IO.Path]::GetFullPath((Join-Path $base $rel))

        if (-not $full.StartsWith($baseFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            $res.StatusCode = 403
        } elseif (Test-Path $full -PathType Leaf) {
            $bytes = [System.IO.File]::ReadAllBytes($full)
            $ext = [System.IO.Path]::GetExtension($full).ToLower()
            $ct = $mime[$ext]; if (-not $ct) { $ct = 'application/octet-stream' }
            $res.ContentType = $ct
            $res.ContentLength64 = $bytes.Length
            $res.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $res.StatusCode = 404
            $msg = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found: $rel")
            $res.OutputStream.Write($msg, 0, $msg.Length)
        }
    } catch {
        try { $res.StatusCode = 500 } catch {}
    } finally {
        try { $res.OutputStream.Close() } catch {}
    }
}
