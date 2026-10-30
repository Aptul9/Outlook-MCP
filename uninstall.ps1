[CmdletBinding()]
param(
    [string]$InstallDir = "$env:LOCALAPPDATA\Programs\OutlookMCP"
)

Write-Host "=== Uninstalling Outlook MCP ===" -ForegroundColor Cyan

# Remove Claude Code registration
$claudeCmd = Get-Command claude -ErrorAction SilentlyContinue
if ($claudeCmd) {
    Write-Host "Removing from Claude Code..."
    try {
        & claude mcp remove --scope user outlook
    } catch {}
}

# Remove Claude Desktop registration
$desktopCfgPath = "$env:APPDATA\Claude\claude_desktop_config.json"
if (Test-Path $desktopCfgPath) {
    try {
        $cfg = Get-Content $desktopCfgPath -Raw | ConvertFrom-Json -AsHashtable
        if ($cfg.ContainsKey("mcpServers") -and $cfg["mcpServers"].ContainsKey("outlook")) {
            $cfg["mcpServers"].Remove("outlook")
            $cfg | ConvertTo-Json -Depth 10 | Set-Content $desktopCfgPath -Encoding utf8
            Write-Host "Removed from Claude Desktop configuration."
        }
    } catch {}
}

# Remove files and caches
if (Test-Path $InstallDir) {
    Remove-Item -Path $InstallDir -Recurse -Force
    Write-Host "Removed $InstallDir"
}

$cacheDir = "$env:LOCALAPPDATA\OutlookMCP"
if (Test-Path $cacheDir) {
    Remove-Item -Path $cacheDir -Recurse -Force
    Write-Host "Removed $cacheDir"
}

Write-Host "Outlook MCP uninstalled." -ForegroundColor Green
