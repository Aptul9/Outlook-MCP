[CmdletBinding()]
param(
    [string]$InstallDir = "$env:LOCALAPPDATA\Programs\OutlookMCP",
    [string]$Python = "",
    [switch]$ClaudeDesktop,
    [switch]$NoClaudeCode
)

$ErrorActionPreference = "Stop"

Write-Host "=== Installing Outlook MCP ===" -ForegroundColor Cyan

# 1. Locate Python
if (-not $Python) {
    $pyCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pyCmd) {
        $Python = $pyCmd.Source
    } else {
        $pyCmd = Get-Command py -ErrorAction SilentlyContinue
        if ($pyCmd) {
            $Python = $pyCmd.Source
        }
    }
}

if (-not $Python -or -not (Test-Path $Python)) {
    Write-Error "Python 3.8+ executable not found. Pass -Python <path_to_python.exe>"
}
Write-Host "Using Python: $Python" -ForegroundColor Green

# 2. Setup install directory
Write-Host "Target install directory: $InstallDir"
if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
}

$srcDir = $PSScriptRoot
Copy-Item (Join-Path $srcDir "outlook_mcp.py") $InstallDir -Force
Copy-Item (Join-Path $srcDir "outlook_api.py") $InstallDir -Force
Copy-Item (Join-Path $srcDir "outlook_tokens.py") $InstallDir -Force

$entryPoint = Join-Path $InstallDir "outlook_mcp.py"
Write-Host "Copied server files to $InstallDir" -ForegroundColor Green

# 3. Verify server runs & detects accounts
Write-Host "Verifying server and token harvest..." -ForegroundColor Cyan
& $Python $entryPoint --call list_accounts
if ($LASTEXITCODE -ne 0) {
    Write-Warning "Initial account check returned non-zero exit code. Make sure New Outlook is logged in."
}

# 4. Register with Claude Code CLI if available
if (-not $NoClaudeCode) {
    $claudeCmd = Get-Command claude -ErrorAction SilentlyContinue
    if ($claudeCmd) {
        Write-Host "Registering with Claude Code (user scope)..." -ForegroundColor Cyan
        try {
            & claude mcp remove --scope user outlook 2>$null
        } catch {}
        & claude mcp add --scope user outlook -- $Python $entryPoint
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Registered in Claude Code successfully." -ForegroundColor Green
        } else {
            Write-Warning "Could not register automatically with Claude Code CLI."
        }
    }
}

# 5. Register with Claude Desktop if requested
if ($ClaudeDesktop) {
    $desktopCfgPath = "$env:APPDATA\Claude\claude_desktop_config.json"
    Write-Host "Configuring Claude Desktop at $desktopCfgPath..." -ForegroundColor Cyan
    $desktopDir = Split-Path $desktopCfgPath
    if (-not (Test-Path $desktopDir)) {
        New-Item -ItemType Directory -Path $desktopDir -Force | Out-Null
    }
    $cfg = if (Test-Path $desktopCfgPath) {
        Get-Content $desktopCfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } else {
        [pscustomobject]@{}
    }
    if (-not $cfg.PSObject.Properties["mcpServers"]) {
        $cfg | Add-Member mcpServers ([pscustomobject]@{})
    }
    $cfg.mcpServers | Add-Member outlook ([pscustomobject]@{
        command = $Python
        args = @($entryPoint)
    }) -Force
    [IO.File]::WriteAllText($desktopCfgPath, ($cfg | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
    Write-Host "Updated Claude Desktop config successfully." -ForegroundColor Green
}

Write-Host "`nOutlook MCP installation complete!" -ForegroundColor Green
Write-Host "To test manually:" -ForegroundColor Yellow
Write-Host "  $Python `"$entryPoint`" --call list_emails" -ForegroundColor Gray
