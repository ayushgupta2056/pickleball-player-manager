$ErrorActionPreference = "Stop"

$projectPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$secretsDirectory = Join-Path $projectPath ".streamlit"
$secretsPath = Join-Path $secretsDirectory "secrets.toml"

Set-Location -LiteralPath $projectPath

$needsToken = -not (Test-Path -LiteralPath $secretsPath)
if (-not $needsToken) {
    $currentSecrets = Get-Content -Raw -LiteralPath $secretsPath
    $needsToken = $currentSecrets -match "PASTE_YOUR_DUPR_TOKEN_HERE|YOUR_DUPR_TOKEN_HERE"
}

if ($needsToken) {
    Write-Host ""
    Write-Host "Pickleball Player Manager - First-time setup" -ForegroundColor Cyan
    Write-Host "Paste your DUPR token below. It will remain hidden while you type." -ForegroundColor Yellow
    $secureToken = Read-Host "DUPR token" -AsSecureString
    $tokenPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken)
    try {
        $plainToken = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($tokenPointer)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($tokenPointer)
    }

    if ([string]::IsNullOrWhiteSpace($plainToken)) {
        throw "No token was entered. Run this launcher again and enter the DUPR token."
    }

    New-Item -ItemType Directory -Path $secretsDirectory -Force | Out-Null
    $tomlToken = $plainToken.Replace('\', '\\').Replace('"', '\"')
    $secretsContent = 'DUPR_TOKEN = "' + $tomlToken + '"'
    $utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($secretsPath, $secretsContent, $utf8WithoutBom)
    Remove-Variable plainToken
    Write-Host "Token saved securely in the Git-ignored local secrets file." -ForegroundColor Green
}

Write-Host "Starting Pickleball Player Manager..." -ForegroundColor Green
python -m streamlit run Main.py
