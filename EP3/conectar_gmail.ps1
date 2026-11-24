param([string]$Cliente='', [string]$Correo='')
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot

function Find-GmailClientFile {
    Add-Type -AssemblyName System.Windows.Forms
    $Picker = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $Picker.Title = 'Selecciona el JSON de OAuth de Google para Gmail'
        $Picker.Filter = 'Cliente OAuth de Google (*.json)|*.json'
        $Picker.InitialDirectory = Join-Path $env:USERPROFILE 'Downloads'
        $Picker.CheckFileExists = $true
        if ($Picker.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $Picker.FileName
        }
    } finally {
        $Picker.Dispose()
    }
}

try {
    if (-not $Cliente) {
        $Cliente = Join-Path $PSScriptRoot 'gmail-cliente.local.json'
        if (-not (Test-Path -LiteralPath $Cliente -PathType Leaf)) {
            Write-Host 'Falta el cliente OAuth de Gmail. Si ya lo descargaste, selecciona su archivo JSON.'
            $Cliente = Find-GmailClientFile
        }
    }
    if (-not $Cliente -or -not (Test-Path -LiteralPath $Cliente -PathType Leaf)) {
        throw 'Necesitas el JSON de un cliente OAuth de tipo Aplicacion de escritorio, con Gmail API habilitada.'
    }
    $Cliente = (Resolve-Path -LiteralPath $Cliente).Path
    try {
        $Installed = (Get-Content -LiteralPath $Cliente -Raw -Encoding UTF8 | ConvertFrom-Json).installed
    } catch {
        throw 'El archivo elegido no es un JSON valido. Selecciona la descarga de Google Cloud.'
    }
    if (-not $Installed -or $Installed.client_id -notlike '*.apps.googleusercontent.com' -or
        -not $Installed.client_secret -or
        $Installed.auth_uri -notin @('https://accounts.google.com/o/oauth2/auth','https://accounts.google.com/o/oauth2/v2/auth') -or
        $Installed.token_uri -ne 'https://oauth2.googleapis.com/token') {
        throw 'El archivo elegido no corresponde a un cliente OAuth de escritorio descargado de Google.'
    }
    if (-not $Correo) { $Correo = Read-Host 'Correo institucional completo' }
    . (Join-Path $PSScriptRoot 'python_correo.ps1')
    $Python = Find-Ep3Python
    $Runtime = Join-Path $PSScriptRoot '.runtime-correo'
    Push-Location -LiteralPath $Root
    try {
        & $Python -m pip install --disable-pip-version-check --target $Runtime -r (Join-Path $PSScriptRoot 'requirements-correo.txt')
        if ($LASTEXITCODE) { throw 'No se pudieron preparar las dependencias de correo.' }
        & $Python -m auxiliaries.ep3_conectar_gmail $Correo $Cliente
        if ($LASTEXITCODE) { throw 'Gmail no se conecto. Revisa el aviso de Google.' }
    } finally {
        Pop-Location
    }
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    exit 1
}
