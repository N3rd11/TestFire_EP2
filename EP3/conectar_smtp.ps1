param([string]$Correo='')
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'python_correo.ps1')
$Python = Find-Ep3Python
$Runtime = Join-Path $PSScriptRoot '.runtime-correo'
if (-not $Correo) { $Correo = Read-Host 'Tu correo Duoc o Gmail para enviarte el reporte' }
if ($Correo -notmatch '^[A-Z0-9_.+-]+@(duocuc\.cl|gmail\.com)$') {
    Write-Host 'Ingresa un unico correo Duoc o Gmail.' -ForegroundColor Yellow
    exit 1
}
$AppPassword = Read-Host 'Contrasena de aplicacion de Google de 16 caracteres, sin espacios' -AsSecureString
if ($AppPassword.Length -ne 16) {
    Write-Host 'Debe ser una contrasena de aplicacion de 16 caracteres. No ingreses tu clave habitual.' -ForegroundColor Yellow
    exit 1
}
$Temp = Join-Path $PSScriptRoot ('smtp-preparacion-{0}.clixml' -f [guid]::NewGuid())
Push-Location -LiteralPath $Root
try {
    [Management.Automation.PSCredential]::new($Correo, $AppPassword) | Export-Clixml -LiteralPath $Temp
    if (-not (Test-Path -LiteralPath (Join-Path $Runtime 'flask'))) {
        & $Python -m pip install --disable-pip-version-check --target $Runtime 'Flask>=3.1,<4'
        if ($LASTEXITCODE) { throw 'No se pudo preparar la cola local.' }
    }
    & $Python -m auxiliaries.ep3_conectar_smtp $Correo $Temp
    if ($LASTEXITCODE) { exit $LASTEXITCODE }
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    exit 1
} finally {
    if (Test-Path -LiteralPath $Temp) { Remove-Item -LiteralPath $Temp }
    Pop-Location
}
