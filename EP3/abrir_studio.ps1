param([switch]$Demo, [switch]$CorreoLocal, [string]$RocketbotExe='', [switch]$SoloPreparar)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Database = Join-Path $Root 'rocketbot\MAIN_TESTFIRE_EP3\robot.db'
$Robot = if ($CorreoLocal) { 'DEMO_CORREO_LOCAL_EP3' } elseif ($Demo) { 'DEMO_TESTFIRE_EP3' } else { 'MAIN_TESTFIRE_EP3' }
if (-not (Test-Path -LiteralPath $Database -PathType Leaf)) { throw 'Falta el proyecto EP3.' }
$Studio = 'http://127.0.0.1:5000'
function Test-Studio {
    try { (Invoke-WebRequest -UseBasicParsing "$Studio/editor/" -TimeoutSec 4).StatusCode -eq 200 }
    catch { $false }
}
if (-not (Test-Studio)) {
    . (Join-Path $Root 'scripts\rocketbot_path.ps1')
    if (-not $RocketbotExe -and $env:LOCALAPPDATA) {
        $LocalExe = Join-Path $env:LOCALAPPDATA 'RocketbotEP3\rocketbot.exe'
        if (Test-Path -LiteralPath $LocalExe -PathType Leaf) { $RocketbotExe = $LocalExe }
    }
    $RocketbotExe = Resolve-RocketbotExecutable -PreferredPath $RocketbotExe -ChooseIfMissing
    Start-Process -FilePath $RocketbotExe -WorkingDirectory (Split-Path -Parent $RocketbotExe) -WindowStyle Hidden
    $Deadline = [datetime]::UtcNow.AddSeconds(25)
    do {
        Start-Sleep -Seconds 1
        $Ready = Test-Studio
    } until ($Ready -or [datetime]::UtcNow -gt $Deadline)
    if (-not $Ready) {
        throw 'Studio no responde. Completar la activacion o responder la pregunta inicial de drivers y volver a iniciar.'
    }
}
if ($SoloPreparar) { return }
$Encoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Database))
Start-Process "$Studio/editor/#/edit/$Robot/db/$Encoded"
