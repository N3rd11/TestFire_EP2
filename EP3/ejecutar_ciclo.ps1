param([Parameter(Mandatory)][string]$RocketbotExe, [switch]$Demo)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Config = Join-Path $PSScriptRoot 'operacion.local.json'
if (-not (Test-Path -LiteralPath $Config -PathType Leaf)) {
    throw 'Configure operacion.local.json antes de ejecutar sin supervisión.'
}
. (Join-Path $Root 'scripts\rocketbot_path.ps1')
$Exe = Resolve-RocketbotExecutable -PreferredPath $RocketbotExe
$DB = Join-Path $Root 'rocketbot\MAIN_TESTFIRE_EP3\robot.db'
$Robot = if ($Demo) { 'DEMO_TESTFIRE_EP3' } else { 'MAIN_TESTFIRE_EP3' }
$Process = Start-Process -FilePath $Exe -ArgumentList @("-start=$Robot", ('-db="{0}"' -f $DB)) -WorkingDirectory (Split-Path -Parent $Exe) -WindowStyle Hidden -PassThru
if (-not $Process.WaitForExit(720000)) {
    throw 'Rocketbot excedió el tiempo de espera. Revisar sus registros y no iniciar otro proceso local al mismo tiempo.'
}
exit $Process.ExitCode
