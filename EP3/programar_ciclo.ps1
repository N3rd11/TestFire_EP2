param([Parameter(Mandatory)][string]$RocketbotExe, [switch]$Demo)
$ErrorActionPreference = 'Stop'
$Mode = if ($Demo) { 'Demo' } else { 'Web' }
$TaskName = "TestFire-EP3-$Mode"
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    throw "Ya existe $TaskName. Revise su configuracion en el Programador de tareas."
}
$Exe = (Resolve-Path -LiteralPath $RocketbotExe).Path
if ((Split-Path -Leaf $Exe) -ine 'rocketbot.exe') { throw 'Seleccione rocketbot.exe.' }
$Cycle = Join-Path $PSScriptRoot 'ejecutar_ciclo.ps1'
$Arg = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -RocketbotExe "{1}"' -f $Cycle, $Exe
if ($Demo) { $Arg += ' -Demo' }
$Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $Arg -WorkingDirectory $PSScriptRoot
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) -RepetitionInterval (New-TimeSpan -Minutes 5)
$Options = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 14)
$Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$Principal = New-ScheduledTaskPrincipal -UserId $Account -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Options -Principal $Principal
