param([switch]$NoAbrirResultado, [switch]$CorreoLocal, [switch]$Principal, [switch]$Visible, [string]$RocketbotExe='')
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$DemoScript = Join-Path $PSScriptRoot 'ejecutar_demo.cjs'
$Mutex = [Threading.Mutex]::new($false, 'Local\TestFireEP3Demo')
$Entered = $false
try {
    try { $Entered = $Mutex.WaitOne(0) }
    catch [Threading.AbandonedMutexException] { $Entered = $true }
    if (-not $Entered) { throw 'Ya se esta ejecutando la demo en este equipo.' }

    $Node = (Get-Command node.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1).Source
    $Runtime = (Get-AppxPackage 'OpenAI.CodexPrimaryRuntime*' -ErrorAction SilentlyContinue | Select-Object -First 1).InstallLocation
    if ($Runtime -and -not $env:EP3_PLAYWRIGHT) {
        $BundledNode = Join-Path $Runtime 'dependencies\node\bin\node.exe'
        $Playwright = Join-Path $Runtime 'dependencies\node\node_modules\playwright'
        if ((Test-Path -LiteralPath $BundledNode) -and (Test-Path -LiteralPath $Playwright)) {
            $Node = $BundledNode
            $env:EP3_PLAYWRIGHT = $Playwright
        }
    }
    if (-not $Node) {
        throw 'Falta el entorno Node. Abre ATAJOS/EP3/INICIAR_EP3.bat para elegir el robot y pulsar Run en Studio.'
    }

    & (Join-Path $PSScriptRoot 'abrir_studio.ps1') -Demo -CorreoLocal:$CorreoLocal -RocketbotExe $RocketbotExe -SoloPreparar
    $Robot = if ($Principal) { 'MAIN_TESTFIRE_EP3' } elseif ($CorreoLocal) { 'DEMO_CORREO_LOCAL_EP3' } else { 'DEMO_TESTFIRE_EP3' }
    Write-Host "Ejecutando $Robot en Rocketbot Studio..."
    $BrowserOptions = @()
    if ($Visible) { $BrowserOptions += '--visible' }
    if ($NoAbrirResultado) { $BrowserOptions += '--close-after-run' }
    $Result = $null
    & $Node $DemoScript $Robot @BrowserOptions | ForEach-Object {
        if ($_ -notlike '{*') {
            Write-Host $_
        } else {
            $Result = $_ | ConvertFrom-Json
            $Mutex.ReleaseMutex()
            $Entered = $false
            $Seconds = $Result.duration_ms / 1000
            Write-Host ('FINALIZADO | {0} | {1:N2} s | Correo: {2}' -f $Result.status, $Seconds, $Result.outlook)
            Write-Host "Origen: $($Result.origin) | Validos: $($Result.valid_rows) | Rechazados: $($Result.rejected_rows)"
            if ($Result.message) { Write-Host $Result.message }
            Write-Host "Resultados: $($Result.folder)"
            if ($Result.report) { Write-Host "Excel: $($Result.report)" }
            if ($CorreoLocal -and $Result.outlook -eq 'CLASIFICADO') {
                Write-Host 'Correo recibido en el buzon local: Correo_EP3.eml'
            }
            if (-not $NoAbrirResultado -and $Principal -and $Result.outlook -eq 'SMTP_ACEPTADO') {
                try {
                    & (Join-Path $PSScriptRoot 'abrir_evidencia_correo.ps1') -RunFolder $Result.folder
                    Write-Host 'Gmail abierto buscando el mensaje de esta ejecucion.'
                } catch {
                    Write-Warning ('Correo enviado. No se pudo abrir Gmail: ' + $_.Exception.Message)
                }
            }
        }
    }
    if ($LASTEXITCODE -ne 0 -or -not $Result) { throw 'Studio no confirmo el resultado de la ejecucion.' }
    if (-not $Result.report) { throw "La ejecucion no genero un reporte. Estado de la cola: $($Result.queue). Revisa execution.json." }
    if ($Result.status -eq 'ERROR') { throw 'La ejecucion termino con error. Revisa execution.json.' }
    if ($CorreoLocal) {
        if ($Result.status -ne 'OK' -or $Result.outlook -ne 'CLASIFICADO') {
            throw 'El buzon local no confirmo el flujo completo. Revisa execution.json.'
        }
    }
    if (-not $NoAbrirResultado) {
        Start-Process -FilePath explorer.exe -ArgumentList ('"{0}"' -f $Result.folder)
    }
} finally {
    if ($Entered) { $Mutex.ReleaseMutex() }
    $Mutex.Dispose()
}
