param([Parameter(Mandatory)][guid]$RunId)
$ErrorActionPreference = 'Stop'
$Folder = Join-Path (Join-Path $PSScriptRoot 'ejecuciones') $RunId.ToString('D')
$State = Join-Path $Folder 'outlook.json'
$Previous = Get-Content -LiteralPath $State -Raw | ConvertFrom-Json
if ($Previous.estado -notin @('INTENTO_REGISTRADO','ENVIADO_PENDIENTE')) {
    throw 'La reanudacion requiere un intento de envio previo.'
}
& (Join-Path $PSScriptRoot 'office_ep3.ps1') -Action Outlook -RunFolder $Folder -ReceiptOnly
$Receipt = Get-Content -LiteralPath $State -Raw | ConvertFrom-Json
if ($Receipt.estado -eq 'CLASIFICADO') {
    # El acuse posterior conserva el resultado y el tiempo de la ejecucion original.
    Write-Host 'Outlook confirmado y clasificado. Revisar outlook.json en la carpeta original.'
}
