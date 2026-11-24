param(
    [Parameter(Mandatory)][ValidateSet('Excel','Outlook','OutlookCheck')][string]$Action,
    [Parameter(Mandatory)][string]$RunFolder,
    [switch]$AllowSend,
    [switch]$ReceiptOnly
)
$ErrorActionPreference = 'Stop'
$RunFolder = (Resolve-Path -LiteralPath $RunFolder).Path
$RunId = Split-Path -Leaf $RunFolder
$ParsedId = [guid]::Empty
if (-not [guid]::TryParseExact($RunId, 'D', [ref]$ParsedId)) { throw 'Carpeta de ejecucion invalida.' }
$ReportName = 'Reporte_TestFire_EP3.xlsx'
$Workbook = Join-Path $RunFolder $ReportName
$Artifact = Get-Content -LiteralPath (Join-Path $RunFolder 'artefacto.json') -Raw | ConvertFrom-Json
if ($Artifact.execution_id -ne $RunId -or (Get-FileHash -LiteralPath $Workbook).Hash -ine $Artifact.sha256) {
    throw 'El Excel no coincide con el artefacto validado.'
}
function Write-Evidence($Name, $Data) {
    $Target = Join-Path $RunFolder $Name
    $Temp = "$Target.tmp"
    $Data | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $Temp -Encoding UTF8
    if ([IO.File]::Exists($Target)) { [IO.File]::Replace($Temp, $Target, [NullString]::Value) }
    else { [IO.File]::Move($Temp, $Target) }
}

if ($Action -eq 'Excel') {
    if (-not [type]::GetTypeFromProgID('Excel.Application')) { exit 12 }
    $Excel = $null
    $Book = $null
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.DisplayAlerts = $false
        $Excel.AutomationSecurity = 3
        $Book = $Excel.Workbooks.Open($Workbook, 0, $true)
        if ($Book.Worksheets.Item('Resumen').Range('B2').Value2 -ne $RunId) {
            throw 'Excel abrio un libro que no corresponde a la ejecucion.'
        }
        Write-Evidence 'excel.json' @{ estado='ABIERTO'; hojas=$Book.Worksheets.Count; solo_lectura=$Book.ReadOnly }
    } finally {
        if ($Book) { $Book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Book) }
        if ($Excel) { $Excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) }
    }
    return
}

$ConfigPath = Join-Path $PSScriptRoot 'outlook.local.json'
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    Write-Evidence 'outlook.json' @{estado='OMITIDO'; motivo='Falta outlook.local.json'}
    return
}
$Config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
if ($Config.mode -notin @('BORRADOR','ENVIO') -or $Config.mailbox -notmatch '^[^@\s;,"<>]+@[^@\s;,"<>]+\.[^@\s;,"<>]+$') {
    throw 'Configurar BORRADOR o ENVIO y exactamente una cuenta de pruebas valida.'
}
if (-not (Test-Path -LiteralPath (Join-Path $RunFolder 'excel.json') -PathType Leaf)) {
    throw 'Falta la evidencia de apertura de Microsoft Excel.'
}
$Lock = [IO.File]::Open((Join-Path $RunFolder 'outlook.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
try {
    $StateFile = Join-Path $RunFolder 'outlook.json'
    $Previous = $null
    if (Test-Path -LiteralPath $StateFile) {
        $Previous = Get-Content -LiteralPath $StateFile -Raw | ConvertFrom-Json
        if ($Previous.estado -eq 'CLASIFICADO') { return }
        if ($Previous.estado -eq 'BORRADOR' -and $Config.mode -eq 'BORRADOR') { return }
    }
    if ($ReceiptOnly -and $Previous.estado -notin @('INTENTO_REGISTRADO','ENVIADO_PENDIENTE')) { return }
    if (-not [type]::GetTypeFromProgID('Outlook.Application')) { exit 12 }
    $Outlook = New-Object -ComObject Outlook.Application
    $Session = $Outlook.GetNamespace('MAPI')
    $Accounts = @($Session.Accounts | Where-Object { $_.SmtpAddress -ieq $Config.mailbox })
    if ($Accounts.Count -ne 1 -or -not $Accounts[0].DeliveryStore) {
        throw 'La cuenta de pruebas debe existir una sola vez en el perfil de Outlook clasico.'
    }
    $Account = $Accounts[0]
    $Store = $Account.DeliveryStore
    if ($Action -eq 'OutlookCheck') { return }
    $Subject = "[RPA EP3] TestFire $RunId"
    $Filter = "[Subject] = '$Subject'"
    $Sent = $Store.GetDefaultFolder(5)
    $Inbox = $Store.GetDefaultFolder(6)
    if ($Previous.estado -notin @('INTENTO_REGISTRADO','ENVIADO_PENDIENTE')) {
        if ($Sent.Items.Restrict($Filter).Count -gt 0 -or $Inbox.Items.Restrict($Filter).Count -gt 0) {
            throw 'Ya existe correo para esta ejecucion. Verificar el estado antes de reintentar.'
        }
        if ($Config.mode -eq 'ENVIO' -and $Previous.estado -eq 'BORRADOR') {
            throw 'Esta ejecucion tiene un borrador. Ejecutar el robot otra vez para una prueba de envio nueva.'
        }
        $Mail = $Store.GetDefaultFolder(16).Items.Add('IPM.Note')
        $Mail.SendUsingAccount = $Account
        $Mail.SaveSentMessageFolder = $Sent
        [void]$Mail.Recipients.Add($Account.SmtpAddress)
        if (-not $Mail.Recipients.ResolveAll()) { throw 'No se resolvio la cuenta de pruebas.' }
        $Mail.Subject = $Subject
        $Mail.Body = "Evaluacion Parcial III de Automatizacion de Procesos Roboticos. Adjunto reporte agregado de TestFire con referencia publica del Banco Mundial. Ejecucion $RunId."
        [void]$Mail.Attachments.Add($Workbook)
        if ($Config.mode -eq 'BORRADOR') {
            $Mail.Save()
            Write-Evidence 'outlook.json' @{estado='BORRADOR'; asunto=$Subject; adjunto=$ReportName}
            return
        }
        if (-not $AllowSend) { throw 'No hay reserva de envio vigente.' }
        Write-Evidence 'outlook.json' @{estado='INTENTO_REGISTRADO' ; asunto=$Subject; adjunto=$ReportName}
        $Mail.Send()
        Write-Evidence 'outlook.json' @{estado='ENVIADO_PENDIENTE'; asunto=$Subject; adjunto=$ReportName}
    }
    $Category = 'RPA EP3 Procesado'
    $Separator = (Get-Culture).TextInfo.ListSeparator
    for ($Attempt=0; $Attempt -lt 12; $Attempt++) {
        $Receipt = $Inbox.Items.Restrict($Filter)
        $SentCount = $Sent.Items.Restrict($Filter).Count
        if ($Receipt.Count -gt 1 -or $SentCount -gt 1) { throw 'Correo duplicado. Revisar manualmente.' }
        if ($Receipt.Count -eq 1 -and $SentCount -eq 1) {
            $Message = $Receipt.Item(1)
            if ($Message.Class -ne 43 -or $Message.Attachments.Count -ne 1 -or
                    $Message.Attachments.Item(1).FileName -ne $ReportName) {
                throw 'El mensaje recibido no coincide con el adjunto esperado.'
            }
            $Existing = @($Message.Categories -split [regex]::Escape($Separator) | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            if ($Category -notin $Existing) {
                $Message.Categories = ($Existing + $Category) -join "$Separator "
                $Message.Save()
            }
            Write-Evidence 'outlook.json' @{estado='CLASIFICADO'; asunto=$Subject; adjunto=$ReportName;
                enviados=1; recibidos=1; categoria=$Category; confirmado_utc=[datetime]::UtcNow.ToString('o')}
            return
        }
        Start-Sleep -Seconds 4
    }
    Write-Host 'Outlook aun no confirma la recepcion. Reanudar esta carpeta; el robot no reenviara.'
} finally {
    $Lock.Dispose()
}
