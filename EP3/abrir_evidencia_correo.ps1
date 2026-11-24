param([Parameter(Mandatory = $true)][string]$RunFolder)

$ErrorActionPreference = 'Stop'
$directory = Get-Item -LiteralPath $RunFolder
$receipt = Get-Content -LiteralPath (Join-Path $directory.FullName 'outlook.json') -Raw -Encoding UTF8 |
    ConvertFrom-Json
$expectedId = '<{0}@testfire-ep3.invalid>' -f $directory.Name
if ($receipt.canal -ne 'GMAIL_SMTP_DUOC' -or $receipt.estado -ne 'SMTP_ACEPTADO' -or
    $receipt.destinatario -notmatch '^[a-zA-Z0-9_.+-]+@duocuc[.]cl$' -or
    $directory.Name -notmatch '^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$' -or
    $receipt.message_id -cne $expectedId) {
    throw 'Esta corrida no tiene un acuse de envio a Duoc.'
}

$account = [uri]::EscapeDataString($receipt.destinatario)
$query = [uri]::EscapeDataString('rfc822msgid:' + $receipt.message_id.Trim('<', '>'))
$mailUrl = 'https://mail.google.com/mail/?authuser={0}#search/{1}' -f $account, $query
Start-Process -FilePath $mailUrl
