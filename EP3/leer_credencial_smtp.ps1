param([Parameter(Mandatory)][string]$Ruta)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$Credential = Import-Clixml -LiteralPath $Ruta
if ($Credential -isnot [Management.Automation.PSCredential]) { throw 'Archivo de credencial no valido.' }
@{username=$Credential.UserName;password=$Credential.GetNetworkCredential().Password} | ConvertTo-Json -Compress
