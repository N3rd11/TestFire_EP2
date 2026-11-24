function Find-Ep3Python {
    $Root = Split-Path -Parent $PSScriptRoot
    $Venv = Join-Path $Root '.venv-ep2\Scripts\python.exe'
    if (Test-Path -LiteralPath $Venv) { return $Venv }
    $Runtime = (Get-AppxPackage 'OpenAI.CodexPrimaryRuntime*' -ErrorAction SilentlyContinue | Select-Object -First 1).InstallLocation
    if ($Runtime) {
        $Bundled = Join-Path $Runtime 'dependencies\python\python.exe'
        if (Test-Path -LiteralPath $Bundled) { return $Bundled }
    }
    $Installed = Get-Command python.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($Installed -and $Installed.Source -notlike '*\Microsoft\WindowsApps\*') { return $Installed.Source }
    throw 'Falta Python para Gmail. Prepara el entorno de EP2 o instala Python.'
}
