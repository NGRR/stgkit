$ErrorActionPreference = 'Stop'
$catalogRoot = $PSScriptRoot
$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCommand) {
    $pythonCommand = Get-Command py -ErrorAction Stop
}

$listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
$listener.Start()
$port = $listener.LocalEndpoint.Port
$listener.Stop()
$quotedRoot = '"' + $catalogRoot + '"'
$serverArgs = @('-m', 'http.server', $port, '--bind', '127.0.0.1', '--directory', $quotedRoot)
$server = Start-Process -FilePath $pythonCommand.Source -ArgumentList $serverArgs -WindowStyle Hidden -PassThru

try {
    Start-Sleep -Milliseconds 700
    if ($server.HasExited) {
        throw 'No fue posible iniciar el servidor local del catálogo.'
    }
    Start-Process "http://127.0.0.1:$port/index.html"
    Write-Host "Catálogo STGND disponible en http://127.0.0.1:$port/index.html"
    Write-Host 'Presione Enter para cerrar el catálogo.'
    [void](Read-Host)
}
finally {
    if ($server -and -not $server.HasExited) {
        Stop-Process -Id $server.Id
    }
}
