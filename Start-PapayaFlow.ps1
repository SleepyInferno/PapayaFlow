# Start-PapayaFlow.ps1 -- Entry point.

$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot

# Load modules in dependency order
. "$projectRoot\lib\PdfParser.ps1"
. "$projectRoot\lib\EntraMapper.ps1"
. "$projectRoot\lib\Aggregator.ps1"
. "$projectRoot\lib\Server.ps1"

$port = 8080
Write-Host "Starting PapayaFlow on http://localhost:$port"
Start-Process "http://localhost:$port"
Start-Server -Port $port -ProjectRoot $projectRoot
