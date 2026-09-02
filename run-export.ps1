[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $ExporterArguments
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$localModules = Join-Path $projectRoot "node_modules"
$bundledRoot = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies"
$bundledNode = Join-Path $bundledRoot "node\bin\node.exe"
$bundledModules = Join-Path $bundledRoot "node\node_modules"

if (-not (Test-Path -LiteralPath (Join-Path $localModules "@oai\artifact-tool"))) {
    if (-not (Test-Path -LiteralPath (Join-Path $bundledModules "@oai\artifact-tool"))) {
        throw "Bundled @oai/artifact-tool was not found. Run this tool from Codex Desktop or install the dependency first."
    }
    if (Test-Path -LiteralPath $localModules) {
        throw "node_modules exists but does not contain @oai/artifact-tool. Move that folder elsewhere and try again."
    }
    New-Item -ItemType Junction -Path $localModules -Target $bundledModules | Out-Null
}

if (Test-Path -LiteralPath $bundledNode) {
    $nodeExecutable = $bundledNode
} else {
    $nodeExecutable = (Get-Command node -ErrorAction Stop).Source
}

Push-Location $projectRoot
try {
    & $nodeExecutable "src/woodpecker-export.mjs" @ExporterArguments
    exit $LASTEXITCODE
} finally {
    Pop-Location
}
