# Runs the tests, builds the single exe and packs the Claude Desktop extension file (.mcpb).
# Needs: .NET 10 SDK and Node.js (for the official mcpb packer). Developers only;
# end users never run this.
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$out = Join-Path $root "extension\server"

dotnet test (Join-Path $root "tests\OutlookClassicMcp.Tests")
if ($LASTEXITCODE -ne 0) { throw "tests failed" }

dotnet publish (Join-Path $root "src\OutlookClassicMcp") -c Release -o $out
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed" }

$exe = Join-Path $out "outlook-classic-mcp.exe"
"{0}: {1:N1} MB" -f $exe, ((Get-Item $exe).Length / 1MB)

New-Item -ItemType Directory -Force (Join-Path $root "dist") | Out-Null
npx --yes @anthropic-ai/mcpb pack (Join-Path $root "extension") (Join-Path $root "dist\outlook-classic-mcp.mcpb")
if ($LASTEXITCODE -ne 0) { throw "mcpb pack failed" }
