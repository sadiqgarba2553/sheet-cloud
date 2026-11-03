param([string]$BuildDirectory = 'src-tauri/target/release')
$ErrorActionPreference = 'Stop'
$releaseRoot = Split-Path -Parent $PSScriptRoot
$releaseConfig = Get-Content -LiteralPath (Join-Path $releaseRoot 'release.config.json') -Raw | ConvertFrom-Json
$releaseBuild = (Resolve-Path -LiteralPath (Join-Path $releaseRoot $BuildDirectory)).Path
$releaseOutput = Join-Path $releaseRoot 'releases/windows'
$releaseManaged = Join-Path $releaseOutput 'managed'
$releaseExe = Join-Path $releaseBuild $releaseConfig.package.executable
$releaseSetup = Join-Path $releaseBuild ('bundle/nsis/{0}_{1}_x64-setup.exe' -f ($releaseConfig.sourceName.Replace('-',' ')), $releaseConfig.version)
if (!(Test-Path -LiteralPath $releaseExe -PathType Leaf) -or !(Test-Path -LiteralPath $releaseSetup -PathType Leaf)) { throw 'Build the app executable and NSIS installer first.' }
$releaseZip = Join-Path $releaseRoot $releaseConfig.localAssets[0]
if (Test-Path -LiteralPath $releaseZip) { throw 'Release archive already exists. Do not replace published release bytes.' }
New-Item -ItemType Directory -Force -Path $releaseManaged | Out-Null
Copy-Item -LiteralPath $releaseExe -Destination $releaseManaged
foreach ($releaseFile in @('LICENSE','README.md','RELEASE-NOTES.md','docs/KNOWN_LIMITATIONS.md','public/THIRD_PARTY_NOTICES.txt')) {
    Copy-Item -LiteralPath (Join-Path $releaseRoot $releaseFile) -Destination $releaseManaged
}
Copy-Item -LiteralPath $releaseSetup -Destination (Join-Path $releaseRoot $releaseConfig.localAssets[1])
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($releaseManaged, $releaseZip, [System.IO.Compression.CompressionLevel]::Optimal, $false)
$releaseArtifacts = foreach ($releaseAsset in $releaseConfig.localAssets) {
    $releaseAssetPath = Join-Path $releaseRoot $releaseAsset
    [ordered]@{file=$releaseAsset; bytes=(Get-Item -LiteralPath $releaseAssetPath).Length; sha256=(Get-FileHash -LiteralPath $releaseAssetPath -Algorithm SHA256).Hash.ToLowerInvariant(); build='unsigned Windows x64 release-mode alpha'; installation_tested=$false}
}
$releaseArtifacts | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $releaseRoot 'docs/BUILD_ARTIFACTS.json')
Write-Output ('Prepared {0} {1} managed ZIP and standalone installer.' -f $releaseConfig.sourceName,$releaseConfig.version)
