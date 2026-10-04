<#
  build.ps1 - pack/unpack the PPT Tools deck

  .\build.ps1 pack     Copies ribbon\customUI.xml into deck\ and zips deck\ into build\PPTTools.pptm
  .\build.ps1 unpack   Unzips build\PPTTools.pptm back into deck\ (run after importing macros and saving)
  .\build.ps1 release  Zips build\PPTTools.ppam + installer\*.cmd into build\PPT-Tools.zip (the one-step install download)
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('pack', 'unpack', 'release')]
    [string]$Action
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$root  = $PSScriptRoot
$deck  = Join-Path $root 'deck'
$build = Join-Path $root 'build'
$pptm  = Join-Path $build 'PPTTools.pptm'

if ($Action -eq 'pack') {
    # The ribbon XML in ribbon\ is the master copy; the deck gets a fresh copy on every build
    Copy-Item -LiteralPath (Join-Path $root 'ribbon\customUI.xml') -Destination (Join-Path $deck 'customUI\customUI.xml') -Force

    New-Item -ItemType Directory -Force -Path $build | Out-Null
    if (Test-Path -LiteralPath $pptm) { Remove-Item -LiteralPath $pptm }

    # [Content_Types].xml goes first, matching how Office writes its files
    $files = Get-ChildItem -LiteralPath $deck -Recurse -File -Force |
        Sort-Object { $_.Name -ne '[Content_Types].xml' }, FullName

    $zip = [System.IO.Compression.ZipFile]::Open($pptm, 'Create')
    try {
        foreach ($f in $files) {
            $entry = $f.FullName.Substring($deck.Length + 1).Replace('\', '/')
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $f.FullName, $entry, 'Optimal') | Out-Null
        }
    }
    finally {
        $zip.Dispose()
    }
    "Packed $($files.Count) files into $pptm"
}
elseif ($Action -eq 'release') {
    $ppam = Join-Path $build 'PPTTools.ppam'
    if (-not (Test-Path -LiteralPath $ppam)) { throw "Not found: $ppam - build the add-in first." }
    $zipPath = Join-Path $build 'PPT-Tools.zip'
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath }

    $zip = [System.IO.Compression.ZipFile]::Open($zipPath, 'Create')
    try {
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $ppam, 'PPTTools.ppam', 'Optimal') | Out-Null
        foreach ($f in Get-ChildItem -LiteralPath (Join-Path $root 'installer') -Filter '*.cmd') {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $f.FullName, $f.Name, 'Optimal') | Out-Null
        }
    }
    finally {
        $zip.Dispose()
    }
    "Created $zipPath"
}
else {
    if (-not (Test-Path -LiteralPath $pptm)) { throw "Not found: $pptm - run '.\build.ps1 pack' first." }

    Remove-Item -LiteralPath $deck -Recurse -Force
    [System.IO.Compression.ZipFile]::ExtractToDirectory($pptm, $deck)
    "Unpacked $pptm into $deck"
}
