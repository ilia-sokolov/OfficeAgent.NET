<#
.SYNOPSIS
Builds the seed corpus for the ingestion fuzz target from the repository's own fixtures.

.DESCRIPTION
Collects every .docx, .pptx and .xlsx under samples/documents and tests/OfficeAgent.Tests/Corpus
into one flat folder outside the repository, so no corpus is committed. Each file is written twice:

  orig-<path>    the fixture byte for byte, deflate-compressed as it was saved;
  stored-<path>  the same package re-written with every entry stored uncompressed.

Byte mutations on a deflated entry almost always break the deflate stream, so the fuzzer would
spend its time in the archive reader. A stored copy lets a mutation land in the XML itself and
reach the Open XML SDK and the format modules. Files that are not ZIP archives (the encrypted
fixtures are OLE compound files) are kept only as orig-. Duplicates are dropped by SHA-256, and
anything over 1 MiB is skipped, because libfuzzer-dotnet refuses inputs larger than that.

.PARAMETER Output
The folder to write. It is emptied first. Defaults to $env:TEMP\officeagent-fuzz\seeds.

.EXAMPLE
./tests/OfficeAgent.Fuzz/build-seeds.ps1
#>
param(
    [string]$Output = (Join-Path $env:TEMP 'officeagent-fuzz\seeds')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
$MaxInput = 1MB
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$sources = @(
    (Join-Path $root 'samples\documents'),
    (Join-Path $root 'tests\OfficeAgent.Tests\Corpus')
)

if (Test-Path $Output) { Remove-Item -Recurse -Force $Output }
New-Item -ItemType Directory -Force $Output | Out-Null

function Get-Sha256([byte[]]$bytes) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '') }
    finally { $sha.Dispose() }
}

# Re-writes a package with every entry stored, keeping entry order and names. Returns $null when
# the bytes are not a readable archive.
function ConvertTo-StoredZip([byte[]]$bytes) {
    try {
        $source = New-Object System.IO.Compression.ZipArchive(
            (New-Object System.IO.MemoryStream(, $bytes)), [System.IO.Compression.ZipArchiveMode]::Read)
    }
    catch { return $null }

    $buffer = New-Object System.IO.MemoryStream
    try {
        $target = New-Object System.IO.Compression.ZipArchive(
            $buffer, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            foreach ($entry in $source.Entries) {
                $copy = $target.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::NoCompression)
                $copy.LastWriteTime = $entry.LastWriteTime
                $in = $entry.Open(); $out = $copy.Open()
                try { $in.CopyTo($out) } finally { $out.Dispose(); $in.Dispose() }
            }
        }
        finally { $target.Dispose() }
    }
    catch { return $null }
    finally { $source.Dispose() }

    return $buffer.ToArray()
}

$seen = @{}
$written = 0; $skipped = 0

function Write-Seed([string]$name, [byte[]]$bytes) {
    if ($bytes.Length -gt $MaxInput) { $script:skipped++; return }
    $hash = Get-Sha256 $bytes
    if ($seen.ContainsKey($hash)) { return }
    $seen[$hash] = $true
    [System.IO.File]::WriteAllBytes((Join-Path $Output $name), $bytes)
    $script:written++
}

$files = foreach ($source in $sources) {
    Get-ChildItem -Recurse -File $source | Where-Object { $_.Extension -in '.docx', '.pptx', '.xlsx' }
}

foreach ($file in ($files | Sort-Object FullName)) {
    $relative = $file.FullName.Substring($root.Length + 1)
    $flat = ($relative -replace '[\\/]', '_')
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

    Write-Seed "orig-$flat" $bytes

    $stored = ConvertTo-StoredZip $bytes
    if ($null -ne $stored) { Write-Seed "stored-$flat" $stored }
}

Write-Output "Wrote $written seeds to $Output ($skipped over 1 MiB skipped)."
