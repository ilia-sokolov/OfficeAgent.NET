<#
.SYNOPSIS
Publishes, instruments and runs the ingestion fuzz target under libFuzzer on Windows.

.DESCRIPTION
1. Checks libfuzzer-dotnet-windows.exe against the SHA-256 recorded in README.md.
2. Publishes tests/OfficeAgent.Fuzz to <WorkDir>\bin, outside the repository.
3. Instruments the assemblies that parse the input with the sharpfuzz tool.
4. Builds the seed corpus with build-seeds.ps1 if <Seeds> does not exist yet.
5. Runs one libFuzzer worker. New inputs go to <Corpus>; crashes, timeouts and their
   minimised forms go to <Findings>.

Everything this writes is outside the repository unless you point it inside.

.PARAMETER LibFuzzer
Path to libfuzzer-dotnet-windows.exe from the pinned libfuzzer-dotnet release.

.PARAMETER SharpFuzz
The sharpfuzz command from SharpFuzz.CommandLine 2.3.0. Defaults to <WorkDir>\tools\sharpfuzz.exe.

.PARAMETER MaxTotalTime
Seconds to fuzz before stopping. 0 runs until a finding or Ctrl+C.

.PARAMETER Timeout
Seconds one input may run before libFuzzer reports it as a hang.

.PARAMETER KeepGoing
Run libFuzzer in fork mode with one job, so a crash or a hang is saved and fuzzing continues
instead of stopping at the first finding.

.PARAMETER InstrumentOpenXmlSchema
Also instrument DocumentFormat.OpenXml.dll, the SDK's generated schema classes. Off by default;
see README.md.

.PARAMETER SkipBuild
Reuse the instrumented build already in <WorkDir>\bin.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$LibFuzzer,
    [string]$WorkDir = (Join-Path $env:TEMP 'officeagent-fuzz'),
    [string]$SharpFuzz = '',
    [string]$Seeds = '',
    [string]$Corpus = '',
    [string]$Findings = (Join-Path $env:TEMP 'officeagent-fuzz-findings'),
    [int]$MaxTotalTime = 0,
    [int]$Timeout = 10,
    [switch]$KeepGoing,
    [switch]$InstrumentOpenXmlSchema,
    [switch]$SkipBuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# libfuzzer-dotnet v2025.05.02.0904 publishes no checksum or signature; this is the SHA-256 of the
# release asset as downloaded, so a later download that differs is noticed.
$LibFuzzerSha256 = '17AF5B3F6FF4D2C57B44B9A35C13051B570EB66F0557D00015DF3832709050BF'

if ($SharpFuzz -eq '') { $SharpFuzz = Join-Path $WorkDir 'tools\sharpfuzz.exe' }
if ($Seeds -eq '') { $Seeds = Join-Path $WorkDir 'seeds' }
if ($Corpus -eq '') { $Corpus = Join-Path $WorkDir 'corpus' }
$bin = Join-Path $WorkDir 'bin'
$project = Join-Path $PSScriptRoot 'OfficeAgent.Fuzz.csproj'

$actual = (Get-FileHash -Algorithm SHA256 $LibFuzzer).Hash
if ($actual -ne $LibFuzzerSha256) {
    throw "libfuzzer-dotnet-windows.exe has SHA-256 $actual, expected $LibFuzzerSha256. Download the pinned release again."
}

# The assemblies that see the input's bytes. OfficeAgent's own, the SDK's parsing core, the OPC
# reader, and the number-format parser Excel inspection runs on every cell's format string.
# System.IO.Compression and System.Xml ship in the shared runtime and cannot be instrumented.
$targets = @(
    'OfficeAgent.Abstractions.dll',
    'OfficeAgent.Core.dll',
    'OfficeAgent.Word.dll',
    'OfficeAgent.PowerPoint.dll',
    'OfficeAgent.Excel.dll',
    'DocumentFormat.OpenXml.Framework.dll',
    'System.IO.Packaging.dll',
    'ExcelNumberFormat.dll'
)
if ($InstrumentOpenXmlSchema) { $targets += 'DocumentFormat.OpenXml.dll' }

if (-not $SkipBuild) {
    if (Test-Path $bin) { Remove-Item -Recurse -Force $bin }
    dotnet publish $project -c Release -o $bin
    if ($LASTEXITCODE -ne 0) { throw 'dotnet publish failed.' }

    foreach ($target in $targets) {
        Write-Output "Instrumenting $target"
        & $SharpFuzz (Join-Path $bin $target)
        if ($LASTEXITCODE -ne 0) { throw "sharpfuzz failed on $target." }
    }
}

if (-not (Test-Path $Seeds)) { & (Join-Path $PSScriptRoot 'build-seeds.ps1') -Output $Seeds }
New-Item -ItemType Directory -Force $Corpus, $Findings | Out-Null

$arguments = @(
    "-timeout=$Timeout",
    # libfuzzer-dotnet refuses an input over 1 MiB.
    '-max_len=1048576',
    "-dict=$(Join-Path $PSScriptRoot 'ooxml.dict')",
    "-artifact_prefix=$Findings\",
    '-print_final_stats=1',
    '-report_slow_units=5'
)
if ($MaxTotalTime -gt 0) { $arguments += "-max_total_time=$MaxTotalTime" }
if ($KeepGoing) { $arguments += @('-fork=1', '-ignore_crashes=1', '-ignore_timeouts=1', '-ignore_ooms=1') }
$arguments += @(
    '--target_path=dotnet',
    "--target_arg=$(Join-Path $bin 'OfficeAgent.Fuzz.dll')",
    $Corpus,
    $Seeds
)

Write-Output "$LibFuzzer $($arguments -join ' ')"
# libFuzzer reports on stderr. Windows PowerShell turns redirected stderr lines into errors, which
# must not stop the run.
$ErrorActionPreference = 'Continue'
& $LibFuzzer @arguments
exit $LASTEXITCODE
