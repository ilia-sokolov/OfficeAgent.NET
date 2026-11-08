# Ingestion fuzz target

A [SharpFuzz](https://github.com/Metalnem/sharpfuzz) target for the path every document takes
into OfficeAgent. It takes arbitrary bytes, wraps them in a `StreamHandle`, and runs
`InspectAsync` through an `OfficeAgentClient` with the Word, PowerPoint and Excel modules
registered and the default [ingestion limits](../../docs/ingestion-limits.md).

## What counts as a finding

An input may produce a result, or one of the three documented refusals:

| Refusal | Code |
| --- | --- |
| `OpenXmlIngestionLimitException` | `input-too-large` |
| `OpenXmlPackageRejectedException` | `malformed-package` |
| `EncryptedDocumentException` | `encrypted-document` |

Anything else is a finding:

- **Any other exception.** It escapes the target and libFuzzer saves the input as `crash-*`.
- **A hang.** An input still running after `-timeout` seconds (10 by default) is saved as
  `timeout-*`.
- **Live memory past 2 GiB.** The project sets `System.GC.HeapHardLimit`, so an input that needs
  more managed heap fails with `OutOfMemoryException`, which is a crash like any other.
- **Work past 1 GiB of allocation.** An inspection that allocates more than 1 GiB over its
  lifetime throws `FuzzFindingException`. Inputs are at most 1 MiB and the default expansion
  ceiling bounds that to about 200 MiB of XML, so more means the ceilings did not bound the work.
  `OFFICEAGENT_FUZZ_MAX_ALLOCATED_MB` changes it.
- **A long refusal.** A refusal message of 600 characters or more throws `FuzzFindingException`.
  Refusals are documented to stay short, and `BoundedIngestionTests` holds them to the same bound.

libFuzzer's own `-rss_limit_mb` measures the libFuzzer process, not the .NET process it drives,
which is why the memory bounds live in the target.

## Pinned tools

| Tool | Version | Where |
| --- | --- | --- |
| `SharpFuzz` NuGet package | 2.3.0 | `OfficeAgent.Fuzz.csproj`, locked in `packages.lock.json` |
| `SharpFuzz.CommandLine` (`sharpfuzz`) | 2.3.0 | installed into a tool folder outside the repository |
| `libfuzzer-dotnet-windows.exe` | release `v2025.05.02.0904` | downloaded outside the repository |

The libfuzzer-dotnet release publishes no checksum and the binary is not signed. It is built by
that repository's release workflow from `libfuzzer-dotnet-windows.cc`. The size published by the
GitHub API is 1,512,448 bytes, and `fuzz.ps1` refuses any file whose SHA-256 is not:

```text
17AF5B3F6FF4D2C57B44B9A35C13051B570EB66F0557D00015DF3832709050BF
```

No binary, corpus or finding is committed. Everything below writes under `$env:TEMP`.

## Run it on Windows

Requires Windows, the .NET 10 SDK and PowerShell. Run from the repository root. Keep the work
folder on a path without spaces, because libfuzzer-dotnet starts `dotnet <target>` unquoted.

1. Download libfuzzer-dotnet into a new, empty folder and check it:

   ```powershell
   $lf = Join-Path $env:TEMP 'libfuzzer-dotnet-v2025.05.02.0904'
   New-Item -ItemType Directory $lf | Out-Null
   Invoke-WebRequest -UseBasicParsing -OutFile "$lf\libfuzzer-dotnet-windows.exe" `
       -Uri 'https://github.com/Metalnem/libfuzzer-dotnet/releases/download/v2025.05.02.0904/libfuzzer-dotnet-windows.exe'
   (Get-FileHash "$lf\libfuzzer-dotnet-windows.exe" -Algorithm SHA256).Hash
   ```

2. Install the instrumentation tool into the work folder, not globally:

   ```powershell
   dotnet tool install SharpFuzz.CommandLine --version 2.3.0 --tool-path "$env:TEMP\officeagent-fuzz\tools"
   ```

3. Build the seed corpus. This writes `$env:TEMP\officeagent-fuzz\seeds`:

   ```powershell
   ./tests/OfficeAgent.Fuzz/build-seeds.ps1
   ```

4. Publish, instrument and fuzz for an hour, continuing past findings:

   ```powershell
   ./tests/OfficeAgent.Fuzz/fuzz.ps1 -LibFuzzer "$lf\libfuzzer-dotnet-windows.exe" -MaxTotalTime 3600 -KeepGoing
   ```

   Without `-KeepGoing`, libFuzzer stops at the first finding. Without `-MaxTotalTime`, it runs
   until a finding or Ctrl+C. Add `-SkipBuild` to reuse the instrumented build. New corpus inputs
   go to `$env:TEMP\officeagent-fuzz\corpus`, and findings go to
   `$env:TEMP\officeagent-fuzz-findings`.

`fuzz.ps1` runs one libFuzzer worker. It passes `-timeout=10`, `-max_len=1048576`, because
libfuzzer-dotnet refuses larger inputs, and the [`ooxml.dict`](ooxml.dict) dictionary.

Keep the machine awake for the whole run. On a laptop in Modern Standby the processes stop, and
libFuzzer's clock does not: a campaign that slept reports hours of `time` with almost no new
executions, and an input that was running when the machine slept is saved as `slow-unit-*`. Count
the time up to the last line whose executions still grew.

### Fuzzing past a known finding

In `-KeepGoing` mode every crash ends the current job, so a shallow finding that most mutations hit
again keeps the fuzzer from getting further. While such a finding waits for its fix, set it aside
with `OFFICEAGENT_FUZZ_KNOWN`: entries separated by `;`, each an exception type, `@`, and text its
stack trace contains. A matching exception counts as handled; anything else is still a finding.
Name a frame large enough not to be inlined: once the JIT optimises a hot path, a small method
such as a one-line helper disappears from the stack trace, and the entry stops matching.

```powershell
$env:OFFICEAGENT_FUZZ_KNOWN = 'System.ArgumentOutOfRangeException@OpenXmlPartContainer.GetPartById'
./tests/OfficeAgent.Fuzz/fuzz.ps1 -LibFuzzer "$lf\libfuzzer-dotnet-windows.exe" -MaxTotalTime 3600 -KeepGoing
```

Leave it unset for any run whose result you report as clean.

## Reproduce and minimise a finding

Run one saved input. The output names the outcome and prints the exception, and the exit code is 1
for a finding. An input still running after 10 seconds is reported as a hang; on a busy machine,
set `OFFICEAGENT_FUZZ_TIMEOUT_SECONDS` higher before calling a timeout real. This works with or
without instrumentation:

```powershell
dotnet "$env:TEMP\officeagent-fuzz\bin\OfficeAgent.Fuzz.dll" --repro "$env:TEMP\officeagent-fuzz-findings\crash-<sha1>"
```

To get a debuggable build, run the uninstrumented one instead:

```powershell
dotnet build tests/OfficeAgent.Fuzz -c Release
dotnet tests/OfficeAgent.Fuzz/bin/Release/net10.0/OfficeAgent.Fuzz.dll --repro <file>
```

Shrink a finding before reporting it. `--minimize` re-writes the package with every entry stored,
then drops entries one at a time while the same exception still escapes from the same OfficeAgent
frame, and prints the entries that are left:

```powershell
dotnet tests/OfficeAgent.Fuzz/bin/Release/net10.0/OfficeAgent.Fuzz.dll --minimize `
    "$env:TEMP\officeagent-fuzz-findings\crash-<sha1>" "$env:TEMP\officeagent-fuzz-findings\minimised-<sha1>"
```

When the finding depends on the archive's own bytes, such as a damaged header or entry name, a
re-written package no longer reproduces it, and the input is kept as it is. libFuzzer's
`-minimize_crash=1` works through libfuzzer-dotnet too, but it removes bytes, and almost any byte
removed from a ZIP breaks the archive before the finding is reached, so on these inputs it rarely
gets smaller.

To check that every seed is still a result or a documented refusal, for example after changing
the ingestion path:

```powershell
dotnet tests/OfficeAgent.Fuzz/bin/Release/net10.0/OfficeAgent.Fuzz.dll --check "$env:TEMP\officeagent-fuzz\seeds"
```

## The seed corpus

[`build-seeds.ps1`](build-seeds.ps1) collects every `.docx`, `.pptx` and `.xlsx` under
`samples/documents` and `tests/OfficeAgent.Tests/Corpus`, so the corpus grows with the fixtures
and nothing extra is committed. Each fixture is written as it is (`orig-`) and again with every
ZIP entry stored uncompressed (`stored-`). A byte flip inside a deflated entry nearly always
breaks the deflate stream, so mutations of the original files mostly exercise the archive reader;
in a stored copy they land in the XML and reach the SDK and the format modules. Duplicates are
dropped, and the encrypted fixtures, which are not archives, are kept only as they are.

## What is instrumented, and why

SharpFuzz rewrites the assemblies so libFuzzer can see which branches an input reaches.
`fuzz.ps1` instruments the code that reads the input's bytes:

- `OfficeAgent.Abstractions`, `OfficeAgent.Core`, `OfficeAgent.Word`, `OfficeAgent.PowerPoint`
  and `OfficeAgent.Excel`;
- `DocumentFormat.OpenXml.Framework`, the SDK's parsing core: package and part loading,
  relationships, the XML reader and element loading;
- `System.IO.Packaging`, which reads the OPC structure, content types and relationship parts;
- `ExcelNumberFormat`, which Excel inspection runs on every cell's number format string.

`DocumentFormat.OpenXml.dll` itself is not instrumented by default. It holds the SDK's generated
schema classes, one per element type. Instrumenting it raised the features libFuzzer counts over
the seeds by about 70% but kept only 4 more of the 253 seeds as distinct, cut throughput to
about a third, and crowds the fixed 64 KiB coverage map that every assembly shares, hiding
OfficeAgent's own branches behind collisions. Pass `-InstrumentOpenXmlSchema` to include it.

`System.IO.Compression` and `System.Xml` ship in the shared runtime and cannot be instrumented.
They still run on every input; libFuzzer just cannot see inside them. The project targets
`net10.0`, so those are the .NET 10 versions.

## Why it is in the solution

`OfficeAgent.NET.sln` includes this project, so CI compiles it on every change and the target
cannot drift from the API it calls. It is an executable, not a test project, so `dotnet test`
skips it, and it is not packable. Like every project in the solution, it restores from its
committed `packages.lock.json`; refresh it with `dotnet restore OfficeAgent.NET.sln
--force-evaluate` after changing the SharpFuzz version. Building it does not fuzz anything:
fuzzing needs the steps above.

## Reporting a finding

File each finding as its own issue with the minimised input attached, unless it could be a
security problem: a hang, unbounded memory, or anything that reads, writes or reaches the network
outside the document. Report those privately as [SECURITY.md](../../SECURITY.md) describes, and do
not attach the input publicly.
