# v0.8.0 acceptance corpus

This directory freezes the public v0.8.0 behavior used as the starting point for v0.9.0. All documents are fictional test fixtures generated from repository-owned inputs. They are not customer documents or independent adoption evidence.

`manifest.json` records each case's origin, permission, SHA-256, authoring tool, operation sequence, semantic expectations, protected package parts, refusal codes, and allowed transformations. `V08CorpusTests` verifies manifest uniqueness and completeness, fixture hashes, Open XML validity, declared semantics, preservation boundaries, and rejection without provider mutation.

## Ordinary verification

The ordinary corpus tests require neither network access nor Microsoft Office:

```powershell
dotnet test tests/OfficeAgent.Tests/OfficeAgent.Tests.csproj `
  --configuration Release `
  --filter "FullyQualifiedName~V08CorpusTests"
```

Regenerate fixtures only when deliberately updating the baseline. Generation writes to an explicit directory and never overwrites this corpus implicitly:

```powershell
$env:OFFICEAGENT_V08_CORPUS_OUTPUT = Join-Path $env:TEMP 'officeagent-v08-corpus'
dotnet test tests/OfficeAgent.Tests/OfficeAgent.Tests.csproj `
  --configuration Release `
  --filter "FullyQualifiedName~Generate_versioned_corpus"
```

Open XML package compression and relationship identifiers can vary between generations. Review the generated package shape and semantics, then update committed files and manifest hashes together. Never weaken a hash or semantic assertion merely to accept unexplained drift.

## Clean package-consumer recipe

This manual recipe validates the packed libraries from a disposable project. It is separate from ordinary unit tests because a clean NuGet restore can require a configured package source.

```powershell
$repository = (Resolve-Path '.').Path
$scratch = Join-Path $env:TEMP 'officeagent-v08-library-smoke'
$feed = Join-Path $scratch 'feed'
$consumer = Join-Path $scratch 'consumer'
New-Item -ItemType Directory -Force -Path $feed | Out-Null
dotnet pack "$repository\OfficeAgent.NET.sln" --configuration Release --output $feed
dotnet new console --output $consumer
dotnet add "$consumer\consumer.csproj" package OfficeAgent.Core `
  --version 0.8.0 `
  --no-restore
dotnet restore "$consumer\consumer.csproj" `
  "--source=$feed" `
  '--source=https://api.nuget.org/v3/index.json'
dotnet build "$consumer\consumer.csproj" --configuration Release --no-restore
```

Run the recipe in a new disposable directory or remove the previous directory before repeating it. Do not reuse a consumer whose assets file points at a source checkout.

## Clean MCP tool recipe

This manual recipe installs only from the freshly packed local feed into a disposable tool path. Starting the process in stdio mode is the minimum packaging smoke test; an MCP client should then initialize the server and list its tools.

```powershell
$repository = (Resolve-Path '.').Path
$scratch = Join-Path $env:TEMP 'officeagent-v08-mcp-smoke'
$feed = Join-Path $scratch 'feed'
$toolPath = Join-Path $scratch 'tool'
New-Item -ItemType Directory -Force -Path $feed,$toolPath | Out-Null
dotnet pack "$repository\OfficeAgent.NET.sln" --configuration Release --output $feed
dotnet tool install OfficeAgent.Mcp `
  --tool-path $toolPath `
  --version 0.8.0 `
  --add-source $feed `
  --ignore-failed-sources
& (Join-Path $toolPath 'officeagent-mcp.exe') --stdio
```

The installed-package recipes must be rerun for a release candidate. A successful source-tree test is not a substitute for an isolated package install.
