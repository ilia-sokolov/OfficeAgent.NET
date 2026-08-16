# Contributing to OfficeAgent.NET

Thank you for considering a contribution. Bug reports, design discussion, and pull requests are welcome.

## Local development

```bash
# Prerequisite: a stable .NET 8 SDK. global.json selects the latest installed
# 8.0 feature band at or above 8.0.100.
dotnet restore OfficeAgent.NET.sln
dotnet build OfficeAgent.NET.sln --no-restore
dotnet test OfficeAgent.NET.sln --no-build
```

The library multi-targets `netstandard2.0;net8.0`. Tests run on `net8.0`; build all TFM legs locally before opening a PR.

## What we want PRs for

- Bug fixes with a regression test.
- New operations or format handlers, with tests under `tests/OfficeAgent.Tests/`,
  XML documentation for public members, and updates to the operation matrix and
  affected format guide.
- New document providers (`IDocumentProvider`) in a focused assembly under
  `src/`, with the filesystem and SharePoint providers as the security and
  concurrency baseline.
- Documentation improvements, samples, and operational guidance.

## What we don't want without discussion

- Breaking changes to `DocumentPlan`, `DocumentReference`, `Anchor`, or the JSON wire shapes. Open an issue first; pre-1.0 we still take these but want to talk through the implications.
- New direct dependencies in `OfficeAgent.Abstractions` or `OfficeAgent.Core`. Both are kept small on purpose.
- New `[Obsolete]` markers - pre-1.0 we delete deprecated API outright. Discuss before adding obsolete shims.

## Style

- C# 12; nullable enabled; XML docs on every public member.
- Prefer struct/record DTOs for plan-shaped objects, classes for handler implementations.
- No `// TODO` comments - open an issue instead.
- Tests use xUnit; helper workspaces live inside the test class.

Documentation changes should keep examples executable, use connection-relative
or cross-platform paths where practical, and update [the documentation
hub](docs/README.md). Run `git diff --check` and verify every changed link before
opening the pull request.

## Branching

- Work on feature branches.
- Squash on merge.

## License

By contributing, you agree your work will be licensed under the MIT License (see [LICENSE](LICENSE)).
