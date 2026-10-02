# Contributing to OfficeAgent.NET

Thank you for considering a contribution. Bug reports, design discussion, and pull requests are welcome.

## Finding something to work on

- [`good first issue`](https://github.com/ilia-sokolov/OfficeAgent.NET/labels/good%20first%20issue)
  issues are small and fully specified: each names the files to change, the test or check that
  proves it, and the documentation to update.
- [`help wanted`](https://github.com/ilia-sokolov/OfficeAgent.NET/labels/help%20wanted) issues
  are larger, and some need more context about the engine.
- [`corpus`](https://github.com/ilia-sokolov/OfficeAgent.NET/labels/corpus) issues need no code:
  a real document that breaks naive editing, or a run of the native Office checks on your Office
  build, is a valuable contribution on its own.
- Questions and early ideas go to
  [Discussions](https://github.com/ilia-sokolov/OfficeAgent.NET/discussions), not issues.

An issue labelled `needs-decision` is waiting for a maintainer decision, recorded in the issue.
Please don't start on it until that decision is there.

## Claiming an issue

Comment on the issue to ask for it. The first person to ask is assigned, and each person holds
one open claim at a time. If no pull request has appeared after 14 days, the claim lapses and the
issue is freed, unless you have said you are still working on it. Saying "I can't get to this"
is always fine.

## Your pull request

- The first time you open a pull request here, CI waits for a maintainer to approve the run. That
  is a GitHub rule for first-time contributors, not a judgement on the change.
- The `build` and `analyze (csharp)` checks must pass before a pull request can merge.
- Expect a first review within two working days.
- Documentation changes are checked with `python scripts/validate_docs.py` (Python 3), which
  checks links and code fences.

## Local development

```bash
# Prerequisite: a stable .NET 8 SDK. global.json selects the latest installed
# 8.0 feature band at or above 8.0.100.
dotnet restore OfficeAgent.NET.sln
dotnet build OfficeAgent.NET.sln --no-restore
dotnet test OfficeAgent.NET.sln --no-build
```

The library multi-targets `netstandard2.0;net8.0`. Tests run on `net8.0`; build all TFM legs locally before opening a PR. CI also runs the tests on the .NET 10 runtime; to do the same locally, install the .NET 10 runtime and run `dotnet test` with `DOTNET_ROLL_FORWARD=LatestMajor`.

The [maintenance guide](docs/maintenance.md) describes the checks, the release process and the systems a maintainer needs.

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

- Breaking changes to the public .NET API, the JSON wire shapes, configuration keys or documented defaults. From 1.0 these wait for the next major version; see [compatibility](docs/compatibility.md). CI fails a pull request whose change shows up in `docs/csharp-api.md` or `docs/wire-contract.md` until the baseline is regenerated, and a regenerated baseline needs a recorded decision.
- New direct dependencies in `OfficeAgent.Abstractions` or `OfficeAgent.Core`. Both are kept small on purpose.
- Removing or renaming public API. Mark it `[Obsolete]` with its replacement named instead, and note it in the changelog; it is removed only in the next major version, as [the deprecation policy](docs/compatibility.md#deprecation) states.

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
