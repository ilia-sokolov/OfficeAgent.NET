# Contributing to OfficeAgent.NET

Thank you for considering a contribution. Bug reports, design discussion, and pull requests are welcome.

## Local development

```bash
# Prerequisites: .NET 8 SDK (any 8.0.x; rollForward picks the latest installed).
dotnet build OfficeAgent.NET.sln
dotnet test  OfficeAgent.NET.sln
```

The library multi-targets `netstandard2.0;net8.0`. Tests run on `net8.0`; build all TFM legs locally before opening a PR.

## What we want PRs for

- Bug fixes with a regression test.
- New verbs / handlers for the Word module, with tests under `tests/OfficeAgent.Tests/` and an XML-doc comment that explains what the verb does.
- New document providers (`IDocumentProvider`) - please ship them in their own assembly under `src/OfficeAgent.Providers.<Name>/` with sample-grade hardening (see `FileSystemDocumentProvider` as the bar).
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

## Branching

- Work on feature branches.
- Squash on merge.

## License

By contributing, you agree your work will be licensed under the MIT License (see [LICENSE](LICENSE)).
