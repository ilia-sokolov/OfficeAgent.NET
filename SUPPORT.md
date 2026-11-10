# Support and compatibility

## Getting help

Search the [documentation](docs/README.md) and [troubleshooting guide](docs/troubleshooting.md)
first. For a reproducible defect or documentation problem, open a
[GitHub issue](https://github.com/ilia-sokolov/OfficeAgent.NET/issues/new/choose) with the
OfficeAgent package version, operating system, document format, provider, operation, stable
error code, and a minimal sanitized reproduction.

Do not attach confidential documents, credentials, tokens, registration indexes, or tenant
identifiers. Use the private route in [SECURITY.md](SECURITY.md) for suspected vulnerabilities.

Community support is best-effort: issues are read and triaged by the project maintainer,
Ilia Sokolov, in English, with no response-time guarantee. The project has one maintainer and no
backup; [maintenance and continuity](docs/maintenance.md) records what that means and how the
project can be carried on without private knowledge. The only response targets the project states
are for [private vulnerability reports](SECURITY.md#report-a-vulnerability-privately).

Managed hosting and commercial support are offered separately by dotaction at
[contact@dotaction.io](mailto:contact@dotaction.io?subject=OfficeAgent.NET%20commercial%20support).
Any such service is governed by its own agreement. It does not change the library's license, its
compatibility promise or this support policy, and nothing in this repository is a service-level
commitment.

## Version policy

OfficeAgent follows Semantic Versioning. From 1.0, a minor or patch release does not break the
public .NET API, the JSON an agent or host exchanges, configuration keys, or documented defaults;
breaking changes wait for the next major version. A security fix may tighten a check, never loosen
one. A bug fix may refuse an input that used to commit a silently wrong document, at validation,
before anything is written. The changelog calls out each, and the compatibility policy lists
[the behaviour changes in 1.3](docs/compatibility.md#behaviour-changes-in-13). Engine extensibility types marked `OFFICEAGENT001` are
excluded. The [compatibility policy](docs/compatibility.md) defines each promise and how CI
enforces it.

The NuGet packages, the `officeagent-mcp` PyPI package and the MCP Registry entry are released
as one versioned set, and the Claude Code plugin in `plugins/officeagent` pins the PyPI package
of its own version. Do not mix
OfficeAgent assemblies from different minor versions. Security fixes are provided for the
latest published minor version only, which from 1.0 is the latest 1.x minor; all 0.x versions
are unsupported. See [SECURITY.md](SECURITY.md#supported-versions).

A deprecated member keeps working until the next major version and is marked `[Obsolete]`, with
its replacement named, at least one minor release before it is removed. A target framework whose
runtime Microsoft still supports is removed only in a major version, announced in the changelog
before the release that makes it. A target whose runtime has left Microsoft support can be removed
in a minor release, and that release's changelog says so.

## Runtime and platform compatibility

The abstractions, core, format, storage provider (SharePoint, iManage Work, Google Drive, Box,
Azure Blob Storage, Amazon S3), and Agent Framework libraries target `netstandard2.0` and
`net10.0`. `OfficeAgent.Rendering` and the standalone MCP tool target `net10.0`.

OfficeAgent supports running on .NET versions that Microsoft supports and that CI executes. CI
executes .NET 10, the current long-term-support release.
[Microsoft's support policy](https://dotnet.microsoft.com/platform/support/policy/dotnet-core)
ended .NET 8 support on 2026-11-10 and ends .NET 10 support on 2028-11-14, so 1.3.0 dropped the
`net8.0` target. A .NET 8 application still compiles against the libraries' `netstandard2.0`
assets, but that runtime is neither supported nor tested, and `OfficeAgent.Rendering` and the MCP
tool need .NET 10. The MCP tool rolls forward to a newer major with `RollForward=LatestMajor`, and
the container image runs on the .NET 10 runtime. The standalone MCP server downloads contain the .NET 10
runtime they were built with, so a .NET security fix reaches them with the next OfficeAgent
release rather than through a runtime update. A problem that reproduces only on a runtime Microsoft no
longer supports is not guaranteed a fix. .NET 9 is a short-term release, supported until
2026-11-10, and is not tested separately.

The build workflow compiles, tests, and runs an installed-package smoke on current
GitHub-hosted Ubuntu, Windows, and macOS runners, built with the .NET 10 SDK and tested on the
.NET 10 runtime. The smoke installs
the packed `officeagent-mcp` tool into an empty tool directory backed by a fresh package
cache and an explicit local-only feed, drives a create, inspect, find, preview, apply, and
export workflow over stdio, and builds a direct consumer that loads the Word, PowerPoint,
and Excel modules from package references.

Compilation and tested execution are different claims:

| Target | Compiled | Tested by CI |
| --- | --- | --- |
| `net10.0` on the .NET 10 runtime: Linux x64, Windows x64, macOS | Yes | Yes, on the GitHub-hosted runner architectures |
| `netstandard2.0` assets on .NET 8 or .NET 9, which Microsoft no longer supports | Yes | No |
| `netstandard2.0` consumers (.NET Framework, Mono, Xamarin, Unity) | Yes | No. The libraries compile for it, but no runtime test executes there |
| Linux arm64, Windows arm64, and other architectures | Yes | No. Not covered by the hosted runners used here |
| Alpine and other musl distributions | Yes | No |

A row marked untested is not a statement that it fails. It means this project has no
automated evidence for it, so validate it yourself before depending on it.

OfficeAgent reads and writes OOXML `.docx`, `.pptx`, and `.xlsx` packages through the
documented operation set. Compatibility means the produced package opens in supported desktop
Office applications for the tested workflow; the engine does not provide Word or PowerPoint's
layout and pagination, calculate fields or Excel formulas, or render Office documents. Validate
representative documents in the Office application used by your reviewers before production
adoption.

## Release assurances

Every release is expected to provide:

- tests of the exact release tag on Ubuntu;
- the normal main-branch build, test, and packaged-artifact smoke matrix on Ubuntu, Windows, and macOS;
- a direct and transitive NuGet vulnerability audit;
- deterministic NuGet packages, symbols, Source Link metadata, and a matching changelog;
- SHA-256 manifests for original release files, machine-readable CycloneDX dependency inventories,
  and GitHub build and SBOM attestations bound to the release workflow, tag, and source commit;
- BuildKit SBOM and provenance attestations plus GitHub build provenance for the container digest;
- self-contained MCP server archives and Claude Desktop bundles, each smoke-tested on Linux x64 at
  release and covered by the same checksums, SBOMs, and attestations;
- PyPI wheels built from those archives after their checksums and build attestations are
  verified, with PyPI's own publish attestation for each wheel;
- version alignment across the GitHub release, NuGet package, PyPI package, and MCP Registry entry.

NuGet trusted publishing authenticates the release workflow without a long-lived API key.
NuGet.org adds a repository signature after upload, so the registry-distributed archive bytes can
differ from the original packages attached to GitHub. The repository signature identifies the
repository, not the package author. OfficeAgent packages are not currently author-signed and no
service-level commitment is provided. Checksums detect changed bytes but do not establish author
identity, authorization, or provenance by themselves. See the [release runbook](docs/releasing.md)
for the separate verification paths and their trust assumptions.
