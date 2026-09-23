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
one, and the changelog calls it out. Engine extensibility types marked `OFFICEAGENT001` are
excluded. The [compatibility policy](docs/compatibility.md) defines each promise and how CI
enforces it.

The NuGet packages and MCP Registry entry are released as one versioned set. Do not mix
OfficeAgent assemblies from different minor versions. Security fixes are provided for the
latest published minor version only, which from 1.0 is the latest 1.x minor; all 0.x versions
are unsupported. See [SECURITY.md](SECURITY.md#supported-versions).

A deprecated member keeps working until the next major version and is marked `[Obsolete]`, with
its replacement named, at least one minor release before it is removed. A change of supported
runtime is announced in the changelog before the release that makes it.

## Runtime and platform compatibility

The abstractions, core, format, SharePoint, and Agent Framework libraries target
`netstandard2.0` and `net8.0`. `OfficeAgent.Rendering` and the standalone MCP tool target
`net8.0`.

OfficeAgent supports running on .NET versions that Microsoft supports and that CI executes. Today
those are .NET 8 and .NET 10, the two long-term-support releases.
[Microsoft's support policy](https://dotnet.microsoft.com/platform/support/policy/dotnet-core)
ends .NET 8 support on 2026-11-10 and .NET 10 support on 2028-11-14. After 2026-11-10, run
OfficeAgent on .NET 10: a .NET 10 application consumes the `net8.0` assets unchanged, the MCP
tool rolls forward to the newest installed runtime (`RollForward=LatestMajor`), and the container
image runs on the .NET 10 runtime. A problem that reproduces only on a runtime Microsoft no
longer supports is not guaranteed a fix. .NET 9 is a short-term release, supported until
2026-11-10, and is not tested separately.

The build workflow compiles, tests, and runs an installed-package smoke on current
GitHub-hosted Ubuntu, Windows, and macOS runners with the .NET 8 SDK and runtime. A second leg,
on Ubuntu and Windows, builds with the same SDK and runs the test suite and the packaged smoke on
the .NET 10 runtime; a test fails that leg if the suite did not actually run on .NET 10. The smoke installs
the packed `officeagent-mcp` tool into an empty tool directory backed by a fresh package
cache and an explicit local-only feed, drives a create, inspect, find, preview, apply, and
export workflow over stdio, and builds a direct consumer that loads the Word, PowerPoint,
and Excel modules from package references.

Compilation and tested execution are different claims:

| Target | Compiled | Tested by CI |
| --- | --- | --- |
| `net8.0` on the .NET 8 runtime: Linux x64, Windows x64, macOS | Yes | Yes, on the GitHub-hosted runner architectures |
| `net8.0` assets on the .NET 10 runtime: Linux x64, Windows x64 | Yes | Yes |
| `net8.0` assets on the .NET 10 runtime: macOS | Yes | No. The .NET 10 leg runs on Ubuntu and Windows |
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
- version alignment across the GitHub release, NuGet package, and MCP Registry entry.

NuGet trusted publishing authenticates the release workflow without a long-lived API key.
NuGet.org adds a repository signature after upload, so the registry-distributed archive bytes can
differ from the original packages attached to GitHub. The repository signature identifies the
repository, not the package author. OfficeAgent packages are not currently author-signed and no
service-level commitment is provided. Checksums detect changed bytes but do not establish author
identity, authorization, or provenance by themselves. See the [release runbook](docs/releasing.md)
for the separate verification paths and their trust assumptions.
