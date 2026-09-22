# Maintenance and continuity

Who maintains OfficeAgent.NET, which systems a release depends on, and how someone else could
build, verify, release or carry on the project using only this repository. It states capacity as
it is; it is not a service commitment. Support terms are in [SUPPORT.md](../SUPPORT.md) and
security terms in [SECURITY.md](../SECURITY.md).

## Maintainers

The project has one maintainer, Ilia Sokolov, who triages issues and security reports, reviews
pull requests, and publishes releases. There is no second maintainer or backup responder. The only
stated response times are the targets for
[private vulnerability reports](../SECURITY.md#report-a-vulnerability-privately); everything else
is best-effort.

## Systems a release depends on

| System | Owner today | Used for |
| --- | --- | --- |
| GitHub repository `ilia-sokolov/OfficeAgent.NET` | Ilia Sokolov (owner and sole collaborator) | Source, issues, private vulnerability reports, CI, releases, the `nuget` environment (Ilia Sokolov is its required reviewer) |
| NuGet.org account `ilia-sokolov` | Ilia Sokolov | Owns every `OfficeAgent.*` package ID; Trusted Publishing policy for `publish.yml` in environment `nuget` |
| GitHub Container Registry `ghcr.io/ilia-sokolov/officeagent-mcp` | Ilia Sokolov's GitHub account | The MCP server image, published by the release workflow |
| MCP Registry namespace `io.github.ilia-sokolov` | Ilia Sokolov's GitHub account | The `officeagent` server entry |
| `contact@dotaction.io` | dotaction | Fallback security reporting route and commercial enquiries |

Each of these is tied to a personal account. Package IDs, the container path and the registry
namespace cannot move to another person without the owner acting: NuGet.org package ownership is
granted by an existing owner, and the other two follow the GitHub account. The
[release runbook](releasing.md#prerequisites) lists exactly what each step needs.

## If the maintainer is unavailable

The code is MIT-licensed, and everything needed to build, test, package, verify and release it is
in this repository. Nothing depends on a private script, machine or secret: publishing uses NuGet
Trusted Publishing and the workflow token, not a stored API key.

What cannot be carried on without the maintainer is publishing under the existing names. A
successor who is not given ownership would fork the repository and publish under new package IDs,
a new container path and a new registry namespace, and users would switch to those names.

## Building and verifying

- **Toolchain.** [`global.json`](../global.json) pins the .NET 8 SDK (8.0.100 or a later 8.0
  feature band). CI installs it with `actions/setup-dotnet`.
- **Reproducible output.** Builds are deterministic, with Source Link and embedded untracked
  sources, and CI builds set `ContinuousIntegrationBuild`. Package references name exact
  versions, which NuGet resolves to that version while it is available; there are no lock files.
- **The checks.** [CONTRIBUTING](../CONTRIBUTING.md#local-development) has the local commands.
  CI runs the same build, the generated contract gate, the tests on the .NET 8 runtime on Ubuntu,
  Windows and macOS and on the .NET 10 runtime on Ubuntu and Windows, the packaged-artifact smoke,
  and the renderer sandbox reference tests.
- **Native Office evidence.** The [native compatibility](native-compatibility.md) harness needs a
  Windows machine with desktop Office and is run per release, not in CI.

## Dependencies and runtimes

- **Vulnerabilities.** Every build on Ubuntu and every release runs
  `scripts/check_vulnerable_packages.py`, which fails on any reported direct or transitive
  vulnerability that `security/dependency-audit-allowlist.json` does not document. Dependency
  updates are made by hand. No automated update service is configured.
- **Runtimes.** OfficeAgent follows Microsoft's
  [.NET support policy](https://dotnet.microsoft.com/platform/support/policy/dotnet-core). .NET 8
  support ends on 2026-11-10 and .NET 10 support on 2028-11-14. The libraries' `net8.0` assets run
  on .NET 10, which CI tests, and the container image runs on the .NET 10 runtime. The build SDK
  pin and the [renderer reference](../deploy/renderer/README.md) base image are still .NET 8; moving
  them changes no package a consumer uses.

## Releases and recovery

The [release runbook](releasing.md) is the complete procedure: prerequisites, tagging, the single
gated publish workflow, verification of every artifact, the MCP Registry entry, and container
checks. It also covers recovery:

- A NuGet version cannot be replaced. A bad release is fixed by a new patch version, never by
  republishing.
- If NuGet publishes but the container stage fails, the gated `publish-container.yml` workflow
  rebuilds the image from the immutable tag without touching NuGet.
- The weekly `distribution` workflow reports drift between the GitHub release, NuGet and the MCP
  Registry.

## Contributors

[CONTRIBUTING](../CONTRIBUTING.md) covers local development, what changes need discussion first,
and style. The [compatibility policy](compatibility.md) is the rule for anything public: a change
that alters `docs/csharp-api.md` or `docs/wire-contract.md` is a contract change, which CI makes
visible on every pull request.
