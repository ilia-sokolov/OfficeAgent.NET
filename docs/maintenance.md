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
| PyPI project `officeagent-mcp` | Ilia Sokolov | The Python wheels; a trusted publisher for `publish-python.yml` in environment `pypi` |
| MCP Registry namespace `io.github.ilia-sokolov` | Ilia Sokolov's GitHub account | The `officeagent` server entry |
| Anthropic plugin directory | Not listed | The Claude plugin in `plugins/officeagent`; submission is held until Anthropic grants Verified organization status |
| `contact@dotaction.io` | dotaction | Fallback security reporting route and commercial enquiries |

Each of these is tied to a personal account. Package IDs, the container path and the registry
namespace cannot move to another person without the owner acting: NuGet.org package ownership is
granted by an existing owner, and the other two follow the GitHub account. The
[release runbook](releasing.md#prerequisites) lists exactly what each step needs.

## If the maintainer is unavailable

The code is MIT-licensed, and everything needed to build, test, package, verify and release it is
in this repository. Nothing depends on a private script, machine or secret: publishing uses NuGet
Trusted Publishing, PyPI trusted publishing and the workflow token, not a stored API key.

What cannot be carried on without the maintainer is publishing under the existing names. A
successor who is not given ownership would fork the repository and publish under new package IDs,
a new container path and a new registry namespace, and users would switch to those names.

## Building and verifying

- **Toolchain.** [`global.json`](../global.json) pins the .NET 10 SDK (10.0.100 or a later 10.0
  feature band). CI installs it with `actions/setup-dotnet`, and the tests run on its .NET 10
  runtime. The libraries target netstandard2.0 and net10.0.
- **Reproducible output.** Builds are deterministic, with Source Link and embedded untracked
  sources, and CI builds set `ContinuousIntegrationBuild`. Package references name exact
  versions, and every project in `OfficeAgent.NET.sln` has a committed `packages.lock.json`
  (`RestorePackagesWithLockFile`). CI restores in locked mode, so a restore cannot pick up a
  different transitive version.
  After changing a reference, refresh the lock files with
  `dotnet restore OfficeAgent.NET.sln --force-evaluate`.
- **The checks.** [CONTRIBUTING](../CONTRIBUTING.md#local-development) has the local commands.
  CI runs the documentation, MCP manifest and Python tooling checks, the same build, the
  generated contract gate, the tests on the .NET 10 runtime on Ubuntu, Windows and macOS, the
  packaged-artifact smoke, the standalone server
  and Python wheel smoke, the container health check, and the renderer sandbox reference tests.
- **Native Office evidence.** The [native compatibility](native-compatibility.md) harness needs a
  Windows machine with desktop Office and is run per release, not in CI.

## Dependencies and runtimes

- **Vulnerabilities.** Every build on Ubuntu and every release runs
  `scripts/check_vulnerable_packages.py`, which fails on any reported direct or transitive
  vulnerability that `security/dependency-audit-allowlist.json` does not document.
- **Updates.** [Dependabot](../.github/dependabot.yml) opens weekly grouped pull requests for
  NuGet packages, GitHub Actions, the container base images and the CI Python packages in
  `scripts`. It leaves major NuGet and container versions alone; take those deliberately. The
  Claude plugin's `uv.lock` and the version pins in `deploy` move with the release, as the
  [release runbook](releasing.md) describes.
- **Runtimes.** OfficeAgent follows Microsoft's
  [.NET support policy](https://dotnet.microsoft.com/platform/support/policy/dotnet-core). .NET 8
  support ended on 2026-11-10, and 1.3.0 replaced the `net8.0` target with `net10.0`; .NET 10
  support ends on 2028-11-14. The container image runs on the .NET 10 runtime. The
  [renderer reference](../deploy/renderer/README.md) publishes its worker self-contained on .NET 10
  over a plain Debian base image.

### What the packages depend on

Every package reference in `src`, from its project file. There is no central
`Directory.Packages.props`; each project names its own versions.

| Package | Third-party and Microsoft packages it references |
| --- | --- |
| `OfficeAgent.Abstractions` | `System.Text.Json` 8.0.6 (netstandard2.0 only) |
| `OfficeAgent.Core` | `DocumentFormat.OpenXml` 3.5.1, `Microsoft.Extensions.DependencyInjection.Abstractions` 8.0.2, `Microsoft.Extensions.Logging.Abstractions` 8.0.3, `System.Diagnostics.DiagnosticSource` 8.0.1; `Microsoft.Bcl.TimeProvider` 8.0.1 on netstandard2.0 |
| `OfficeAgent.Word`, `OfficeAgent.PowerPoint` | `DocumentFormat.OpenXml` 3.5.1, `Microsoft.Extensions.DependencyInjection.Abstractions` 8.0.2; `Microsoft.Bcl.TimeProvider` 8.0.1 on netstandard2.0 |
| `OfficeAgent.Excel` | `DocumentFormat.OpenXml` 3.5.1, **`ExcelNumberFormat` 1.1.0** (new in 1.3, MIT), `Microsoft.Extensions.DependencyInjection.Abstractions` 8.0.2; `Microsoft.Bcl.TimeProvider` 8.0.1 on netstandard2.0 |
| `OfficeAgent.Excel.Calculation` | `ClosedXML` `[0.105.1, 0.106.0)`, `DocumentFormat.OpenXml` 3.5.1, `Microsoft.Extensions.DependencyInjection.Abstractions` 8.0.2 |
| `OfficeAgent.AgentFramework` | `Microsoft.Extensions.AI` and `Microsoft.Extensions.AI.Abstractions` 10.11.0 |
| `OfficeAgent.Rendering` | None; it starts LibreOffice and Poppler as external processes |
| `OfficeAgent.SharePoint`, `OfficeAgent.IManage`, `OfficeAgent.GoogleDrive`, `OfficeAgent.Box` | `Microsoft.Extensions.DependencyInjection.Abstractions` 8.0.2; `System.Text.Json` 8.0.6 and `System.Net.Http` 4.3.4 on netstandard2.0. No vendor SDK: each calls its service's REST API |
| `OfficeAgent.AzureBlob` | **`Azure.Storage.Blobs` 12.30.1** and **`Azure.Identity` 1.21.0** (new in 1.3), `Microsoft.Extensions.DependencyInjection.Abstractions` 10.0.9, which matches what `Azure.Core` brings |
| `OfficeAgent.AmazonS3` | **`AWSSDK.S3` 4.0.104.4** (new in 1.3), `Microsoft.Extensions.DependencyInjection.Abstractions` 8.0.2; `System.Text.Json` 8.0.6 on netstandard2.0 |
| `OfficeAgent.Mcp` | `ModelContextProtocol.AspNetCore` 1.4.1, `Microsoft.Extensions.AI` 10.11.0, the ASP.NET Core shared framework, and every package above except `OfficeAgent.Rendering` |

Projects that are not packages add their own: the fuzz target uses `SharpFuzz` 2.3.0, and the
samples use Microsoft Agent Framework, Semantic Kernel, Azure OpenAI and ClosedXML. The PyPI
wheels and the Claude plugin's launcher carry no Python dependencies beyond `officeagent-mcp`
itself.

## Releases and recovery

The [release runbook](releasing.md) is the complete procedure: prerequisites, tagging, the single
gated publish workflow, verification of every artifact, the MCP Registry entry, and container
checks. It also covers recovery:

- Neither a NuGet version nor a PyPI version can be replaced. A bad release is fixed by a new
  patch version, never by republishing.
- If NuGet publishes but the container stage fails, the gated `publish-container.yml` workflow
  rebuilds the image from the immutable tag without touching NuGet.
- The weekly `distribution` workflow reports drift between the GitHub release, NuGet and the MCP
  Registry.

## Contributors

[CONTRIBUTING](../CONTRIBUTING.md) covers local development, what changes need discussion first,
and style. The [compatibility policy](compatibility.md) is the rule for anything public: a change
that alters `docs/csharp-api.md` or `docs/wire-contract.md` is a contract change, which CI makes
visible on every pull request.
