# Security policy

## Supported versions

Security fixes are released for the latest published minor version only. From 1.0 that is the
latest 1.x minor: a 1.x minor release never breaks the public API or the wire contract (see
[compatibility](docs/compatibility.md)), so the supported way to receive a fix is to upgrade within
1.x. Older minor versions and preview versions are unsupported; upgrade before reporting or
reproducing an issue when possible.

| Version | Security support |
| --- | --- |
| 1.0.x | Supported |
| 0.9.x and earlier | Unsupported |

Every 0.x version is unsupported. This table is updated with each minor release, and a test
checks that it names the version the repository builds.

Fixes also depend on the .NET runtime underneath. A problem that reproduces only on a runtime
Microsoft no longer supports is not guaranteed a fix; see
[runtime support](SUPPORT.md#runtime-and-platform-compatibility).

## Report a vulnerability privately

Use [GitHub private vulnerability reporting](https://github.com/ilia-sokolov/OfficeAgent.NET/security/advisories/new).
If that route is unavailable, email [contact@dotaction.io](mailto:contact@dotaction.io?subject=OfficeAgent.NET%20security%20report).
Do not include a confidential customer document, production credential, access token,
SharePoint registration index, or tenant identifier. A minimal synthetic document and exact
package version are usually enough to reproduce an Office-processing issue.

Include the affected version, deployment mode, document format, impact, reproduction steps,
and any known workaround. Please allow coordinated remediation before public disclosure.
Reports are handled in English.

Reports are triaged by the project maintainer, Ilia Sokolov, who owns these targets:

| Step | Target |
| --- | --- |
| Acknowledge a complete report | Within three business days |
| Initial assessment | Within seven business days |
| Status update while a confirmed report is unresolved | At least every fourteen days |

These are the maintainer's targets, not a contractual service level. Fix and disclosure timing
depends on severity, exploitability, and upstream dependencies. The project has one maintainer
and no backup responder; if a target is missed, the report is still handled, and the delay is
stated in the next update.

In scope: the published OfficeAgent packages, the MCP server and its container image, and the
reference deployments in this repository used as documented. Out of scope: the security of a host
application built on OfficeAgent, a hosted endpoint that lacks the authentication the
[security boundary](#security-boundary) requires, and vulnerabilities in dependencies that are
not reachable through OfficeAgent (report those upstream).

## Security boundary

OfficeAgent validates structured document operations and bounds access through host-configured
document connections. It does not make an untrusted document safe to disclose, provide malware
scanning, or decide which user is allowed to use a hosted endpoint.

The standalone open-source HTTP server does not authenticate callers. A hosted deployment must
add TLS, authentication, and a fail-closed `IConnectionAccessPolicy`. The
[HostedGateway reference](samples/HostedGateway/) demonstrates two authenticated principals,
disjoint connections, and trusted receipt actors, but it is not a production identity system.
SharePoint credentials authorize storage access; they do not protect the public MCP endpoint.
See the [hosted deployment boundary](docs/deployment.md#b2-put-authentication-in-front)
and [hosted deployment checklist](docs/deployment.md#pre-flight-checklist-for-a-hosted-deployment).

Filesystem roots, SharePoint permissions, process identity, secrets, logs, and retention are
controlled by the host. Limit each connection to the documents its callers need, keep secrets
outside prompts and configuration committed to source control, and avoid logging document
content or tokens.

## Release security

Pull requests and pushes to `main` build and test on Ubuntu, Windows and macOS, and a separate
leg runs the tests and the packaged smoke on the .NET 10 runtime. The release workflow checks out
and tests the exact tag, audits direct and transitive NuGet dependencies, creates deterministic
packages and symbols, publishes CycloneDX dependency inventories with GitHub build and SBOM
attestations, and publishes through NuGet trusted publishing. Packages are not author-signed. See
[Support and compatibility](SUPPORT.md#release-assurances) for what each release provides and its
limits.
