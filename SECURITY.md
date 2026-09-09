# Security policy

## Supported versions

OfficeAgent.NET is pre-1.0. Security fixes are released for the latest published minor
version. Older minor versions and preview versions are unsupported; upgrade before reporting
or reproducing an issue when possible.

| Version | Security support |
| --- | --- |
| 0.7.x | Supported |
| 0.6.x and earlier | Unsupported |

This table is updated when a newer minor version is released.

## Report a vulnerability privately

Use [GitHub private vulnerability reporting](https://github.com/ilia-sokolov/OfficeAgent.NET/security/advisories/new).
If that route is unavailable, email [contact@dotaction.io](mailto:contact@dotaction.io?subject=OfficeAgent.NET%20security%20report).
Do not include a confidential customer document, production credential, access token,
SharePoint registration index, or tenant identifier. A minimal synthetic document and exact
package version are usually enough to reproduce an Office-processing issue.

Include the affected version, deployment mode, document format, impact, reproduction steps,
and any known workaround. Please allow coordinated remediation before public disclosure.

We aim to acknowledge a complete report within three business days and provide an initial
assessment within seven business days. Fix and disclosure timing depends on severity,
exploitability, and upstream dependencies. We will send an update at least every fourteen
days while a confirmed report remains unresolved.

## Security boundary

OfficeAgent validates structured document operations and bounds access through host-configured
document connections. It does not make an untrusted document safe to disclose, provide malware
scanning, or decide which user is allowed to use a hosted endpoint.

The open-source HTTP server does not authenticate callers or implement a caller-to-connection
authorization policy. A hosted deployment must put TLS, authentication, and a fail-closed
connection allow-list in front of it. SharePoint credentials authorize storage access; they do
not protect the public MCP endpoint. See the [hosted deployment boundary](docs/deployment.md#b2-put-authentication-in-front)
and [hosted deployment checklist](docs/deployment.md#pre-flight-checklist-for-a-hosted-deployment).

Filesystem roots, SharePoint permissions, process identity, secrets, logs, and retention are
controlled by the host. Limit each connection to the documents its callers need, keep secrets
outside prompts and configuration committed to source control, and avoid logging document
content or tokens.

## Release security

Pull requests and pushes to `main` build and test on Ubuntu and Windows. The release workflow
checks out and tests the exact tag, audits direct and transitive NuGet dependencies, creates
deterministic packages and symbols, and publishes through NuGet trusted publishing. Packages
are not currently author-signed and the release does not currently publish an SBOM. See
[Support and compatibility](SUPPORT.md) for the compatibility policy and current assurance
limits.
