# Support and compatibility

## Getting help

Search the [documentation](docs/README.md) and [troubleshooting guide](docs/troubleshooting.md)
first. For a reproducible defect or documentation problem, open a
[GitHub issue](https://github.com/ilia-sokolov/OfficeAgent.NET/issues/new/choose) with the
OfficeAgent package version, operating system, document format, provider, operation, stable
error code, and a minimal sanitized reproduction.

Do not attach confidential documents, credentials, tokens, registration indexes, or tenant
identifiers. Use the private route in [SECURITY.md](SECURITY.md) for suspected vulnerabilities.

Community support is best-effort and has no response-time guarantee. Managed hosting and
commercial support are available from dotaction at
[contact@dotaction.io](mailto:contact@dotaction.io?subject=OfficeAgent.NET%20commercial%20support).

## Version policy

OfficeAgent follows Semantic Versioning. While the project is below 1.0, a minor release may
change a public API, JSON plan shape, configuration setting, or behavior. Patch releases should
remain compatible within their minor line except where a security or correctness fix cannot be
made safely without a behavior change. Such changes are called out in the changelog.

The NuGet packages and MCP Registry entry are released as one versioned set. Do not mix
OfficeAgent assemblies from different minor versions. Security fixes are provided for the
latest published minor version as described in [SECURITY.md](SECURITY.md#supported-versions).

## Runtime and platform compatibility

The abstractions, core, format, SharePoint, and Agent Framework libraries target
`netstandard2.0` and `net8.0`. `OfficeAgent.Rendering` and the standalone MCP tool target
`net8.0`. The build workflow compiles and tests on current GitHub-hosted Ubuntu and Windows
runners with the .NET 8 SDK. macOS is expected to work through .NET and is included in adoption
trials, but it is not currently part of the automated CI matrix.

OfficeAgent reads and writes OOXML `.docx`, `.pptx`, and `.xlsx` packages through the
documented operation set. Compatibility means the produced package opens in supported desktop
Office applications for the tested workflow; the engine does not provide Word or PowerPoint's
layout and pagination, calculate fields or Excel formulas, or render Office documents. Validate
representative documents in the Office application used by your reviewers before production
adoption.

## Release assurances

Every release is expected to provide:

- tests of the exact release tag on Ubuntu;
- the normal main-branch build and test matrix on Ubuntu and Windows;
- a direct and transitive NuGet vulnerability audit;
- deterministic NuGet packages, symbols, Source Link metadata, and a matching changelog;
- version alignment across the GitHub release, NuGet package, and MCP Registry entry.

NuGet trusted publishing authenticates the release workflow without a long-lived API key.
OfficeAgent packages are not currently author-signed, and releases do not currently include an
SBOM or a service-level commitment. These are explicit assurance limits rather than guarantees
provided by the package.
