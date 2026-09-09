# Releasing

This is the reusable maintainer runbook for publishing an OfficeAgent.NET release. Create a
GitHub issue for release-specific scope, blockers, and approvals; keep this file independent
of any one version.

Publishing writes to GitHub, NuGet, and the MCP Registry. Work through the steps in order and
verify each public artifact before continuing. A NuGet version cannot be replaced after it is
published.

## 0. Prepare the release

Choose the version in `major.minor.patch` form. It must already be present in
[`Directory.Build.props`](../Directory.Build.props), both version fields in
[`server.json`](../server.json), and the first release heading in
[`CHANGELOG.md`](../CHANGELOG.md).

Bash:

```bash
VERSION="$(sed -n 's:.*<Version>\([^<]*\)</Version>.*:\1:p' Directory.Build.props | head -n 1)"
TAG="v${VERSION}"
printf '%s\n' "$TAG"
```

PowerShell:

```powershell
[xml]$officeAgentProps = Get-Content Directory.Build.props
$officeAgentVersion = $officeAgentProps.Project.PropertyGroup.Version |
  Where-Object { $_ } |
  Select-Object -First 1
$officeAgentTag = "v$officeAgentVersion"
$officeAgentTag
```

Then run the release build and repository checks:

```bash
dotnet build OfficeAgent.NET.sln --configuration Release
dotnet test OfficeAgent.NET.sln --no-build --configuration Release
python scripts/check_vulnerable_packages.py
python scripts/validate_docs.py
python scripts/validate_server_manifest.py
```

Update the changelog heading from `unreleased` to the release date. Remove any temporary
prerelease notice for this version from the README, then commit those changes. The tag must
point at this final release commit.

Run the release-only validator before tagging:

```bash
python scripts/validate_release.py "$TAG"
```

```powershell
python scripts/validate_release.py $officeAgentTag
```

The validator refuses a release whose tag, package version, MCP manifest, skill metadata, or
dated changelog heading disagree.

## 1. Tag and create the GitHub release

Create and push an annotated tag from the final release commit:

```bash
git tag -a "$TAG" -m "OfficeAgent.NET $VERSION"
git push origin "$TAG"
```

```powershell
git tag -a $officeAgentTag -m "OfficeAgent.NET $officeAgentVersion"
git push origin $officeAgentTag
```

Use the matching changelog section as the GitHub release body so the two descriptions do not
drift. Extract that section into a temporary file and inspect it before publishing.

Bash:

```bash
awk -v version="$VERSION" '
  $0 ~ "^## " version " " { capture=1 }
  capture && $0 ~ "^## " && $0 !~ "^## " version " " { exit }
  capture { print }
' CHANGELOG.md > release-notes.md

gh release create "$TAG" \
  --title "OfficeAgent.NET $VERSION" \
  --notes-file release-notes.md
```

PowerShell:

```powershell
$officeAgentSections = (Get-Content CHANGELOG.md -Raw) -split '(?m)^## '
$officeAgentReleaseNotes = $officeAgentSections |
  Where-Object { $_.StartsWith("$officeAgentVersion ") } |
  Select-Object -First 1
"## $officeAgentReleaseNotes" | Set-Content release-notes.md

gh release create $officeAgentTag `
  --title "OfficeAgent.NET $officeAgentVersion" `
  --notes-file release-notes.md
```

Publishing the GitHub release triggers [the publish workflow](../.github/workflows/publish.yml).
That workflow checks out the release tag, validates the metadata, packs the packages, pushes
the libraries before `OfficeAgent.Mcp`, and attaches `word-document-review.zip` to the same
release. It is the only NuGet publisher; do not publish the same version manually.

## 2. Verify GitHub and NuGet

Check the release-triggered workflow and wait for it to succeed:

```bash
gh run list --workflow publish.yml --limit 1
```

Verify all of the following before publishing the MCP Registry entry:

- the GitHub release exists at
  [github.com/ilia-sokolov/OfficeAgent.NET/releases](https://github.com/ilia-sokolov/OfficeAgent.NET/releases);
- all expected packages show the new version on NuGet;
- `word-document-review.zip` contains `word-document-review/SKILL.md`;
- the release notes render correctly and their links resolve;
- the global tool installs from NuGet on a clean machine.

Clean-machine check:

```bash
dotnet tool install --global OfficeAgent.Mcp --version "$VERSION"
officeagent-mcp --stdio
```

Use a disposable machine, container, or user profile when the maintainer workstation already
has another version installed.

## 3. Publish the MCP Registry entry

Publish only after NuGet exposes the matching `OfficeAgent.Mcp` version, because the registry
validates that the package exists.

```bash
mcp-publisher login github
mcp-publisher publish
```

Verify the public registry response and confirm that its version, package version,
description, configuration requirements, and repository URL match
[`server.json`](../server.json):

```text
https://registry.modelcontextprotocol.io/v0.1/servers/io.github.ilia-sokolov%2Fofficeagent/versions/latest
```

Run the post-release alignment check. It fails unless the latest GitHub release, NuGet package,
and MCP Registry entry all expose the requested version:

```bash
python scripts/check_distribution.py --version "$VERSION"
```

```powershell
python scripts/check_distribution.py --version $officeAgentVersion
```

The scheduled `distribution` workflow repeats this check weekly against the latest GitHub
release so later registry drift remains visible.

Do not announce the release while the registry still resolves to an older version.

## 4. Check repository discovery metadata

GitHub topics and the homepage are persistent repository settings rather than release
artifacts. Check them on every release because new capabilities may require updated discovery
terms.

```bash
gh api repos/ilia-sokolov/OfficeAgent.NET --jq '{description, homepage, topics}'
```

The description should state the broad outcome and the distinctive reliability property.
Topics should stay aligned with the package tags in
[`Directory.Build.props`](../Directory.Build.props). Suitable topics include `mcp`,
`model-context-protocol`, `openxml`, `docx`, `pptx`, `xlsx`, `word`, `powerpoint`, `excel`, `dotnet`, `csharp`,
`ai-agents`, `tracked-changes`, and `document-automation`.

Update settings when necessary:

```bash
gh repo edit ilia-sokolov/OfficeAgent.NET \
  --homepage "https://www.nuget.org/packages/OfficeAgent.Mcp" \
  --description "Structured Word, PowerPoint, and Excel automation for .NET and coding agents"
```

Add or remove topics deliberately rather than appending the full list on every release.

## 5. Validate adoption and announce

Run the [representative adoption trials](adoption-validation.md) against the published
artifacts. Record failures by workflow and fix repeated installation or documentation
problems before broad promotion.

Lead announcements with the broad outcome: coding agents and .NET applications can create
and edit Word documents and PowerPoint decks through structured, validated operations. Use
one concrete proof point for the intended audience:

- coding-agent users: create or update a document or deck through MCP;
- document-review users: preserve comments and native tracked changes in the sample contract;
- .NET users: run a typed plan without an MCP client or language model;
- enterprise users: show the relevant filesystem, SharePoint, or hosted integration boundary.

Link to the shortest relevant guide and to an output file the reader can verify in Word or
PowerPoint. Rotate proof points across releases so one workflow does not define the apparent
scope of the project.

## 6. Close the release

Complete the release issue with links to:

- the final commit and tag;
- the GitHub release and publish workflow;
- the NuGet packages;
- the MCP Registry version;
- clean-machine verification evidence;
- adoption results and known limitations carried into the next release.

Delete the local `release-notes.md` file if it was created only for publication. Add newly
discovered release problems to this runbook when they apply to future releases; keep
version-specific notes in the release issue.

## Distribution gaps outside the release

The `word-document-review` skill is packaged as a GitHub release asset, but it is not
automatically installed into agent skill directories. If a broadly used skill registry or
installer becomes available, treat submission there as a separate distribution task.

`glama.json` contains repository metadata and does not normally need a version update.
