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
python -m unittest tests/test_agent_evaluation.py tests/test_package_skills.py tests/test_release_evidence.py -v
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
the libraries before `OfficeAgent.Mcp`, and attaches the original packages, symbol packages,
skill archives, CycloneDX SBOMs, `release-manifest.json`, and `SHA256SUMS` to the same release.
It also creates GitHub build provenance and package-specific SBOM attestations. It is the only
NuGet publisher; do not publish the same version manually.

The manual workflow path is intentionally fail-closed. Dispatch it from the release tag so
GitHub's signed workflow identity and the checked-out source describe the same ref:

```bash
gh workflow run publish.yml --ref "$TAG" -f version="$VERSION"
```

Before tagging, a maintainer can exercise the evidence generator against locally packed files.
The exact source ref must still be a release tag, even during this unpublished dry run:

```bash
dotnet tool restore
python scripts/package_skills.py --output artifacts
python scripts/release_evidence.py sbom --version "$VERSION" --artifacts artifacts
python scripts/release_evidence.py create \
  --version "$VERSION" \
  --source-ref "refs/tags/$TAG" \
  --source-sha "$(git rev-parse HEAD)" \
  --artifacts artifacts
python scripts/release_evidence.py verify \
  --repository ilia-sokolov/OfficeAgent.NET \
  --workflow .github/workflows/publish.yml \
  --source-ref "refs/tags/$TAG" \
  --source-sha "$(git rev-parse HEAD)" \
  --artifacts artifacts
```

## 2. Verify GitHub and NuGet

Check the release-triggered workflow and wait for it to succeed:

```bash
gh run list --workflow publish.yml --limit 1
```

Download the complete release asset set and ask the release API for the authoritative attachment
list. The verifier rejects altered bytes, an unexpected repository, workflow, tag or source SHA,
a missing SBOM mapping, and missing or extra attachments:

```bash
VERIFY_DIR="verification/$TAG"
mkdir -p "$VERIFY_DIR"
gh release download "$TAG" --repo ilia-sokolov/OfficeAgent.NET --dir "$VERIFY_DIR"
gh release view "$TAG" --repo ilia-sokolov/OfficeAgent.NET \
  --json assets --jq '.assets[].name' > "$VERIFY_DIR/release-assets.txt"
SOURCE_SHA="$(git rev-list -n 1 "$TAG")"
python scripts/release_evidence.py verify \
  --repository ilia-sokolov/OfficeAgent.NET \
  --workflow .github/workflows/publish.yml \
  --source-ref "refs/tags/$TAG" \
  --source-sha "$SOURCE_SHA" \
  --release-asset-list "$VERIFY_DIR/release-assets.txt" \
  --artifacts "$VERIFY_DIR"
```

Then verify GitHub's signed provenance and SBOM claims for each downloaded product artifact.
For example:

```bash
gh attestation verify "$VERIFY_DIR/OfficeAgent.Core.$VERSION.nupkg" \
  --repo ilia-sokolov/OfficeAgent.NET \
  --signer-workflow ilia-sokolov/OfficeAgent.NET/.github/workflows/publish.yml \
  --source-ref "refs/tags/$TAG" \
  --source-digest "$SOURCE_SHA"
gh attestation verify "$VERIFY_DIR/OfficeAgent.Core.$VERSION.nupkg" \
  --repo ilia-sokolov/OfficeAgent.NET \
  --signer-workflow ilia-sokolov/OfficeAgent.NET/.github/workflows/publish.yml \
  --source-ref "refs/tags/$TAG" \
  --source-digest "$SOURCE_SHA" \
  --predicate-type https://cyclonedx.org/bom
```

Repeat both commands for every `.nupkg`, `.snupkg`, and skill archive. The manifest maps each
artifact to its aggregate SBOM and, for multi-targeted packages, separate `netstandard2.0` and
`net8.0` inventories. Missing license values mean upstream NuGet metadata did not supply them;
they are not inferred.

Verify all of the following before publishing the MCP Registry entry:

- the GitHub release exists at
  [github.com/ilia-sokolov/OfficeAgent.NET/releases](https://github.com/ilia-sokolov/OfficeAgent.NET/releases);
- all expected packages show the new version on NuGet;
- `word-document-review.zip` contains `word-document-review/SKILL.md`;
- `officeagent-integration.zip` contains `officeagent-integration/SKILL.md`, its recipe reference, and its recipe assets;
- the release notes render correctly and their links resolve;
- the global tool installs from NuGet on a clean machine.

The `.nupkg` and `.snupkg` files attached to GitHub are the exact pre-upload build outputs covered
by `SHA256SUMS` and GitHub attestations. NuGet.org repository-signs packages during ingestion, so a
package downloaded from NuGet can have different archive bytes. Do not compare its whole-file hash
to `SHA256SUMS`. Instead, verify the registry-distributed package's repository signature and
package identity:

```bash
curl --fail --location \
  "https://api.nuget.org/v3-flatcontainer/officeagent.core/$VERSION/officeagent.core.$VERSION.nupkg" \
  --output "OfficeAgent.Core.$VERSION.nuget-org.nupkg"
dotnet nuget verify "OfficeAgent.Core.$VERSION.nuget-org.nupkg" --all
```

The repository signature proves NuGet.org distribution, not author identity. The matching GitHub
original plus its attestation establishes the workflow/source identity of the pre-upload bytes.
Checksums detect byte changes but are not signatures or authorization.

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

## 4. Verify the container image

The container workflow builds only from the matching tag. BuildKit publishes an in-registry SBOM
and maximum-mode provenance for the immutable image digest, and GitHub records a separate signed
build attestation. Verify the version tag by digest, never through the mutable `latest` tag:

```bash
IMAGE="ghcr.io/ilia-sokolov/officeagent-mcp:$VERSION"
docker buildx imagetools inspect "$IMAGE"
docker buildx imagetools inspect "$IMAGE" --format '{{ json .SBOM }}'
docker buildx imagetools inspect "$IMAGE" --format '{{ json .Provenance }}'
gh attestation verify "oci://$IMAGE" \
  --repo ilia-sokolov/OfficeAgent.NET \
  --signer-workflow ilia-sokolov/OfficeAgent.NET/.github/workflows/publish-image.yml \
  --source-ref "refs/tags/$TAG" \
  --source-digest "$SOURCE_SHA"
```

## 5. Check repository discovery metadata

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

## 6. Validate adoption and announce

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

## 7. Close the release

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

The `word-document-review` and `officeagent-integration` skills are packaged as GitHub release
assets, but they are not automatically installed into agent skill directories. If a broadly
used skill registry or installer becomes available, treat submission there as a separate
distribution task.

`glama.json` contains repository metadata and does not normally need a version update.
