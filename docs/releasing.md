# Releasing

Everything in the repository is prepared for 0.7.0. The steps below are the ones that
publish, and they are deliberately left for the maintainer — each writes to a registry or to
repository settings, and none can be undone quietly.

Nothing here has been run. Work through it in order; each step says how to verify it.

## 0. Preconditions

```bash
dotnet build OfficeAgent.NET.sln --configuration Release
dotnet test OfficeAgent.NET.sln --configuration Release
```

Confirm the version is the one you mean to ship — it is set once and flows everywhere:

```bash
grep '<Version>' Directory.Build.props        # 0.7.0
grep '"version"' server.json                   # 0.7.0, twice
```

Change the date line in [CHANGELOG.md](../CHANGELOG.md) from `unreleased` to the release
date, and commit before tagging so the tag points at the final text.

## 1. Tag and publish the GitHub release

The body is the 0.7.0 section of [CHANGELOG.md](../CHANGELOG.md) — paste it rather than
writing new text, so the changelog and the release cannot drift.

```bash
git tag -a v0.7.0 -m "OfficeAgent.NET 0.7.0"
git push origin v0.7.0
```

Extract the section to a file, then create the release from it — two steps rather than one
pipeline, so this works the same in PowerShell as in bash:

```bash
sed -n '/^## 0.7.0/,/^## 0.6.0/p' CHANGELOG.md | sed '$d' > release-notes.md
gh release create v0.7.0 --title "OfficeAgent.NET 0.7.0" --notes-file release-notes.md
```

```powershell
(Get-Content CHANGELOG.md -Raw) -split '(?m)^## ' |
  Where-Object { $_.StartsWith('0.7.0') } |
  ForEach-Object { "## $_" } | Set-Content release-notes.md
gh release create v0.7.0 --title "OfficeAgent.NET 0.7.0" --notes-file release-notes.md
```

Publishing the GitHub release triggers `.github/workflows/publish.yml`. That workflow is
the single NuGet publisher: it checks out the release tag, packs all seven packages, pushes
the six libraries first, and pushes `OfficeAgent.Mcp` last through NuGet trusted publishing.
Do not push the same release manually.

## 2. Verify NuGet publication

Check the release-triggered workflow and wait for it to succeed:

```bash
gh run list --workflow publish.yml --limit 1
```

Then install the package version from a clean machine and check that the server starts:

```bash
dotnet tool install --global OfficeAgent.Mcp --version 0.7.0
officeagent-mcp --stdio --config ./officeagent.json
```

Verify that the release appears at
[github.com/ilia-sokolov/OfficeAgent.NET/releases](https://github.com/ilia-sokolov/OfficeAgent.NET/releases).

**Not reversible.** A pushed NuGet version cannot be replaced, only delisted.

Check the rendered notes before announcing: the changelog links are repository-relative and
GitHub resolves them from the release page, but the sample `.docx` link is worth clicking.

## 3. MCP Registry

[`server.json`](../server.json) is already consistent with the NuGet package — same version,
and its `identifier` is `OfficeAgent.Mcp`. Publish **after** NuGet, because the registry
validates that the package exists.

```bash
mcp-publisher login github
mcp-publisher publish
```

Verify at `https://registry.modelcontextprotocol.io/v0/servers?search=officeagent` that the
returned version is 0.7.0 and the environment-variable descriptions match `server.json`.

## 4. Repository topics and homepage

Neither is set today, and both are how the repository is found from outside. These are
settings rather than releases, so they can be changed freely afterwards.

```bash
gh repo edit ilia-sokolov/OfficeAgent.NET \
  --homepage "https://www.nuget.org/packages/OfficeAgent.Mcp" \
  --add-topic mcp \
  --add-topic model-context-protocol \
  --add-topic mcp-server \
  --add-topic openxml \
  --add-topic docx \
  --add-topic pptx \
  --add-topic word \
  --add-topic powerpoint \
  --add-topic dotnet \
  --add-topic csharp \
  --add-topic ai-agents \
  --add-topic llm-tools \
  --add-topic tracked-changes \
  --add-topic document-automation \
  --add-topic agent-skills
```

These match the NuGet `PackageTags` in [Directory.Build.props](../Directory.Build.props), so
someone arriving from either side sees the same vocabulary. GitHub allows 20 topics; this
uses 15, leaving room.

On the homepage: pointing at the NuGet tool page sends visitors somewhere they can act. If
documentation is published to a site later, that becomes the better target.

Verify:

```bash
gh api repos/ilia-sokolov/OfficeAgent.NET --jq '{homepage, topics}'
```

## 5. Announce

Only after the registry entry resolves and the tool installs from a clean machine. The
quickest honest demo is the one in the README: install, drop the sample contract in a folder,
ask for the payment terms to change, open the result in Word and see a tracked change.

## Still open

- The `word-document-review` skill is not published to any skill registry; users copy it out
  of the repository. If a registry becomes the norm, that is a separate submission.
- `glama.json` carries only the maintainer list and needs nothing per release.
