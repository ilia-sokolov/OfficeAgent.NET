# Install OfficeAgent agent skills

The v1.0.0 release workflow packages two optional agent skills as GitHub release assets:

| Skill | Use it for |
| --- | --- |
| `officeagent-integration` | Choosing packages and building direct .NET, MCP, or Microsoft Agent Framework integrations |
| `word-document-review` | Safely inspecting and editing an existing Word document after the integration is available |

A skill is guidance and bundled recipe material. Installing one does not install the .NET runtime, NuGet packages, `OfficeAgent.Mcp`, or an agent framework. Install only the skill that matches the work.

## Install a released skill

The examples pin release `v1.0.0`. Replace `$skillName` only with one of the names above, and keep the skill and OfficeAgent packages on the same release.

Bash:

```bash
release=v1.0.0
skillName=officeagent-integration
destination="$HOME/.codex/skills"
temporary="$(mktemp -d)"
archive="$temporary/$skillName.zip"
mkdir -p "$destination"
curl -fL "https://github.com/ilia-sokolov/OfficeAgent.NET/releases/download/$release/$skillName.zip" -o "$archive"
unzip -q "$archive" -d "$destination"
test -f "$destination/$skillName/SKILL.md"
```

PowerShell:

```powershell
$release = "v1.0.0"
$skillName = "officeagent-integration"
$destination = Join-Path $env:USERPROFILE ".codex\skills"
$archive = Join-Path $env:TEMP "$skillName.zip"
New-Item -ItemType Directory -Force $destination | Out-Null
Invoke-WebRequest `
  "https://github.com/ilia-sokolov/OfficeAgent.NET/releases/download/$release/$skillName.zip" `
  -OutFile $archive
Expand-Archive -Force $archive $destination
if (-not (Test-Path (Join-Path $destination "$skillName\SKILL.md"))) {
  throw "Skill installation did not produce the expected layout."
}
```

Use `~/.claude/skills` instead of `~/.codex/skills` for Claude Code. Install into both only when both clients need the skill. Restart the client after installation so it can rediscover skills.

The `officeagent-integration` skill includes a disposable recipe project. Its [recipe instructions](../skills/officeagent-integration/references/recipes.md) cover a tracked Word edit, template population, complete-plan comparison, and an unsupported-operation refusal. Install the required OfficeAgent packages separately as those instructions describe.

For MCP, install the tool separately and follow [deployment and client setup](deployment.md#option-a---local-stdio-for-claude-code-and-codex). Confirm OfficeAgent appears in the client's MCP tool list before relying on it.

## Remove a skill

Bash:

```bash
skillName=officeagent-integration
rm -r "$HOME/.codex/skills/$skillName"
test ! -e "$HOME/.codex/skills/$skillName"
```

PowerShell:

```powershell
$skillName = "officeagent-integration"
$installed = Join-Path $env:USERPROFILE ".codex\skills\$skillName"
Remove-Item -LiteralPath $installed -Recurse
if (Test-Path -LiteralPath $installed) {
  throw "Skill removal was incomplete."
}
```

Use the matching Claude Code directory when that is where the skill was installed. Removing a skill does not uninstall `OfficeAgent.Mcp` or remove application package references.

## Maintainer verification before publication

From the repository root, after a Release build and pack into `artifacts`:

```bash
python scripts/package_skills.py --output artifacts
python scripts/verify_integration_kit.py --artifacts artifacts
```

The verification uses an isolated temporary agent home. It checks archive layout, relative references, release-version pins, and uncertain-write guidance; reads the repository package version; restores a disposable consumer from locally packed packages with that exact version; runs all four recipes and the refusal assertion; removes the installed skill; and confirms removal. On the `v1.0.0` release tag that version must be `1.0.0`. The verifier does not install into the user's active agent directories, connect a live model, or perform native Office visual review.
