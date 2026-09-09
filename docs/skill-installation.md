# Install the Word review skill

The `word-document-review` skill teaches an agent how to inspect an existing `.docx`,
preserve review state, make explicit tracked edits, and recover from OfficeAgent error
codes. It complements the MCP server; it does not install or start the server itself.
The skill is optional and scoped to review tasks. Creating documents, making ordinary
direct edits, and working with PowerPoint do not require it.

## 1. Install the MCP server

Install the released tool:

```bash
dotnet tool install --global OfficeAgent.Mcp
```

The skill and expanded review workflow described on `main` require OfficeAgent 0.7.0. Until
that version is published, clone the repository, pack all projects into one local package
source, and install the tool from there:

```bash
git clone --depth 1 https://github.com/ilia-sokolov/OfficeAgent.NET.git officeagent-net
cd officeagent-net
dotnet pack OfficeAgent.NET.sln --configuration Release --output ./artifacts/local
dotnet tool install --global OfficeAgent.Mcp --version 0.7.0 --add-source ./artifacts/local
```

## 2. Install the skill

Starting with the 0.7.0 release, download `word-document-review.zip` from the latest GitHub
release. The archive contains the correctly named skill directory.

Bash:

```bash
mkdir -p ~/.claude/skills ~/.codex/skills
curl -L https://github.com/ilia-sokolov/OfficeAgent.NET/releases/latest/download/word-document-review.zip \
  -o /tmp/word-document-review.zip
unzip -q /tmp/word-document-review.zip -d ~/.claude/skills
unzip -q /tmp/word-document-review.zip -d ~/.codex/skills
```

PowerShell:

```powershell
$claudeSkills = Join-Path $env:USERPROFILE ".claude\skills"
$codexSkills = Join-Path $env:USERPROFILE ".codex\skills"
$skillArchive = Join-Path $env:TEMP "word-document-review.zip"
New-Item -ItemType Directory -Force $claudeSkills, $codexSkills | Out-Null
Invoke-WebRequest `
  https://github.com/ilia-sokolov/OfficeAgent.NET/releases/latest/download/word-document-review.zip `
  -OutFile $skillArchive
Expand-Archive -Force $skillArchive $claudeSkills
Expand-Archive -Force $skillArchive $codexSkills
```

Before 0.7.0 is released, copy `skills/word-document-review` from the clone made in step 1
into either destination instead. Install into just one destination when you use only Claude
Code or only Codex.

## 3. Connect the server

Use the Claude Code or Codex recipe in
[deployment and client setup](deployment.md#option-a---local-stdio-for-claude-code-and-codex).
Point the filesystem connection at a dedicated directory containing the documents the agent
may edit.

Restart the client after installing the skill, confirm OfficeAgent appears in its MCP tool
list, then use a request that matches the skill:

> Review `services-agreement.docx`. Preserve its existing comments and revisions, and
> change the payment term from thirty days to forty-five days as a tracked change.

The agent should inspect the review state before editing and send `"mode": "Tracked"`
explicitly. Follow the [three sample workflows](../samples/documents/README.md#reproducible-review-workflows)
to verify clause editing, comment resolution, and table editing.

## Updating

Update the MCP tool and refresh the copied skill from the same release tag. Keeping them on
the same version avoids teaching the agent operations that its installed server does not yet
support.
