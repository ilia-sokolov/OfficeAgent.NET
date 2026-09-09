# Adoption validation

Use this check before broad promotion of a release. It measures whether developers can
discover a relevant OfficeAgent workflow and produce a correct Office file. Repository stars
and package downloads measure attention, not successful use.

## Success definition

A trial succeeds when a participant, without maintainer intervention:

1. understands that OfficeAgent creates and edits Word documents, PowerPoint decks, and Excel workbooks;
2. chooses the shortest setup path for the assigned task;
3. installs or starts the MCP server, or adds the required .NET packages;
4. completes the assigned document outcome; and
5. opens the saved file in Word, PowerPoint, or Excel and verifies the requested result and preserved
   unrelated content.

The Word review skill is required only for the review trial. Record setup failures separately
from operation and verification failures. Stop the timer when the participant verifies the
file in Office, rather than when the agent or application reports success.

## Five-trial portfolio

Run at least five fresh-user trials across the representative tasks below. For MCP trials,
include Claude Code and Codex, Windows and one Unix platform, and a fresh client configuration
or disposable user profile. A participant may complete more than one task, but each row must
start from a clean setup.

| Trial | Representative outcome | Suggested path |
| --- | --- | --- |
| 1 | Make a targeted change in an existing Word document and preserve unrelated structure | [README Word edit](../README.md#try-a-word-edit) |
| 2 | Create a new Word document and add structured content | [MCP tools](mcp-server.md#tools) or [.NET `CreateAsync`](getting-started.md#create-a-document-instead) |
| 3 | Create or update a PowerPoint deck with text plus one rich-content operation | [PowerPoint support](powerpoint.md#creating-a-deck) |
| 4 | Review a Word document with tracked changes or comments | [Sample review workflows](../samples/documents/README.md#reproducible-review-workflows) and the [optional review skill](skill-installation.md) |
| 5 | Complete one workflow through a second integration or storage path | [Agent integration](agent-integration.md), [SharePoint](document-providers.md#the-sharepoint-provider), or [session storage](mcp-server.md#documents-with-no-storage) |

Record each attempt here:

| Trial | Outcome | Client or API and version | OS | Server/package version | Skill, if used | Setup result | Correct saved file | Time to verification | Blocker or confusion |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 |  |  |  |  |  |  |  |  |  |
| 2 |  |  |  |  |  |  |  |  |  |
| 3 |  |  |  |  |  |  |  |  |  |
| 4 |  |  |  |  |  |  |  |  |  |
| 5 |  |  |  |  |  |  |  |  |  |

For each failure, capture the exact command or API step, stable OfficeAgent error code, and
the first documentation sentence that caused uncertainty. Do not collect documents,
credentials, tenant details, or private model transcripts.

## Promotion gate

Start broader promotion when all five outcomes produce a correct saved file and no single
documentation problem blocks more than one trial. Fix repeated setup failures first, rerun
the affected trials from a clean profile, and retain both attempts in the worksheet.

After the activation gate passes, measure whether users discover and complete a second useful
task. Track these separately by workflow so a popular review demo does not hide friction in
creation, presentation, or integration paths:

- completion rate and median time to a verified file;
- Word creation, Word editing, PowerPoint, and review activation;
- optional skill discovery and activation for review tasks;
- second-workflow completion and the workflow chosen;
- installation, configuration, tool-selection, operation, and verification failures.

Publish only aggregate results with the client, model, server, package, and optional skill
versions plus the test date. Keep exploratory observations labelled as such when raw traces
or version metadata were not retained.
