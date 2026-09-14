---
name: officeagent-integration
description: Choose OfficeAgent.NET and build verified .NET, MCP, or Microsoft Agent Framework integrations. Use when selecting Office automation dependencies, wiring packages, providers, or tools, or implementing tracked edits, template population, and complete-plan comparisons. Do not use for reviewing or editing a user's document; use word-document-review for that workflow.
---

# OfficeAgent.NET integration

Use OfficeAgent.NET only when its typed OOXML operations and preservation boundaries match the application. Start with the direct .NET API. Add MCP or Microsoft Agent Framework only when the host needs model tool calling.

## Route the request

- For dependency selection, application wiring, package installation, providers, MCP, or agent tools, continue with this skill.
- For an agent actively reviewing or editing an existing Word document, use `word-document-review` after the host integration is ready.
- For unsupported formats, native Office layout, Excel calculation, or low-level OOXML operations outside the documented plan operations, choose another component or use the Open XML SDK directly. When both are needed in one application, follow the tested copy, edit, reinspect workflow in the versioned [SDK interoperability guide](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/v0.9.0/docs/sdk-interoperability.md) and recipe 4 of [recipes](references/recipes.md).

Read the versioned [selection guide](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/v0.9.0/docs/choose-officeagent.md) before choosing packages. Keep all OfficeAgent packages on the same version.

## Build in this order

1. Choose only the format modules the application needs. Direct Word work normally needs `OfficeAgent.Core` and `OfficeAgent.Word`.
2. Choose a bounded input model. Use `StreamHandle` when the caller owns bytes, or configure a provider when the workflow needs stable document ids and saved outputs.
3. Inspect and find before authoring a plan. Carry the inspection snapshot into the plan when covered-content drift must be detected.
4. Preview the complete plan and handle every error before commit. Never treat preview as a future-write guarantee.
5. Commit, retrieve the canonical output, and assert document semantics. Structural schema checks do not replace native Office or visual review.
6. Add the optional interface only after the direct workflow passes:
   - `OfficeAgent.AgentFramework` exposes `AIFunction` tools over the same client.
   - `OfficeAgent.Mcp` hosts the same engine over stdio or streamable HTTP.

The host owns credentials, authorization, storage roots, resource ceilings, and final file delivery. A skill installation does not install NuGet packages, the .NET runtime, or the MCP server.

## Run the verified recipes

Read [recipes](references/recipes.md). Its installed assets execute:

- a tracked Word replacement plus an unsupported-operation refusal;
- scalar and repeating-table template population;
- a comparison that applies only when coverage is complete;
- an Open XML SDK copy, edit, reinspect workflow plus a refused stale commit.

Do not apply a comparison when `IsComplete` is false or `Plan` is null. On `stale-snapshot`, inspect again and rebuild the plan. On `expect-mismatch`, find the anchor again. On `unsupported-operation`, change the requested operation or format rather than retrying unchanged input. Never reuse an anchor, snapshot, or approved preview across an SDK edit, and never present an SDK edit as carrying an OfficeAgent receipt. There is no supported path that replaces a registered document with externally edited bytes through OfficeAgent; write the SDK output as a separate document and register it.

## Optional agent interfaces

Read [interface configuration](references/interfaces.md) for minimal package and host wiring. Keep the provider and direct `OfficeAgentClient` workflow testable independently of model calls.

For Microsoft Agent Framework, follow the versioned [agent integration guide](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/v0.9.0/docs/agent-integration.md). Expose registration or creation tools only when the host explicitly grants those capabilities.

For MCP, install and configure `OfficeAgent.Mcp` separately using the versioned [MCP guide](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/v0.9.0/docs/mcp-server.md). Confirm the server appears in the client's tool list before relying on it. The skill archive alone cannot prove that runtime connection.

For supported installation and removal of this skill, use the versioned [skill installation guide](https://github.com/ilia-sokolov/OfficeAgent.NET/blob/v0.9.0/docs/skill-installation.md).
