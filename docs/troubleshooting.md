# Troubleshooting

Start with the exact error code and the `(connectionId, documentId)` involved.
OfficeAgent returns validation failures without writing; provider failures can
occur after remote storage accepted a request, so their recovery is different.

## Server does not start

The MCP server requires at least one configured filesystem or SharePoint
connection. For a local smoke test:

```powershell
New-Item -ItemType Directory -Force C:\officeagent-documents | Out-Null
$env:OfficeAgent__FileSystemConnections__0__ConnectionId = "documents"
$env:OfficeAgent__FileSystemConnections__0__RootPath = "C:\officeagent-documents"
officeagent-mcp --stdio
```

An invalid SharePoint `AuthMode` is a configuration error. Valid values are
`appOnly` and `onBehalfOf` (including the documented hyphenated aliases); the
server never silently falls back to a shared identity.

## The MCP client cannot connect

1. Run `officeagent-mcp --stdio` with the same environment and confirm it remains
   running.
2. Check `claude mcp list` or `codex mcp list`; in a client session, use `/mcp`.
3. Keep client options before the server name in a `claude mcp add` command.
4. In Codex remote configuration, put `url` directly under
   `[mcp_servers.officeagent]`.
5. In stdio mode, stdout is reserved for JSON-RPC. Send application diagnostics
   through `ILogger`, which writes to stderr.

## Registration is refused

| Error | Check |
| --- | --- |
| `not-found` | The file/item exists and the opaque id has not been removed. |
| `access-denied` | A filesystem path is under the configured root and does not traverse a symlink/reparse point; the SharePoint identity can reach the item. |
| `extension-not-allowed` | The connection allows the lower-case extension, including `.pptx` for decks. |
| `content-too-large` | The connection cap is high enough. SharePoint simple upload cannot exceed 250 MB. |
| `configuration-error` | The connection id is unique and all required provider settings are present. |

For SharePoint app-only access with `Sites.Selected`, confirm both tenant admin
consent and an explicit application permission assignment on the intended site.
For On-Behalf-Of, confirm that the incoming user token targets the middle-tier
API, the gateway forwards it, and the app has delegated Graph consent.

## A plan is not valid

- `stale-snapshot`: inspect again and rebuild the complete plan.
- `expect-mismatch`: find the text again and use the newly issued anchor.
- `ambiguous-anchor`: use the returned candidate index or a more specific string.
- `unsupported-operation`: use the operation matrix for the document format.
- `invalid-operation` with `Tracked` on a deck: send `Direct` or configure the
  connection's default change mode.
- `operation-conflict`: combine compatible payloads or split the work into two
  plans with an inspection between them.

Do not remove a snapshot or `expect` merely to make a drift error disappear;
those checks prevent an edit from moving to the wrong content.

## Commit or creation has an uncertain result

`Replace` checks the source version. On `version-conflict`, inspect and re-author
the plan; never retry the stale request unchanged.

Graph requests are not retried automatically. After a timeout, 429, or transient
5xx during a write, first reopen the registered item or inspect the requested
destination name. Remote storage may have accepted the content even though its
response was lost. Creation never overwrites an occupied name.

## The output file is not in the model response

That is intentional. A successful report returns `outputConnectionId` and
`outputDocumentId`; the host retrieves the content with:

```csharp
using var content = await client.OpenReadAsync(
    outputConnectionId, outputDocumentId, cancellationToken);
```

Copy `content.Stream` to the chat platform's attachment API or an authenticated
download response. Do not send OOXML bytes as base64 through model context.

## Health reports OK but SharePoint fails

`/healthz` is a liveness check, not provider readiness. It does not call Graph,
validate effective permissions, or prove that creation destinations exist. Add
a deployment preflight against a harmless test document when the platform needs
a readiness gate.

## Collect a safe reproduction

Record the package versions, target format, operation JSON, stable error code,
and whether the call was preview or commit. Replace business text and identifiers
with synthetic values. Do not attach confidential documents, access tokens,
client secrets, registration indexes, or production paths to a public issue.
