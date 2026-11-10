# Troubleshooting

Start with the exact error code and the `(connectionId, documentId)` involved.
OfficeAgent returns validation failures without writing; provider failures can
occur after remote storage accepted a request, so their recovery is different.

## Server does not start

With no storage configuration, the MCP server starts with an in-memory connection named
`session`, enables document creation there, and prints that choice to stderr. Its documents
disappear when the process stops. If you expected persistent storage and see that notice,
check the names of your `OfficeAgent__` variables.

For a persistent local smoke test:

```powershell
New-Item -ItemType Directory -Force C:\officeagent-documents | Out-Null
$env:OfficeAgent__FileSystemConnections__0__ConnectionId = "documents"
$env:OfficeAgent__FileSystemConnections__0__RootPath = "C:\officeagent-documents"
officeagent-mcp --stdio
```

An invalid SharePoint `AuthMode` is a configuration error. Valid values are
`appOnly` and `onBehalfOf` (including the documented hyphenated aliases); the
server never silently falls back to a shared identity.

A missing or malformed setting on an iManage Work, Google Drive, Box, Azure Blob Storage or
Amazon S3 connection stops the server at startup, and the message names the environment
variable to set, such as `OfficeAgent__IManageConnections__0__ServerUrl`. So do a
`TemplateLocation` that is not `Connection`, `Personal`, `Workgroup` or `Organization`, a
`TemplateLocation` without a `TemplatesFolder` on one of those five connections, a
filesystem `TemplatesFolder` outside the root, and a `ReadOnly` connection that names a creation
folder (`CreationFolderId`, or SharePoint's `CreationDriveId` and `CreationFolderItemId`). With `OfficeTemplates:Enabled`, an Office template folder that does not exist is skipped
with a line on stderr rather than stopping the server. See
[MCP server settings](mcp-server.md#settings).

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
| `access-denied` | A filesystem path is under the configured root and does not traverse a symlink/reparse point; the SharePoint, iManage, Drive or Box identity can reach the item; a blob name or S3 key stays under the connection's prefix; a Box Shield policy is not blocking it. A read-only connection refuses every save and creation this way, and a read-only filesystem connection also refuses removing a registration. |
| `extension-not-allowed` | The connection allows the lower-case extension, including `.pptx` for decks and `.dotx`, `.potx` or `.xltx` for templates. Native Google Docs, Sheets and Slides are refused this way; export them to Office files first. |
| `content-too-large` | The connection's `MaximumBytes` is high enough. SharePoint uploads large files in chunks, so `MaximumBytes` is its only limit. |
| `configuration-error` | The connection id is unique and all required provider settings are present. |
| `encrypted-document` | The file is protected by a sensitivity label or a password. OfficeAgent does not decrypt; remove the protection in Office, or use a copy without it. |

For SharePoint app-only access with `Sites.Selected`, confirm both tenant admin
consent and an explicit application permission assignment on the intended site.
For On-Behalf-Of, confirm that the incoming user token targets the middle-tier
API, the gateway forwards it, and the app has delegated Graph consent.

## A plan is not valid

- `stale-snapshot`: inspect again and rebuild the complete plan.
- `expect-mismatch`: find the text again and use the newly issued anchor.
- `ambiguous-anchor`: use the returned candidate index, `"match": "all"`, or a more specific
  string. The message names each candidate's location, such as `match 0 (body):`.
- `anchor-not-found` for a plan with `"match": "all"`: every such target matched nothing, so
  nothing was saved. Check the text, `caseSensitive` and `wholeWord`.
- `invalid-argument` on a `find` target: `match` must be an index or `"all"`.
- `unsupported-operation`: use the operation matrix for the document format.
- `invalid-operation` with `Tracked` on a deck: send `Direct` or configure the
  connection's default change mode.
- `invalid-operation` on `insertSlide`: the layout has no placeholder for the title or the
  body. Choose another layout, or leave that part out.
- A Word style refused as undefined: use one of the styles the message lists, or define it
  in the same plan with `defineStyle`.
- `operation-conflict`: combine compatible payloads or split the work into two
  plans with an inspection between them.

## Creating from a template is refused

- `macro-enabled-document`: the template or document is macro-enabled (`.dotm`, `.potm`,
  `.xltm`, `.docm`, `.pptm`, `.xlsm`, or any package with a VBA project). Save it in Office as a
  macro-free template, and use that.
- `output-format-mismatch`: the output name must end in the document extension of the
  template's format, such as `.docx` for a `.dotx` template.
- `extension-not-allowed`: the connection that holds the template must allow its extension.

Do not remove a snapshot or `expect` merely to make a drift error disappear;
those checks prevent an edit from moving to the wrong content.

## Commit or creation has an uncertain result

`Replace` checks the source version. On `version-conflict`, inspect and re-author
the plan; never retry the stale request unchanged.

Apart from the chunks of a large SharePoint upload, Graph requests are not retried
automatically. After a timeout, 429, or transient 5xx during a write, first reopen the
registered item or inspect the requested destination name. Remote storage may have accepted
the content even though its response was lost. Creation never overwrites an occupied name.

A SharePoint save of 4 MiB or more (`UploadSessionThresholdBytes`) goes up in chunks, and a failed chunk is retried from where
Graph says to resume. A failure before the last chunk is `write-rejected`: nothing was stored,
so the save can be repeated. A failure on the last chunk, or a lost reply to it, is
`outcome-unknown`: reconcile storage before retrying, as above.

`write-rejected` also means the file is checked out or locked: in SharePoint, by someone who
has it checked out or open for exclusive editing; in iManage Work, by a checkout or lock; in
Box, by another user's lock. The message names the lock. Have it released, or save as a new
document.

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
