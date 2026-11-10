# Privacy policy

OfficeAgent.NET is open-source software that runs on your own machine or on a server you
operate. Nobody runs it for you: there is no OfficeAgent service, account, or cloud component.
This policy covers the libraries, the MCP server (`OfficeAgent.Mcp`), its standalone downloads,
the Claude Desktop bundles, the `officeagent-mcp` PyPI package, the Claude Code plugin, and the
container image.

## Data collection

OfficeAgent collects nothing about you. It has no telemetry, analytics, crash reporting,
update check, or licence check, and it never contacts the project or its maintainers.

## How documents are used and stored

OfficeAgent reads and edits only the documents you, or the host that runs it, point it at:

- **Folder connections**, including the one the Claude Desktop bundle asks you to choose and
  the project folder the Claude Code plugin uses. OfficeAgent reads and writes files inside
  that folder and nowhere else. To give each document an id, it keeps
  `.officeagent/index.json` inside the folder, which maps ids to file paths. A read-only folder
  connection writes nothing, and keeps its ids in memory instead.
- **Office's own template folders**, only if a host turns on `OfficeTemplates:Enabled`.
  OfficeAgent reads where Word, PowerPoint and Excel keep personal and workgroup templates, from
  the Office registry values on Windows or the Office group container on macOS, and reads the
  templates there. It never writes to them.
- **SharePoint connections**, if a host configures one. OfficeAgent reads and writes the
  documents it is given through Microsoft Graph, using the credentials the host supplies. A
  host can choose to keep the id-to-document mapping in a file it names.
- **iManage Work connections**, if a host configures one. OfficeAgent reads documents and saves
  new versions and documents through the iManage Work REST API, using the credentials the host
  supplies. A host can choose to keep the id-to-document mapping in a file it names.
- **Google Drive, Box, Azure Blob Storage and Amazon S3 connections**, if a host configures one,
  in code or in the MCP server's settings. OfficeAgent reads and writes the documents it is
  given through that service's API, using the credentials the host supplies, and keeps the
  id-to-document mapping the same way.
- **Session connections** keep documents in memory only, for as long as the server runs.
- **Inline content tools** receive a document in the request and return the result in the
  response; they store nothing.

When an AI assistant uses OfficeAgent through MCP, the tool results it receives, such as
document text, previews, and change reports, go to that assistant. What happens to them next
is decided by the assistant and its provider, under their privacy policy, not this one.

On a hosted SharePoint connection that acts as the signed-in user (`onBehalfOf`), each apply
receipt records that user as its `actor`, taken from the token Microsoft Entra ID accepted. If
the host turns on `RevisionAuthorFromActor`, the user's display name is also written into the
document as the author of its tracked changes and comments. Both stay with the receipt and the
document, wherever the host keeps them.

OfficeAgent writes operational log messages, for example which document was saved and its
size, through the standard .NET logging system to wherever the host sends its logs. The MCP
server writes them to its own console. They stay on the machine unless the host forwards them.

## Third-party sharing

OfficeAgent sends nothing to any third party. The exceptions are storage connections that a
host configures, each of which talks only to the service the host names, to read and save
documents and to obtain access tokens:

- A SharePoint connection talks to Microsoft Graph and Microsoft Entra ID in that host's own
  Microsoft 365 tenant. Microsoft's privacy statement governs that service.
- An iManage Work connection talks to the iManage server the host names, iManage Cloud or the
  host's own. The host's agreement with iManage governs that service.
- A Google Drive connection talks to the Google Drive API and Google's OAuth token endpoint, and
  a Box connection to the Box API and Box's token endpoint. Google's and Box's privacy policies
  govern those services.
- An Azure Blob Storage or Amazon S3 connection reads and saves documents in that host's own
  storage account or bucket. It uses a connection string the host supplies, or obtains
  credentials the way the host's Azure or AWS environment provides them: through
  `DefaultAzureCredential`, such as a managed identity or Microsoft Entra ID, or the AWS SDK's
  default credential chain. Microsoft's and Amazon's
  privacy statements govern those services.

## Data retention

- Edited documents are yours and remain where they were saved.
- `.officeagent/index.json` stays in the folder until you delete it. `remove_document` removes
  an entry; it never deletes the document itself.
- A registration file a host names for a SharePoint, iManage Work, Google Drive, Box, Azure Blob
  Storage or Amazon S3 connection stays until the host deletes it. Without one, those ids are
  kept in memory and discarded when the server stops.
- Session documents are discarded when the server stops.
- Logs are kept according to the host's own logging configuration.

## Contact

Questions about this policy: open an issue at
https://github.com/ilia-sokolov/OfficeAgent.NET/issues. To report a security or privacy problem
privately, follow [SECURITY.md](SECURITY.md).

Changes to this policy are recorded in the repository's history and noted in the changelog.
