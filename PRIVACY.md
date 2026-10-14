# Privacy policy

OfficeAgent.NET is open-source software that runs on your own machine or on a server you
operate. Nobody runs it for you: there is no OfficeAgent service, account, or cloud component.
This policy covers the libraries, the MCP server (`OfficeAgent.Mcp`), its standalone downloads,
the Claude Desktop bundles, and the container image.

## Data collection

OfficeAgent collects nothing about you. It has no telemetry, analytics, crash reporting,
update check, or licence check, and it never contacts the project or its maintainers.

## How documents are used and stored

OfficeAgent reads and edits only the documents you, or the host that runs it, point it at:

- **Folder connections**, including the one the Claude Desktop bundle asks you to choose.
  OfficeAgent reads and writes files inside that folder and nowhere else. To give each
  document an id, it keeps `.officeagent/index.json` inside the folder, which maps ids to file
  paths.
- **SharePoint connections**, if a host configures one. OfficeAgent reads and writes the
  documents it is given through Microsoft Graph, using the credentials the host supplies. A
  host can choose to keep the id-to-document mapping in a file it names.
- **Session connections** keep documents in memory only, for as long as the server runs.
- **Inline content tools** receive a document in the request and return the result in the
  response; they store nothing.

When an AI assistant uses OfficeAgent through MCP, the tool results it receives, such as
document text, previews, and change reports, go to that assistant. What happens to them next
is decided by the assistant and its provider, under their privacy policy, not this one.

OfficeAgent writes operational log messages, for example which document was saved and its
size, through the standard .NET logging system to wherever the host sends its logs. The MCP
server writes them to its own console. They stay on the machine unless the host forwards them.

## Third-party sharing

OfficeAgent sends nothing to any third party. The one exception is a SharePoint connection
that a host configures: OfficeAgent then talks to Microsoft Graph and Microsoft Entra ID in
that host's own Microsoft 365 tenant, to read and save documents and to obtain access tokens.
Microsoft's privacy statement governs that service.

## Data retention

- Edited documents are yours and remain where they were saved.
- `.officeagent/index.json` stays in the folder until you delete it. `remove_document` removes
  an entry; it never deletes the document itself.
- Session documents are discarded when the server stops.
- Logs are kept according to the host's own logging configuration.

## Contact

Questions about this policy: open an issue at
https://github.com/ilia-sokolov/OfficeAgent.NET/issues. To report a security or privacy problem
privately, follow [SECURITY.md](SECURITY.md).

Changes to this policy are recorded in the repository's history and noted in the changelog.
