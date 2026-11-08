# python-openai-agents - OfficeAgent from the OpenAI Agents SDK

An agent built with the [OpenAI Agents SDK](https://openai.github.io/openai-agents-python/)
edits a Word contract through OfficeAgent's MCP server and leaves the edit as a tracked change.
It is the Python counterpart of [IChatClientWordEdit](../IChatClientWordEdit/), and reads as a
pair with [python-langgraph](../python-langgraph/): same task, same document, same output.

The sample:

1. copies [`services-agreement.docx`](../documents/) into `work/`;
2. starts `officeagent-mcp --stdio` with one filesystem connection, `contracts`, rooted at
   `work/`, configured through environment variables as in
   [the MCP server guide](../../docs/mcp-server.md#install-and-run);
3. gives the agent the server's own instructions, which the server sends when the MCP session
   starts, instead of restating the inspect, find, preview and apply loop;
4. asks for one edit: *change the initial term in clause 5 from twelve months to twenty-four
   months as a tracked change, leave clause 4 (which also says twelve months) as it is, and
   save the result as a new version beside the original*;
5. prints each tool call and the change report; and
6. opens the saved file and checks the package for the tracked revision.

The input in `work/` is left as it was. The result is `work/services-agreement.v2.docx`.

## Prerequisites

- Python 3.10 or later (tested with 3.12)
- The OfficeAgent MCP server, 1.2.0 or later, from one of:
  - the .NET tool: `dotnet tool install --global OfficeAgent.Mcp` (needs the .NET 8 or later
    SDK; run `dotnet tool update --global OfficeAgent.Mcp` if an older version is installed);
  - PyPI, with no .NET: `uvx officeagent-mcp --stdio` (see [below](#use-another-server-command));
  - this repository, built with the .NET 10 SDK (see [below](#use-another-server-command)).
- An OpenAI API key in the `OPENAI_API_KEY` environment variable. The model must support
  function calling.

No key or secret is in the sample. The key is read from the environment by the OpenAI client.

## Run

From the repository root.

Bash:

```bash
cd samples/python-openai-agents
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export OPENAI_API_KEY="<your key>"
python edit_contract.py
```

PowerShell:

```powershell
cd samples/python-openai-agents
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:OPENAI_API_KEY = "<your key>"
python edit_contract.py
```

`OPENAI_MODEL` chooses the model; it defaults to `gpt-5.6-luna`, the Agents SDK's own default.

To check the server connection without a model or a key:

```bash
python edit_contract.py --list-tools
```

It prints the server's version, the length of its instructions, its tools, and the
`list_connections` result, which should name the `contracts` connection.

### Use another server command

`OFFICEAGENT_MCP_COMMAND`, or `--server-command`, replaces `officeagent-mcp --stdio`. From PyPI,
with no .NET needed:

```bash
export OFFICEAGENT_MCP_COMMAND="uvx officeagent-mcp --stdio"
```

Built from this repository: build once from the repository root, then point the command at the
built assembly. The path below is relative to this sample's folder, where you run the script.
Do not use `dotnet run` here: it writes build output to standard output, which the MCP client
reads as protocol messages.

```bash
dotnet build src/OfficeAgent.Mcp -c Release
export OFFICEAGENT_MCP_COMMAND="dotnet ../../src/OfficeAgent.Mcp/bin/Release/net10.0/OfficeAgent.Mcp.dll --stdio"
```

`--document` takes another `.docx` and `--workdir` another folder for the connection root.

## What you should see

The output has this shape. Which tools the model calls, and in what order, varies from run to
run; it may open the document and call `edit_document`, or find the clause and use `apply_plan`:

```text
-> open_document {"connectionId": "contracts", "source": "services-agreement.docx", ...
-> find_in_document {"connectionId": "contracts", "documentId": "...", "pattern": "twelve months", ...
-> apply_plan {"connectionId": "contracts", "documentId": "...", "planJson": "{\"snapshot\": ...
   isValid=True committed=True writeOutcome=committed outputName=services-agreement.v2.docx

Change report:
[
  {
    "verb": "changeText",
    "before": "continues for twelve months",
    "after": "continues for twenty-four months",
    ...
  }
]

Tracked insertion of 'twenty-four months': True
Tracked deletion of 'twelve months':       True
Clause 4 still says twelve months:         True
```

The three checks at the end read `word/document.xml` in the saved package: the new words must be
inside a `w:ins` element, the old ones inside a `w:del`, and clause 4 must be unchanged.

To check by eye, open `work/services-agreement.v2.docx` in Word and open **Review > Reviewing
Pane**. Clause 5 shows `twelve months` struck through and `twenty-four months` inserted, a
deletion and an insertion by `OfficeAgent` that you can accept or reject. The document's existing revision in clause 3 (2% to 3%) and
Priya Raman's comment on clause 4 are still there.

`twelve months` appears in clauses 4 and 5, which is why the task names the clause. An edit that
targets `{ "find": "twelve months" }` alone is refused as `ambiguous-anchor`, with both
candidates listed and nothing written; the agent then re-issues it for clause 5.
