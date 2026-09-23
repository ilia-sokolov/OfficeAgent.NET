# Generate manual Office fixtures

This maintainer tool creates a fictional statement of work, quarterly business review, and
blank PowerPoint deck for manual MCP and native Office trials. It is fixture preparation, not
a regression test, and therefore does not run during `dotnet test`.

By default it writes to the operating system's temporary `officeagent-realworld` directory.
Set `OFFICEAGENT_REALWORLD_ROOT` to an explicit disposable directory to choose another target:

```powershell
$env:OFFICEAGENT_REALWORLD_ROOT = Join-Path $env:TEMP "officeagent-realworld"
dotnet run --project tools/OfficeAgent.RealWorldFixtures
```

The generated files are fictional. Open them in Word or PowerPoint when a manual workflow calls
for native inspection; their generation alone is not evidence that an edit or rendering worked.
