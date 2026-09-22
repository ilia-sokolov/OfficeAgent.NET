# v1.0.0 native Office corpus

These are the exact files desktop Office opened for the
[native compatibility matrix](../../../../../docs/native-compatibility.md). Every document is
generated in code from fictional content. None is a customer document or adoption evidence.

| Path | Contents |
| --- | --- |
| `inputs/` | Each case's input, as generated |
| `outputs/` | Each case's OfficeAgent output: the file Office opened |
| `controls/` | Deliberately corrupt files that Office must refuse |
| `manifest.json` | Every case: operation family, description, input and output SHA-256, and what Office must observe |
| `results.json` | What Office observed for every check, with the application build, the OS and the manifest hash. Local paths are removed |

## Rerunning

The package and semantic layers run on every build, without Office:

```powershell
dotnet test tests/OfficeAgent.Tests --filter "FullyQualifiedName~NativeCorpusTests"
```

Those tests regenerate the corpus in memory. ZIP metadata is not deterministic, so regenerated
files have different hashes from the ones here, with the same parts and meaning. The native
evidence belongs to these exact files. To repeat it on a Windows machine with desktop Office,
write a fresh corpus and run the harness against it:

```powershell
$env:OFFICEAGENT_NATIVE_CORPUS_OUTPUT = 'C:\temp\officeagent-native'
dotnet test tests/OfficeAgent.Tests --filter "FullyQualifiedName~Writes_the_corpus_for_native_verification_when_asked"
Remove-Item Env:OFFICEAGENT_NATIVE_CORPUS_OUTPUT
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/native/office_check.ps1 -Corpus C:\temp\officeagent-native
```

To replace this record after a passing run:

```powershell
python scripts/native/publish_results.py C:\temp\officeagent-native
```

The script refuses to publish a run with any failed case, failed check or missed control, or
whose results were not recorded against these exact manifest and output bytes. The same bindings
are enforced by `NativeCorpusTests` on every build, so a regenerated corpus cannot keep an old
native record. The JSON and Markdown here are LF text on every platform; the packages are byte
exact.
