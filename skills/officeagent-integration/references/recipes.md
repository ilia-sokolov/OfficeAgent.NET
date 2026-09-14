# Installed-package recipes

These recipes target OfficeAgent.NET `0.9.0` and use fictional documents created at runtime. The project asset defaults to the exact `OfficeAgent.Core` and `OfficeAgent.Word` `0.9.0` package version. Maintainers can override `OfficeAgentPackageVersion` only when verifying the same source against a locally packed candidate. It writes only under its disposable working directory.

Copy [IntegrationRecipes.csproj](../assets/IntegrationRecipes.csproj) and [Program.cs](../assets/Program.cs) into an empty directory, then run:

```bash
dotnet restore
dotnet run --configuration Release --no-restore
```

For an unpublished candidate, put locally packed `0.9.0` packages in a directory and add that directory as the first package source in a temporary `NuGet.Config`. Do not mistake a branch name or source build for a published package.

## Recipe 1: tracked Word edit

Input: a generated `proposal.docx` containing `Prepared for Northwind Labs.`

Dependencies: `OfficeAgent.Core` and `OfficeAgent.Word` `0.9.0`. The sample also uses the transitive Open XML SDK types to construct and inspect its fixture.

The recipe inspects the file, finds `Northwind Labs`, previews a `ChangeTextOp` with `ChangeMode.Tracked`, and commits `Contoso Research`. It asserts:

- the preview is valid and describes the intended replacement;
- the committed file contains `Contoso Research` as visible text;
- `Northwind Labs` remains in Word deletion markup;
- the output has no Office 2019 Open XML schema errors.

It also previews an Excel-only `SetCellOp` against the Word file. The expected recovery evidence is `unsupported-operation`, a failed preview, and an unchanged input SHA-256. Do not retry that operation against Word.

## Recipe 2: scalar and repeating template population

Input: a generated quote template with a `CustomerName` tagged control and one repeating line-item row containing `{{Description}}`, `{{Quantity}}`, and `{{Price}}`.

The recipe binds `Fabrikam Services` and two fictional line items. It asserts that the saved document contains the scalar value and both repeated records, contains no unresolved `{{...}}` marker, and has no Office 2019 schema errors.

If the API returns an uncommitted item, report its diagnostic codes. Fix missing tags or row bindings before retrying; do not present a partial output as success.

## Recipe 3: complete-plan document comparison

Inputs: generated original and revised agreements. The payment term changes from 30 to 45 days and the revised document adds a fictional notice clause.

The recipe requires `comparison.IsComplete` and a non-null `comparison.Plan`, previews that plan, then commits it as a new document. It asserts that the redline contains the revised visible text, preserves replaced text in deletion markup, includes inserted revision elements, and has no Office 2019 schema errors.

When comparison coverage is incomplete, stop and surface the diagnostics. Never apply partial findings as if they were a complete plan.

## Expected output

```text
tracked-edit=passed
refusal=unsupported-operation bytes-unchanged=True
template-population=passed
complete-comparison=passed
all-recipes=passed
```

The repository's clean fixture additionally proves that the skill ZIP has the expected directory layout, every relative skill reference resolves after extraction, these recipes run against locally packed NuGet packages, and the installed directory can be removed. That is deterministic integration evidence, not a live-model trial or native Office visual review.
