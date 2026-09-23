#!/usr/bin/env python3
"""Run the stored v0.8 and v0.9 records through installed candidate packages.

A release candidate must keep what earlier releases promised, as a consumer sees it: packages
restored from a feed, not project references. This builds a fresh console consumer against the
locally packed candidate and checks, from the repository's stored corpora:

- the stored v0.8 edit plan still applies to the contract it was written for (generated here as
  the repository's tests generate it), and an unknown contractVersion is still refused;
- the stored v0.8 receipt and validation error still read with the host settings the upgrade
  guide documents;
- the v0.9 Word preservation guarantee still holds: a tracked edit on the rich fixture leaves every
  protected part byte identical;
- the stored v0.8 PowerPoint and Excel fixtures still open and inspect;
- the 1.0 recovery types are reachable from the packages;
- every loaded OfficeAgent assembly carries exactly the candidate version: its informational version,
  without the source-control suffix after '+', equals the requested package version;
- when OFFICEAGENT_EXPECT_RUNTIME_MAJOR is set, the consumer ran on that .NET major version, so a
  runtime cell of a CI matrix cannot pass on another runtime.

It proves packaging and compatibility for the candidate, not publication.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from smoke_packaged_artifacts import (  # noqa: E402
    SmokeError,
    isolated_env,
    repository_version,
    run,
    write_nuget_config,
)

ROOT = Path(__file__).resolve().parents[1]
CORPUS_V08 = ROOT / "tests" / "OfficeAgent.Tests" / "Corpus" / "v0.8.0"
CORPUS_V09 = ROOT / "tests" / "OfficeAgent.Tests" / "Corpus" / "v0.9.0" / "word-preservation"

# What the consumer reads. Checked before building so a moved fixture fails here, not as a
# confusing runtime error inside the consumer.
REQUIRED_RECORDS = (
    CORPUS_V08 / "word-edit-plan.json",
    CORPUS_V08 / "apply-receipt.json",
    CORPUS_V08 / "validation-error.json",
    CORPUS_V08 / "deck-with-chart.pptx",
    CORPUS_V08 / "workbook-styles.xlsx",
    CORPUS_V09 / "manifest.json",
    CORPUS_V09 / "rich-word-features.docx",
)

PACKAGES = ("OfficeAgent.Core", "OfficeAgent.Word", "OfficeAgent.PowerPoint", "OfficeAgent.Excel",
            "OfficeAgent.AgentFramework")

# SemVer 2.0 build metadata: dot-separated, non-empty identifiers of ASCII alphanumerics and hyphens.
BUILD_METADATA = re.compile(r"[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*")

# The assemblies whose version the consumer reports. Abstractions arrives transitively.
ASSEMBLIES = ("OfficeAgent.Abstractions",) + PACKAGES

PROJECT = """<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
{references}
  </ItemGroup>
</Project>
"""

PROGRAM = r"""using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

var v08 = args[0];
var v09 = args[1];
var failures = new List<string>();

void Check(string name, Func<Task<string?>> check)
{
    string? problem;
    try { problem = check().GetAwaiter().GetResult(); }
    catch (Exception ex) { problem = ex.GetType().Name + ": " + ex.Message; }
    Console.WriteLine(problem is null ? $"PASS {name}" : $"FAIL {name}: {problem}");
    if (problem is not null) failures.Add(name);
}

OfficeAgentTools Tools()
{
    var client = new OfficeAgentClient(
        new DocumentProviderRegistry(new IDocumentProvider[] { new MemoryDocumentProvider("mem") }),
        new WordModule(), new PowerPointModule(), new ExcelModule());
    return new OfficeAgentTools(client);
}

async Task<(OfficeAgentTools Tools, string Id)> Imported(string name, byte[] bytes)
{
    var tools = Tools();
    var imported = JsonNode.Parse(await tools.ImportDocumentContent("mem", name, Convert.ToBase64String(bytes)))!;
    return (tools, imported["documentId"]!.GetValue<string>());
}

// The contract the stored v0.8 plan was written for: the repository's DocxFactory.Contract clause,
// paragraph 00000002, whose "Acme Corp" the plan replaces. Built with the Open XML SDK the packages
// bring in, since a consumer cannot use the repository's test helpers.
static byte[] Contract()
{
    using var stream = new MemoryStream();
    using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
    {
        static DocumentFormat.OpenXml.Wordprocessing.Run Run(string text) =>
            new(new DocumentFormat.OpenXml.Wordprocessing.Text(text) { Space = SpaceProcessingModeValues.Preserve });
        document.AddMainDocumentPart().Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
            new DocumentFormat.OpenXml.Wordprocessing.Body(
                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(Run("Service Agreement")) { ParagraphId = "00000001" },
                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    Run("Acme "), Run("Corp"), Run(" shall provide services to "), Run("Acme Corp"), Run(".")) { ParagraphId = "00000002" }));
    }
    return stream.ToArray();
}

static string? ErrorCode(JsonNode node) => node["errors"]?.AsArray().FirstOrDefault()?["code"]?.GetValue<string>();

Check("v0.8 stored plan applies through the tools", async () =>
{
    var (tools, id) = await Imported("contract.docx", Contract());
    var result = JsonNode.Parse(await tools.ApplyPlan("mem", id, File.ReadAllText(Path.Combine(v08, "word-edit-plan.json"))))!;
    if (result["committed"]?.GetValue<bool>() != true) return "not committed: " + result.ToJsonString();
    if (result["writeOutcome"]?.GetValue<string>() != "committed") return "writeOutcome " + result["writeOutcome"];
    var exported = JsonNode.Parse(await tools.ExportDocumentContent("mem", id))!;
    using var archive = new ZipArchive(new MemoryStream(Convert.FromBase64String(exported["contentBase64"]!.GetValue<string>())));
    using var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open());
    var xml = reader.ReadToEnd();
    return xml.Contains("<w:ins", StringComparison.Ordinal) && xml.Contains("Globex Inc.", StringComparison.Ordinal)
        ? null : "the tracked replacement is not in the document";
});

Check("an unknown contractVersion is still refused before any change", async () =>
{
    var (tools, id) = await Imported("contract.docx", Contract());
    var plan = JsonNode.Parse(File.ReadAllText(Path.Combine(v08, "word-edit-plan.json")))!;
    plan["contractVersion"] = "9.9";
    var result = JsonNode.Parse(await tools.ApplyPlan("mem", id, plan.ToJsonString()))!;
    return result["committed"]?.GetValue<bool>() == false && ErrorCode(result) == ValidationErrorCodes.ContractMismatch
        ? null : "not refused with contract-mismatch: " + result.ToJsonString();
});

var host = new JsonSerializerOptions(JsonSerializerDefaults.Web) { Converters = { new JsonStringEnumConverter() } };

Check("the stored v0.8 receipt reads into the 1.0 receipt type", () =>
{
    var receipt = JsonSerializer.Deserialize<ApplyReceipt>(File.ReadAllText(Path.Combine(v08, "apply-receipt.json")), host)!;
    return Task.FromResult<string?>(receipt.Outcome == ApplyOutcome.Committed && receipt.PlanSha256 == new string('1', 64)
        ? null : "receipt fields differ");
});

Check("the stored v0.8 validation error still uses a catalogued code", () =>
{
    var error = JsonNode.Parse(File.ReadAllText(Path.Combine(v08, "validation-error.json")))!;
    return Task.FromResult<string?>(ErrorCode(error) == ValidationErrorCodes.StaleSnapshot ? null : "code " + ErrorCode(error));
});

Check("the v0.9 preservation guarantee holds: a tracked edit leaves protected parts byte identical", async () =>
{
    var manifest = JsonNode.Parse(File.ReadAllText(Path.Combine(v09, "manifest.json")))!;
    var tracked = manifest["cases"]!.AsArray().First(c => c!["id"]!.GetValue<string>() == "mixed-runs-tracked-edit")!;
    var protectedParts = tracked["protectedParts"]!.AsArray().Select(p => p!.GetValue<string>()).ToArray();
    var input = File.ReadAllBytes(Path.Combine(v09, "rich-word-features.docx"));

    var client = new OfficeAgentClient(new WordModule());
    var hit = (await client.FindAsync(new StreamHandle(new MemoryStream(input, writable: false), "rich.docx"), new FindQuery("target"))).First();
    using var result = client.Commit(new StreamHandle(new MemoryStream(input, writable: false), "rich.docx"), new DocumentPlan
    {
        Revision = new RevisionMetadata { Author = "Candidate Consumer", TimestampUtc = new DateTimeOffset(2026, 9, 23, 0, 0, 0, TimeSpan.Zero) },
        Operations = new PlanOperation[] { new ChangeTextOp { Target = hit.Anchor, With = "objective", Mode = ChangeMode.Tracked } }
    });
    if (!result.Committed) return "not committed: " + string.Join("; ", result.Report.Errors.Select(e => e.Code));
    var output = result.ToBytes();

    static Dictionary<string, string> Parts(byte[] bytes)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes));
        return zip.Entries.ToDictionary(e => e.FullName, e =>
        {
            using var s = e.Open();
            return Convert.ToHexString(SHA256.HashData(s));
        });
    }
    var before = Parts(input);
    var after = Parts(output);
    var changed = protectedParts.Where(p => !after.TryGetValue(p, out var h) || h != before[p]).ToArray();
    return changed.Length == 0 ? null : "protected parts changed: " + string.Join(", ", changed);
});

Check("the stored v0.8 PowerPoint and Excel fixtures still open and inspect", async () =>
{
    foreach (var (file, format) in new[] { ("deck-with-chart.pptx", "PowerPoint"), ("workbook-styles.xlsx", "Excel") })
    {
        var (tools, id) = await Imported(file, File.ReadAllBytes(Path.Combine(v08, file)));
        var inspected = JsonNode.Parse(await tools.InspectDocument("mem", id))!;
        if (inspected["format"]?.GetValue<string>() != format) return $"{file}: {inspected.ToJsonString()[..Math.Min(200, inspected.ToJsonString().Length)]}";
    }
    return null;
});

Check("the 1.0 recovery types are reachable from the packages", () =>
    Task.FromResult<string?>(typeof(DocumentWriteRecoveryException).IsAssignableFrom(typeof(DocumentRegistrationFailedException)) &&
                             typeof(DocumentWriteRecoveryException).IsAssignableFrom(typeof(DocumentWriteOutcomeUnknownException))
        ? null : "the recovery hierarchy is missing"));

Console.WriteLine($"RUNTIME\t{Environment.Version}");

// Reported, not judged here: the exact comparison is scripts/verify_candidate_consumer.py's
// version_mismatches, which its unit tests exercise with near matches.
foreach (var type in new[] { typeof(DocumentPlan), typeof(OfficeAgentClient), typeof(WordModule), typeof(PowerPointModule),
                             typeof(ExcelModule), typeof(OfficeAgentTools) })
{
    var assembly = type.Assembly;
    Console.WriteLine($"ASSEMBLY\t{assembly.GetName().Name}\t{assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? ""}");
}

Console.WriteLine(failures.Count == 0 ? "candidate-consumer=passed" : $"candidate-consumer=failed ({failures.Count})");
return failures.Count == 0 ? 0 : 1;
"""


def normalized_informational_version(value: str | None) -> str | None:
    """The package version an assembly claims: its informational version without the
    source-control suffix from the first '+'. Nothing else is normalized: no trimming and no case
    folding, so whitespace anywhere makes the version unequal. The removed suffix must itself be
    SemVer build metadata; otherwise, as for whitespace after the commit or an empty suffix, the
    value is returned whole, so it cannot equal a package version and is reported as it was read."""
    if not value:
        return None
    version, plus, metadata = value.partition("+")
    if plus and not BUILD_METADATA.fullmatch(metadata):
        return value
    return version


def version_mismatches(observed: dict[str, str | None], expected: str) -> list[str]:
    """Every assembly that does not carry exactly the expected version, by name.

    Exact ordinal equality after normalization: 1.0.0-rc.30 is not 1.0.0-rc.3, and 1.0.0-rc.3 is
    not 1.0.0. A missing assembly or a missing informational version is a mismatch.
    """
    problems = []
    for name in ASSEMBLIES:
        if name not in observed:
            problems.append(f"{name}: not reported by the consumer")
            continue
        normalized = normalized_informational_version(observed[name])
        if normalized is None:
            problems.append(f"{name}: no informational version")
        elif normalized != expected:
            problems.append(f"{name}: informational version {observed[name]!r}, expected {expected!r}")
    return problems


def runtime_mismatch(stdout: str, expected_major: str | None) -> str | None:
    """None when no runtime is required or the consumer ran on the required major version."""
    if not expected_major:
        return None
    reported = next((line.split("\t", 1)[1] for line in stdout.splitlines() if line.startswith("RUNTIME\t")), None)
    if reported is None:
        return "the consumer did not report its runtime"
    return None if reported.split(".", 1)[0] == expected_major.strip() else \
        f"the consumer ran on .NET {reported}, expected major version {expected_major}"


def reported_assembly_versions(stdout: str) -> dict[str, str | None]:
    versions: dict[str, str | None] = {}
    for line in stdout.splitlines():
        if line.startswith("ASSEMBLY\t"):
            _, name, informational = (line.split("\t", 2) + [""])[:3]
            versions[name] = informational or None
    return versions


def verify(artifacts: Path, version: str, dotnet: str) -> None:
    missing = [str(path.relative_to(ROOT)) for path in REQUIRED_RECORDS if not path.is_file()]
    if missing:
        raise SmokeError(f"stored records are missing: {', '.join(missing)}")
    packages = [artifacts / f"{name}.{version}.nupkg" for name in PACKAGES]
    absent = [path.name for path in packages if not path.is_file()]
    if absent:
        raise SmokeError(f"missing candidate packages: {', '.join(absent)}")

    with tempfile.TemporaryDirectory(prefix="officeagent-candidate-consumer-") as temporary:
        consumer = Path(temporary) / "consumer"
        consumer.mkdir()
        cache = Path(temporary) / "package-cache"
        cache.mkdir()
        references = "\n".join(
            f'    <PackageReference Include="{name}" Version="{version}" />' for name in PACKAGES)
        (consumer / "CandidateConsumer.csproj").write_text(
            PROJECT.format(references=references), encoding="utf-8", newline="\n")
        (consumer / "Program.cs").write_text(PROGRAM, encoding="utf-8", newline="\n")
        config = consumer / "NuGet.Config"
        write_nuget_config(config, artifacts, include_nuget_org=True)
        env = isolated_env(cache, config)

        run([dotnet, "restore", "--configfile", str(config)], consumer, env)
        assets = json.loads((consumer / "obj" / "project.assets.json").read_text(encoding="utf-8"))
        resolved = set(assets.get("libraries", {}))
        for name in PACKAGES:
            if f"{name}/{version}" not in resolved:
                raise SmokeError(f"the consumer did not resolve {name} {version} from the candidate feed")

        result = subprocess.run(
            [dotnet, "run", "--configuration", "Release", "--no-restore", "--",
             str(CORPUS_V08), str(CORPUS_V09)],
            cwd=consumer, env=env, text=True, capture_output=True, check=False, timeout=900)
        print(result.stdout.strip())
        if result.returncode or "candidate-consumer=passed" not in result.stdout:
            raise SmokeError(f"candidate consumer failed ({result.returncode}):\n{result.stderr[-2000:]}")
        mismatches = version_mismatches(reported_assembly_versions(result.stdout), version)
        if mismatches:
            raise SmokeError("assemblies do not carry the candidate version: " + "; ".join(mismatches))
        print(f"assembly-versions=passed expected={version} assemblies={len(ASSEMBLIES)}")
        runtime_problem = runtime_mismatch(result.stdout, os.environ.get("OFFICEAGENT_EXPECT_RUNTIME_MAJOR"))
        if runtime_problem:
            raise SmokeError(runtime_problem)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, required=True)
    parser.add_argument("--version", default=None)
    parser.add_argument("--dotnet", default="dotnet")
    args = parser.parse_args()
    try:
        verify(args.artifacts.resolve(), args.version or repository_version(), args.dotnet)
    except (OSError, SmokeError, subprocess.TimeoutExpired) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print("Candidate consumer verification completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
