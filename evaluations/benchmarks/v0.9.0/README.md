# v0.9.0 benchmark evidence

Raw measurements for the v0.9.0 candidate, and the report generated from them. See
[docs/benchmarks.md](../../../docs/benchmarks.md) for what is measured and why the threshold
policy is what it is.

## What is here

| File | What it is |
| --- | --- |
| `baseline-run-1.json` … `baseline-run-5.json` | Five independent runs of the direct .NET path on one commit, used as the baseline |
| `verification-run.json` | A sixth direct run of that same commit, compared against those five |
| `report.md` | Generated from the direct runs, not written by hand |
| `mcp-baseline-run-1.json` … `mcp-baseline-run-3.json` | Three runs of the installed MCP path on the same commit |
| `mcp-verification-run.json` | A fourth MCP run, compared against those three |
| `mcp-report.md` | Generated from the MCP runs |

## The scenario set changed on 2026-09-16

An earlier version of these files covered six scenarios. Two of them were weaker than their
names suggested: `excel-bounded-edit` inspected a workbook and never edited one, and
`powerpoint-chart-deck` inspected a deck containing a chart rather than binding chart data.
Both are replaced, and image-heavy decks and a concurrent-edit scenario are added, so the set
now matches what the task asked for.

The old baselines were **not** carried forward. Comparing them against the new set would have
produced `missing` and `new` rows, which block, and that is the reporter working rather than a
problem to route around: numbers from a different set of scenarios are not a baseline for this
one. The baselines here were recollected against the current set.

## Why six runs of one commit

The verification run exists to answer a question that has to be answered before any of these
numbers mean anything: **does this setup report a regression when nothing changed?**

It did, at first. An earlier threshold policy derived its limits from the spread of
repetitions inside a single run, and comparing two consecutive runs of one commit produced two
scenarios "regressed" by 24%. That is a measurement of the method, not of the code, and it is
why the policy now derives thresholds from movement between runs and gates on allocated bytes
rather than wall-clock time. The numbers behind that decision:

| Metric | Movement between the five baseline runs |
| --- | --- |
| Scenario median wall-clock time | 8.6% to 21.1% |
| Scenario median allocated bytes | 0.016% to 0.052% |

With the corrected policy, the verification run reports zero blocking rows against the
baseline, while still reporting one scenario as `slower` rather than hiding it.

## The MCP path has something better to gate on

The installed-server runs measure what an agent pays in protocol traffic, and those numbers
behave nothing like wall-clock time. Across ten iterations:

| Metric | Variation between iterations |
| --- | --- |
| Tool calls | none, identical every time |
| Bytes sent | none, identical every time |
| Bytes received | about 1.2%, generated content differing slightly |

So tool calls are compared exactly and any increase blocks. An adapter change that needs more
round trips to do the same work costs every user something real, and it is invisible in a
timing number on a machine this noisy.

The MCP scenarios also carry no allocation figure. The workflow runs inside the installed
server process, which this harness drives from outside over stdio, so managed allocation is
not observable from here. It is recorded as zero rather than estimated, and the allocation
gate is inert for those scenarios by design.

## Reproducing

From the repository root, on the commit named in the report:

```bash
dotnet run --project tests/OfficeAgent.Benchmarks --configuration Release -- --repetitions 15 --warmup 3 --output my-run.json
```

```bash
python scripts/benchmark_report.py --results my-run.json --baseline evaluations/benchmarks/v0.9.0/baseline-run-1.json evaluations/benchmarks/v0.9.0/baseline-run-2.json evaluations/benchmarks/v0.9.0/baseline-run-3.json evaluations/benchmarks/v0.9.0/baseline-run-4.json evaluations/benchmarks/v0.9.0/baseline-run-5.json
```

Your absolute timings will differ, because they are a property of your machine. The
correctness assertions, the allocation figures and the work units should not.

## What these numbers are not

These runs come from one ordinary developer workstation, recorded in each file's environment
block. They are not a claim about any other hardware, not a comparison with any other library,
and not adoption evidence. The corpus is generated, fictional and small.

They are also not a measurement of model behaviour. The `mcp-*` scenarios count the protocol
traffic of a fixed, scripted workflow, which is a property of the product. How many calls a
model would choose to make, how often it would retry, and what that would cost are measured by
the [agent-selection evaluation](../../agent-selection/v0.9.0/README.md), are currently
`NOT_RUN` there for want of model access, and are deliberately not duplicated here.
