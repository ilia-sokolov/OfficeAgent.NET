# Benchmarks

Reproducible performance and correctness measurements for a release candidate, taken against
the tracked v0.8 acceptance corpus.

Read [what "it worked" means](getting-started.md#what-it-worked-means-here) first. Everything
here is structural and semantic: no native Office opening and no rendered visual parity.

## Published results

The v0.9.0 measurements, the baselines they were compared against, and the generated reports
for both the direct and the installed MCP paths are in
[evaluations/benchmarks/v0.9.0](../evaluations/benchmarks/v0.9.0/README.md). Any performance
statement in this repository should point at those files rather than at a remembered figure
from an older package.

### What is not measured here

Engine performance and model behaviour are separate questions, and this page answers only the
first. How many tool calls a *model* spends to reach a goal, how often it retries, and what
that costs are measured by the agent-selection evaluation, not here, and they are currently
`NOT_RUN` there because no model access was available. See
[adoption validation](adoption-validation.md) and the
[agent-selection evaluation](../evaluations/agent-selection/v0.9.0/README.md) for that state.
The `mcp-*` scenarios below count the protocol traffic of a **fixed** workflow, which is a
property of the product; they say nothing about how many calls a model would choose to make.

## What is measured

Six scenarios, each driven through the provider-backed .NET client against fixtures committed
in `tests/OfficeAgent.Tests/Corpus/v0.8.0/`:

| Scenario | Work | At once |
| --- | --- | ---: |
| `word-contract-tracked-edit` | Inspect a contract, apply one tracked text change, read the saved output | 1 |
| `word-comparison` | Compare two revisions and produce a tracked-change plan | 1 |
| `template-batch-repeated` | Discover a template, preflight a batch, populate eight outputs under a token-bound commit | 1 |
| `word-assembly` | Preview and commit an assembly of two documents | 1 |
| `powerpoint-chart-binding` | Discover a deck's native chart slot and rewrite the chart's data | 1 |
| `powerpoint-image-heavy-deck` | Grow a deck to eight slides and place a background image on each | 1 |
| `excel-bounded-cell-edit` | Edit eight cells in a styled workbook under a bounded cell budget | 1 |
| `word-concurrent-edits` | Apply a tracked edit to four documents simultaneously through one client | 4 |

Every scenario asserts its own correctness. A scenario that produced the wrong result records
a failed iteration; it is never dropped, and a scenario with any failure is reported as failed
regardless of how fast its successful iterations were.

### Preservation is asserted separately from correctness

Three scenarios also compare named package parts byte for byte between the staged input and
the produced output, using the parts the v0.8 corpus manifest declares protected for that
workflow. They are different claims: an edit can be exactly right and still have rewritten a
styles part or dropped an image on the way out.

That check earned its place immediately. Editing a cell in column H of a workbook whose table
occupies A1:B2 leaves `xl/styles.xml` byte-identical but **not** `xl/tables/table1.xml`: the
SDK re-serialises the part it read and moves the namespace declaration to the front of the
attribute list. Same length, same content, different bytes. Word leaves untouched parts
byte-identical; Excel does not, for parts it reads. The difference is presentational rather
than semantic, and it is pinned by `ExcelPreservationTests` rather than papered over by
loosening what the benchmark asserts.

### What each row reports

Latency as median, p95 and max; peak allocation; throughput against two denominators, input
bytes per second and work units per second, because neither alone describes every scenario;
and the concurrency the scenario ran at. Input sizes are recorded per iteration so throughput
has a denominator a reader can check.

## The installed MCP path

The direct benchmark measures the engine. It does not measure what an agent pays to reach the
engine through a packaged server, which is a different question with a different answer.

```bash
dotnet pack OfficeAgent.NET.sln --configuration Release --output ./artifacts
```

```bash
python scripts/benchmark_mcp.py --artifacts ./artifacts --output artifacts/benchmarks/mcp-run-1.json --repetitions 10
```

This installs the packed tool into an empty tool directory with a fresh package cache, exactly
as the packaged smoke does, and drives the same stdio protocol an MCP client would. It records
wall-clock time, and two things the direct path has no equivalent of:

| Metric | What it is |
| --- | --- |
| Tool calls | Round trips an agent spends to complete the workflow |
| Bytes sent and received | The payload it pays for in each direction |

Because it packs and installs, this takes minutes and is not part of the unit suite. It runs
for release evidence.

## Running it

```bash
dotnet run --project tests/OfficeAgent.Benchmarks --configuration Release -- --repetitions 15 --warmup 3 --output artifacts/benchmarks/run-1.json
```

```bash
python scripts/benchmark_report.py --results artifacts/benchmarks/run-1.json --output artifacts/benchmarks/report.md
```

Warmup iterations are measured and discarded, because the first pass through a code path pays
for JIT compilation and the Open XML SDK's static initialisation. The count is recorded in the
report so a reader can see what was excluded.

The raw JSON is the artifact to keep. It records the commit, whether the working tree was
clean, the runtime, operating system, architecture, processor count, garbage collection mode,
build configuration, repetition and warmup counts, the measurement method, and every
individual iteration. The report is derived from it and can be regenerated at any time.

## Comparing against a baseline

A baseline is **several runs**, not one:

```bash
python scripts/benchmark_report.py --results artifacts/benchmarks/candidate.json --baseline artifacts/benchmarks/run-1.json artifacts/benchmarks/run-2.json artifacts/benchmarks/run-3.json --output artifacts/benchmarks/report.md
```

Add `--fail-on-blocking` to use it as a release gate.

## Why allocation gates and wall-clock time does not

This is the part worth understanding before trusting any number here.

Thresholds are derived from measured variance rather than chosen. The first version of this
policy derived them from the spread of repetitions *inside* one baseline run. Two consecutive
runs of the same commit then reported two scenarios as regressed by 24%. Repetitions inside
one process share a warm cache, a warm JIT, one thermal state and one set of background
neighbours, so their spread describes the machine at one moment and badly understates how far
the same benchmark moves when it is run again.

Measuring five runs of a single commit on an ordinary developer workstation:

| Metric | Between-run movement |
| --- | --- |
| Scenario median wall-clock time | 9% to 21% |
| Scenario median allocated bytes | 0.016% to 0.052% |

Allocated bytes are roughly three orders of magnitude steadier, because allocation is a
property of the code and wall-clock time is a property of the machine's mood. Even with a
five-run baseline at three standard deviations, wall clock still produced a 46% false
regression comparing a commit against itself.

So:

- **Allocation regressions block.** The limit is three times the baseline's between-run
  allocation spread, with a 5% floor, which is roughly a hundred times the measured noise.
- **Timing regressions are reported and never block.** A `slower` row is a real signal worth
  investigating; it is not evidence of a regression on this class of hardware.

A timing gate here would fail for no reason, and a gate that fails for no reason is one people
learn to ignore. If a machine with pinned cores and a controlled thermal envelope is available
later, timing can be promoted to a gate on that machine, and the recorded between-run spreads
are what should justify it.

### Protocol cost is gated, and call counts exactly

The MCP path adds two metrics steadier than anything the direct path offers. Measured across
ten iterations against an installed server:

| Metric | Variation between iterations |
| --- | --- |
| Tool calls | none, identical every time |
| Bytes sent | none, identical every time |
| Bytes received | about 1.2%, because generated content differs slightly |

So the number of round trips is compared **exactly**: any increase blocks, because an adapter
that needs more calls to do the same work has cost every one of its users something real, and
no timing figure on a noisy machine would have shown it. Byte totals get a threshold, with a 5%
floor that comfortably covers the measured 1.2%. A decrease in calls is reported as
`fewer-calls` and does not block.

Other statuses: `inputs-changed` and `missing` block, because a scenario compared against
different fixtures, or deleted from the candidate, must not be able to look like an
improvement. `new` does not block, because there is nothing to compare against and calling it
either a pass or a regression would be a guess.

## What these numbers are not

- Not a comparison with any other library. Nothing here measures another product, so no
  relative claim is supported.
- Not adoption evidence. Performance is distinct from whether anyone chose the library; see
  [adoption validation](adoption-validation.md) for what that evidence is and is not.
- Not a guarantee for your documents. The corpus is generated, fictional and small, and real
  documents differ in ways that matter. Run this against your own inputs.
- Not a memory limit measurement. Per-iteration allocated bytes are attributable to the
  scenario; the peak working set recorded in the raw JSON is process-level and cumulative
  across the whole run, and it is not a per-scenario peak.
