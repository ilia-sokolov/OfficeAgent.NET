# Dependency audit exceptions

`scripts/check_vulnerable_packages.py` fails CI and release publishing when NuGet reports a
vulnerable direct or transitive dependency. It also fails when the audit command or its JSON
output cannot be read.

Fix or remove the dependency whenever possible. If an upstream fix is unavailable and the
maintainer deliberately accepts the exposure for a bounded period, add an entry to
`dependency-audit-allowlist.json`:

```json
[
  {
    "package": "Example.Package",
    "advisory": "https://github.com/advisories/GHSA-xxxx-xxxx-xxxx",
    "expires": "2026-10-01",
    "reason": "Not reachable in OfficeAgent; upgrade tracked in issue #123"
  }
]
```

The package and advisory URL must exactly match NuGet's audit output. The expiry must be an ISO
date, the reason must name the exposure decision and follow-up, and the exception must be
removed when the finding disappears. Expired and stale exceptions fail the gate.
