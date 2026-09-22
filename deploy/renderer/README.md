# Sandboxed renderer reference

A tested way to run the optional renderer on untrusted documents. Each render runs in a fresh,
locked-down Linux container. The document goes in on stdin and the page images come out on
stdout, so the container needs no network and no host mount. Rendering stays optional: nothing
outside `OfficeAgent.Rendering` depends on it, and structural preview works without it.

This is a deployment reference, not a supported library type. Copy
[`host/SandboxedDocumentRenderer.cs`](host/SandboxedDocumentRenderer.cs) into your host and
register it as the `IDocumentRenderer`. Discovery then reports `renderingAvailable: true`.

## What it contains

| Path | Purpose |
| --- | --- |
| `Dockerfile` | The image. Base image pinned by digest, every package pinned by version |
| `officeagent-hardening.xcd` | LibreOffice registry layer, finalized so no profile can override it |
| `worker/` | The in-container worker: bounded stdin request, `LibreOfficeDocumentRenderer`, JSON result on stdout |
| `host/SandboxedDocumentRenderer.cs` | The host side: an `IDocumentRenderer` that runs one container per render |
| `test/officeagent-permissive.xcd` | Test only: the control image for the hardening tests. Never deploy it |

## Build

```bash
dotnet publish deploy/renderer/worker -c Release -o deploy/renderer/worker/bin/publish
docker build --target hardened -t officeagent-renderer:reference deploy/renderer
```

Build the image yourself. It contains GPL-licensed Poppler, so redistributing it carries GPL
obligations; see [Provenance](#provenance).

The pins make the build fail rather than drift when Debian supersedes a security update. When
that happens, update the pins, rebuild, and rerun the reference tests as a recorded change.

## Deployment checklist

Each line is enforced by the container arguments in `SandboxedDocumentRenderer.RunArguments`,
and proved by the named test in [`tests/OfficeAgent.RenderSandbox.Tests`](../../tests/OfficeAgent.RenderSandbox.Tests).
Every protection was also removed one at a time to confirm that its test then fails.

| Boundary | How | Proved by |
| --- | --- | --- |
| No network | `--network none` | `Network_is_denied`, with a control that connects on the default network |
| Read-only root, one bounded scratch directory | `--read-only`, `--tmpfs /tmp:size=…` | `Only_the_scratch_directory_is_writable` (mount checked `ro`), `Scratch_space_is_bounded` |
| No host path shared | input on stdin, no `-v` | `No_host_path_is_mounted` |
| Unprivileged | `--user 10001:10001`, `--cap-drop ALL`, `no-new-privileges` | `The_worker_runs_unprivileged` (effective and bounding capabilities empty) |
| Process limit | `--pids-limit` | `The_process_limit_holds`, with a control that fits |
| Memory limit, swap included | `--memory`, `--memory-swap` | `The_memory_limit_is_enforced_by_the_kernel`, with a control that fits |
| CPU quota | `--cpus` | argument only; not measured |
| Time limit and cleanup | host deadline, `docker kill`, `--rm` | `A_timed_out_render_leaves_no_container_behind` |
| Input, page and output bounds | enforced on the host and in the worker | `An_oversized_input_…`, `The_page_limit_holds_…`, `The_output_limit_holds` |
| Only the declared OOXML package reaches LibreOffice | `LibreOfficeDocumentRenderer` checks `[Content_Types].xml` and refuses macros | `Content_that_is_not_the_declared_package_…`, `Bytes_that_are_not_a_document_…` |
| No macro runs, no local file is linked in | finalized `.xcd` layer | `A_document_event_macro_does_not_run`, `A_linked_local_file_…` (see residual risks) |
| Failures carry no document content | stable codes; stderr drained and discarded | `Failure_messages_carry_no_document_content` |

Also set by the deployer and not proved here: run the host's container engine with its own
hardening (a rootless engine or user namespaces, and the default seccomp and AppArmor or
SELinux profiles), keep the renderer host away from service credentials, and never log renderer
output.

## Residual risks

- **Format confusion is closed only at the package check.** LibreOffice picks an import filter
  from content, not the file name. Any bytes it receives are parsed by whichever of its filters
  matches, which is why the renderer refuses anything that is not the declared macro-free OOXML
  package before the backend starts.
- **The registry hardening is not proven effective on this backend.** Macro security and
  "never update links" are set and finalized. On LibreOffice 7.4.7, though, headless
  conversion neither ran a document-event macro nor followed a linked section even in the
  permissive control. So the tests prove the property, that nothing ran and nothing was
  included, but cannot prove the settings are the reason. Treat them as defense in depth.
- **The container is the boundary, not a guarantee against a kernel or engine escape.** It
  protects the host as far as the container runtime does.
- **Pagination is LibreOffice's**, and can differ from Word's; see [rendering](../../docs/rendering.md).

## Portability

Proved on Linux containers only: Docker Desktop's Linux engine on Windows, and GitHub's
`ubuntu-latest` runners. Windows containers, macOS without a Linux VM, other runtimes such as
Podman or containerd, and other architectures such as arm64 have not been verified and are not
claimed.

## Provenance

| Component | Version | Licence |
| --- | --- | --- |
| `mcr.microsoft.com/dotnet/runtime:8.0-bookworm-slim` | `sha256:37466ea190f696105c1c3ae67c15e32d4e199face9a0b2ad5b9a37c464db8f30` (Debian 12.15) | MIT (.NET); Debian packages under their own licences |
| LibreOffice (`libreoffice-*-nogui`) | `4:7.4.7-1+deb12u14`, reporting `LibreOffice 7.4.7.2` | MPL-2.0, Apache-2.0 |
| Poppler (`poppler-utils`) | `22.12.0-2+deb12u3`, reporting `pdftoppm version 22.12.0` | GPL-2 or GPL-3 |
| `fonts-dejavu-core` | `2.37-6` | Bitstream Vera, GPL-2+ |
| `fonts-liberation2` | `2.1.5-1` | SIL OFL 1.1, GPL-2+ |
| `fontconfig` | `2.14.1-4` | permissive (Keith Packard notice) |

Licences were read from `/usr/share/doc/*/copyright` in the built image.
