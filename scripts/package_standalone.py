#!/usr/bin/env python3
"""Publish the MCP server as self-contained executables and Claude Desktop bundles.

Every runtime in the release inventory gets an archive holding one executable that needs no
.NET installation. The Windows and macOS runtimes also get an MCPB bundle, which Claude
Desktop installs in one step and configures with a folder picker.
"""

from __future__ import annotations

import argparse
import gzip
import io
import json
import re
import subprocess
import sys
import tarfile
import tempfile
import zipfile
from pathlib import Path

from release_evidence import (
    DEFAULT_INVENTORY,
    EvidenceError,
    load_inventory,
    single_file_dependencies,
    standalone_archive,
    standalone_archive_member,
    standalone_bundle,
    standalone_bundle_member,
    standalone_executable,
)


ROOT = Path(__file__).resolve().parents[1]
VERSION_RE = re.compile(r"<Version>([^<]+)</Version>")
REPOSITORY = "https://github.com/ilia-sokolov/OfficeAgent.NET"
# 1980-01-01, the earliest time a zip entry can carry; tar entries use the same instant.
ZIP_TIMESTAMP = (1980, 1, 1, 0, 0, 0)
TAR_MTIME = 315532800
# What dotnet publish writes beside the single-file executable that the download leaves out:
# the IIS in-process module, which only IIS loads, and appsettings.json, which the dotnet
# tool does not ship either. Anything else is a file the server may need, so it stops the build.
IGNORED_PUBLISH_FILES = {"appsettings.json", "aspnetcorev2_inprocess.dll"}

# The tools a bundle exposes with its default settings. Allowing creation adds the
# authoring and template tools, which is why the manifest also sets tools_generated.
BUNDLE_TOOLS = [
    ("list_connections", "List the folders you can address documents in."),
    ("describe_capabilities", "Report the formats, edit operations, change modes, and limits this server supports."),
    ("open_document", "Open a document by path and return its inspection in one call."),
    ("inspect_document", "Inspect a Word, PowerPoint, or Excel document."),
    ("find_in_document", "Find text in a Word, PowerPoint, or Excel document."),
    ("preview_plan", "Dry-run an edit plan and report what it would change, without saving."),
    ("apply_plan", "Apply an edit plan and save the document."),
    ("edit_document", "Edit a document by path in one call."),
    ("compare_documents", "Compare two Word documents and return their paragraph differences."),
    ("preview_document_merge", "Preview assembling Word documents into one, without saving."),
    ("register_document", "Register an existing document and return its id."),
    ("remove_document", "Remove a document registration."),
]

# (Mach-O cputype, ELF e_machine, PE machine) for each architecture.
ARCHITECTURES = {
    "x64": (0x01000007, 0x3E, 0x8664),
    "arm64": (0x0100000C, 0xB7, 0xAA64),
}
LC_CODE_SIGNATURE = 0x1D


class PackagingError(RuntimeError):
    pass


def current_version() -> str:
    match = VERSION_RE.search((ROOT / "Directory.Build.props").read_text(encoding="utf-8"))
    if not match:
        raise PackagingError("Directory.Build.props has no Version")
    return match.group(1)


def has_code_signature(content: bytes) -> bool:
    commands = int.from_bytes(content[16:20], "little")
    offset = 32
    for _ in range(commands):
        command = int.from_bytes(content[offset:offset + 4], "little")
        size = int.from_bytes(content[offset + 4:offset + 8], "little")
        if command == LC_CODE_SIGNATURE:
            return True
        if size < 8:
            return False
        offset += size
    return False


def check_executable_format(content: bytes, rid: str) -> None:
    """The executable is in the runtime's format and architecture, and runnable on macOS.

    Apple silicon refuses to run code with no signature at all. The SDK signs macOS
    executables ad hoc, which is enough to run; this check keeps a build that lost the
    signature from reaching users as a binary that dies on launch.
    """
    system, architecture = rid.split("-", 1)
    macho, elf, pe = ARCHITECTURES[architecture]
    if system == "win":
        header = int.from_bytes(content[0x3C:0x40], "little") if content[:2] == b"MZ" else 0
        valid = header > 0 and content[header:header + 4] == b"PE\0\0" and (
            int.from_bytes(content[header + 4:header + 6], "little") == pe)
    elif system == "linux":
        valid = content[:4] == b"\x7fELF" and int.from_bytes(content[18:20], "little") == elf
    else:
        valid = content[:4] == b"\xcf\xfa\xed\xfe" and int.from_bytes(content[4:8], "little") == macho
        if valid and not has_code_signature(content):
            raise PackagingError(f"the {rid} executable carries no code signature")
    if not valid:
        raise PackagingError(f"the {rid} executable is not a {rid} binary")


def publish(dotnet: str, standalone: dict, rid: str, version: str, work: Path) -> bytes:
    project = ROOT / standalone["project"]
    destination = work / "publish" / rid
    command = [
        dotnet, "publish", str(project),
        "--configuration", "Release",
        "--framework", standalone["framework"],
        "--runtime", rid,
        "--self-contained", "true",
        "--output", str(destination),
        # A runtime-specific restore rewrites obj/project.assets.json, which the release SBOM
        # step reads for the NuGet packages. Building under a separate artifacts path leaves the
        # repository's obj folders exactly as the solution build left them.
        "--artifacts-path", str(work / "build"),
        "--nologo",
        "-p:PublishSingleFile=true",
        "-p:DebugType=none",
        "-p:GenerateDocumentationFile=false",
    ]
    if subprocess.run(command, cwd=ROOT, check=False).returncode:
        raise PackagingError(f"dotnet publish failed for {rid}")

    built = project.stem + (".exe" if rid.startswith("win-") else "")
    unexpected = sorted(
        path.name for path in destination.iterdir()
        if path.name != built and path.name not in IGNORED_PUBLISH_FILES
    )
    if unexpected:
        raise PackagingError(
            f"{rid} publish wrote files beside the executable that the download would drop: {unexpected}"
        )
    content = (destination / built).read_bytes()
    check_executable_format(content, rid)
    try:
        dependencies = single_file_dependencies(content, standalone, rid)
    except EvidenceError as exc:
        raise PackagingError(str(exc)) from exc
    stale = sorted(
        f"{name} {value}" for name, value in dependencies.items()
        if name.startswith("OfficeAgent.") and value != version
    )
    if stale:
        raise PackagingError(f"the {rid} executable was built from another version: {stale}")
    return content


def archive_readme(standalone: dict, rid: str, version: str) -> bytes:
    executable = standalone_executable(standalone, rid)
    command = executable if rid.startswith("win-") else f"./{executable}"
    return f"""OfficeAgent MCP server {version} ({rid})

A self-contained build of the OfficeAgent.NET MCP server. Nothing else needs to be installed.

Run it as a stdio server from an MCP client:

    {command} --stdio

Give it a folder of documents with environment variables:

    OfficeAgent__FileSystemConnections__0__ConnectionId=documents
    OfficeAgent__FileSystemConnections__0__RootPath=<absolute path to the folder>
    OfficeAgent__FileSystemConnections__0__AllowedExtensions__0=.docx
    OfficeAgent__FileSystemConnections__0__AllowedExtensions__1=.pptx
    OfficeAgent__FileSystemConnections__0__AllowedExtensions__2=.xlsx
    OfficeAgent__AllowCreation=true      (optional: lets the agent create new files)

With nothing configured, the server keeps documents in memory for the life of the process.

Configuration, security notes, and tool contracts:
{REPOSITORY}/blob/main/docs/mcp-server.md

Check this download against the release's SHA256SUMS and build attestation:
{REPOSITORY}/blob/main/docs/mcp-server.md#standalone-downloads
""".replace("\n", "\r\n" if rid.startswith("win-") else "\n").encode()


def bundle_manifest(standalone: dict, rid: str, version: str) -> dict:
    executable = standalone_executable(standalone, rid)
    extensions = (".docx", ".pptx", ".xlsx")
    env = {
        "OfficeAgent__FileSystemConnections__0__ConnectionId": "documents",
        "OfficeAgent__FileSystemConnections__0__RootPath": "${user_config.documents_folder}",
        "OfficeAgent__AllowCreation": "${user_config.allow_creation}",
    }
    env.update(
        (f"OfficeAgent__FileSystemConnections__0__AllowedExtensions__{index}", extension)
        for index, extension in enumerate(extensions)
    )
    return {
        "manifest_version": "0.3",
        "name": "officeagent",
        "display_name": "OfficeAgent",
        "version": version,
        "description": "Inspect and edit Word, PowerPoint, and Excel files in a folder you choose, previewing every change.",
        "long_description": (
            "OfficeAgent edits Office files in place with structured plans. It inspects a document, "
            "previews each change before saving it, and records Word edits as tracked changes by "
            "default, so you review them in Word before accepting them. It reads and writes only "
            "inside the folder you choose; turn on *Allow creating new documents* to let Claude "
            "create files there too.\n\n"
            f"This is a self-contained build of the open-source [OfficeAgent.NET]({REPOSITORY}) "
            "MCP server, so nothing else needs to be installed."
        ),
        "author": {"name": "Ilia Sokolov", "url": "https://github.com/ilia-sokolov"},
        "repository": {"type": "git", "url": REPOSITORY},
        "homepage": REPOSITORY,
        "documentation": f"{REPOSITORY}/blob/main/docs/mcp-server.md",
        "support": f"{REPOSITORY}/issues",
        "icon": "icon.png",
        "license": standalone["license"],
        "keywords": ["word", "docx", "powerpoint", "pptx", "excel", "xlsx", "office", "tracked changes"],
        "server": {
            "type": "binary",
            "entry_point": standalone_bundle_member(standalone, rid),
            "mcp_config": {
                "command": "${__dirname}/" + standalone_bundle_member(standalone, rid),
                "args": ["--stdio"],
                "env": env,
            },
        },
        "tools": [{"name": name, "description": description} for name, description in BUNDLE_TOOLS],
        "tools_generated": True,
        "compatibility": {"platforms": ["win32" if rid.startswith("win-") else "darwin"]},
        "user_config": {
            "documents_folder": {
                "type": "directory",
                "title": "Documents folder",
                "description": "The folder OfficeAgent may read and edit .docx, .pptx, and .xlsx files in. Nothing outside it is reachable.",
                "required": True,
                "default": "${DOCUMENTS}",
            },
            "allow_creation": {
                "type": "boolean",
                "title": "Allow creating new documents",
                "description": "Let Claude create new documents in the folder, not only edit existing ones.",
                "required": False,
                "default": False,
            },
        },
    }


def write_zip(path: Path, entries: list[tuple[str, bytes, bool]]) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        for name, content, executable in entries:
            info = zipfile.ZipInfo(name, ZIP_TIMESTAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            # Unix as the creator system, so extractors on macOS and Linux honour the mode.
            info.create_system = 3
            info.external_attr = (0o100755 if executable else 0o100644) << 16
            archive.writestr(info, content, compresslevel=9)


def write_tar_gz(path: Path, directory: str, entries: list[tuple[str, bytes, bool]]) -> None:
    def entry(name: str, mode: int) -> tarfile.TarInfo:
        info = tarfile.TarInfo(name)
        info.mode, info.mtime = mode, TAR_MTIME
        info.uid = info.gid = 0
        info.uname = info.gname = ""
        return info

    buffer = io.BytesIO()
    with tarfile.open(fileobj=buffer, mode="w", format=tarfile.USTAR_FORMAT) as archive:
        folder = entry(directory, 0o755)
        folder.type = tarfile.DIRTYPE
        archive.addfile(folder)
        for name, content, executable in entries:
            info = entry(name, 0o755 if executable else 0o644)
            info.size = len(content)
            archive.addfile(info, io.BytesIO(content))
    # No file name and a zero timestamp in the gzip header keep the archive reproducible.
    with path.open("wb") as raw, gzip.GzipFile(
        filename="", mode="wb", fileobj=raw, mtime=0, compresslevel=9
    ) as stream:
        stream.write(buffer.getvalue())


def package_runtime(
    standalone: dict, rid: str, version: str, executable: bytes, output: Path, bundle: bool,
) -> list[Path]:
    license_text = (ROOT / "LICENSE").read_bytes()
    directory = f"{standalone['name']}-{rid}"
    entries = sorted([
        (standalone_archive_member(standalone, rid), executable, True),
        (f"{directory}/LICENSE", license_text, False),
        (f"{directory}/README.txt", archive_readme(standalone, rid, version), False),
    ])
    archive = output / standalone_archive(standalone, rid)
    if rid.startswith("win-"):
        write_zip(archive, entries)
    else:
        write_tar_gz(archive, directory, entries)
    written = [archive]

    if bundle:
        manifest = json.dumps(bundle_manifest(standalone, rid, version), indent=2) + "\n"
        package = output / standalone_bundle(standalone, rid)
        write_zip(package, sorted([
            ("manifest.json", manifest.encode(), False),
            ("icon.png", (ROOT / "icon.png").read_bytes(), False),
            ("LICENSE", license_text, False),
            (standalone_bundle_member(standalone, rid), executable, True),
        ]))
        written.append(package)
    return written


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--version", default=None)
    parser.add_argument("--dotnet", default="dotnet")
    parser.add_argument("--inventory", type=Path, default=DEFAULT_INVENTORY)
    parser.add_argument(
        "--runtime", action="append",
        help="package one runtime identifier; repeat as needed (default: every inventoried runtime)",
    )
    args = parser.parse_args()
    try:
        inventory = load_inventory(args.inventory.resolve())
        standalone = inventory.get("standalone")
        if not standalone:
            raise PackagingError("the release inventory has no standalone entry")
        runtimes = args.runtime or standalone["runtimes"]
        unknown = sorted(set(runtimes) - set(standalone["runtimes"]))
        if unknown:
            raise PackagingError(f"runtimes not in the release inventory: {unknown}")
        version = args.version or current_version()
        output = args.output.resolve()
        output.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix="officeagent-standalone-") as work:
            for rid in runtimes:
                executable = publish(args.dotnet, standalone, rid, version, Path(work))
                for path in package_runtime(
                    standalone, rid, version, executable, output, rid in standalone.get("bundles", [])
                ):
                    print(f"Packaged {path} ({path.stat().st_size // (1024 * 1024)} MB)")
    except (EvidenceError, OSError, PackagingError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
