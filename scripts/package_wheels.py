#!/usr/bin/env python3
"""Build the officeagent-mcp Python wheels from a release's standalone server archives.

Each wheel carries one platform's self-contained executable, taken byte for byte from the
release archive after that archive is checked against the release's SHA256SUMS. Nothing is
rebuilt, so a wheel installs exactly the binary the release attested. The Python code in the
wheel is only the launcher in packaging/pypi.

    python scripts/package_wheels.py --artifacts release --version 1.2.0 --output dist
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import io
import re
import sys
import tarfile
import zipfile
from pathlib import Path

from release_evidence import (
    DEFAULT_INVENTORY,
    load_inventory,
    standalone_archive,
    standalone_archive_member,
)

ROOT = Path(__file__).resolve().parents[1]
PACKAGE_SOURCE = ROOT / "packaging" / "pypi"
REPOSITORY = "https://github.com/ilia-sokolov/OfficeAgent.NET"
DISTRIBUTION = "officeagent-mcp"
MODULE = "officeagent_mcp"
ZIP_TIMESTAMP = (1980, 1, 1, 0, 0, 0)

# Wheel platform tags, from the oldest systems .NET 10 supports for each build: glibc 2.27
# on Linux, macOS 15. The Linux builds are glibc only; there is no musl (Alpine) wheel.
PLATFORM_TAGS = {
    "win-x64": "win_amd64",
    "win-arm64": "win_arm64",
    "osx-x64": "macosx_15_0_x86_64",
    "osx-arm64": "macosx_15_0_arm64",
    "linux-x64": "manylinux_2_27_x86_64",
    "linux-arm64": "manylinux_2_27_aarch64",
}


class PackagingError(Exception):
    pass


def python_version(version: str) -> str:
    """The PEP 440 form of a release version: 1.2.0 stays, 1.2.0-rc.3 becomes 1.2.0rc3."""
    match = re.fullmatch(r"(\d+\.\d+\.\d+)(?:-rc\.(\d+))?", version)
    if not match:
        raise PackagingError(f"unsupported release version {version!r}: expected 1.2.3 or 1.2.3-rc.4")
    return match.group(1) + (f"rc{match.group(2)}" if match.group(2) else "")


def read_checksums(path: Path) -> dict[str, str]:
    if not path.is_file():
        raise PackagingError(f"{path.name} is missing; every archive is checked against it")
    sums = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        digest, _, name = line.partition("  ")
        if digest and name:
            sums[name.strip().lstrip("*")] = digest.strip().lower()
    return sums


def archive_files(archive: Path, members: list[str]) -> dict[str, bytes]:
    """The named members of a release archive (.zip or .tar.gz)."""
    found: dict[str, bytes] = {}
    if archive.name.endswith(".zip"):
        with zipfile.ZipFile(archive) as source:
            for member in members:
                try:
                    found[member] = source.read(member)
                except KeyError:
                    pass
    else:
        with tarfile.open(archive, "r:gz") as source:
            for member in members:
                try:
                    entry = source.getmember(member)
                except KeyError:
                    continue
                handle = source.extractfile(entry) if entry.isfile() else None
                if handle is not None:
                    found[member] = handle.read()
    return found


def record_hash(data: bytes) -> str:
    digest = hashlib.sha256(data).digest()
    return "sha256=" + base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")


def metadata(version: str, tag: str) -> str:
    readme = (PACKAGE_SOURCE / "README.md").read_text(encoding="utf-8")
    return "\n".join([
        "Metadata-Version: 2.4",
        f"Name: {DISTRIBUTION}",
        f"Version: {version}",
        "Summary: MCP server for inspecting and editing Word, PowerPoint and Excel files, with no .NET install",
        "Author: Ilia Sokolov",
        "License-Expression: MIT",
        "License-File: LICENSE",
        f"Project-URL: Homepage, {REPOSITORY}",
        f"Project-URL: Documentation, {REPOSITORY}/blob/{tag}/docs/mcp-server.md",
        f"Project-URL: Changelog, {REPOSITORY}/blob/{tag}/CHANGELOG.md",
        f"Project-URL: Issues, {REPOSITORY}/issues",
        "Keywords: mcp,model-context-protocol,word,docx,excel,xlsx,powerpoint,pptx,office,tracked-changes",
        "Classifier: Development Status :: 5 - Production/Stable",
        "Classifier: Intended Audience :: Developers",
        "Classifier: Operating System :: Microsoft :: Windows",
        "Classifier: Operating System :: MacOS",
        "Classifier: Operating System :: POSIX :: Linux",
        "Classifier: Programming Language :: Python :: 3",
        "Classifier: Topic :: Office/Business :: Office Suites",
        "Requires-Python: >=3.9",
        "Description-Content-Type: text/markdown",
        "",
        readme,
    ])


def build_wheel(archive: Path, rid: str, release_version: str, inventory: dict, output: Path) -> Path:
    standalone = inventory["standalone"]
    version = python_version(release_version)
    tag = f"v{release_version}"
    platform_tag = PLATFORM_TAGS[rid]

    executable_member = standalone_archive_member(standalone, rid)
    license_member = executable_member.rsplit("/", 1)[0] + "/LICENSE"
    found = archive_files(archive, [executable_member, license_member])
    if executable_member not in found:
        raise PackagingError(f"{archive.name} has no {executable_member}")
    if license_member not in found:
        raise PackagingError(f"{archive.name} has no {license_member}")
    executable_name = executable_member.rsplit("/", 1)[1]

    init = (PACKAGE_SOURCE / MODULE / "__init__.py").read_text(encoding="utf-8")
    if "__VERSION__" not in init:
        raise PackagingError("the launcher's __init__.py has no __VERSION__ placeholder")
    dist_info = f"{MODULE}-{version}.dist-info"
    files: list[tuple[str, bytes, int]] = [
        (f"{MODULE}/__init__.py", init.replace("__VERSION__", version).encode("utf-8"), 0o644),
        (f"{MODULE}/__main__.py", (PACKAGE_SOURCE / MODULE / "__main__.py").read_bytes(), 0o644),
        (f"{MODULE}/bin/{executable_name}", found[executable_member], 0o755),
        (f"{dist_info}/METADATA", metadata(version, tag).encode("utf-8"), 0o644),
        (f"{dist_info}/WHEEL", "\n".join([
            "Wheel-Version: 1.0",
            "Generator: OfficeAgent.NET scripts/package_wheels.py",
            "Root-Is-Purelib: false",
            f"Tag: py3-none-{platform_tag}",
            "",
        ]).encode("utf-8"), 0o644),
        (f"{dist_info}/entry_points.txt",
         f"[console_scripts]\n{DISTRIBUTION} = {MODULE}.__main__:main\n".encode("utf-8"), 0o644),
        (f"{dist_info}/licenses/LICENSE", found[license_member], 0o644),
    ]
    record = "".join(f"{name},{record_hash(data)},{len(data)}\n" for name, data, _ in files)
    record += f"{dist_info}/RECORD,,\n"
    files.append((f"{dist_info}/RECORD", record.encode("utf-8"), 0o644))

    output.mkdir(parents=True, exist_ok=True)
    wheel = output / f"{MODULE}-{version}-py3-none-{platform_tag}.whl"
    with zipfile.ZipFile(wheel, "w") as target:
        for name, data, mode in files:
            info = zipfile.ZipInfo(name, ZIP_TIMESTAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            # Unix as the creator system, so installers keep the executable bit.
            info.create_system = 3
            info.external_attr = (0o100000 | mode) << 16
            target.writestr(info, data, compresslevel=9)
    return wheel


def build(artifacts: Path, release_version: str, output: Path, runtimes: list[str] | None = None) -> list[Path]:
    inventory = load_inventory(DEFAULT_INVENTORY)
    standalone = inventory.get("standalone")
    if not standalone:
        raise PackagingError("the release inventory ships no standalone server")
    available = standalone["runtimes"]
    unknown = sorted(set(available) - set(PLATFORM_TAGS))
    if unknown:
        raise PackagingError(f"no wheel platform tag for {', '.join(unknown)}")
    selected = runtimes or available
    for rid in selected:
        if rid not in available:
            raise PackagingError(f"the release inventory ships no standalone build for {rid}")

    sums = read_checksums(artifacts / "SHA256SUMS")
    wheels = []
    for rid in selected:
        archive = artifacts / standalone_archive(standalone, rid)
        if not archive.is_file():
            raise PackagingError(f"missing standalone archive: {archive.name}")
        expected = sums.get(archive.name)
        if expected is None:
            raise PackagingError(f"SHA256SUMS has no entry for {archive.name}")
        actual = hashlib.sha256(archive.read_bytes()).hexdigest()
        if actual != expected:
            raise PackagingError(f"{archive.name} does not match SHA256SUMS ({actual} != {expected})")
        wheels.append(build_wheel(archive, rid, release_version, inventory, output))
    return wheels


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--artifacts", type=Path, required=True,
                        help="directory with the release's standalone archives and SHA256SUMS")
    parser.add_argument("--version", required=True, help="release version, for example 1.2.0")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--runtime", action="append", help="build only this runtime (repeatable)")
    args = parser.parse_args()
    try:
        for wheel in build(args.artifacts, args.version, args.output, args.runtime):
            print(f"Packaged {wheel} ({wheel.stat().st_size // (1024 * 1024)} MB)")
    except PackagingError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
