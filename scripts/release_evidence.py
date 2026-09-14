#!/usr/bin/env python3
"""Generate and verify release SBOMs, checksums, and artifact identity."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INVENTORY = ROOT / ".config" / "release-artifacts.json"
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
VERSION_RE = re.compile(r"^\d+\.\d+\.\d+$")
SHA_RE = re.compile(r"^[0-9a-f]{40}$")


class EvidenceError(ValueError):
    """A fail-closed release evidence validation error."""


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise EvidenceError(f"cannot read JSON {path}: {exc}") from exc


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def load_inventory(path: Path) -> dict[str, Any]:
    value = read_json(path)
    if not isinstance(value, dict) or value.get("schemaVersion") != 1:
        raise EvidenceError("unsupported release artifact inventory schema")
    packages = value.get("packages")
    if not isinstance(packages, list) or not packages:
        raise EvidenceError("release artifact inventory has no packages")
    ids = [package.get("id") for package in packages]
    if any(not isinstance(package_id, str) for package_id in ids) or len(ids) != len(set(ids)):
        raise EvidenceError("release artifact package IDs must be unique strings")
    return value


def sbom_names(package: dict[str, Any], version: str) -> list[str]:
    prefix = f"{package['id']}.{version}"
    return [f"{prefix}.cdx.json"] + [
        f"{prefix}.{framework}.cdx.json" for framework in package["frameworks"]
    ]


def expected_products(inventory: dict[str, Any], version: str) -> list[dict[str, Any]]:
    products: list[dict[str, Any]] = []
    for package in inventory["packages"]:
        inventories = sbom_names(package, version)
        kind = package.get("packageKind", "nuget-package")
        products.extend(
            {
                "name": f"{package['id']}.{version}.{extension}",
                "kind": kind if extension == "nupkg" else "nuget-symbol-package",
                "distribution": "github-original-and-nuget",
                "dependencyInventory": inventories,
            }
            for extension in ("nupkg", "snupkg")
        )
    skill = inventory["skill"]
    products.append(
        {
            "name": skill["archive"],
            "kind": "agent-skill",
            "distribution": "github-original",
            "dependencyInventory": [f"{skill['name']}.{version}.cdx.json"],
        }
    )
    return products


def expected_sboms(inventory: dict[str, Any], version: str) -> list[str]:
    names = [name for package in inventory["packages"] for name in sbom_names(package, version)]
    names.append(f"{inventory['skill']['name']}.{version}.cdx.json")
    return sorted(names)


def run_cyclonedx(
    inventory: dict[str, Any], version: str, output: Path, dotnet: str,
    cyclonedx_dll: Path | None = None, package_ids: set[str] | None = None,
) -> None:
    if not VERSION_RE.fullmatch(version):
        raise EvidenceError("version must use major.minor.patch form")
    known_package_ids = {package["id"] for package in inventory["packages"]}
    if package_ids and not package_ids <= known_package_ids:
        raise EvidenceError(f"unknown package IDs: {sorted(package_ids - known_package_ids)}")
    output.mkdir(parents=True, exist_ok=True)
    tool_prefix = (
        [dotnet, str(cyclonedx_dll.resolve())]
        if cyclonedx_dll
        else [dotnet, "tool", "run", "dotnet-CycloneDX", "--"]
    )
    tool_cwd = ROOT.parent if cyclonedx_dll else ROOT
    version_result = subprocess.run(
        tool_prefix + ["--version"],
        cwd=tool_cwd,
        text=True,
        capture_output=True,
        check=False,
    )
    tool_output = f"{version_result.stdout}\n{version_result.stderr}"
    expected_tool_version = inventory["sbomTool"]["version"]
    if version_result.returncode or expected_tool_version not in tool_output:
        raise EvidenceError(
            f"CycloneDX {expected_tool_version} is required; tool output was {tool_output.strip()!r}"
        )

    for package in inventory["packages"]:
        if package_ids and package["id"] not in package_ids:
            continue
        project = ROOT / package["project"]
        aggregate, *per_framework = sbom_names(package, version)
        variants: list[tuple[str, str | None]] = [(aggregate, None)] + list(
            zip(per_framework, package["frameworks"], strict=True)
        )
        for filename, framework in variants:
            command = tool_prefix + [
                str(project),
                "--output",
                str(output),
                "--filename",
                filename,
                "--output-format",
                "Json",
                "--recursive",
                "--exclude-dev",
                "--disable-package-restore",
                "--no-serial-number",
                "--set-name",
                package["id"],
                "--set-version",
                version,
                "--set-type",
                "Library" if package.get("packageKind") != "dotnet-tool" else "Application",
                "--set-nuget-purl",
            ]
            if framework:
                command.extend(["--framework", framework])
            result = subprocess.run(command, cwd=tool_cwd, check=False)
            if result.returncode:
                raise EvidenceError(f"CycloneDX failed for {package['id']} {framework or 'all frameworks'}")
            bom_path = output / filename
            bom = read_json(bom_path)
            component = bom.get("metadata", {}).get("component")
            if not isinstance(component, dict):
                raise EvidenceError(f"{filename} has no metadata component")
            properties = component.setdefault("properties", [])
            properties.append(
                {
                    "name": "officeagent:targetFramework",
                    "value": framework or "all",
                }
            )
            bom_path.write_text(
                json.dumps(bom, indent=2, sort_keys=True) + "\n", encoding="utf-8"
            )

    skill = inventory["skill"]
    archive = output / skill["archive"]
    if not archive.is_file():
        raise EvidenceError(f"missing skill archive: {archive.name}")
    skill_bom = {
        "bomFormat": "CycloneDX",
        "specVersion": inventory["sbomTool"]["specVersion"],
        "version": 1,
        "metadata": {
            "tools": {
                "components": [
                    {
                        "type": "application",
                        "name": inventory["sbomTool"]["package"],
                        "version": expected_tool_version,
                    }
                ]
            },
            "component": {
                "type": "file",
                "name": skill["name"],
                "version": version,
                "hashes": [{"alg": "SHA-256", "content": sha256(archive)}],
                "licenses": [{"license": {"id": skill["license"]}}],
                "properties": [
                    {"name": "officeagent:targetFramework", "value": "not-applicable"}
                ],
            },
        },
        "components": [],
        "dependencies": [],
    }
    (output / f"{skill['name']}.{version}.cdx.json").write_text(
        json.dumps(skill_bom, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )


def validate_sbom(
    path: Path, expected_name: str, version: str, spec_version: str,
    expected_framework: str,
) -> None:
    value = read_json(path)
    if not isinstance(value, dict) or value.get("bomFormat") != "CycloneDX":
        raise EvidenceError(f"{path.name} is not a CycloneDX SBOM")
    if value.get("specVersion") != spec_version:
        raise EvidenceError(f"{path.name} has unexpected CycloneDX specVersion")
    metadata = value.get("metadata")
    component = metadata.get("component") if isinstance(metadata, dict) else None
    if not isinstance(component, dict):
        raise EvidenceError(f"{path.name} has no metadata component")
    if component.get("name") != expected_name or component.get("version") != version:
        raise EvidenceError(f"{path.name} has incorrect component identity")
    properties = component.get("properties")
    framework_values = {
        item.get("value")
        for item in properties or []
        if isinstance(item, dict) and item.get("name") == "officeagent:targetFramework"
    }
    if framework_values != {expected_framework}:
        raise EvidenceError(f"{path.name} has incorrect target framework identity")
    if not isinstance(value.get("components"), list) or not isinstance(value.get("dependencies"), list):
        raise EvidenceError(f"{path.name} omits components or dependencies")


def create_evidence(
    inventory: dict[str, Any], version: str, source_sha: str, source_ref: str, output: Path
) -> None:
    if not VERSION_RE.fullmatch(version):
        raise EvidenceError("version must use major.minor.patch form")
    if not SHA_RE.fullmatch(source_sha):
        raise EvidenceError("source SHA must be a lowercase 40-character Git SHA")
    expected_ref = f"refs/tags/v{version}"
    if source_ref != expected_ref:
        raise EvidenceError(f"source ref must be {expected_ref}")

    products = expected_products(inventory, version)
    sboms = expected_sboms(inventory, version)
    for product in products:
        path = output / product["name"]
        if not path.is_file():
            raise EvidenceError(f"missing release artifact: {path.name}")
        product["sha256"] = sha256(path)

    package_by_id = {package["id"]: package for package in inventory["packages"]}
    for filename in sboms:
        path = output / filename
        if not path.is_file():
            raise EvidenceError(f"missing SBOM: {filename}")
        if filename.startswith(f"{inventory['skill']['name']}."):
            component_name = inventory["skill"]["name"]
        else:
            component_name = next(
                package_id for package_id in package_by_id if filename.startswith(f"{package_id}.")
            )
        framework = "not-applicable" if component_name == inventory["skill"]["name"] else (
            filename.removeprefix(f"{component_name}.{version}.").removesuffix(".cdx.json")
            if filename != f"{component_name}.{version}.cdx.json"
            else "all"
        )
        validate_sbom(
            path, component_name, version, inventory["sbomTool"]["specVersion"], framework
        )

    release_assets = sorted(
        [product["name"] for product in products]
        + sboms
        + ["release-manifest.json", "SHA256SUMS"]
    )
    manifest = {
        "schemaVersion": 1,
        "release": {
            "version": version,
            "tag": f"v{version}",
            "repository": inventory["repository"],
            "workflow": inventory["workflow"],
            "sourceRef": source_ref,
            "sourceSha": source_sha,
        },
        "generator": {
            "name": "scripts/release_evidence.py",
            "sbomTool": inventory["sbomTool"],
        },
        "artifacts": sorted(products, key=lambda item: item["name"]),
        "verificationMaterials": [
            {"name": name, "kind": "cyclonedx-sbom", "sha256": sha256(output / name)}
            for name in sboms
        ],
        "releaseAssets": release_assets,
        "container": {
            "name": inventory["container"]["name"],
            "workflow": inventory["container"]["workflow"],
            "versionTag": version,
            "identity": "The registry digest, BuildKit SBOM/provenance, and GitHub build attestation are emitted by the container workflow from the same tag.",
        },
        "unsupportedInventoryFields": [
            "Runtime-selected native assets are not predicted beyond restored NuGet graphs.",
            "License fields remain absent when upstream package metadata does not provide them.",
            "NuGet repository signatures are added after upload and are not present in the attached original packages.",
        ],
    }
    manifest_path = output / "release-manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    checksummed = sorted([product["name"] for product in products] + sboms + [manifest_path.name])
    (output / "SHA256SUMS").write_text(
        "".join(f"{sha256(output / name)}  {name}\n" for name in checksummed),
        encoding="utf-8",
        newline="\n",
    )


def parse_checksums(path: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        raise EvidenceError(f"cannot read {path}: {exc}") from exc
    for line in lines:
        parts = line.split("  ", 1)
        if len(parts) != 2 or not SHA256_RE.fullmatch(parts[0]) or not parts[1]:
            raise EvidenceError(f"invalid SHA256SUMS line: {line!r}")
        if parts[1] in result:
            raise EvidenceError(f"duplicate SHA256SUMS entry: {parts[1]}")
        result[parts[1]] = parts[0]
    return result


def verify_evidence(
    inventory: dict[str, Any], output: Path, repository: str, workflow: str,
    source_ref: str, source_sha: str, release_asset_names: set[str] | None = None,
) -> None:
    manifest = read_json(output / "release-manifest.json")
    if not isinstance(manifest, dict) or manifest.get("schemaVersion") != 1:
        raise EvidenceError("unsupported release manifest schema")
    release = manifest.get("release")
    if not isinstance(release, dict):
        raise EvidenceError("release manifest has no release identity")
    version = release.get("version")
    expected_identity = {
        "version": version,
        "tag": source_ref.removeprefix("refs/tags/"),
        "repository": repository,
        "workflow": workflow,
        "sourceRef": source_ref,
        "sourceSha": source_sha,
    }
    if release != expected_identity or not isinstance(version, str) or not VERSION_RE.fullmatch(version):
        raise EvidenceError("release identity does not match expected repository, workflow, ref, SHA, and version")
    if source_ref != f"refs/tags/v{version}":
        raise EvidenceError("expected source ref does not match manifest version")

    expected_product_list = expected_products(inventory, version)
    expected_product_map = {item["name"]: item for item in expected_product_list}
    actual_products = manifest.get("artifacts")
    if not isinstance(actual_products, list):
        raise EvidenceError("release manifest has no artifact list")
    actual_product_map = {
        item.get("name"): item for item in actual_products if isinstance(item, dict)
    }
    if set(actual_product_map) != set(expected_product_map) or len(actual_product_map) != len(actual_products):
        raise EvidenceError("release manifest artifact inventory is incomplete or duplicated")

    sboms = expected_sboms(inventory, version)
    material_map = {
        item.get("name"): item
        for item in manifest.get("verificationMaterials", [])
        if isinstance(item, dict)
    }
    if set(material_map) != set(sboms):
        raise EvidenceError("release manifest SBOM inventory is incomplete")

    for name, expected in expected_product_map.items():
        actual = actual_product_map[name]
        for field in ("kind", "distribution", "dependencyInventory"):
            if actual.get(field) != expected[field]:
                raise EvidenceError(f"{name} has an incorrect {field}")
        if actual.get("sha256") != sha256(output / name):
            raise EvidenceError(f"{name} digest does not match release manifest")
        if not actual["dependencyInventory"]:
            raise EvidenceError(f"{name} has no dependency inventory")
        for sbom in actual["dependencyInventory"]:
            if sbom not in material_map:
                raise EvidenceError(f"{name} references missing SBOM entry {sbom}")

    for sbom in sboms:
        if material_map[sbom].get("sha256") != sha256(output / sbom):
            raise EvidenceError(f"{sbom} digest does not match release manifest")
        component_name = (
            inventory["skill"]["name"]
            if sbom.startswith(f"{inventory['skill']['name']}.")
            else next(package["id"] for package in inventory["packages"] if sbom.startswith(f"{package['id']}."))
        )
        framework = "not-applicable" if component_name == inventory["skill"]["name"] else (
            sbom.removeprefix(f"{component_name}.{version}.").removesuffix(".cdx.json")
            if sbom != f"{component_name}.{version}.cdx.json"
            else "all"
        )
        validate_sbom(
            output / sbom, component_name, version,
            inventory["sbomTool"]["specVersion"], framework,
        )

    expected_assets = sorted(
        list(expected_product_map) + sboms + ["release-manifest.json", "SHA256SUMS"]
    )
    if manifest.get("releaseAssets") != expected_assets:
        raise EvidenceError("release attachment inventory is incomplete or unexpected")
    if release_asset_names is not None and release_asset_names != set(expected_assets):
        missing = sorted(set(expected_assets) - release_asset_names)
        unexpected = sorted(release_asset_names - set(expected_assets))
        raise EvidenceError(f"release attachments differ; missing={missing}, unexpected={unexpected}")

    checksums = parse_checksums(output / "SHA256SUMS")
    expected_checksummed = set(expected_assets) - {"SHA256SUMS"}
    if set(checksums) != expected_checksummed:
        raise EvidenceError("SHA256SUMS inventory is incomplete or unexpected")
    for name, digest in checksums.items():
        path = output / name
        if not path.is_file() or sha256(path) != digest:
            raise EvidenceError(f"checksum verification failed for {name}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--inventory", type=Path, default=DEFAULT_INVENTORY)
    subparsers = parser.add_subparsers(dest="command", required=True)

    sbom = subparsers.add_parser("sbom", help="generate all release SBOMs")
    sbom.add_argument("--version", required=True)
    sbom.add_argument("--artifacts", required=True, type=Path)
    sbom.add_argument("--dotnet", default="dotnet")
    sbom.add_argument(
        "--cyclonedx-dll",
        type=Path,
        help="direct tool DLL path for SDK workarounds; normal releases use the pinned local tool",
    )
    sbom.add_argument(
        "--package",
        action="append",
        help="generate one package ID; repeat as needed (default: all packages)",
    )

    create = subparsers.add_parser("create", help="create manifest and checksums")
    create.add_argument("--version", required=True)
    create.add_argument("--source-sha", required=True)
    create.add_argument("--source-ref", required=True)
    create.add_argument("--artifacts", required=True, type=Path)

    verify = subparsers.add_parser("verify", help="verify downloaded release evidence")
    verify.add_argument("--artifacts", required=True, type=Path)
    verify.add_argument("--repository", required=True)
    verify.add_argument("--workflow", required=True)
    verify.add_argument("--source-ref", required=True)
    verify.add_argument("--source-sha", required=True)
    verify.add_argument(
        "--release-asset-list",
        type=Path,
        help="newline-delimited names returned by the GitHub release API",
    )

    args = parser.parse_args()
    try:
        inventory = load_inventory(args.inventory.resolve())
        artifacts = args.artifacts.resolve()
        if args.command == "sbom":
            run_cyclonedx(
                inventory, args.version, artifacts, args.dotnet, args.cyclonedx_dll,
                set(args.package) if args.package else None,
            )
        elif args.command == "create":
            create_evidence(inventory, args.version, args.source_sha, args.source_ref, artifacts)
        else:
            names = None
            if args.release_asset_list:
                names = set(args.release_asset_list.read_text(encoding="utf-8").splitlines())
            verify_evidence(
                inventory, artifacts, args.repository, args.workflow,
                args.source_ref, args.source_sha, names,
            )
    except (EvidenceError, OSError, StopIteration) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(f"Release evidence {args.command} completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
