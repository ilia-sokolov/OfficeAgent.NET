#!/usr/bin/env python3
"""Validate server.json against the MCP Registry schema it declares, and the Claude Desktop
bundle manifests against the MCPB manifest schema they use."""

import json
from pathlib import Path
from urllib.request import urlopen

import jsonschema

from package_standalone import bundle_manifest, current_version
from release_evidence import DEFAULT_INVENTORY, load_inventory


# Pinned to a commit so a schema change upstream cannot silently change what passes here.
MCPB_SCHEMA = (
    "https://raw.githubusercontent.com/modelcontextprotocol/mcpb/"
    "257af308122753c311825523d19e8c939aeaccc5/schemas/mcpb-manifest-v0.3.schema.json"
)


def fetch_json(url: str):
    with urlopen(url, timeout=30) as response:
        return json.load(response)


root = Path(__file__).resolve().parents[1]
manifest = json.loads((root / "server.json").read_text(encoding="utf-8"))
jsonschema.validate(manifest, fetch_json(manifest["$schema"]))
print("server.json matches its declared MCP Registry schema.")

standalone = load_inventory(DEFAULT_INVENTORY)["standalone"]
schema = fetch_json(MCPB_SCHEMA)
for rid in standalone["bundles"]:
    bundle = bundle_manifest(standalone, rid, current_version())
    if bundle["manifest_version"] != schema["properties"]["manifest_version"]["const"]:
        raise SystemExit(f"{rid} bundle manifest_version does not match the pinned MCPB schema")
    jsonschema.validate(bundle, schema)
print(f"{len(standalone['bundles'])} MCPB bundle manifests match the MCPB {schema['properties']['manifest_version']['const']} schema.")
