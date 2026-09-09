#!/usr/bin/env python3
"""Validate server.json against the MCP Registry schema it declares."""

import json
from pathlib import Path
from urllib.request import urlopen

import jsonschema


root = Path(__file__).resolve().parents[1]
manifest = json.loads((root / "server.json").read_text(encoding="utf-8"))
with urlopen(manifest["$schema"], timeout=30) as response:
    schema = json.load(response)
jsonschema.validate(manifest, schema)
print("server.json matches its declared MCP Registry schema.")
