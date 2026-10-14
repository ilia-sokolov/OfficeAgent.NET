#!/usr/bin/env python3
"""Install this machine's officeagent-mcp wheel into a fresh virtual environment and use it.

The server is started through the installed `officeagent-mcp` command, so the launcher, the
executable bit and argument passing are all exercised, with no .NET on the path and no network
for the install. It then runs the same document workflow and stdin-EOF check as the standalone
smoke test.

    python scripts/smoke_python_wheel.py --wheels dist --version 1.2.0
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
import venv
from pathlib import Path

from package_wheels import MODULE, PLATFORM_TAGS, python_version
from smoke_packaged_artifacts import (
    SmokeError,
    StdioServer,
    host_runtime,
    verify_document_workflow,
    verify_stdin_eof_exit,
)


def without_dotnet(env: dict[str, str]) -> dict[str, str]:
    env = dict(env)
    for name in ("DOTNET_ROOT", "DOTNET_ROOT_X64", "DOTNET_ROOT_ARM64", "DOTNET_HOST_PATH"):
        env.pop(name, None)
    env["PATH"] = os.pathsep.join(
        entry for entry in env.get("PATH", "").split(os.pathsep)
        if entry and not (Path(entry) / ("dotnet.exe" if os.name == "nt" else "dotnet")).exists()
    )
    return env


def smoke(wheels: Path, release_version: str) -> None:
    rid = host_runtime()
    wheel = wheels / f"{MODULE}-{python_version(release_version)}-py3-none-{PLATFORM_TAGS[rid]}.whl"
    if not wheel.is_file():
        raise SmokeError(f"no wheel for this machine ({rid}): {wheel.name}")

    # Windows can hold the just-exited executable open for a moment, so cleanup may lag.
    with tempfile.TemporaryDirectory(prefix="officeagent-wheel-smoke-", ignore_cleanup_errors=True) as temporary:
        root = Path(temporary)
        environment = root / "venv"
        venv.EnvBuilder(with_pip=True).create(environment)
        scripts = environment / ("Scripts" if os.name == "nt" else "bin")
        python = scripts / ("python.exe" if os.name == "nt" else "python")
        install = subprocess.run(
            [str(python), "-m", "pip", "install", "--no-index", "--no-deps", "--disable-pip-version-check", str(wheel)],
            capture_output=True, text=True, check=False)
        if install.returncode:
            raise SmokeError(f"pip could not install {wheel.name}:\n{install.stderr[-2000:]}")
        print(f"install=passed wheel={wheel.name}")

        command_path = scripts / ("officeagent-mcp.exe" if os.name == "nt" else "officeagent-mcp")
        if not command_path.is_file():
            raise SmokeError(f"installing the wheel created no {command_path.name} command")
        bundled = subprocess.run(
            [str(python), "-c", f"import {MODULE}, os; p = {MODULE}.executable(); print(p.is_file(), os.access(p, os.X_OK), {MODULE}.__version__)"],
            capture_output=True, text=True, check=False).stdout.split()
        if bundled[:2] != ["True", "True"]:
            raise SmokeError(f"the installed server executable is missing or not executable: {bundled}")
        if bundled[2:] != [python_version(release_version)]:
            raise SmokeError(f"the launcher reports version {bundled[2:]}, expected {python_version(release_version)}")

        env = without_dotnet(os.environ)
        command = [str(command_path), "--stdio"]
        workdir = root / "empty"
        workdir.mkdir()
        with StdioServer(command, workdir, env) as server:
            verify_document_workflow(server, release_version)
        verify_stdin_eof_exit(command, workdir, env)
    print("wheel-smoke=passed")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--wheels", type=Path, required=True)
    parser.add_argument("--version", required=True, help="release version, for example 1.2.0")
    args = parser.parse_args()
    try:
        smoke(args.wheels, args.version)
    except SmokeError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
