"""
Guards the directories whose files are compared byte for byte.

This defect class has been found twice. The acceptance corpus was silently rewritten on every
Windows checkout, which the platform matrix caught only because the corpus hashes are asserted
in a test; the agent-selection evaluation had exactly the same problem and was missed by the
first audit, because that audit searched for SHA-256 comparisons and the evaluation test
compares file bytes directly.

The trap is that `git status` reports the tree as clean either way, because git normalises on
the way back in. So nothing shows up until an assertion fails on one platform and passes on
another.

This test does not care how any particular test checks its fixtures. It asserts the property
those tests depend on: for the directories declared byte exact in `.gitattributes`, what is on
disk is what was committed.
"""

from __future__ import annotations

import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# Directories whose contents are regenerated and compared against the committed bytes.
BYTE_EXACT = ("tests/OfficeAgent.Tests/Corpus", "evaluations")


def git(*args: str) -> bytes:
    return subprocess.run(
        ["git", *args], cwd=ROOT, check=True, capture_output=True
    ).stdout


class ByteExactFixtureTests(unittest.TestCase):
    def test_gitattributes_declares_every_byte_exact_directory(self) -> None:
        """A rule removed as apparent dead configuration must fail here, not in CI."""
        attributes = (ROOT / ".gitattributes").read_text(encoding="utf-8")
        for directory in BYTE_EXACT:
            self.assertIn(
                f"{directory}/** -text",
                attributes,
                f"{directory} is compared byte for byte but is not protected from end-of-line "
                "translation",
            )

    def test_the_working_tree_matches_the_committed_bytes(self) -> None:
        """
        The property every byte-for-byte fixture comparison depends on.

        A failure here means the checkout rewrote a fixture, so any test that regenerates one
        and compares it will fail on this platform and pass on another.
        """
        mangled: list[str] = []
        for directory in BYTE_EXACT:
            listed = git("ls-files", "--", directory).decode("utf-8").splitlines()
            self.assertTrue(listed, f"{directory} tracks no files; the guard would be vacuous")
            for name in listed:
                path = ROOT / name
                if not path.exists():
                    continue
                committed = git("cat-file", "-p", f"HEAD:{name}")
                if committed != path.read_bytes():
                    mangled.append(name)

        self.assertEqual(
            [],
            mangled,
            "these files differ from their committed bytes, which git reports as a clean tree; "
            "check .gitattributes and re-check out the files",
        )


if __name__ == "__main__":
    unittest.main()
