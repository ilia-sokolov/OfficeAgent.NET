"""Copy a verified native corpus into the repository and render its matrix.

Usage: python scripts/native/publish_results.py <corpus-dir>

Copies inputs, outputs, controls, the manifest and a sanitized results.json into
tests/OfficeAgent.Tests/Corpus/v1.0.0/native, then replaces the generated matrix section of
docs/native-compatibility.md. Local paths are removed from results before they are copied.
"""
import json
import os
import re
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TARGET = os.path.join(ROOT, "tests", "OfficeAgent.Tests", "Corpus", "v1.0.0", "native")
DOC = os.path.join(ROOT, "docs", "native-compatibility.md")
BEGIN, END = "<!-- native-matrix:begin -->", "<!-- native-matrix:end -->"

FORMAT_NAMES = {"word": "Word", "powerpoint": "PowerPoint", "excel": "Excel"}


def sanitize(value, corpus):
    if isinstance(value, dict):
        return {k: sanitize(v, corpus) for k, v in value.items()}
    if isinstance(value, list):
        return [sanitize(v, corpus) for v in value]
    if isinstance(value, str):
        value = value.replace(corpus, "<corpus>").replace(corpus.replace("\\", "/"), "<corpus>")
        return re.sub(r"[A-Za-z]:\\Users\\[^\\\s]+", r"<user>", value)
    return value


def main(corpus):
    corpus = os.path.abspath(corpus)
    results = json.load(open(os.path.join(corpus, "results.json"), encoding="utf-8-sig"))
    manifest = json.load(open(os.path.join(corpus, "manifest.json"), encoding="utf-8"))
    failed = [c["id"] for c in results["cases"] if not c["pass"]]
    missed = [c["file"] for c in results["controls"] if not c["detected"]]
    if failed or missed or len(results["cases"]) != len(manifest):
        sys.exit(f"refusing to publish: failed {failed}, missed controls {missed}, "
                 f"{len(results['cases'])} of {len(manifest)} cases run")

    # Only the generated files are replaced; the corpus README is maintained by hand.
    os.makedirs(TARGET, exist_ok=True)
    for sub in ("inputs", "outputs", "controls"):
        shutil.rmtree(os.path.join(TARGET, sub), ignore_errors=True)
        shutil.copytree(os.path.join(corpus, sub), os.path.join(TARGET, sub))
    shutil.copy(os.path.join(corpus, "manifest.json"), TARGET)
    with open(os.path.join(TARGET, "results.json"), "w", encoding="utf-8", newline="\n") as out:
        json.dump(sanitize(results, corpus), out, indent=2, ensure_ascii=False)
        out.write("\n")

    by_id = {c["id"]: c for c in results["cases"]}
    lines = ["| Format | Operation | Case | Native checks | Visual | Output SHA-256 |",
             "| --- | --- | --- | --- | --- | --- |"]
    for case in manifest:
        r = by_id[case["id"]]
        checks = "; ".join(f"{c['name']}" for c in r["checks"])
        visual = "yes" if case["visual"] else ""
        lines.append(f"| {FORMAT_NAMES[case['format']]} | `{case['family']}` | {case['description']} "
                     f"| {len(r['checks'])} passed: {checks} | {visual} | `{case['outputSha256'][:16]}…` |")
    env = results["environment"]
    summary = (f"{len(manifest)} of {len(manifest)} cases pass and "
               f"{len(results['controls'])} of {len(results['controls'])} corrupt controls are caught, "
               f"on Office {env['officeBuild']}, {env['os']}. Full hashes are in the corpus manifest.")
    block = BEGIN + "\n\n" + summary + "\n\n" + "\n".join(lines) + "\n\n" + END
    doc = open(DOC, encoding="utf-8").read()
    if BEGIN in doc:
        doc = doc[:doc.index(BEGIN)] + block + doc[doc.index(END) + len(END):]
    else:
        doc = doc.replace("RESULTS_TABLE", block)
    open(DOC, "w", encoding="utf-8", newline="\n").write(doc)
    print(f"published {len(manifest)} cases to {TARGET}")


if __name__ == "__main__":
    main(sys.argv[1])
