"""Run cellproof on public formula-engine test workbooks. Not a test: run it by hand (needs network and `gh`-free curl).

    python tests/corpus.py DIR

Downloads the .xlsx/.xlsm test files of pycel, formulas and xlcalculator into DIR/corpus (once), converts each into
DIR/runs, and prints one line per workbook plus the totals. The files are public; nothing is uploaded.
"""
import json
import subprocess
import sys
import urllib.request
from collections import Counter
from pathlib import Path

REPOS = {"dgorissen/pycel": "master", "vinci1it2000/formulas": "master", "bradbase/xlcalculator": "master"}


def fetch(corpus: Path) -> None:
    for repo, branch in REPOS.items():
        folder = corpus / repo.replace("/", "_")
        if folder.exists():
            continue
        folder.mkdir(parents=True)
        tree = json.load(urllib.request.urlopen(f"https://api.github.com/repos/{repo}/git/trees/{branch}?recursive=1"))
        for item in tree["tree"]:
            if item["path"].endswith((".xlsx", ".xlsm")):
                url = f"https://raw.githubusercontent.com/{repo}/{branch}/{item['path']}"
                (folder / item["path"].replace("/", "_")).write_bytes(urllib.request.urlopen(url).read())


def main() -> None:
    root = Path(sys.argv[1])
    corpus, runs = root / "corpus", root / "runs"
    fetch(corpus)
    totals, exits = Counter(), Counter()
    for f in sorted(corpus.rglob("*.xls[xm]")):
        out = runs / f"{f.parent.name}__{f.name}"
        done = subprocess.run([sys.executable, "-I", "-m", "cellproof", "convert", str(f), "--out", str(out)],
                              capture_output=True, text=True, timeout=600)
        exits[done.returncode] += 1
        last = (done.stdout + done.stderr).strip().splitlines()[-1:] or [""]
        print(f"{done.returncode} | {f.parent.name}/{f.name} | {last[0][:100]}")
        report = out / "report.json"
        if report.exists():
            totals.update(json.loads(report.read_text(encoding="utf-8"))["counts"])
    print("exit codes:", dict(exits), "cells:", dict(totals))


if __name__ == "__main__":
    main()
