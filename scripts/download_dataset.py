#!/usr/bin/env python3
"""Download the MBABench tasks from Hugging Face into the data root and verify them.

    uv run python scripts/download_dataset.py                 # into local.data_root (default: <repo>/data)
    uv run python scripts/download_dataset.py --verify-only   # re-check an existing download
    uv run python scripts/download_dataset.py --dest /else/where

No login is needed for the public dataset; a private or gated copy needs `hf auth login` (or HF_TOKEN).
After the download every file is checked against MANIFEST.json (size + sha256).
Exit code 1 if anything is missing or differs.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "config" / "python"))
from config import Config  # noqa: E402

DEFAULT_REPO = "namkoong-lab/MBABench"
MANIFEST = "MANIFEST.json"


def data_root() -> Path:
    env = os.environ.get("MBABENCH_DATA_ROOT")
    if env:
        return Path(env).expanduser()
    cfg = Config.load(create_missing=False, check_required=False)
    value = cfg.get("local.data_root") or "data"
    path = Path(str(value)).expanduser()
    return path if path.is_absolute() else ROOT / path


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def verify(dest: Path) -> int:
    manifest = dest / MANIFEST
    if not manifest.exists():
        print(f"no {MANIFEST} under {dest}; run without --verify-only first")
        return 1
    entries = json.loads(manifest.read_text())["files"]
    bad = 0
    for e in entries:
        p = dest / e["path"]
        if not p.exists():
            print(f"MISSING  {e['path']}")
            bad += 1
        elif p.stat().st_size != e["size"] or sha256(p) != e["sha256"]:
            print(f"DIFFERS  {e['path']}")
            bad += 1
    tasks = sorted(dest.glob("tasks/task_id=*/task.json"))
    print(f"{len(entries) - bad}/{len(entries)} files verified, {len(tasks)} tasks under {dest}")
    return 1 if bad else 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--repo", default=DEFAULT_REPO, help=f"dataset repo id (default {DEFAULT_REPO})")
    ap.add_argument("--revision", default=None, help="branch, tag or commit (default: main)")
    ap.add_argument("--dest", type=Path, default=None, help="where to put the files (default: local.data_root)")
    ap.add_argument("--verify-only", action="store_true", help="only check an existing download against MANIFEST.json")
    args = ap.parse_args()

    dest = (args.dest or data_root()).resolve()
    if not args.verify_only:
        from huggingface_hub import snapshot_download
        from huggingface_hub.errors import GatedRepoError, RepositoryNotFoundError

        dest.mkdir(parents=True, exist_ok=True)
        print(f"downloading {args.repo} -> {dest}")
        try:
            snapshot_download(repo_id=args.repo, repo_type="dataset", revision=args.revision, local_dir=str(dest))
        except (GatedRepoError, RepositoryNotFoundError) as e:
            print(f"cannot access {args.repo}: {e.__class__.__name__}. Log in with `hf auth login` using an account that can see the dataset.")
            return 1
        for junk in (".cache", ".gitattributes"):
            p = dest / junk
            if p.is_dir():
                import shutil
                shutil.rmtree(p, ignore_errors=True)
            elif p.is_file():
                p.unlink()
    return verify(dest)


if __name__ == "__main__":
    sys.exit(main())
