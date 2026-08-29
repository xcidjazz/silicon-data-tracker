"""Pull the cloud archive into this vault copy, filling any days the laptop missed.

The GitHub Actions job scrapes three times a day regardless of whether this machine
is on. This script brings that archive down and MERGES it into the local CSVs, then
syncs the workbook. It runs before the local scrape, so the order each day is:

    sync_from_repo.py   <- fill gaps from the cloud
    scrape_llm_index.py <- capture anything neither side has yet

Merging is a union keyed on each dataset's natural key. Local-only rows are never
dropped: if the cloud missed a day this machine caught, that day survives. Where both
sides hold the same key and disagree, the cloud value wins and the change is logged -
the cloud run is the one whose cross-checks gate every write.

Run:  python sync_from_repo.py [--dry-run]
"""

import argparse
import csv
import io
import sys

import requests

import scrape_llm_index as S

REPO = "xcidjazz/silicon-data-tracker"
RAW = f"https://raw.githubusercontent.com/{REPO}/main"
API = f"https://api.github.com/repos/{REPO}/contents"
UA = {"User-Agent": "silicon-data-vault-sync"}

# (local path, remote path, key fields, column list)
FILES = [
    (S.DATASETS["llm"]["csv"], "data/llm_token_index.csv", ["date"], S.DATASETS["llm"]["columns"]),
    (S.DATASETS["gpu"]["csv"], "data/gpu_rental_index.csv", ["date"], S.DATASETS["gpu"]["columns"]),
    (S.FC_CSV, "data/forward_curve.csv", ["date", "gpu", "rate_type"], S.FC_COLUMNS),
]
for _which, _spec in S.RAMP_SHEETS.items():
    FILES.append((_spec["csv"], f"data/{_spec['csv'].name}",
                  ["date_month", "breakdown", "dimension"],
                  ["date_month", "breakdown", "dimension"] + _spec["metrics"]))

RAW_DIRS = [(S.FC_RAW_DIR, "data/forward_curve_raw"), (S.RAMP_RAW_DIR, "data/ramp_raw")]


def fetch_csv(remote):
    r = requests.get(f"{RAW}/{remote}", timeout=60, headers=UA)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return list(csv.DictReader(io.StringIO(r.text)))


def load_local(path, keys):
    if not path.exists():
        return {}
    with open(path, newline="", encoding="utf-8") as fh:
        return {tuple(row[k] for k in keys): row for row in csv.DictReader(fh)}


def merge(local, remote_rows, keys, columns):
    """Union of both sides. Returns (merged, n_added, changed[])."""
    added, changed = 0, []
    for row in remote_rows:
        key = tuple(row[k] for k in keys)
        cur = local.get(key)
        if cur is None:
            local[key] = {c: row.get(c, "") for c in columns}
            added += 1
            continue
        for c in columns:
            a, b = (cur.get(c) or "").strip(), (row.get(c) or "").strip()
            if a and b and a != b:
                changed.append(f"{'/'.join(key)} {c}: {a} -> {b} (cloud)")
                cur[c] = b
            elif not a and b:
                cur[c] = b
    return local, added, changed


def sync_raw(local_dir, remote_dir, dry):
    try:
        r = requests.get(f"{API}/{remote_dir}", timeout=60, headers=UA)
        if r.status_code == 404:
            return 0
        r.raise_for_status()
        listing = r.json()
    except Exception as exc:
        S.log(f"  WARN could not list {remote_dir}: {type(exc).__name__}")
        return 0
    pulled = 0
    for entry in listing:
        if entry.get("type") != "file" or not entry["name"].endswith(".json"):
            continue
        dest = local_dir / entry["name"]
        if dest.exists():
            continue
        if dry:
            pulled += 1
            continue
        try:
            body = requests.get(entry["download_url"], timeout=60, headers=UA)
            body.raise_for_status()
            local_dir.mkdir(parents=True, exist_ok=True)
            dest.write_text(body.text, encoding="utf-8")
            pulled += 1
        except Exception as exc:
            S.log(f"  WARN could not pull {entry['name']}: {type(exc).__name__}")
    return pulled


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="report what would change, write nothing")
    args = ap.parse_args()

    S.log(f"=== sync start (from {REPO})")
    total_added, total_changed, touched = 0, 0, False

    for path, remote, keys, columns in FILES:
        try:
            remote_rows = fetch_csv(remote)
        except Exception as exc:
            S.log(f"  WARN {remote}: {type(exc).__name__}: {exc}")
            continue
        if remote_rows is None:
            S.log(f"  {path.name}: not in the repo yet")
            continue
        local = load_local(path, keys)
        before = len(local)
        merged, added, changed = merge(local, remote_rows, keys, columns)
        if len(merged) < before:
            raise ValueError(f"{path.name}: refusing to write, merge shrank {before} -> {len(merged)}")
        total_added += added
        total_changed += len(changed)
        if added or changed:
            touched = True
            if not args.dry_run:
                S.backup(path)
                S.write_keyed_csv(path, columns, merged)
            S.log(f"  {path.name}: +{added} filled from cloud, {len(changed)} revised "
                  f"({before} -> {len(merged)} rows)")
            for c in changed[:10]:
                S.log(f"     REVISED {c}")
        else:
            S.log(f"  {path.name}: already in sync ({before} rows)")

    for local_dir, remote_dir in RAW_DIRS:
        n = sync_raw(local_dir, remote_dir, args.dry_run)
        if n:
            touched = True
            S.log(f"  {remote_dir}: pulled {n} raw snapshot(s)")

    if args.dry_run:
        S.log(f"=== sync dry-run: {total_added} rows would be filled, {total_changed} revised")
        return 0

    if touched:
        try:
            S.backup(S.XLSX_PATH)
            archives = {k: S.load_csv(spec["csv"]) for k, spec in S.DATASETS.items()}
            curves = S.load_fc_csv()
            ramp = {w: S.load_keyed_csv(sp["csv"], ["date_month", "breakdown", "dimension"])
                    for w, sp in S.RAMP_SHEETS.items()}
            for key, (n, mode) in S.update_xlsx(archives, curves, ramp).items():
                if n:
                    sheet = (S.FC_SHEET if key == "fc" else
                             S.RAMP_SHEETS[key]["sheet"] if key in S.RAMP_SHEETS else
                             S.DATASETS[key]["sheet"])
                    S.log(f"  xlsx '{sheet}' {mode}: {n} row(s)")
        except PermissionError:
            S.log("  WARN xlsx locked (open in Excel?) - CSVs synced, sheet catches up next run")

    S.log(f"=== sync ok  {total_added} row(s) filled from cloud, {total_changed} revised")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        S.log(f"=== sync FAILED: {type(exc).__name__}: {exc}")
        if "--quiet" not in sys.argv:
            S.notify_failure(f"Vault sync from {REPO} FAILED\n\n{type(exc).__name__}: {exc}")
        sys.exit(1)
