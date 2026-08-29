"""Daily Telegram summary. Reads the archives and reports what actually moved.

Usage:  python notify.py --url <site url> [--changed 0|1]
Sends nothing unless TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID are set.
"""
import argparse
import csv
import os
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

D = Path(__file__).resolve().parent / "data"


def rows(name):
    p = D / name
    if not p.exists():
        return []
    with open(p, newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def pct(cur, prev):
    if cur is None or prev in (None, 0):
        return ""
    d = (cur / prev - 1) * 100
    return f"  {'+' if d >= 0 else ''}{d:.2f}%"


def line(label, cur, prev, dp=4, unit=""):
    if cur is None:
        return f"{label}: n/a"
    return f"{label}: {cur:.{dp}f}{unit}{pct(cur, prev)}"


def send(text):
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat = os.environ.get("TELEGRAM_CHAT_ID")
    if not (token and chat):
        print("no telegram credentials in env - printing instead:\n")
        print(text)
        return False
    body = urllib.parse.urlencode({"chat_id": chat, "text": text[:4000],
                                   "disable_web_page_preview": "true"}).encode()
    req = urllib.request.Request(f"https://api.telegram.org/bot{token}/sendMessage", data=body)
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.status == 200


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", default="")
    ap.add_argument("--changed", default="1")
    args = ap.parse_args()

    llm = rows("llm_token_index.csv")
    gpu = rows("gpu_rental_index.csv")
    fc = rows("forward_curve.csv")
    ad = rows("ramp_adoption.csv")

    out = ["AI Compute Tape - daily update", ""]

    if llm:
        c, p = llm[-1], (llm[-2] if len(llm) > 1 else {})
        out += [f"TOKEN PRICES  (as of {c['date']}, USD/M tokens)",
                "  " + line("All", num(c.get("llm_token")), num(p.get("llm_token"))),
                "  " + line("Proprietary", num(c.get("proprietary_llm")), num(p.get("proprietary_llm"))),
                "  " + line("Open", num(c.get("open_llm")), num(p.get("open_llm"))), ""]

    if gpu:
        c, p = gpu[-1], (gpu[-2] if len(gpu) > 1 else {})
        out.append(f"GPU RENTAL  (as of {c['date']}, USD/hr, neo-cloud)")
        for key, lbl in [("h100_neo", "H100"), ("h200_neo", "H200"), ("a100_neo", "A100"),
                         ("b200_neo", "B200"), ("mi300x_neo", "MI300X")]:
            out.append("  " + line(lbl, num(c.get(key)), num(p.get(key)), dp=2, unit=""))
        out.append("  " + line("H100 hyperscaler", num(c.get("h100_hyper")),
                               num(p.get("h100_hyper")), dp=2))
        out.append("")

    if fc:
        snaps = sorted({r["date"] for r in fc})
        latest = snaps[-1]
        h = {r["rate_type"]: r for r in fc if r["date"] == latest and r["gpu"] == "H100"}
        if "term_rate" in h:
            spot, m36 = num(h["term_rate"].get("t0")), num(h["term_rate"].get("t36"))
            back = f"  spot {spot:.2f} vs 36m {m36:.2f}" if spot and m36 else ""
            prem = f"  ({(spot/m36-1)*100:+.1f}% backwardation)" if spot and m36 else ""
            out += [f"FORWARD CURVE  (H100 term, as of {latest}){back}{prem}",
                    f"  snapshots archived: {len(snaps)}", ""]

    if ad:
        head = [r for r in ad if r["breakdown"] == "Headline"]
        us = [r for r in ad if r["breakdown"] == "US estimate"]
        if head:
            c = head[-1]
            extra = f"  |  US est {num(us[-1].get('adoption_rate_pct')):.1f}%" if us else ""
            out += [f"BUSINESS AI  (Ramp, {c['date_month'][:7]})",
                    f"  Adoption {num(c.get('adoption_rate_pct')):.1f}%"
                    f"  MoM {num(c.get('mom_change_pp')):+.2f}pp{extra}", ""]

    if args.changed == "0":
        out.append("No new data published since the last run.")
        out.append("")
    if args.url:
        out.append(args.url)
    out.append(f"run {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%MZ')}")
    send("\n".join(out))


if __name__ == "__main__":
    main()
