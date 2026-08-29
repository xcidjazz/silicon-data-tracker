"""Render the published site from the CSV archives.

Reads data/*.csv, inlines them into site/template.html, writes site/index.html.
The template is fully self-contained - no CDN, no external JS.
"""
import csv
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent
D = BASE / "data"
TEMPLATE = BASE / "site" / "template.html"
OUT = BASE / "site" / "index.html"


def rows(name):
    with open(D / name, newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def f(v):
    try:
        return round(float(v), 6)
    except (TypeError, ValueError):
        return None


def build_data():
    data = {}
    llm = rows("llm_token_index.csv")
    gpu = rows("gpu_rental_index.csv")
    data["token"] = {"dates": [r["date"] for r in llm],
                     "series": {k: [f(r[k]) for r in llm]
                                for k in ("llm_token", "open_llm", "proprietary_llm")}}
    gcols = ["h100_neo", "h100_hyper", "a100_neo", "a100_hyper",
             "b200_neo", "mi300x_neo", "h200_neo"]
    data["gpu"] = {"dates": [r["date"] for r in gpu],
                   "series": {k: [f(r[k]) for r in gpu] for k in gcols}}

    fc = rows("forward_curve.csv")
    tenors = sorted([c for c in fc[0] if c.startswith("t")], key=lambda c: int(c[1:]))
    data["fc"] = {"tenors": [int(c[1:]) for c in tenors],
                  "rows": [{"date": r["date"], "gpu": r["gpu"], "rate": r["rate_type"],
                            "v": [f(r[c]) for c in tenors]} for r in fc]}

    def ramp(name, metrics):
        out = {}
        for r in rows(name):
            b = r["breakdown"]
            dim = r["dimension"] or "__single__"
            out.setdefault(b, {}).setdefault(dim, {"months": [], **{m: [] for m in metrics}})
            out[b][dim]["months"].append(r["date_month"][:7])
            for m in metrics:
                out[b][dim][m].append(f(r.get(m)))
        return out

    data["ramp_adoption"] = ramp("ramp_adoption.csv",
                                 ["adoption_rate_pct", "mom_change_pp", "yoy_change_pp"])
    data["ramp_spend"] = ramp("ramp_spend.csv",
                              ["median_pepm", "p90_pepm", "p99_pepm",
                               "top_10_percent_median_pepm", "top_1_percent_median_pepm",
                               "p99_winsorized_weighted_pepm"])
    return data


def main():
    tpl = TEMPLATE.read_text(encoding="utf-8")
    if "/*__DATA__*/" not in tpl:
        raise SystemExit("template is missing the /*__DATA__*/ placeholder")
    data = build_data()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(tpl.replace("/*__DATA__*/", json.dumps(data, separators=(",", ":"))),
                   encoding="utf-8")
    kb = OUT.stat().st_size / 1024
    print(f"built {OUT.name}  {kb:.0f} KB")
    print(f"  token   {data['token']['dates'][0]} -> {data['token']['dates'][-1]}")
    print(f"  gpu     {data['gpu']['dates'][0]} -> {data['gpu']['dates'][-1]}")
    snaps = sorted({r['date'] for r in data['fc']['rows']})
    print(f"  curves  {len(snaps)} snapshot(s), {snaps[0]} -> {snaps[-1]}")
    for k in ("ramp_adoption", "ramp_spend"):
        months = sorted({m for b in data[k].values() for d in b.values() for m in d["months"]})
        print(f"  {k:<15} {months[0]} -> {months[-1]}")


if __name__ == "__main__":
    main()
