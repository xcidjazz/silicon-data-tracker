# AI Compute Tape

Daily capture of AI compute economics from **Silicon Data** and the **Ramp AI Index**.

**Live site:** https://xcidjazz.github.io/silicon-data-tracker/

Built to reproduce the data spine of Citadel Securities' *Elastic Expectations*
(18 Aug 2026) — the argument that falling token prices are expanding compute demand
rather than shrinking it.

## Why this repo exists

**None of these sources publish history.** Silicon Data's token and GPU endpoints
serve a rolling 7-day window; its forward curve serves a single snapshot per day with
no date parameters at all. Once a day rolls off, it is gone from the web permanently.

This repo is the archive. The CSVs under `data/` are the source of truth and cannot be
reconstructed from the internet. Everything else — the site, the workbook — is a render
of them.

## What is captured

| Sheet / file | Series | Cadence | If a run is missed |
|---|---|---|---|
| `llm_token_index.csv` | LLM Token, Open LLM, Proprietary LLM (4dp) | daily | self-heals up to 6 days |
| `gpu_rental_index.csv` | H100 / A100 / H200 / MI300X neo-cloud, H100 / A100 hyperscaler | daily | self-heals up to 6 days |
| `forward_curve.csv` | Term + forward rate, 0–36 months, H100 / A100 / B200 | daily | **permanent loss** |
| `ramp_adoption.csv` | Headline, US estimate, Sector, Business size, Models | monthly | no risk — full history re-served |
| `ramp_spend.csv` | Overall, Business size, Sector, Financing status | monthly | no risk — full history re-served |

Raw payloads are archived verbatim in `data/forward_curve_raw/` (all 145 tenors,
including the interpolated quarter-months the CSV drops) and `data/ramp_raw/`
(one vintage per published month — Ramp restates history).

## Endpoint traps worth knowing

- **Hyperscaler exists for H100 and A100 only.** Requesting `mainTab=hyperscaler` for
  B200/MI300X/H200 does not error — the server returns the *neo-cloud* series. The
  scraper asserts the echoed tab matches the request and refuses to store on mismatch.
- **H200 is hidden** from the public site (filtered unless `standalone=true`) but the
  series is real and populated.
- **The forward curve runs to 36 months.** The public page only ever plots 0–6.
- **Ramp's "90th percentile"** in published commentary means `top_10_percent_median_pepm`,
  not `p90_pepm`. 650/11.95 = 54×; p90/median is only 22×.

## Verification

Every dataset is cross-checked against a second surface before anything is written, and
the run aborts rather than store an unverified number:

- Token: 4dp portal values must round to the 2dp figures on the public page (3/3 required).
- GPU: checked against the "Other Silicon Indices" cards on that same page (4/4 required).
- Forward curve: its tenor-0 node must equal the independently scraped neo-cloud spot
  index — two separate endpoints agreeing on one number.
- Ramp: structural invariants, and `mom_change_pp` must equal the actual level difference.

The archive is also refused if it would ever shrink, and every file is snapshotted to
`backups/` before a write.

## Running it

```bash
pip install requests openpyxl
python scrape_llm_index.py      # all datasets; --only llm|gpu|fc|ramp to narrow
python build_site.py            # renders site/index.html from data/
```

`--rebuild` regenerates the workbook from the CSVs. It discards any sheets or columns
added by hand, so it is for recovery only.

## Automation

`.github/workflows/daily.yml` runs three times a day. GitHub's scheduler is best-effort —
runs are frequently late and occasionally skipped — and the forward curve cannot be
back-filled, so the redundancy is deliberate rather than wasteful.

A Telegram summary goes out on success; the scraper reports its own error on failure, and
the workflow has a backstop alert for everything else that can break a run.

## Keeping a local copy in sync

`sync_from_repo.py` pulls this archive down and merges it into a local vault copy,
filling any days that machine missed while it was off. The merge is a union on each
dataset's natural key: rows only the local machine caught are never dropped, and where
both sides hold the same key and disagree the cloud value wins and the change is logged.

It runs before the local scrape, so the daily order on that machine is sync, then scrape —
covering the case where the cloud missed a day as well as the case where the laptop did.
