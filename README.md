# AI Ops Agent — Excel-to-Deck for Client Portfolio Reviews

A small Python agent that turns a structured Excel tracker into an executive-ready
PowerPoint deck: one portfolio summary slide plus one performance scorecard per
enterprise client.

## Why I built this

As an interim operations advisor and former Chief of Staff, a recurring part of my job
is keeping leadership current on business performance — and that usually means turning
a spreadsheet someone maintains by hand into a deck someone else has to read. Rebuilding
the same slides every reporting cycle, copying numbers across and re-checking what
changed, is exactly the kind of repetitive work that's worth automating.

I built an early version of this using Claude Code for a real engagement — an agent that
pulled business performance data from Excel and updated a management presentation
automatically, saving a few hours every reporting cycle. This repo is a generalized,
from-scratch rebuild of that idea for a fictional consulting company, using synthetic
data, so I can show how it works without touching anything confidential.

## What it does

**Input:** `client_performance_data.xlsx` — a `Clients` sheet (one row per account:
revenue, margin, headcount, utilization, delivery status, NPS, eNPS, attrition, contract
dates...) and a `Company_Summary` sheet that rolls those up with live formulas.

**Output:** `Meridian_Advisory_Portfolio_Review.pptx` — a 5-slide deck:
1. **Portfolio summary** — KPI cards and a flagged-accounts callout, computed from the
   client-level data (not from the Excel formulas — the agent recomputes the rollups
   itself, so it doesn't depend on a spreadsheet engine having recalculated the file)
2–5. **One scorecard per client** — financials, resourcing, delivery & risk,
   relationship health — with engagement-model-aware logic: utilization is only shown
   for staff-augmentation accounts (the client manages delivery), and delivery status is
   only shown for project-based accounts (we own delivery).

The agent also auto-flags accounts that need attention: contract renewal due within 6
months, an NPS/CSAT score last measured more than 6 months ago, or account-team
attrition above 15%.

## Tech stack

Python, [`openpyxl`](https://openpyxl.readthedocs.io/) (reads the Excel input),
[`python-pptx`](https://python-pptx.readthedocs.io/) (writes the PowerPoint output).

## Project structure

```
AI Ops Agent/
├── client_performance_data.xlsx        # sample input (synthetic data)
├── Meridian_Advisory_Portfolio_Review.pptx   # generated output
├── requirements.txt
└── scripts/
    ├── build_sample_data.py            # generates the sample Excel input
    └── build_deck.py                   # the agent — reads Excel, writes the deck
```

## How to run it

```bash
pip install -r requirements.txt

# optional — regenerates client_performance_data.xlsx from scratch
python3 scripts/build_sample_data.py

# reads client_performance_data.xlsx, writes the .pptx
python3 scripts/build_deck.py
```

## Note on the data

**All company and client names, and every figure in this repo, are fictional** —
generated for demonstration purposes. No real client or employer data is used anywhere
in this project.
