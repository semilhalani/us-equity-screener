# US Equity Screener & Scoring Engine

A personal ETL and scoring system for screening US equities, built entirely on Google Apps Script.

## Summary

This system pulls fundamental and market data for US equities from the Finnhub API, scrapes supplementary data from Finviz, benchmarks each stock against its industry peers using a tiered comparison approach, and scores stocks using a rule based multi factor scoring engine. Google Apps Script enforces a hard 6 minute execution limit per run, so the pipeline is built around a custom chunked execution engine with lock based concurrency control, letting a full data pull and scoring run complete reliably across many short executions instead of one long running process.

## Tech stack

- Google Apps Script, written in JavaScript
- Finnhub API for fundamentals and market data
- Finviz, scraped for supplementary data
- Google Sheets as the data store and output layer
- clasp, used to pull this project into version control

## Key technical decisions

**Chunked execution instead of a standard job scheduler.** Apps Script kills any execution after 6 minutes, and a full run across the roughly 5,336 tickers currently in the universe takes far longer than that. The pipeline breaks work into chunks driven by a per minute trigger, tracks progress in Script Properties, and uses LockService so overlapping runs can't corrupt shared state. The full reasoning, plus a real cell limit crash this design had to work around, is in `project-description.md`.

**Tier based peer benchmarking.** A mega cap and a nano cap trading at the same PE do not mean the same thing, so every stock is benchmarked against its own tier and industry or sector median instead of the whole market. How tiers get assigned and how the benchmarks are computed is in `tiers-gates-and-scoring.md`.

**Disqualifier gates run separately from scoring.** Scoring asks how good a stock is, on a scale. Five hard gates ask whether something specifically disqualifies it, a plain yes or no, and a failed gate overrides everything else. One real example: QXO scores 65 on Fundamentals_Score and still ends up disqualified for excessive dilution. Gate by gate detail is in `tiers-gates-and-scoring.md`.

**Rule based scoring rather than an ML model.** The code doesn't explain this one beyond it being the current stage of the project. Moving toward machine learning is a stated future direction, not something built yet, so I'm leaving the why not ML yet question open rather than guessing at a philosophy that isn't written down anywhere.

## Known limitations

This is a rule based, deterministic system today, with no backtesting yet and a handful of honest metric substitutions where Finnhub's free tier falls short of what the ideal metric would be. The full list of what's not built yet and what's still manual is in `project-description.md`.

## How to explore this code

This is not a typical clone and run repo. It's a Google Apps Script project that depends on a bound Google Sheet, a Finnhub API key, and Google's execution environment to actually run. The code here was pulled from the live Apps Script project using clasp.

There is no single onOpen or menu driven entry point. Each stage of the pipeline has its own run function, called manually or from a time based trigger, and stages depend on each other's output sheets rather than being chained through one master controller. Start with:

- `Universe.gs`, the natural starting point. Its `runUniverseRefresh()` builds the master ticker list and assigns tiers, and every other stage depends on Raw_Universe existing first.
- `Fundamentals.gs`, `Earnings.gs`, `Insiders.gs`, `DisqualifierGates.gs`, and `RefData.gs` for the rest of the data pulls. Each calls the Finviz and Finnhub APIs for its own slice of data: financial ratios, quarterly EPS history, Form 4 insider transactions, hard pass or fail checks like liquidity and dilution, and industry or sector reference averages.
- `ScoringEngine.gs`, the multi factor scoring logic. `runScoringEngineRefresh()` combines Raw_Fundamentals, Raw_TierBenchmarks, Raw_DisqualifierGates, Raw_EPSCategory, and Raw_InsiderSignals into Quality, Value, Growth, Moonshot, Volatility, and Reliability scores, then a final Verdict per ticker.
- `ChunkEngine.gs`, the chunked execution and locking engine described above. `startChunkProcess()`, `runChunkWorker()`, and `autoResumeChunk()` are the core pieces.

## Documentation

The `/docs` folder has three companion write ups with more depth than fits comfortably in this README:

- `project-description.md`, the full technical writeup. Covers what the system does end to end, the decision framework behind it, engineering principles, automation and reliability, current limitations, and ideas for where it could go next.
- `tiers-gates-and-scoring.md`, how tiers get assigned, how the five hard pass or fail gates work, and a column by column walkthrough of the scoring engine, all with real examples pulled from a live run.
- `presentation-sheets.md`, how to actually use the four output sheets together to make a decision.

## Setup

If you want to run your own copy:

```bash
npm install -g @google/clasp
clasp login
clasp clone <your-script-id>
```

You'll also need a Finnhub API key, available at https://finnhub.io/, and a Google Sheet bound to the script for output.

This repo's `Constants.gs` has the real API keys replaced with a placeholder. Add your own Finnhub API key or keys to Script Properties under Apps Script's Project Settings, or paste them directly into the `FINNHUB_KEYS` placeholder in `Constants.gs`, before running. Don't commit real keys to a public repo.
