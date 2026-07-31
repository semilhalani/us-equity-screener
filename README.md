# US Equity Screener & Scoring Engine

A personal ETL and scoring system for screening US equities, built entirely on Google Apps Script.

## Summary

This system pulls fundamental and market data for US equities from the Finnhub API, scrapes supplementary data from Finviz, benchmarks each stock against its industry peers using a tiered comparison approach, and scores stocks using a rule based multi factor scoring engine. Google Apps Script enforces a hard 6 minute execution limit per run, so the pipeline is built around a custom chunked execution engine with lock based concurrency control, letting a full data pull and scoring run complete reliably across many short executions instead of one long running process.

## Tech stack

- Google Apps Script (JavaScript)
- Finnhub API for fundamentals and market data
- Finviz, scraped for supplementary data
- Google Sheets as the data store and output layer
- clasp, used to pull this project into version control

## Key technical decisions

**Why a custom chunked execution engine instead of a standard job scheduler**

Google Apps Script kills any execution after 6 minutes with no way to extend it. The universe this pipeline scrapes covers all US common stock, ADR, and EQS tickers across a whitelisted set of exchanges, and the exact count depends on how many pass the index and cap filters at run time rather than being a fixed number in the code. Either way, a full data pull, benchmarking, and scoring run takes far longer than 6 minutes. Instead of relying only on Apps Script's built in time based triggers, the pipeline breaks work into discrete chunks, persists progress between runs using PropertiesService, and uses Apps Script's LockService so overlapping executions cannot corrupt shared state. ChunkEngine.gs also runs an hourly safety net trigger that recreates the per minute worker trigger if it ever gets deleted, so a job cannot get silently stuck.

**Why tier based peer benchmarking**

Finviz's published industry and sector averages blend every market cap together, and a mega cap trading at the same PE as a nano cap does not mean the same thing for both companies. TierBenchmarks.gs computes its own median PE, ROE, margin, growth, and debt to equity per industry or sector, split by tier (Tier1_Core, Tier2_LargeMid, Tier3_Small, Tier4_Moonshot), using medians instead of means so a handful of extreme outliers cannot distort the benchmark. Tiers are assigned in Universe.gs, first by index membership (S&P 500, Nasdaq 100, Dow 30 land in Tier1), then by market cap bucket for everything else, with micro and nano cap stocks routed into their own Moonshot tier that gets scored on a different model entirely instead of being compared to large companies on the same metrics.

**Why a rule based scoring engine rather than an ML model**

The code and comments don't give a reason for this one, so I'm leaving it open rather than guessing. Worth having your own answer ready if it comes up.

## How to explore this code

This is not a typical clone and run repo. It's a Google Apps Script project that depends on a bound Google Sheet, a Finnhub API key, and Google's execution environment to actually run. The code here was pulled from the live Apps Script project using clasp.

There is no single onOpen or menu driven entry point. Each stage of the pipeline has its own run function, called manually or from a time based trigger, and stages depend on each other's output sheets rather than being chained through one master controller. Start with:

- `Universe.gs`, the natural starting point. Its `runUniverseRefresh()` builds the master ticker list and assigns tiers, and every other stage depends on Raw_Universe existing first.
- `Fundamentals.gs`, `Earnings.gs`, `Insiders.gs`, `DisqualifierGates.gs`, and `RefData.gs` for the rest of the data pulls. Each calls the Finviz and Finnhub APIs for its own slice of data: financial ratios, quarterly EPS history, Form 4 insider transactions, hard pass or fail checks like liquidity and dilution, and industry or sector reference averages.
- `ScoringEngine.gs`, the multi factor scoring logic. `runScoringEngineRefresh()` combines Raw_Fundamentals, Raw_TierBenchmarks, Raw_DisqualifierGates, Raw_EPSCategory, and Raw_InsiderSignals into Quality, Value, Growth, Moonshot, Volatility, and Reliability scores, then a final Verdict per ticker.
- `ChunkEngine.gs`, the chunked execution and locking engine described above. `startChunkProcess()`, `runChunkWorker()`, and `autoResumeChunk()` are the core pieces.

## Setup (if you want to run your own copy)

```bash
npm install -g @google/clasp
clasp login
clasp clone <your-script-id>
```

You'll also need:

- A [Finnhub API key](https://finnhub.io/)
- A Google Sheet bound to the script for output

This repo's `Constants.gs` has the real API keys replaced with a placeholder. Add your own Finnhub API key or keys to Script Properties (Apps Script editor, Project Settings, Script Properties), or paste them directly into the `FINNHUB_KEYS` placeholder in `Constants.gs`, before running. Don't commit real keys to a public repo.
