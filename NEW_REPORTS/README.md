# NEW_REPORTS – GT IBKR full-history report

Builds one Excel workbook in the `FLEX_FORMATTING.xlsx` layout
(Dashboard, Pending Order, Trades_Status, Open Position, All Trades, Trade Summary).

| Data | Source |
|---|---|
| Every fill, commission, realized PnL (FIFO, net of commission) | IBKR Activity Flex Query `GT_FLEX_QUERY_ID` |
| FX pair realized PnL | Flex `FxTransactions` section (IBKR reports 0 on the trade itself) |
| CFD financing charges | Flex `CFDCharges` section |
| Pending orders, trigger / limit / offset / SL %, cancelled orders, unrealized PnL | TWS / IB Gateway (optional) |
| Today's fills (after the last Flex day), with commission | TWS / IB Gateway executions (optional) |

## Run

```
cp .env.example .env      # then fill in the token / query id
bash run.sh               # Linux / EC2
.\run.ps1                 # Windows
```

Each run writes both outputs:

- `reports/GT_Report_<from>_to_<to>.xlsx` – full history, with live TWS data
- `reports/daily/GT_Report_<day>.xlsx` – one workbook per trading day

Options (pass after `run.sh` / `run.ps1`):

```
--from 2026-10-01 --to 2026-10-31   only that period
--no-tws                            Flex only
--xml reports/raw/flex_....xml      rebuild from a saved Flex download
```

Reports go to `reports/`; the raw Flex XML is kept in `reports/raw/`.

## Notes

- Flex Activity data is end-of-day: the latest day included is IBKR's last completed day. Fills after that
  come live from IB Gateway (`reqExecutions`, with IBKR's commission) and are marked
  `executed (TWS live, not yet in Flex)`. The next day Flex has them and they are replaced by the Flex
  figures (matched by execution id, so nothing is counted twice). Gateway forgets executions after its
  daily restart, so they are also kept in `data/tws_executions_cache.json` until Flex has them.
- SL % = |entry price - stop trigger| / entry price. The stop is found by its bracket tag even while it is
  still pending; when Flex has no order price for an older entry, its average fill price is used.
- Metals are booked in IBKR's `...F` sub-account in Flex but under the main account in TWS; both are
  treated as one account when matching orders and positions (the Account column shows the real one).
- Each default run replaces the previous full-history report (the name carries the end date); reports
  made with `--from` / `--to` are kept. `fetch_ec2_reports.bat` mirrors that on the PC.
- Gateway execution times are read in the time zone from `~/Jts/jts.ini` (`TimeZone=`), or `TWS_TIMEZONE` in `.env`.
- The Flex query covers the last 365 days. Use `--from` for anything older; the script downloads it in 365-day chunks.
- Flex has no order prices, so trigger, limit, offset and SL % come from TWS. TWS only remembers recent
  orders, so each run saves them in `data/tws_orders_cache.json`. Schedule a daily run on EC2 so the history builds up, e.g. with cron:
  `0 6 * * 2-6 cd /home/ubuntu/IBKR_REPORT/NEW_REPORTS && mkdir -p reports && bash run.sh >> reports/cron.log 2>&1`
- Sr No comes from the algo's `orderRef` tags: `ENTRY_..._c12_n1_s18a8` gets `N` and its `BR_..._c12_n1_s18a8` stop gets `N.1`.
  Untagged orders (manual, ActivityMonitor) join the position cycle they trade in.
- Flex times are US/Eastern (`FLEX_TIMEZONE` in `.env`), shown converted to UTC and GST.
