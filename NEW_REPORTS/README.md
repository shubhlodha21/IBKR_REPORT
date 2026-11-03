# NEW_REPORTS – GT IBKR full-history report

Builds one Excel workbook in the `FLEX_FORMATTING.xlsx` layout
(Dashboard, Pending Order, Trades_Status, Open Position, All Trades, Trade Summary).

| Data | Source |
|---|---|
| Every fill, commission, realized PnL (FIFO, net of commission) | IBKR Activity Flex Query `GT_FLEX_QUERY_ID` |
| FX pair realized PnL | Flex `FxTransactions` section (IBKR reports 0 on the trade itself) |
| CFD financing charges | Flex `CFDCharges` section |
| Pending orders, trigger / limit / offset / SL %, cancelled orders, unrealized PnL | TWS / IB Gateway (optional) |

## Run

```
cp .env.example .env      # then fill in the token / query id
bash run.sh               # Linux / EC2
.\run.ps1                 # Windows
```

Options (pass after `run.sh` / `run.ps1`):

```
--from 2026-10-01 --to 2026-10-31   only that period
--daily                             also one workbook per trading day (reports/daily/)
--no-tws                            Flex only
--xml reports/raw/flex_....xml      rebuild from a saved Flex download
```

Reports go to `reports/`; the raw Flex XML is kept in `reports/raw/`.

## Notes

- Flex Activity data is end-of-day: the latest day included is IBKR's last completed day.
- The Flex query covers the last 365 days. Use `--from` for anything older; the script downloads it in 365-day chunks.
- Flex has no order prices, so trigger, limit, offset and SL % come from TWS. TWS only remembers recent
  orders, so each run saves them in `data/tws_orders_cache.json`. Schedule a daily run on EC2 so the history builds up, e.g. with cron:
  `30 1 * * * cd /home/ubuntu/IBKR_REPORING_TOOL/NEW_REPORTS && bash run.sh >> reports/cron.log 2>&1`
- Sr No comes from the algo's `orderRef` tags: `ENTRY_..._c12_n1_s18a8` gets `N` and its `BR_..._c12_n1_s18a8` stop gets `N.1`.
  Untagged orders (manual, ActivityMonitor) join the position cycle they trade in.
- Flex times are US/Eastern (`FLEX_TIMEZONE` in `.env`), shown converted to UTC and GST.
