r"""
GT IBKR full-history report.

Pulls every execution from the GT Activity Flex Query -- with IBKR's own commission and
FIFO realized PnL for each fill -- and, when TWS / IB Gateway is reachable (EC2), adds the
live pending orders, order prices (trigger / limit), cancelled orders and unrealized PnL.
Writes one workbook in the FLEX_FORMATTING.xlsx layout:

    Dashboard | Pending Order | Trades_Status | Open Position | All Trades | Trade Summary

Usage:
    python gt_flex_report.py                                  # all history in the Flex query
    python gt_flex_report.py --from 2026-10-01 --to 2026-10-31
    python gt_flex_report.py --daily                          # also one workbook per trading day
    python gt_flex_report.py --xml reports/raw/flex_x.xml     # rebuild from a saved Flex XML
    python gt_flex_report.py --no-tws                         # Flex only, do not contact TWS

Settings come from environment variables or the .env file next to this script
(see .env.example): GT_FLEX_TOKEN, GT_FLEX_QUERY_ID, TWS_HOST, TWS_PORT, ...
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import threading
import time
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Settings
# --------------------------------------------------------------------------- #
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_OUT = SCRIPT_DIR / "reports"
ORDER_CACHE = SCRIPT_DIR / "data" / "tws_orders_cache.json"

FLEX_SEND_URL = "https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService/SendRequest"
FLEX_GET_URL = "https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService/GetStatement"
FLEX_VERSION = "3"
FLEX_RETRY_CODES = {"1001", "1009", "1018", "1019", "1021"}  # busy / generating / rate limit
FLEX_MAX_DAYS = 365
POLL_SECONDS = 5
POLL_ATTEMPTS = 36

TZ_GST = ZoneInfo("Asia/Dubai")

# TWS / IB Gateway
TWS_INFO_CODES = {2100, 2103, 2104, 2105, 2106, 2107, 2108, 2119, 2150, 2157, 2158}
CANCELLED_STATUSES = {"Cancelled", "ApiCancelled", "Inactive"}
OPEN_STATUSES = {"PendingSubmit", "ApiPending", "PreSubmitted", "Submitted", "PendingCancel"}

# Bracket tags the algo puts in orderRef, e.g. ENTRY_BUY_100_AMZN_c12_n1_s18a8 / BR_SELL_100_AMZN_c12_n1_s18a8
REF_RE = re.compile(r"^(?P<role>[A-Za-z]+)_(?:BUY|SELL)_[\d.]+_[^_]+_(?P<key>c\d+_n\d+_s[0-9a-z]+)$", re.I)

# Styles (from FLEX_FORMATTING.xlsx)
NAVY = "FF1F3864"
FONT_TITLE = Font(name="Calibri", size=13, bold=True, color="FFFFFFFF")
FONT_HEAD = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
FONT_LABEL = Font(name="Calibri", size=10, bold=True, color=NAVY)
FONT_DATA = Font(name="Calibri", size=10, color="FF262626")
FONT_NOTE = Font(name="Calibri", size=9, italic=True, color="FF595959")
FILL_NAVY = PatternFill("solid", fgColor=NAVY)
FILL_RED = PatternFill("solid", fgColor="FFFFC7CE")
FILL_GREEN = PatternFill("solid", fgColor="FFC6EFCE")
THIN = Side(style="thin", color="FFBFBFBF")
MEDIUM = Side(style="medium", color="FF00B050")
BORDER_DATA = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_TITLE = Border(left=MEDIUM, right=MEDIUM, top=MEDIUM, bottom=MEDIUM)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

FMT_DT = "dd-mmm-yyyy hh:mm:ss"
FMT_MONEY = "#,##0.00"
FMT_PRICE = "#,##0.00####"
FMT_AVG = "#,##0.00000"
FMT_PCT = "0.00%"


def load_env():
    """Read KEY=VALUE (or `export KEY="VALUE"`) lines from .env; real environment variables win."""
    env_file = SCRIPT_DIR / ".env"
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line.startswith("export "):
            line = line[len("export "):].strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def setting(name, default=""):
    return os.environ.get(name, default).strip()


def setting_num(name, default):
    try:
        return float(setting(name, str(default)).replace(",", ""))
    except ValueError:
        return default


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
DATE_FORMATS = ("%Y%m%d", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%b-%y")
TIME_FORMATS = ("%H%M%S", "%H:%M:%S")


def parse_date(text):
    text = (text or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def parse_datetime(text, tz):
    """Flex dateTime such as '20261030;094007' (in the query's time zone) -> aware datetime."""
    text = (text or "").strip()
    if not text:
        return None
    d_part, t_part = text, ""
    for sep in (";", ",", " "):
        if sep in text:
            d_part, t_part = text.split(sep, 1)
            break
    d = parse_date(d_part)
    if d is None:
        return None
    for fmt in TIME_FORMATS:
        try:
            return datetime.combine(d, datetime.strptime(t_part.strip(), fmt).time(), tzinfo=tz)
        except ValueError:
            pass
    return datetime.combine(d, datetime.min.time(), tzinfo=tz)


def parse_tws_time(text, default_tz):
    """TWS completedTime: '20260810 13:07:46 America/New_York' or '20260810-13:07:46' (UTC)."""
    text = (text or "").strip()
    if not text:
        return None
    if "-" in text.split(" ")[0]:
        try:
            return datetime.strptime(text.split(" ")[0], "%Y%m%d-%H:%M:%S").replace(tzinfo=timezone.utc)
        except ValueError:
            return None
    parts = text.split()
    tz = default_tz
    if len(parts) >= 3:
        try:
            tz = ZoneInfo(parts[2])
        except Exception:
            pass
    try:
        return datetime.strptime(" ".join(parts[:2]), "%Y%m%d %H:%M:%S").replace(tzinfo=tz)
    except ValueError:
        return None


def num(text):
    try:
        return float(text) if text not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def price_or_none(v):
    """TWS uses 0 or Double.MAX for 'not set'."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f == 0 or f != f or abs(f) >= 1e300:
        return None
    return f


def clean_num(x, digits=8):
    """Round away float noise; return int when whole, so Excel shows 50000 not 50000.0."""
    if x is None:
        return None
    x = round(x, digits)
    return int(x) if x == int(x) else x


def to_utc(dt):
    return dt.astimezone(timezone.utc).replace(tzinfo=None) if dt else None


def to_gst(dt):
    return dt.astimezone(TZ_GST).replace(tzinfo=None) if dt else None


def ref_key(account, conid, order_ref):
    """(group key, is_entry) from the algo's orderRef, or (None, False)."""
    m = REF_RE.match((order_ref or "").strip())
    if not m:
        return None, False
    return ("ref", account, conid, m.group("key").lower()), m.group("role").upper() == "ENTRY"


# --------------------------------------------------------------------------- #
# Flex Web Service
# --------------------------------------------------------------------------- #
class FlexError(Exception):
    def __init__(self, code, msg):
        super().__init__(f"[{code}] {msg}")
        self.code = code


def _flex_error(root):
    return root.findtext("ErrorCode", ""), root.findtext("ErrorMessage", "")


def download_flex(token, query_id, from_date=None, to_date=None):
    """One Flex statement (the query's own period, or [from_date, to_date]) as XML text."""
    session = requests.Session()
    session.headers["User-Agent"] = "Python/GT-Flex-Report"
    params = {"t": token, "q": query_id, "v": FLEX_VERSION}
    if from_date and to_date:
        params.update(fd=from_date.strftime("%Y%m%d"), td=to_date.strftime("%Y%m%d"))

    for attempt in range(1, 6):
        resp = session.get(FLEX_SEND_URL, params=params, timeout=60)
        resp.raise_for_status()
        root = ET.fromstring(resp.content)
        if root.findtext("Status") == "Success":
            break
        code, msg = _flex_error(root)
        if code not in FLEX_RETRY_CODES or attempt == 5:
            raise FlexError(code, msg)
        print(f"  IBKR busy [{code}] {msg} -- retrying in 15s")
        time.sleep(15)
    ref = root.findtext("ReferenceCode")
    get_url = root.findtext("Url") or FLEX_GET_URL
    print(f"  Flex request accepted (reference {ref}); waiting for the statement...")

    for _ in range(POLL_ATTEMPTS):
        resp = session.get(get_url, params={"t": token, "q": ref, "v": FLEX_VERSION}, timeout=300)
        resp.raise_for_status()
        if "<FlexQueryResponse" in resp.text[:1000]:
            return resp.text
        code, msg = _flex_error(ET.fromstring(resp.content))
        if code not in FLEX_RETRY_CODES:
            raise FlexError(code, msg)
        time.sleep(POLL_SECONDS)
    raise FlexError("timeout", "statement was not ready in time -- please run again in a minute")


def download_all(token, query_id, from_date=None):
    """The query's default period (normally the last 365 days), plus older 365-day chunks
    back to from_date when that is earlier."""
    print("Downloading GT Activity Flex statement ...")
    texts = [download_flex(token, query_id)]
    if from_date:
        first = FlexData.statement_from(texts[0])
        end = (first or date.today()) - timedelta(days=1)
        while from_date <= end:
            start = max(from_date, end - timedelta(days=FLEX_MAX_DAYS - 1))
            print(f"Downloading older history {start} -> {end} ...")
            texts.append(download_flex(token, query_id, start, end))
            end = start - timedelta(days=1)
    return texts


# --------------------------------------------------------------------------- #
# Flex statement model
# --------------------------------------------------------------------------- #
@dataclass
class Fill:
    account: str
    conid: str
    symbol: str
    name: str
    asset: str
    currency: str
    fx: float                 # trade currency -> base
    multiplier: float
    raw_dt: str
    dt: datetime              # aware
    order_dt: Optional[datetime]
    trade_date: date
    side: str                 # BUY / SELL
    qty: float                # absolute
    price: float
    money: float              # absolute trade value in trade currency
    comm_raw: float
    comm_ccy: str
    fifo_pnl: float           # IBKR FIFO realized PnL in trade currency (net of commissions)
    exchange: str
    order_type: str
    perm_id: str
    order_ref: str
    txn_id: str
    auto_fx: bool             # IBKR automatic conversion of a leftover cash balance
    value: float = 0.0        # base currency
    commission: float = 0.0   # base currency (negative = paid)
    pnl: float = 0.0          # base currency
    cycle: Optional[int] = None

    @property
    def signed(self):
        return self.qty if self.side == "BUY" else -self.qty


class FlexData:
    def __init__(self, xml_texts, tz):
        self.tz = tz
        self.account_names = {}
        self.base_ccy = None
        self.from_date = None
        self.to_date = None
        self._fills = {}
        self._fx_txns = {}
        self.cfd_charges = {}          # transactionID -> (account, date, amount base, description)
        self.flex_unrealized = {}      # (account, conid) -> (reportDate, unrealized base)
        for text in xml_texts:
            for st in ET.fromstring(text).iter("FlexStatement"):
                self._load_statement(st)
        self.base_ccy = self.base_ccy or "USD"
        self.fills = sorted(self._fills.values(), key=lambda f: (f.dt, f.txn_id))
        self._finalize()

    @staticmethod
    def statement_from(xml_text):
        dates = [parse_date(st.get("fromDate")) for st in ET.fromstring(xml_text).iter("FlexStatement")]
        dates = [d for d in dates if d]
        return min(dates) if dates else None

    @property
    def accounts(self):
        return sorted(self.account_names)

    def _load_statement(self, st):
        acct = st.get("accountId", "")
        self.account_names.setdefault(acct, "")
        for d, pick in ((parse_date(st.get("fromDate")), min), (parse_date(st.get("toDate")), max)):
            if d:
                attr = "from_date" if pick is min else "to_date"
                cur = getattr(self, attr)
                setattr(self, attr, d if cur is None else pick(cur, d))

        info = st.find("AccountInformation")
        if info is not None:
            if info.get("name"):
                self.account_names[acct] = info.get("name")
            if info.get("currency") and not self.base_ccy:
                self.base_ccy = info.get("currency")

        for e in st.iter("Trade"):
            if (e.get("levelOfDetail") or "EXECUTION").upper() != "EXECUTION":
                continue
            f = self._fill(acct, e.attrib)
            if f:
                self._fills[(f.account, f.txn_id or f"{f.raw_dt}|{f.conid}|{f.signed}")] = f

        for e in st.iter("FxTransaction"):
            a = e.attrib
            key = (acct, a.get("dateTime", ""), a.get("fxCurrency", ""), a.get("quantity", ""),
                   a.get("code", ""), a.get("realizedPL", ""))
            self._fx_txns[key] = a

        for e in st.iter("OpenPosition"):
            a = e.attrib
            if (a.get("levelOfDetail") or "SUMMARY").upper() != "SUMMARY":
                continue
            rd = parse_date(a.get("reportDate")) or date.min
            key = (acct, a.get("conid", ""))
            value = num(a.get("fifoPnlUnrealized")) * (num(a.get("fxRateToBase")) or 1.0)
            if key not in self.flex_unrealized or rd >= self.flex_unrealized[key][0]:
                self.flex_unrealized[key] = (rd, value)

        for e in st.iter("CFDCharge"):
            a = e.attrib
            self.cfd_charges[a.get("transactionID") or str(len(self.cfd_charges))] = (
                acct, parse_date(a.get("date")) or parse_date(a.get("reportDate")),
                num(a.get("total")) * (num(a.get("fxRateToBase")) or 1.0), a.get("activityDescription", ""))

    def _fill(self, acct, a):
        dt = parse_datetime(a.get("dateTime"), self.tz)
        if dt is None:
            return None
        side = (a.get("buySell") or "").upper()
        side = "SELL" if side.startswith("SELL") else "BUY"
        qty = abs(num(a.get("quantity")))
        price = num(a.get("tradePrice"))
        multiplier = num(a.get("multiplier")) or 1.0
        asset = a.get("assetCategory", "")
        return Fill(
            account=a.get("accountId") or acct,
            conid=a.get("conid") or a.get("symbol", ""),
            symbol=a.get("symbol", ""),
            name=a.get("description", "") or a.get("symbol", ""),
            asset=asset,
            currency=a.get("currency", ""),
            fx=num(a.get("fxRateToBase")) or 1.0,
            multiplier=multiplier,
            raw_dt=a.get("dateTime", ""),
            dt=dt,
            order_dt=parse_datetime(a.get("orderTime"), self.tz),
            trade_date=parse_date(a.get("tradeDate")) or dt.date(),
            side=side,
            qty=qty,
            price=price,
            money=abs(num(a.get("tradeMoney"))) or qty * price * multiplier,
            comm_raw=num(a.get("ibCommission")),
            comm_ccy=a.get("ibCommissionCurrency", ""),
            fifo_pnl=num(a.get("fifoPnlRealized")),
            exchange=a.get("exchange", ""),
            order_type=a.get("orderType", ""),
            perm_id=a.get("ibOrderID", "") or "",
            order_ref=a.get("orderReference", "") or "",
            txn_id=a.get("transactionID") or a.get("tradeID") or "",
            auto_fx=asset == "CASH" and not a.get("orderType"),
        )

    def _finalize(self):
        base = self.base_ccy
        fx_closed = defaultdict(float)     # (account, dateTime, currency) -> IBKR realized FX PnL (base)
        for (acct, dt_raw, ccy, _q, code, _r), a in self._fx_txns.items():
            if "C" in code.split(";") and ";" in dt_raw:
                fx_closed[(acct, dt_raw, ccy)] += num(a.get("realizedPL"))

        cash_groups = defaultdict(list)
        for f in self.fills:
            pair = f.symbol.split(".")
            if f.asset == "CASH" and pair[0] == base:
                f.value = f.qty                       # e.g. USD.JPY: quantity is already in USD
            else:
                f.value = f.money * f.fx
            if f.comm_ccy in ("", base):
                f.commission = f.comm_raw
            elif f.comm_ccy == f.currency:
                f.commission = f.comm_raw * f.fx
            else:
                f.commission = f.comm_raw * self.fx_for(f.comm_ccy)
            if f.asset == "CASH":
                cash_groups[(f.account, f.raw_dt, f.symbol)].append(f)
            else:
                f.pnl = f.fifo_pnl * f.fx

        # IBKR reports FX pair PnL in the FxTransactions section (fifoPnlRealized is 0 there)
        for (acct, dt_raw, symbol), fills in cash_groups.items():
            total = sum(fx_closed.get((acct, dt_raw, ccy), 0.0) for ccy in symbol.split(".") if ccy != base)
            qty = sum(f.qty for f in fills)
            for f in fills:
                f.pnl = total * f.qty / qty if qty else 0.0

        # Position cycles: a new cycle starts each time a contract's position leaves flat
        net = defaultdict(float)
        cycle_no = defaultdict(int)
        for f in self.fills:
            if f.auto_fx:
                continue
            key = (f.account, f.conid)
            if abs(net[key]) < 1e-9:
                cycle_no[key] += 1
            f.cycle = cycle_no[key]
            net[key] += f.signed
            if abs(net[key]) < 1e-9:
                net[key] = 0.0

    def fx_for(self, ccy):
        """Latest Flex rate ccy -> base."""
        if not ccy or ccy == self.base_ccy:
            return 1.0
        rates = [f.fx for f in getattr(self, "fills", self._fills.values()) if f.currency == ccy and f.fx]
        return rates[-1] if rates else 1.0

    def name_for(self, conid, fallback):
        for f in reversed(self.fills):
            if f.conid == conid:
                return f.name
        return fallback


# --------------------------------------------------------------------------- #
# TWS / IB Gateway (optional): pending orders, order prices, cancelled orders, unrealized PnL
# --------------------------------------------------------------------------- #
def tws_order(contract, order, status):
    is_cash = contract.secType == "CASH"
    return {
        "perm_id": str(order.permId),
        "order_id": order.orderId,
        "parent_id": order.parentId,
        "account": order.account or "",
        "conid": str(contract.conId),
        "symbol": contract.localSymbol if is_cash and contract.localSymbol else contract.symbol,
        "sec_type": contract.secType,
        "currency": contract.currency,
        "multiplier": num(contract.multiplier) or 1.0,
        "action": order.action,
        "order_type": order.orderType,
        "aux": price_or_none(order.auxPrice),
        "lmt": price_or_none(order.lmtPrice),
        "trail": price_or_none(getattr(order, "trailStopPrice", None)),
        "qty": float(order.totalQuantity),
        "order_ref": order.orderRef or "",
        "status": status,
    }


def fetch_tws(host, port, client_id, wait):
    """Snapshot of TWS: open orders, completed orders and portfolio. None when TWS is not reachable."""
    try:
        from ibapi.client import EClient
        from ibapi.wrapper import EWrapper
    except ImportError:
        print("  ibapi is not installed -- skipping TWS (pip install ibapi).")
        return None

    class TwsClient(EWrapper, EClient):
        def __init__(self):
            EWrapper.__init__(self)
            EClient.__init__(self, wrapper=self)
            self.ready = threading.Event()
            self.open_done = threading.Event()
            self.completed_done = threading.Event()
            self.account_done = threading.Event()
            self.accounts = []
            self.open_orders = {}
            self.completed_orders = {}
            self.portfolio = {}

        def error(self, reqId, *args):
            # (code, msg[, json]) in older APIs, (time, code, msg, json) from API 10.33
            if len(args) >= 3 and isinstance(args[1], int):
                code, msg = args[1], args[2]
            else:
                code, msg = args[0], args[1] if len(args) > 1 else ""
            if code not in TWS_INFO_CODES:
                print(f"  [TWS] {code}: {msg}", file=sys.stderr)

        def nextValidId(self, orderId):
            self.ready.set()

        def managedAccounts(self, accountsList):
            self.accounts = [a.strip() for a in accountsList.split(",") if a.strip()]

        def openOrder(self, orderId, contract, order, orderState):
            self.open_orders[str(order.permId)] = tws_order(contract, order, orderState.status)

        def openOrderEnd(self):
            self.open_done.set()

        def completedOrder(self, contract, order, orderState):
            d = tws_order(contract, order, orderState.status)
            d["completed"] = getattr(orderState, "completedTime", "") or ""
            self.completed_orders[d["perm_id"]] = d

        def completedOrdersEnd(self):
            self.completed_done.set()

        def updatePortfolio(self, contract, position, marketPrice, marketValue, averageCost,
                            unrealizedPNL, realizedPNL, accountName):
            unreal = float(unrealizedPNL)
            self.portfolio[(accountName, str(contract.conId))] = {
                "symbol": contract.localSymbol if contract.secType == "CASH" else contract.symbol,
                "currency": contract.currency,
                "position": float(position),
                "avg_cost": float(averageCost),
                "unrealized": None if abs(unreal) >= 1e300 else unreal,
            }

        def accountDownloadEnd(self, accountName):
            self.account_done.set()

    print(f"Connecting to TWS / IB Gateway at {host}:{port} (client id {client_id}) ...")
    app = TwsClient()
    try:
        app.connect(host, port, client_id)
    except Exception as e:
        print(f"  TWS not reachable ({e}) -- continuing with Flex data only.")
        return None
    if not app.isConnected():
        print("  TWS not reachable -- continuing with Flex data only.")
        return None
    threading.Thread(target=app.run, daemon=True).start()
    try:
        if not app.ready.wait(10):
            print("  TWS did not complete the handshake -- continuing with Flex data only.")
            return None
        time.sleep(0.5)                                 # managedAccounts arrives with the handshake
        app.reqAllOpenOrders()
        app.open_done.wait(wait)
        app.reqCompletedOrders(False)
        app.completed_done.wait(wait)
        for acct in app.accounts:
            app.account_done.clear()
            app.reqAccountUpdates(True, acct)
            app.account_done.wait(wait)
            app.reqAccountUpdates(False, acct)
    finally:
        app.disconnect()
    print(f"  TWS: {len(app.open_orders)} open orders, {len(app.completed_orders)} completed orders, "
          f"{len(app.portfolio)} positions (accounts {', '.join(app.accounts) or '-'})")
    return app


def update_order_cache(tws, path):
    """TWS only remembers recent orders, so every run stores their prices / status here.
    Over time this keeps trigger, limit and cancelled orders for the whole history."""
    cache = {}
    if path.exists():
        try:
            cache = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            print(f"  !! {path.name} is unreadable -- starting a new order cache.")
    if tws is None:
        return cache
    for d in list(tws.open_orders.values()) + list(tws.completed_orders.values()):   # completed wins
        if d["perm_id"] in ("", "0"):
            continue
        old = cache.get(d["perm_id"], {})
        cache[d["perm_id"]] = {**old, **{k: v for k, v in d.items() if v not in (None, "")}}
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(cache, indent=1, sort_keys=True), encoding="utf-8")
    return cache


# --------------------------------------------------------------------------- #
# Orders, groups and serial numbers
# --------------------------------------------------------------------------- #
@dataclass
class Order:
    account: str
    perm_id: str
    conid: str
    symbol: str
    name: str
    side: str
    order_type: str
    qty: float
    value: float
    placed: Optional[datetime]
    done: Optional[datetime]
    status: str
    order_ref: str = ""
    fills: list = field(default_factory=list)
    trigger: Optional[float] = None
    lmt: Optional[float] = None
    numbered: bool = True
    group: tuple = ()
    is_entry: bool = False
    sr: str = ""
    sl_pct: Optional[float] = None

    @property
    def sort_time(self):
        t = self.placed or self.done
        return t.astimezone(timezone.utc) if t else datetime.max.replace(tzinfo=timezone.utc)

    @property
    def offset(self):
        if self.trigger is not None and self.lmt is not None:
            return clean_num(abs(self.lmt - self.trigger))
        return None


def trigger_of(info):
    if (info.get("order_type") or "").upper().startswith("TRAIL"):
        return info.get("trail") or info.get("aux")
    return info.get("aux")


def order_value(data, info, price):
    """Approximate order value in base currency for orders that have no fills."""
    qty = info.get("qty") or 0.0
    symbol = info.get("symbol", "")
    if info.get("sec_type") == "CASH" and symbol.split(".")[0] == data.base_ccy:
        return qty
    if price is None:
        return None
    return qty * price * (info.get("multiplier") or 1.0) * data.fx_for(info.get("currency"))


def executed_orders(fills, cache):
    by_order = {}
    for f in fills:
        by_order.setdefault((f.account, f.perm_id or f.txn_id), []).append(f)
    orders = []
    for (acct, pid), fs in by_order.items():
        f0 = fs[0]
        info = cache.get(pid, {})
        orders.append(Order(
            account=acct, perm_id=pid, conid=f0.conid, symbol=f0.symbol, name=f0.name, side=f0.side,
            order_type=f0.order_type or ("AUTO-FX" if f0.auto_fx else ""),
            qty=sum(f.qty for f in fs), value=sum(f.value for f in fs),
            placed=f0.order_dt, done=fs[-1].dt, order_ref=f0.order_ref, fills=fs,
            status="executed (IBKR auto FX conversion)" if f0.auto_fx else "executed",
            trigger=trigger_of(info), lmt=info.get("lmt"), numbered=not f0.auto_fx,
        ))
    return orders


def tws_only_orders(data, cache, d_from, d_to):
    """Orders TWS knows but Flex does not: cancelled ones, and fills newer than the Flex data."""
    flex_perm_ids = {f.perm_id for f in data.fills}
    orders = []
    for pid, info in cache.items():
        if pid in flex_perm_ids:
            continue
        status = info.get("status", "")
        done = parse_tws_time(info.get("completed", ""), data.tz)
        if done is None:
            continue
        day = done.astimezone(data.tz).date()
        if status in CANCELLED_STATUSES:
            label = "cancelled"
        elif status == "Filled" and (data.to_date is None or day > data.to_date):
            label = "executed (TWS, not yet in Flex)"
        else:
            continue
        if not d_from <= day <= d_to:
            continue
        trig, lmt = trigger_of(info), info.get("lmt")
        orders.append(Order(
            account=info.get("account", ""), perm_id=pid, conid=info.get("conid", ""),
            symbol=info.get("symbol", ""), name=data.name_for(info.get("conid", ""), info.get("symbol", "")),
            side=info.get("action", ""), order_type=info.get("order_type", ""), qty=info.get("qty") or 0.0,
            value=order_value(data, info, lmt or trig), placed=None, done=done, status=label,
            order_ref=info.get("order_ref", ""), trigger=trig, lmt=lmt, numbered=label != "cancelled",
        ))
    return orders


def cycle_groups(data):
    """Group key of each position cycle = bracket tag of the order that opened it."""
    groups = {}
    for f in data.fills:
        if f.cycle is None:
            continue
        key = (f.account, f.conid, f.cycle)
        if key not in groups:
            ref, _ = ref_key(f.account, f.conid, f.order_ref)
            groups[key] = ref or ("cycle",) + key
    return groups


def assign_groups(orders, cyc_groups):
    for o in orders:
        key, o.is_entry = ref_key(o.account, o.conid, o.order_ref)
        if key is None and o.fills and o.fills[0].cycle is not None:
            # no bracket tag (manual / ActivityMonitor order): joins the position cycle it trades in
            key = cyc_groups.get((o.account, o.conid, o.fills[0].cycle))
        o.group = key or ("order", o.account, o.perm_id)


def number_groups(orders):
    """Sr No: each bracket / position cycle gets N, its entry order is N, follow-ups N.1, N.2 ...
    Cancelled orders are listed inside their group without a number. Returns orders in display order."""
    groups = defaultdict(list)
    for o in orders:
        groups[o.group].append(o)
    ordered = sorted(groups.values(), key=lambda g: min(o.sort_time for o in g))
    result = []
    n = 0
    for g in ordered:
        g.sort(key=lambda o: o.sort_time)
        live = [o for o in g if o.numbered]
        parent = next((o for o in live if o.is_entry), live[0] if live else None)
        if parent is not None:
            n += 1
            parent.sr = str(n)
            for k, o in enumerate((o for o in live if o is not parent), 1):
                o.sr = f"{n}.{k}"
            entry_price = parent.trigger if parent.trigger is not None else parent.lmt
            stops = [o for o in g if o is not parent and o.trigger is not None
                     and o.order_type.replace(" ", "").upper().startswith(("STP", "TRAIL"))]
            if entry_price and stops:
                parent.sl_pct = abs(entry_price - stops[0].trigger) / entry_price
            result.append(parent)
        result.extend(o for o in g if o is not parent)
    return result


# --------------------------------------------------------------------------- #
# Positions and per-contract statistics
# --------------------------------------------------------------------------- #
def contract_stats(fills):
    buys = [f for f in fills if f.side == "BUY"]
    sells = [f for f in fills if f.side == "SELL"]
    bq = sum(f.qty for f in buys)
    sq = sum(f.qty for f in sells)
    exchanges = []
    for f in fills:
        if f.exchange and f.exchange not in exchanges:
            exchanges.append(f.exchange)
    return {
        "first": fills[0].dt, "last": fills[-1].dt, "symbol": fills[0].symbol, "name": fills[0].name,
        "buys": bq, "sells": sq, "net": bq - sq,
        "avg_buy": sum(f.price * f.qty for f in buys) / bq if bq else 0.0,
        "avg_sell": sum(f.price * f.qty for f in sells) / sq if sq else 0.0,
        "tot_buy": sum(f.value for f in buys), "tot_sell": sum(f.value for f in sells),
        "exchanges": ", ".join(exchanges),
        "commission": sum(f.commission for f in fills), "pnl": sum(f.pnl for f in fills),
    }


def stats_row(s, unreal):
    return [to_utc(s["first"]), to_gst(s["first"]), to_utc(s["last"]), to_gst(s["last"]),
            s["symbol"], s["name"], clean_num(s["buys"]), clean_num(s["sells"]), clean_num(s["net"]),
            s["avg_buy"], s["avg_sell"], s["tot_buy"], s["tot_sell"], s["exchanges"],
            s["commission"], s["pnl"], unreal]


def open_positions(data, tws, day, live):
    """[(stats, unrealized)] for contracts with a position at the close of `day`."""
    by_contract = defaultdict(list)
    for f in data.fills:
        if not f.auto_fx and f.trade_date <= day:
            by_contract[(f.account, f.conid)].append(f)
    portfolio = tws.portfolio if (tws is not None and live) else {}
    rows = []
    for key, fills in by_contract.items():
        net = sum(f.signed for f in fills)
        if abs(net) < 1e-9:
            continue
        cycle = [f for f in fills if f.cycle == fills[-1].cycle]
        unreal = None
        if key in portfolio and portfolio[key]["unrealized"] is not None:
            unreal = portfolio[key]["unrealized"] * data.fx_for(portfolio[key]["currency"])
        elif key in data.flex_unrealized and data.flex_unrealized[key][0] == day:
            unreal = data.flex_unrealized[key][1]
        rows.append((contract_stats(cycle), unreal, key))
    # positions TWS holds that Flex has not reported yet (e.g. opened today)
    known = {k for _, _, k in rows}
    for key, p in portfolio.items():
        if key in known or abs(p["position"]) < 1e-9:
            continue
        s = {"first": None, "last": None, "symbol": p["symbol"], "name": data.name_for(key[1], p["symbol"]),
             "buys": None, "sells": None, "net": p["position"], "avg_buy": p["avg_cost"], "avg_sell": None,
             "tot_buy": None, "tot_sell": None, "exchanges": "", "commission": None, "pnl": None}
        unreal = p["unrealized"] * data.fx_for(p["currency"]) if p["unrealized"] is not None else None
        rows.append((s, unreal, key))
    rows.sort(key=lambda r: r[0]["last"].astimezone(timezone.utc) if r[0]["last"] else
              datetime.max.replace(tzinfo=timezone.utc), reverse=True)
    return rows


# --------------------------------------------------------------------------- #
# Excel writer
# --------------------------------------------------------------------------- #
def pnl_fill(cell):
    if isinstance(cell.value, (int, float)):
        if cell.value < 0:
            cell.fill = FILL_RED
        elif cell.value > 0:
            cell.fill = FILL_GREEN


def table_sheet(wb, title, banner, headers, rows, widths, formats=None, pnl_cols=()):
    ws = wb.create_sheet(title)
    ncol = len(headers)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = ws.cell(1, 1, banner)
    c.font, c.fill, c.alignment = FONT_TITLE, FILL_NAVY, ALIGN_CENTER
    for col in range(1, ncol + 1):
        ws.cell(1, col).border = BORDER_TITLE
        ws.cell(1, col).fill = FILL_NAVY
    for col, h in enumerate(headers, 1):
        c = ws.cell(2, col, h)
        c.font, c.fill, c.alignment, c.border = FONT_HEAD, FILL_NAVY, ALIGN_CENTER, BORDER_DATA
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 30
    formats = formats or {}
    for r, row in enumerate(rows, 3):
        for col, value in enumerate(row, 1):
            c = ws.cell(r, col, value)
            c.font, c.alignment, c.border = FONT_DATA, ALIGN_CENTER, BORDER_DATA
            if isinstance(value, datetime):
                c.number_format = FMT_DT
            elif col in formats and isinstance(value, (int, float)):
                c.number_format = formats[col]
            if col in pnl_cols:
                pnl_fill(c)
    for col in range(1, ncol + 1):
        ws.column_dimensions[get_column_letter(col)].width = widths.get(col, 13)
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(ncol)}{max(2, len(rows) + 2)}"
    return ws


def dashboard_sheet(wb, rows):
    ws = wb.active
    ws.title = "Dashboard"
    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value = "ACCOUNT DASHBOARD"
    c.font, c.fill, c.alignment = FONT_TITLE, FILL_NAVY, ALIGN_CENTER
    for ref in ("A1", "B1"):
        ws[ref].border, ws[ref].fill = BORDER_TITLE, FILL_NAVY
    for ref, h in (("A2", "Parameter"), ("B2", "Value")):
        ws[ref] = h
        ws[ref].font, ws[ref].fill, ws[ref].alignment = FONT_HEAD, FILL_NAVY, ALIGN_CENTER
    ws.row_dimensions[1].height = 24
    r = 3
    for row in rows:
        if row is None:
            r += 1
            continue
        kind = row[0]
        if kind == "section":
            ws.cell(r, 1, row[1])
            for col in (1, 2):
                ws.cell(r, col).font, ws.cell(r, col).fill = FONT_HEAD, FILL_NAVY
            ws.cell(r, 1).alignment = ALIGN_LEFT
        elif kind == "note":
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            ws.cell(r, 1, row[1]).font = FONT_NOTE
            ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        else:
            label, value, fmt, *flags = row
            ws.cell(r, 1, label).font = FONT_LABEL
            b = ws.cell(r, 2, value)
            b.font, b.alignment = FONT_DATA, ALIGN_RIGHT
            if fmt and isinstance(value, (int, float)):
                b.number_format = fmt
            elif isinstance(value, datetime):
                b.number_format = FMT_DT
            if "pnl" in flags:
                pnl_fill(b)
        r += 1
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 30
    ws.freeze_panes = "A3"


PENDING_HEADERS = ["Sr No", "Contract", "Name", "Action", "Type", "Trigger", "Limit", "Offset",
                   "SL-Percentage", "Quantity", "Total Amount"]
STATUS_HEADERS = ["Sr No", "Order Placed Date & Time (GST)", "Order Executed Date & Time (GST)",
                  "Contract", "Name", "Action", "Type", "Trigger", "Limit", "Offset", "SL-Percentage",
                  "Quantity", "Total Amount", "Order_Status"]
STAT_HEADERS = ["First Trade Date (UTC)", "First Trade Date (GST)", "Last Trade Date (UTC)",
                "Last Trade Date (GST)", "Contract", "Name", "Buys", "Sells", "Net", "Avg (bought)",
                "Avg (sold)", "Total (bought)", "Total (sold)", "Exchange List", "Commission", "PnL",
                "Unrealized PnL"]
TRADE_HEADERS = ["Date & Time (UTC)", "Date & Time (GST)", "Sr No", "Contract", "Name", "Action",
                 "Quantity", "Price", "Order Type", "Trigger Price", "Offset", "Limit Price", "SL %",
                 "Exchange", "Commission", "Realized PnL", "Account"]
STAT_WIDTHS = {1: 25, 2: 25, 3: 24, 4: 24, 5: 13, 6: 17, 7: 13, 10: 15, 11: 13, 12: 17,
               13: 15, 14: 16, 15: 13, 17: 17}
STAT_FORMATS = {10: FMT_AVG, 11: FMT_AVG, 12: FMT_MONEY, 13: FMT_MONEY, 15: FMT_MONEY, 16: FMT_MONEY,
                17: FMT_MONEY}


def pending_orders(data, tws):
    if tws is None:
        return []
    orders = []
    for pid, info in tws.open_orders.items():
        if info.get("status") not in OPEN_STATUSES:
            continue
        trig, lmt = trigger_of(info), info.get("lmt")
        o = Order(account=info["account"], perm_id=pid, conid=info["conid"], symbol=info["symbol"],
                  name=data.name_for(info["conid"], info["symbol"]), side=info["action"],
                  order_type=info["order_type"], qty=info["qty"], value=order_value(data, info, lmt or trig),
                  placed=None, done=None, status=info["status"], order_ref=info["order_ref"],
                  trigger=trig, lmt=lmt)
        key, o.is_entry = ref_key(o.account, o.conid, o.order_ref)
        if key is None and info.get("order_id"):
            parent = info.get("parent_id") or info.get("order_id")
            key = ("tws", o.account, parent)
            o.is_entry = not info.get("parent_id")
        o.group = key or ("order", o.account, pid)
        orders.append(o)
    # no placement time from TWS: keep permId order (sorting in number_groups is stable)
    orders.sort(key=lambda o: int(o.perm_id) if o.perm_id.isdigit() else 0)
    return number_groups(orders)


def build_report(data, cache, tws, d_from, d_to, out_path, include_live):
    fills = [f for f in data.fills if d_from <= f.trade_date <= d_to]
    live = include_live and tws is not None

    orders = executed_orders(fills, cache) + tws_only_orders(data, cache, d_from, d_to)
    assign_groups(orders, cycle_groups(data))
    orders = number_groups(orders)
    order_of = {(o.account, o.perm_id): o for o in orders}

    positions = open_positions(data, tws, d_to, live)
    unreal_known = [u for _, u, _ in positions if u is not None]
    unreal_total = sum(unreal_known) if unreal_known else None
    open_keys = {k: u for _, u, k in positions}

    wb = Workbook()

    # Dashboard
    buys = [f for f in fills if f.side == "BUY"]
    sells = [f for f in fills if f.side == "SELL"]
    realized = sum(f.pnl for f in fills)
    charges = [c for c in data.cfd_charges.values() if c[1] and d_from <= c[1] <= d_to]
    charge_total = sum(c[2] for c in charges)
    closing = [o for o in orders if o.numbered and o.fills and abs(sum(f.pnl for f in o.fills)) > 1e-9]
    winners = [o for o in closing if sum(f.pnl for f in o.fills) > 0]
    names = sorted({n for n in data.account_names.values() if n})
    rows = [
        ("section", "REPORT INFO"),
        ("Account ID", ", ".join(data.accounts), None),
        ("Account Name", ", ".join(names), None),
        ("Report Period", f"{d_from:%d-%b-%Y}  to  {d_to:%d-%b-%Y}", None),
        ("IBKR Flex Data Through", f"{data.to_date:%d-%b-%Y}" if data.to_date else "", None),
        ("Generated (GST)", datetime.now(TZ_GST).replace(tzinfo=None, microsecond=0), None),
        ("TWS Live Data", "connected" if live else "not used", None),
        None,
        ("section", f"ALL TRADES SUMMARY ({data.base_ccy})"),
        ("Total Trades (fills)", len(fills), "0"),
        ("Orders Executed", len([o for o in orders if o.fills]), "0"),
        ("Contracts Traded", len({(f.account, f.conid) for f in fills}), "0"),
        ("Total Bought Value", sum(f.value for f in buys), FMT_MONEY),
        ("Total Sold Value", sum(f.value for f in sells), FMT_MONEY),
        ("Commission", sum(f.commission for f in fills), FMT_MONEY, "pnl"),
        ("Realized PnL", realized, FMT_MONEY, "pnl"),
        ("CFD Financing Charges", charge_total, FMT_MONEY, "pnl"),
        ("Net Realized PnL (after charges)", realized + charge_total, FMT_MONEY, "pnl"),
        ("Unrealized PnL", unreal_total, FMT_MONEY, "pnl"),
        ("Open Positions", len(positions), "0"),
        ("Winning / Losing Closes", f"{len(winners)} / {len(closing) - len(winners)}", None),
        ("Max Exposure", setting_num("MAX_EXPOSURE", 150000), FMT_MONEY),
        ("Daily Max Loss", setting_num("DAILY_MAX_LOSS", 2000), FMT_MONEY),
        ("Loss due to bugs", setting_num("LOSS_DUE_TO_BUGS", 0), FMT_MONEY),
    ]
    if len(data.accounts) > 1:
        rows += [None, ("section", "BY ACCOUNT")]
        for acct in data.accounts:
            af = [f for f in fills if f.account == acct]
            rows.append((f"{acct}  Commission", sum(f.commission for f in af), FMT_MONEY, "pnl"))
            rows.append((f"{acct}  Realized PnL", sum(f.pnl for f in af), FMT_MONEY, "pnl"))
    rows += [None, ("note", "Commission and Realized PnL are IBKR's own figures per fill (FIFO; realized "
                            "PnL is already net of commissions). Negative commission = paid.")]
    dashboard_sheet(wb, rows)

    # Pending Order (live from TWS)
    pend_rows = []
    for o in (pending_orders(data, tws) if live else []):
        pend_rows.append([o.sr, o.symbol, o.name, o.side, o.order_type, o.trigger, o.lmt, o.offset,
                          o.sl_pct, clean_num(o.qty), o.value])
    table_sheet(wb, "Pending Order", "PENDING ORDERS", PENDING_HEADERS, pend_rows,
                {1: 7, 2: 12, 3: 30, 4: 9, 5: 10, 9: 15, 10: 10, 11: 14},
                formats={6: FMT_PRICE, 7: FMT_PRICE, 8: FMT_PRICE, 9: FMT_PCT, 11: FMT_MONEY})

    # Trades_Status -- one row per order
    status_rows = [[o.sr, to_gst(o.placed), to_gst(o.done), o.symbol, o.name, o.side, o.order_type,
                    o.trigger, o.lmt, o.offset, o.sl_pct, clean_num(o.qty), o.value, o.status]
                   for o in orders]
    table_sheet(wb, "Trades_Status", "TRADES STATUS", STATUS_HEADERS, status_rows,
                {1: 7, 2: 26, 3: 26, 4: 12, 5: 30, 6: 9, 7: 10, 11: 15, 12: 10, 13: 14, 14: 20},
                formats={8: FMT_PRICE, 9: FMT_PRICE, 10: FMT_PRICE, 11: FMT_PCT, 13: FMT_MONEY})

    # Open Position
    open_rows = [stats_row(s, u) for s, u, _ in positions]
    table_sheet(wb, "Open Position", "OPEN POSITIONS", STAT_HEADERS, open_rows, STAT_WIDTHS,
                formats=STAT_FORMATS, pnl_cols=(16, 17))

    # All Trades -- one row per fill, newest first
    trade_rows = []
    for f in sorted(fills, key=lambda f: (f.dt, f.txn_id), reverse=True):
        o = order_of.get((f.account, f.perm_id or f.txn_id))
        trade_rows.append([
            to_utc(f.dt), to_gst(f.dt), o.sr if o else "", f.symbol, f.name, f.side, clean_num(f.qty),
            clean_num(f.price), o.order_type if o else f.order_type, o.trigger if o else None,
            o.offset if o else None, o.lmt if o else None, o.sl_pct if o else None,
            f.exchange, f.commission, f.pnl, f.account,
        ])
    table_sheet(wb, "All Trades", "ALL TRADES", TRADE_HEADERS, trade_rows,
                {1: 22, 2: 22, 3: 7, 4: 14, 5: 30, 6: 10, 7: 12, 8: 14, 9: 12, 10: 14, 11: 12,
                 12: 14, 13: 10, 14: 12, 15: 14, 16: 13, 17: 13},
                formats={8: FMT_PRICE, 10: FMT_PRICE, 11: FMT_PRICE, 12: FMT_PRICE, 13: FMT_PCT,
                         15: FMT_MONEY, 16: FMT_MONEY}, pnl_cols=(16,))

    # Trade Summary -- one row per contract traded in the period
    by_contract = {}
    for f in fills:
        if not f.auto_fx:
            by_contract.setdefault((f.account, f.conid), []).append(f)
    summary_rows = [stats_row(contract_stats(fs), open_keys.get(key, 0.0))
                    for key, fs in sorted(by_contract.items(), key=lambda kv: kv[1][0].dt)]
    table_sheet(wb, "Trade Summary", "TRADE SUMMARY", STAT_HEADERS, summary_rows, STAT_WIDTHS,
                formats=STAT_FORMATS, pnl_cols=(16, 17))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return {"fills": len(fills), "orders": len(orders), "open": len(positions),
            "commission": sum(f.commission for f in fills), "pnl": realized}


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def parse_day(text):
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        raise argparse.ArgumentTypeError(f"date must be YYYY-MM-DD, got {text!r}")


def main():
    load_env()
    ap = argparse.ArgumentParser(description="GT IBKR full-history report (Flex + TWS)")
    ap.add_argument("--from", dest="d_from", type=parse_day, help="first trade date YYYY-MM-DD (default: all)")
    ap.add_argument("--to", dest="d_to", type=parse_day, help="last trade date YYYY-MM-DD (default: latest)")
    ap.add_argument("--daily", action="store_true", help="also write one workbook per trading day")
    ap.add_argument("--xml", nargs="+", help="use saved Flex XML file(s) instead of downloading")
    ap.add_argument("--no-tws", action="store_true", help="do not connect to TWS / IB Gateway")
    ap.add_argument("--out", help=f"output folder (default: {DEFAULT_OUT})")
    args = ap.parse_args()

    out_dir = Path(args.out).expanduser().resolve() if args.out else DEFAULT_OUT
    tz = ZoneInfo(setting("FLEX_TIMEZONE", "America/New_York"))

    if args.xml:
        texts = [Path(p).read_text(encoding="utf-8") for p in args.xml]
    else:
        token, query_id = setting("GT_FLEX_TOKEN"), setting("GT_FLEX_QUERY_ID")
        if not token or not query_id:
            sys.exit("Missing GT_FLEX_TOKEN / GT_FLEX_QUERY_ID -- set them in NEW_REPORTS/.env "
                     "(copy .env.example) or as environment variables.")
        try:
            texts = download_all(token, query_id, args.d_from)
        except FlexError as e:
            sys.exit(f"IBKR Flex request failed {e}")
        except requests.RequestException as e:
            sys.exit(f"Could not reach IBKR Flex Web Service: {e}")
        raw_dir = out_dir / "raw"
        raw_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        for i, text in enumerate(texts):
            (raw_dir / f"flex_{stamp}_{i}.xml").write_text(text, encoding="utf-8")
        print(f"  Raw Flex XML saved in {raw_dir}")

    data = FlexData(texts, tz)
    print(f"Accounts {', '.join(data.accounts)} | base {data.base_ccy} | {len(data.fills)} fills | "
          f"Flex data {data.from_date} -> {data.to_date}")
    if not data.fills:
        print("  !! No executions in the Flex statement -- check the query includes the Trades section "
              "with 'Execution' level of detail.")

    use_tws = not args.no_tws and setting("USE_TWS", "1").lower() not in ("0", "false", "no")
    tws = None
    if use_tws:
        tws = fetch_tws(setting("TWS_HOST", "127.0.0.1"), int(setting_num("TWS_PORT", 4001)),
                        int(setting_num("TWS_CLIENT_ID", 997)), setting_num("TWS_WAIT_SECONDS", 8))
    cache = update_order_cache(tws, ORDER_CACHE)

    first_fill = data.fills[0].trade_date if data.fills else data.from_date or date.today()
    last_day = max([d for d in (data.to_date, data.fills[-1].trade_date if data.fills else None) if d],
                   default=date.today())
    d_from = args.d_from or first_fill
    d_to = args.d_to or last_day
    if d_from > d_to:
        sys.exit(f"--from {d_from} is after --to {d_to}")
    live = args.d_to is None or d_to >= last_day

    out = out_dir / f"GT_Report_{d_from:%Y-%m-%d}_to_{d_to:%Y-%m-%d}.xlsx"
    try:
        s = build_report(data, cache, tws, d_from, d_to, out, include_live=live)
        print(f"{d_from:%d-%b-%Y} -> {d_to:%d-%b-%Y}: {s['fills']} fills, {s['orders']} orders, "
              f"{s['open']} open positions, commission {s['commission']:,.2f}, realized PnL {s['pnl']:,.2f}")
        print(f"  -> {out}")
        if args.daily:
            days = sorted({f.trade_date for f in data.fills if d_from <= f.trade_date <= d_to})
            for day in days:
                path = out_dir / "daily" / f"GT_Report_{day:%Y-%m-%d}.xlsx"
                s = build_report(data, cache, tws, day, day, path, include_live=False)
                print(f"  {day:%a %d-%b-%Y}: {s['fills']:3d} fills, realized PnL {s['pnl']:10,.2f} -> {path.name}")
    except PermissionError as e:
        sys.exit(f"  !! Cannot write {Path(e.filename).name if e.filename else 'the report'} "
                 f"-- close it in Excel and run again.")


if __name__ == "__main__":
    main()
