#!/usr/bin/env python3

"""
Buffett Indicator + Shiller CAPE + Berkshire cash dashboard (v5)
==========================================================

What the script does
--------------------
1. Downloads the latest public data from:
   - Yahoo Finance (Wilshire 5000 and S&P 500)
   - FRED (US nominal GDP)
   - Robert Shiller's data workbook (S&P and CAPE history)
   - CompaniesMarketCap pages (Berkshire Hathaway cash on hand and total assets)
2. Builds an interactive Plotly dashboard with shared time axis.
3. Saves an Excel workbook with raw/merged historical data, plus dedicated S&P
   and Berkshire audit sheets.

Outputs
-------
- buffett_dashboard.html : interactive multi-panel chart
- buffett_dashboard.xlsx : raw and merged workbook

Usage
-----
python buffett_dashboard.py --user-agent "Your Name your.email@company.com"

Notes
-----
- The legend acts as the requested on/off switch. Clicking a legend item toggles
  the corresponding output series. For grouped items such as Buffett bands and
  CAPE bands, the full group toggles together.
- The Buffett trend and +/- standard deviation bands are computed on a log scale
  so the reference lines remain proportional over long horizons, similar to the
  reference screenshots.
- Berkshire cash/assets are sourced only from CompaniesMarketCap in v5 because
  the narrower SEC cash taxonomy used in prior versions could materially
  understate Berkshire's broader liquidity.
- v5 removes all Berkshire SEC download code so the script no longer mixes two
  incompatible Berkshire liquidity definitions in the same chart/workbook.
- Progress messages remain visible during execution so long-running downloads are
  easier to understand.
"""


from __future__ import annotations

import argparse
import io
import math
import os
import re
import time
from urllib.parse import urljoin
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import yfinance as yf
from plotly.subplots import make_subplots


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
FRED_GDP_CSV = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=GDP"
FRED_WILSHIRE_CSV = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=WILL5000IND"
SHILLER_URLS = [
    "https://www.econ.yale.edu/~shiller/data/ie_data.xls",
    "http://www.econ.yale.edu/~shiller/data/ie_data.xls",
]
SHILLER_DISCOVERY_PAGE = "https://shillerdata.com/"
CMC_BERKSHIRE_TOTAL_ASSETS_URL = "https://companiesmarketcap.com/berkshire-hathaway/total-assets/"
CMC_BERKSHIRE_CASH_ON_HAND_URL = "https://companiesmarketcap.com/berkshire-hathaway/cash-on-hand/"
DEFAULT_START = "1989-01-01"
MIN_DISPLAY_START = "1989-01-01"
DEFAULT_OUTPUT_HTML = "buffett_dashboard.html"
DEFAULT_OUTPUT_XLSX = "buffett_dashboard.xlsx"
STDDEV_MULTIPLIERS = (0.5, 1.0, 1.5, 2.0)


# ---------------------------------------------------------------------------
# Data containers
# ---------------------------------------------------------------------------
@dataclass
class SeriesBundle:
    name: str
    frame: pd.DataFrame



@dataclass
class BerkshireHistoryBundle:
    """Container for Berkshire output plus raw audit sheets.

    v5 intentionally removes every Berkshire SEC frame. The dashboard now uses
    CompaniesMarketCap as the sole Berkshire source for both cash-on-hand and
    total-assets so the workbook audit trail stays aligned with the visualized
    series and with the broader liquidity definition requested by the user.
    """

    merged: pd.DataFrame
    cmc_cash: pd.DataFrame
    cmc_assets: pd.DataFrame
    source_audit: pd.DataFrame


# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------
def make_http_session(user_agent: str) -> requests.Session:
    """Create a requests session with polite headers.

    v5 no longer downloads Berkshire facts from SEC EDGAR, but keeping a
    descriptive user-agent is still a good practice for public data sources.
    The same session is reused across all downloads to centralize timeout and
    header behavior in one place.
    """
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": user_agent,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "en-US,en;q=0.9",
        }
    )
    return session


def ensure_output_dir(path: Path) -> Path:
    """Create output directory if needed and return the normalized path."""
    path.mkdir(parents=True, exist_ok=True)
    return path.resolve()


def clean_numeric_series(series: pd.Series) -> pd.Series:
    """Convert strings to numeric while safely coercing non-numeric values."""
    return pd.to_numeric(series, errors="coerce")


def _format_stddev_suffix(multiplier: float) -> str:
    """Return a stable column-name suffix for a standard-deviation multiplier.

    Examples
    --------
    0.5 -> ``0_5``
    1.0 -> ``1``
    1.5 -> ``1_5``

    Using underscores instead of decimal points keeps the exported Excel column
    names easy to reference from formulas and downstream automation.
    """
    if float(multiplier).is_integer():
        return str(int(multiplier))
    return str(multiplier).replace(".", "_")


def add_stddev_level_columns(
    df: pd.DataFrame,
    source_col: str,
    prefix: str,
    *,
    include_mean: bool = True,
    floor_at_zero: bool = False,
) -> pd.DataFrame:
    """Append mean and +/- standard-deviation level columns for one series.

    The user requested these extra Excel columns for the raw exported sheets.
    This helper computes a single historical mean/std pair from ``source_col``
    and writes constant reference-level columns for +/-0.5, +/-1.0, +/-1.5, and
    +/-2.0 standard deviations.

    Parameters
    ----------
    df:
        Source dataframe to augment.
    source_col:
        Name of the numeric column whose historical distribution is used.
    prefix:
        Prefix used for the generated output column names.
    include_mean:
        When True, also writes ``{prefix}_mean``.
    floor_at_zero:
        When True, negative lower-band values are clipped to zero which is useful
        for non-negative series such as CAPE, index levels, and market-cap/GDP %.
    """
    out = df.copy()
    series = clean_numeric_series(out[source_col]).dropna()

    if series.empty:
        mean_value = np.nan
        std_value = np.nan
    else:
        mean_value = float(series.mean())
        std_value = float(series.std(ddof=1)) if len(series) > 1 else 0.0

    if include_mean:
        out[f"{prefix}_mean"] = mean_value

    for multiplier in STDDEV_MULTIPLIERS:
        suffix = _format_stddev_suffix(multiplier)
        plus_value = mean_value + (multiplier * std_value)
        minus_value = mean_value - (multiplier * std_value)
        if floor_at_zero and pd.notna(minus_value):
            minus_value = max(0.0, minus_value)
        out[f"{prefix}_plus_{suffix}sd"] = plus_value
        out[f"{prefix}_minus_{suffix}sd"] = minus_value

    return out


def normalize_ticker(text: str) -> str:
    """Normalize tickers so BRK.B, BRK-B, and BRKB can be compared easily."""
    return re.sub(r"[^A-Z0-9]", "", text.upper())




def parse_human_number(value: object) -> float | np.nan:
    """Convert strings like '$123.4B' or '987,654' into numeric floats.

    This helper is used by the CompaniesMarketCap parser, where values
    may appear either in HTML tables or inside inline chart configuration text.
    The parser intentionally accepts magnitude suffixes (K/M/B/T) because public
    finance sites often mix plain numbers with abbreviated display strings.
    """
    if pd.isna(value):
        return np.nan
    text = str(value).strip()
    if not text or text in {"-", "—", "N/A", "nan"}:
        return np.nan
    text = text.replace(",", "").replace("$", "").replace("€", "").replace("£", "")
    multiplier = 1.0
    match = re.search(r"([KMBT])$", text, flags=re.IGNORECASE)
    if match:
        multiplier = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12}[match.group(1).upper()]
        text = text[:-1]
    try:
        return float(text) * multiplier
    except ValueError:
        return np.nan


def normalize_statement_date(series: pd.Series) -> pd.Series:
    """Normalize arbitrary dates to month-end timestamps without time values."""
    out = pd.to_datetime(series, errors="coerce")
    return out.dt.to_period("M").dt.to_timestamp(how="end").dt.normalize()


def ensure_history_frame(df: pd.DataFrame, value_name: str) -> pd.DataFrame:
    """Standardize a two-column history frame and drop duplicate month-ends.

    This helper keeps the Berkshire history parsers small and predictable.
    Each returned row represents a single month-end date plus one numeric value.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=["date", value_name])
    work = df.copy()
    work["date"] = normalize_statement_date(work["date"])
    work[value_name] = clean_numeric_series(work[value_name])
    work = work.dropna(subset=["date", value_name])
    work = work.sort_values("date").drop_duplicates(subset=["date"], keep="last")
    return work[["date", value_name]].reset_index(drop=True)


def log_progress(message: str, started_at: float | None = None) -> None:
    """Print a user-friendly progress line that flushes immediately.

    Why this exists:
    - Some public data sources are slow or temporarily rate-limited.
    - Without any interim output, users may think the script is frozen.
    - Flushing on each message ensures the text appears immediately in terminals
      and notebook consoles instead of being buffered until the end.
    """
    timestamp = pd.Timestamp.now().strftime("%H:%M:%S")
    suffix = ""
    if started_at is not None:
        suffix = f" | elapsed {time.perf_counter() - started_at:,.1f}s"
    print(f"[{timestamp}] {message}{suffix}", flush=True)


def parse_cmc_js_history(html: str, value_name: str) -> pd.DataFrame:
    """Best-effort parser for CompaniesMarketCap history embedded inside HTML.

    The site may expose data either as a visible table or inside script-tag text
    that powers the browser chart. Instead of relying on one fragile selector,
    the parser tries a few safe date/value patterns and returns whatever valid
    history it can infer. If the site structure changes again, the parser raises an error or returns
    an empty frame, making the source issue visible instead of silently mixing
    in a different metric definition from another provider.
    """
    records: list[tuple[pd.Timestamp, float]] = []

    iso_patterns = [
        # Matches YYYY-MM-DD followed by any content (incl. HTML tags / newlines) up to
        # 150 chars, then a numeric value optionally followed by a magnitude suffix
        # (B/M/K/T). The suffix is inside the capture group so parse_human_number can
        # apply the correct multiplier. re.DOTALL is NOT used globally; [\s\S] handles
        # newlines between the date cell and value cell.
        re.compile(r'(?P<date>(?:19|20)\d{2}-\d{2}-\d{2})[\s\S]{0,150}?(?<!\d)(?P<value>-?\d[\d,]*(?:\.\d+)?(?:[ \t]*[BbMmKkTt])?)(?![a-zA-Z\d])'),
        re.compile(r'Date\.UTC\((?P<year>\d{4}),(?P<month>\d{1,2}),(?P<day>\d{1,2})\)\s*,\s*(?P<value>-?\d[\d,]*(?:\.\d+)?)'),
        re.compile(r"[\"'](?:date|d)[\"']\s*:\s*[\"'](?P<date>(?:19|20)\d{2}-\d{2}-\d{2})[\"'][^{}]{0,120}?[\"'](?:value|val|v|y)[\"']\s*:\s*(?P<value>-?\d[\d,]*(?:\.\d+)?)"),
    ]

    for match in iso_patterns[0].finditer(html):
        dt = pd.to_datetime(match.group("date"), errors="coerce")
        val = parse_human_number(match.group("value"))
        if pd.notna(dt) and pd.notna(val):
            records.append((dt.normalize(), float(val)))

    for match in iso_patterns[1].finditer(html):
        dt = pd.Timestamp(
            year=int(match.group("year")),
            month=int(match.group("month")) + 1,
            day=int(match.group("day")),
        ) + pd.offsets.MonthEnd(0)
        val = parse_human_number(match.group("value"))
        if pd.notna(val):
            records.append((dt.normalize(), float(val)))

    for match in iso_patterns[2].finditer(html):
        dt = pd.to_datetime(match.group("date"), errors="coerce")
        val = parse_human_number(match.group("value"))
        if pd.notna(dt) and pd.notna(val):
            records.append((dt.normalize(), float(val)))

    if not records:
        return pd.DataFrame(columns=["date", value_name])

    out = pd.DataFrame(records, columns=["date", value_name])
    out = ensure_history_frame(out, value_name)
    return out[out[value_name] > 0].reset_index(drop=True)


def fetch_companiesmarketcap_history(session: requests.Session, url: str, value_name: str) -> pd.DataFrame:
    """Fetch historical Berkshire metrics from CompaniesMarketCap.

    v5 treats CompaniesMarketCap as the only Berkshire source. The parser first
    tries visible HTML tables and then falls back to best-effort script parsing.
    """
    response = session.get(url, timeout=60)
    response.raise_for_status()
    html = response.text

    try:
        tables = pd.read_html(io.StringIO(html))
    except ValueError:
        tables = []

    for table in tables:
        work = table.copy()
        work.columns = [str(col).strip() for col in work.columns]
        lower = {str(col).strip().lower(): col for col in work.columns}
        date_col = next((orig for low, orig in lower.items() if "date" in low), None)
        value_col = next((orig for low, orig in lower.items() if any(tok in low for tok in ["assets", "cash", "amount", "value"]) and orig != date_col), None)

        # Fallback: headers are unnamed (e.g. empty <th> cells) — detect columns by content.
        if date_col is None:
            for col in work.columns:
                sample = work[col].dropna().astype(str).head(10)
                if any(re.match(r'(?:19|20)\d{2}-\d{2}-\d{2}$', v.strip()) for v in sample):
                    date_col = col
                    break
        if date_col is not None and value_col is None:
            for col in work.columns:
                if col == date_col:
                    continue
                sample = work[col].dropna().astype(str).head(5)
                if any(re.search(r'\d', v) for v in sample):
                    value_col = col
                    break

        if date_col and value_col:
            parsed = pd.DataFrame({"date": work[date_col], value_name: work[value_col].map(parse_human_number)})
            parsed = ensure_history_frame(parsed, value_name)
            if not parsed.empty:
                return parsed

    return parse_cmc_js_history(html, value_name)


def year_month_decimal_to_timestamp(value: object) -> pd.Timestamp | pd.NaT:
    """Parse Robert Shiller's decimal date format into a month-end timestamp.

    The Yale workbook typically stores dates like 1871.01, 1871.02, ..., 2026.04.
    The digits after the decimal point represent the month number, not a fraction
    of the year. This parser is intentionally defensive and returns NaT if the
    input is badly formed.
    """
    if pd.isna(value):
        return pd.NaT
    try:
        as_float = float(value)
    except (TypeError, ValueError):
        return pd.NaT

    year = int(as_float)
    month = int(round((as_float - year) * 100))
    if month < 1 or month > 12:
        return pd.NaT
    return pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)


# ---------------------------------------------------------------------------
# Download functions
# ---------------------------------------------------------------------------
def fetch_fred_csv(session: requests.Session, url: str, value_name: str) -> pd.DataFrame:
    """Download a simple two-column FRED CSV and return a standardized frame."""
    response = session.get(url, timeout=60)
    response.raise_for_status()
    df = pd.read_csv(io.StringIO(response.text))
    date_col = df.columns[0]
    value_col = df.columns[1]
    df = df.rename(columns={date_col: "date", value_col: value_name})
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df[value_name] = clean_numeric_series(df[value_name])
    df = df.dropna(subset=["date"]).sort_values("date")
    return df.reset_index(drop=True)


def fetch_yahoo_history(ticker: str, start: str, interval: str = "1mo") -> pd.DataFrame:
    """Fetch Yahoo Finance history via yfinance and standardize the output.

    The interval defaults to monthly because Buffett / CAPE comparison works
    best on a longer-run time axis and monthly frequency also reduces workbook size.
    """
    data = yf.download(
        tickers=ticker,
        start=start,
        interval=interval,
        auto_adjust=False,
        progress=False,
        threads=False,
    )
    if data.empty:
        raise ValueError(f"Yahoo Finance returned no data for {ticker!r}")

    if isinstance(data.columns, pd.MultiIndex):
        # yfinance can return a column MultiIndex even for a single ticker.
        data.columns = data.columns.get_level_values(0)

    close_col = "Adj Close" if "Adj Close" in data.columns else "Close"
    result = data[[close_col]].copy()
    result = result.rename(columns={close_col: "value"})
    result.index = pd.to_datetime(result.index)
    result.index = result.index.to_period("M").to_timestamp(how="end").normalize()
    result = result[~result.index.duplicated(keep="last")]
    result = result.reset_index().rename(columns={"Date": "date", "index": "date"})
    return result


def _discover_shiller_urls(session: requests.Session) -> list[str]:
    """Discover candidate Shiller workbook URLs.

    Why this helper exists:
    - Robert Shiller's historical workbook has moved over time.
    - Public Shiller wrappers explicitly scrape shillerdata.com to discover the
      current workbook URL, which is often hosted behind a changing wsimg URL.
    - The function returns a de-duplicated list with the original Yale URLs
      first, then any discovered shillerdata.com download URLs.
    """
    candidates: list[str] = list(SHILLER_URLS)
    try:
        response = session.get(SHILLER_DISCOVERY_PAGE, timeout=30)
        response.raise_for_status()
        html = response.text
        discovered: list[str] = []
        href_pattern = r'(?:href|src)=["\']([^"\']*ie_data[^"\']*\.xls[^"\']*)["\']'
        direct_pattern = r'https?://[^"\'\s>]+ie_data[^"\'\s>]*\.xls(?:\?[^"\'\s>]*)?'
        for match in re.findall(href_pattern, html, flags=re.IGNORECASE):
            discovered.append(urljoin(SHILLER_DISCOVERY_PAGE, match))
        for match in re.findall(direct_pattern, html, flags=re.IGNORECASE):
            discovered.append(match)
        for url in discovered:
            if url not in candidates:
                candidates.append(url)
    except Exception:
        # Discovery failure should not break the dashboard; the static Yale URLs
        # are still attempted first.
        pass
    return candidates


def _standardize_shiller_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize Shiller workbook columns into a consistent monthly DataFrame.

    Supported workbook layouts:
    - The older Yale workbook where the useful sheet starts at header row 7.
    - The shillerdata.com workbook variant documented by public parser projects,
      where columns are named 'P', 'E', 'D', 'CAPE', etc.
    """
    if df is None or df.empty:
        raise ValueError('Shiller workbook parsed into an empty DataFrame')

    work = df.copy()
    work.columns = [str(col).strip() for col in work.columns]
    work = work.dropna(how='all')
    if work.empty:
        raise ValueError('Shiller workbook contains no usable rows')

    lower_map = {str(c).strip().lower(): c for c in work.columns}

    def find_col(candidates: Iterable[str]) -> str:
        for candidate in candidates:
            candidate_lower = candidate.lower()
            for lower_name, original_name in lower_map.items():
                if lower_name == candidate_lower or candidate_lower in lower_name:
                    return original_name
        raise KeyError(f'Unable to find any of the expected columns: {list(candidates)}')

    date_col = find_col(['date'])
    sp_col = find_col(['s&p comp.', 'p', 'sp500'])
    cape_col = find_col(['cape', 'pe10'])

    out = pd.DataFrame(
        {
            'date': work[date_col].map(year_month_decimal_to_timestamp),
            'sp500_shiller': clean_numeric_series(work[sp_col]),
            'cape': clean_numeric_series(work[cape_col]),
        }
    )
    out = out.dropna(subset=['date']).sort_values('date')
    out = out[(out['cape'].notna()) & (out['date'] >= pd.Timestamp(DEFAULT_START))]
    if out.empty:
        raise ValueError('Shiller workbook did not produce any valid CAPE rows after normalization')
    return out.reset_index(drop=True)


def fetch_shiller_workbook(session: requests.Session) -> pd.DataFrame:
    """Download and parse Robert Shiller's latest workbook into monthly S&P/CAPE data.

    Why the original implementation failed:
    - The workbook structure varies by source and over time.
    - On some runs, sheet 0 or the default header strategy yields an empty frame,
      causing an out-of-bounds error when the code later assumes row 7 exists.

    This version is more robust because it:
    1. Tries multiple candidate URLs, including a scraped shillerdata.com link.
    2. Tries multiple sheet/header layouts that are known to exist in public
       Shiller data mirrors.
    3. Validates row counts before touching header_row.
    4. Falls back with a detailed error if all strategies fail.
    """
    errors: list[str] = []
    for url in _discover_shiller_urls(session):
        try:
            response = session.get(url, timeout=90)
            response.raise_for_status()
            content = response.content
        except Exception as exc:
            errors.append(f'{url} -> download failed: {exc}')
            continue

        for strategy in [
            {'sheet_name': 'Data', 'header': 7},
            {'sheet_name': 0, 'header': 7},
        ]:
            try:
                parsed = pd.read_excel(io.BytesIO(content), engine='xlrd', **strategy)
                return _standardize_shiller_frame(parsed)
            except Exception as exc:
                errors.append(f'{url} -> strategy {strategy} failed: {exc}')

        try:
            raw = pd.read_excel(io.BytesIO(content), sheet_name=0, header=None, engine='xlrd')
            if raw is None or raw.empty:
                raise ValueError('raw workbook is empty')

            header_row = None
            for idx in range(min(30, len(raw))):
                row_text = [str(x).strip().lower() for x in raw.iloc[idx].tolist()]
                has_date = any(cell == 'date' for cell in row_text)
                has_cape = any('cape' in cell or 'pe10' in cell for cell in row_text)
                if has_date and has_cape:
                    header_row = idx
                    break

            if header_row is None:
                raise ValueError('no header row with date + CAPE markers was found in first 30 rows')
            if header_row >= len(raw):
                raise ValueError(f'discovered header row {header_row} is outside the workbook bounds ({len(raw)} rows)')

            header_values = [str(x).strip() for x in raw.iloc[header_row].tolist()]
            body = raw.iloc[header_row + 1 :].copy()
            body.columns = header_values
            return _standardize_shiller_frame(body)
        except Exception as exc:
            errors.append(f'{url} -> raw-sheet strategy failed: {exc}')

    raise RuntimeError('Unable to download and parse Shiller workbook. Detailed attempts: ' + ' | '.join(errors))



def fetch_berkshire_history(session: requests.Session, ticker: str = "BRK-B") -> BerkshireHistoryBundle:
    """Download Berkshire cash/assets from CompaniesMarketCap only.

    Why v5 removes SEC here:
    - Berkshire's SEC cash series in prior versions only used narrow cash XBRL
      concepts such as CashAndCashEquivalentsAtCarryingValue.
    - That narrow SEC taxonomy can materially understate Berkshire's broader
      liquidity when short-dated Treasury bills or similar current investments
      are reported outside those exact cash tags.
    - The user explicitly requested one consistent data definition that matches
      the broader cash-on-hand metric shown on CompaniesMarketCap, so v5 now
      scrapes only the Berkshire CompaniesMarketCap history pages and removes
      every Berkshire SEC download path.

    Parameters
    ----------
    session:
        Shared HTTP session used for public downloads.
    ticker:
        Retained for interface compatibility with v4. v5 is Berkshire-specific
        and therefore only accepts the normalized BRK.B/BRK-B/BRKB ticker.
    """
    normalized = normalize_ticker(ticker)
    if normalized != "BRKB":
        raise ValueError(
            "fetch_berkshire_history() is Berkshire-specific in this script and only supports BRK.B / BRK-B."
        )

    try:
        cmc_cash = fetch_companiesmarketcap_history(session, CMC_BERKSHIRE_CASH_ON_HAND_URL, "brk_cash_usd")
    except Exception:
        cmc_cash = pd.DataFrame(columns=["date", "brk_cash_usd"])

    try:
        cmc_assets = fetch_companiesmarketcap_history(session, CMC_BERKSHIRE_TOTAL_ASSETS_URL, "brk_total_assets_usd")
    except Exception:
        cmc_assets = pd.DataFrame(columns=["date", "brk_total_assets_usd"])

    cmc_cash = ensure_history_frame(cmc_cash, "brk_cash_usd")
    cmc_assets = ensure_history_frame(cmc_assets, "brk_total_assets_usd")

    if cmc_cash.empty and cmc_assets.empty:
        raise RuntimeError(
            "CompaniesMarketCap returned no Berkshire cash-on-hand or total-assets history; the page structure may have changed."
        )

    merged = pd.merge(cmc_cash, cmc_assets, on="date", how="outer").sort_values("date").reset_index(drop=True)
    merged["brk_cash_source"] = np.where(merged["brk_cash_usd"].notna(), "CompaniesMarketCap", pd.NA)
    merged["brk_total_assets_source"] = np.where(
        merged["brk_total_assets_usd"].notna(),
        "CompaniesMarketCap",
        pd.NA,
    )
    merged["brk_cash_to_assets_pct"] = np.where(
        merged["brk_cash_usd"].notna() & merged["brk_total_assets_usd"].notna() & (merged["brk_total_assets_usd"] != 0),
        (merged["brk_cash_usd"] / merged["brk_total_assets_usd"]) * 100.0,
        np.nan,
    )
    merged["brk_data_status"] = np.select(
        [
            merged["brk_cash_usd"].notna() & merged["brk_total_assets_usd"].notna(),
            merged["brk_cash_usd"].notna() & merged["brk_total_assets_usd"].isna(),
            merged["brk_cash_usd"].isna() & merged["brk_total_assets_usd"].notna(),
        ],
        ["complete", "cash_only", "assets_only"],
        default="missing",
    )

    source_audit = merged[
        [
            "date",
            "brk_cash_usd",
            "brk_cash_source",
            "brk_total_assets_usd",
            "brk_total_assets_source",
            "brk_cash_to_assets_pct",
            "brk_data_status",
        ]
    ].copy()

    return BerkshireHistoryBundle(
        merged=merged[
            [
                "date",
                "brk_cash_usd",
                "brk_total_assets_usd",
                "brk_cash_to_assets_pct",
                "brk_cash_source",
                "brk_total_assets_source",
                "brk_data_status",
            ]
        ].copy(),
        cmc_cash=cmc_cash,
        cmc_assets=cmc_assets,
        source_audit=source_audit,
    )


# ---------------------------------------------------------------------------
# Metric builders
# ---------------------------------------------------------------------------
def build_buffett_series(session: requests.Session, start: str) -> pd.DataFrame:
    """Build the Buffett Indicator from Wilshire 5000 and nominal GDP.

    Priority order for market data:
    1. Yahoo Finance ticker ^W5000 (monthly market proxy used by several public
       Buffett dashboards)
    2. FRED WILL5000IND if Yahoo is unavailable in the runtime environment

    GDP is taken from FRED's nominal GDP series and forward-filled monthly to
    align with the market series.
    """
    try:
        wilshire = fetch_yahoo_history("^W5000", start=start, interval="1mo").rename(columns={"value": "wilshire_proxy"})
    except Exception:
        wilshire = fetch_fred_csv(session, FRED_WILSHIRE_CSV, "wilshire_proxy")
        wilshire["date"] = pd.to_datetime(wilshire["date"]).dt.to_period("M").dt.to_timestamp(how="end").dt.normalize()
        wilshire = wilshire.drop_duplicates(subset=["date"], keep="last")
        wilshire = wilshire[wilshire["date"] >= pd.Timestamp(start)]

    gdp = fetch_fred_csv(session, FRED_GDP_CSV, "nominal_gdp_billions")
    gdp["date"] = pd.to_datetime(gdp["date"]).dt.to_period("M").dt.to_timestamp(how="end").dt.normalize()

    # Monthly alignment by forward-filling the most recent quarterly GDP value.
    monthly_dates = pd.DataFrame({"date": pd.date_range(wilshire["date"].min(), wilshire["date"].max(), freq="ME")})
    gdp_monthly = pd.merge_asof(
        monthly_dates.sort_values("date"),
        gdp.sort_values("date"),
        on="date",
        direction="backward",
    )

    merged = pd.merge(wilshire, gdp_monthly, on="date", how="inner")
    merged = merged.dropna(subset=["wilshire_proxy", "nominal_gdp_billions"]).copy()
    merged["buffett_index_pct"] = (merged["wilshire_proxy"] / merged["nominal_gdp_billions"]) * 100.0

    # Log-linear trend similar to the reference site visual style.
    work = merged[["date", "buffett_index_pct"]].dropna().copy()
    work = work[work["buffett_index_pct"] > 0]
    t = np.arange(len(work), dtype=float)
    coeffs = np.polyfit(t, np.log(work["buffett_index_pct"].to_numpy()), deg=1)
    trend = np.exp(coeffs[1] + coeffs[0] * t)
    residual_log = np.log(work["buffett_index_pct"].to_numpy() / trend)
    sigma_log = float(np.nanstd(residual_log, ddof=1))

    work["buffett_trend"] = trend
    work["buffett_trend_plus_1sd"] = trend * np.exp(1.0 * sigma_log)
    work["buffett_trend_plus_2sd"] = trend * np.exp(2.0 * sigma_log)
    work["buffett_trend_minus_1sd"] = trend * np.exp(-1.0 * sigma_log)
    work["buffett_trend_minus_2sd"] = trend * np.exp(-2.0 * sigma_log)

    # Add fixed historical mean +/- standard-deviation reference levels to the
    # Excel export without disturbing the existing dynamic trend-band columns
    # used by the dashboard chart.
    work = add_stddev_level_columns(
        work,
        source_col="buffett_index_pct",
        prefix="buffett_index",
        include_mean=True,
        floor_at_zero=True,
    )
    return work.reset_index(drop=True)


def build_shiller_series(session: requests.Session, start: str) -> pd.DataFrame:
    """Build the Shiller CAPE / S&P history plus requested standard-deviation levels.

    The updated dashboard needs both series in the same figure panels, so this
    builder now precomputes the constant historical mean +/- standard deviation
    bands for *both* the CAPE ratio and the S&P 500 index. Keeping the derived
    columns here makes the plotting code much cleaner and ensures the Excel
    workbook receives the same reference levels used by the chart.
    """
    shiller = fetch_shiller_workbook(session)
    shiller = shiller[shiller["date"] >= pd.Timestamp(start)].copy()
    shiller = shiller.rename(columns={"sp500_shiller": "sp500_index", "cape": "shiller_cape"})

    # Constant historical CAPE mean and +/-0.5, +/-1.0, +/-1.5, +/-2.0 sigma
    # guide levels used in the combined top panel.
    shiller = add_stddev_level_columns(
        shiller,
        source_col="shiller_cape",
        prefix="cape",
        include_mean=True,
        floor_at_zero=True,
    )

    # Constant historical S&P 500 mean and +/-0.5, +/-1.0, +/-1.5, +/-2.0 sigma
    # guide levels used in the updated middle panel.
    shiller = add_stddev_level_columns(
        shiller,
        source_col="sp500_index",
        prefix="sp500",
        include_mean=True,
        floor_at_zero=True,
    )
    return shiller.reset_index(drop=True)



def build_combined_monthly_sheet(
    buffett: pd.DataFrame,
    shiller: pd.DataFrame,
    brk: pd.DataFrame,
    sp500_yahoo: pd.DataFrame,
) -> pd.DataFrame:
    """Create a user-friendly merged sheet with a unified monthly date spine.

    v5 keeps the dedicated Yahoo S&P export in the combined sheet and also
    preserves Berkshire source columns so users can audit the final CompaniesMarketCap-only history.
    """
    start_date = min(buffett["date"].min(), shiller["date"].min(), brk["date"].min(), sp500_yahoo["date"].min())
    end_date = max(buffett["date"].max(), shiller["date"].max(), brk["date"].max(), sp500_yahoo["date"].max())
    spine = pd.DataFrame({"date": pd.date_range(start_date, end_date, freq="ME")})

    merged = spine.merge(buffett, on="date", how="left")
    merged = merged.merge(
        shiller[[
            "date",
            "sp500_index",
            "shiller_cape",
            "cape_mean",
            "cape_plus_0_5sd",
            "cape_minus_0_5sd",
            "cape_plus_1sd",
            "cape_minus_1sd",
            "cape_plus_1_5sd",
            "cape_minus_1_5sd",
            "cape_plus_2sd",
            "cape_minus_2sd",
        ]],
        on="date",
        how="left",
    )
    merged = merged.merge(
        sp500_yahoo[[
            "date",
            "sp500_yahoo",
            "sp500_yahoo_mean",
            "sp500_yahoo_plus_0_5sd",
            "sp500_yahoo_minus_0_5sd",
            "sp500_yahoo_plus_1sd",
            "sp500_yahoo_minus_1sd",
            "sp500_yahoo_plus_1_5sd",
            "sp500_yahoo_minus_1_5sd",
            "sp500_yahoo_plus_2sd",
            "sp500_yahoo_minus_2sd",
        ]],
        on="date",
        how="left",
    )
    merged = merged.merge(
        brk[["date", "brk_cash_usd", "brk_total_assets_usd", "brk_cash_to_assets_pct", "brk_cash_source", "brk_total_assets_source", "brk_data_status"]],
        on="date",
        how="left",
    )
    return merged.sort_values("date").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Visualization
# ---------------------------------------------------------------------------
def add_grouped_trace(fig: go.Figure, trace: go.BaseTraceType, row: int, col: int = 1, secondary_y: bool = False) -> None:
    """Small wrapper so the main plotting code remains readable."""
    fig.add_trace(trace, row=row, col=col, secondary_y=secondary_y)


def _make_year_axis_settings(start_date: pd.Timestamp) -> dict:
    """Return a reusable Plotly x-axis configuration with yearly grid lines.

    The dashboard uses monthly data, but the requested presentation should mark
    each calendar year explicitly on the top and middle charts. Returning the
    settings as a dictionary keeps the row-specific axis updates concise while
    ensuring every visible x-axis uses the same tick spacing, solid yearly grid
    lines, and readable year labels.
    """
    if pd.isna(start_date):
        start_date = pd.Timestamp(DEFAULT_START)
    year_start = pd.Timestamp(year=start_date.year, month=1, day=1)
    return {
        "title_text": "Year",
        "tickformat": "%Y",
        "dtick": "M12",
        "tick0": year_start,
        "showticklabels": True,
        "showgrid": True,
        "gridcolor": "rgba(120, 120, 120, 0.45)",
        "griddash": "solid",
        "gridwidth": 1.0,
        "showline": True,
        "linecolor": "rgba(70, 70, 70, 0.80)",
        "ticks": "outside",
        "rangeslider_visible": False,
    }


def build_dashboard_figure(buffett: pd.DataFrame, shiller: pd.DataFrame, brk: pd.DataFrame) -> go.Figure:
    """Create the three-panel Plotly dashboard with the requested styling fixes.

    The updated figure intentionally:
    - shows year labels and solid yearly vertical grid lines on the shared time
      axis so the top and middle panels are easier to read;
    - colors the y-axis titles/tick labels to match their underlying series;
    - limits the middle-panel guide bands to the Buffett Indicator only; and
    - removes the duplicated Berkshire cash/assets ratio from the unified hover.
    """
    fig = make_subplots(
        rows=3,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.05,
        row_heights=[0.40, 0.32, 0.28],
        specs=[[{"secondary_y": True}], [{"secondary_y": True}], [{}]],
        subplot_titles=(
            "Buffett Indicator and Shiller CAPE ratio",
            "S&P 500 index and Buffett Indicator",
            "Berkshire Hathaway cash on hand and total assets (CompaniesMarketCap source)",
        ),
    )

    # Requested band styling: keep Buffett guide bands clearly red and CAPE
    # guide bands clearly green, while using progressively lighter shades so the
    # outer/inner levels remain visually distinct without overpowering the main
    # series. The lighter shades read like softer grey levels against a white
    # background while still preserving the requested red/green identity.
    buffett_band_colors = {
        "buffett_trend_plus_2sd": "rgba(214, 39, 40, 0.95)",
        "buffett_trend_plus_1sd": "rgba(214, 39, 40, 0.75)",
        "buffett_trend_minus_1sd": "rgba(214, 39, 40, 0.55)",
        "buffett_trend_minus_2sd": "rgba(214, 39, 40, 0.35)",
    }
    cape_band_colors = {
        "cape_plus_2sd": "rgba(44, 160, 44, 0.95)",
        "cape_plus_1sd": "rgba(44, 160, 44, 0.75)",
        "cape_minus_1sd": "rgba(44, 160, 44, 0.55)",
        "cape_minus_2sd": "rgba(44, 160, 44, 0.35)",
    }

    # --- Panel 1: Buffett Indicator + CAPE (dual axis) ---
    add_grouped_trace(
        fig,
        go.Scatter(
            x=buffett["date"],
            y=buffett["buffett_index_pct"],
            mode="lines",
            name="Buffett Indicator",
            legendgroup="buffett",
            line=dict(color="#d62728", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>Buffett Indicator: %{y:.2f}%<extra></extra>",
        ),
        row=1,
        secondary_y=False,
    )
    for ycol, label in [
        ("buffett_trend_plus_2sd", "Buffett +2 Std Dev"),
        ("buffett_trend_plus_1sd", "Buffett +1 Std Dev"),
        ("buffett_trend_minus_1sd", "Buffett -1 Std Dev"),
        ("buffett_trend_minus_2sd", "Buffett -2 Std Dev"),
    ]:
        add_grouped_trace(
            fig,
            go.Scatter(
                x=buffett["date"],
                y=buffett[ycol],
                mode="lines",
                name=label,
                legendgroup="buffett",
                showlegend=False,
                line=dict(color=buffett_band_colors[ycol], width=1.5, dash="dash"),
                hovertemplate=f"%{{x|%Y-%m}}<br>{label}: %{{y:.2f}}%<extra></extra>",
            ),
            row=1,
            secondary_y=False,
        )

    add_grouped_trace(
        fig,
        go.Scatter(
            x=shiller["date"],
            y=shiller["shiller_cape"],
            mode="lines",
            name="Shiller CAPE Ratio",
            legendgroup="cape",
            line=dict(color="#2ca02c", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>Shiller CAPE Ratio: %{y:.2f}<extra></extra>",
        ),
        row=1,
        secondary_y=True,
    )
    for ycol, label in [
        ("cape_plus_2sd", "CAPE +2 Std Dev"),
        ("cape_plus_1sd", "CAPE +1 Std Dev"),
        ("cape_minus_1sd", "CAPE -1 Std Dev"),
        ("cape_minus_2sd", "CAPE -2 Std Dev"),
    ]:
        add_grouped_trace(
            fig,
            go.Scatter(
                x=shiller["date"],
                y=shiller[ycol],
                mode="lines",
                name=label,
                legendgroup="cape",
                showlegend=False,
                line=dict(color=cape_band_colors[ycol], width=1.5, dash="dash"),
                hovertemplate=f"%{{x|%Y-%m}}<br>{label}: %{{y:.2f}}<extra></extra>",
            ),
            row=1,
            secondary_y=True,
        )

    # --- Panel 2: S&P 500 + Buffett Indicator (dual axis) ---
    add_grouped_trace(
        fig,
        go.Scatter(
            x=shiller["date"],
            y=shiller["sp500_index"],
            mode="lines",
            name="S&P 500 Index",
            legendgroup="spx",
            line=dict(color="#1f77b4", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>S&P 500 Index: %{y:,.2f}<extra></extra>",
        ),
        row=2,
        secondary_y=False,
    )

    add_grouped_trace(
        fig,
        go.Scatter(
            x=buffett["date"],
            y=buffett["buffett_index_pct"],
            mode="lines",
            name="Buffett Indicator (Middle Panel)",
            legendgroup="buffett_mid",
            line=dict(color="#d62728", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>Buffett Indicator: %{y:.2f}%<extra></extra>",
        ),
        row=2,
        secondary_y=True,
    )
    for ycol, label in [
        ("buffett_trend_plus_2sd", "Buffett +2 Std Dev (Middle Panel)"),
        ("buffett_trend_plus_1sd", "Buffett +1 Std Dev (Middle Panel)"),
        ("buffett_trend_minus_1sd", "Buffett -1 Std Dev (Middle Panel)"),
        ("buffett_trend_minus_2sd", "Buffett -2 Std Dev (Middle Panel)"),
    ]:
        clean_label = label.replace(" (Middle Panel)", "")
        add_grouped_trace(
            fig,
            go.Scatter(
                x=buffett["date"],
                y=buffett[ycol],
                mode="lines",
                name=label,
                legendgroup="buffett_mid",
                showlegend=False,
                line=dict(color=buffett_band_colors[ycol], width=1.5, dash="dash"),
                hovertemplate=f"%{{x|%Y-%m}}<br>{clean_label}: %{{y:.2f}}%<extra></extra>",
            ),
            row=2,
            secondary_y=True,
        )

    # --- Panel 3: Berkshire grouped bars with a single ratio in unified hover ---
    hover_ratio = brk["brk_cash_to_assets_pct"].map(lambda x: "N/A" if pd.isna(x) else f"{x:.2f}%")

    add_grouped_trace(
        fig,
        go.Bar(
            x=brk["date"],
            y=brk["brk_cash_usd"] / 1_000_000_000.0,
            name="BRK.B cash on hand",
            legendgroup="brk_cash",
            marker_color="#10b981",
            opacity=0.82,
            hovertemplate="%{x|%Y-%m-%d}<br>Cash: %{y:,.2f} B USD<extra></extra>",
        ),
        row=3,
        secondary_y=False,
    )
    add_grouped_trace(
        fig,
        go.Bar(
            x=brk["date"],
            y=brk["brk_total_assets_usd"] / 1_000_000_000.0,
            name="BRK.B total assets",
            legendgroup="brk_assets",
            marker_color="#f59e0b",
            opacity=0.70,
            customdata=np.column_stack([hover_ratio]),
            hovertemplate="%{x|%Y-%m-%d}<br>Total Assets: %{y:,.2f} B USD<br>Cash / Assets: %{customdata[0]}<extra></extra>",
        ),
        row=3,
        secondary_y=False,
    )

    sp500_max = float(shiller["sp500_index"].max()) if not shiller.empty else 10000.0
    sp500_axis_max = max(10000, int(math.ceil(sp500_max / 1000.0) * 1000.0))
    sp500_gridvals = list(range(1000, 10001, 1000))

    all_dates = pd.concat(
        [
            buffett[["date"]].rename(columns={"date": "date"}),
            shiller[["date"]].rename(columns={"date": "date"}),
            brk[["date"]].rename(columns={"date": "date"}),
        ],
        ignore_index=True,
    )
    first_date = pd.to_datetime(all_dates["date"], errors="coerce").dropna().min()
    year_axis_settings = _make_year_axis_settings(first_date)

    fig.update_yaxes(
        title_text="Buffett Indicator (% of GDP)",
        row=1,
        col=1,
        secondary_y=False,
        title_font=dict(color="#d62728"),
        tickfont=dict(color="#d62728"),
    )
    fig.update_yaxes(
        title_text="Shiller CAPE Ratio",
        row=1,
        col=1,
        secondary_y=True,
        title_font=dict(color="#2ca02c"),
        tickfont=dict(color="#2ca02c"),
    )
    fig.update_yaxes(
        title_text="S&P 500 Index",
        row=2,
        col=1,
        secondary_y=False,
        tickmode="array",
        tickvals=sp500_gridvals,
        ticktext=[f"{value:,}" for value in sp500_gridvals],
        range=[0, sp500_axis_max],
        showgrid=True,
        gridcolor="rgba(128, 128, 128, 0.50)",
        griddash="solid",
        gridwidth=1.0,
        zeroline=False,
        title_font=dict(color="#1f77b4"),
        tickfont=dict(color="#1f77b4"),
    )
    fig.update_yaxes(
        title_text="Buffett Indicator (% of GDP)",
        row=2,
        col=1,
        secondary_y=True,
        title_font=dict(color="#d62728"),
        tickfont=dict(color="#d62728"),
    )
    fig.update_yaxes(title_text="USD billions", row=3, col=1)

    for row in (1, 2, 3):
        fig.update_xaxes(row=row, col=1, **year_axis_settings)

    fig.update_layout(
        title="Buffett Indicator dashboard with combined Buffett/CAPE and S&P/Buffett panels",
        template="plotly_white",
        hovermode="x unified",
        barmode="group",
        height=1120,
        legend=dict(title="Click legend items to switch each output on/off", groupclick="togglegroup"),
        margin=dict(l=70, r=70, t=80, b=50),
    )
    return fig


# ---------------------------------------------------------------------------
# Export helpers
# ---------------------------------------------------------------------------

def save_excel_workbook(
    output_path: Path,
    buffett: pd.DataFrame,
    shiller: pd.DataFrame,
    sp500_yahoo: pd.DataFrame,
    brk_bundle: BerkshireHistoryBundle,
    combined: pd.DataFrame,
) -> None:
    """Write native-frequency sheets and Berkshire audit tabs.

    v5 removes the SEC Berkshire sheets completely. The workbook now exposes the
    final merged Berkshire series, the raw CompaniesMarketCap cash/assets pages,
    and a simplified audit sheet that only reflects the sole retained source.
    """
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        buffett.to_excel(writer, sheet_name="Buffett", index=False)
        shiller.to_excel(writer, sheet_name="Shiller_SP500", index=False)
        sp500_yahoo.to_excel(writer, sheet_name="SP500_Yahoo", index=False)
        brk_bundle.merged.to_excel(writer, sheet_name="Berkshire_Merged", index=False)
        brk_bundle.cmc_cash.to_excel(writer, sheet_name="Berkshire_CMC_Cash", index=False)
        brk_bundle.cmc_assets.to_excel(writer, sheet_name="Berkshire_CMC_Assets", index=False)
        brk_bundle.source_audit.to_excel(writer, sheet_name="Berkshire_Audit", index=False)
        combined.to_excel(writer, sheet_name="Combined", index=False)


# ---------------------------------------------------------------------------
# Main orchestration
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Buffett/CAPE/Berkshire dashboard and Excel workbook.")
    parser.add_argument("--start", default=DEFAULT_START, help=f"Start date for displayed data (default: {DEFAULT_START})")
    parser.add_argument("--output-dir", default=".", help="Directory where the HTML and Excel files will be saved")
    parser.add_argument("--html-name", default=DEFAULT_OUTPUT_HTML, help=f"HTML file name (default: {DEFAULT_OUTPUT_HTML})")
    parser.add_argument("--excel-name", default=DEFAULT_OUTPUT_XLSX, help=f"Excel file name (default: {DEFAULT_OUTPUT_XLSX})")
    parser.add_argument(
        "--user-agent",
        default=os.environ.get("HTTP_USER_AGENT", os.environ.get("SEC_USER_AGENT", "Your Name your.email@example.com")),
        help="Descriptive HTTP user-agent with contact information.",
    )
    return parser.parse_args()


def main() -> None:
    """Orchestrate download, transformation, charting, and workbook export.

    This function emits progress messages before and after the slowest steps so
    users can tell which source is currently being downloaded or processed.
    """
    args = parse_args()
    output_dir = ensure_output_dir(Path(args.output_dir))
    session = make_http_session(args.user_agent)
    started_at = time.perf_counter()

    requested_start = pd.Timestamp(args.start)
    enforced_cutoff = pd.Timestamp(MIN_DISPLAY_START)
    start = max(requested_start, enforced_cutoff).strftime("%Y-%m-%d")
    log_progress("Starting dashboard build", started_at)

    log_progress("Downloading Buffett Indicator inputs (Yahoo/FRED)", started_at)
    buffett = build_buffett_series(session, start)
    log_progress(f"Buffett dataset ready with {len(buffett):,} rows", started_at)

    log_progress("Downloading Shiller workbook and calculating CAPE series", started_at)
    shiller = build_shiller_series(session, start)
    log_progress(f"Shiller dataset ready with {len(shiller):,} rows", started_at)

    log_progress("Downloading dedicated Yahoo S&P 500 monthly series for Excel", started_at)
    try:
        sp500_yahoo = fetch_yahoo_history("^GSPC", start=start, interval="1mo").rename(columns={"value": "sp500_yahoo"})
        log_progress(f"Yahoo S&P dataset ready with {len(sp500_yahoo):,} rows", started_at)
    except Exception as exc:
        log_progress(f"Yahoo S&P download failed ({exc}); falling back to Shiller S&P column", started_at)
        sp500_yahoo = shiller[["date", "sp500_index"]].rename(columns={"sp500_index": "sp500_yahoo"}).copy()

    # Add the requested historical mean +/- standard-deviation levels directly
    # in the dedicated S&P worksheet before the workbook is written.
    sp500_yahoo = add_stddev_level_columns(
        sp500_yahoo,
        source_col="sp500_yahoo",
        prefix="sp500_yahoo",
        include_mean=True,
        floor_at_zero=True,
    )

    log_progress("Downloading Berkshire history from CompaniesMarketCap only", started_at)
    brk_bundle = fetch_berkshire_history(session, ticker="BRK-B")
    brk = brk_bundle.merged.copy()
    log_progress(f"Berkshire merged dataset ready with {len(brk):,} rows", started_at)

    # Apply the enforced 1989-01 minimum cutoff consistently across all exported sheets.
    cutoff = pd.Timestamp(start)
    buffett = buffett[buffett["date"] >= cutoff].copy()
    shiller = shiller[shiller["date"] >= cutoff].copy()
    sp500_yahoo = sp500_yahoo[sp500_yahoo["date"] >= cutoff].copy()
    brk_bundle.merged = brk_bundle.merged[brk_bundle.merged["date"] >= cutoff].copy()
    brk_bundle.cmc_cash = brk_bundle.cmc_cash[brk_bundle.cmc_cash["date"] >= cutoff].copy()
    brk_bundle.cmc_assets = brk_bundle.cmc_assets[brk_bundle.cmc_assets["date"] >= cutoff].copy()
    brk_bundle.source_audit = brk_bundle.source_audit[brk_bundle.source_audit["date"] >= cutoff].copy()
    brk = brk_bundle.merged

    if buffett.empty or shiller.empty or brk.empty or sp500_yahoo.empty:
        raise RuntimeError("One or more required datasets came back empty. Please check the source endpoints.")

    log_progress("Building combined monthly sheet", started_at)
    combined = build_combined_monthly_sheet(buffett, shiller, brk, sp500_yahoo)
    log_progress(f"Combined sheet ready with {len(combined):,} rows", started_at)

    log_progress("Rendering Plotly dashboard", started_at)
    figure = build_dashboard_figure(buffett=buffett, shiller=shiller, brk=brk)

    html_path = output_dir / args.html_name
    excel_path = output_dir / args.excel_name

    log_progress("Saving HTML dashboard", started_at)
    figure.write_html(str(html_path), include_plotlyjs="cdn")

    log_progress("Saving Excel workbook", started_at)
    save_excel_workbook(
        excel_path,
        buffett=buffett,
        shiller=shiller,
        sp500_yahoo=sp500_yahoo,
        brk_bundle=brk_bundle,
        combined=combined,
    )

    log_progress(f"HTML dashboard saved to: {html_path}", started_at)
    log_progress(f"Excel workbook saved to: {excel_path}", started_at)
    log_progress("Dashboard build completed", started_at)


if __name__ == "__main__":
    main()
