#!/usr/bin/env python3

"""
Buffett Indicator + Shiller CAPE + Berkshire cash dashboard 
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
- Berkshire cash/assets are sourced only from CompaniesMarketCap because
  the narrower SEC cash taxonomy used in prior versions could materially
  understate Berkshire's broader liquidity.
- Removes all Berkshire SEC download code so the script no longer mixes two
  incompatible Berkshire liquidity definitions in the same chart/workbook.
- Progress messages remain visible during execution so long-running downloads are
  easier to understand.
"""


from __future__ import annotations

import argparse
import io
import json
import math
import os
import re
import time
from urllib.parse import urljoin, urlparse
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
# Treat the legacy Yale workbook as a fallback only. When this source is used,
# the downloaded CAPE history may lag the current date by many months.
SHILLER_FRESHNESS_WARNING_DAYS = 120
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


def add_log_trend_stddev_columns(
    df: pd.DataFrame,
    source_col: str,
    prefix: str,
    *,
    include_trend: bool = True,
) -> pd.DataFrame:
    """Append log-linear trend and +/- log-residual standard-deviation bands."""
    out = df.copy()
    work = out[["date", source_col]].copy()
    work[source_col] = clean_numeric_series(work[source_col])
    work = work.dropna(subset=[source_col])
    work = work[work[source_col] > 0].copy()

    if include_trend:
        out[f"{prefix}_trend"] = np.nan
    for multiplier in STDDEV_MULTIPLIERS:
        suffix = _format_stddev_suffix(multiplier)
        out[f"{prefix}_trend_plus_{suffix}sd"] = np.nan
        out[f"{prefix}_trend_minus_{suffix}sd"] = np.nan

    if work.empty:
        return out

    t = np.arange(len(work), dtype=float)
    values = work[source_col].to_numpy(dtype=float)
    if len(work) >= 2:
        coeffs = np.polyfit(t, np.log(values), deg=1)
        trend = np.exp(coeffs[1] + coeffs[0] * t)
    else:
        trend = values.copy()

    residual_log = np.log(values / trend)
    sigma_log = float(np.nanstd(residual_log, ddof=1)) if len(work) > 1 else 0.0

    idx = work.index
    if include_trend:
        out.loc[idx, f"{prefix}_trend"] = trend
    for multiplier in STDDEV_MULTIPLIERS:
        suffix = _format_stddev_suffix(multiplier)
        out.loc[idx, f"{prefix}_trend_plus_{suffix}sd"] = trend * np.exp(multiplier * sigma_log)
        out.loc[idx, f"{prefix}_trend_minus_{suffix}sd"] = trend * np.exp(-multiplier * sigma_log)

    return out


def selected_stddev_multipliers(stddev_line_count: int) -> tuple[float, ...]:
    """Return the positive multipliers to plot for the requested line count."""
    if stddev_line_count == 8:
        return (0.5, 1.0, 1.5, 2.0)
    return (1.0, 2.0)


def make_stddev_band_specs(prefix: str, mode: str, stddev_line_count: int | None = None) -> list[tuple[str, str, float]]:
    """Build chart column/label specs for raw or log standard-deviation bands."""
    suffix_label = f"Std({mode})"
    multipliers = STDDEV_MULTIPLIERS if stddev_line_count is None else selected_stddev_multipliers(stddev_line_count)
    specs: list[tuple[str, str, float]] = []
    for multiplier in reversed(multipliers):
        suffix = _format_stddev_suffix(multiplier)
        value_label = f"+{multiplier:g} {suffix_label}"
        col = f"{prefix}_trend_plus_{suffix}sd" if mode == "log" else f"{prefix}_plus_{suffix}sd"
        specs.append((col, value_label, multiplier))
    for multiplier in multipliers:
        suffix = _format_stddev_suffix(multiplier)
        value_label = f"-{multiplier:g} {suffix_label}"
        col = f"{prefix}_trend_minus_{suffix}sd" if mode == "log" else f"{prefix}_minus_{suffix}sd"
        specs.append((col, value_label, -multiplier))
    return specs


def stddev_band_color(base_rgb: str, signed_multiplier: float) -> str:
    """Return a readable RGBA color for a standard-deviation band line."""
    opacity_by_abs = {0.5: 0.35, 1.0: 0.55, 1.5: 0.75, 2.0: 0.95}
    opacity = opacity_by_abs.get(abs(float(signed_multiplier)), 0.65)
    return f"rgba({base_rgb}, {opacity})"


def stddev_initial_visibility(show_std_lines: bool, mode: str, selected_mode: str, multiplier: float, stddev_line_count: int) -> bool:
    """Return initial Plotly visibility for custom-JS-controlled standard-deviation traces."""
    return bool(show_std_lines and mode == selected_mode and abs(float(multiplier)) in selected_stddev_multipliers(stddev_line_count))


def _matching_numeric_columns(df: pd.DataFrame, *, exact_names: Iterable[str] = (), prefixes: Iterable[str] = ()) -> list[str]:
    """Return existing numeric-like columns that match the provided names/prefixes.

    The dashboard overlays many optional standard-deviation guide lines. Plotly
    autoranges an axis from *visible* traces, so hiding/showing guide lines can
    change the main chart scale. This helper lets the figure builder collect the
    full family of columns that belong to one axis so a stable, explicit axis
    range can be calculated once during figure creation.
    """
    exact = {str(name) for name in exact_names}
    prefix_values = tuple(str(prefix) for prefix in prefixes)
    matches: list[str] = []
    for column in df.columns:
        if column in exact or any(column.startswith(prefix) for prefix in prefix_values):
            matches.append(column)
    return matches



def _stable_axis_range(
    df: pd.DataFrame,
    *,
    exact_names: Iterable[str] = (),
    prefixes: Iterable[str] = (),
    floor_at_zero: bool = False,
    pad_ratio: float = 0.06,
    fallback: tuple[float, float] = (0.0, 1.0),
) -> list[float]:
    """Build a fixed axis range that covers the main series and all guide lines.

    Why this exists:
    - the Std On/Off buttons should only show or hide the related guide lines;
    - with Plotly autorange, making those traces visible can change the y-axis;
    - using one precomputed range keeps Graph 1/2/3 visually stable.

    The function inspects every matched column, ignores missing/non-numeric
    values, adds a small padding margin, and optionally clamps the lower bound
    to zero for naturally non-negative finance metrics.
    """
    columns = _matching_numeric_columns(df, exact_names=exact_names, prefixes=prefixes)
    if not columns:
        return [float(fallback[0]), float(fallback[1])]

    numeric_series = [pd.to_numeric(df[column], errors="coerce") for column in columns]
    merged = pd.concat(numeric_series, axis=0, ignore_index=True).dropna()
    if merged.empty:
        return [float(fallback[0]), float(fallback[1])]

    min_value = float(merged.min())
    max_value = float(merged.max())
    if floor_at_zero:
        min_value = 0.0 if min_value >= 0 else min_value

    span = max_value - min_value
    if not np.isfinite(span) or span <= 0:
        span = max(abs(max_value), 1.0) * 0.10
    pad = span * float(pad_ratio)

    lower = min_value - pad
    upper = max_value + pad
    if floor_at_zero:
        lower = max(0.0, lower)
    if upper <= lower:
        upper = lower + max(abs(lower), 1.0) * 0.10
    return [float(lower), float(upper)]


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


def _is_yale_shiller_url(url: str) -> bool:
    """Return True when a candidate workbook URL points at the legacy Yale host.

    Why this helper exists:
    - The original script treated the Yale workbook as the primary source.
    - Multiple public projects now document that the Yale workbook can lag the
      current month, while shillerdata.com often exposes a newer mirrored file.
    - The dashboard still keeps Yale as a safety fallback so users get *some*
      CAPE history even when discovery fails, but that fallback should be labeled
      clearly because the newest observations may be missing.
    """
    try:
        host = urlparse(url).netloc.lower()
    except Exception:
        return False
    return 'econ.yale.edu' in host


def _shiller_data_lag_days(last_date: pd.Timestamp, as_of: Optional[pd.Timestamp] = None) -> int | None:
    """Return how many days the parsed Shiller workbook trails the latest month-end.

    Parameters
    ----------
    last_date:
        Latest month-end found in the parsed workbook.
    as_of:
        Optional override mostly useful for tests. When omitted, the function
        compares against the most recent completed month-end according to the
        local runtime clock.

    Returns
    -------
    int | None
        Number of trailing days if both dates are valid, otherwise ``None``.

    Why this helper matters:
    - URL download success alone does not guarantee fresh CAPE data.
    - The stale-data bug happened because an older Yale workbook parsed fine, so
      the code stopped before checking whether a newer source was available.
    - Measuring the gap lets the downloader prefer fresher sources and emit a
      clear warning whenever it must settle for an old fallback file.
    """
    if pd.isna(last_date):
        return None
    reference_month_end = latest_complete_month_end(as_of)
    return int((reference_month_end - pd.Timestamp(last_date)).days)


def _discover_shiller_urls(session: requests.Session) -> list[str]:
    """Discover candidate Shiller workbook URLs.

    Important behavior change:
    - Discovered shillerdata.com workbook links are now tried *before* the hard-
      coded Yale URLs. This directly fixes the stale-data bug where a readable—
      but outdated—Yale workbook caused an early return before the newer source
      could even be attempted.
    - The legacy Yale URLs remain in the list as a resilience fallback only.

    Returns
    -------
    list[str]
        Deduplicated candidate URLs ordered from most preferred to least
        preferred. Discovery failures are tolerated so the dashboard can still
        operate in degraded mode.
    """
    discovered: list[str] = []
    try:
        response = session.get(SHILLER_DISCOVERY_PAGE, timeout=30)
        response.raise_for_status()
        html = response.text
        href_pattern = r'(?:href|src)=["\']([^"\']*ie_data[^"\']*\.xls[^"\']*)["\']'
        direct_pattern = r'https?://[^"\'\s>]+ie_data[^"\'\s>]*\.xls(?:\?[^"\'\s>]*)?'
        for match in re.findall(href_pattern, html, flags=re.IGNORECASE):
            discovered.append(urljoin(SHILLER_DISCOVERY_PAGE, match))
        for match in re.findall(direct_pattern, html, flags=re.IGNORECASE):
            discovered.append(match)
    except Exception:
        # Discovery failure should not break the dashboard. The code will fall
        # back to the older Yale URLs, but a later warning makes that downgrade
        # visible to the user if a Yale workbook is ultimately selected.
        pass

    candidates: list[str] = []
    for url in [*discovered, *SHILLER_URLS]:
        if url not in candidates:
            candidates.append(url)
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

    Key robustness upgrades:
    1. Tries discovered shillerdata.com workbook links before the stale Yale URLs.
    2. Measures data freshness after parsing instead of assuming every successful
       download is current.
    3. Prints an explicit warning whenever the old Yale URL is used as fallback,
       because that path may omit the newest CAPE observations.

    The function still returns the best available workbook it can parse, even if
    every source is stale, but it makes any such degradation obvious in the log.
    """
    errors: list[str] = []
    stale_candidates: list[tuple[pd.Timestamp, bool, int | None, str, pd.DataFrame]] = []

    for url in _discover_shiller_urls(session):
        try:
            response = session.get(url, timeout=90)
            response.raise_for_status()
            content = response.content
        except Exception as exc:
            errors.append(f'{url} -> download failed: {exc}')
            continue

        parsed_frame: pd.DataFrame | None = None
        parse_errors: list[str] = []

        for strategy in [
            {'sheet_name': 'Data', 'header': 7},
            {'sheet_name': 0, 'header': 7},
        ]:
            try:
                parsed = pd.read_excel(io.BytesIO(content), engine='xlrd', **strategy)
                parsed_frame = _standardize_shiller_frame(parsed)
                break
            except Exception as exc:
                parse_errors.append(f'{url} -> strategy {strategy} failed: {exc}')

        if parsed_frame is None:
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
                parsed_frame = _standardize_shiller_frame(body)
            except Exception as exc:
                parse_errors.append(f'{url} -> raw-sheet strategy failed: {exc}')

        if parsed_frame is None:
            errors.extend(parse_errors)
            continue

        last_date = pd.Timestamp(parsed_frame['date'].max())
        lag_days = _shiller_data_lag_days(last_date)
        is_yale = _is_yale_shiller_url(url)

        # A discovered mirror is accepted immediately when it appears current.
        if (not is_yale) and (lag_days is None or lag_days <= SHILLER_FRESHNESS_WARNING_DAYS):
            log_progress(
                f"Using Shiller workbook source: {url} | latest CAPE row {last_date.date()}"
            )
            return parsed_frame

        # Otherwise keep the parsed frame as a fallback candidate but continue
        # searching for something fresher. The sort order later prefers the most
        # recent date and, on ties, prefers non-Yale sources over Yale.
        stale_candidates.append((last_date, not is_yale, lag_days, url, parsed_frame))
        if is_yale:
            if lag_days is None:
                log_progress(
                    f"WARNING: Parsed Shiller fallback from legacy Yale URL: {url}. "
                    "This fallback may miss the newest CAPE data."
                )
            else:
                log_progress(
                    f"WARNING: Parsed Shiller fallback from legacy Yale URL: {url} | "
                    f"latest CAPE row {last_date.date()} trails the latest completed month-end by {lag_days} days. "
                    "The newest CAPE data may be missing from this fallback source."
                )

    if stale_candidates:
        stale_candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
        last_date, prefers_non_yale, lag_days, url, frame = stale_candidates[0]
        if _is_yale_shiller_url(url):
            if lag_days is None:
                log_progress(
                    f"WARNING: Falling back to the old Yale Shiller workbook: {url}. "
                    "Latest CAPE data may be lost when this fallback is used."
                )
            else:
                log_progress(
                    f"WARNING: Falling back to the old Yale Shiller workbook: {url} | "
                    f"latest CAPE row {last_date.date()} trails the latest completed month-end by {lag_days} days. "
                    "Latest CAPE data may be lost when this fallback is used."
                )
        else:
            if lag_days is None:
                log_progress(
                    f"WARNING: Returning best available Shiller workbook from {url}, but freshness could not be assessed."
                )
            else:
                log_progress(
                    f"WARNING: Returning best available Shiller workbook from {url} | "
                    f"latest CAPE row {last_date.date()} trails the latest completed month-end by {lag_days} days."
                )
        return frame

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

    # Build both standard-deviation display modes:
    # - raw: fixed historical mean +/- raw standard deviation levels;
    # - log: proportional bands using std of log residuals around a log-linear trend.
    work = merged[["date", "buffett_index_pct"]].dropna().copy()
    work = add_log_trend_stddev_columns(
        work,
        source_col="buffett_index_pct",
        prefix="buffett",
        include_trend=True,
    )
    work = add_stddev_level_columns(
        work,
        source_col="buffett_index_pct",
        prefix="buffett_index",
        include_mean=True,
        floor_at_zero=True,
    )
    # Chart raw-mode aliases. Keep original buffett_index_* export columns for
    # workbook compatibility, but expose buffett_* names for unified plotting.
    for multiplier in STDDEV_MULTIPLIERS:
        suffix = _format_stddev_suffix(multiplier)
        work[f"buffett_plus_{suffix}sd"] = work[f"buffett_index_plus_{suffix}sd"]
        work[f"buffett_minus_{suffix}sd"] = work[f"buffett_index_minus_{suffix}sd"]
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
    shiller = add_log_trend_stddev_columns(
        shiller,
        source_col="shiller_cape",
        prefix="cape",
        include_trend=True,
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



def prepare_sp500_yahoo_frame_for_dashboard(sp500_yahoo: pd.DataFrame) -> pd.DataFrame:
    """Return a dashboard-ready S&P 500 Yahoo frame with stable export columns.

    Why this helper exists:
    - ``fetch_yahoo_history()`` returns a generic two-column frame named
      ``date`` + ``value`` rather than the dashboard-specific ``sp500_yahoo``
      column expected by later merge/plot functions.
    - ``build_combined_monthly_sheet()`` also expects the raw mean and standard-
      deviation export columns to exist, but those columns are only guaranteed in
      the plotting path unless we normalize the Yahoo frame explicitly.

    What this helper does:
    1) renames the Yahoo price column to ``sp500_yahoo`` when needed;
    2) coerces the date/value columns into the expected types and order; and
    3) adds the raw standard-deviation export columns when they are missing.

    Keeping this logic in one reusable helper prevents future regressions where
    one call path prepares the S&P frame but another call path forgets to do so.
    """
    frame = sp500_yahoo.copy()
    if frame.empty:
        return frame

    # ``fetch_yahoo_history()`` returns ``value`` for all tickers.  The dashboard
    # code, workbook export, and Panel 2/3 plotting logic all expect the more
    # explicit ``sp500_yahoo`` column name, so normalize it here.
    if "sp500_yahoo" not in frame.columns:
        if "value" in frame.columns:
            frame = frame.rename(columns={"value": "sp500_yahoo"})
        elif "close" in frame.columns:
            frame = frame.rename(columns={"close": "sp500_yahoo"})
        elif "Close" in frame.columns:
            frame = frame.rename(columns={"Close": "sp500_yahoo"})
        else:
            raise KeyError(
                "S&P 500 Yahoo frame is missing the expected price column. "
                f"Available columns: {list(frame.columns)}"
            )

    frame["date"] = pd.to_datetime(frame["date"], errors="coerce")
    frame["sp500_yahoo"] = pd.to_numeric(frame["sp500_yahoo"], errors="coerce")
    frame = frame.dropna(subset=["date", "sp500_yahoo"]).sort_values("date").copy()

    # Rebuild the raw mean +/- std-dev export columns every time this helper runs.
    # This keeps the workbook/chart export consistent even when callers first
    # download the raw Yahoo frame, then later filter it down to the latest fully
    # completed month before exporting or plotting.
    raw_std_columns = [
        "sp500_yahoo_mean",
        "sp500_yahoo_plus_0_5sd",
        "sp500_yahoo_minus_0_5sd",
        "sp500_yahoo_plus_1sd",
        "sp500_yahoo_minus_1sd",
        "sp500_yahoo_plus_1_5sd",
        "sp500_yahoo_minus_1_5sd",
        "sp500_yahoo_plus_2sd",
        "sp500_yahoo_minus_2sd",
    ]
    frame = frame.drop(columns=[col for col in raw_std_columns if col in frame.columns])
    frame = add_stddev_level_columns(
        frame,
        source_col="sp500_yahoo",
        prefix="sp500_yahoo",
        include_mean=True,
        floor_at_zero=True,
    )

    return frame.reset_index(drop=True)


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
    # Normalize the Yahoo S&P frame here so workbook export remains robust even
    # if future callers pass the raw ``fetch_yahoo_history()`` output directly.
    sp500_yahoo = prepare_sp500_yahoo_frame_for_dashboard(sp500_yahoo)

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

    The centered Plotly x-axis title is intentionally disabled.  The dashboard
    adds a custom bold black ``Year`` annotation at the far right of each panel
    so the label sits to the right of the latest visible year instead of below
    the middle of the axis.
    """
    if pd.isna(start_date):
        start_date = pd.Timestamp(DEFAULT_START)
    year_start = pd.Timestamp(year=start_date.year, month=1, day=1)
    return {
        "title_text": "",
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
        "tickfont": dict(color="black"),
        "rangeslider_visible": False,
    }


def latest_complete_month_end(as_of: Optional[pd.Timestamp] = None) -> pd.Timestamp:
    """Return the latest completed month-end for monthly market data.

    Example: when the dashboard is run on 2026-06-04, June is still incomplete,
    so the latest complete S&P 500 month-end must be 2026-05-31.
    """
    as_of_ts = pd.Timestamp.today().normalize() if as_of is None else pd.Timestamp(as_of).normalize()
    if as_of_ts.is_month_end:
        return as_of_ts
    return (as_of_ts.to_period("M") - 1).to_timestamp(how="end").normalize()


def build_dashboard_figure(
    buffett: pd.DataFrame,
    shiller: pd.DataFrame,
    brk: pd.DataFrame,
    sp500_yahoo: Optional[pd.DataFrame] = None,
    *,
    stddev_modes: dict[str, str],
    show_std_lines: dict[str, bool],
    stddev_line_counts: dict[str, int],
) -> go.Figure:
    """Create the four-panel Plotly dashboard with the requested graph fixes.

    Fixes included:
    - larger vertical spacing between the four panels;
    - custom bold black ``Year`` labels placed next to the latest year tick;
    - extra right-side plotting room so standard-deviation labels and right y-axes do not collide;
    - the middle-panel S&P 500 line uses the dedicated Yahoo series so it
      continues past the Shiller workbook cutoff; and
    - the Berkshire bar hover puts the year and Cash / Assets (%) at the top.
    """
    if sp500_yahoo is None or sp500_yahoo.empty:
        sp500_plot = shiller[["date", "sp500_index"]].rename(columns={"sp500_index": "sp500_yahoo"}).copy()
    else:
        sp500_plot = sp500_yahoo[["date", "sp500_yahoo"]].copy()
    sp500_plot["date"] = pd.to_datetime(sp500_plot["date"], errors="coerce")
    sp500_plot["sp500_yahoo"] = pd.to_numeric(sp500_plot["sp500_yahoo"], errors="coerce")
    sp500_plot = sp500_plot.dropna(subset=["date", "sp500_yahoo"]).sort_values("date")

    sp500_band_source = sp500_yahoo.copy() if sp500_yahoo is not None and not sp500_yahoo.empty else sp500_plot.copy()
    sp500_band_source["date"] = pd.to_datetime(sp500_band_source["date"], errors="coerce")
    sp500_band_source["sp500_yahoo"] = pd.to_numeric(sp500_band_source["sp500_yahoo"], errors="coerce")
    sp500_band_source = sp500_band_source.dropna(subset=["date", "sp500_yahoo"]).sort_values("date")
    if "sp500_yahoo_plus_1sd" not in sp500_band_source.columns:
        sp500_band_source = add_stddev_level_columns(
            sp500_band_source,
            source_col="sp500_yahoo",
            prefix="sp500_yahoo",
            include_mean=True,
            floor_at_zero=True,
        )
    if "sp500_yahoo_trend_plus_1sd" not in sp500_band_source.columns:
        sp500_band_source = add_log_trend_stddev_columns(
            sp500_band_source,
            source_col="sp500_yahoo",
            prefix="sp500_yahoo",
            include_trend=True,
        )

    fig = make_subplots(
        rows=4,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.08,
        row_heights=[0.30, 0.24, 0.24, 0.22],
        specs=[[{"secondary_y": True}], [{"secondary_y": True}], [{"secondary_y": True}], [{}]],
        subplot_titles=(
            "Buffett Indicator and Shiller CAPE ratio",
            "S&P 500 index and Buffett Indicator",
            "S&P 500 index and Shiller CAPE ratio",
            "Berkshire Hathaway cash on hand and total assets (CompaniesMarketCap source)",
        ),
    )

    stddev_mode_defaults = {
        "cape_graph1": "raw",
        "buffett_graph1": "log",
        "sp500_graph2": "raw",
        "buffett_graph2": "log",
        "sp500_graph3": "raw",
        "cape_graph3": "raw",
    }
    stddev_visible_defaults = {
        "cape_graph1": True,
        "buffett_graph1": True,
        "sp500_graph2": False,
        "buffett_graph2": True,
        "sp500_graph3": False,
        "cape_graph3": True,
    }
    stddev_modes = {
        group: str(stddev_modes.get(group, default)).lower()
        for group, default in stddev_mode_defaults.items()
    }
    show_std_lines = {
        group: bool(show_std_lines.get(group, default))
        for group, default in stddev_visible_defaults.items()
    }
    stddev_line_counts = {
        group: int(stddev_line_counts.get(group, 4))
        for group in stddev_mode_defaults
    }
    for group, mode in stddev_modes.items():
        if mode not in {"raw", "log"}:
            raise ValueError(f"stddev mode for {group} must be 'raw' or 'log'")
    for group, line_count in stddev_line_counts.items():
        if line_count not in {4, 8}:
            raise ValueError(f"stddev line count for {group} must be 4 or 8")

    # Add all raw/log and 4/8-line candidates. Custom JavaScript controls decide
    # which mutually-exclusive mode/count is visible in the exported HTML.
    buffett_band_specs_by_mode = {
        "raw": make_stddev_band_specs("buffett", "raw", None),
        "log": make_stddev_band_specs("buffett", "log", None),
    }
    cape_band_specs_by_mode = {
        "raw": make_stddev_band_specs("cape", "raw", None),
        "log": make_stddev_band_specs("cape", "log", None),
    }
    sp500_band_specs_by_mode = {
        "raw": make_stddev_band_specs("sp500_yahoo", "raw", None),
        "log": make_stddev_band_specs("sp500_yahoo", "log", None),
    }
    # Graph 3 reuses the same underlying S&P and CAPE calculations, but it gets
    # its own Std-control groups so the top switchers affect only the matching graph.
    sp500_graph3_band_specs_by_mode = {
        "raw": make_stddev_band_specs("sp500_yahoo", "raw", None),
        "log": make_stddev_band_specs("sp500_yahoo", "log", None),
    }
    cape_graph3_band_specs_by_mode = {
        "raw": make_stddev_band_specs("cape", "raw", None),
        "log": make_stddev_band_specs("cape", "log", None),
    }

    # --- Panel 1: Buffett Indicator + CAPE (dual axis) ---
    add_grouped_trace(
        fig,
        go.Scatter(
            x=buffett["date"],
            y=buffett["buffett_index_pct"],
            mode="lines",
            name="Buffett Indicator",
            legendgroup="graph_1",
            legendgrouptitle_text="Graph 1 · Buffett Indicator + CAPE",
            line=dict(color="#d62728", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>Buffett Indicator: %{y:.2f}%<extra></extra>",
        ),
        row=1,
        secondary_y=False,
    )
    for mode, specs in buffett_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            label = f"Buffett {short_label}"
            visible = stddev_initial_visibility(show_std_lines["buffett_graph1"], mode, stddev_modes["buffett_graph1"], multiplier, stddev_line_counts["buffett_graph1"])
            add_grouped_trace(
                fig,
                go.Scatter(
                    x=buffett["date"],
                    y=buffett[ycol],
                    mode="lines",
                    name=label,
                    legendgroup="buffett_std",
                    showlegend=False,
                    visible=visible,
                    meta=dict(stdControl=True, stdGroup="buffett_graph1", stdMode=mode, stdMultiplier=abs(float(multiplier))),
                    line=dict(color=stddev_band_color("214, 39, 40", multiplier), width=1.5, dash="dash"),
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
            legendgroup="graph_1",
            line=dict(color="#2ca02c", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>Shiller CAPE Ratio: %{y:.2f}<extra></extra>",
        ),
        row=1,
        secondary_y=True,
    )
    for mode, specs in cape_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            label = f"CAPE {short_label}"
            visible = stddev_initial_visibility(show_std_lines["cape_graph1"], mode, stddev_modes["cape_graph1"], multiplier, stddev_line_counts["cape_graph1"])
            add_grouped_trace(
                fig,
                go.Scatter(
                    x=shiller["date"],
                    y=shiller[ycol],
                    mode="lines",
                    name=label,
                    legendgroup="cape_std",
                    showlegend=False,
                    visible=visible,
                    meta=dict(stdControl=True, stdGroup="cape_graph1", stdMode=mode, stdMultiplier=abs(float(multiplier))),
                    line=dict(color=stddev_band_color("44, 160, 44", multiplier), width=1.5, dash="dash"),
                    hovertemplate=f"%{{x|%Y-%m}}<br>{label}: %{{y:.2f}}<extra></extra>",
                ),
                row=1,
                secondary_y=True,
            )

    # --- Panel 2: S&P 500 + Buffett Indicator (dual axis) ---
    add_grouped_trace(
        fig,
        go.Scatter(
            x=sp500_plot["date"],
            y=sp500_plot["sp500_yahoo"],
            mode="lines",
            name="S&P 500 Index",
            legendgroup="graph_2",
            legendgrouptitle_text="Graph 2 · S&P 500 + Buffett Indicator",
            line=dict(color="#1f77b4", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>S&P 500 Index: %{y:,.2f}<extra></extra>",
        ),
        row=2,
        secondary_y=False,
    )
    for mode, specs in sp500_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            label = f"S&P 500 {short_label}"
            visible = stddev_initial_visibility(show_std_lines["sp500_graph2"], mode, stddev_modes["sp500_graph2"], multiplier, stddev_line_counts["sp500_graph2"])
            add_grouped_trace(
                fig,
                go.Scatter(
                    x=sp500_band_source["date"],
                    y=sp500_band_source[ycol],
                    mode="lines",
                    name=label,
                    legendgroup="sp500_std",
                    showlegend=False,
                    visible=visible,
                    meta=dict(stdControl=True, stdGroup="sp500_graph2", stdMode=mode, stdMultiplier=abs(float(multiplier))),
                    line=dict(color=stddev_band_color("31, 119, 180", multiplier), width=1.5, dash="dash"),
                    hovertemplate=f"%{{x|%Y-%m}}<br>{label}: %{{y:,.2f}}<extra></extra>",
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
            legendgroup="graph_2",
            line=dict(color="#d62728", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>Buffett Indicator: %{y:.2f}%<extra></extra>",
        ),
        row=2,
        secondary_y=True,
    )
    for mode, specs in buffett_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            label = f"Buffett {short_label} (Middle Panel)"
            clean_label = label.replace(" (Middle Panel)", "")
            visible = stddev_initial_visibility(show_std_lines["buffett_graph2"], mode, stddev_modes["buffett_graph2"], multiplier, stddev_line_counts["buffett_graph2"])
            add_grouped_trace(
                fig,
                go.Scatter(
                    x=buffett["date"],
                    y=buffett[ycol],
                    mode="lines",
                    name=label,
                    legendgroup="buffett_std",
                    showlegend=False,
                    visible=visible,
                    meta=dict(stdControl=True, stdGroup="buffett_graph2", stdMode=mode, stdMultiplier=abs(float(multiplier))),
                    line=dict(color=stddev_band_color("214, 39, 40", multiplier), width=1.5, dash="dash"),
                    hovertemplate=f"%{{x|%Y-%m}}<br>{clean_label}: %{{y:.2f}}%<extra></extra>",
                ),
                row=2,
                secondary_y=True,
            )

    # --- Panel 3: S&P 500 + Shiller CAPE (dual axis) ---
    add_grouped_trace(
        fig,
        go.Scatter(
            x=sp500_plot["date"],
            y=sp500_plot["sp500_yahoo"],
            mode="lines",
            name="S&P 500 Index",
            legendgroup="graph_3",
            legendgrouptitle_text="Graph 3 · S&P 500 + Shiller CAPE",
            line=dict(color="#1f77b4", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>S&P 500 Index: %{y:,.2f}<extra></extra>",
        ),
        row=3,
        secondary_y=False,
    )
    for mode, specs in sp500_graph3_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            label = f"S&P 500 {short_label}"
            clean_label = label
            visible = stddev_initial_visibility(show_std_lines["sp500_graph3"], mode, stddev_modes["sp500_graph3"], multiplier, stddev_line_counts["sp500_graph3"])
            add_grouped_trace(
                fig,
                go.Scatter(
                    x=sp500_band_source["date"],
                    y=sp500_band_source[ycol],
                    mode="lines",
                    name=label,
                    legendgroup="sp500_graph3_std",
                    showlegend=False,
                    visible=visible,
                    meta=dict(stdControl=True, stdGroup="sp500_graph3", stdMode=mode, stdMultiplier=abs(float(multiplier))),
                    line=dict(color=stddev_band_color("31, 119, 180", multiplier), width=1.5, dash="dash"),
                    hovertemplate=f"%{{x|%Y-%m}}<br>{clean_label}: %{{y:,.2f}}<extra></extra>",
                ),
                row=3,
                secondary_y=False,
            )

    add_grouped_trace(
        fig,
        go.Scatter(
            x=shiller["date"],
            y=shiller["shiller_cape"],
            mode="lines",
            name="Shiller CAPE Ratio",
            legendgroup="graph_3",
            line=dict(color="#2ca02c", width=3.2),
            hovertemplate="%{x|%Y-%m}<br>Shiller CAPE Ratio: %{y:.2f}<extra></extra>",
        ),
        row=3,
        secondary_y=True,
    )
    for mode, specs in cape_graph3_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            label = f"CAPE {short_label}"
            clean_label = label
            visible = stddev_initial_visibility(
                show_std_lines["cape_graph3"],
                mode,
                stddev_modes["cape_graph3"],
                multiplier,
                stddev_line_counts["cape_graph3"],
            )
            add_grouped_trace(
                fig,
                go.Scatter(
                    x=shiller["date"],
                    y=shiller[ycol],
                    mode="lines",
                    name=label,
                    legendgroup="cape_graph3_std",
                    showlegend=False,
                    visible=visible,
                    meta=dict(stdControl=True, stdGroup="cape_graph3", stdMode=mode, stdMultiplier=abs(float(multiplier))),
                    line=dict(color=stddev_band_color("44, 160, 44", multiplier), width=1.5, dash="dash"),
                    hovertemplate=f"%{{x|%Y-%m}}<br>{clean_label}: %{{y:.2f}}<extra></extra>",
                ),
                row=3,
                secondary_y=True,
            )

    # --- Panel 4: Berkshire grouped bars ---
    # A single invisible hover carrier prevents duplicated year rows in unified hover.
    # The visible bars keep the legend and visual encoding, but do not emit hover rows.
    # The unified hover title remains the single year; the body starts with Cash / Assets (%).
    hover_ratio = brk["brk_cash_to_assets_pct"].map(lambda x: "N/A" if pd.isna(x) else f"{x:.2f}%")
    hover_cash = (brk["brk_cash_usd"] / 1_000_000_000.0).map(lambda x: "N/A" if pd.isna(x) else f"{x:,.2f} B USD")
    hover_assets = (brk["brk_total_assets_usd"] / 1_000_000_000.0).map(lambda x: "N/A" if pd.isna(x) else f"{x:,.2f} B USD")
    brk_customdata = np.column_stack([hover_cash, hover_assets, hover_ratio])

    add_grouped_trace(
        fig,
        go.Scatter(
            x=brk["date"],
            y=brk["brk_total_assets_usd"] / 1_000_000_000.0,
            mode="markers",
            name="Berkshire hover details",
            showlegend=False,
            marker=dict(size=28, color="rgba(0,0,0,0)"),
            customdata=brk_customdata,
            hovertemplate=(
                "<b>Cash / Assets (%) : %{customdata[2]}</b>"
                "<br><br><span style='color:#10b981'>■</span> "
                "<b><span style='color:#10b981'>Cash: %{customdata[0]}</span></b>"
                "<br><br><span style='color:#f59e0b'>■</span> "
                "<b><span style='color:#f59e0b'>Total Assets: %{customdata[1]}</span></b>"
                "<extra></extra>"
            ),
        ),
        row=4,
        secondary_y=False,
    )

    add_grouped_trace(
        fig,
        go.Bar(
            x=brk["date"],
            y=brk["brk_cash_usd"] / 1_000_000_000.0,
            name="BRK.B cash on hand",
            legendgroup="panel_4",
            legendgrouptitle_text="Graph 4 · Berkshire cash + assets",
            marker_color="#10b981",
            opacity=0.82,
            hoverinfo="skip",
            hovertemplate=None,
        ),
        row=4,
        secondary_y=False,
    )
    add_grouped_trace(
        fig,
        go.Bar(
            x=brk["date"],
            y=brk["brk_total_assets_usd"] / 1_000_000_000.0,
            name="BRK.B total assets",
            legendgroup="panel_4",
            marker_color="#f59e0b",
            opacity=0.70,
            hoverinfo="skip",
            hovertemplate=None,
        ),
        row=4,
        secondary_y=False,
    )

    sp500_range_columns = _matching_numeric_columns(
        sp500_band_source,
        exact_names=("sp500_yahoo",),
        prefixes=("sp500_yahoo_",),
    )
    sp500_max = 10000.0
    if sp500_range_columns:
        sp500_numeric = pd.concat(
            [pd.to_numeric(sp500_band_source[column], errors="coerce") for column in sp500_range_columns],
            axis=0,
            ignore_index=True,
        ).dropna()
        if not sp500_numeric.empty:
            sp500_max = float(sp500_numeric.max())
    sp500_axis_max = max(10000, int(math.ceil(sp500_max / 1000.0) * 1000.0))
    sp500_gridvals = list(range(1000, sp500_axis_max + 1, 1000))
    buffett_axis_range = _stable_axis_range(
        buffett,
        exact_names=("buffett_index_pct",),
        prefixes=("buffett_",),
        floor_at_zero=True,
        fallback=(0.0, 100.0),
    )
    cape_axis_range = _stable_axis_range(
        shiller,
        exact_names=("shiller_cape",),
        prefixes=("cape_",),
        floor_at_zero=True,
        fallback=(0.0, 40.0),
    )

    all_dates = pd.concat(
        [
            buffett[["date"]].rename(columns={"date": "date"}),
            shiller[["date"]].rename(columns={"date": "date"}),
            sp500_plot[["date"]].rename(columns={"date": "date"}),
            brk[["date"]].rename(columns={"date": "date"}),
        ],
        ignore_index=True,
    )
    clean_dates = pd.to_datetime(all_dates["date"], errors="coerce").dropna()
    first_date = clean_dates.min()
    last_data_date = clean_dates.max()
    # Keep generous right-side whitespace for Std Dev labels before the right-side y-axis text.
    # This also extends the horizontal grid lines to the right.
    label_right_date = (last_data_date + pd.DateOffset(months=28)).normalize()
    x_title_date = (last_data_date + pd.DateOffset(months=6)).normalize()
    year_axis_settings = _make_year_axis_settings(first_date)
    year_axis_settings["range"] = [first_date, label_right_date]

    fig.update_yaxes(
        title_text="Buffett Indicator(Stock Market Value/GDP)",
        row=1,
        col=1,
        secondary_y=False,
        range=buffett_axis_range,
        autorange=False,
        title_font=dict(color="#d62728"),
        tickfont=dict(color="#d62728"),
        automargin=True,
        title_standoff=34,
    )
    fig.update_yaxes(
        title_text="Shiller CAPE Ratio",
        row=1,
        col=1,
        secondary_y=True,
        range=cape_axis_range,
        autorange=False,
        title_font=dict(color="#2ca02c"),
        tickfont=dict(color="#2ca02c"),
        automargin=True,
        title_standoff=34,
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
        autorange=False,
        showgrid=True,
        gridcolor="rgba(128, 128, 128, 0.50)",
        griddash="solid",
        gridwidth=1.0,
        zeroline=False,
        title_font=dict(color="#1f77b4"),
        tickfont=dict(color="#1f77b4"),
        automargin=True,
        title_standoff=34,
    )
    fig.update_yaxes(
        title_text="Buffett Indicator(Stock Market Value/GDP)",
        row=2,
        col=1,
        secondary_y=True,
        range=buffett_axis_range,
        autorange=False,
        title_font=dict(color="#d62728"),
        tickfont=dict(color="#d62728"),
        automargin=True,
        title_standoff=34,
    )
    fig.update_yaxes(
        title_text="S&P 500 Index",
        row=3,
        col=1,
        secondary_y=False,
        tickmode="array",
        tickvals=sp500_gridvals,
        ticktext=[f"{value:,}" for value in sp500_gridvals],
        range=[0, sp500_axis_max],
        autorange=False,
        showgrid=True,
        gridcolor="rgba(128, 128, 128, 0.50)",
        griddash="solid",
        gridwidth=1.0,
        zeroline=False,
        title_font=dict(color="#1f77b4"),
        tickfont=dict(color="#1f77b4"),
        automargin=True,
        title_standoff=34,
    )
    fig.update_yaxes(
        title_text="Shiller CAPE Ratio",
        row=3,
        col=1,
        secondary_y=True,
        range=cape_axis_range,
        autorange=False,
        title_font=dict(color="#2ca02c"),
        tickfont=dict(color="#2ca02c"),
        automargin=True,
        title_standoff=34,
    )
    fig.update_yaxes(title_text="USD billions", row=4, col=1, automargin=True, title_standoff=34)

    for row in (1, 2, 3, 4):
        fig.update_xaxes(row=row, col=1, **year_axis_settings)

    fig.update_layout(
        title="Buffett Indicator dashboard with Buffett/CAPE, S&P/Buffett, S&P/CAPE, and Berkshire panels",
        template="plotly_white",
        hovermode="x unified",
        hoverlabel=dict(font=dict(color="black")),
        barmode="group",
        height=1640,
        legend=dict(
            title="Click legend items to switch each output on/off",
            groupclick="toggleitem",
            tracegroupgap=12,
        ),
        margin=dict(l=70, r=340, t=90, b=70),
    )

    def _last_valid_point(df: pd.DataFrame, ycol: str) -> tuple[pd.Timestamp, float] | None:
        work = df[["date", ycol]].copy()
        work["date"] = pd.to_datetime(work["date"], errors="coerce")
        work[ycol] = pd.to_numeric(work[ycol], errors="coerce")
        work = work.dropna(subset=["date", ycol]).sort_values("date")
        if work.empty:
            return None
        row = work.iloc[-1]
        return pd.Timestamp(row["date"]), float(row[ycol])

    def _add_line_end_label(
        df: pd.DataFrame,
        ycol: str,
        label: str,
        color: str,
        yref: str,
        x_month_offset: int = 4,
        yshift: int = 0,
    ) -> None:
        point = _last_valid_point(df, ycol)
        if point is None:
            return
        x_value, y_value = point
        fig.add_annotation(
            x=x_value + pd.DateOffset(months=x_month_offset),
            y=y_value,
            xref="x",
            yref=yref,
            text=f"<b>{label}</b>",
            showarrow=False,
            xanchor="left",
            yanchor="middle",
            font=dict(color=color, size=11),
            bgcolor="rgba(255,255,255,0.70)",
            borderpad=1,
            yshift=yshift,
        )

    # Label all standard-deviation guide lines at their right endpoints. Custom
    # JavaScript toggles each label with the corresponding raw/log/count state.
    def _add_std_label(
        df: pd.DataFrame,
        ycol: str,
        short_label: str,
        color: str,
        yref: str,
        group: str,
        mode: str,
        multiplier: float,
        visible: bool,
    ) -> None:
        before_count = len(fig.layout.annotations) if fig.layout.annotations else 0
        _add_line_end_label(df, ycol, short_label, color, yref=yref, yshift=0)
        after_count = len(fig.layout.annotations) if fig.layout.annotations else 0
        if after_count > before_count:
            fig.layout.annotations[-1].update(
                visible=visible,
                name=f"stdLabel:{group}:{mode}:{abs(float(multiplier)):g}",
            )

    for mode, specs in buffett_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            buffett_graph1_visible = stddev_initial_visibility(show_std_lines["buffett_graph1"], mode, stddev_modes["buffett_graph1"], multiplier, stddev_line_counts["buffett_graph1"])
            _add_std_label(buffett, ycol, short_label, stddev_band_color("214, 39, 40", multiplier), "y", "buffett_graph1", mode, multiplier, buffett_graph1_visible)
            buffett_graph2_visible = stddev_initial_visibility(show_std_lines["buffett_graph2"], mode, stddev_modes["buffett_graph2"], multiplier, stddev_line_counts["buffett_graph2"])
            _add_std_label(buffett, ycol, short_label, stddev_band_color("214, 39, 40", multiplier), "y4", "buffett_graph2", mode, multiplier, buffett_graph2_visible)
    for mode, specs in cape_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            visible = stddev_initial_visibility(show_std_lines["cape_graph1"], mode, stddev_modes["cape_graph1"], multiplier, stddev_line_counts["cape_graph1"])
            _add_std_label(shiller, ycol, short_label, stddev_band_color("44, 160, 44", multiplier), "y2", "cape_graph1", mode, multiplier, visible)
    for mode, specs in sp500_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            visible = stddev_initial_visibility(show_std_lines["sp500_graph2"], mode, stddev_modes["sp500_graph2"], multiplier, stddev_line_counts["sp500_graph2"])
            _add_std_label(sp500_band_source, ycol, short_label, stddev_band_color("31, 119, 180", multiplier), "y3", "sp500_graph2", mode, multiplier, visible)
    for mode, specs in sp500_graph3_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            visible = stddev_initial_visibility(show_std_lines["sp500_graph3"], mode, stddev_modes["sp500_graph3"], multiplier, stddev_line_counts["sp500_graph3"])
            _add_std_label(sp500_band_source, ycol, short_label, stddev_band_color("31, 119, 180", multiplier), "y5", "sp500_graph3", mode, multiplier, visible)
    for mode, specs in cape_graph3_band_specs_by_mode.items():
        for ycol, short_label, multiplier in specs:
            visible = stddev_initial_visibility(show_std_lines["cape_graph3"], mode, stddev_modes["cape_graph3"], multiplier, stddev_line_counts["cape_graph3"])
            _add_std_label(shiller, ycol, short_label, stddev_band_color("44, 160, 44", multiplier), "y6", "cape_graph3", mode, multiplier, visible)

    # Custom x-axis titles, one per panel, positioned just to the right of the latest visible year.
    for axis_name in ("yaxis", "yaxis3", "yaxis5", "yaxis7"):
        domain = getattr(fig.layout, axis_name).domain
        fig.add_annotation(
            x=x_title_date,
            y=max(domain[0] - 0.025, 0.01),
            xref="x",
            yref="paper",
            text="<b>Year</b>",
            showarrow=False,
            xanchor="left",
            yanchor="top",
            font=dict(color="black", size=12),
        )

    return fig


# ---------------------------------------------------------------------------
# Export helpers
# ---------------------------------------------------------------------------

def build_stddev_control_panel_html(initial_state: dict[str, object], plot_id: str) -> str:
    """Return grouped HTML/JS controls for standard-deviation traces.

    Each graph owns its own complete control set so every switcher affects only
    the matching graph series. The JavaScript state is keyed by the same
    graph-owned ``stdGroup`` values used by Plotly trace metadata, which keeps
    HTML controls, trace visibility, and CLI startup settings aligned.
    """
    state_json = json.dumps(initial_state)
    return f"""
<div id="stddev-control-panel" class="stddev-control-panel">
  <div class="stddev-control-title">Std line controls</div>
  <div class="stddev-control-grid">
    <section class="stddev-control-group">
      <div class="stddev-control-group-title">Graph 1 · Buffett Indicator + CAPE</div>
      <div class="stddev-control-metric">
        <div class="stddev-control-row">
          <span class="stddev-control-label">CAPE mode</span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="mode" data-group="cape_graph1" data-value="raw">Raw</button>
            <button type="button" data-std-control="mode" data-group="cape_graph1" data-value="log">Log</button>
          </div>
        </div>
        <div class="stddev-control-row stddev-control-row--toggle">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <button type="button" data-std-control="toggle" data-group="cape_graph1">CAPE Std On/Off</button>
        </div>
        <div class="stddev-control-row stddev-control-row--line-count">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="lineCount" data-group="cape_graph1" data-value="4">4 lines</button>
            <button type="button" data-std-control="lineCount" data-group="cape_graph1" data-value="8">8 lines</button>
          </div>
        </div>
      </div>
      <div class="stddev-control-metric">
        <div class="stddev-control-row">
          <span class="stddev-control-label">Buffett mode</span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="mode" data-group="buffett_graph1" data-value="raw">Raw</button>
            <button type="button" data-std-control="mode" data-group="buffett_graph1" data-value="log">Log</button>
          </div>
        </div>
        <div class="stddev-control-row stddev-control-row--toggle">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <button type="button" data-std-control="toggle" data-group="buffett_graph1">Buffett Std On/Off</button>
        </div>
        <div class="stddev-control-row stddev-control-row--line-count">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="lineCount" data-group="buffett_graph1" data-value="4">4 lines</button>
            <button type="button" data-std-control="lineCount" data-group="buffett_graph1" data-value="8">8 lines</button>
          </div>
        </div>
      </div>
    </section>

    <section class="stddev-control-group">
      <div class="stddev-control-group-title">Graph 2 · S&amp;P 500 + Buffett Indicator</div>
      <div class="stddev-control-metric">
        <div class="stddev-control-row">
          <span class="stddev-control-label">S&amp;P 500 mode</span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="mode" data-group="sp500_graph2" data-value="raw">Raw</button>
            <button type="button" data-std-control="mode" data-group="sp500_graph2" data-value="log">Log</button>
          </div>
        </div>
        <div class="stddev-control-row stddev-control-row--toggle">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <button type="button" data-std-control="toggle" data-group="sp500_graph2">S&amp;P 500 Std On/Off</button>
        </div>
        <div class="stddev-control-row stddev-control-row--line-count">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="lineCount" data-group="sp500_graph2" data-value="4">4 lines</button>
            <button type="button" data-std-control="lineCount" data-group="sp500_graph2" data-value="8">8 lines</button>
          </div>
        </div>
      </div>
      <div class="stddev-control-metric">
        <div class="stddev-control-row">
          <span class="stddev-control-label">Buffett mode</span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="mode" data-group="buffett_graph2" data-value="raw">Raw</button>
            <button type="button" data-std-control="mode" data-group="buffett_graph2" data-value="log">Log</button>
          </div>
        </div>
        <div class="stddev-control-row stddev-control-row--toggle">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <button type="button" data-std-control="toggle" data-group="buffett_graph2">Buffett Std On/Off</button>
        </div>
        <div class="stddev-control-row stddev-control-row--line-count">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="lineCount" data-group="buffett_graph2" data-value="4">4 lines</button>
            <button type="button" data-std-control="lineCount" data-group="buffett_graph2" data-value="8">8 lines</button>
          </div>
        </div>
      </div>
    </section>

    <section class="stddev-control-group">
      <div class="stddev-control-group-title">Graph 3 · S&amp;P 500 + Shiller CAPE</div>
      <div class="stddev-control-metric">
        <div class="stddev-control-row">
          <span class="stddev-control-label">S&amp;P 500 mode</span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="mode" data-group="sp500_graph3" data-value="raw">Raw</button>
            <button type="button" data-std-control="mode" data-group="sp500_graph3" data-value="log">Log</button>
          </div>
        </div>
        <div class="stddev-control-row stddev-control-row--toggle">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <button type="button" data-std-control="toggle" data-group="sp500_graph3">S&amp;P 500 Std On/Off</button>
        </div>
        <div class="stddev-control-row stddev-control-row--line-count">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="lineCount" data-group="sp500_graph3" data-value="4">4 lines</button>
            <button type="button" data-std-control="lineCount" data-group="sp500_graph3" data-value="8">8 lines</button>
          </div>
        </div>
      </div>
      <div class="stddev-control-metric">
        <div class="stddev-control-row">
          <span class="stddev-control-label">CAPE mode</span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="mode" data-group="cape_graph3" data-value="raw">Raw</button>
            <button type="button" data-std-control="mode" data-group="cape_graph3" data-value="log">Log</button>
          </div>
        </div>
        <div class="stddev-control-row stddev-control-row--toggle">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <button type="button" data-std-control="toggle" data-group="cape_graph3">CAPE Std On/Off</button>
        </div>
        <div class="stddev-control-row stddev-control-row--line-count">
          <span class="stddev-control-label stddev-control-label--spacer"></span>
          <div class="stddev-control-choice-group">
            <button type="button" data-std-control="lineCount" data-group="cape_graph3" data-value="4">4 lines</button>
            <button type="button" data-std-control="lineCount" data-group="cape_graph3" data-value="8">8 lines</button>
          </div>
        </div>
      </div>
    </section>
  </div>
</div>
<style>
  .stddev-control-panel {{
    font-family: Arial, sans-serif;
    border: 1px solid #d0d7de;
    border-radius: 8px;
    padding: 14px 16px;
    margin: 10px 18px 0 18px;
    background: #f8fafc;
  }}
  .stddev-control-title {{ font-weight: 700; margin-bottom: 10px; color: #111827; }}
  .stddev-control-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 12px; }}
  .stddev-control-group {{
    border: 1px solid #dbe3ee;
    border-radius: 10px;
    background: #ffffff;
    padding: 10px 12px;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.05);
  }}
  .stddev-control-group-title {{ font-weight: 700; color: #1f2937; margin-bottom: 10px; }}
  .stddev-control-metric + .stddev-control-metric {{ margin-top: 10px; }}
  .stddev-control-row {{ display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }}
  .stddev-control-row + .stddev-control-row {{ margin-top: 6px; }}
  .stddev-control-row--toggle, .stddev-control-row--line-count {{ padding-left: 0; }}
  .stddev-control-label {{ min-width: 110px; font-weight: 600; color: #374151; }}
  .stddev-control-label--spacer {{ visibility: hidden; }}
  .stddev-control-choice-group {{ display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }}
  .stddev-control-panel button {{
    border: 1px solid #9ca3af;
    border-radius: 999px;
    padding: 5px 12px;
    background: #ffffff;
    color: #111827;
    cursor: pointer;
  }}
  .stddev-control-panel button.active {{ background: #1f77b4; border-color: #1f77b4; color: #ffffff; }}
  .stddev-control-panel button.toggle-active {{ background: #16a34a; border-color: #16a34a; color: #ffffff; }}
  .stddev-control-panel button.toggle-inactive {{ background: #e5e7eb; border-color: #9ca3af; color: #374151; }}
</style>
<script>
(function() {{
  const plotId = {json.dumps(plot_id)};
  const state = {state_json};
  const EPS = 1e-9;

  function selectedMultipliers(lineCount) {{
    return Number(lineCount) === 8 ? [0.5, 1, 1.5, 2] : [1, 2];
  }}

  function lineCountForGroup(group) {{
    return Number((state.lineCounts || {{}})[group] || 4);
  }}

  function traceShouldShow(trace) {{
    const meta = trace.meta || {{}};
    if (!meta.stdControl) return trace.visible === undefined ? true : trace.visible;
    const group = meta.stdGroup;
    const mode = meta.stdMode;
    const multiplier = Number(meta.stdMultiplier);
    const allowed = selectedMultipliers(lineCountForGroup(group)).some(v => Math.abs(v - multiplier) < EPS);
    return Boolean(state.show[group] && state.modes[group] === mode && allowed);
  }}

  function annotationShouldShow(annotation) {{
    if (!annotation || !annotation.name || !annotation.name.startsWith('stdLabel:')) {{
      return annotation.visible === undefined ? true : annotation.visible;
    }}
    const parts = annotation.name.split(':');
    const group = parts[1];
    const mode = parts[2];
    const multiplier = Number(parts[3]);
    const allowed = selectedMultipliers(lineCountForGroup(group)).some(v => Math.abs(v - multiplier) < EPS);
    return Boolean(state.show[group] && state.modes[group] === mode && allowed);
  }}

  function updateButtonStates() {{
    document.querySelectorAll('#stddev-control-panel button').forEach(button => {{
      const control = button.dataset.stdControl;
      const group = button.dataset.group;
      const value = button.dataset.value;
      button.classList.remove('active', 'toggle-active', 'toggle-inactive');
      if (control === 'mode' && state.modes[group] === value) button.classList.add('active');
      if (control === 'lineCount' && String(lineCountForGroup(group)) === String(value)) button.classList.add('active');
      if (control === 'toggle') button.classList.add(state.show[group] ? 'toggle-active' : 'toggle-inactive');
    }});
  }}

  function applyStdControls() {{
    const plot = document.getElementById(plotId);
    if (!plot || !plot.data) return;
    const traceIndexes = [];
    const visibility = [];
    plot.data.forEach((trace, index) => {{
      if (trace.meta && trace.meta.stdControl) {{
        traceIndexes.push(index);
        visibility.push(traceShouldShow(trace));
      }}
    }});
    if (traceIndexes.length) {{
      Plotly.restyle(plot, {{ visible: visibility }}, traceIndexes);
    }}
    const currentAnnotations = (plot.layout.annotations || []).map(annotation => {{
      const copy = Object.assign({{}}, annotation);
      copy.visible = annotationShouldShow(copy);
      return copy;
    }});
    Plotly.relayout(plot, {{ annotations: currentAnnotations }});
    updateButtonStates();
  }}

  function attachHandlers() {{
    document.querySelectorAll('#stddev-control-panel button').forEach(button => {{
      button.addEventListener('click', () => {{
        const control = button.dataset.stdControl;
        const group = button.dataset.group;
        const value = button.dataset.value;
        if (control === 'mode') state.modes[group] = value;
        if (control === 'toggle') state.show[group] = !state.show[group];
        if (control === 'lineCount') state.lineCounts[group] = Number(value);
        applyStdControls();
      }});
    }});
  }}

  function waitForPlot() {{
    const plot = document.getElementById(plotId);
    if (plot && plot.data) {{
      attachHandlers();
      applyStdControls();
    }} else {{
      setTimeout(waitForPlot, 50);
    }}
  }}

  if (document.readyState === 'loading') {{
    document.addEventListener('DOMContentLoaded', waitForPlot);
  }} else {{
    waitForPlot();
  }}
}})();
</script>
"""


def save_dashboard_html(
    output_path: Path,
    figure: go.Figure,
    *,
    stddev_modes: dict[str, str],
    show_std_lines: dict[str, bool],
    stddev_line_counts: dict[str, int],
) -> None:
    """Save Plotly HTML with per-graph custom controls for Std lines."""
    plot_id = "buffett-dashboard-plot"
    initial_state = {
        "modes": {group: str(value) for group, value in stddev_modes.items()},
        "show": {group: bool(value) for group, value in show_std_lines.items()},
        "lineCounts": {group: int(value) for group, value in stddev_line_counts.items()},
    }
    controls = build_stddev_control_panel_html(initial_state, plot_id)
    html = figure.to_html(full_html=True, include_plotlyjs="cdn", div_id=plot_id)
    marker = f'<div id="{plot_id}"'
    if marker in html:
        html = html.replace(marker, controls + "\n" + marker, 1)
    else:
        html = html.replace("<body>", "<body>\n" + controls, 1)
    output_path.write_text(html, encoding="utf-8")


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

def add_stddev_cli_group_args(
    parser: argparse.ArgumentParser,
    *,
    arg_prefix: str,
    label: str,
    default_mode: str,
    default_visible: bool,
    default_line_count: int = 4,
) -> None:
    """Register one complete Std-dev CLI control set for a single graph series.

    ``arg_prefix`` is the stable command-line stem (for example ``graph1-cape``).
    The function adds three arguments that all apply only to that one chart
    series: initial mode, initial visibility, and initial line-count selector.
    Centralizing the argument creation keeps the six graph-series definitions
    consistent and makes it much harder for shared/coupled CLI behavior to creep
    back into the script in later edits.
    """
    parser.add_argument(
        f"--{arg_prefix}-stddev-mode",
        choices=["raw", "log"],
        default=default_mode,
        help=f"Initial std-dev mode for {label} only: raw mean +/- std or log-residual bands (default: {default_mode}).",
    )
    parser.add_argument(
        f"--{arg_prefix}-std-lines",
        action=argparse.BooleanOptionalAction,
        default=default_visible,
        help=(
            f"Initial Std line visibility for {label} only "
            f"(default: {'on' if default_visible else 'off'})."
        ),
    )
    parser.add_argument(
        f"--{arg_prefix}-stddev-lines",
        type=int,
        choices=[4, 8],
        default=default_line_count,
        help=f"Initial Std line count for {label} only: 4 = +/-1 and +/-2; 8 also includes +/-0.5 and +/-1.5 (default: {default_line_count}).",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Buffett/CAPE/Berkshire dashboard and Excel workbook.")
    parser.add_argument("--start", default=DEFAULT_START, help=f"Start date for displayed data (default: {DEFAULT_START})")
    parser.add_argument("--output-dir", default=".", help="Directory where the HTML and Excel files will be saved")
    parser.add_argument("--html-name", default=DEFAULT_OUTPUT_HTML, help=f"HTML file name (default: {DEFAULT_OUTPUT_HTML})")
    parser.add_argument("--excel-name", default=DEFAULT_OUTPUT_XLSX, help=f"Excel file name (default: {DEFAULT_OUTPUT_XLSX})")

    add_stddev_cli_group_args(parser, arg_prefix="graph1-cape", label="Graph 1 CAPE", default_mode="raw", default_visible=True)
    add_stddev_cli_group_args(parser, arg_prefix="graph1-buffett", label="Graph 1 Buffett Indicator", default_mode="log", default_visible=True)
    add_stddev_cli_group_args(parser, arg_prefix="graph2-sp500", label="Graph 2 S&P 500", default_mode="raw", default_visible=False)
    add_stddev_cli_group_args(parser, arg_prefix="graph2-buffett", label="Graph 2 Buffett Indicator", default_mode="log", default_visible=True)
    add_stddev_cli_group_args(parser, arg_prefix="graph3-sp500", label="Graph 3 S&P 500", default_mode="raw", default_visible=False)
    add_stddev_cli_group_args(parser, arg_prefix="graph3-cape", label="Graph 3 CAPE", default_mode="raw", default_visible=True)

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

    latest_sp500_month_end = latest_complete_month_end()
    log_progress(f"Latest completed S&P 500 month-end cutoff: {latest_sp500_month_end.date()}", started_at)

    log_progress("Downloading latest S&P 500 history from Yahoo Finance", started_at)
    sp500_yahoo = fetch_yahoo_history("^GSPC", start)
    if sp500_yahoo.empty:
        raise RuntimeError("Yahoo Finance returned no usable S&P 500 history for ^GSPC.")

    # Normalize Yahoo output into the dashboard/export schema right away so all
    # later code paths can rely on a single stable column layout.
    sp500_yahoo = prepare_sp500_yahoo_frame_for_dashboard(sp500_yahoo)
    canonical_sp500_cutoff = latest_sp500_month_end.normalize()
    sp500_yahoo = sp500_yahoo.loc[sp500_yahoo["date"].dt.normalize() <= canonical_sp500_cutoff].copy()
    sp500_yahoo = prepare_sp500_yahoo_frame_for_dashboard(sp500_yahoo)
    if sp500_yahoo.empty:
        raise RuntimeError("Yahoo Finance S&P 500 history did not contain the latest completed month-end after cutoff filtering.")
    latest_downloaded_sp500_date = sp500_yahoo["date"].max().normalize()
    if latest_downloaded_sp500_date != canonical_sp500_cutoff:
        raise RuntimeError(
            "Yahoo Finance S&P 500 history is missing the latest completed month-end "
            f"({canonical_sp500_cutoff.date()}); latest available row was {latest_downloaded_sp500_date.date()}."
        )
    log_progress(f"S&P 500 history ready through {latest_downloaded_sp500_date.date()} with {len(sp500_yahoo):,} rows", started_at)

    log_progress("Downloading Shiller workbook", started_at)
    shiller = build_shiller_series(session, start)
    log_progress(f"Shiller dataset ready with {len(shiller):,} rows", started_at)

    log_progress("Downloading Berkshire cash/assets history", started_at)
    brk_bundle = fetch_berkshire_history(session)
    brk = brk_bundle.merged.copy()
    log_progress(f"Berkshire dataset ready with {len(brk):,} rows", started_at)

    log_progress("Combining monthly datasets for Excel export", started_at)
    combined = build_combined_monthly_sheet(buffett, shiller, brk, sp500_yahoo)

    stddev_modes = {
        "cape_graph1": args.graph1_cape_stddev_mode,
        "buffett_graph1": args.graph1_buffett_stddev_mode,
        "sp500_graph2": args.graph2_sp500_stddev_mode,
        "buffett_graph2": args.graph2_buffett_stddev_mode,
        "sp500_graph3": args.graph3_sp500_stddev_mode,
        "cape_graph3": args.graph3_cape_stddev_mode,
    }
    show_std_lines = {
        "cape_graph1": bool(args.graph1_cape_std_lines),
        "buffett_graph1": bool(args.graph1_buffett_std_lines),
        "sp500_graph2": bool(args.graph2_sp500_std_lines),
        "buffett_graph2": bool(args.graph2_buffett_std_lines),
        "sp500_graph3": bool(args.graph3_sp500_std_lines),
        "cape_graph3": bool(args.graph3_cape_std_lines),
    }
    stddev_line_counts = {
        "cape_graph1": int(args.graph1_cape_stddev_lines),
        "buffett_graph1": int(args.graph1_buffett_stddev_lines),
        "sp500_graph2": int(args.graph2_sp500_stddev_lines),
        "buffett_graph2": int(args.graph2_buffett_stddev_lines),
        "sp500_graph3": int(args.graph3_sp500_stddev_lines),
        "cape_graph3": int(args.graph3_cape_stddev_lines),
    }

    log_progress("Building Plotly dashboard figure", started_at)
    figure = build_dashboard_figure(
        buffett=buffett,
        shiller=shiller,
        brk=brk,
        sp500_yahoo=sp500_yahoo,
        stddev_modes=stddev_modes,
        show_std_lines=show_std_lines,
        stddev_line_counts=stddev_line_counts,
    )

    html_path = output_dir / args.html_name
    excel_path = output_dir / args.excel_name

    log_progress("Saving HTML dashboard", started_at)
    save_dashboard_html(
        html_path,
        figure,
        stddev_modes=stddev_modes,
        show_std_lines=show_std_lines,
        stddev_line_counts=stddev_line_counts,
    )

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
