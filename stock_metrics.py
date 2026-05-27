import requests
import pandas as pd
from bs4 import BeautifulSoup
import re
import sys
import time
from collections import defaultdict
import argparse
import json
import os
from typing import Union
from datetime import datetime, timedelta
from pathlib import Path

# Selenium imports for dynamic content rendering
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}


EXCEL_DIR = Path("excel")

CACHE_DIR = Path(".cache")
CACHE_EXPIRY_HOURS = 24


def get_cache_path(ticker: str) -> Path:
    """Return the cache file path for a given ticker."""
    CACHE_DIR.mkdir(exist_ok=True)
    return CACHE_DIR / f"{ticker.lower()}_metrics.json"


def save_cache(ticker: str, df: pd.DataFrame) -> None:
    """Save DataFrame to cache file."""
    cache_path = get_cache_path(ticker)
    df.to_json(cache_path, orient="records", force_ascii=False)
    print(f"Cache saved to {cache_path}")


def load_cache(ticker: str) -> pd.DataFrame | None:
    """Load DataFrame from cache if exists and not expired."""
    cache_path = get_cache_path(ticker)
    if not cache_path.exists():
        return None
    try:
        cache_time = datetime.fromtimestamp(cache_path.stat().st_mtime)
        if datetime.now() - cache_time > timedelta(hours=CACHE_EXPIRY_HOURS):
            print(f"Cache expired ({(datetime.now() - cache_time).total_seconds() / 3600:.1f}h ago), will refresh")
            return None
        df = pd.read_json(cache_path)
        print(f"Loaded cached data from {cache_path}")
        return df
    except Exception as e:
        print(f"Failed to load cache: {e}")
        return None


def display_metric(df: pd.DataFrame) -> None:
    """Print the metrics DataFrame in a simple table format."""
    print("\nCollected data:")
    print(df.to_string(index=False))


def _extract_headers(header_cells):
    """Extract and convert year strings from table header cells.

    Skips the first cell (metric name), converts Buddhist years (>2500)
    to Gregorian by subtracting 543.

    Args:
        header_cells: List of BeautifulSoup th/td elements from table header

    Returns:
        List of year strings
    """
    years = []
    for cell in header_cells[1:]:  # skip first cell which is metric name
        txt = cell.get_text(strip=True)
        if txt.isdigit():
            year = int(txt)
            # If Buddhist year (> 2500) convert to Gregorian.
            if year > 2500:
                year -= 543
            years.append(str(year))
        else:
            years.append(txt)  # fallback, keep as‑is
    return years


def _build_metric_rows(rows, format=0):
    """Build dictionary mapping metric names to lists of cell values from table rows.

    Args:
        rows: List of BeautifulSoup tr elements (table rows)

    Returns:
        Dictionary mapping metric names to lists of cell values
    """
    metric_rows = {}
    metric_name_list = []
    for row in rows[1:]:  # skip header row
        cells = row.find_all(["th", "td"])
        if not cells:
            continue
        metric_name = cells[0].get_text(strip=True)
        if format == 1:
            metric_name_list.append(metric_name) 
        metric_rows[metric_name] = [c.get_text(strip=True) for c in cells[1:]]
    return metric_rows if format==0 else (metric_rows, metric_name_list)


def _create_metric_dataframe(years, metric_rows, metric_mapping):
    """Create a DataFrame for two metrics given years, metric rows, and mapping.

    Args:
        years: List of year strings
        metric_rows: Dictionary mapping metric names to lists of cell values
        metric_mapping: List of tuples [(thai_name, col_name), ...] for metrics to extract

    Returns:
        DataFrame with Year and metric columns, or None if empty
    """
    import pandas as pd
    import re

    data = []
    for metric, col_name in metric_mapping:
        if metric not in metric_rows:
            continue
        values = metric_rows[metric]
        for yr, str_val in zip(years, values):
            # Slice to get 2 digit float as requested: from start to first dot + 2
            float_val = __convert_float_str_to_float(str_val)
            data.append({"Year": yr, col_name: float_val})

    df = pd.DataFrame(data)
    
    if df.empty:
        return None
    df = df.groupby('Year', as_index=False).first()

    # The pivot above creates a MultiIndex column; flatten it.
    df.columns = df.columns.get_level_values(0)
    df = df.reset_index()
    return df[["Year"] + [col_name for _, col_name in metric_mapping]]

def __convert_float_str_to_float(val):
    """
    Convert float string value to float

    Args:
        val: float strings  

    Returns:
        float value or None
    """
    if '.' in val:
            dot_index = val.index('.')
            val = val[0:dot_index + 2 + 1 ].strip().replace(',', '')
    try:
        return float(val) if val else None
    except ValueError:
        return None

def _parse_dividend_table(soup):
    """Parse dividend table using helper functions for consistency.

    Args:
        soup: BeautifulSoup object of rendered page

    Returns:
        DataFrame with Year and Dividend columns, or None if not found
    """
    import pandas as pd
    import re

    table = soup.find("table")
    if not table:
        return None
    rows = table.find_all("tr")
    if not rows:
        return None
    header_cells = rows[0].find_all(["th", "td"])
    # Detect year and dividend column positions.
    year_idx = div_idx = None
    for i, cell in enumerate(header_cells):
        txt = cell.get_text(strip=True)
        if re.search(r"\d{4}", txt):
            year_idx = i
        if "ปันผล" in txt or "Dividend" in txt:
            div_idx = i
    if year_idx is None or div_idx is None:
        raise ValueError("Year or dividend column not found in dividend table")

    # Extract years using similar logic to _extract_headers but adjusted for dividend table format
    years = []
    for cell in header_cells[year_idx:]:
        txt = cell.get_text(strip=True)
        if txt.isdigit():
            year = int(txt)
            if year > 2500:
                year -= 543
            years.append(str(year))
        elif re.search(r"\d{4}", txt):  # Already Gregorian
            years.append(txt)
        # Skip non-year cells

    data = []
    for row in rows[1:]:
        cells = row.find_all(["th", "td"])
        if len(cells) <= max(year_idx, div_idx):
            continue
        year = cells[year_idx].get_text(strip=True)
        div = cells[div_idx].get_text(strip=True)
        # Convert Buddhist year to Gregorian if needed.
        if year.isdigit() and int(year) > 2500:
            year = str(int(year) - 543)
        temp_num = re.sub(r"[^0-9.]", "", div)
        num = __convert_float_str_to_float(temp_num)
        data.append({"Year": year, "Dividend": num})
    return pd.DataFrame(data) if data else None

def _aggregate_num_with_year(data_list):
    """
    Aggregate dictionary with 2 keys [year, num] list 
    to sum data with the same year to only 1 row
    and return to Dictionary list with the same keys

    Args:
        data_list: Dictionary 2 keys (year, num) list

    Return  
        Dictionary 2 keys (year, num) list
    """
    aggregated = defaultdict(int)

    # 1. ขั้นตอนการรวมกลุ่ม (Aggregation)
    for item in data_list:
        year_val, num_val = item.values()
    
        # ใช้ str() เพื่อบังคับให้ Key ของตารางสรุปผลเป็น String เสมอ
        aggregated[str(year_val)] += num_val

    # 2. ขั้นตอนการสร้างผลลัพธ์กลับเป็น List โดยใช้คีย์เดิม
    k0, k1 = data_list[0].keys()

    # ผลลัพธ์ที่ได้ คีย์ 'year' จะจับคู่กับข้อมูลที่เป็น String แน่นอน
    return [{k0: y, k1: n} for y, n in aggregated.items()]




def get_driver():
    """Create a headless Chrome driver using webdriver‑manager.
    The driver is configured with the custom User‑Agent header via Chrome
    options. This function returns a ready‑to‑use driver instance.
    """
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument(f"--user-agent={HEADERS['User-Agent']}")
    # Suppress unnecessary logs
    chrome_options.add_argument("--log-level=3")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def fetch_soup(url):
    """Load *url* in a headless browser, wait for the first <table> to appear,
    then return a BeautifulSoup object of the fully rendered page source.
    """
    driver = get_driver()
    try:
        driver.get(url)
        # Wait up to 15 seconds for a table element to be present.
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        # Give a short pause for any asynchronous scripts to fill the table.
        time.sleep(1)
        html = driver.page_source
        return BeautifulSoup(html, "html.parser")
    finally:
        driver.quit()

def parse_pl_table(soup):
    # Find the first table – the page contains only one financial table.
    table = soup.find("table")
    if not table:
        return None

    rows = table.find_all("tr")
    if len(rows) < 3:
        return None

    # Extract years and metric rows using helpers
    header_cells = rows[0].find_all(["th", "td"])
    years = _extract_headers(header_cells)
    metric_rows = _build_metric_rows(rows)

    # Build DataFrame from the two metrics we need using helper
    df = _create_metric_dataframe(years, metric_rows, [("กำไรต่อหุ้น", "EPS"), ("P/E", "P/E")])
    return df

def parse_bl_table(soup):
    table = soup.find("table")
    if not table:
        return None
    rows = table.find_all("tr")
    if len(rows) < 3:
        return None

    # Extract years and metric rows using helpers
    header_cells = rows[0].find_all(["th", "td"])
    years = _extract_headers(header_cells)
    metric_rows = _build_metric_rows(rows)

    # Look for Equity – the Thai term is "ทุนสุทธิ"; English fallback "Equity".
    df = _create_metric_dataframe(years, metric_rows, [("รวมส่วนของเจ้าของ", "Equity"), ("อัตราผลตอบแทนผู้ถือหุ้น", "ROE")])
    return df

def parse_dividend_table(soup):
    table = soup.find("table")
    if not table:
        return None
    rows = table.find_all("tr")
    if not rows:
        return None
    header_cells = rows[0].find_all(["th", "td"])
    # Detect year and dividend column positions.
    year_idx = div_idx = None
    for i, cell in enumerate(header_cells):
        txt = cell.get_text(strip=True)
        if "ปีที่" in txt or "year" in txt:
            year_idx = i
        if "จำนวน" in txt or "Dividend" in txt:
            div_idx = i
    if year_idx is None or div_idx is None:
        raise ValueError("Year or dividend column not found in dividend table")

    data = []
    for row in rows[1:]:
        cells = row.find_all(["th", "td"])
        if len(cells) <= max(year_idx, div_idx):
            continue
        if len(cells) == len(header_cells):
            year = cells[year_idx].get_text(strip=True)
        else:
            if len(data) <= 0: continue
            year = data[-1]["Year"] 
        div = cells[div_idx-(len(header_cells)-len(cells))].get_text(strip=True)
        # Convert Buddhist year to Gregorian if needed.
        if year.isdigit() and int(year) > 2500:
            year = str(int(year) - 543)
        temp_num = re.sub(r"[^0-9.]", "", div)
        num = __convert_float_str_to_float(temp_num)
        data.append({"Year": year, "Dividend": num})
    
    result = _aggregate_num_with_year(data) 

    return pd.DataFrame(result) if data else None

def get_eps_pe(ticker):
    url = f"https://aio.panphol.com/stock/{ticker}/pl"
    soup = fetch_soup(url)
    df = parse_pl_table(soup)
    if df is None:
        raise ValueError(f"No PL data found for {ticker}")
    return df

def get_equity(ticker):
    url = f"https://aio.panphol.com/stock/{ticker}/bl"
    soup = fetch_soup(url)
    df = parse_bl_table(soup)
    if df is None:
        raise ValueError(f"No BL data found for {ticker}")
    return df

def get_dividend(ticker):
    url = f"https://aio.panphol.com/stock/{ticker}/dividend"
    soup = fetch_soup(url)
    df = parse_dividend_table(soup)
    if df is None:
        raise ValueError(f"No dividend data found for {ticker}")
    return df

def get_shares_outstanding(ticker):
    url = f"https://www.set.or.th/th/market/product/stock/quote/{ticker}/factsheet"
    soup = fetch_soup(url)
    
    # For number of share are at table index 7
    all_table = soup.find_all("table")
    table = all_table[7] if len(all_table) > 7 else None
    if not table:
        return None
    rows = table.find_all("tr")
    if len(rows) < 3:
        return None
    
    # Build headers
    header_cells = rows[0].find_all(["th", "td"])
    headers = _extract_headers(header_cells)

    # Extract years and metric rows using helpers
    metric_rows, metric_names = _build_metric_rows(rows, 1)

    float_str = metric_rows[metric_names[0]][0]
    
    return __convert_float_str_to_float(float_str) if float_str else None


def _generate_excel_path(stock_name: str) -> str:
    fname = f"{stock_name}_project.xlsx"
    return os.path.join(EXCEL_DIR, fname)

def save_excel_operation(path: str, df: pd.DataFrame) -> None | str:
    """Ask user to save dataframe and save"""
    save = input("\nSave to Excel? (y/n): ").strip().lower()

    if not EXCEL_DIR.is_dir():
        EXCEL_DIR.mkdir(parents=True, exist_ok=True)
    
    if save == "y":
        df.to_excel(path, index=False)
        print(f"Data saved to {path}")
    return path
    

def open_excel(path: Union[str, Path]) -> None:
    """Open excel file by default program"""

    file_path = None
    if path == None: file_path = Path("")
    else: file_path = Path(path)

    if file_path.is_file():
        os.startfile(path)
    else:
        print(f"File is not exist: '{path}' ")

def open_excel_operation(path: Union[str, Path]) -> None:
    """Ask to open excel and operate"""
    is_desire_open = input("\nOpen the excel? (y/n): ").strip().lower()
    if is_desire_open:
        open_excel(path)

def excel_operation(ticker: str, df: pd.DataFrame) -> None:
    save_path = _generate_excel_path(ticker)
    save_excel_operation(save_path, df)
    open_excel_operation(save_path)

def assemble(ticker):
    eps_pe = get_eps_pe(ticker)
    equity = get_equity(ticker)
    dividend = get_dividend(ticker)
    # Merge on Year (inner join keeps only years present in all tables).
    df = eps_pe.merge(equity, on="Year", how="inner")
    
    df = df.merge(dividend, on="Year", how="inner")
    shares = get_shares_outstanding(ticker)

    df['Equity/share'] = (df['Equity'] / shares).round(2) 

    df = df[["Year", "P/E", "ROE", "Equity", "Equity/share", "EPS", "Dividend"]]

    print(f"Avaliable shares: {shares}")

    filter_df = df.sort_values(by="Year", ascending=False).head(10).reset_index(drop=True)
    sorted_df = filter_df.sort_values(by="Year", ascending=True)
    # Ensure Year is integer when possible.
    try:
        sorted_df["Year"] = sorted_df["Year"].astype(int)
    except Exception:
        pass
    return sorted_df

def process_ticker(ticker: str, force_refresh: bool = False) -> None:
    """Core workflow: check cache, optionally refresh, and display results."""
    cached_df = load_cache(ticker) if not force_refresh else None
    if cached_df is not None and not force_refresh:
        display_metric(cached_df)

        path = _generate_excel_path(ticker)
        if not os.path.exists(path):
            excel_operation(ticker, cached_df)
        else:
            open_excel(path)
        return
        
    print(f"Fetching data for {ticker} …")
    df = assemble(ticker)
    display_metric(df)
    save_cache(ticker, df)
    excel_operation(ticker, df)

def main() -> None:
    parser = argparse.ArgumentParser(description="Fetch and display stock metrics with optional caching.")
    parser.add_argument("ticker", nargs="?", help="Stock ticker symbol (e.g., AOT).")
    parser.add_argument("-r", "--refresh", action="store_true", help="Force refresh data.")
    args = parser.parse_args()

    ticker = args.ticker or input("Enter stock ticker (e.g., AOT): ").strip()
    if not ticker:
        print("No ticker entered – exiting.")
        sys.exit(1)

    process_ticker(ticker, force_refresh=args.refresh)

if __name__ == "__main__":
    main()