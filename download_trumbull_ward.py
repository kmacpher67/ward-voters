#!/usr/bin/env python3
"""Download the latest Ohio SOS county voter file and filter it by Warren ward.

By default this fetches the Trumbull County voter file from the Ohio SOS
download page, saves a copy of the raw .txt file, and writes filtered outputs
for Warren City Ward 4.

Examples:
  python3 download_trumbull_ward.py
  python3 download_trumbull_ward.py --ward 3
  python3 download_trumbull_ward.py --ward 4 --city "WARREN CITY"
"""

from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


HOME_URL = "https://www6.ohiosos.gov/ords/f?p=VOTERFTP:HOME::::::"
DOWNLOAD_FALLBACK_URL = (
    "https://www6.ohiosos.gov/ords/f?p=VOTERFTP:DOWNLOAD::FILE:NO:2:P2_PRODUCT_NUMBER:{product_number}"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download the latest Ohio SOS county voter file and filter it by Warren ward."
    )
    parser.add_argument("--county", default="TRUMBULL", help="County name to download (default: TRUMBULL)")
    parser.add_argument("--ward", default="4", help="Ward number to keep (default: 4)")
    parser.add_argument(
        "--city",
        default="WARREN CITY",
        help='Optional city filter (default: "WARREN CITY"). Use empty string to disable.',
    )
    parser.add_argument(
        "--product-number",
        default="78",
        help="Ohio SOS product number used as a fallback direct download URL (default: 78 for Trumbull).",
    )
    parser.add_argument(
        "--download-dir",
        default="downloads",
        help="Directory for the raw downloaded .txt file (default: downloads)",
    )
    parser.add_argument(
        "--input",
        default="",
        help="Optional local raw .txt file to filter instead of downloading a fresh copy.",
    )
    parser.add_argument(
        "--output-dir",
        default="outputs",
        help="Directory for filtered CSV/XLSX outputs (default: outputs)",
    )
    parser.add_argument(
        "--headless",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Run Chrome headless (default: true). Use --no-headless to watch it.",
    )
    return parser.parse_args()


def build_driver(download_dir: Path, headless: bool) -> webdriver.Chrome:
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    prefs = {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options,
    )
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": str(download_dir)},
    )
    return driver


def wait_for_download(download_dir: Path, start_time: float, timeout: int = 120) -> Path:
    deadline = time.time() + timeout
    while time.time() < deadline:
        in_progress = list(download_dir.glob("*.crdownload"))
        candidates = [
            path
            for path in download_dir.glob("*.txt")
            if path.stat().st_mtime >= start_time - 1
        ]
        if candidates and not in_progress:
            return max(candidates, key=lambda path: path.stat().st_mtime)
        time.sleep(1)
    raise TimeoutError(f"No completed .txt download appeared in {download_dir} within {timeout} seconds.")


def download_latest_county_file(county: str, product_number: str, download_dir: Path, headless: bool) -> Path:
    driver = build_driver(download_dir, headless=headless)
    start_time = time.time()
    try:
        driver.get(HOME_URL)
        wait = WebDriverWait(driver, 30)
        try:
            county_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, county.upper())))
            county_link.click()
        except Exception:
            # Fallback in case the county link text or page layout changes.
            driver.get(DOWNLOAD_FALLBACK_URL.format(product_number=product_number))
        raw_download = wait_for_download(download_dir, start_time=start_time)
        target_name = f"{county.upper()}-latest.txt"
        target_path = download_dir / target_name
        if target_path.exists():
            target_path.unlink()
        raw_download.replace(target_path)
        return target_path
    finally:
        driver.quit()


def latest_local_txt(download_dir: Path, county: str) -> Path | None:
    candidates = [p for p in download_dir.glob("*.txt") if county.lower() in p.name.lower()]
    if not candidates:
        candidates = list(download_dir.glob("*.txt"))
    if not candidates:
        return None
    return max(candidates, key=lambda path: path.stat().st_mtime)


def filter_county_file(raw_path: Path, county: str, ward: str, city: str, output_dir: Path) -> tuple[Path, Path, int]:
    df = pd.read_csv(raw_path, dtype=str, keep_default_na=False)
    df.columns = df.columns.str.strip().str.upper()

    filtered = df.copy()
    if "WARD" in filtered.columns:
        filtered = filtered[filtered["WARD"].str.contains(f"WARREN-WARD {ward}", case=False, na=False)]
    else:
        raise KeyError("The downloaded file does not contain a WARD column.")

    if city and "CITY" in filtered.columns:
        filtered = filtered[filtered["CITY"].str.contains(city, case=False, na=False)]

    sort_columns = [col for col in ["PRECINCT_NAME", "LAST_NAME", "FIRST_NAME"] if col in filtered.columns]
    if sort_columns:
        filtered = filtered.sort_values(by=sort_columns, kind="stable")

    output_dir.mkdir(parents=True, exist_ok=True)
    today_str = datetime.today().strftime("%Y-%m-%d")
    city_slug = (city or "ALL_CITIES").replace(" ", "_").upper()
    csv_path = output_dir / f"{county.upper()}_{city_slug}_WARD{ward}_{today_str}.csv"
    xlsx_path = output_dir / f"{county.upper()}_{city_slug}_WARD{ward}_{today_str}.xlsx"

    filtered.to_csv(csv_path, index=False)
    filtered.to_excel(xlsx_path, index=False, engine="openpyxl")
    return csv_path, xlsx_path, len(filtered)


def main() -> int:
    args = parse_args()
    download_dir = Path(args.download_dir).resolve()
    output_dir = Path(args.output_dir).resolve()
    download_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    raw_path: Path | None = None
    if args.input:
        raw_path = Path(args.input).expanduser().resolve()
        if not raw_path.exists():
            print(f"Error: input file does not exist: {raw_path}", file=sys.stderr)
            return 1
    else:
        try:
            raw_path = download_latest_county_file(
                county=args.county,
                product_number=args.product_number,
                download_dir=download_dir,
                headless=args.headless,
            )
        except Exception as exc:
            fallback = latest_local_txt(download_dir, args.county)
            if fallback is None:
                print(f"Error: {exc}", file=sys.stderr)
                return 1
            print(
                f"Download blocked or failed ({exc}). Using the latest local file instead: {fallback}",
                file=sys.stderr,
            )
            raw_path = fallback

    try:
        csv_path, xlsx_path, row_count = filter_county_file(
            raw_path=raw_path,
            county=args.county,
            ward=args.ward,
            city=args.city,
            output_dir=output_dir,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Raw download: {raw_path}")
    print(f"Filtered rows: {row_count}")
    print(f"CSV output: {csv_path}")
    print(f"XLSX output: {xlsx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
