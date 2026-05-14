#!/usr/bin/env python3
"""
TerraBrasilis/INPE public data downloader
-----------------------------------------

This script lists, downloads, resumes, and validates public ZIP files from the
official TerraBrasilis/INPE download page:

https://terrabrasilis.dpi.inpe.br/en/download-files/

The repository provides code only. It does not redistribute TerraBrasilis/INPE
datasets. Downloaded files are stored locally on the user's machine. Users are
responsible for checking the current TerraBrasilis/INPE terms of use, citation
requirements, and data-use conditions before publishing or sharing derived
products.
"""

from __future__ import annotations

import argparse
import csv
import json
import time
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import TypedDict
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from tqdm import tqdm
from urllib3.util.retry import Retry

BASE_URL = "https://terrabrasilis.dpi.inpe.br/en/download-files/"
HTTP_TIMEOUT = 30
DOWNLOAD_TIMEOUT = 600
CHUNK_SIZE = 8 * 1024 * 1024
SEP = "=" * 70

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

BIOME_MAP: dict[str, str] = {
    "amz-prodes": "Amazon_Biome",
    "amz-aux": "Amazon_Biome",
    "amz-terraclass": "Amazon_Biome",
    "legal-amz-prodes": "Legal_Amazon",
    "legal-amz-aux": "Legal_Amazon",
    "caatinga-prodes": "Caatinga",
    "caatinga-aux": "Caatinga",
    "cerrado-prodes": "Cerrado",
    "cerrado-aux": "Cerrado",
    "cerrado-vegetation": "Cerrado",
    "mata-atlantica-prodes": "Mata_Atlantica",
    "mata-atlantica-aux": "Mata_Atlantica",
    "pampa-prodes": "Pampa",
    "pampa-aux": "Pampa",
    "pantanal-prodes": "Pantanal",
    "pantanal-aux": "Pantanal",
    "brasil-prodes": "Brazil",
    "vs": "Vegetacao_Secundaria",
}


class ZipEntry(TypedDict):
    url: str
    filename: str
    filename_local: str
    text: str
    biome: str
    category: str


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(HEADERS)
    retry = Retry(
        total=5,
        backoff_factor=2.0,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
        respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def infer_from_url(url: str) -> tuple[str, str]:
    parts = urlparse(url).path.strip("/").split("/")
    try:
        slug = parts[2]
        file_type = parts[3].capitalize()
    except IndexError:
        return "N_A", "N_A"
    return BIOME_MAP.get(slug, slug.replace("-", "_")), file_type


def extract_zip_links(html: str, base_url: str) -> list[ZipEntry]:
    soup = BeautifulSoup(html, "lxml")
    results: list[ZipEntry] = []
    seen: set[str] = set()

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href.lower().endswith(".zip"):
            continue

        full_url = urljoin(base_url, href)
        if full_url in seen:
            continue

        seen.add(full_url)
        biome, category = infer_from_url(full_url)
        filename = urlparse(full_url).path.split("/")[-1]
        results.append(
            {
                "url": full_url,
                "filename": filename,
                "filename_local": filename,
                "text": a.get_text(strip=True) or "(no text)",
                "biome": biome,
                "category": category,
            }
        )

    return results


def fetch_static(url: str, session: requests.Session) -> list[ZipEntry]:
    print(f"[scrape] Static request: {url}")
    response = session.get(url, timeout=HTTP_TIMEOUT)
    response.raise_for_status()
    return extract_zip_links(response.text, base_url=url)


def fetch_dynamic(url: str, wait_seconds: int = 8) -> list[ZipEntry]:
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
    except ImportError as exc:
        raise RuntimeError(
            "No ZIP links were found by static scraping and Selenium is not installed. "
            "Install optional dependencies with: pip install selenium webdriver-manager"
        ) from exc

    print(f"[scrape] Dynamic rendering with Selenium: {url}")
    options = Options()
    for arg in ("--headless=new", "--no-sandbox", "--disable-dev-shm-usage"):
        options.add_argument(arg)
    options.add_argument(f"user-agent={HEADERS['User-Agent']}")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )
    try:
        driver.get(url)
        time.sleep(wait_seconds)
        return extract_zip_links(driver.page_source, base_url=url)
    finally:
        driver.quit()


def dest_path(entry: ZipEntry, dest_folder: Path) -> Path:
    return dest_folder / entry["biome"] / entry["category"] / entry["filename_local"]


def tmp_path(entry: ZipEntry, dest_folder: Path) -> Path:
    return dest_path(entry, dest_folder).with_suffix(".tmp")


def find_existing(filename: str, root_folder: Path) -> Path | None:
    if not root_folder.exists():
        return None
    for candidate in root_folder.rglob(filename):
        if candidate.suffix.lower() == ".tmp":
            continue
        if candidate.is_file() and candidate.stat().st_size > 0:
            return candidate
    return None


def is_already_downloaded(entry: ZipEntry, root_folder: Path, dest_folder: Path) -> Path | None:
    current = dest_path(entry, dest_folder)
    if current.is_file() and current.stat().st_size > 0:
        return current
    return find_existing(entry["filename_local"], root_folder)


def folder_label(path: Path, root_folder: Path) -> str:
    try:
        rel_parts = path.relative_to(root_folder).parts
        return rel_parts[0] if rel_parts else str(path.parent)
    except ValueError:
        return str(path.parent)


def print_table(rows: list[tuple[str, str]], headers: tuple[str, str]) -> None:
    h1, h2 = headers
    w1 = max([len(h1)] + [len(row[0]) for row in rows])
    w2 = max([len(h2)] + [len(row[1]) for row in rows])
    print(f"  {h1:<{w1}}  {h2:<{w2}}")
    print(f"  {'-' * w1}  {'-' * w2}")
    for c1, c2 in rows:
        print(f"  {c1:<{w1}}  {c2:<{w2}}")


def print_inventory(zips: list[ZipEntry], root_folder: Path, dest_folder: Path) -> list[ZipEntry]:
    already: list[tuple[str, str]] = []
    pending_rows: list[tuple[str, str]] = []
    pending: list[ZipEntry] = []

    ordered = sorted(zips, key=lambda x: (x["biome"], x["category"], x["filename_local"]))
    for entry in ordered:
        existing = is_already_downloaded(entry, root_folder, dest_folder)
        if existing is not None:
            already.append((entry["filename_local"], folder_label(existing, root_folder)))
        else:
            pending_rows.append((entry["filename_local"], str(dest_folder)))
            pending.append(entry)

    print(f"\n{SEP}\n  INVENTORY\n{SEP}")
    print(f"\n  [A] ALREADY DOWNLOADED ({len(already)} file(s))\n")
    print_table(already, ("FILE", "FOLDER")) if already else print("  (none)")

    print(f"\n  [B] TO BE DOWNLOADED ({len(pending)} file(s))\n")
    print_table(pending_rows, ("FILE", "DESTINATION")) if pending_rows else print("  (none)")

    print(f"\n  Total on site: {len(zips)}")
    print(f"  Already have : {len(already)}")
    print(f"  To download  : {len(pending)}")
    print(SEP)
    return pending


def ask_confirmation(pending: list[ZipEntry], yes: bool) -> bool:
    if not pending:
        print("\n  Nothing to download. All files are already present.")
        return False
    if yes:
        return True

    while True:
        answer = input(f"\n  Download {len(pending)} file(s)? [y/n]: ").strip().lower()
        if answer in {"y", "yes"}:
            return True
        if answer in {"n", "no"}:
            print("  Download cancelled.")
            return False
        print("  Please enter 'y' or 'n'.")


def verify_zip(path: Path) -> tuple[bool, str]:
    if not path.exists() or path.stat().st_size == 0:
        return False, "file missing or empty"
    try:
        with zipfile.ZipFile(path, "r") as z:
            bad = z.testzip()
        return (False, f"corrupt member: {bad}") if bad else (True, "ok")
    except zipfile.BadZipFile as exc:
        return False, f"BadZipFile: {exc}"
    except Exception as exc:
        return False, f"error: {exc}"


def download_one(
    entry: ZipEntry,
    root_folder: Path,
    dest_folder: Path,
    session: requests.Session,
    force: bool = False,
    index: int = 0,
    total: int = 0,
) -> tuple[str, str, str | None]:
    name = entry["filename_local"]
    dest = dest_path(entry, dest_folder)
    tmp = tmp_path(entry, dest_folder)
    label = f"[{index}/{total}] {name}"

    if not force:
        existing = is_already_downloaded(entry, root_folder, dest_folder)
        if existing is not None:
            print(f"  SKIP  {label}  (exists: {existing})")
            return name, "skipped", str(existing)

    dest.parent.mkdir(parents=True, exist_ok=True)
    bytes_done = tmp.stat().st_size if tmp.exists() else 0
    print(f"  {'RESUME' if bytes_done else 'START '} {label}")

    extra_headers = {"Range": f"bytes={bytes_done}-"} if bytes_done else {}

    try:
        with session.get(
            entry["url"], headers=extra_headers, stream=True, timeout=DOWNLOAD_TIMEOUT
        ) as response:
            if response.status_code == 416:
                bytes_done = 0
                tmp.unlink(missing_ok=True)
                response.close()
                response = session.get(entry["url"], stream=True, timeout=DOWNLOAD_TIMEOUT)

            if response.status_code == 403:
                return name, "forbidden", "403 Forbidden"
            if response.status_code not in (200, 206):
                response.raise_for_status()

            content_length = response.headers.get("Content-Length")
            total_bytes = int(content_length) + bytes_done if content_length else None
            mode = "ab" if bytes_done else "wb"

            with open(tmp, mode) as file, tqdm(
                total=total_bytes,
                initial=bytes_done,
                unit="B",
                unit_scale=True,
                unit_divisor=1024,
                desc=f"    {name[:55]}",
                ncols=90,
            ) as bar:
                for chunk in response.iter_content(chunk_size=CHUNK_SIZE):
                    if chunk:
                        file.write(chunk)
                        bar.update(len(chunk))

        tmp.rename(dest)
        print(f"  DONE   {label}")
        return name, "ok", None

    except KeyboardInterrupt:
        print(f"\n  Interrupted. Progress saved in: {tmp}")
        raise
    except Exception as exc:
        print(f"  ERROR  {label}: {exc}")
        return name, "error", str(exc)


def download_all(
    zips: list[ZipEntry],
    root_folder: Path,
    dest_folder: Path,
    session: requests.Session,
) -> dict[str, list]:
    dest_folder.mkdir(parents=True, exist_ok=True)
    summary: dict[str, list] = {"ok": [], "error": [], "skipped": [], "forbidden": []}

    print(f"\n{SEP}\n  SEQUENTIAL DOWNLOAD ({len(zips)} file(s))\n{SEP}\n")
    for i, entry in enumerate(zips, start=1):
        name, status, detail = download_one(entry, root_folder, dest_folder, session, False, i, len(zips))
        if status == "error":
            summary["error"].append({"file": name, "error": detail})
        else:
            summary[status].append(name)
    return summary


def validate_and_repair(
    zips: list[ZipEntry],
    root_folder: Path,
    dest_folder: Path,
    session: requests.Session,
    max_attempts: int = 3,
) -> dict[str, list[str]]:
    print(f"\n{SEP}\n  POST-DOWNLOAD VALIDATION ({len(zips)} file(s) expected)\n{SEP}")

    intact: list[str] = []
    missing: list[ZipEntry] = []
    corrupt: list[ZipEntry] = []

    for entry in tqdm(zips, desc="  Checking", unit="file", ncols=80):
        existing = is_already_downloaded(entry, root_folder, dest_folder)
        if existing is None:
            missing.append(entry)
            continue
        ok, reason = verify_zip(existing)
        if ok:
            intact.append(entry["filename_local"])
        else:
            tqdm.write(f"    [CORRUPT] {entry['filename_local']}: {reason}")
            corrupt.append(entry)

    need_repair = missing + corrupt
    failed: list[str] = []

    for attempt in range(1, max_attempts + 1):
        if not need_repair:
            break
        print(f"\n{SEP}\n  REPAIR ATTEMPT {attempt}/{max_attempts} ({len(need_repair)} file(s))\n{SEP}")

        for entry in need_repair:
            for path in (dest_path(entry, dest_folder), tmp_path(entry, dest_folder)):
                if path.exists():
                    path.unlink()

        for i, entry in enumerate(need_repair, start=1):
            download_one(entry, root_folder, dest_folder, session, True, i, len(need_repair))

        still_bad: list[ZipEntry] = []
        for entry in need_repair:
            existing = is_already_downloaded(entry, root_folder, dest_folder)
            ok = existing is not None and verify_zip(existing)[0]
            if ok:
                intact.append(entry["filename_local"])
            else:
                still_bad.append(entry)
        need_repair = still_bad

    failed = [entry["filename_local"] for entry in need_repair]
    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "intact": len(set(intact)),
            "missing_or_corrupt_after_repair": len(failed),
        },
        "intact": sorted(set(intact)),
        "failed_after_repair": failed,
    }
    report_path = dest_folder / "validation_report.json"
    with open(report_path, "w", encoding="utf-8") as file:
        json.dump(report, file, ensure_ascii=False, indent=2)
    print(f"\n  Validation report saved to: {report_path}")
    return {"intact": sorted(set(intact)), "failed": failed}


def save_metadata(zips: list[ZipEntry], dest_folder: Path) -> None:
    dest_folder.mkdir(parents=True, exist_ok=True)

    csv_path = dest_folder / "terrabrasilis_zips.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["biome", "category", "filename", "text", "url"])
        writer.writeheader()
        writer.writerows(zips)

    json_path = dest_folder / "terrabrasilis_zips.json"
    with open(json_path, "w", encoding="utf-8") as file:
        json.dump(
            {
                "scraped_at": datetime.now(timezone.utc).isoformat(),
                "count": len(zips),
                "files": zips,
            },
            file,
            ensure_ascii=False,
            indent=2,
        )

    print(f"  Metadata saved to: {csv_path}")
    print(f"  Metadata saved to: {json_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="List, download, resume, and validate public TerraBrasilis/INPE ZIP files."
    )
    parser.add_argument(
        "--root",
        type=Path,
        default=Path("data/raw/terrabrasilis"),
        help="Root folder where all TerraBrasilis downloads are stored.",
    )
    parser.add_argument(
        "--date-folder",
        default=datetime.now().strftime("%Y-%m-%d"),
        help="Subfolder name for the current download batch.",
    )
    parser.add_argument(
        "--url",
        default=BASE_URL,
        help="TerraBrasilis/INPE download page URL.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Download without interactive confirmation.",
    )
    parser.add_argument(
        "--dynamic",
        action="store_true",
        help="Force Selenium rendering instead of static scraping.",
    )
    parser.add_argument(
        "--max-repair-attempts",
        type=int,
        default=3,
        help="Maximum number of repair attempts for missing or corrupt ZIP files.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    root_folder = args.root
    dest_folder = root_folder / args.date_folder
    session = make_session()

    print(f"\n{SEP}\n  STEP 1 -- Scraping TerraBrasilis/INPE\n{SEP}")
    zips = fetch_dynamic(args.url) if args.dynamic else fetch_static(args.url, session)
    if not zips:
        print("  No ZIPs found by static scraping. Trying Selenium...")
        zips = fetch_dynamic(args.url)

    if not zips:
        print("  No ZIP files found. Nothing to do.")
        return

    print(f"  Found {len(zips)} ZIP file(s).")

    print(f"\n{SEP}\n  STEP 2 -- Inventory check\n{SEP}")
    pending = print_inventory(zips, root_folder, dest_folder)

    print(f"\n{SEP}\n  STEP 3 -- Confirmation\n{SEP}")
    if not ask_confirmation(pending, yes=args.yes):
        print("\n  Exiting without downloading.")
        return

    save_metadata(zips, dest_folder)

    print(f"\n{SEP}\n  STEP 4 -- Download and validation\n{SEP}")
    to_download = [entry for entry in zips if is_already_downloaded(entry, root_folder, dest_folder) is None]
    download_summary = download_all(to_download, root_folder, dest_folder, session)
    print(f"\n  Downloaded: {len(download_summary['ok'])}")
    print(f"  Skipped   : {len(download_summary['skipped'])}")
    print(f"  Forbidden : {len(download_summary['forbidden'])}")
    print(f"  Errors    : {len(download_summary['error'])}")

    validation = validate_and_repair(
        zips,
        root_folder,
        dest_folder,
        session,
        max_attempts=args.max_repair_attempts,
    )
    print(f"\n  Intact after validation: {len(validation['intact'])}")
    print(f"  Failed after repair    : {len(validation['failed'])}")
    print("\n  Done.")


if __name__ == "__main__":
    main()
