"""
01_download_zips.py
===================
Discovers, lists, downloads, and verifies all .zip files at:
  https://terrabrasilis.dpi.inpe.br/en/download-files/

Saves to: C:\\Amintas\\Prodes\\zip\\<today's-date>\\

Workflow
--------
1. Scrape the TerraBrasilis download page for all .zip links.
2. For each file, check whether it already exists anywhere under ROOT_FOLDER.
3. Print a two-section inventory table:
      Section A - files already downloaded (filename + dated folder found)
      Section B - files still missing (filename + will be downloaded)
4. Ask the user to confirm before downloading anything.
5. Download only the missing files, one at a time, with resume support.
6. Validate every expected file (ZIP integrity) and repair if needed.

Skip logic
----------
A file is skipped if a non-empty, non-.tmp file with the same name exists
anywhere under ROOT_FOLDER, regardless of subfolder depth or structure.

Resume support
--------------
Each file is streamed to a .tmp file first. If the script is interrupted,
the .tmp file is kept on disk. On the next run, an HTTP Range request
resumes from where it stopped — no bytes are re-downloaded.

Usage
-----
    python 01_download_zips.py

Author
------
Amintas Brandão Jr. <abrandaojr@gmail.com>
Imazon — Instituto do Homem e Meio Ambiente da Amazônia

License
-------
MIT
"""

from __future__ import annotations

__version__ = "1.0.0"
__all__: list[str] = []

import importlib.util
import subprocess
import sys


# ---------------------------------------------------------------------------
# Dependency bootstrap
# ---------------------------------------------------------------------------

def _bootstrap(*packages: tuple[str, str]) -> None:
    """Install missing packages at runtime.

    Tries uv first (better wheel resolution for native libs on Windows).
    Self-installs uv via pip if it is not found on PATH.
    Falls back to plain pip as a last resort.
    """
    import shutil

    missing = [pip for pip, mod in packages if not importlib.util.find_spec(mod)]
    if not missing:
        return

    if not shutil.which("uv"):
        subprocess.call(
            [sys.executable, "-m", "pip", "install", "--quiet", "uv"],
            stderr=subprocess.DEVNULL,
        )

    strategies = [
        ["uv", "pip", "install", "--python", sys.executable, "--quiet", *missing],
        [sys.executable, "-m", "uv", "pip", "install", "--python", sys.executable, "--quiet", *missing],
        [sys.executable, "-m", "pip", "install", "--quiet", *missing],
        [sys.executable, "-m", "pip", "install", "--quiet", "--break-system-packages", *missing],
    ]
    for cmd in strategies:
        try:
            subprocess.check_call(cmd, stderr=subprocess.DEVNULL)
            return
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass
    sys.exit(f"[FATAL] Could not install: {' '.join(missing)}")


_bootstrap(
    ("requests",       "requests"),
    ("beautifulsoup4", "bs4"),
    ("lxml",           "lxml"),
    ("tqdm",           "tqdm"),
)


# ---------------------------------------------------------------------------
# Imports
# ---------------------------------------------------------------------------

import csv
import json
import time
import zipfile as zf
from datetime import datetime, timezone
from pathlib import Path
from typing import TypedDict
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from tqdm import tqdm
from urllib3.util.retry import Retry


# ---------------------------------------------------------------------------
# Types
# ---------------------------------------------------------------------------

class ZipEntry(TypedDict):
    url:            str
    filename:       str
    filename_local: str
    text:           str
    biome:          str
    category:       str


# ---------------------------------------------------------------------------
# CONFIG  ← the only section that needs to be edited
# ---------------------------------------------------------------------------

CONFIG: dict[str, object] = {
    "base_url":         "https://terrabrasilis.dpi.inpe.br/en/download-files/",
    "root_folder":      r"C:\Amintas\Prodes\zip",
    "http_timeout":     30,
    "download_timeout": 600,
    "chunk_size":       8 * 1024 * 1024,   # 8 MB
    # Files to permanently skip (exact filename, case-sensitive).
    "skip_files": [
        "prodes_brasil_2023_arte.zip",
    ],
}

# ---------------------------------------------------------------------------
# Module-level constants derived from CONFIG
# ---------------------------------------------------------------------------

BASE_URL         = str(CONFIG["base_url"])
ROOT_FOLDER      = Path(str(CONFIG["root_folder"]))
DEST_FOLDER      = ROOT_FOLDER / datetime.now().strftime("%Y-%m-%d")
HTTP_TIMEOUT     = int(CONFIG["http_timeout"])
DOWNLOAD_TIMEOUT = int(CONFIG["download_timeout"])
CHUNK_SIZE       = int(CONFIG["chunk_size"])
SKIP_FILES       = frozenset(CONFIG["skip_files"])  # type: ignore[arg-type]
SEP              = "=" * 65
DIV              = "-" * 65

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


# ---------------------------------------------------------------------------
# HTTP session
# ---------------------------------------------------------------------------

def _make_session() -> requests.Session:
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
    session.mount("http://",  adapter)
    return session


# ---------------------------------------------------------------------------
# Scraping - static
# ---------------------------------------------------------------------------

def fetch_static(url: str) -> list[ZipEntry]:
    print(f"[scrape] Static request at: {url}")
    resp = _make_session().get(url, timeout=HTTP_TIMEOUT)
    resp.raise_for_status()
    return _extract_zip_links(BeautifulSoup(resp.text, "lxml"), base_url=url)


def _extract_zip_links(soup: BeautifulSoup, base_url: str) -> list[ZipEntry]:
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
        biome, category = _infer_from_url(full_url)
        results.append({
            "url":            full_url,
            "filename":       urlparse(full_url).path.split("/")[-1],
            "filename_local": "",
            "text":           a.get_text(strip=True) or "(no text)",
            "biome":          biome,
            "category":       category,
        })
    return results


_BIOME_MAP: dict[str, str] = {
    "amz-prodes":            "Amazon Biome",
    "amz-aux":               "Amazon Biome",
    "amz-terraclass":        "Amazon Biome",
    "legal-amz-prodes":      "Legal Amazon",
    "legal-amz-aux":         "Legal Amazon",
    "caatinga-prodes":       "Caatinga",
    "caatinga-aux":          "Caatinga",
    "cerrado-prodes":        "Cerrado",
    "cerrado-aux":           "Cerrado",
    "cerrado-vegetation":    "Cerrado",
    "mata-atlantica-prodes": "Mata Atlantica",
    "mata-atlantica-aux":    "Mata Atlantica",
    "pampa-prodes":          "Pampa",
    "pampa-aux":             "Pampa",
    "pantanal-prodes":       "Pantanal",
    "pantanal-aux":          "Pantanal",
    "brasil-prodes":         "Brazil",
    "vs":                    "Vegetacao Secundaria",
}


def _infer_from_url(url: str) -> tuple[str, str]:
    parts = urlparse(url).path.strip("/").split("/")
    try:
        slug      = parts[2]
        file_type = parts[3].capitalize()
    except IndexError:
        return "N/A", "N/A"
    return _BIOME_MAP.get(slug, slug), file_type


# ---------------------------------------------------------------------------
# Scraping - dynamic (Selenium fallback)
# ---------------------------------------------------------------------------

def fetch_dynamic(url: str, wait_seconds: int = 8) -> list[ZipEntry]:
    _bootstrap(
        ("selenium",          "selenium"),
        ("webdriver-manager", "webdriver_manager"),
    )
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    print(f"[scrape] Dynamic rendering (Selenium) at: {url}")
    opts = Options()
    for arg in ("--headless=new", "--no-sandbox", "--disable-dev-shm-usage"):
        opts.add_argument(arg)
    opts.add_argument(f"user-agent={HEADERS['User-Agent']}")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=opts
    )
    try:
        driver.get(url)
        print(f"    Waiting {wait_seconds}s for JS to render...")
        time.sleep(wait_seconds)
        _expand_all_menus(driver)
        time.sleep(3)
        return _extract_zip_links(BeautifulSoup(driver.page_source, "lxml"), base_url=url)
    finally:
        driver.quit()


def _expand_all_menus(driver) -> None:
    try:
        from selenium.webdriver.common.by import By
        toggles = driver.find_elements(
            By.CSS_SELECTOR,
            "li.menu-item-has-children > a, .accordion-header, details > summary",
        )
        for el in toggles:
            try:
                driver.execute_script("arguments[0].click();", el)
                time.sleep(0.3)
            except Exception:
                pass
        print(f"    {len(toggles)} menu(s) expanded.")
    except Exception as exc:
        print(f"    Warning while expanding menus: {exc}")


# ---------------------------------------------------------------------------
# Filename resolution + path helpers
# ---------------------------------------------------------------------------

def resolve_unique_filenames(zips: list[ZipEntry]) -> list[ZipEntry]:
    result: list[ZipEntry] = []
    for z in zips:
        z = z.copy()
        z["filename_local"] = z["filename"]
        result.append(z)
    return result


def _dest_path(z: ZipEntry, folder: Path) -> Path:
    return folder / z["biome"] / z["category"] / z["filename_local"]


def _tmp_path(z: ZipEntry, folder: Path) -> Path:
    return _dest_path(z, folder).with_suffix(".tmp")


def _find_existing(filename: str) -> Path | None:
    """
    Walk the entire ROOT_FOLDER tree and return the first non-empty, non-.tmp
    file whose name matches `filename`, regardless of subfolder depth or
    structure used in previous runs.
    """
    if not ROOT_FOLDER.exists():
        return None
    for candidate in ROOT_FOLDER.rglob(filename):
        if candidate.suffix.lower() == ".tmp":
            continue
        if candidate.is_file() and candidate.stat().st_size > 0:
            return candidate
    return None


def _is_already_downloaded(z: ZipEntry, folder: Path) -> Path | None:
    """
    Return the path of a valid existing copy, or None.
    Checks today's destination first (fast), then the full ROOT_FOLDER tree.
    """
    dest = _dest_path(z, folder)
    if dest.is_file() and dest.stat().st_size > 0:
        return dest
    return _find_existing(z["filename_local"])


# ---------------------------------------------------------------------------
# Pre-download inventory table + confirmation
# ---------------------------------------------------------------------------

def _folder_label(path: Path) -> str:
    """
    Return the dated subfolder name (e.g. '2024-11-03') where the file lives,
    rather than the full path.
    """
    try:
        rel_parts = path.relative_to(ROOT_FOLDER).parts
        return rel_parts[0] if rel_parts else str(path.parent)
    except ValueError:
        return str(path.parent)


def _print_table(rows: list[tuple[str, str]], col_headers: tuple[str, str]) -> None:
    """Print a simple two-column fixed-width table."""
    h1, h2 = col_headers
    w1 = max((len(r[0]) for r in rows), default=0)
    w1 = max(w1, len(h1))
    w2 = max((len(r[1]) for r in rows), default=0)
    w2 = max(w2, len(h2))

    divider = f"  {'-' * w1}  {'-' * w2}"
    print(f"  {h1:<{w1}}  {h2:<{w2}}")
    print(divider)
    for c1, c2 in rows:
        print(f"  {c1:<{w1}}  {c2:<{w2}}")
    print(divider)


def print_inventory_table(zips: list[ZipEntry]) -> list[ZipEntry]:
    """
    Print two separate tables:
      A) Files already downloaded -- filename + dated folder where found
      B) Files to be downloaded   -- filename + destination folder

    Returns the list of ZipEntry items that still need downloading.
    """
    already:  list[tuple[str, str]] = []   # (filename, dated-folder)
    pending_rows: list[tuple[str, str]] = []   # (filename, dest)
    pending_entries: list[ZipEntry]     = []

    for z in sorted(zips, key=lambda x: (x["biome"], x["category"], x["filename_local"])):
        existing = _is_already_downloaded(z, DEST_FOLDER)
        if existing is not None:
            already.append((z["filename_local"], _folder_label(existing)))
        else:
            pending_rows.append((z["filename_local"], str(DEST_FOLDER)))
            pending_entries.append(z)

    print(f"\n{SEP}")
    print("  INVENTORY  --  files found on TerraBrasilis")
    print(SEP)

    # --- Section A: already downloaded ---
    print(f"\n  [A] ALREADY DOWNLOADED ({len(already)} file(s))\n")
    if already:
        _print_table(already, ("FILE", "FOLDER"))
    else:
        print("  (none)")

    # --- Section B: to be downloaded ---
    print(f"\n  [B] TO BE DOWNLOADED ({len(pending_entries)} file(s))\n")
    if pending_rows:
        _print_table(pending_rows, ("FILE", "DESTINATION"))
    else:
        print("  (none -- all files already present)")

    # --- Summary line ---
    print(f"\n  Total on site  : {len(zips)}")
    print(f"  Already have   : {len(already)}")
    print(f"  To download    : {len(pending_entries)}")
    print(SEP)

    return pending_entries


def ask_confirmation(pending: list[ZipEntry]) -> bool:
    """
    Ask the user whether to proceed with the download.
    Returns True if the user confirms, False otherwise.
    """
    if not pending:
        print("\n  Nothing to download. All files are already present.")
        return False

    print(f"\n  {len(pending)} file(s) will be downloaded to:")
    print(f"  {DEST_FOLDER}")
    print()

    while True:
        answer = input("  Proceed with download? [y/n]: ").strip().lower()
        if answer in ("y", "yes"):
            return True
        if answer in ("n", "no"):
            print("  Download cancelled.")
            return False
        print("  Please enter 'y' or 'n'.")


# ---------------------------------------------------------------------------
# Single-file sequential download with resume
# ---------------------------------------------------------------------------

def _download_one(
    z: ZipEntry,
    folder: Path,
    session: requests.Session,
    force: bool = False,
    index: int = 0,
    total: int = 0,
) -> tuple[str, str, str | None]:
    """
    Download one file.  Skips if a valid copy exists anywhere under ROOT_FOLDER.
    Resumes from a .tmp partial file if present.

    Returns (filename_local, status, detail).
    status: 'ok' | 'skipped' | 'error' | 'forbidden'
    """
    name  = z["filename_local"]
    dest  = _dest_path(z, folder)
    tmp   = _tmp_path(z, folder)
    label = f"[{index}/{total}] {name}"

    if not force:
        existing = _is_already_downloaded(z, folder)
        if existing is not None:
            print(f"  SKIP  {label}  (exists: {existing})")
            return name, "skipped", str(existing)

    dest.parent.mkdir(parents=True, exist_ok=True)

    bytes_done = tmp.stat().st_size if tmp.exists() else 0
    if bytes_done:
        print(f"  RESUME {label}  (have {bytes_done / 1_048_576:.1f} MB)")
    else:
        print(f"  START  {label}")

    extra_headers = {"Range": f"bytes={bytes_done}-"} if bytes_done else {}

    try:
        with session.get(
            z["url"], headers=extra_headers, stream=True, timeout=DOWNLOAD_TIMEOUT
        ) as r:

            if r.status_code == 416:
                print(f"    Range rejected, restarting from byte 0...")
                bytes_done    = 0
                extra_headers = {}
                tmp.unlink(missing_ok=True)
                r.close()
                r = session.get(z["url"], stream=True, timeout=DOWNLOAD_TIMEOUT)

            if r.status_code == 403:
                return name, "forbidden", "403 Forbidden"

            if r.status_code not in (200, 206):
                r.raise_for_status()

            content_length = r.headers.get("Content-Length")
            total_bytes    = (int(content_length) + bytes_done) if content_length else None

            mode = "ab" if bytes_done else "wb"
            with open(tmp, mode) as f, tqdm(
                total=total_bytes,
                initial=bytes_done,
                unit="B",
                unit_scale=True,
                unit_divisor=1024,
                desc=f"    {name[:55]}",
                ncols=90,
                miniters=1,
            ) as bar:
                for chunk in r.iter_content(chunk_size=CHUNK_SIZE):
                    f.write(chunk)
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


# ---------------------------------------------------------------------------
# Sequential download loop
# ---------------------------------------------------------------------------

def download_all(
    zips: list[ZipEntry],
    folder: Path,
    session: requests.Session,
) -> dict:
    folder.mkdir(parents=True, exist_ok=True)
    n = len(zips)
    summary: dict[str, list] = {"ok": [], "error": [], "skipped": [], "forbidden": []}

    print(f"\n{SEP}")
    print(f"  SEQUENTIAL DOWNLOAD  ({n} file(s))")
    print(f"  Destination  : {folder}")
    print(f"  Skip if found: anywhere under {ROOT_FOLDER}")
    print(f"{SEP}\n")

    for i, z in enumerate(zips, start=1):
        name, status, detail = _download_one(
            z, folder, session, force=False, index=i, total=n
        )
        if status == "ok":
            summary["ok"].append(name)
        elif status == "skipped":
            summary["skipped"].append(name)
        elif status == "forbidden":
            summary["forbidden"].append(name)
        else:
            summary["error"].append({"file": name, "error": detail})

    return summary


def print_download_summary(summary: dict, folder: Path) -> None:
    print(f"\n{SEP}\n  DOWNLOAD SUMMARY\n{SEP}")
    print(f"  Successfully downloaded  : {len(summary['ok'])}")
    print(f"  Already existed (skipped): {len(summary['skipped'])}")
    print(f"  Forbidden (403)          : {len(summary['forbidden'])}")
    print(f"  Errors                   : {len(summary['error'])}")
    if summary["error"]:
        print("\n  Files with errors (will be retried in validation phase):")
        for e in summary["error"]:
            print(f"    * {e['file']}: {e['error']}")
    print(f"\n  Destination folder: {folder}\n{SEP}")


# ---------------------------------------------------------------------------
# ZIP integrity check
# ---------------------------------------------------------------------------

def verify_zip(path: Path) -> tuple[bool, str]:
    if not path.exists() or path.stat().st_size == 0:
        return False, "file missing or empty"
    try:
        with zf.ZipFile(path, "r") as z:
            bad = z.testzip()
            return (False, f"corrupt member: {bad}") if bad else (True, "ok")
    except zf.BadZipFile as exc:
        return False, f"BadZipFile: {exc}"
    except Exception as exc:
        return False, f"error: {exc}"


# ---------------------------------------------------------------------------
# Post-download validation + repair
# ---------------------------------------------------------------------------

def validate_and_repair(
    zips: list[ZipEntry],
    folder: Path,
    session: requests.Session,
    max_attempts: int = 3,
) -> dict:
    """
    Phase 1 -- For each expected file, search anywhere under ROOT_FOLDER.
               Only files absent from the entire tree are marked missing.
               Found files are integrity-checked; corrupt ones are repaired.

    Phase 2 -- Repair loop: re-download missing/corrupt into today's folder.

    Phase 3 -- Save validation_report.json.
    """
    n = len(zips)
    print(f"\n{SEP}")
    print(f"  POST-DOWNLOAD VALIDATION  ({n} file(s) expected)")
    print(f"  Searching anywhere under: {ROOT_FOLDER}")
    print(SEP)

    # ---- Phase 1 --------------------------------------------------------

    intact:  list[str]      = []
    missing: list[ZipEntry] = []
    corrupt: list[ZipEntry] = []

    for z in tqdm(zips, desc="  Checking", unit="file", ncols=80):
        name     = z["filename_local"]
        existing = _is_already_downloaded(z, folder)

        if existing is None:
            tqdm.write(f"    [MISSING]  {name}")
            missing.append(z)
            continue

        ok, reason = verify_zip(existing)
        if ok:
            intact.append(name)
        else:
            tqdm.write(f"    [CORRUPT]  {name}  ({existing}): {reason}")
            corrupt.append(z)

    need_repair = missing + corrupt

    print(f"\n  Intact  : {len(intact)}")
    print(f"  Missing : {len(missing)}")
    print(f"  Corrupt : {len(corrupt)}")

    if not need_repair:
        print(f"\n  All {len(intact)} file(s) passed validation.")
        _save_report(folder, intact, [], [], [])
        return {"intact": intact, "missing": [], "corrupt": [], "failed": []}

    print(f"\n  {len(need_repair)} file(s) need repair.\n")

    # ---- Phase 2 --------------------------------------------------------

    failed: list[str] = []

    for attempt in range(1, max_attempts + 1):
        if not need_repair:
            break

        print(f"\n{SEP}")
        print(f"  REPAIR ATTEMPT {attempt}/{max_attempts}  ({len(need_repair)} file(s))")
        print(SEP + "\n")

        # Remove bad copies in today's folder only; leave previous dated folders untouched
        for z in need_repair:
            for p in (_dest_path(z, folder), _tmp_path(z, folder)):
                if p.exists():
                    p.unlink()

        for i, z in enumerate(need_repair, start=1):
            _download_one(z, folder, session, force=True, index=i, total=len(need_repair))

        still_bad: list[ZipEntry] = []
        for z in need_repair:
            name     = z["filename_local"]
            existing = _is_already_downloaded(z, folder)
            ok       = existing is not None and verify_zip(existing)[0]
            if ok:
                intact.append(name)
                missing[:] = [m for m in missing if m["filename_local"] != name]
                corrupt[:] = [c for c in corrupt if c["filename_local"] != name]
            else:
                reason = "not found" if existing is None else verify_zip(existing)[1]
                print(f"  [STILL FAILING] {name}: {reason}")
                still_bad.append(z)

        need_repair = still_bad

    for z in need_repair:
        name = z["filename_local"]
        if name not in failed:
            failed.append(name)

    # ---- Phase 3 --------------------------------------------------------

    _save_report(
        folder,
        intact,
        [z["filename_local"] for z in missing],
        [z["filename_local"] for z in corrupt],
        failed,
    )

    return {
        "intact":  intact,
        "missing": [z["filename_local"] for z in missing],
        "corrupt": [z["filename_local"] for z in corrupt],
        "failed":  failed,
    }


def _save_report(
    folder: Path,
    intact:  list[str],
    missing: list[str],
    corrupt: list[str],
    failed:  list[str],
) -> None:
    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "intact":              len(intact),
            "missing":             len(missing),
            "corrupt":             len(corrupt),
            "failed_after_repair": len(failed),
        },
        "intact":              intact,
        "missing":             missing,
        "corrupt":             corrupt,
        "failed_after_repair": failed,
    }
    path = folder / "validation_report.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"\n  Validation report saved to: {path}")


def print_validation_summary(summary: dict) -> None:
    print(f"\n{SEP}\n  VALIDATION RESULT\n{SEP}")
    print(f"  Intact                   : {len(summary['intact'])}")
    print(f"  Missing (not downloaded) : {len(summary['missing'])}")
    print(f"  Corrupt (bad ZIP)        : {len(summary['corrupt'])}")
    print(f"  Failed after all repairs : {len(summary['failed'])}")
    if summary["failed"]:
        print("\n  Files that could not be recovered:")
        for name in summary["failed"]:
            print(f"    * {name}")
    else:
        print(f"\n  All files intact.")
    print(SEP)


# ---------------------------------------------------------------------------
# Metadata export
# ---------------------------------------------------------------------------

def save_csv(zips: list[ZipEntry], folder: Path) -> None:
    path = folder / "terrabrasilis_zips.csv"
    fields = ["biome", "category", "filename", "text", "url"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(zips)
    print(f"  CSV saved to: {path}")


def save_json(zips: list[ZipEntry], folder: Path) -> None:
    path = folder / "terrabrasilis_zips.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "scraped_at": datetime.now(timezone.utc).isoformat(),
                "count":      len(zips),
                "files":      zips,
            },
            f, ensure_ascii=False, indent=2,
        )
    print(f"  JSON saved to: {path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  TerraBrasilis Download  v{__version__}  |  {now}")
    print(f"{SEP}")

    session     = _make_session()
    MAX_PASSES  = 20
    RETRY_DELAY = 30

    # ------------------------------------------------------------------ #
    # Step 1: scrape the site                                             #
    # ------------------------------------------------------------------ #
    print(f"\n{SEP}\n  STEP 1 OF 4  --  Scraping TerraBrasilis\n{SEP}")

    zips = fetch_static(BASE_URL)
    if zips:
        print(f"  -> {len(zips)} ZIP(s) found via static request.")
    else:
        print("  -> No ZIPs found via static request (JS rendering required).")
        print("  Trying Selenium...")
        zips = fetch_dynamic(BASE_URL)
        print(f"  -> {len(zips)} ZIP(s) found via Selenium.")

    if not zips:
        print("  No ZIP files found on the page. Nothing to do.")
        return

    zips = resolve_unique_filenames(zips)

    if SKIP_FILES:
        before = len(zips)
        zips   = [z for z in zips if z["filename"] not in SKIP_FILES]
        print(f"  -> {before - len(zips)} file(s) excluded by skip_files config.")

    # ------------------------------------------------------------------ #
    # Step 2: show inventory table                                        #
    # ------------------------------------------------------------------ #
    print(f"\n{SEP}\n  STEP 2 OF 4  --  Inventory check\n{SEP}")
    print(f"  Checking what is already present under: {ROOT_FOLDER}\n")

    pending = print_inventory_table(zips)

    # ------------------------------------------------------------------ #
    # Step 3: ask for confirmation                                        #
    # ------------------------------------------------------------------ #
    print(f"\n{SEP}\n  STEP 3 OF 4  --  Confirmation\n{SEP}")

    if not ask_confirmation(pending):
        print("\n  Exiting without downloading.")
        return

    # Save metadata only after user confirms (creates DEST_FOLDER)
    DEST_FOLDER.mkdir(parents=True, exist_ok=True)
    save_csv(zips, DEST_FOLDER)
    save_json(zips, DEST_FOLDER)

    # ------------------------------------------------------------------ #
    # Step 4: download + validate (with retry passes)                    #
    # ------------------------------------------------------------------ #
    print(f"\n{SEP}\n  STEP 4 OF 4  --  Download & Validation\n{SEP}")

    for pass_num in range(1, MAX_PASSES + 1):
        if pass_num > 1:
            print(f"\n{SEP}\n  RETRY PASS {pass_num}/{MAX_PASSES}\n{SEP}")

        # Recompute pending in case a previous pass partially succeeded
        pending = [z for z in zips if _is_already_downloaded(z, DEST_FOLDER) is None]

        if not pending:
            print(f"  All {len(zips)} ZIP(s) already present under {ROOT_FOLDER}. Done.")
            break

        print(f"  {len(pending)} ZIP(s) to download in this pass.")

        dl_summary = download_all(zips, DEST_FOLDER, session)
        print_download_summary(dl_summary, DEST_FOLDER)

        val_summary = validate_and_repair(zips, DEST_FOLDER, session, max_attempts=3)
        print_validation_summary(val_summary)

        remaining = len(val_summary["failed"]) + len(val_summary["missing"])
        if remaining == 0:
            print("\n  All files downloaded and validated successfully. Done.")
            break

        print(f"\n  {remaining} file(s) still incomplete after pass {pass_num}.")
        if pass_num < MAX_PASSES:
            print(f"  Retrying in {RETRY_DELAY}s...")
            time.sleep(RETRY_DELAY)
    else:
        print(f"\n  Reached {MAX_PASSES} passes. Check validation_report.json for details.")


if __name__ == "__main__":
    main()