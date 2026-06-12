from __future__ import annotations

__version__ = "1.0.0"

import importlib.util
import subprocess
import sys
import time
import zipfile as zf
from datetime import datetime, timezone
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from typing import Any, TypedDict
from urllib.parse import urljoin, urlparse

# ---------------------------------------------------------------------------
# Dependency bootstrap
# ---------------------------------------------------------------------------


def _bootstrap(*packages: tuple[str, str]) -> None:
    """Install missing packages into the current Python environment.

    Strategy order (most to least reliable for targeting sys.executable):
      1. python -m pip          â€” always installs into the running interpreter
      2. uv pip --python        â€” faster wheel resolution for native libs
      3. python -m uv pip       â€” uv via module, same target guarantee
      4. pip --break-system-pkg â€” last resort for externally-managed envs

    After each attempt, importlib.invalidate_caches() re-scans site-packages
    so that newly installed packages are immediately discoverable.
    Only packages that remain missing are retried with subsequent strategies.
    """
    import importlib
    import shutil

    mod_by_pip = {pip: mod for pip, mod in packages}

    def _still_missing(pkgs: list[str]) -> list[str]:
        importlib.invalidate_caches()
        return [p for p in pkgs if not importlib.util.find_spec(mod_by_pip[p])]

    missing = _still_missing(list(mod_by_pip))
    if not missing:
        return

    # Check for `uv` and install it if missing to enable faster dependency resolution
    if not shutil.which("uv"):
        subprocess.call(
            [sys.executable, "-m", "pip", "install", "--quiet", "uv"],
            stderr=subprocess.DEVNULL,
        )

    strategies = [
        [sys.executable, "-m", "pip", "install", "--quiet"],
        ["uv", "pip", "install", "--python", sys.executable, "--quiet"],
        [sys.executable, "-m", "uv", "pip", "install", "--python", sys.executable, "--quiet"],
        [sys.executable, "-m", "pip", "install", "--quiet", "--break-system-packages"],
    ]
    for base in strategies:
        if not missing:
            return
        try:
            subprocess.check_call(base + missing, stderr=subprocess.DEVNULL)
            missing = _still_missing(missing)
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass

    if missing:
        sys.exit(f"[FATAL] Could not install: {' '.join(missing)}")


_bootstrap(
    ("requests", "requests"),
    ("beautifulsoup4", "bs4"),
    ("lxml", "lxml"),
    ("tqdm", "tqdm"),
)


# ---------------------------------------------------------------------------
# Imports
# ---------------------------------------------------------------------------

import csv

import json
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from tqdm import tqdm
from urllib3.util.retry import Retry

from prodes_pipeline.data_quality import (
    LineageRecord,
    StageTimer,
    atomic_write_json,
    configure_json_logging,
    file_inventory,
    freshness_metrics,
    to_jsonable,
    write_run_report,
)
from prodes_pipeline.pipeline_contracts import ZIP_ARCHIVE_CONTRACT, ZIP_INVENTORY_CONTRACT
from prodes_pipeline.config import REPORTS_DIR, ZIP_ROOT, ensure_pipeline_dirs


# ---------------------------------------------------------------------------
# Types
# ---------------------------------------------------------------------------


class ZipEntry(TypedDict):
    url: str
    filename: str
    filename_local: str
    text: str
    biome: str
    category: str


# ---------------------------------------------------------------------------
# CONFIG  â† the only section that needs to be edited
# ---------------------------------------------------------------------------

CONFIG: dict[str, Any] = {
    "base_url": "https://terrabrasilis.dpi.inpe.br/en/download-files/",
    "root_folder": ZIP_ROOT,
    "http_timeout": 30,
    "download_timeout": 600,
    "chunk_size": 32 * 1024 * 1024,  # 32 MB per stream chunk â€” maximize single-file throughput
    # Files to permanently skip (exact filename, case-sensitive).
    "skip_files": [
        "prodes_brasil_2023_arte.zip",
    ],
}

# ---------------------------------------------------------------------------
# Module-level constants derived from CONFIG
# ---------------------------------------------------------------------------

BASE_URL = str(CONFIG["base_url"])
ROOT_FOLDER = Path(str(CONFIG["root_folder"]))
DEST_FOLDER = ROOT_FOLDER / datetime.now().strftime("%Y-%m-%d")
HTTP_TIMEOUT = int(CONFIG["http_timeout"])
DOWNLOAD_TIMEOUT = int(CONFIG["download_timeout"])
CHUNK_SIZE = int(CONFIG["chunk_size"])
SKIP_FILES = frozenset(CONFIG["skip_files"])
STATIC_HTML_LIMIT_BYTES = 8 * 1024 * 1024
_HTML_PARSER = "lxml" if importlib.util.find_spec("lxml") else "html.parser"
SEP = "=" * 65
DIV = "-" * 65
REPORT_DIR = REPORTS_DIR
OBS_LOG = configure_json_logging(REPORT_DIR / "observability.jsonl")
_EXISTING_ZIP_INDEX: dict[str, list[Path]] | None = None

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
    """Create a requests session with retry logic and custom headers."""
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


# ---------------------------------------------------------------------------
# Scraping - static
# ---------------------------------------------------------------------------


def fetch_static(url: str) -> list[ZipEntry]:
    """Fetch and parse ZIP links from a static HTML page."""
    print(f"[scrape] Static request at: {url}")
    session = _make_session()
    html = _fetch_static_html(session, url)
    return _extract_zip_links(BeautifulSoup(html, _HTML_PARSER), base_url=url)


def _fetch_static_html(session: requests.Session, url: str) -> bytes:
    """Fetch HTML with bounded reads so slow responses can fall back cleanly."""
    chunks: list[bytes] = []
    bytes_read = 0
    deadline = time.monotonic() + HTTP_TIMEOUT
    timeout = (min(10, HTTP_TIMEOUT), min(10, HTTP_TIMEOUT))

    with session.get(url, timeout=timeout, stream=True) as resp:
        resp.raise_for_status()
        for chunk in resp.iter_content(chunk_size=64 * 1024):
            if not chunk:
                continue
            chunks.append(chunk)
            bytes_read += len(chunk)
            if bytes_read > STATIC_HTML_LIMIT_BYTES:
                raise ValueError(
                    f"Static HTML exceeded {STATIC_HTML_LIMIT_BYTES / 1_048_576:.0f} MB"
                )
            if time.monotonic() > deadline:
                raise requests.Timeout(
                    f"Static HTML read exceeded {HTTP_TIMEOUT}s"
                )

    return b"".join(chunks)


def _extract_zip_links(soup: BeautifulSoup, base_url: str) -> list[ZipEntry]:
    """Extract .zip file links from BeautifulSoup object."""
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
        filename = urlparse(full_url).path.split("/")[-1]
        results.append({
            "url": full_url,
            "filename": filename,
            "filename_local": filename,  # Initially, local name is the same as remote
            "text": a.get_text(strip=True) or "(no text)",
            "biome": biome,
            "category": category,
        })
    return results


_BIOME_MAP: dict[str, str] = {
    "amz-prodes": "Amazon Biome",
    "amz-aux": "Amazon Biome",
    "amz-terraclass": "Amazon Biome",
    "legal-amz-prodes": "Legal Amazon",
    "legal-amz-aux": "Legal Amazon",
    "caatinga-prodes": "Caatinga",
    "caatinga-aux": "Caatinga",
    "cerrado-prodes": "Cerrado",
    "cerrado-aux": "Cerrado",
    "cerrado-vegetation": "Cerrado",
    "mata-atlantica-prodes": "Mata Atlantica",
    "mata-atlantica-aux": "Mata Atlantica",
    "pampa-prodes": "Pampa",
    "pampa-aux": "Pampa",
    "pantanal-prodes": "Pantanal",
    "pantanal-aux": "Pantanal",
    "brasil-prodes": "Brazil",
    "vs": "Vegetacao Secundaria",
}


def _infer_from_url(url: str) -> tuple[str, str]:
    """Infer biome and category from URL path slugs."""
    parts = urlparse(url).path.strip("/").split("/")
    try:
        slug = parts[2]
        file_type = parts[3].capitalize()
    except IndexError:
        return "N/A", "N/A"
    return _BIOME_MAP.get(slug, slug), file_type


# ---------------------------------------------------------------------------
# Scraping - dynamic (Selenium fallback)
# ---------------------------------------------------------------------------


def _find_chrome() -> str | None:
    """Return the Chrome executable path, or None if not found."""
    import shutil

    if exe := shutil.which("google-chrome") or \
               shutil.which("chromium-browser") or \
               shutil.which("chromium"):
        return exe
    for candidate in (
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ):
        if Path(candidate).exists():
            return candidate
    return None


def fetch_dynamic(url: str, wait_seconds: int = 8) -> list[ZipEntry]:
    """Fetch and parse ZIP links using Selenium for JavaScript-rendered pages."""
    chrome_path = _find_chrome()
    if not chrome_path:
        print(
            "  [skip] Selenium fallback requires Chrome or Brave Browser.\n"
            "         Install one and re-run, or download files manually."
        )
        return []

    _bootstrap(("selenium", "selenium"))

    import threading
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    print(f"[scrape] Dynamic rendering (Selenium) at: {url}")

    opts = Options()
    opts.binary_location = str(chrome_path)
    for arg in (
        "--headless=new", "--no-sandbox", "--disable-dev-shm-usage",
        "--disable-gpu", "--no-first-run",
    ):
        opts.add_argument(arg)
    opts.add_argument(f"user-agent={HEADERS['User-Agent']}")

    _slot: dict[str, Any] = {}

    def _init_driver() -> None:
        try:
            _slot["driver"] = webdriver.Chrome(options=opts)
        except Exception as exc:
            _slot["error"] = exc

    t = threading.Thread(target=_init_driver, daemon=True)
    t.start()
    t.join(timeout=60)  # Wait up to 60 seconds for driver initialization

    if t.is_alive():
        print(
            "  [skip] ChromeDriver init timed out after 60 s.\n"
            "         Opening the download page in your browser instead."
        )
        import webbrowser
        webbrowser.open(url)
        return []

    if "error" in _slot:
        exc = _slot["error"]
        print(f"  [skip] ChromeDriver init failed: {type(exc).__name__}: {exc}")
        return []

    driver = _slot["driver"]
    try:
        driver.set_page_load_timeout(30)
        driver.set_script_timeout(15)

        driver.get(url)
        print(f"    Waiting {wait_seconds}s for JS to render...")
        time.sleep(wait_seconds)
        _expand_all_menus(driver)
        time.sleep(3)
        return _extract_zip_links(BeautifulSoup(driver.page_source, _HTML_PARSER), base_url=url)
    except KeyboardInterrupt:
        raise
    except BaseException as exc:
        if isinstance(exc, KeyboardInterrupt):
            raise
        print(f"  [skip] Selenium failed: {type(exc).__name__}: {exc}")
        return []
    finally:
        try:
            driver.quit()
        except Exception:
            pass


def _expand_all_menus(driver: Any) -> None:
    """Expand all accordion/menu elements on the page."""
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


def _dest_path(z: ZipEntry, folder: Path) -> Path:
    """Return the final destination path for a given ZipEntry."""
    return folder / z["biome"] / z["category"] / z["filename_local"]


def _tmp_path(z: ZipEntry, folder: Path) -> Path:
    """Return the temporary download path for a given ZipEntry."""
    return _dest_path(z, folder).with_suffix(".tmp")


def _index_existing_zips() -> dict[str, list[Path]]:
    """Index valid local ZIPs once so inventory checks do not rescan the tree."""
    index: dict[str, list[Path]] = {}
    if not ROOT_FOLDER.exists():
        return {}

    for candidate in ROOT_FOLDER.rglob("*.zip"):
        try:
            st = candidate.stat()
            if not candidate.is_file() or st.st_size <= 0:
                continue
            index.setdefault(candidate.name, []).append(candidate)
        except OSError:
            continue
    for candidates in index.values():
        candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return index


def _find_existing(z: ZipEntry) -> Path | None:
    """
    Return the best local match for a remote file.

    Filename-only matching is unsafe because TerraBrasilis reuses names such
    as biome_border.zip across biomes. Prefer matches whose path also contains
    the expected biome and category; use filename-only only when unambiguous.
    """
    global _EXISTING_ZIP_INDEX
    if _EXISTING_ZIP_INDEX is None:
        _EXISTING_ZIP_INDEX = _index_existing_zips()
    candidates_raw = _EXISTING_ZIP_INDEX.get(z["filename_local"], [])
    candidates = [candidates_raw] if isinstance(candidates_raw, Path) else candidates_raw
    if not candidates:
        return None

    biome = z["biome"].lower()
    category = z["category"].lower()
    contextual = [
        path
        for path in candidates
        if biome in {part.lower() for part in path.parts}
        and category in {part.lower() for part in path.parts}
    ]
    if contextual:
        return contextual[0]
    if len(candidates) == 1:
        return candidates[0]
    return None


def _is_already_downloaded(z: ZipEntry, folder: Path) -> Path | None:
    """
    Return the path of a valid existing copy, or None.
    Checks today's destination first (fast), then the full ROOT_FOLDER tree.
    """
    dest = _dest_path(z, folder)
    if dest.is_file() and dest.stat().st_size > 0:
        return dest
    return _find_existing(z)


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
    already: list[tuple[str, str]] = []  # (filename, dated-folder)
    pending_rows: list[tuple[str, str]] = []  # (filename, dest)
    pending_entries: list[ZipEntry] = []

    for z in sorted(
        zips, key=lambda x: (x["biome"], x["category"], x["filename_local"])
    ):
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
    Download one file. Skips if a valid copy exists anywhere under ROOT_FOLDER.
    Resumes from a .tmp partial file if present.

    Returns (filename_local, status, detail).
    status: 'ok' | 'skipped' | 'error' | 'forbidden'
    """
    name = z["filename_local"]
    dest = _dest_path(z, folder)
    tmp = _tmp_path(z, folder)
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
                print("    Range rejected, restarting from byte 0...")
                bytes_done = 0
                extra_headers = {}
                tmp.unlink(missing_ok=True)
                r.close()
                r = session.get(z["url"], stream=True, timeout=DOWNLOAD_TIMEOUT)

            if r.status_code == 403:
                return name, "forbidden", "403 Forbidden"

            if r.status_code not in (200, 206):
                r.raise_for_status()

            content_length = r.headers.get("Content-Length")
            total_bytes = (int(content_length) + bytes_done) if content_length else None

            mode = "ab" if bytes_done else "wb"
            with open(tmp, mode) as f, tqdm(
                total=total_bytes,
                initial=bytes_done,
                unit="B",
                unit_scale=True,
                unit_divisor=1024,
                desc=f"    {name[:55]}",
                ncols=90,
                mininterval=0.5,  # update bar at most twice per second
            ) as bar:
                for chunk in r.iter_content(chunk_size=CHUNK_SIZE):
                    f.write(chunk)
                    bar.update(len(chunk))

        tmp.rename(dest)
        if _EXISTING_ZIP_INDEX is not None:
            cached = _EXISTING_ZIP_INDEX.get(name, [])
            candidates = [cached] if isinstance(cached, Path) else cached
            _EXISTING_ZIP_INDEX[name] = [dest, *candidates]
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
    """Download all pending ZIP files."""
    folder.mkdir(parents=True, exist_ok=True)
    n = len(zips)
    summary: dict[str, list[Any]] = {"ok": [], "error": [], "skipped": [], "forbidden": []}

    print(f"\n{SEP}")
    print(f"  DOWNLOAD  ({n} file(s)  |  full bandwidth per file)")
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


def print_download_summary(summary: dict[str, Any], folder: Path) -> None:
    """Print a summary of the download operation."""
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
    """Verify the integrity of a ZIP file."""
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
) -> dict[str, Any]:
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

    intact: list[str] = []
    missing: list[ZipEntry] = []
    corrupt: list[ZipEntry] = []

    for z in tqdm(zips, desc="  Checking", unit="file", ncols=80):
        name = z["filename_local"]
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
            name = z["filename_local"]
            existing = _is_already_downloaded(z, folder)
            ok = existing is not None and verify_zip(existing)[0]
            if ok:
                intact.append(name)
                missing = [m for m in missing if m["filename_local"] != name]
                corrupt = [c for c in corrupt if c["filename_local"] != name]
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
        "intact": intact,
        "missing": [z["filename_local"] for z in missing],
        "corrupt": [z["filename_local"] for z in corrupt],
        "failed": failed,
    }


def _save_report(
    folder: Path,
    intact: list[str],
    missing: list[str],
    corrupt: list[str],
    failed: list[str],
) -> None:
    """Save the validation report to a JSON file."""
    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "intact": len(intact),
            "missing": len(missing),
            "corrupt": len(corrupt),
            "failed_after_repair": len(failed),
        },
        "intact": intact,
        "missing": missing,
        "corrupt": corrupt,
        "failed_after_repair": failed,
    }
    path = folder / "validation_report.json"
    atomic_write_json(path, report)
    print(f"\n  Validation report saved to: {path}")


def print_validation_summary(summary: dict[str, Any]) -> None:
    """Print a summary of the validation results."""
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
        print("\n  All files intact.")
    print(SEP)


# ---------------------------------------------------------------------------
# Metadata export
# ---------------------------------------------------------------------------


def save_csv(zips: list[ZipEntry], folder: Path) -> None:
    """Save metadata of discovered ZIPs to a CSV file."""
    path = folder / "terrabrasilis_zips.csv"
    fields = ["biome", "category", "filename", "text", "url"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(zips)
    print(f"  CSV saved to: {path}")


def save_json(zips: list[ZipEntry], folder: Path) -> None:
    """Save metadata of discovered ZIPs to a JSON file."""
    path = folder / "terrabrasilis_zips.json"
    atomic_write_json(
        path,
        {
            "scraped_at": datetime.now(timezone.utc).isoformat(),
            "count": len(zips),
            "files": zips,
        },
    )
    print(f"  JSON saved to: {path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def _count_local_zips() -> int:
    """Count non-empty, non-.tmp .zip files anywhere under ROOT_FOLDER."""
    return sum(len(paths) for paths in _index_existing_zips().values())


def _remote_inventory_quality(zips: list[ZipEntry]) -> dict[str, Any]:
    """Return lightweight data-quality metrics for the remote ZIP inventory."""
    missing_required = [
        z
        for z in zips
        if not z.get("url")
        or not z.get("filename_local")
        or not z.get("biome")
        or not z.get("category")
    ]
    names: dict[str, int] = {}
    logical_keys: dict[tuple[str, str, str], int] = {}
    for z in zips:
        names[z["filename_local"]] = names.get(z["filename_local"], 0) + 1
        key = (z["biome"], z["category"], z["filename_local"])
        logical_keys[key] = logical_keys.get(key, 0) + 1
    duplicate_filenames = sorted(name for name, count in names.items() if count > 1)
    duplicate_logical_keys = [
        {"biome": b, "category": c, "filename": n, "count": count}
        for (b, c, n), count in sorted(logical_keys.items())
        if count > 1
    ]
    return {
        "remote_count": len(zips),
        "contract": to_jsonable(ZIP_INVENTORY_CONTRACT),
        "missing_required_count": len(missing_required),
        "duplicate_filename_count": len(duplicate_filenames),
        "duplicate_filenames": duplicate_filenames,
        "duplicate_logical_keys": duplicate_logical_keys,
        "biome_counts": {
            biome: sum(1 for z in zips if z["biome"] == biome)
            for biome in sorted({z["biome"] for z in zips})
        },
        "category_counts": {
            category: sum(1 for z in zips if z["category"] == category)
            for category in sorted({z["category"] for z in zips})
        },
    }


def _local_zip_quality() -> dict[str, Any]:
    paths = [p for paths in _index_existing_zips().values() for p in paths]
    return {
        "contract": to_jsonable(ZIP_ARCHIVE_CONTRACT),
        "inventory": file_inventory(paths),
        "freshness": freshness_metrics(paths, ZIP_ARCHIVE_CONTRACT.freshness),
    }


def main() -> None:
    """Main function to discover, download, and validate TerraBrasilis ZIP files."""
    ensure_pipeline_dirs()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{SEP}")
    print(f"  TerraBrasilis Download  v{__version__}  |  {now}")
    print(f"{SEP}")

    session = _make_session()
    MAX_PASSES = 20
    RETRY_DELAY = 30
    all_remote_zips: list[ZipEntry] = []

    # ------------------------------------------------------------------ #
    # Step 1: scrape the site                                            #
    # ------------------------------------------------------------------ #
    print(f"\n{SEP}\n  STEP 1 OF 4  --  Scraping TerraBrasilis\n{SEP}")
    scrape_timer = StageTimer("01_scrape_terrabrasilis")

    # Attempt static scrape
    try:
        all_remote_zips = fetch_static(BASE_URL)
    except Exception as exc:
        print(f"  -> Static request failed: {type(exc).__name__}: {exc}")

    if all_remote_zips:
        print(f"  -> {len(all_remote_zips)} ZIP(s) found via static request.")
    else:
        local_zip_count = _count_local_zips()
        print("  -> No ZIPs found via static request (JS rendering required).")
        if local_zip_count:
            print(
                f"  -> {local_zip_count} ZIP(s) already present under {ROOT_FOLDER}.\n"
                "  Attempting Selenium anyway to refresh the remote inventory..."
            )
        else:
            print("  No local files found. Attempting dynamic scraping with Selenium...")

        try:
            all_remote_zips = fetch_dynamic(BASE_URL)
        except Exception as exc:
            print(f"  -> Selenium error: {type(exc).__name__}: {exc}")
        print(f"  -> {len(all_remote_zips)} ZIP(s) found via Selenium.")

    OBS_LOG.emit(
        "stage_metrics",
        **to_jsonable(
            scrape_timer.finish(
                "ok" if all_remote_zips else "degraded",
                input_row_count=None,
                output_row_count=len(all_remote_zips),
            )
        ),
    )

    if not all_remote_zips:
        # If still no remote ZIPs after all scraping attempts
        existing_local_count = _count_local_zips()
        if existing_local_count:
            print(
                f"\n  Scraping failed to discover remote files, but {existing_local_count} "
                f"ZIP(s) already present under {ROOT_FOLDER}.\n"
                "  Remote inventory was not refreshed. Continuing with local ZIPs."
            )
            report_path = write_run_report(
                REPORT_DIR,
                Path(__file__).name,
                {
                    "status": "degraded",
                    "reason": "remote inventory unavailable; local ZIPs present",
                    "root_folder": str(ROOT_FOLDER),
                    "local_zip_count": existing_local_count,
                    "local_zip_quality": _local_zip_quality(),
                    "lineage": LineageRecord(
                        stage_name="01_download_zips",
                        upstream_sources=[BASE_URL],
                        transformation="Remote scrape unavailable; downstream stages will consume existing local ZIP archives.",
                        downstream_outputs=[str(ROOT_FOLDER)],
                        contracts=[ZIP_ARCHIVE_CONTRACT.name],
                    ),
                },
            )
            print(f"  Quality report: {report_path}")
            sys.exit(0)  # Let the pipeline continue with the local ZIP archive set.
        else:
            print(
                "\n  No ZIP files found via scraping and none present locally.\n"
                "  Opening the download page in your default browser for manual download."
            )
            import webbrowser
            webbrowser.open(BASE_URL)
            print(f"  URL: {BASE_URL}\n"
                  f"  Download files manually to: {ROOT_FOLDER}\n"
                  "  Then re-run this script."
            )
            sys.exit(1)  # Exit with error, implies "failed to get files"

    if SKIP_FILES:
        before = len(all_remote_zips)
        all_remote_zips = [z for z in all_remote_zips if z["filename"] not in SKIP_FILES]
        if before - len(all_remote_zips) > 0:
            print(f"  -> {before - len(all_remote_zips)} file(s) excluded by skip_files config.")

    inventory_quality = _remote_inventory_quality(all_remote_zips)
    local_zip_quality = _local_zip_quality()
    OBS_LOG.emit(
        "data_contract",
        stage_name="01_remote_inventory_contract",
        contract=to_jsonable(ZIP_INVENTORY_CONTRACT),
        metrics=inventory_quality,
    )
    OBS_LOG.emit(
        "data_contract",
        stage_name="01_local_zip_contract",
        contract=to_jsonable(ZIP_ARCHIVE_CONTRACT),
        metrics=local_zip_quality,
    )
    if inventory_quality["missing_required_count"]:
        report_path = write_run_report(
            REPORT_DIR,
            Path(__file__).name,
            {
                "status": "failed",
                "reason": "remote inventory has missing required fields",
                "inventory_quality": inventory_quality,
            },
        )
        print(f"  Quality report: {report_path}")
        sys.exit("[FATAL] Remote ZIP inventory has missing required fields.")
    if inventory_quality["duplicate_filename_count"]:
        print(
            "  [quality] Duplicate filenames detected across the remote inventory; "
            "biome/category-aware matching is enabled."
        )

    # ------------------------------------------------------------------ #
    # Step 2: show inventory table                                       #
    # ------------------------------------------------------------------ #
    print(f"\n{SEP}\n  STEP 2 OF 4  --  Inventory check\n{SEP}")
    print(f"  Checking what is already present under: {ROOT_FOLDER}\n")

    inventory_timer = StageTimer("01_inventory_check")
    pending = print_inventory_table(all_remote_zips)
    OBS_LOG.emit(
        "stage_metrics",
        **to_jsonable(
            inventory_timer.finish(
                "ok",
                input_row_count=len(all_remote_zips),
                output_row_count=len(pending),
                anomalies={
                    "schema": [
                        "duplicate filenames require biome/category-aware matching"
                    ]
                    if inventory_quality["duplicate_filename_count"]
                    else []
                },
            )
        ),
    )
    lineage = LineageRecord(
        stage_name="01_download_zips",
        upstream_sources=[BASE_URL],
        transformation="Scrape TerraBrasilis ZIP inventory, compare with local ZIP archive set, optionally download and validate ZIP files.",
        downstream_outputs=[str(ROOT_FOLDER), str(DEST_FOLDER)],
        contracts=[ZIP_INVENTORY_CONTRACT.name, ZIP_ARCHIVE_CONTRACT.name],
    )

    # ------------------------------------------------------------------ #
    # Step 3: ask for confirmation                                       #
    # ------------------------------------------------------------------ #
    print(f"\n{SEP}\n  STEP 3 OF 4  --  Confirmation\n{SEP}")

    if not ask_confirmation(pending):
        print("\n  Exiting without downloading.")
        report_path = write_run_report(
            REPORT_DIR,
            Path(__file__).name,
            {
                "status": "ok",
                "action": "no_download_needed" if not pending else "cancelled",
                "root_folder": str(ROOT_FOLDER),
                "destination_folder": str(DEST_FOLDER),
                "local_zip_count": _count_local_zips(),
                "pending_count": len(pending),
                "inventory_quality": inventory_quality,
                "local_zip_quality": local_zip_quality,
                "lineage": lineage,
            },
        )
        print(f"  Quality report: {report_path}")
        sys.exit(0)

    # Save metadata only after user confirms (creates DEST_FOLDER if not exists)
    DEST_FOLDER.mkdir(parents=True, exist_ok=True)
    save_csv(all_remote_zips, DEST_FOLDER)
    save_json(all_remote_zips, DEST_FOLDER)

    # ------------------------------------------------------------------ #
    # Step 4: download + validate (with retry passes)                    #
    # ------------------------------------------------------------------ #
    print(f"\n{SEP}\n  STEP 4 OF 4  --  Download & Validation\n{SEP}")

    for pass_num in range(1, MAX_PASSES + 1):
        if pass_num > 1:
            print(f"\n{SEP}\n  RETRY PASS {pass_num}/{MAX_PASSES}\n{SEP}")

        # Recompute pending based on current state (some might have been downloaded)
        pending = [
            z for z in all_remote_zips
            if _is_already_downloaded(z, DEST_FOLDER) is None
        ]

        if not pending:
            print(f"  All {len(all_remote_zips)} ZIP(s) already present under {ROOT_FOLDER}. Done.")
            report_path = write_run_report(
                REPORT_DIR,
                Path(__file__).name,
                {
                    "status": "ok",
                    "action": "already_present",
                    "root_folder": str(ROOT_FOLDER),
                    "destination_folder": str(DEST_FOLDER),
                    "local_zip_count": _count_local_zips(),
                    "inventory_quality": inventory_quality,
                    "local_zip_quality": local_zip_quality,
                    "lineage": lineage,
                },
            )
            print(f"  Quality report: {report_path}")
            break

        print(f"  {len(pending)} ZIP(s) to download in this pass.")

        dl_summary = download_all(pending, DEST_FOLDER, session)
        print_download_summary(dl_summary, DEST_FOLDER)

        val_summary = validate_and_repair(all_remote_zips, DEST_FOLDER, session, max_attempts=3)
        print_validation_summary(val_summary)

        remaining_failures = len(val_summary["failed"]) + len(val_summary["missing"])
        if remaining_failures == 0:
            print("\n  All files downloaded and validated successfully. Done.")
            report_path = write_run_report(
                REPORT_DIR,
                Path(__file__).name,
                {
                    "status": "ok",
                    "action": "downloaded_and_validated",
                    "root_folder": str(ROOT_FOLDER),
                    "destination_folder": str(DEST_FOLDER),
                    "local_zip_count": _count_local_zips(),
                    "inventory_quality": inventory_quality,
                    "local_zip_quality": local_zip_quality,
                    "validation_summary": val_summary,
                    "lineage": lineage,
                },
            )
            print(f"  Quality report: {report_path}")
            break

        print(f"\n  {remaining_failures} file(s) still incomplete after pass {pass_num}.")
        if pass_num < MAX_PASSES:
            print(f"  Retrying in {RETRY_DELAY}s...")
            time.sleep(RETRY_DELAY)
    else:
        print(f"\n  Reached {MAX_PASSES} passes. Check validation_report.json for details.")
        sys.exit(1)


if __name__ == "__main__":
    main()

