"""
MEXF_CONSAR Scraper — Mexico CONSAR Pension Fund Statistics

Sources:
  md=61:  Investment allocation Excel (Playwright __doPostBack download)
  cd=60:  Inflows Excel export    (Playwright check-all + export button)
  cd=141: Outflows Excel export   (Playwright check-all + export button)
  cd=209: Total assets Excel export (Playwright check-all + export button)

Usage: fetch_data(downloads_dir) → dict of {source: local_path}
"""

import os

BASE_URL  = "https://www.consar.gob.mx/gobmx/aplicativo/siset"
MD61_URL  = f"{BASE_URL}/CuadroInicial.aspx?md=61"
CD60_URL  = f"{BASE_URL}/Series.aspx?cd=60&cdAlt=False"
CD141_URL = f"{BASE_URL}/Series.aspx?cd=141&cdAlt=False"
CD209_URL = f"{BASE_URL}/Series.aspx?cd=209&cdAlt=False"


def _fetch_investments_playwright(downloads_dir: str) -> str:
    """Download investment allocation Excel from md=61 via Playwright."""
    from playwright.sync_api import sync_playwright

    out_path = os.path.join(downloads_dir, "md61_investments.xlsx")

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        ctx = browser.new_context(accept_downloads=True)
        page = ctx.new_page()

        print("[scraper] Navigating to md=61...")
        page.goto(MD61_URL, timeout=60000, wait_until="networkidle")

        with page.expect_download(timeout=30000) as dl_info:
            page.evaluate("__doPostBack('ctl00$ContentPlaceHolder1$lnkBtn_Excel', '')")
        download = dl_info.value
        download.save_as(out_path)
        print(f"[scraper] Investment Excel saved: {out_path}")

        ctx.close()
        browser.close()

    return out_path


def _fetch_series_export(url: str, out_path: str, label: str) -> str:
    """
    Download a Series.aspx Excel export via Playwright.
    Checks all series checkboxes then clicks the export button.
    The downloaded file is HTML-as-XLS (parseable by BeautifulSoup).
    """
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        ctx = browser.new_context(accept_downloads=True)
        page = ctx.new_page()

        print(f"[scraper] Navigating to {label}...")
        page.goto(url, timeout=60000, wait_until="networkidle")

        # onclick="selectAll()" fires when #checkAll is checked — checks all series boxes
        page.check("#checkAll")
        page.wait_for_timeout(300)
        checked = page.evaluate(
            "document.querySelectorAll('input[type=checkbox]:checked').length"
        )
        print(f"[scraper]   {checked} checkboxes checked")

        with page.expect_download(timeout=30000) as dl_info:
            page.click("#ctl00_ContentPlaceHolder1_btn_ExportaSeries")
        download = dl_info.value
        download.save_as(out_path)
        size = os.path.getsize(out_path)
        print(f"[scraper] {label} export saved: {out_path} ({size:,} bytes)")

        ctx.close()
        browser.close()

    return out_path


def fetch_data(downloads_dir: str = "downloads") -> dict:
    """
    Download all CONSAR data sources.
    Returns dict: {'investments': path, 'inflows': path, 'outflows': path, 'assets': path}
    """
    os.makedirs(downloads_dir, exist_ok=True)

    investments_path = _fetch_investments_playwright(downloads_dir)
    inflows_path = _fetch_series_export(
        CD60_URL,
        os.path.join(downloads_dir, "cd60_export.xls"),
        "Inflows (cd=60)"
    )
    outflows_path = _fetch_series_export(
        CD141_URL,
        os.path.join(downloads_dir, "cd141_export.xls"),
        "Outflows (cd=141)"
    )
    assets_path = _fetch_series_export(
        CD209_URL,
        os.path.join(downloads_dir, "cd209_export.xls"),
        "Total assets (cd=209)"
    )

    return {
        "investments": investments_path,
        "inflows":     inflows_path,
        "outflows":    outflows_path,
        "assets":      assets_path,
    }
