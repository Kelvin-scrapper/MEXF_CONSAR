"""
MEXF_CONSAR mapper — Mexico CONSAR Pension Fund Statistics

Sources and layouts:

1. md61_investments.xlsx (actual XLSX, WEB_SISTEMA sheet):
   Row 2: period string "Cifras porcentuales al cierre de [month] de [year]"
   Row 4: column headers; col 14 = "TOTAL" (system-wide allocation)
   Rows 5-48: data rows (col1=category group, col2=sub-label, cols 3-14=Siefore values)
   Row 49: TOTAL row (col1="TOTAL", col2=NaN)
   BPA182 and BPAT are deprecated — not present in sheet → always output "NA".
   BONDESG, FONADIN, Otros Gubernamental exist in sheet but are NOT in headers → skipped.

2. cd60_export.xls / cd141_export.xls / cd209_export.xls (HTML disguised as .xls):
   Header row: ['', 'Descripción del Concepto', 'MMM-YYYY', 'MMM-YYYY']
               Last header cell = latest available period (e.g. "Abr-2026" → 2026-04)
   Data rows:  ['', label, prev_value, curr_value]  (4 cells)
               Last cell = most recent period value (full precision float string)

50 series (+ period column) → loaded from headers.json.
"""

import json
import os
import re
import unicodedata

import pandas as pd
from bs4 import BeautifulSoup

if hasattr(__import__("sys").stdout, "reconfigure"):
    __import__("sys").stdout.reconfigure(encoding="utf-8", errors="replace")

with open(os.path.join(os.path.dirname(__file__), "headers.json"), encoding="utf-8") as _f:
    COLUMN_HEADERS = json.load(_f)

# Deprecated series — not in WEB_SISTEMA → always "NA"
_ALWAYS_NA = {
    "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBPA182.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBPAT.ACTUALALLOCATION.NONE.M.1@CONSAR",
}

_MONTH_ES_FULL = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
    "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
    "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12",
}
_MONTH_ES_ABBR = {
    "ene": "01", "feb": "02", "mar": "03", "abr": "04",
    "may": "05", "jun": "06", "jul": "07", "ago": "08",
    "sep": "09", "oct": "10", "nov": "11", "dic": "12",
}

# Category group headers in WEB_SISTEMA — not data series, skip them
_INVEST_SKIP = {
    "privados nacionales", "gubernamental ii", "otros activos i",
    "composicion de las inversiones", "tipo de instrumento",
}

# Normalized label (accents stripped, lowercased) → full code string
# BONDESG, FONADIN, "otros gubernamental" intentionally excluded (not in reference)
_INVEST_LABEL_MAP = {
    "renta variable nacional":             "MEXPENSIONFUNDS.DOMESTICEQUITIES.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "renta variable internacional":        "MEXPENSIONFUNDS.FOREIGNEQUITIES.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "mercancias":                          "MEXPENSIONFUNDS.COMMODITIES.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "aerolineas":                          "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTAEROLINEAS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "alimentos":                           "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTALIMENTOS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "automotriz":                          "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTAUTOMOTRIZ.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "banca de desarrollo":                 "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTBANCADEDESARROLLO.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "bancario":                            "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTBANCARIO.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "bebidas":                             "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTBEBIDAS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "cemento":                             "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCEMENTO.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "centros comerciales":                 "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCENTROSCOMERCIALES.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "consumo":                             "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCONSUMO.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "deuda cp":                            "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTDEUDACP.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "empresas publicas del estado":        "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTEMPRESASPRODUCTIVASDELESTADO.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "empresas productivas del estado":     "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTEMPRESASPRODUCTIVASDELESTADO.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "estados":                             "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTESTADOS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "europesos":                           "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTEUROBONOS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "eurobonos":                           "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTEUROBONOS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "grupos industriales":                 "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTGRUPOSINDUSTRIALES.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "construccion":                        "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCONSTRUCCION.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "infraestructura":                     "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTINFRAESTRUCTURA.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "inmobiliario":                        "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTINMOBILIARIO.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "otros":                               "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTOTROS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "papel":                               "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTPAPEL.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "serv. financieros":                   "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTSERVFINANCIEROS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "siderurgica":                         "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTSIDERURGICA.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "telecom":                             "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTTELECOM.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "transporte":                          "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTTRANSPORTE.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "vivienda":                            "MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTVIVIENDA.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "estructurados":                       "MEXPENSIONFUNDS.STRUCTUREDASSETS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "fibras":                              "MEXPENSIONFUNDS.REIT.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "deuda internacional":                 "MEXPENSIONFUNDS.FOREIGNBONDS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "bond182":                             "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBOND182.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "bondesd":                             "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONDESD.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "bondesf":                             "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONDESF.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "bonos":                               "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONOS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "bpas":                                "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBPAS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "cbic":                                "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSCBIC.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "cetes":                               "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSCETES.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "depbmx":                              "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSDEPBMX.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "udibono":                             "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSUDIBONO.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "ums":                                 "MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSUMS.ACTUALALLOCATION.NONE.M.1@CONSAR",
    "otros activos":                       "MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR",
    "total":                               "MEXPENSIONFUNDS.TOTAL.ACTUALALLOCATION.NONE.M.1@CONSAR",
}

_CD60_LABEL_MAP = {
    "rcv":        "MEXPENSIONFUNDS.RCV.FLOW.BIGPIPE.M.1@CONSAR",
    "rcv imss":   "MEXPENSIONFUNDS.RCVIMSS.FLOW.BIGPIPE.M.1@CONSAR",
    "rcv issste": "MEXPENSIONFUNDS.RCVISSTE.FLOW.BIGPIPE.M.1@CONSAR",
}

_CD141_LABEL_MAP = {
    "retiro de recursos imss":   "MEXPENSIONFUNDS.IMSSWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR",
    "retiro de recursos issste": "MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR",
}

_CD209_LABEL_MAP = {
    "activos netos de las siefores": "MEXPENSIONFUNDS.TOTAL.LEVEL.NONE.M.1@CONSAR",
}


def _norm(s: str) -> str:
    """Strip accents and lowercase for label matching."""
    nfkd = unicodedata.normalize("NFKD", str(s))
    return nfkd.encode("ascii", "ignore").decode("ascii").strip().lower()


def _es_period_full(text: str) -> str | None:
    """'cierre de abril de 2026' → '2026-04'."""
    text = text.lower().strip()
    year_m = re.search(r"(\d{4})", text)
    month_m = re.search(
        r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto"
        r"|septiembre|octubre|noviembre|diciembre)",
        text,
    )
    if year_m and month_m:
        return f"{year_m.group(1)}-{_MONTH_ES_FULL[month_m.group(1)]}"
    return None


def _es_period_abbr(text: str) -> str | None:
    """'Abr-2026' → '2026-04'."""
    m = re.match(r"(\w{3})-(\d{4})", text.strip())
    if m:
        num = _MONTH_ES_ABBR.get(m.group(1).lower())
        if num:
            return f"{m.group(2)}-{num}"
    return None


# ── Investment allocation (md=61 XLSX, WEB_SISTEMA sheet) ────────────────────

def _parse_investments(path: str) -> tuple[str | None, dict]:
    """Parse WEB_SISTEMA sheet. Returns (period, {full_code: float})."""
    df = pd.read_excel(path, sheet_name="WEB_SISTEMA", header=None, engine="openpyxl")

    period_text = str(df.iloc[2, 1]) if pd.notna(df.iloc[2, 1]) else ""
    period = _es_period_full(period_text)

    header_row = df.iloc[3].tolist()
    total_col = next(
        (i for i, h in enumerate(header_row) if isinstance(h, str) and h.strip().upper() == "TOTAL"),
        len(header_row) - 1,
    )

    data: dict = {}
    for idx in range(4, len(df)):
        row = df.iloc[idx]
        col1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        col2 = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        label_raw = col2 if col2 else col1
        if not label_raw:
            continue
        label_norm = _norm(label_raw)
        if label_norm in _INVEST_SKIP:
            continue
        code = _INVEST_LABEL_MAP.get(label_norm)
        if code is None:
            continue
        val = row.iloc[total_col]
        if pd.notna(val):
            data[code] = float(val)

    print(f"[mapper] investments: period={period}, matched={len(data)} series")
    return period, data


# ── Series.aspx HTML-as-XLS export parser ───────────────────────────────────

def _parse_series_export(path: str, label_map: dict, source_label: str) -> tuple[str | None, dict]:
    """
    Parse HTML-as-XLS export from Series.aspx (downloaded via Playwright).
    Header row: ['', 'Descripción del Concepto', 'MMM-YYYY', 'MMM-YYYY']
    Data rows:  ['', label, prev_val, curr_val]  — last cell = latest period.
    Returns (period, {full_code: float}).
    """
    with open(path, "rb") as f:
        content = f.read()
    soup = BeautifulSoup(content, "html.parser")
    tables = soup.find_all("table")
    if not tables:
        raise ValueError(f"No tables found in {path}")

    period = None
    data: dict = {}

    for row in tables[0].find_all("tr"):
        cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
        if len(cells) < 2:
            continue

        # Header row: second cell contains "descripci"
        if "descripci" in cells[1].lower():
            period = _es_period_abbr(cells[-1])
            continue

        if len(cells) != 4 or cells[0] != "":
            continue

        # Strip trailing footnote digits, normalize
        clean = re.sub(r"\d+$", "", cells[1]).strip().lower()
        code = label_map.get(clean)
        if code:
            try:
                data[code] = float(cells[-1])
            except (ValueError, TypeError):
                pass

    print(f"[mapper] {source_label}: period={period}, matched={len(data)} series")
    return period, data


# ── Output row assembly ───────────────────────────────────────────────────────

def _record_to_row(period: str, data: dict) -> list:
    codes = COLUMN_HEADERS["codes"]
    row = [period] + [None] * (len(codes) - 1)
    for i, code in enumerate(codes[1:], start=1):
        if code in _ALWAYS_NA:
            row[i] = "NA"
            continue
        val = data.get(code)
        if val is None:
            continue
        try:
            row[i] = float(val)
        except (ValueError, TypeError):
            row[i] = val
    return row


# ── Public interface ──────────────────────────────────────────────────────────

def map_to_output(data_paths: dict, existing_path: str | None = None) -> pd.DataFrame:
    """
    Parse all downloaded files and merge into existing history.

    data_paths: dict returned by scraper.fetch_data()
    existing_path: path to existing DATA xlsx (rows 0/1 = headers, row 2+ = data)

    Returns DataFrame with 2 header rows + date-sorted data rows.
    """
    existing_rows: dict[str, list] = {}
    if existing_path and os.path.exists(existing_path):
        ex = pd.read_excel(existing_path, header=None)
        for _, r in ex.iloc[2:].iterrows():
            period = str(r.iloc[0]) if pd.notna(r.iloc[0]) else None
            if period:
                existing_rows[period] = list(r)
        print(f"[mapper] Loaded {len(existing_rows)} existing rows from {existing_path}")

    period_inv, inv_data = _parse_investments(data_paths["investments"])
    period_inf, inf_data = _parse_series_export(data_paths["inflows"],  _CD60_LABEL_MAP,  "inflows (cd=60)")
    period_out, out_data = _parse_series_export(data_paths["outflows"], _CD141_LABEL_MAP, "outflows (cd=141)")
    period_ast, ast_data = _parse_series_export(data_paths["assets"],   _CD209_LABEL_MAP, "assets (cd=209)")

    period = period_inv or period_inf or period_ast or period_out
    if period is None:
        raise ValueError("[mapper] Could not determine reporting period from any source")

    combined: dict = {}
    combined.update(ast_data)
    combined.update(inf_data)
    combined.update(out_data)
    combined.update(inv_data)  # investments last (most authoritative)

    new_row = _record_to_row(period, combined)

    if period in existing_rows:
        merged = list(existing_rows[period])
        for i, val in enumerate(new_row):
            if val is not None:
                merged[i] = val
        existing_rows[period] = merged
    else:
        existing_rows[period] = new_row

    sorted_rows = [existing_rows[d] for d in sorted(existing_rows)]
    all_rows = [COLUMN_HEADERS["codes"], COLUMN_HEADERS["descriptions"]] + sorted_rows
    return pd.DataFrame(all_rows)


PAGE_URL = "https://www.consar.gob.mx/gobmx/aplicativo/siset/CuadroInicial.aspx?md=61"


def build_metadata_rows() -> list[dict]:
    """Return metadata dicts for each series — used by main.py for META xlsx."""
    codes = COLUMN_HEADERS["codes"][1:]
    descs = COLUMN_HEADERS["descriptions"][1:]
    rows = []
    for code, desc in zip(codes, descs):
        if ".LEVEL." in code or ".FLOW." in code:
            unit = "millions MXN"
        else:
            unit = "%"
        rows.append({
            "CODE":              code,
            "DESCRIPTION":       desc,
            "FREQUENCY":         "Monthly",
            "UNIT":              unit,
            "SOURCE_NAME":       "CONSAR - Comision Nacional del Sistema de Ahorro para el Retiro",
            "SOURCE_URL":        PAGE_URL,
            "DATASET":           "MEXF_CONSAR",
            "NEXT_RELEASE_DATE": "",
        })
    return rows
