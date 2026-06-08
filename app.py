from fastapi import FastAPI
from fastapi.responses import HTMLResponse, JSONResponse
from pathlib import Path
from datetime import datetime
import csv
import html
import io
import urllib.request

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

app = FastAPI(title="Klimaat Dashboard Kas4")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
EXCEL_FILE = DATA_DIR / "Overzicht_alle_dagen.xlsx"
CSV_DIR = DATA_DIR / "per_dag"
CSV_FILE = CSV_DIR / "Overzicht_alle_dagen.csv"
REMOTE_CSV_URL = "https://raw.githubusercontent.com/peekbv-rgb/historischvergelijk-uden/main/data/per_dag/Overzicht_alle_dagen.csv"
REFERENCE_KAS = 4
KASSEN = [1, 2, 3, 4, 5, 6]


def to_float(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def fmt(value, digits=1):
    number = to_float(value)
    if number is None:
        return "-"
    return f"{number:.{digits}f}"


def avg(values):
    numbers = [v for v in values if v is not None]
    return sum(numbers) / len(numbers) if numbers else None


def max_or_none(values):
    numbers = [v for v in values if v is not None]
    return max(numbers) if numbers else None


def read_excel_rows():
    if load_workbook is None or not EXCEL_FILE.exists():
        return []
    workbook = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    sheet_name = "Alle dagen" if "Alle dagen" in workbook.sheetnames else workbook.sheetnames[0]
    worksheet = workbook[sheet_name]
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    output = []
    for raw_row in rows[1:]:
        item = {}
        for key, value in zip(headers, raw_row):
            if key:
                item[key] = value
        if any(value is not None for value in item.values()):
            output.append(item)
    return output


def read_local_csv_rows():
    output = []
    if CSV_FILE.exists():
        paths = [CSV_FILE]
    elif CSV_DIR.exists():
        paths = sorted(CSV_DIR.rglob("*.csv"))
    else:
        paths = []

    for path in paths:
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                reader = csv.DictReader(handle)
                output.extend(dict(row) for row in reader)
        except Exception:
            continue
    return output


def read_remote_csv_rows():
    try:
        with urllib.request.urlopen(REMOTE_CSV_URL, timeout=10) as response:
            text = response.read().decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(text))
        return [dict(row) for row in reader]
    except Exception:
        return []


def filesystem_debug():
    local_csvs = []
    if CSV_DIR.exists():
        local_csvs = [str(p.relative_to(BASE_DIR)) for p in sorted(CSV_DIR.rglob("*.csv"))]
    return {
        "base_dir": str(BASE_DIR),
        "excel_exists": EXCEL_FILE.exists(),
        "csv_dir_exists": CSV_DIR.exists(),
        "csv_file_exists": CSV_FILE.exists(),
        "local_csvs": local_csvs,
        "remote_csv_url": REMOTE_CSV_URL,
    }


def load_rows():
    rows = read_excel_rows()
    if rows:
        return rows, "data/Overzicht_alle_dagen.xlsx"

    rows = read_local_csv_rows()
    if rows:
        return rows, "lokale CSV: data/per_dag/Overzicht_alle_dagen.csv"

    rows = read_remote_csv_rows()
    if rows:
        return rows, "GitHub CSV fallback"

    return [], "geen databestand gevonden"


def get_date(row):
    for key in ("Datum", "datum", "Date", "date"):
        value = row.get(key)
        if value is None:
            continue
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d")
        return str(value)[:10]
    return "-"


def get_metric(row, kas, metric):
    keys = [
        f"{metric}_Afd{kas}",
        f"{metric}_Kas{kas}",
        f"Afd{kas}_{metric}",
        f"Kas{kas}_{metric}",
    ]
    for key in keys:
        if key in row:
            return to_float(row.get(key))
    return None


def get_tmax(row, kas):
    for key in (f"Temp_Afd{kas}_max", f"Temp_Kas{kas}_max", f"Tmax_Afd{kas}", f"Tmax_Kas{kas}"):
        if key in row:
            return to_float(row.get(key))
    return get_metric(row, kas, "Temp_max")


def get_vd_max(row, kas):
    for key in (f"VD_Afd{kas}_max", f"VD_Kas{kas}_max", f"VDmax_Afd{kas}", f"VDmax_Kas{kas}"):
        if key in row:
            return to_float(row.get(key))
    return get_metric(row, kas, "VD_max")


def get_vd_avg(row, kas):
    for key in (f"VD_Afd{kas}_gem", f"VD_Kas{kas}_gem", f"VD_Afd{kas}_avg", f"VD_Kas{kas}_avg"):
        if key in row:
            return to_float(row.get(key))
    return get_metric(row, kas, "VD_gem")


def summarize(rows):
    summary = []
    for kas in KASSEN:
        tmax_values = [get_tmax(row, kas) for row in rows]
        vdmax_values = [get_vd_max(row, kas) for row in rows]
        vdavg_values = [get_vd_avg(row, kas) for row in rows]
        summary.append({
            "kas": kas,
            "tmax_avg": avg(tmax_values),
            "tmax_max": max_or_none(tmax_values),
            "vdmax_avg": avg(vdmax_values),
            "vdmax_max": max_or_none(vdmax_values),
            "vdavg_avg": avg(vdavg_values),
        })
    return summary


def build_analysis():
    rows, source = load_rows()
    summary = summarize(rows)
    reference = next((item for item in summary if item["kas"] == REFERENCE_KAS), None)
    differences = []
    if reference:
        for item in summary:
            differences.append({
                "kas": item["kas"],
                "dtmax": None if item["tmax_avg"] is None or reference["tmax_avg"] is None else item["tmax_avg"] - reference["tmax_avg"],
                "dvdmax": None if item["vdmax_avg"] is None or reference["vdmax_avg"] is None else item["vdmax_avg"] - reference["vdmax_avg"],
                "dvdavg": None if item["vdavg_avg"] is None or reference["vdavg_avg"] is None else item["vdavg_avg"] - reference["vdavg_avg"],
            })

    critical = []
    for row in rows:
        values = []
        for kas in KASSEN:
            vd = get_vd_max(row, kas)
            if vd is not None:
                values.append((kas, vd))
        if values:
            kas, vd = max(values, key=lambda item: item[1])
            critical.append({"date": get_date(row), "kas": kas, "vd": vd})
    critical = sorted(critical, key=lambda item: item["vd"], reverse=True)[:10]

    return {
        "status": "ok",
        "reference_kas": REFERENCE_KAS,
        "rows_count": len(rows),
        "source": source,
        "debug": filesystem_debug(),
        "summary": summary,
        "differences": differences,
        "critical": critical,
    }


def cell(value):
    return f"<td>{html.escape(str(value))}</td>"


def render_table(headers, rows):
    header_html = "".join(f"<th>{html.escape(str(header))}</th>" for header in headers)
    body_html = "".join("<tr>" + "".join(cell(c) for c in row) + "</tr>" for row in rows)
    return f"<table><tr>{header_html}</tr>{body_html}</table>"


def render_page(data):
    warning = ""
    if data["rows_count"] == 0:
        debug = data.get("debug", {})
        warning = f"""
        <div class="warning">
            <b>Geen historische data gevonden.</b><br>
            De app draait goed, maar ziet lokaal nog geen databestand.<br>
            Gecontroleerd: Excel={debug.get('excel_exists')}, CSV-map={debug.get('csv_dir_exists')}, CSV-bestand={debug.get('csv_file_exists')}.<br>
            Remote fallback: {html.escape(str(debug.get('remote_csv_url')))}
        </div>
        """

    summary_rows = [[f"Kas {i['kas']}", f"{fmt(i['tmax_avg'])} °C", f"{fmt(i['tmax_max'])} °C", fmt(i["vdmax_avg"]), fmt(i["vdmax_max"]), fmt(i["vdavg_avg"])] for i in data["summary"]]
    diff_rows = [[f"Kas {i['kas']}", f"{fmt(i['dtmax'])} °C", fmt(i["dvdmax"]), fmt(i["dvdavg"])] for i in data["differences"]]
    critical_rows = [[i["date"], f"Kas {i['kas']}", fmt(i["vd"])] for i in data["critical"]]

    return f"""<!doctype html>
<html lang="nl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Klimaat Dashboard Kas4</title>
  <style>
    body {{ margin:0; font-family:Arial,sans-serif; background:#eef3f0; color:#13251d; }}
    header {{ background:#123c2c; color:white; padding:28px 36px; }}
    main {{ max-width:1200px; margin:auto; padding:28px 36px; }}
    .cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:16px; margin-bottom:26px; }}
    .card {{ background:white; border-radius:14px; padding:18px; box-shadow:0 8px 20px rgba(0,0,0,.08); }}
    .value {{ font-size:28px; font-weight:700; margin-top:8px; }}
    .ok {{ color:#16803c; }}
    .warning {{ background:#fff3cd; border:1px solid #e2c46d; padding:14px 16px; border-radius:12px; margin-bottom:18px; line-height:1.55; }}
    table {{ width:100%; border-collapse:collapse; background:white; border-radius:14px; overflow:hidden; margin:18px 0 32px; box-shadow:0 8px 20px rgba(0,0,0,.06); }}
    th, td {{ padding:10px 12px; border-bottom:1px solid #e3e8e4; text-align:left; }}
    th {{ background:#dfeae3; }}
  </style>
</head>
<body>
  <header><h1>Klimaat Dashboard</h1><div>Historische analyse met <b>Kas4</b> als referentiekas</div></header>
  <main>
    {warning}
    <div class="cards">
      <div class="card"><div>Status</div><div class="value ok">Online</div></div>
      <div class="card"><div>Referentie</div><div class="value">Kas4</div></div>
      <div class="card"><div>Datapunten</div><div class="value">{data['rows_count']}</div></div>
      <div class="card"><div>Bron</div><div>{html.escape(data['source'])}</div></div>
    </div>
    <h2>Samenvatting per kas</h2>
    {render_table(["Kas", "Gem. Tmax", "Hoogste Tmax", "Gem. VD max", "Hoogste VD max", "Gem. VD gemiddeld"], summary_rows)}
    <h2>Afwijking t.o.v. Kas4</h2>
    {render_table(["Kas", "Δ gem. Tmax", "Δ gem. VD max", "Δ gem. VD gemiddeld"], diff_rows)}
    <h2>Top 10 kritische VD-pieken</h2>
    {render_table(["Datum", "Kas", "VD max"], critical_rows)}
  </main>
</body>
</html>"""


@app.get("/health")
def health():
    return {"status": "ok", "reference_kas": REFERENCE_KAS}


@app.get("/debug/files")
def debug_files():
    rows, source = load_rows()
    info = filesystem_debug()
    info["rows_count"] = len(rows)
    info["source"] = source
    return JSONResponse(info)


@app.get("/api/data")
def api_data():
    return JSONResponse(build_analysis())


@app.get("/", response_class=HTMLResponse)
def dashboard():
    try:
        return HTMLResponse(render_page(build_analysis()))
    except Exception as exc:
        safe = html.escape(str(exc))
        return HTMLResponse(f"<h1>Klimaat Dashboard</h1><p>App draait, maar analyse gaf een fout:</p><pre>{safe}</pre>", status_code=200)
