
from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import JSONResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
DEFAULT_EXCEL = DATA_DIR / "Overzicht_alle_dagen.xlsx"
PER_DAG_DIR = DATA_DIR / "per_dag"
REFERENCE_KAS = 4

app = FastAPI(title="Klimaat historisch dashboard", version="1.0.0")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


def _to_float(value: Any) -> float | None:
    """Maak van Excel/CSV waarden veilige floats voor JSON/grafieken."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except TypeError:
        pass
    if isinstance(value, str):
        value = value.replace(",", ".").strip()
        if value == "":
            return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return round(number, 2)


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Verwijder lege kolommen/regels.
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")

    # Datumkolom herkennen.
    date_candidates = [c for c in df.columns if c.lower() in {"datum", "date", "dag"}]
    if date_candidates:
        date_col = date_candidates[0]
    else:
        date_col = df.columns[0]
        df = df.rename(columns={date_col: "Datum"})
        date_col = "Datum"

    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df[df[date_col].notna()].sort_values(date_col)
    df["Datum"] = df[date_col].dt.strftime("%Y-%m-%d")

    # Alles wat lijkt op klimaatwaarde numeriek maken.
    for col in df.columns:
        if col == "Datum":
            continue
        if any(key.lower() in col.lower() for key in ["temp", "vd", "rv", "rh", "straling", "wind", "metingen"]):
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", ".", regex=False), errors="coerce")

    return df


def _read_excel(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    sheet = "Alle dagen" if "Alle dagen" in xls.sheet_names else xls.sheet_names[0]
    return _clean_dataframe(pd.read_excel(path, sheet_name=sheet))


def _read_csv_folder(path: Path) -> pd.DataFrame:
    files = sorted(path.glob("**/*.csv"))
    if not files:
        return pd.DataFrame()

    frames: list[pd.DataFrame] = []
    for file in files:
        try:
            frames.append(pd.read_csv(file, sep=None, engine="python"))
        except Exception:
            continue

    if not frames:
        return pd.DataFrame()

    return _clean_dataframe(pd.concat(frames, ignore_index=True))


def load_data() -> tuple[pd.DataFrame, str, str | None]:
    """Laad eerst uploads, daarna standaard Excel, daarna CSV-map."""
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    uploaded_excels = sorted(UPLOAD_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    candidates = uploaded_excels + ([DEFAULT_EXCEL] if DEFAULT_EXCEL.exists() else [])

    for path in candidates:
        try:
            df = _read_excel(path)
            if not df.empty:
                return df, path.name, None
        except Exception as exc:
            return pd.DataFrame(), str(path.name), f"Excelbestand kon niet worden gelezen: {exc}"

    try:
        df = _read_csv_folder(PER_DAG_DIR)
        if not df.empty:
            return df, "data/per_dag/*.csv", None
    except Exception as exc:
        return pd.DataFrame(), "data/per_dag/*.csv", f"CSV-map kon niet worden gelezen: {exc}"

    return pd.DataFrame(), "geen data", "Geen data gevonden. Upload Overzicht_alle_dagen.xlsx of plaats het bestand in data/."


def kas_numbers(df: pd.DataFrame) -> list[int]:
    numbers: set[int] = set()
    for col in df.columns:
        for match in re.finditer(r"(?:Afd|Kas)\s*_?\s*(\d+)", col, flags=re.IGNORECASE):
            numbers.add(int(match.group(1)))
    return sorted(numbers)


def find_col(df: pd.DataFrame, kas: int, contains: list[str]) -> str | None:
    candidates = []
    for col in df.columns:
        low = col.lower().replace(" ", "")
        kas_match = f"afd{kas}" in low or f"kas{kas}" in low or f"afd_{kas}" in low or f"kas_{kas}" in low
        if kas_match and all(item.lower() in low for item in contains):
            candidates.append(col)
    return candidates[0] if candidates else None


def latest_non_null(df: pd.DataFrame, col: str | None) -> float | None:
    if not col or col not in df.columns:
        return None
    series = df[col].dropna()
    if series.empty:
        return None
    return _to_float(series.iloc[-1])


def avg_col(df: pd.DataFrame, col: str | None) -> float | None:
    if not col or col not in df.columns:
        return None
    value = pd.to_numeric(df[col], errors="coerce").mean()
    return _to_float(value)


def max_col(df: pd.DataFrame, col: str | None) -> float | None:
    if not col or col not in df.columns:
        return None
    value = pd.to_numeric(df[col], errors="coerce").max()
    return _to_float(value)


def build_dashboard(df: pd.DataFrame) -> dict[str, Any]:
    if df.empty:
        return {
            "kas_numbers": [],
            "stats": [],
            "critical_days": [],
            "chart": {"labels": [], "vd": [], "temp": []},
            "summary": {},
        }

    numbers = kas_numbers(df)
    if not numbers:
        numbers = [1, 2, 3, 4, 5, 6]

    labels = df["Datum"].astype(str).tolist() if "Datum" in df.columns else [str(i + 1) for i in range(len(df))]

    stats = []
    ref_vd_col = find_col(df, REFERENCE_KAS, ["vd", "max"])
    ref_temp_col = find_col(df, REFERENCE_KAS, ["temp", "max"])
    ref_vd_avg = avg_col(df, ref_vd_col)
    ref_temp_avg = avg_col(df, ref_temp_col)

    vd_chart = []
    temp_chart = []

    for kas in numbers:
        vd_max_col = find_col(df, kas, ["vd", "max"])
        vd_gem_col = find_col(df, kas, ["vd", "gem"])
        temp_max_col = find_col(df, kas, ["temp", "max"])

        kas_vd_avg = avg_col(df, vd_max_col)
        kas_temp_avg = avg_col(df, temp_max_col)

        stats.append({
            "kas": kas,
            "is_reference": kas == REFERENCE_KAS,
            "temp_max_avg": kas_temp_avg,
            "temp_max_highest": max_col(df, temp_max_col),
            "vd_max_avg": kas_vd_avg,
            "vd_max_highest": max_col(df, vd_max_col),
            "vd_gem_avg": avg_col(df, vd_gem_col),
            "delta_temp_vs_ref": _to_float(kas_temp_avg - ref_temp_avg) if kas_temp_avg is not None and ref_temp_avg is not None else None,
            "delta_vd_vs_ref": _to_float(kas_vd_avg - ref_vd_avg) if kas_vd_avg is not None and ref_vd_avg is not None else None,
        })

        if vd_max_col:
            vd_chart.append({
                "label": f"Kas {kas} VD max",
                "data": [_to_float(v) for v in df[vd_max_col].tolist()],
                "reference": kas == REFERENCE_KAS,
            })
        if temp_max_col:
            temp_chart.append({
                "label": f"Kas {kas} Tmax",
                "data": [_to_float(v) for v in df[temp_max_col].tolist()],
                "reference": kas == REFERENCE_KAS,
            })

    # Kritische dagen: hoogste VD over beschikbare kassen.
    vd_cols = [find_col(df, kas, ["vd", "max"]) for kas in numbers]
    vd_cols = [c for c in vd_cols if c]
    critical_days = []
    if vd_cols:
        tmp = df[["Datum"] + vd_cols].copy()
        tmp["VD_piek"] = tmp[vd_cols].max(axis=1)
        tmp = tmp.sort_values("VD_piek", ascending=False).head(10)
        for _, row in tmp.iterrows():
            critical_days.append({
                "datum": str(row["Datum"]),
                "vd_piek": _to_float(row["VD_piek"]),
            })

    warmer = [s for s in stats if s["kas"] != REFERENCE_KAS and s["delta_temp_vs_ref"] is not None and s["delta_temp_vs_ref"] > 0.75]
    droger = [s for s in stats if s["kas"] != REFERENCE_KAS and s["delta_vd_vs_ref"] is not None and s["delta_vd_vs_ref"] > 1.5]

    advice = []
    if warmer:
        advice.append("Kassen " + ", ".join(str(s["kas"]) for s in warmer) + " zijn gemiddeld duidelijk warmer dan referentie Kas4.")
    if droger:
        advice.append("Kassen " + ", ".join(str(s["kas"]) for s in droger) + " zijn gemiddeld duidelijk droger/hoger in VD dan Kas4.")
    if not advice:
        advice.append("Geen grote structurele afwijking t.o.v. Kas4 gevonden in de geladen data.")

    return {
        "kas_numbers": numbers,
        "stats": stats,
        "critical_days": critical_days,
        "chart": {"labels": labels, "vd": vd_chart, "temp": temp_chart},
        "summary": {
            "days": len(df),
            "start": labels[0] if labels else None,
            "end": labels[-1] if labels else None,
            "reference_kas": REFERENCE_KAS,
            "advice": advice,
        },
    }


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/api/data")
def api_data() -> JSONResponse:
    df, source, error = load_data()
    payload = build_dashboard(df)
    payload["source"] = source
    payload["error"] = error
    return JSONResponse(payload)


@app.get("/")
def dashboard(request: Request):
    df, source, error = load_data()
    payload = build_dashboard(df)
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "source": source,
            "error": error,
            "dashboard": payload,
        },
    )


@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        return JSONResponse({"error": "Upload een .xlsx bestand."}, status_code=400)

    target = UPLOAD_DIR / "Overzicht_alle_dagen_uploaded.xlsx"
    content = await file.read()
    target.write_bytes(content)
    return RedirectResponse(url="/", status_code=303)
