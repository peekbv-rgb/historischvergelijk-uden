from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, Query, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
DEFAULT_EXCEL = DATA_DIR / "Overzicht_alle_dagen.xlsx"
PER_DAY_DIR = DATA_DIR / "per_dag"

DATA_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Klimaat Historisch Dashboard")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


# ============================================================
# Hulpfuncties
# ============================================================

def _clean_col(col: Any) -> str:
    """Maak kolomnamen voorspelbaar zonder de betekenis te wijzigen."""
    return str(col).strip().replace("\ufeff", "")


def _to_float_series(series: pd.Series) -> pd.Series:
    """Converteer getallen robuust, ook bij decimale komma's."""
    if series.dtype == object:
        series = series.astype(str).str.replace(",", ".", regex=False)
    return pd.to_numeric(series, errors="coerce")


def _find_first_column(df: pd.DataFrame, patterns: list[str]) -> str | None:
    for col in df.columns:
        low = col.lower()
        if all(p.lower() in low for p in patterns):
            return col
    return None


def _json_safe(value: Any) -> Any:
    """Maak waarden geschikt voor Jinja/JSON: geen NaN en geen numpy-scalars."""
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "item"):
        try:
            value = value.item()
        except Exception:
            pass
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        return round(value, 3)
    return value


def _row_to_dict(row: pd.Series, cols: list[str]) -> dict[str, Any]:
    return {col: _json_safe(row.get(col)) for col in cols}


def _detect_afdelingen(df: pd.DataFrame) -> list[int]:
    ids: set[int] = set()
    for col in df.columns:
        match = re.search(r"Afd\s*([0-9]+)", col, flags=re.IGNORECASE)
        if match:
            ids.add(int(match.group(1)))
    return sorted(ids)


def _normalize_markering(row: pd.Series) -> str:
    raw_value = row.get("Markering", "")
    if not pd.isna(raw_value):
        value = str(raw_value).strip().upper()
        if value and value != "NAN":
            return value
    t = row.get("Temp_buiten_max")
    try:
        t = float(t)
    except Exception:
        return ""
    if t >= 28:
        return "HEET"
    if t >= 22:
        return "WARM"
    if t <= 10:
        return "KOEL"
    return "NORMAAL"


# ============================================================
# Data inlezen
# ============================================================

def load_excel(path: Path = DEFAULT_EXCEL) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Excelbestand niet gevonden: {path}")

    # Het bestand uit de analyse gebruikt tabblad 'Alle dagen'.
    # Als dat tabblad ontbreekt, pakken we het eerste tabblad.
    try:
        df = pd.read_excel(path, sheet_name="Alle dagen")
    except ValueError:
        df = pd.read_excel(path, sheet_name=0)

    df.columns = [_clean_col(c) for c in df.columns]
    return normalize_history_df(df, source=f"Excel: {path.name}")


def _read_csv_flexible(path: Path) -> pd.DataFrame:
    """Lees CSV met ; of , separator en zowel punt als komma decimalen."""
    attempts = [
        {"sep": ";", "decimal": ","},
        {"sep": ";", "decimal": "."},
        {"sep": ",", "decimal": "."},
        {"sep": None, "engine": "python"},
    ]
    last_error: Exception | None = None
    for kwargs in attempts:
        try:
            df = pd.read_csv(path, **kwargs)
            if len(df.columns) > 1:
                df.columns = [_clean_col(c) for c in df.columns]
                return df
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"CSV kon niet worden gelezen: {path} ({last_error})")


def load_per_day_csvs(folder: Path = PER_DAY_DIR) -> pd.DataFrame:
    files = sorted(folder.glob("**/*.csv"))
    if not files:
        raise FileNotFoundError(f"Geen CSV-bestanden gevonden in {folder}")

    rows: list[dict[str, Any]] = []
    already_aggregated: list[pd.DataFrame] = []

    for file in files:
        raw = _read_csv_flexible(file)
        raw.columns = [_clean_col(c) for c in raw.columns]

        # Als het CSV-bestand al dagregels bevat met dezelfde kolommen, direct meenemen.
        if any(c.startswith("Temp_Afd") and c.endswith("_max") for c in raw.columns):
            already_aggregated.append(raw)
            continue

        # Anders: ruwe meetdata per dag aggregeren.
        rows.append(aggregate_raw_csv_to_day(raw, fallback_date=file.stem))

    if already_aggregated:
        df = pd.concat(already_aggregated, ignore_index=True)
    else:
        df = pd.DataFrame(rows)

    return normalize_history_df(df, source=f"CSV-map: {folder}")


def aggregate_raw_csv_to_day(raw: pd.DataFrame, fallback_date: str) -> dict[str, Any]:
    row: dict[str, Any] = {"Datum": fallback_date, "N_metingen": len(raw)}

    # Datum uit kolom heeft voorkeur boven bestandsnaam.
    date_col = _find_first_column(raw, ["datum"]) or _find_first_column(raw, ["date"])
    if date_col is not None:
        dates = pd.to_datetime(raw[date_col], errors="coerce")
        if dates.notna().any():
            row["Datum"] = dates.dropna().iloc[0].strftime("%Y-%m-%d")

    # Buitenklimaat.
    temp_buiten = (
        _find_first_column(raw, ["temp", "buiten"])
        or _find_first_column(raw, ["temperature", "outside"])
        or _find_first_column(raw, ["t_out"])
    )
    if temp_buiten:
        s = _to_float_series(raw[temp_buiten])
        row["Temp_buiten_max"] = s.max()
        row["Temp_buiten_gem"] = s.mean()

    straling = _find_first_column(raw, ["straling"]) or _find_first_column(raw, ["radiation"])
    if straling:
        row["Straling_piek"] = _to_float_series(raw[straling]).max()

    wind = _find_first_column(raw, ["wind", "ms"]) or _find_first_column(raw, ["windspeed"])
    if wind:
        row["Wind_gem_ms"] = _to_float_series(raw[wind]).mean()

    windrichting = _find_first_column(raw, ["windrichting"]) or _find_first_column(raw, ["wind", "dir"])
    if windrichting:
        row["Windrichting_gem"] = _to_float_series(raw[windrichting]).mean()

    # Afdelingkolommen herkennen: Afd1/Afdeling 1/Department 1 met temp of VD/VPD.
    for afd in range(1, 21):
        afd_regex = re.compile(rf"(afd|afdeling|department)[_\s-]*{afd}\b", re.IGNORECASE)
        afd_cols = [c for c in raw.columns if afd_regex.search(c)]
        if not afd_cols:
            continue

        temp_cols = [c for c in afd_cols if re.search(r"temp|temperature|\bt\b", c, re.IGNORECASE)]
        vd_cols = [c for c in afd_cols if re.search(r"\bvd\b|vpd", c, re.IGNORECASE)]

        if temp_cols:
            s = _to_float_series(raw[temp_cols[0]])
            row[f"Temp_Afd{afd}_max"] = s.max()
        if vd_cols:
            s = _to_float_series(raw[vd_cols[0]])
            row[f"VD_Afd{afd}_max"] = s.max()
            row[f"VD_Afd{afd}_gem"] = s.mean()

    return row


def normalize_history_df(df: pd.DataFrame, source: str) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_clean_col(c) for c in df.columns]

    if "Datum" not in df.columns:
        maybe_date = _find_first_column(df, ["datum"]) or _find_first_column(df, ["date"])
        if maybe_date:
            df = df.rename(columns={maybe_date: "Datum"})

    if "Datum" not in df.columns:
        raise ValueError("Kolom 'Datum' ontbreekt. Kan historische data niet op tijdlijn zetten.")

    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
    df = df[df["Datum"].notna()].copy()
    df = df.sort_values("Datum")

    for col in df.columns:
        if col == "Datum" or col == "Markering":
            continue
        df[col] = _to_float_series(df[col])

    if "Markering" not in df.columns:
        df["Markering"] = ""
    df["Markering"] = df.apply(_normalize_markering, axis=1)
    df["Bron"] = source

    return df.reset_index(drop=True)


def get_history(source: str = "auto", uploaded_filename: str | None = None) -> pd.DataFrame:
    if uploaded_filename:
        path = UPLOAD_DIR / uploaded_filename
        return load_excel(path)

    if source == "excel":
        return load_excel(DEFAULT_EXCEL)
    if source == "csv":
        return load_per_day_csvs(PER_DAY_DIR)

    # Auto: eerst Excel, daarna CSV-map.
    if DEFAULT_EXCEL.exists():
        return load_excel(DEFAULT_EXCEL)
    return load_per_day_csvs(PER_DAY_DIR)


# ============================================================
# Analyse
# ============================================================

def build_analysis(df: pd.DataFrame) -> dict[str, Any]:
    afdelingen = _detect_afdelingen(df)
    if not afdelingen:
        raise ValueError("Geen afdelingskolommen gevonden, bijvoorbeeld Temp_Afd1_max of VD_Afd1_max.")

    vd_max_cols = [f"VD_Afd{a}_max" for a in afdelingen if f"VD_Afd{a}_max" in df.columns]
    vd_gem_cols = [f"VD_Afd{a}_gem" for a in afdelingen if f"VD_Afd{a}_gem" in df.columns]
    temp_max_cols = [f"Temp_Afd{a}_max" for a in afdelingen if f"Temp_Afd{a}_max" in df.columns]

    df = df.copy()
    if vd_max_cols:
        df["VD_max_all"] = df[vd_max_cols].max(axis=1)
        df["VD_max_afd"] = df[vd_max_cols].idxmax(axis=1).str.extract(r"Afd(\d+)")[0]
    else:
        df["VD_max_all"] = None
        df["VD_max_afd"] = None

    if temp_max_cols:
        df["Temp_max_all"] = df[temp_max_cols].max(axis=1)
        df["Temp_max_afd"] = df[temp_max_cols].idxmax(axis=1).str.extract(r"Afd(\d+)")[0]
    else:
        df["Temp_max_all"] = None
        df["Temp_max_afd"] = None

    day_count = len(df)
    date_min = df["Datum"].min().strftime("%Y-%m-%d") if day_count else ""
    date_max = df["Datum"].max().strftime("%Y-%m-%d") if day_count else ""

    max_measurements = int(df["N_metingen"].max()) if "N_metingen" in df.columns and df["N_metingen"].notna().any() else None
    incomplete_days = 0
    if max_measurements:
        incomplete_days = int((df["N_metingen"] < max_measurements).sum())

    markers = df["Markering"].value_counts(dropna=False).to_dict() if "Markering" in df.columns else {}

    department_summary = []
    for afd in afdelingen:
        department_summary.append(
            {
                "afd": afd,
                "temp_max_avg": _json_safe(df.get(f"Temp_Afd{afd}_max", pd.Series(dtype=float)).mean()),
                "temp_max_high": _json_safe(df.get(f"Temp_Afd{afd}_max", pd.Series(dtype=float)).max()),
                "vd_max_avg": _json_safe(df.get(f"VD_Afd{afd}_max", pd.Series(dtype=float)).mean()),
                "vd_max_high": _json_safe(df.get(f"VD_Afd{afd}_max", pd.Series(dtype=float)).max()),
                "vd_gem_avg": _json_safe(df.get(f"VD_Afd{afd}_gem", pd.Series(dtype=float)).mean()),
            }
        )

    warm_hot = df[df["Markering"].isin(["WARM", "HEET"])] if "Markering" in df.columns else df.iloc[0:0]
    group_compare = []
    groups = {
        "Afd 1-3": [1, 2, 3],
        "Afd 5-6": [5, 6],
    }
    for label, ids in groups.items():
        group_temp_cols = [f"Temp_Afd{i}_max" for i in ids if f"Temp_Afd{i}_max" in df.columns]
        group_vd_cols = [f"VD_Afd{i}_max" for i in ids if f"VD_Afd{i}_max" in df.columns]
        if len(warm_hot) and group_temp_cols and group_vd_cols:
            group_compare.append(
                {
                    "groep": label,
                    "temp_max_avg_warm_hot": _json_safe(warm_hot[group_temp_cols].mean(axis=1).mean()),
                    "vd_max_avg_warm_hot": _json_safe(warm_hot[group_vd_cols].mean(axis=1).mean()),
                }
            )

    critical_cols = ["Datum", "Markering", "Temp_buiten_max", "Straling_piek", "N_metingen", "VD_max_all", "VD_max_afd", "Temp_max_all", "Temp_max_afd"]
    critical_cols = [c for c in critical_cols if c in df.columns]
    critical_days = []
    if "VD_max_all" in df.columns:
        critical_days = [
            _row_to_dict(row, critical_cols)
            for _, row in df.sort_values("VD_max_all", ascending=False).head(10).iterrows()
        ]

    labels = [d.strftime("%d-%m") for d in df["Datum"]]
    chart_vd = {
        f"Afd {afd}": [_json_safe(v) for v in df.get(f"VD_Afd{afd}_max", pd.Series([None] * len(df))).tolist()]
        for afd in afdelingen
        if f"VD_Afd{afd}_max" in df.columns
    }
    chart_temp = {
        f"Afd {afd}": [_json_safe(v) for v in df.get(f"Temp_Afd{afd}_max", pd.Series([None] * len(df))).tolist()]
        for afd in afdelingen
        if f"Temp_Afd{afd}_max" in df.columns
    }
    if "Temp_buiten_max" in df.columns:
        chart_temp["Buiten"] = [_json_safe(v) for v in df["Temp_buiten_max"].tolist()]

    table_cols = [
        "Datum",
        "Markering",
        "N_metingen",
        "Temp_buiten_max",
        "Straling_piek",
        "VD_max_all",
        "VD_max_afd",
        "Temp_max_all",
        "Temp_max_afd",
    ]
    table_cols = [c for c in table_cols if c in df.columns]
    table_rows = [_row_to_dict(row, table_cols) for _, row in df.tail(20).iloc[::-1].iterrows()]

    advice = make_advice(department_summary, group_compare, incomplete_days, max_measurements)

    return {
        "source": str(df["Bron"].iloc[0]) if "Bron" in df.columns and len(df) else "",
        "day_count": day_count,
        "date_min": date_min,
        "date_max": date_max,
        "max_measurements": max_measurements,
        "incomplete_days": incomplete_days,
        "markers": markers,
        "afdelingen": afdelingen,
        "department_summary": department_summary,
        "group_compare": group_compare,
        "critical_days": critical_days,
        "labels": labels,
        "chart_vd": chart_vd,
        "chart_temp": chart_temp,
        "table_cols": table_cols,
        "table_rows": table_rows,
        "advice": advice,
    }


def make_advice(
    department_summary: list[dict[str, Any]],
    group_compare: list[dict[str, Any]],
    incomplete_days: int,
    max_measurements: int | None,
) -> list[str]:
    advice: list[str] = []

    if department_summary:
        worst_vd = max(department_summary, key=lambda x: x.get("vd_max_high") or -999)
        worst_avg = max(department_summary, key=lambda x: x.get("vd_max_avg") or -999)
        advice.append(
            f"Afdeling {worst_vd['afd']} heeft de hoogste VD-piek ({worst_vd['vd_max_high']}). Controleer daar eerst sensorpositie, bevochtiging/koeling en luchtverdeling."
        )
        if worst_avg["afd"] != worst_vd["afd"]:
            advice.append(
                f"Gemiddeld over de periode is afdeling {worst_avg['afd']} het meest kritisch op VD-max ({worst_avg['vd_max_avg']})."
            )

    if len(group_compare) == 2:
        g1, g2 = group_compare
        temp_diff = (g1.get("temp_max_avg_warm_hot") or 0) - (g2.get("temp_max_avg_warm_hot") or 0)
        vd_diff = (g1.get("vd_max_avg_warm_hot") or 0) - (g2.get("vd_max_avg_warm_hot") or 0)
        if temp_diff > 1.5 or vd_diff > 2:
            advice.append(
                f"Op warme/hete dagen lopen {g1['groep']} duidelijk zwaarder dan {g2['groep']}: circa {round(temp_diff, 1)} °C warmer en {round(vd_diff, 1)} VD hoger. Stuur dus niet op kasgemiddelde maar per afdeling."
            )

    if incomplete_days:
        advice.append(
            f"Er zijn {incomplete_days} onvolledige dag(en). Gebruik piekwaarden wel, maar wees voorzichtig met daggemiddelden. Volledige dag lijkt {max_measurements} metingen te hebben."
        )

    advice.append("Aanbevolen vervolgstap: maak alarmregels per afdeling, bijvoorbeeld VD_max > 20 of temperatuurverschil afdeling 1-3 versus 5-6 > 3 °C.")
    return advice


# ============================================================
# Routes
# ============================================================

@app.get("/", response_class=HTMLResponse)
def dashboard(
    request: Request,
    source: str = Query("auto", pattern="^(auto|excel|csv)$"),
    upload: str | None = None,
):
    try:
        df = get_history(source=source, uploaded_filename=upload)
        analysis = build_analysis(df)
        error = None
    except Exception as exc:
        analysis = None
        error = str(exc)

    uploads = sorted([p.name for p in UPLOAD_DIR.glob("*.xlsx")], reverse=True)

    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "analysis": analysis,
            "error": error,
            "source": source,
            "uploads": uploads,
            "selected_upload": upload,
        },
    )


@app.get("/api/data")
def api_data(source: str = Query("auto", pattern="^(auto|excel|csv)$"), upload: str | None = None):
    try:
        df = get_history(source=source, uploaded_filename=upload)
        return JSONResponse(build_analysis(df))
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        return JSONResponse({"error": "Upload alleen .xlsx bestanden."}, status_code=400)

    safe_name = re.sub(r"[^a-zA-Z0-9_.-]", "_", file.filename)
    target = UPLOAD_DIR / safe_name
    target.write_bytes(await file.read())
    return RedirectResponse(url=f"/?upload={safe_name}", status_code=303)


@app.get("/health")
def health():
    return {"status": "ok", "default_excel_exists": DEFAULT_EXCEL.exists(), "csv_folder_exists": PER_DAY_DIR.exists()}
