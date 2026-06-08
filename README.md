# Klimaat historisch dashboard

FastAPI-dashboard voor historische kasdata uit `Overzicht_alle_dagen.xlsx`.

Referentiekas: **Kas4**.

## Lokaal starten

```bash
pip install -r requirements.txt
uvicorn app:app --reload
```

Open daarna: http://127.0.0.1:8000

## Render

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
uvicorn app:app --host 0.0.0.0 --port $PORT
```

## Data

Plaats het Excelbestand hier:

```text
data/Overzicht_alle_dagen.xlsx
```

Of upload het via de uploadknop in het dashboard.
