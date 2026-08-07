from __future__ import annotations

import contextlib
import io
import os
import re
import secrets
import time
from datetime import datetime
from threading import Lock
from typing import Iterable, Literal
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from fastapi import FastAPI, Header, HTTPException, Query
from fastapi.responses import Response, StreamingResponse
from mcp.server.fastmcp import FastMCP
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

APP_VERSION = "2.0.0"
SEACE_BASE = (
    "https://prod4.seace.gob.pe:8086/api/oportunidades/"
    "codObjeto/codDepartamento/sintesisProceso/codTipoProceso"
)
OBJECTS = {
    "servicios": {"code": "65", "label": "Servicio", "filename": "Servicios"},
    "servicio": {"code": "65", "label": "Servicio", "filename": "Servicios"},
    "obras": {"code": "64", "label": "Obra", "filename": "Obras"},
    "obra": {"code": "64", "label": "Obra", "filename": "Obras"},
}
PREFERRED_COLUMNS = [
    "idProcedimiento", "detEntidad", "detTipoProceso", "nomenclatura",
    "fechaConvocatoria", "codObjeto", "detObjeto", "sintesisProceso",
    "descripcionItem", "item", "cubso", "moneda", "valorReferencial",
    "documentoBase", "ubigeo",
]
DOWNLOAD_TTL_SECONDS = int(os.getenv("DOWNLOAD_TTL_SECONDS", "1800"))
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "https://seace-export-api.onrender.com").rstrip("/")

# In-memory short-lived downloads. This avoids moving file bytes through ChatGPT/MCP.
_downloads: dict[str, dict] = {}
_downloads_lock = Lock()

mcp = FastMCP(
    "SEACE Export",
    instructions=(
        "Exporta oportunidades públicas del SEACE a Excel. "
        "Usa 'servicios' para código 65 y 'obras' para código 64. "
        "Devuelve un enlace HTTPS temporal al XLSX ya generado."
    ),
    stateless_http=True,
    json_response=True,
    streamable_http_path="/",
)


def _auth(x_api_key: str | None) -> None:
    expected = os.getenv("SEACE_EXPORT_API_KEY", "").strip()
    if expected and x_api_key != expected:
        raise HTTPException(status_code=401, detail="API key inválida")


def _normalize_object(value: str) -> dict:
    key = value.strip().lower()
    if key not in OBJECTS:
        raise HTTPException(status_code=400, detail="Objeto inválido. Use servicios u obras.")
    return OBJECTS[key]


def _parse_user_date(value: str) -> datetime:
    value = value.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    raise HTTPException(status_code=400, detail=f"Fecha inválida: {value}. Use DD/MM/YYYY o YYYY-MM-DD.")


def _parse_seace_date(value: object) -> datetime | None:
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})", str(value or ""))
    if not m:
        return None
    try:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def _fetch_json(code: str, retries: int = 3, timeout: int = 45) -> list[dict]:
    url = f"{SEACE_BASE}/{code}/0/0/0"
    last_error: Exception | None = None
    for attempt in range(retries):
        try:
            req = Request(url, headers={"User-Agent": "SEACE-Export-API/2.0", "Accept": "application/json"})
            with urlopen(req, timeout=timeout) as resp:
                raw = resp.read()
            import json
            payload = json.loads(raw.decode("utf-8"))
            if isinstance(payload, list):
                return [r for r in payload if isinstance(r, dict)]
            if isinstance(payload, dict) and isinstance(payload.get("data"), list):
                return [r for r in payload["data"] if isinstance(r, dict)]
            raise RuntimeError("La API SEACE respondió con un esquema inesperado")
        except (HTTPError, URLError, TimeoutError, RuntimeError, ValueError) as exc:
            last_error = exc
            if attempt < retries - 1:
                time.sleep(1.5 * (attempt + 1))
    raise HTTPException(status_code=502, detail=f"No se pudo consultar SEACE: {last_error}")


def _filter_rows(
    rows: list[dict],
    code: str,
    exact_dates: set[datetime],
    start: datetime | None,
    end: datetime | None,
) -> list[dict]:
    exact_days = {x.date() for x in exact_dates}
    out: list[dict] = []
    for row in rows:
        if str(row.get("codObjeto", "")) != code:
            continue
        d = _parse_seace_date(row.get("fechaConvocatoria"))
        if d is None:
            continue
        dn = d.date()
        if exact_days:
            if dn not in exact_days:
                continue
        elif start and end:
            if not (start.date() <= dn <= end.date()):
                continue
        out.append(row)
    return out


def _columns(rows: Iterable[dict]) -> list[str]:
    keys: list[str] = []
    seen = set()
    for row in rows:
        for key in row.keys():
            if key not in seen:
                seen.add(key)
                keys.append(key)
    return [c for c in PREFERRED_COLUMNS if c in seen] + [c for c in keys if c not in PREFERRED_COLUMNS]


def _autosize(ws) -> None:
    for idx, col in enumerate(ws.columns, 1):
        width = 10
        for cell in col[:150]:
            if cell.value is not None:
                width = max(width, min(len(str(cell.value)) + 2, 55))
        ws.column_dimensions[get_column_letter(idx)].width = width


def _build_workbook(rows: list[dict], columns: list[str], meta: list[tuple[str, object]]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Oportunidades"
    ws.sheet_view.showGridLines = False

    if columns:
        ws.append(columns)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(vertical="center")
        for row in rows:
            ws.append([row.get(c, "") for c in columns])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    else:
        ws.append(["Sin resultados"])

    _autosize(ws)

    ctrl = wb.create_sheet("Control")
    ctrl.sheet_view.showGridLines = False
    ctrl.append(["Campo", "Valor"])
    for cell in ctrl[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="7030A0")
    for key, value in meta:
        ctrl.append([key, value])
    ctrl.column_dimensions["A"].width = 28
    ctrl.column_dimensions["B"].width = 80

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _export(
    object_name: str,
    dates: list[str],
    start: str | None,
    end: str | None,
) -> tuple[bytes, str, int, int]:
    obj = _normalize_object(object_name)
    exact_dates = {_parse_user_date(d) for d in dates if d.strip()}
    start_dt = _parse_user_date(start) if start else None
    end_dt = _parse_user_date(end) if end else None

    if exact_dates and (start_dt or end_dt):
        raise HTTPException(status_code=400, detail="Use dates o start/end, no ambos.")
    if bool(start_dt) != bool(end_dt):
        raise HTTPException(status_code=400, detail="Debe indicar start y end juntos.")
    if start_dt and end_dt and start_dt > end_dt:
        raise HTTPException(status_code=400, detail="start no puede ser posterior a end.")
    if not exact_dates and not (start_dt and end_dt):
        raise HTTPException(status_code=400, detail="Indique dates o un rango start/end.")

    source = _fetch_json(obj["code"])
    filtered = _filter_rows(source, obj["code"], exact_dates, start_dt, end_dt)
    cols = _columns(filtered if filtered else source)

    if exact_dates:
        label = "_".join(sorted(d.strftime("%Y%m%d") for d in exact_dates))
        date_desc = ", ".join(sorted(d.strftime("%d/%m/%Y") for d in exact_dates))
    else:
        label = f"{start_dt:%Y%m%d}_{end_dt:%Y%m%d}"
        date_desc = f"{start_dt:%d/%m/%Y} - {end_dt:%d/%m/%Y}"

    filename = f"SEACE_{obj['filename']}_{label}.xlsx"
    meta = [
        ("Objeto", f"{obj['label']} ({obj['code']})"),
        ("Fecha(s)", date_desc),
        ("Registros descargados", len(source)),
        ("Registros filtrados", len(filtered)),
        ("Fuente", f"{SEACE_BASE}/{obj['code']}/0/0/0"),
        ("Versión servicio", APP_VERSION),
    ]
    return _build_workbook(filtered, cols, meta), filename, len(source), len(filtered)


def _cleanup_downloads() -> None:
    now = time.time()
    with _downloads_lock:
        expired = [token for token, item in _downloads.items() if item["expires_at"] <= now]
        for token in expired:
            _downloads.pop(token, None)


def _store_download(content: bytes, filename: str) -> str:
    _cleanup_downloads()
    token = secrets.token_urlsafe(24)
    with _downloads_lock:
        _downloads[token] = {
            "content": content,
            "filename": filename,
            "expires_at": time.time() + DOWNLOAD_TTL_SECONDS,
        }
    return f"{PUBLIC_BASE_URL}/download/{token}"


@mcp.tool()
def exportar_oportunidades(
    objeto: Literal["servicios", "obras"],
    fechas: list[str] | None = None,
    inicio: str | None = None,
    fin: str | None = None,
) -> dict:
    """Genera un XLSX de oportunidades SEACE y devuelve un enlace HTTPS temporal.

    Usa `fechas` para uno o varios días exactos (DD/MM/YYYY o YYYY-MM-DD),
    o usa `inicio` y `fin` juntos para un rango inclusivo. No mezcles ambos modos.
    """
    content, filename, source_count, filtered_count = _export(objeto, fechas or [], inicio, fin)
    url = _store_download(content, filename)
    return {
        "ok": True,
        "objeto": objeto,
        "archivo": filename,
        "registros_descargados": source_count,
        "registros_filtrados": filtered_count,
        "download_url": url,
        "expira_en_segundos": DOWNLOAD_TTL_SECONDS,
        "instruccion": "Entrega este enlace al usuario como descarga del archivo Excel.",
    }


@contextlib.asynccontextmanager
async def lifespan(app: FastAPI):
    async with mcp.session_manager.run():
        yield


app = FastAPI(
    title="SEACE Export API",
    version=APP_VERSION,
    description="Consulta oportunidades SEACE y devuelve Excel filtrado por objeto y fecha.",
    lifespan=lifespan,
)


@app.get("/")
def root():
    return {
        "ok": True,
        "service": "SEACE Export API",
        "version": APP_VERSION,
        "health": "/health",
        "mcp": "/mcp",
    }


@app.get("/health")
def health():
    return {"ok": True, "version": APP_VERSION, "mcp": "/mcp"}


@app.get("/v1/export")
def export_get(
    object: str = Query(..., description="servicios u obras"),
    dates: list[str] = Query(default=[]),
    start: str | None = Query(default=None),
    end: str | None = Query(default=None),
    x_api_key: str | None = Header(default=None),
):
    _auth(x_api_key)
    content, filename, source_count, filtered_count = _export(object, dates, start, end)
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "X-SEACE-Source-Count": str(source_count),
        "X-SEACE-Filtered-Count": str(filtered_count),
    }
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/download/{token}")
def download_generated(token: str):
    _cleanup_downloads()
    with _downloads_lock:
        item = _downloads.get(token)
    if not item:
        raise HTTPException(status_code=404, detail="Archivo no encontrado o enlace expirado")
    headers = {"Content-Disposition": f'attachment; filename="{item["filename"]}"'}
    return Response(
        content=item["content"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


# Streamable HTTP MCP endpoint at exactly /mcp
app.mount("/mcp", mcp.streamable_http_app())
