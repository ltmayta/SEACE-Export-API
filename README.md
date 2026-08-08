# API de exportación SEACE + MCP v3

Servicio para exportar oportunidades públicas del SEACE a Excel y exponer una herramienta MCP remota para ChatGPT.

## Endpoints

- `GET /health` — estado del servicio.
- `GET /v1/export` — descarga directa de Excel por HTTP.
- `/mcp` — endpoint MCP Streamable HTTP.
- `GET /download/{token}` — descarga temporal del Excel generado por la herramienta MCP.

## Herramienta MCP

`exportar_oportunidades` acepta:

- `objeto`: `servicios` o `obras`.
- `fechas`: lista de fechas exactas `DD/MM/YYYY` o `YYYY-MM-DD`.
- `inicio` y `fin`: rango inclusivo alternativo a `fechas`.

La herramienta consulta SEACE, genera el XLSX en el servidor y devuelve un enlace HTTPS temporal.

## Códigos SEACE verificados

- Servicios: `65`
- Obras: `64`

## MCP SDK

Esta versión fija `mcp[cli]==2.0.0b2` y usa la API v2 `MCPServer`. El host público de Render está incluido explícitamente en `TransportSecuritySettings`, requisito del transporte Streamable HTTP para despliegues remotos.

## Render

El `Dockerfile` inicia `uvicorn app.main:app`. `render.yaml` configura `/health` como health check.

Después del despliegue, verifica primero:

```text
https://seace-export-api.onrender.com/health
```

Luego registra como URL MCP:

```text
https://seace-export-api.onrender.com/mcp
```
