# SEACE Export API + MCP

Servicio para exportar oportunidades públicas del SEACE a Excel y exponer una herramienta MCP para ChatGPT.

## Endpoints

- `GET /health` — estado del servicio.
- `GET /v1/export` — descarga directa por HTTP.
- `POST /mcp` — endpoint MCP Streamable HTTP para ChatGPT.
- `GET /download/{token}` — descarga temporal de un XLSX ya generado por la herramienta MCP.

## Herramienta MCP

`exportar_oportunidades`

Parámetros:

- `objeto`: `servicios` u `obras`.
- `fechas`: lista de fechas exactas (`DD/MM/YYYY` o `YYYY-MM-DD`).
- `inicio`, `fin`: rango inclusivo; usar ambos y no combinarlos con `fechas`.

La herramienta consulta SEACE una sola vez, genera el XLSX en el servidor y devuelve un enlace HTTPS temporal. El archivo no se transfiere a través del contenido MCP ni del chat.

## URL en Render

MCP: `https://seace-export-api.onrender.com/mcp`

## Prueba HTTP

`/v1/export?object=servicios&dates=05/08/2026`

## Seguridad

Durante la validación inicial el MCP está sin autenticación. Después de confirmar la integración con ChatGPT, añadir autenticación del conector antes de usarlo como servicio de producción.
