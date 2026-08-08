# API de exportación SEACE + MCP v4 endurecida

Servicio de propósito único para exportar oportunidades públicas del SEACE a Excel y exponer una sola herramienta MCP a ChatGPT.

## Superficie permitida

- `GET /health` — estado del servicio.
- `GET /security` — resumen público de controles de seguridad.
- `/mcp` — MCP Streamable HTTP.
- `GET /download/{token}` — descarga temporal, aleatoria y de un solo uso.
- `GET /v1/export` — deshabilitado por defecto. Solo se habilita con `ENABLE_DIRECT_EXPORT=1` y una `SEACE_EXPORT_API_KEY` configurada.

## Única herramienta MCP

`exportar_oportunidades`

Parámetros:

- `objeto`: solo `servicios` o `obras`.
- `fechas`: uno o varios días exactos (`DD/MM/YYYY` o `YYYY-MM-DD`).
- `inicio` + `fin`: rango inclusivo alternativo a `fechas`.

No acepta URLs, código, comandos, SQL, prompts libres, credenciales ni otros orígenes de datos.

## Controles v4

- Endpoint SEACE hardcodeado y allowlist de códigos 64/65.
- Máximo 31 fechas exactas o rango de 31 días.
- Límites de tamaño de respuesta, filas de origen y filas exportadas.
- Límite de exportaciones por minuto.
- Descargas de un solo uso y TTL de 10 minutos.
- Fórmulas de Excel provenientes de datos públicos se neutralizan para evitar formula injection.
- No se devuelven filas SEACE al modelo; MCP devuelve solo conteos, nombre de archivo y URL temporal.
- No se exponen docs/OpenAPI públicos en producción.
- Exportación REST directa deshabilitada por defecto.
- No hay escritura en SEACE ni en sistemas internos.

## Importante sobre el aviso de ChatGPT

ChatGPT seguirá mostrando advertencias para apps MCP personalizadas no verificadas. Este endurecimiento reduce la superficie real de riesgo, pero no elimina la etiqueta genérica de riesgo porque OpenAI no revisa automáticamente cada app personalizada del workspace.

## Despliegue en Render

Reemplazar en el repositorio:

- `app/main.py`
- `requirements.txt`
- `render.yaml`
- `README.md`

Render debe desplegar automáticamente el commit. Verificar después:

```text
https://seace-export-api.onrender.com/health
https://seace-export-api.onrender.com/security
https://seace-export-api.onrender.com/mcp
```

La URL MCP para ChatGPT sigue siendo:

```text
https://seace-export-api.onrender.com/mcp
```


## Nota v4.1
La herramienta MCP devuelve metadatos como JSON textual (`structured_output=False`) para mantener compatibilidad con `mcp[cli]==2.0.0b2` y evitar `InvalidSignature` al registrar un retorno `dict`.
