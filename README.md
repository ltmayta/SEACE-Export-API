# SEACE Export API

Servicio API-first para generar Excel de oportunidades SEACE sin Browserless en operación normal.

## Rutas

- `GET /health`
- `GET /v1/export`

### Ejemplos

Una fecha exacta:

```text
/v1/export?object=servicios&dates=05/08/2026
```

Dos fechas exactas:

```text
/v1/export?object=servicios&dates=03/08/2026&dates=04/08/2026
```

Rango inclusivo:

```text
/v1/export?object=obras&start=03/08/2026&end=05/08/2026
```

Si `SEACE_EXPORT_API_KEY` está configurada, enviar cabecera `X-API-Key`.

## Desarrollo local

```bash
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8080
```

## Docker

```bash
docker build -t seace-export-api .
docker run --rm -p 8080:8080 -e SEACE_EXPORT_API_KEY=CAMBIAR seace-export-api
```

## Despliegue sugerido

El archivo `render.yaml` permite desplegar en Render. También puede desplegarse el mismo `Dockerfile` en Railway, Cloud Run, Azure Container Apps u otro servicio compatible.

## Integración con el Agente

El objetivo final es que el Agente haga una sola llamada a `/v1/export`, reciba el archivo y lo entregue. Browserless queda fuera del camino normal y solo se conserva para diagnóstico si SEACE cambia el backend.
