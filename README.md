# Client Reports Automation

Microservicio Python (Flask) para la generación automática de informes mensuales
en formato .pptx a partir de plantillas + datos.

Usado por la agencia **The Hook / IKI Group** para los clientes:
- Malne (implementado)
- Yamaha (pendiente)
- Eskariam (pendiente)
- Ávalos Villas (pendiente)
- Interwetten (pendiente)
- MyWin360 (pendiente)

## Cómo funciona

1. n8n Cloud recopila los datos del mes (GA4, GSC, Ahrefs) y genera los comentarios AI con OpenAI
2. n8n llama a este microservicio vía HTTP POST con todos los datos + URLs de gráficos
3. Este microservicio abre la plantilla `.pptx`, sustituye los placeholders y las imágenes, y devuelve el .pptx generado
4. n8n sube el .pptx generado a Google Drive

## Endpoints

### `GET /`
Health check. Devuelve el estado del servicio y lista las plantillas disponibles.

### `POST /generate`
Genera un informe a partir de una plantilla y datos.

**Body (JSON):**
```json
{
  "template": "Malne-Template.pptx",
  "placeholders": {
    "tra_web_curr": "27.729",
    "comentario_trafico": "El tráfico total ha aumentado...",
    "ev_full_1_name": "page_view"
  },
  "images": {
    "5": "https://quickchart.io/chart?...",
    "6": "https://quickchart.io/chart?..."
  }
}
```

**Response:** archivo `.pptx` binario.

## Deploy en Render.com

- Runtime: Python 3
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app`
- Instance Type: Free

## Añadir un nuevo cliente

1. Preparar la plantilla `.pptx` con los `{{placeholder}}` en las posiciones deseadas
2. Añadir el archivo al repositorio
3. Actualizar el workflow n8n del cliente para llamar al endpoint con `"template": "NombreCliente-Template.pptx"`

## Coste

- GitHub: gratis (plan Free)
- Render.com: gratis (plan Free hasta 750h/mes)
- Total: 0 €/mes
