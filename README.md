# JSON to PPTX DIAL Application

This project packages a [DIAL](https://dialx.ai) application that converts a JSON
instruction payload into a PowerPoint presentation using a template-driven
approach.  The service is implemented with the [ai-dial-sdk](https://github.com/epam/ai-dial-sdk)
and reuses the templating conventions from the
[pptx-template](https://github.com/m3dev/pptx-template) project to replace
placeholders inside `.pptx` files.

## Features

- Runs as a FastAPI application compatible with the DIAL runtime.
- Accepts structured JSON describing slide updates and table data.
- Replaces placeholder text in user-provided templates and fills named tables with
  dynamic content.
- Returns the rendered PowerPoint file as a downloadable attachment in the DIAL chat
  completion response.
- **Option B Implementation**: Uploads files to DIAL storage and provides clickable download links.

## Project Layout

```
.
├── aidial_sdk/                # Local copy of the DIAL SDK runtime
├── app/
│   ├── __init__.py            # Exposes the FastAPI/DIAL application instance
│   ├── main.py                # Chat completion implementation
│   ├── template_engine.py     # Minimal templating logic derived from pptx-template
│   └── templates/             # Directory containing user-provided templates
├── Dockerfile
├── examples/
│   └── sample_payload.json    # Example request payload
├── requirements.txt
└── README.md
```

## JSON Payload Structure

The request **must** be a JSON object.  The top-level fields are:

| Field | Type | Description |
| ----- | ---- | ----------- |
| `template` | string | Optional template name. Defaults to `"default"` and resolves to `app/templates/<name>.pptx`. |
| `output.file_name` | string | Optional filename used for the generated attachment. |
| `context` | object | Global values that can be referenced by placeholders such as `{context.title}`. |
| `slides` | array | Per-slide instructions. Each entry must provide an `index` or an `id` that matches a `{id:...}` marker inside the template. |

Each slide entry supports:

- `index` *(integer)* – zero-based slide index.
- `id` *(string)* – template slide identifier inside `{id:...}` markers.
- `replacements` *(object)* – values merged into the global context before
  resolving placeholders on the slide.
- `tables` *(array)* – table population instructions.  Each item must specify:
  - `shape` *(string)* – the PowerPoint shape name of the table to fill.
  - `header` *(array, optional)* – values inserted as the first row.
  - `data` *(array of arrays)* – body rows written after the optional header.
  - `clear_extra_rows` *(boolean, optional, defaults to `true`)* – whether to
    blank out any unused template rows.

Placeholders inside the template follow the `{path.to.value}` convention.  Dotted
paths are resolved against the merged slide context, supporting nested objects
and list indices.

## Providing Templates

This repository does not include any PowerPoint templates.  Before running the
service, place a `.pptx` file in `app/templates/` named `default.pptx` (or
adjust the JSON payload to reference your chosen filename).  The application
looks up templates by `<name>.pptx` in that directory when processing requests.

## Running Locally

### Using Python

```bash
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Send a request to the DIAL-compatible endpoint:

```bash
curl -X POST http://localhost:8000/openai/deployments/json-to-pptx/chat/completions \
  -H "Content-Type: application/json" \
  -d @examples/sample_payload.json
```

The response contains a single choice with the generated PowerPoint as a downloadable
attachment. When deployed with DIAL_URL configured, the attachment will have a `url` 
field pointing to DIAL file storage, making it clickable in the UI.

```bash
curl -X POST "http://localhost:8000/openai/deployments/json-to-pptx/chat/completions" \
  -H "Content-Type: application/json" \
  -d '{
    "messages": [
      {
        "role": "user",
        "content": "{\"template\": \"default\", \"output\": {\"file_name\": \"test-presentation.pptx\"}, \"slides\": [{\"index\": 0, \"replacements\": {\"text1\": \"This is a sample text for replacement.\"}}]}"
      }
    ]
  }'
```

### Using Docker

```bash
docker build -t json-to-pptx .
docker run -p 8000:8000 json-to-pptx
```

The container exposes the same API on port `8000`.

### Deployment with DIAL (Option B - Recommended)

For downloadable attachments, deploy with DIAL_URL configured:

```bash
docker run -p 8000:8000 -e DIAL_URL=https://your-dial-instance.com json-to-pptx
```

This enables:
- Automatic file upload to DIAL storage
- Clickable download links in the UI  
- Integration with OneDrive/file systems

### Debugging and Logging

For troubleshooting deployment issues or to see detailed request logs, you can enable debug logging by setting the `LOG_LEVEL` environment variable:

```bash
# Enable DEBUG logging to see all incoming requests, headers, and processing steps
docker run -p 8000:8000 -e LOG_LEVEL=DEBUG json-to-pptx

# Or when deploying to DIAL
# Set LOG_LEVEL=DEBUG in your deployment configuration
```

Available log levels:
- `ERROR`: Only errors and exceptions
- `WARN`: Warnings and above  
- `INFO`: General information (default)
- `DEBUG`: Detailed debugging information including:
  - All incoming HTTP requests with headers and body
  - Chat completion processing steps
  - Template rendering progress
  - Response creation details

When `LOG_LEVEL=DEBUG` is set, the application will log:
- Complete request details (method, URL, headers, body)
- Client connection information
- JSON payload parsing and validation
- Template rendering progress
- Response generation steps
- Health check endpoint access

This is especially useful when deployed to DIAL to diagnose connection issues or request processing problems.

## Extending the Template

Add new placeholders to your template files or ship additional templates by
placing more `.pptx` files in `app/templates/`.  Make sure to reference new
tables by their shape names in the JSON payload.  The templating logic is a
modernised subset of `pptx-template`, so only text replacement and basic table
population are currently supported.
