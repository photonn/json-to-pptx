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
- Returns the rendered PowerPoint file as a base64 attachment in the DIAL chat
  completion response.

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

The response contains a single choice with the generated PowerPoint attached as
base64 data under `choices[0].attachments[0].data`.

### Using Docker

```bash
docker build -t json-to-pptx .
docker run -p 8000:8000 json-to-pptx
```

The container exposes the same API on port `8000`.

## Extending the Template

Add new placeholders to your template files or ship additional templates by
placing more `.pptx` files in `app/templates/`.  Make sure to reference new
tables by their shape names in the JSON payload.  The templating logic is a
modernised subset of `pptx-template`, so only text replacement and basic table
population are currently supported.
