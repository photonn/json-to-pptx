"""DIAL application entry point for the JSON to PowerPoint service."""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any, Dict

from aidial_sdk import DIALApp, HTTPException
from aidial_sdk.chat_completion import ChatCompletion, Request, Response

from .template_engine import TemplateEngine, encode_pptx

LOGGER = logging.getLogger(__name__)

TEMPLATES_DIR = Path(__file__).resolve().parent / "templates"
ENGINE = TemplateEngine(str(TEMPLATES_DIR))
DEFAULT_OUTPUT_NAME = "presentation.pptx"
MIME_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


class PresentationApplication(ChatCompletion):
    """Render a PowerPoint file from a JSON description."""

    async def chat_completion(self, request: Request, response: Response) -> None:
        if not request.messages:
            raise HTTPException(status_code=422, message="The request is empty")

        message = request.messages[-1]
        content = message.text()

        try:
            payload = json.loads(content)
            if not isinstance(payload, dict):
                raise ValueError("Payload must be a JSON object")
        except (json.JSONDecodeError, ValueError) as exc:
            LOGGER.exception("Failed to parse request payload")
            raise HTTPException(
                status_code=422, message=f"Invalid JSON payload: {exc}"
            ) from exc

        output_name = _resolve_output_name(payload)

        try:
            pptx_bytes = ENGINE.render(payload)
        except Exception as exc:  # pragma: no cover - converted to HTTP error
            LOGGER.exception("Failed to render presentation")
            raise HTTPException(status_code=500, message=str(exc)) from exc

        pptx_base64 = encode_pptx(pptx_bytes)

        with response.create_single_choice() as choice:
            choice.append_content(
                f"Generated presentation '{output_name}' using template instructions."
            )
            choice.add_attachment(
                type=MIME_TYPE,
                title=output_name,
                data=pptx_base64,
            )


def _resolve_output_name(payload: Dict[str, Any]) -> str:
    output = payload.get("output", {})
    if isinstance(output, dict):
        name = output.get("file_name")
        if isinstance(name, str) and name.strip():
            return name
    return DEFAULT_OUTPUT_NAME


# Configuration for DIAL Core integration
DIAL_URL = os.getenv("DIAL_URL")

# Create DIAL app with optional DIAL Core integration
app = DIALApp(
    dial_url=DIAL_URL,
    propagate_auth_headers=bool(DIAL_URL),  # Only enable if DIAL_URL is set
    add_healthcheck=True
)
app.add_chat_completion("json-to-pptx", PresentationApplication())
