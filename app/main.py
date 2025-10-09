"""DIAL application entry point for the JSON to PowerPoint service."""

from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path
from typing import Any, Dict

from aidial_sdk import DIALApp, HTTPException
from aidial_sdk.chat_completion import ChatCompletion, Request, Response
from fastapi import Request as FastAPIRequest
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import Response as StarletteResponse, StreamingResponse
import asyncio
from typing import AsyncGenerator

from .template_engine import TemplateEngine, encode_pptx

# Configure logging based on environment variable
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s [%(levelname)s] [%(name)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

LOGGER = logging.getLogger(__name__)

TEMPLATES_DIR = Path(__file__).resolve().parent / "templates"
ENGINE = TemplateEngine(str(TEMPLATES_DIR))
DEFAULT_OUTPUT_NAME = "presentation.pptx"
# Use the PowerPoint MIME type that works best with DIAL and OneDrive
MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


class LoggingMiddleware(BaseHTTPMiddleware):
    """Middleware to log all incoming requests and responses with full details."""
    
    async def _log_response_body(self, response_body: bytes, is_chat_completion: bool):
        """Helper to log response body content."""
        if not is_chat_completion:
            return
            
        body_str = response_body.decode('utf-8', errors='ignore')
        LOGGER.debug(f"Response body length: {len(body_str)} characters")
        
        if len(body_str) < 2000:  # Log smaller responses in full
            LOGGER.debug(f"Full response body: {body_str}")
        else:
            # For larger responses, log the beginning and end
            LOGGER.debug(f"Response body (first 1000 chars): {body_str[:1000]}")
            LOGGER.debug(f"Response body (last 500 chars): {body_str[-500:]}")
            
        # Try to parse as JSON and log structure
        try:
            if body_str.strip():
                import json
                parsed = json.loads(body_str)
                LOGGER.debug(f"Response JSON structure: {self._get_json_structure(parsed)}")
        except:
            LOGGER.debug("Response body is not valid JSON")
    
    def _get_json_structure(self, obj, max_depth=3, current_depth=0):
        """Get a summary of JSON structure without logging sensitive data."""
        if current_depth > max_depth:
            return "..."
            
        if isinstance(obj, dict):
            return {key: self._get_json_structure(value, max_depth, current_depth + 1) 
                   for key, value in obj.items()}
        elif isinstance(obj, list):
            if len(obj) > 0:
                return [self._get_json_structure(obj[0], max_depth, current_depth + 1), f"...({len(obj)} items)"]
            return []
        elif isinstance(obj, str):
            return f"string(len={len(obj)})"
        else:
            return type(obj).__name__
    
    async def dispatch(self, request: FastAPIRequest, call_next):
        # Log incoming request details
        LOGGER.info(f"=== INCOMING REQUEST ===")
        LOGGER.info(f"Method: {request.method}")
        LOGGER.info(f"URL: {request.url}")
        LOGGER.info(f"Path: {request.url.path}")
        LOGGER.info(f"Query params: {dict(request.query_params)}")
        LOGGER.info(f"Headers: {dict(request.headers)}")
        LOGGER.info(f"Client: {request.client}")
        
        # Try to log body for POST/PUT requests
        if request.method in ["POST", "PUT", "PATCH"]:
            try:
                body = await request.body()
                LOGGER.info(f"Body length: {len(body)} bytes")
                if len(body) < 1000:  # Only log small bodies in full
                    LOGGER.info(f"Body content: {body.decode('utf-8', errors='ignore')}")
                else:
                    LOGGER.info(f"Body content (first 500 chars): {body[:500].decode('utf-8', errors='ignore')}...")
            except Exception as e:
                LOGGER.error(f"Failed to read request body: {e}")
        
        # Process the request
        response = None
        try:
            response = await call_next(request)
            
            # Log response details
            LOGGER.info(f"=== OUTGOING RESPONSE ===")
            LOGGER.info(f"Response status: {response.status_code}")
            LOGGER.info(f"Response headers: {dict(response.headers)}")
            
            # Enhanced response logging for chat completions
            is_chat_completion = request.url.path.endswith('/chat/completions')
            if is_chat_completion:
                try:
                    LOGGER.info("Chat completion response generated")
                    LOGGER.debug(f"Response media type: {response.media_type}")
                    LOGGER.debug(f"Response charset: {getattr(response, 'charset', 'N/A')}")
                    
                    # Log response size if available
                    content_length = response.headers.get('content-length')
                    if content_length:
                        LOGGER.info(f"Response content length: {content_length} bytes")
                    
                    # For StreamingResponse, we need special handling
                    if isinstance(response, StreamingResponse):
                        LOGGER.debug("Response is a StreamingResponse - cannot easily log body")
                    elif hasattr(response, 'body') and response.body:
                        # Try to read and log the response body
                        try:
                            await self._log_response_body(response.body, True)
                        except Exception as e:
                            LOGGER.debug(f"Could not log response body: {e}")
                    else:
                        LOGGER.debug("Response has no accessible body attribute")
                        
                except Exception as e:
                    LOGGER.debug(f"Could not log detailed response info: {e}")
                    
            LOGGER.info(f"=== RESPONSE SENT ===")
            return response
        except Exception as e:
            LOGGER.error(f"Request failed with exception: {e}")
            LOGGER.info(f"=== REQUEST FAILED ===")
            raise


class PresentationApplication(ChatCompletion):
    """Render a PowerPoint file from a JSON description."""

    async def chat_completion(self, request: Request, response: Response) -> None:
        LOGGER.info("=== CHAT COMPLETION REQUEST ===")
        LOGGER.info(f"Request messages count: {len(request.messages) if request.messages else 0}")
        LOGGER.debug(f"Full request object: {request}")
        
        if not request.messages:
            LOGGER.error("Request is empty - no messages provided")
            raise HTTPException(status_code=422, message="The request is empty")

        message = request.messages[-1]
        content = message.text()
        LOGGER.info(f"Processing message content length: {len(content)} characters")
        LOGGER.debug(f"Message content: {content}")

        try:
            payload = json.loads(content)
            if not isinstance(payload, dict):
                raise ValueError("Payload must be a JSON object")
            LOGGER.info(f"Successfully parsed JSON payload with keys: {list(payload.keys())}")
            LOGGER.debug(f"Full payload: {payload}")
        except (json.JSONDecodeError, ValueError) as exc:
            LOGGER.exception("Failed to parse request payload")
            raise HTTPException(
                status_code=422, message=f"Invalid JSON payload: {exc}"
            ) from exc

        output_name = _resolve_output_name(payload)
        LOGGER.info(f"Output filename: {output_name}")

        try:
            LOGGER.info("Starting presentation rendering...")
            pptx_bytes = ENGINE.render(payload)
            LOGGER.info(f"Successfully rendered presentation, size: {len(pptx_bytes)} bytes")
        except Exception as exc:  # pragma: no cover - converted to HTTP error
            LOGGER.exception("Failed to render presentation")
            raise HTTPException(status_code=500, message=str(exc)) from exc

        LOGGER.info("Encoding presentation to base64...")
        pptx_base64 = encode_pptx(pptx_bytes)
        LOGGER.info(f"Base64 encoding complete, length: {len(pptx_base64)} characters")

        LOGGER.info("Creating response with attachment...")
        LOGGER.info(f"Attachment details - Title: {output_name}, Data length: {len(pptx_base64)} characters")
        LOGGER.debug(f"Base64 data sample (first 100 chars): {pptx_base64[:100]}")
        
        # Create response following the exact pattern from DIAL SDK examples
        with response.create_single_choice() as choice:
            LOGGER.debug("Response choice created successfully")
            
            choice.append_content(
                f"Generated presentation '{output_name}' using template instructions."
            )
            LOGGER.debug("Content appended to choice")
            
            LOGGER.info("Adding PowerPoint attachment to response...")
            LOGGER.debug(f"Attachment parameters: type='{MIME_TYPE}', title='{output_name}', data_length={len(pptx_base64)}")
            
            # Add the PowerPoint file as an attachment - following render_text example pattern
            choice.add_attachment(
                type=MIME_TYPE,
                title=output_name,
                data=pptx_base64
            )
            
            LOGGER.info("Attachment added successfully")
            LOGGER.debug("Choice with attachment completed")
            
        LOGGER.info("=== CHAT COMPLETION SUCCESSFUL ===")
        LOGGER.info("Response created successfully with presentation attachment")
        LOGGER.debug("Response object fully constructed and ready to send")


def _resolve_output_name(payload: Dict[str, Any]) -> str:
    output = payload.get("output", {})
    if isinstance(output, dict):
        name = output.get("file_name")
        if isinstance(name, str) and name.strip():
            # Ensure the filename has the .pptx extension
            if not name.lower().endswith('.pptx'):
                name = name + '.pptx'
            return name
    return DEFAULT_OUTPUT_NAME


# Configuration for DIAL Core integration
DIAL_URL = os.getenv("DIAL_URL")

LOGGER.info("=== APPLICATION STARTUP ===")
LOGGER.info(f"Log level set to: {LOG_LEVEL}")
LOGGER.info(f"DIAL_URL configured: {'Yes' if DIAL_URL else 'No'}")
LOGGER.info(f"Templates directory: {TEMPLATES_DIR}")

# Create DIAL app with optional DIAL Core integration
app = DIALApp(
    dial_url=DIAL_URL,
    propagate_auth_headers=bool(DIAL_URL),  # Only enable if DIAL_URL is set
    add_healthcheck=True
)

# Add logging middleware to capture all HTTP requests
app.add_middleware(LoggingMiddleware)

# Add health check logging
@app.get("/health")
async def health_check():
    LOGGER.info("Health check endpoint accessed")
    return {"status": "healthy", "service": "json-to-pptx"}

app.add_chat_completion("json-to-pptx", PresentationApplication())

LOGGER.info("Application initialized successfully")
LOGGER.info("=== READY TO RECEIVE REQUESTS ===")
LOGGER.info(f"Server will start on port 5000")
