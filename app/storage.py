import base64
import os
from io import BytesIO

import aiohttp


async def upload_pptx_file(dial_url: str, filepath: str, pptx_base64: str, content_type: str = "application/vnd.openxmlformats-officedocument.presentationml.presentation") -> str:
    """Upload a PPTX file (base64) to DIAL storage and return a URL.

    Mirrors the pattern from ai-dial-sdk examples/render_text/app/image.py
    but adapted for PPTX content type and path.
    """
    async with aiohttp.ClientSession() as session:
        async with session.get(f"{dial_url}/v1/bucket") as response:
            response.raise_for_status()
            appdata = (await response.json())["appdata"]

        pptx_bytes = base64.b64decode(pptx_base64)

        data = aiohttp.FormData()
        data.add_field(
            name="file",
            content_type=content_type,
            value=BytesIO(pptx_bytes),
            filename=os.path.basename(filepath),
        )

        async with session.put(
            f"{dial_url}/v1/files/{appdata}/{filepath}", data=data
        ) as response:
            response.raise_for_status()
            metadata = await response.json()

    return metadata["url"]
