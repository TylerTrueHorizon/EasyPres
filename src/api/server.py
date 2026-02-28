import sys
import os
import tempfile
import io

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "agents"))

from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Optional

from presentation_agent import build_presentation

app = FastAPI(title="EasyPres API")


class GenerateRequest(BaseModel):
    content: str
    num_slides: Optional[int] = None
    webhook_url: Optional[str] = None
    webhook_headers: Optional[dict[str, str]] = None


PPTX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


@app.post("/generate")
async def generate(req: GenerateRequest):
    """Generate a PowerPoint presentation and return the .pptx binary."""
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        await build_presentation(
            content=req.content,
            output_path=tmp_path,
            num_slides=req.num_slides,
            webhook_url=req.webhook_url,
            webhook_headers=req.webhook_headers,
        )

        with open(tmp_path, "rb") as f:
            pptx_bytes = f.read()
    finally:
        os.unlink(tmp_path)

    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type=PPTX_MEDIA_TYPE,
        headers={
            "Content-Disposition": "attachment; filename=presentation.pptx",
        },
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
