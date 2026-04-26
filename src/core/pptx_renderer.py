"""Render .pptx slides to PNG via PowerPoint COM automation.

python-pptx modifies XML but doesn't update PowerPoint's cached
rendering data. COM Export() with WithWindow=False or from background
threads without desktop access produces stale PNGs.

Fix: spawn a subprocess with desktop access to do the COM rendering.
The subprocess runs render_worker.py which opens PowerPoint visibly,
waits for rendering, exports, then quits.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


_WORKER_SCRIPT = Path(__file__).parent / "render_worker.py"


def render_slides_to_png(
    items: list[tuple[str, str]],
    slide_index: int = 0,
    width: int = 1920,
) -> list[str]:
    items_abs = []
    for pptx_path, output_png in items:
        pptx_abs = str(Path(pptx_path).resolve())
        out_abs = str(Path(output_png).resolve())
        Path(out_abs).parent.mkdir(parents=True, exist_ok=True)
        items_abs.append((pptx_abs, out_abs))

    payload = json.dumps({
        "items": items_abs,
        "slide_index": slide_index,
        "width": width,
    })

    result = subprocess.run(
        [sys.executable, str(_WORKER_SCRIPT)],
        input=payload,
        capture_output=True,
        text=True,
        timeout=120,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"Render worker failed (exit {result.returncode}):\n{result.stderr}"
        )

    results = []
    for _, out_abs in items_abs:
        if not Path(out_abs).exists():
            raise RuntimeError(f"PowerPoint failed to export: {out_abs}")
        results.append(out_abs)
    return results


def render_slide_to_png(
    pptx_path: str,
    output_png: str,
    slide_index: int = 0,
    width: int = 1920,
) -> str:
    return render_slides_to_png([(pptx_path, output_png)], slide_index, width)[0]
