"""Render .pptx slides to PNG via PowerPoint COM automation.

WithWindow=True is required: PowerPoint uses cached thumbnails when
opening WithWindow=False, so python-pptx text changes won't appear
in the exported PNG. We open with a visible window (minimized) to
force full re-rendering.
"""

from __future__ import annotations

import time
from pathlib import Path

import pythoncom
import win32com.client


def render_slides_to_png(
    items: list[tuple[str, str]],
    slide_index: int = 0,
    width: int = 1920,
) -> list[str]:
    results = []
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        app.WindowState = 2  # ppWindowMinimized
        try:
            for pptx_path, output_png in items:
                pptx_abs = str(Path(pptx_path).resolve())
                out_abs = str(Path(output_png).resolve())
                Path(out_abs).parent.mkdir(parents=True, exist_ok=True)

                pres = app.Presentations.Open(pptx_abs)
                time.sleep(0.5)
                try:
                    slide = pres.Slides[slide_index + 1]
                    slide.Export(out_abs, "PNG", width)
                finally:
                    pres.Close()

                if not Path(out_abs).exists():
                    raise RuntimeError(f"PowerPoint failed to export: {out_abs}")
                results.append(out_abs)
        finally:
            app.Quit()
    finally:
        pythoncom.CoUninitialize()
    return results


def render_slide_to_png(
    pptx_path: str,
    output_png: str,
    slide_index: int = 0,
    width: int = 1920,
) -> str:
    return render_slides_to_png([(pptx_path, output_png)], slide_index, width)[0]
