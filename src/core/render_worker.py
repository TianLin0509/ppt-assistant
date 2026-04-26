"""Standalone COM render worker — runs as subprocess with desktop access.

Reads JSON from stdin: {"items": [[pptx, png], ...], "slide_index": 0, "width": 1920}
Opens PowerPoint visibly, renders each slide, exports PNG, quits.
"""

import json
import sys
import time

import pythoncom
import win32com.client


def main():
    payload = json.loads(sys.stdin.read())
    items = payload["items"]
    slide_index = payload.get("slide_index", 0)
    width = payload.get("width", 1920)

    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        try:
            for pptx_path, output_png in items:
                pres = app.Presentations.Open(pptx_path)
                time.sleep(1.5)
                try:
                    slide = pres.Slides(slide_index + 1)
                    slide.Export(output_png, "PNG", width)
                finally:
                    pres.Close()
        finally:
            app.Quit()
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
