"""Task history viewer (Phase 1 minimal stub)."""

import streamlit as st
from pathlib import Path

from src.utils.config import load_config, resolve_path

st.set_page_config(page_title="历史任务", layout="wide")
st.header("历史任务")

config = load_config()
runs_dir = resolve_path(config.runs_dir)

if not runs_dir.exists():
    st.info("暂无历史任务。")
    st.stop()

task_dirs = sorted([d for d in runs_dir.iterdir() if d.is_dir()], reverse=True)
if not task_dirs:
    st.info("暂无历史任务。")
    st.stop()

for d in task_dirs:
    task_json = d / "task.json"
    preview = d / "preview.png"
    output = d / "output.pptx"

    with st.expander(d.name):
        if preview.exists():
            st.image(str(preview), use_container_width=True)
        if output.exists():
            with open(str(output), "rb") as f:
                st.download_button("下载 .pptx", data=f, file_name=f"{d.name}.pptx", key=f"dl_{d.name}")
