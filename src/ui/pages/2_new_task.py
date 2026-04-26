"""New task workflow: select template -> generate prompt -> paste JSON -> pick candidates -> export."""

from datetime import datetime
from pathlib import Path

import streamlit as st

from src.utils.config import load_config, resolve_path
from src.utils.state_manager import (
    list_template_metas, load_template_meta, save_task, save_raw_response,
)
from src.prompt.text_prompt_builder import build_text_prompt
from src.prompt.json_parser import parse_ai_json
from src.core.pptx_filler import fill_template
from src.core.pptx_renderer import render_slide_to_png
from src.schema import TaskRun, TextCandidates
from src.ui.components.candidate_picker import render_candidate_picker

try:
    from st_copy_to_clipboard import st_copy_to_clipboard
except ImportError:
    st_copy_to_clipboard = None

st.set_page_config(page_title="新建任务", layout="wide")
st.header("新建任务")

config = load_config()
all_metas = list_template_metas()

if not all_metas:
    st.warning("暂无已预处理的模板。请先在「模板管理」页面预处理模板。")
    st.stop()

# --- Step 1: Select template ---
st.subheader("Step 1: 选择模板")
template_ids = [m.template_id for m in all_metas]
selected_id = st.selectbox("模板", template_ids)
meta = load_template_meta(selected_id)

if meta and meta.annotated_image and Path(meta.annotated_image).exists():
    st.image(meta.annotated_image, use_container_width=True)

# --- Step 2: Task description ---
st.subheader("Step 2: 任务描述")
task_desc = st.text_area("描述你的 PPT 主题", placeholder="例如：介绍最新的 L2O 优化算法")

if not task_desc:
    st.stop()

# --- Initialize task ---
if "current_task" not in st.session_state:
    now = datetime.now()
    task_id = f"{now.strftime('%Y-%m-%d')}_{task_desc[:20].replace(' ', '_')}"
    st.session_state["current_task"] = TaskRun(
        task_id=task_id,
        task_description=task_desc,
        template_id=selected_id,
        template_mtime=meta.file_mtime if meta else 0,
    )

task: TaskRun = st.session_state["current_task"]

# --- Step 3: Generate prompt ---
st.subheader("Step 3: 生成 Prompt → 复制到 AI Web")
if meta:
    prompt = build_text_prompt(meta, task_desc, config.candidates_per_element)
    task.text_prompt = prompt
    task.status = "prompting"
    task.current_step = 3

    st.code(prompt, language="markdown")
    if st_copy_to_clipboard:
        st_copy_to_clipboard(prompt, "复制 Prompt")
    else:
        st.info("安装 st-copy-to-clipboard 后可一键复制。手动选中上方文本复制。")

# --- Step 4: Paste JSON ---
st.subheader("Step 4: 粘贴 AI 返回的 JSON")
raw_json = st.text_area("粘贴 AI 返回结果", height=200, key="raw_json_input")

if not raw_json:
    st.stop()

parsed = parse_ai_json(raw_json)
if parsed is None:
    st.error("无法解析 JSON。请检查格式后重试。原文已保存。")
    task.ai_response_raw = raw_json
    save_raw_response(task.task_id, raw_json)
    st.stop()

task.ai_response_raw = raw_json
task.text_candidates = TextCandidates(
    template_id=selected_id,
    task_description=task_desc,
    candidates=parsed,
)
task.status = "selecting"
task.current_step = 5

# --- Step 5: Pick candidates ---
st.subheader("Step 5: 挑选备选")
choices = render_candidate_picker(task.text_candidates, meta)
task.text_choices = choices

# --- Step 6: Generate preview ---
st.subheader("Step 6: 生成预览")
if st.button("生成预览"):
    task.status = "rendering"
    task.current_step = 6

    runs_dir = resolve_path(config.runs_dir) / task.task_id
    runs_dir.mkdir(parents=True, exist_ok=True)

    output_pptx = str(runs_dir / "output.pptx")
    fill_template(meta.file_path, meta, task.text_choices, output_pptx)
    task.output_pptx = output_pptx

    preview_png = str(runs_dir / "preview.png")
    render_slide_to_png(output_pptx, preview_png)
    task.preview_image = preview_png

    task.status = "completed"
    task.current_step = 7
    save_task(task)

    st.image(preview_png, caption="预览", use_container_width=True)

# --- Step 7: Download ---
if task.output_pptx and Path(task.output_pptx).exists():
    st.subheader("Step 7: 导出")
    with open(task.output_pptx, "rb") as f:
        st.download_button(
            "下载 .pptx",
            data=f,
            file_name=f"{task.task_id}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

# --- Auto-save (Task 11) ---
@st.fragment(run_every=config.auto_save_interval_sec)
def auto_save():
    if "current_task" in st.session_state:
        save_task(st.session_state["current_task"])

auto_save()
