"""New task workflow: select template -> generate prompt -> paste JSON -> preview 3 candidates -> export."""

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
from src.core.pptx_renderer import render_slide_to_png, render_slides_to_png
from src.schema import TaskRun, TextCandidates

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

# --- Step 5: Generate 3 preview variants ---
st.subheader("Step 5: 生成三套预览 → 整体对比选择")

candidates = task.text_candidates.candidates
n_variants = min(config.candidates_per_element, min(len(v) for v in candidates.values()) if candidates else 3)
variant_labels = [chr(65 + i) for i in range(n_variants)]

runs_dir = resolve_path(config.runs_dir) / task.task_id
runs_dir.mkdir(parents=True, exist_ok=True)

if st.button("生成三套预览"):
    with st.spinner("正在生成预览（共 3 套，请稍候）..."):
        render_items = []
        for vi in range(n_variants):
            choices_vi = {}
            for role_key, options in candidates.items():
                idx = min(vi, len(options) - 1)
                choices_vi[role_key] = options[idx]

            pptx_path = str(runs_dir / f"variant_{variant_labels[vi]}.pptx")
            fill_template(meta.file_path, meta, choices_vi, pptx_path)

            png_path = str(runs_dir / f"preview_{variant_labels[vi]}.png")
            render_items.append((pptx_path, png_path))

        render_slides_to_png(render_items)
        st.session_state["previews_generated"] = True
    st.rerun()

if st.session_state.get("previews_generated"):
    tab_names = [f"方案 {label}" for label in variant_labels]
    tabs = st.tabs(tab_names)
    for vi, tab in enumerate(tabs):
        label = variant_labels[vi]
        png_path = runs_dir / f"preview_{label}.png"
        with tab:
            if png_path.exists():
                st.image(str(png_path), use_container_width=True)
            else:
                st.error(f"预览图未生成: {png_path}")

            with st.expander("查看文案详情"):
                for role_key, options in candidates.items():
                    idx = min(vi, len(options) - 1)
                    el = next((e for e in meta.editable_text_elements if e.role_key == role_key), None)
                    label_text = el.display_label if el else role_key
                    st.markdown(f"**{label_text}**")
                    st.caption(options[idx])

            if st.button(f"选择方案 {label}", key=f"select_variant_{label}", type="primary"):
                for role_key, options in candidates.items():
                    idx = min(vi, len(options) - 1)
                    task.text_choices[role_key] = options[idx]
                task.output_pptx = str(runs_dir / f"variant_{label}.pptx")
                task.preview_image = str(png_path)
                task.status = "completed"
                task.current_step = 7
                save_task(task)
                st.session_state["selected_variant"] = label
                st.rerun()

# --- Step 6: Download ---
selected = st.session_state.get("selected_variant")
if selected and task.output_pptx and Path(task.output_pptx).exists():
    st.subheader(f"Step 6: 导出（方案 {selected}）")
    st.image(task.preview_image, caption=f"方案 {selected} 预览", use_container_width=True)
    with open(task.output_pptx, "rb") as f:
        st.download_button(
            f"下载方案 {selected} .pptx",
            data=f,
            file_name=f"{task.task_id}_{selected}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

# --- Auto-save ---
@st.fragment(run_every=config.auto_save_interval_sec)
def auto_save():
    if "current_task" in st.session_state:
        save_task(st.session_state["current_task"])

auto_save()
