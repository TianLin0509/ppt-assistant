"""Task history: view, resume, delete tasks."""

import json
import shutil

import streamlit as st
from pathlib import Path

from src.utils.config import load_config, resolve_path
from src.schema import TaskRun

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

STATUS_LABELS = {
    "created": "已创建",
    "prompting": "生成 Prompt 中",
    "selecting": "挑选备选中",
    "rendering": "渲染预览中",
    "completed": "已完成",
    "failed": "失败",
}

for d in task_dirs:
    task_json = d / "task.json"

    task_data = None
    status = "未知"
    template_id = ""
    description = d.name
    if task_json.exists():
        try:
            data = json.loads(task_json.read_text(encoding="utf-8"))
            task_data = TaskRun(**data)
            status = STATUS_LABELS.get(task_data.status, task_data.status)
            template_id = task_data.template_id
            description = task_data.task_description
        except Exception:
            status = "数据损坏"

    status_icon = "✅" if task_data and task_data.status == "completed" else "🔄" if task_data and task_data.status in ("selecting", "prompting", "rendering") else "📝"

    with st.expander(f"{status_icon} {d.name} — {status}"):
        if task_data:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"**任务描述**: {description}")
                st.markdown(f"**模板**: {template_id}")
                st.markdown(f"**状态**: {status} (Step {task_data.current_step}/7)")
                st.markdown(f"**创建时间**: {task_data.created_at[:19]}")
            with c2:
                previews = sorted(d.glob("preview_*.png"))
                variants = sorted(d.glob("variant_*.pptx"))
                if previews:
                    st.image(str(previews[0]), use_container_width=True, caption="预览")
                elif (d / "preview.png").exists():
                    st.image(str(d / "preview.png"), use_container_width=True)

        # Download buttons for completed variants
        variant_files = sorted(d.glob("variant_*.pptx"))
        if variant_files:
            cols = st.columns(len(variant_files))
            for i, vf in enumerate(variant_files):
                with cols[i]:
                    with open(str(vf), "rb") as f:
                        st.download_button(
                            f"下载 {vf.stem}",
                            data=f,
                            file_name=f"{d.name}_{vf.stem}.pptx",
                            key=f"dl_{d.name}_{vf.stem}",
                        )
        elif (d / "output.pptx").exists():
            with open(str(d / "output.pptx"), "rb") as f:
                st.download_button("下载 .pptx", data=f, file_name=f"{d.name}.pptx", key=f"dl_{d.name}")

        # Action buttons
        bc1, bc2 = st.columns(2)
        with bc1:
            if task_data and task_data.status != "completed":
                if st.button("继续任务", key=f"resume_{d.name}"):
                    st.session_state["current_task"] = task_data
                    st.session_state.pop("previews_generated", None)
                    st.session_state.pop("selected_variant", None)
                    st.switch_page("pages/2_new_task.py")
        with bc2:
            if st.button("删除任务", key=f"delete_{d.name}", type="secondary"):
                st.session_state[f"confirm_delete_{d.name}"] = True

            if st.session_state.get(f"confirm_delete_{d.name}"):
                st.warning(f"确定删除 {d.name}？此操作不可恢复。")
                dc1, dc2 = st.columns(2)
                with dc1:
                    if st.button("确认删除", key=f"yes_delete_{d.name}", type="primary"):
                        shutil.rmtree(str(d))
                        st.session_state.pop(f"confirm_delete_{d.name}", None)
                        st.rerun()
                with dc2:
                    if st.button("取消", key=f"no_delete_{d.name}"):
                        st.session_state.pop(f"confirm_delete_{d.name}", None)
                        st.rerun()
