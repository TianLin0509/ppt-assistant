"""Template library: scan, preprocess, review roles."""

import streamlit as st
from pathlib import Path

from src.utils.config import load_config, resolve_path
from src.utils.state_manager import save_template_meta, load_template_meta, list_template_metas
from src.core.pptx_parser import parse_template
from src.core.pptx_renderer import render_slide_to_png
from src.core.shape_annotator import annotate_preview
from src.core.role_inferencer import infer_roles
from src.schema import ShapeType, TextSubtype

st.set_page_config(page_title="模板管理", layout="wide")
st.header("模板管理")

config = load_config()
tpl_dir = resolve_path(config.templates_dir)
meta_dir = resolve_path(config.templates_meta_dir)

if not tpl_dir.exists():
    st.warning(f"模板目录不存在: {tpl_dir}")
    st.stop()

pptx_files = sorted(tpl_dir.glob("*.pptx"))
pptx_files = [f for f in pptx_files if not f.name.startswith("~$")]

if not pptx_files:
    st.info("templates/ 目录下没有 .pptx 文件。请先放入模板。")
    st.stop()

st.subheader(f"找到 {len(pptx_files)} 个模板")

cols = st.columns(3)
for i, pptx_file in enumerate(pptx_files):
    template_id = pptx_file.stem.replace(" ", "_")
    meta = load_template_meta(template_id)

    with cols[i % 3]:
        st.markdown(f"**{pptx_file.name}**")

        if meta and meta.annotated_image and Path(meta.annotated_image).exists():
            st.image(meta.annotated_image, use_container_width=True)
            st.success(f"已预处理 ({len(meta.editable_text_elements)} 文本, {len(meta.editable_image_elements)} 图片)")
        else:
            st.info("未预处理")

        if st.button("预处理", key=f"preprocess_{template_id}"):
            with st.spinner("解析中..."):
                meta = parse_template(str(pptx_file))

                preview_path = str(meta_dir / f"{template_id}_preview.png")
                render_slide_to_png(str(pptx_file), preview_path)
                meta.preview_image = preview_path

                meta.elements = infer_roles(meta.elements, meta.slide_width_emu, meta.slide_height_emu)

                annotated_path = annotate_preview(
                    preview_path, meta.elements,
                    meta.slide_width_emu, meta.slide_height_emu,
                )
                meta.annotated_image = annotated_path

                save_template_meta(meta)
                st.success("预处理完成!")
                st.rerun()

st.divider()

st.subheader("角色修正")
all_metas = list_template_metas()
if not all_metas:
    st.info("暂无已预处理的模板。请先点击上方「预处理」。")
    st.stop()

selected_id = st.selectbox("选择模板", [m.template_id for m in all_metas])
meta = load_template_meta(selected_id)
if meta is None:
    st.stop()

if meta.annotated_image and Path(meta.annotated_image).exists():
    st.image(meta.annotated_image, use_container_width=True, caption="标注图（红框编号 = 元素序号）")

ROLE_OPTIONS = [
    "", "title_main", "subtitle", "body_01", "body_02", "body_03",
    "bullet_01", "bullet_02", "bullet_03",
    "image_01", "image_02", "image_03", "decoration",
]
SUBTYPE_OPTIONS = ["", "title", "bullet", "body"]

changed = False
for idx, el in enumerate(meta.elements):
    if el.type in (ShapeType.GROUP, ShapeType.SMARTART, ShapeType.CHART, ShapeType.TABLE):
        st.markdown(f"**#{idx}** {el.shape_name_original} — :red[{el.type.value}，不可编辑]")
        continue

    c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
    with c1:
        st.text(f"#{idx} {el.shape_name_original}")
        if el.current_content:
            st.caption(el.current_content[:60])
    with c2:
        current_role = el.role_key or ""
        new_role = st.selectbox(
            "角色", ROLE_OPTIONS,
            index=ROLE_OPTIONS.index(current_role) if current_role in ROLE_OPTIONS else 0,
            key=f"role_{idx}",
        )
        if new_role != current_role:
            el.role_key = new_role or None
            changed = True
    with c3:
        current_sub = el.text_subtype.value if el.text_subtype else ""
        new_sub = st.selectbox(
            "子类型", SUBTYPE_OPTIONS,
            index=SUBTYPE_OPTIONS.index(current_sub) if current_sub in SUBTYPE_OPTIONS else 0,
            key=f"subtype_{idx}",
        )
        if new_sub != current_sub:
            el.text_subtype = TextSubtype(new_sub) if new_sub else None
            changed = True
    with c4:
        el.role_confirmed = st.checkbox("确认", value=el.role_confirmed, key=f"confirm_{idx}")

if st.button("保存角色配置"):
    save_template_meta(meta)
    st.success("已保存!")
