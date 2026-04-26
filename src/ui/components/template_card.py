"""Template thumbnail card component."""

from __future__ import annotations

from pathlib import Path

import streamlit as st

from src.schema import TemplateMeta


def render_template_card(meta: TemplateMeta, col) -> bool:
    with col:
        if meta.annotated_image and Path(meta.annotated_image).exists():
            st.image(meta.annotated_image, use_container_width=True)
        elif meta.preview_image and Path(meta.preview_image).exists():
            st.image(meta.preview_image, use_container_width=True)
        else:
            st.info("无预览图")

        st.caption(meta.template_id)
        n_text = len(meta.editable_text_elements)
        n_img = len(meta.editable_image_elements)
        st.caption(f"文本: {n_text} | 图片: {n_img}")

        return st.button("选择", key=f"select_{meta.template_id}")
