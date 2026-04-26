"""Candidate selection component with radio + inline edit."""

from __future__ import annotations

import streamlit as st

from src.schema import TextCandidates, TemplateMeta


def render_candidate_picker(
    candidates: TextCandidates,
    meta: TemplateMeta,
) -> dict[str, str]:
    elements = meta.editable_text_elements
    role_elements = {e.role_key: e for e in elements if e.role_key}

    if st.button("一键全选 A"):
        for role_key in candidates.candidates:
            st.session_state[f"pick_{role_key}"] = "A"
        st.rerun()

    choices: dict[str, str] = {}

    for role_key, options in candidates.candidates.items():
        el = role_elements.get(role_key)
        label = el.display_label if el else role_key

        st.markdown(f"#### {label}")

        option_labels = [chr(65 + i) for i in range(len(options))]

        if f"pick_{role_key}" not in st.session_state:
            st.session_state[f"pick_{role_key}"] = "A"

        choice = st.radio(
            f"选择 {role_key}",
            option_labels,
            key=f"pick_{role_key}",
            horizontal=True,
            label_visibility="collapsed",
        )

        choice_idx = ord(choice) - 65
        selected_text = options[choice_idx] if choice_idx < len(options) else options[0]

        edit_key = f"text_{role_key}_{choice}"
        if edit_key not in st.session_state:
            st.session_state[edit_key] = selected_text

        edited = st.text_area(
            f"编辑 {role_key}",
            value=st.session_state[edit_key],
            key=f"ta_{role_key}_{choice}",
            label_visibility="collapsed",
            height=80,
        )
        st.session_state[edit_key] = edited
        choices[role_key] = edited

        st.divider()

    return choices
