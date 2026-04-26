"""Streamlit main entry point."""

import streamlit as st

st.set_page_config(
    page_title="PPT Assistant",
    page_icon="📊",
    layout="wide",
)

st.title("PPT Assistant — 单页填充助手")
st.markdown("从左侧导航选择功能页面。")
st.markdown("""
- **模板管理**: 导入模板、预处理、标注元素角色
- **新建任务**: 选模板 → 生成 prompt → 粘贴 AI 结果 → 挑选 → 导出
- **历史任务**: 查看过往任务记录
""")
