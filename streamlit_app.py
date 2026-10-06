# -*- coding: utf-8 -*-
"""鸭芯智选 V5.0 —— 入口页：专家模式 / 简单模式切换"""

import streamlit as st

st.set_page_config(
    page_title="鸭芯智选 V5.0",
    page_icon="🦆",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🦆 鸭芯智选 · 肉鸭智能化采食行为分析工具 V5.0")
st.caption("四川农业大学 张雅晨团队 · 大挑项目 · 留种决策模块")

# URL 参数支持：?mode=simple / ?mode=expert
qp = st.query_params
default_idx = 1 if qp.get("mode") == "simple" else 0

with st.sidebar:
    st.markdown("### 🎛️ 工作模式")
    mode = st.radio(
        "选择模式",
        ["👨‍🔬 专家模式（六大模块全功能）",
         "🧑‍🌾 简单模式（一键选留种鸭）"],
        index=default_idx,
        key="app_mode",
        label_visibility="collapsed",
    )
    st.divider()

st.query_params["mode"] = "simple" if mode.startswith("🧑‍🌾") else "expert"

is_expert = mode.startswith("👨‍🔬")

page = st.navigation(
    [st.Page(
        "page_expert.py" if is_expert else "page_simple.py",
        title="专家模式" if is_expert else "简单模式",
        icon="👨‍🔬" if is_expert else "🧑‍🌾",
        default=True,
    )],
    position="hidden",
)

page.run()
