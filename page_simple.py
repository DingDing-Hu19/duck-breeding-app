# -*- coding: utf-8 -*-
"""简单模式 —— 三步选出好鸭（算法来自 duck_algo）"""

from datetime import date, timedelta

import pandas as pd
import streamlit as st
import plotly.express as px

from duck_algo import *   # noqa: F401,F403


def render_plotly(fig, key):
    st.plotly_chart(fig, use_container_width=True, key=key)


if "simple_res" not in st.session_state:
    st.session_state.simple_res = None


st.title("🦆 鸭芯智选 · 简单模式")
st.caption("三步选出好鸭：传数据 → 选目标 → 拿结果。不用懂遗传育种，软件自动完成清洗、筛指标、打分、排名。")

with st.sidebar:
    st.header("① 上传数据")
    feed_files = st.file_uploader("采食表（必传，可多选）", type=["xls", "xlsx"],
                                   accept_multiple_files=True, key="simple_feed")
    bw_files = st.file_uploader("体重表（必传，可多选）", type=["xls", "xlsx"],
                                 accept_multiple_files=True, key="simple_bw")
    ped_file = st.file_uploader("系谱（可选，有系谱更准）", type=["xls", "xlsx"], key="simple_ped")
    idmap_file = st.file_uploader("ID对照表（可选）", type=["xls", "xlsx"], key="simple_idmap")

    st.divider()
    st.header("时间范围（可选）")
    auto_time = st.checkbox("自动使用数据全部时间（默认）", value=True, key="simple_auto_time")
    exp_start = None
    if not auto_time:
        exp_start = st.date_input("试验开始日期", value=date.today() - timedelta(days=30),
                                   key="simple_exp_start")

    st.divider()
    st.header("② 育种目标（可多选）")
    goal_save = st.checkbox("吃得省（料重比低、省饲料）", value=True, key="simple_goal_save")
    goal_grow = st.checkbox("长得快（日增重高、出栏重）", value=True, key="simple_goal_grow")
    goal_fat = st.checkbox("皮脂厚（烤鸭用，未测自动用文献关联指标）", value=False, key="simple_goal_fat")
    goal_meat = st.checkbox("胸肌厚（瘦肉多，未测自动用文献关联指标）", value=False, key="simple_goal_meat")
    goal_rhythm = st.checkbox("采食规律（节律稳定）", value=False, key="simple_goal_rhythm")

    st.divider()
    st.header("③ 留种设置")
    ratio = st.slider("留种比例", min_value=5, max_value=80, value=30, step=1,
                       format="%d %%", key="simple_ratio")
    sex_balance = st.checkbox("尽量公母均衡", value=True, key="simple_sex_balance")

    st.divider()
    run_btn = st.button("🦆 开始选留种鸭", type="primary", use_container_width=True)

    if ped_file is None:
        st.caption("未上传系谱：将自动采用权威文献遗传参数（不影响使用）")
    else:
        st.caption("已上传系谱：将用本场数据实测遗传力，评估更准")


def collect_goals():
    g = []
    if goal_save:
        g.append("save")
    if goal_grow:
        g.append("grow")
    if goal_fat:
        g.append("fat")
    if goal_meat:
        g.append("meat")
    if goal_rhythm:
        g.append("rhythm")
    return g or ["save"]


if run_btn:
    if not feed_files or not bw_files:
        st.error("请上传采食表与体重表。")
    else:
        try:
            with st.spinner("简单模式计算中（约 1-2 分钟）……"):
                st.session_state.simple_res = run_simple_all(
                    feed_files, bw_files, ped_file, idmap_file,
                    goals=collect_goals(), ratio=ratio / 100,
                    sex_balance=sex_balance, auto_time=auto_time,
                    experiment_start=exp_start)
            st.success("计算完成！请查看右侧结果。")
        except Exception as e:
            st.error(f"计算失败：{e}")

res = st.session_state.simple_res

if res is None:
    st.info("👈 请在左侧上传数据、选择育种目标，然后点击『开始选留种鸭』。")
    st.stop()

warn_msgs = list(res["A"].get("warnings") or [])
anomalies = res["C"]["anomalies"]
if len(anomalies):
    warn_msgs.append(f"检出 {len(anomalies)} 条异常记录，见『采食行为』页（暂缓复测，不淘汰）")
if warn_msgs:
    st.warning("\n".join(f"⚠ {m}" for m in warn_msgs))

tabs = st.tabs(["🦆 选留种鸭", "📊 生产报表", "🔍 采食行为"])

with tabs[0]:
    st.subheader("选留结果概要")
    summary = res["A"]["summary"].copy()
    cols = st.columns(min(4, len(summary)))
    for i, (_, row) in enumerate(summary.iterrows()):
        cols[i % len(cols)].metric(str(row["指标"]), str(row["值"]))
    st.caption(f"权重策略：{res['A']['weights_note']}")

    st.subheader(f"留种名单（前 {round(res['A']['ratio'] * 100)}%）")
    st.dataframe(res["A"]["retained"], use_container_width=True, height=420)

    st.subheader("综合得分分布")
    idx = res["A"]["all_index"].copy()
    keep_ids = set(res["A"]["retained"]["Animal_ID"])
    idx["Is_Retained"] = idx["Animal_ID"].isin(keep_ids)
    thr = idx.loc[idx["Is_Retained"], "Index"].min()
    fig = px.histogram(idx, x="Index", color="Is_Retained", nbins=40,
                       color_discrete_map={True: PAL_RPBG[0], False: PAL_RPBG[1]})
    fig.add_vline(x=thr, line_dash="dash", line_color="black",
                  annotation_text=f"留种阈值（前 {round(res['A']['ratio'] * 100)}%）")
    fig.update_layout(xaxis_title="综合得分", yaxis_title="个体数")
    render_plotly(fig, "simple_index_plot")

    st.subheader("指示性状（自动筛选）")
    ind = res["A"]["indicator_traits"]
    if ind is not None and len(ind):
        st.dataframe(ind.round(3), use_container_width=True)
    else:
        st.caption("本次未筛出额外指示性状（目标性状本身已足够）。")

    xlsx_buf = export_simple_xlsx(res)
    st.download_button("📥 导出全部结果（Excel）", data=xlsx_buf,
                        file_name=f"简单模式_结果_{date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tabs[1]:
    st.subheader("周统计 · 采食")
    st.dataframe(res["B"]["week_feed"], use_container_width=True, height=300)

    st.subheader("周统计 · 体重")
    st.dataframe(res["B"]["week_bw"], use_container_width=True, height=300)

    st.subheader("每周 FCR")
    wf = res["B"]["weekly_fcr"]
    if len(wf):
        agg = wf.groupby("Week_Number").agg(Weekly_FCR=("Weekly_FCR", "mean")).reset_index()
        fig = px.bar(agg, x="Week_Number", y="Weekly_FCR", text_auto=".2f")
        fig.update_traces(marker_color=PAL_RPBG[1])
        fig.update_layout(xaxis_title="周次", yaxis_title="每周 FCR")
        render_plotly(fig, "simple_fcr_plot")
        st.dataframe(wf.round(3), use_container_width=True, height=280)

    st.subheader("周平均体重（生长曲线）")
    g = res["B"]["growth"]
    if len(g):
        fig = px.line(g, x="Week", y="Mean_BW", markers=True)
        fig.update_traces(line_color=PAL_RPBG[0])
        fig.update_layout(xaxis_title="周次", yaxis_title="平均体重(kg)")
        render_plotly(fig, "simple_growth_plot")

with tabs[2]:
    st.subheader("日访饲节律")
    d = res["C"]["rhythm_daily"]
    if len(d):
        fig = px.line(d, x="Date", y="Mean_Bouts_Per_Duck", markers=True)
        fig.update_traces(line_color=PAL_RPBG[2])
        fig.update_layout(xaxis_title="日期", yaxis_title="平均访饲次数/只/天")
        render_plotly(fig, "simple_daily_plot")

    st.subheader("24 小时访饲节律（按周）")
    h = res["C"]["rhythm_hourly"]
    if len(h):
        fig = px.line(h, x="Hour_Block", y="Mean_Bouts_Per_Duck",
                      color="Week_Number", markers=True,
                      color_discrete_sequence=PAL_RPBG * 5)
        fig.update_xaxes(tickmode="array", tickvals=list(range(24)))
        fig.update_layout(xaxis_title="时刻 (h)", yaxis_title="平均访问次数/只", color="周次")
        render_plotly(fig, "simple_hourly_plot")

    st.subheader("昼夜分配")
    st.dataframe(res["C"]["daynight"].round(3), use_container_width=True, height=280)

    st.subheader("行为变异性")
    st.dataframe(res["C"]["cv"].round(3), use_container_width=True, height=280)

    st.subheader("异常个体预警（暂缓复测，不淘汰）")
    an = res["C"]["anomalies"]
    if len(an):
        st.dataframe(an, use_container_width=True, height=300)
    else:
        st.success("未检出异常个体 ✅")
