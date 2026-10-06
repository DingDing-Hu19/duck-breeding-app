# -*- coding: utf-8 -*-
"""专家模式 —— 六大模块 UI（算法来自 duck_algo）"""

import io
from datetime import date, timedelta

import pandas as pd
import streamlit as st
import plotly.express as px

from duck_algo import *   # noqa: F401,F403 —— 所有算法函数


# ─────────── UI 辅助 ───────────
def fig_to_download(fig, fname):
    buf = io.BytesIO()
    fig.write_image(buf, format="png", width=1000, height=600, scale=2)
    buf.seek(0)
    return buf, fname


def render_plotly(fig, key, filename):
    st.plotly_chart(fig, use_container_width=True, key=key)
    try:
        buf, fn = fig_to_download(fig, filename)
        st.download_button("下载PNG", data=buf, file_name=fn, mime="image/png",
                           key=f"dl_{key}")
    except Exception:
        html = fig.to_html(include_plotlyjs="cdn")
        st.download_button("下载HTML", data=html,
                           file_name=filename.replace(".png", ".html"),
                           mime="text/html", key=f"dl_{key}")


def df_download(df, fname, key):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    st.download_button("下载本板块 Excel", data=buf, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=key)


def df_show(df, height=400):
    st.dataframe(df, use_container_width=True, height=height)


# ─────────── session_state ───────────
for k, v in [("pipe", None), ("ml", None), ("h2_est", None)]:
    if k not in st.session_state:
        st.session_state[k] = v


tabs = st.tabs([
    "① 数据与清洗",
    "② 个体指标",
    "③ 组间时序相关",
    "④ 创新行为指标",
    "⑤ 遗传评估与留种",
    "⑥ AI 智能选种探索",
])

# ============================================================
# 模块① 数据与清洗
# ============================================================
with tabs[0]:
    st.header("① 数据与清洗")
    col1, col2 = st.columns([1, 2])

    with col1:
        feed_files = st.file_uploader("采食原始数据（可多选，自动读取全部Sheet）",
                                       type=["xls", "xlsx"], accept_multiple_files=True,
                                       key="feed_files")
        bw_files = st.file_uploader("体重原始数据（可多选）",
                                     type=["xls", "xlsx"], accept_multiple_files=True,
                                     key="bw_files")
        ped_file = st.file_uploader("系谱（可选：翅号/父本笼号/母本笼号/性别）",
                                     type=["xls", "xlsx"], key="ped_file")
        idmap_file = st.file_uploader("eID-ID对照表（可选）",
                                       type=["xls", "xlsx"], key="idmap_file")
        st.info("系谱与对照表为可选。不上传系谱时，软件自动使用『文献遗传参数』路线完成留种决策。")

        st.subheader("清洗参数")
        auto_time = st.checkbox("自动读取数据最早/最晚时间（推荐）", value=True)
        exp_start = st.date_input("试验开始日期", value=date.today() - timedelta(days=30))
        week_length = st.number_input("周长度（天）", value=7, min_value=1, max_value=30)
        feed_sd = st.number_input("采食异常阈值（均值±k倍SD）", value=3.0, min_value=1.0, max_value=10.0)
        bw_sd = st.number_input("体重异常阈值（均值±k倍SD）", value=3.0, min_value=1.0, max_value=10.0)
        imi_threshold = st.number_input("餐间隔阈值（秒）", value=300, min_value=30)
        min_intake = st.number_input("最小单次采食量（g）", value=1.0, min_value=0.0)

        if st.button("运行数据清洗", type="primary", use_container_width=True):
            if not feed_files or not bw_files:
                st.error("请上传采食与体重数据。")
            else:
                try:
                    cfg = {
                        "auto_time_range": auto_time,
                        "experiment_start": exp_start,
                        "training_end": None, "training_time": None,
                        "week_length": int(week_length),
                        "feed_sd": float(feed_sd), "bw_sd": float(bw_sd),
                        "imi_threshold": float(imi_threshold),
                        "min_intake": float(min_intake),
                        "hff_method": "median", "hff_cutoff": None,
                        "day_start": "06:00", "day_end": "18:00",
                    }
                    with st.spinner("数据读取与清洗中…"):
                        st.session_state.pipe = run_pipeline_v5(
                            feed_files, bw_files, ped_file, idmap_file, cfg)
                    st.success(f"清洗完成：采食 {len(st.session_state.pipe['clean']['feed'])} 行，"
                               f"体重 {len(st.session_state.pipe['clean']['bw'])} 行，"
                               f"耗时 {st.session_state.pipe['elapsed']:.1f} 秒")
                except Exception as e:
                    st.error(f"清洗失败：{e}")

    with col2:
        if st.session_state.pipe is not None:
            p = st.session_state.pipe
            st.subheader("工作表读取情况")
            df_show(pd.concat([p["feed_raw_qc"], p["bw_raw_qc"]], ignore_index=True), 300)

            st.subheader("时间范围 QC")
            dq = p["clean"]["datetime_qc"]
            st.info(f"数据时间范围：{dq['Start'].iloc[0]} ～ {dq['End'].iloc[0]}，"
                    f"共 {dq['N'].iloc[0]} 条采食记录。")
            df_show(dq, 100)

            st.subheader("字段缺失 QC")
            df_show(p["clean"]["na_qc"], 300)

            st.subheader("周统计（采食）")
            df_show(p["clean"]["feed_week_stats"], 200)

            st.subheader("周统计（体重）")
            df_show(p["clean"]["bw_week_stats"], 200)

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                pd.concat([p["feed_raw_qc"], p["bw_raw_qc"]], ignore_index=True).to_excel(
                    w, sheet_name="工作表读取", index=False)
                p["clean"]["datetime_qc"].to_excel(w, sheet_name="时间范围QC", index=False)
                p["clean"]["na_qc"].to_excel(w, sheet_name="字段缺失QC", index=False)
                p["clean"]["feed_week_stats"].to_excel(w, sheet_name="周统计_采食", index=False)
                p["clean"]["bw_week_stats"].to_excel(w, sheet_name="周统计_体重", index=False)
            buf.seek(0)
            st.download_button("下载本板块 Excel", data=buf,
                               file_name=f"鸭芯智选_模块1_清洗结果_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# 模块② 个体指标
# ============================================================
with tabs[1]:
    st.header("② 个体指标")
    if st.session_state.pipe is None:
        st.warning("请先在模块①运行数据清洗。")
    else:
        p = st.session_state.pipe
        max_n = st.number_input("表格最大行数", value=500, min_value=50, max_value=10000)

        st.subheader("采食行为")
        df_show(p["feeding"].round(2).head(int(max_n)), 400)

        st.subheader("生产性能")
        prod_cols = [c for c in ["Animal_ID", "Feed_Frequency_Group", "TFB", "TFB_Day", "FI_g", "FI_Day_g",
                                  "AMS_g", "TFD_sec", "AFBD_sec", "IMI_sec", "FR_g_sec",
                                  "IBW_kg", "FBW_kg", "Gain_kg", "ADG_g", "ADFI_g", "FCR", "RFI",
                                  "Feed_Efficiency"] if c in p["prod"].columns]
        df_show(p["prod"][prod_cols].round(2).head(int(max_n)), 400)

        bout = p["bout"].copy()
        feed = p["clean"]["feed"].copy()
        feed["Date_only"] = feed["Time1"].dt.date
        wk_map = feed[["Animal_ID", "Date_only", "Week_Number"]].drop_duplicates()
        bout["Date_only"] = pd.to_datetime(bout["Date"]).dt.date
        bout = bout.merge(wk_map, left_on=["Animal_ID", "Date_only"],
                          right_on=["Animal_ID", "Date_only"], how="left")

        def violin_by_week(df, ycol, ylabel, key):
            d = df[df[ycol].notna() & df["Week_Number"].notna()].copy()
            if len(d) == 0:
                st.info(f"没有可用于绘图的 {ylabel} 数据。")
                return
            d["Week_Label"] = "Week " + d["Week_Number"].astype(int).astype(str)
            fig = px.violin(d, x="Week_Label", y=ycol, color="Week_Label",
                            box=True, points=False,
                            color_discrete_sequence=PAL_RPBG * 5,
                            labels={ycol: ylabel, "Week_Label": "Week"})
            fig.update_layout(showlegend=False, height=500)
            render_plotly(fig, key, f"{key}.png")

        st.subheader("采食速率 FR（按周）")
        violin_by_week(bout, "FR_g_sec", "采食速率 (g/s)", "fr_violin")
        st.subheader("采食间隔 IMI（按周）")
        violin_by_week(bout, "IMI_sec", "采食间隔 (s)", "imi_violin")
        st.subheader("采食时长（按周）")
        violin_by_week(bout, "Feed_Duration_sec", "采食时长 (s)", "duration_violin")

        bw = p["clean"]["bw"].copy()
        bw["Date_only"] = bw["Time1"].dt.date
        bw_daily = bw[bw["BW_kg"].notna()].sort_values(["Animal_ID", "Time1"]).groupby(
            ["Animal_ID", "Date_only"]).last().reset_index()
        bw_daily["Exp_Day"] = (pd.to_datetime(bw_daily["Date_only"]) -
                                pd.to_datetime(bw_daily["Date_only"]).min()).dt.days
        overall = bw_daily.groupby("Exp_Day").agg(Mean_BW_kg=("BW_kg", "mean"),
                                                   N=("BW_kg", "size")).reset_index()
        st.subheader("总体生长曲线（每日终末体重均值）")
        fig = px.line(overall, x="Exp_Day", y="Mean_BW_kg", markers=True)
        fig.update_traces(line_color=PAL_RPBG[1])
        render_plotly(fig, "overall_growth_plot", "overall_growth.png")

        if "Feed_Frequency_Group" in p["prod"].columns:
            group_map = p["prod"][["Animal_ID", "Feed_Frequency_Group"]]
            bw_g = bw_daily.merge(group_map, on="Animal_ID", how="left")
            bw_g = bw_g[bw_g["Feed_Frequency_Group"].isin(["HFF", "LFF"])]
            grp = bw_g.groupby(["Feed_Frequency_Group", "Exp_Day"]).agg(
                Mean_BW_kg=("BW_kg", "mean"), N=("BW_kg", "size")).reset_index()
            st.subheader("HFF / LFF 分组生长曲线")
            fig = px.line(grp, x="Exp_Day", y="Mean_BW_kg", color="Feed_Frequency_Group",
                          color_discrete_map={"HFF": PAL_RPBG[0], "LFF": PAL_RPBG[1]},
                          markers=True)
            render_plotly(fig, "group_growth_plot", "group_growth.png")

        df_download(p["feeding"], f"模块2_采食行为_{date.today()}.xlsx", "dl_ind1")
        df_download(p["prod"], f"模块2_生产性能_{date.today()}.xlsx", "dl_ind2")

# ============================================================
# 模块③ 组间时序相关
# ============================================================
with tabs[2]:
    st.header("③ 组间时序相关")
    if st.session_state.pipe is None:
        st.warning("请先在模块①运行数据清洗。")
    else:
        p = st.session_state.pipe
        hff_method = st.radio("分组方法", ["中位数", "自定义阈值"], horizontal=True)
        cutoff_val = None
        if hff_method == "自定义阈值":
            cutoff_val = st.number_input("自定义阈值（次/天）", value=20.0, min_value=1.0)

        if st.button("运行组间分析", type="primary"):
            prod, c = assign_hff_lff(p["prod"],
                                      "custom" if hff_method == "自定义阈值" else "median",
                                      cutoff_val)
            p["prod"] = prod
            p["hff_cutoff"] = c
            st.session_state.pipe = p
            st.success(f"分组完成，阈值 = {c:.1f} 次/天")

        if "Feed_Frequency_Group" in p["prod"].columns:
            hff_tab = p["prod"].groupby("Feed_Frequency_Group").agg(
                N=("Animal_ID", "size"),
                TFB_Day=("TFB_Day", "mean"),
                FI_Day_g=("FI_Day_g", "mean"),
                ADG_g=("ADG_g", "mean"),
                FCR=("FCR", "mean"),
                RFI=("RFI", "mean"),
            ).reset_index().round(2)
            st.subheader("HFF/LFF 指标对比")
            df_show(hff_tab, 200)

            plot_df = p["prod"].melt(
                id_vars=["Animal_ID", "Feed_Frequency_Group"],
                value_vars=[c for c in ["TFB", "FR_g_sec", "FCR", "RFI"] if c in p["prod"].columns],
                var_name="Indicator", value_name="Value").dropna()
            fig = px.violin(plot_df, x="Feed_Frequency_Group", y="Value",
                            color="Feed_Frequency_Group", facet_col="Indicator",
                            facet_col_wrap=2, box=True,
                            color_discrete_map={"HFF": PAL_RPBG[0], "LFF": PAL_RPBG[1]},
                            height=600)
            fig.update_yaxes(matches=None)
            render_plotly(fig, "hff_plot", "hff_plot.png")

        st.subheader("访饲节律（日）")
        daily = p["rhythm"]["daily"]
        if len(daily):
            fig = px.line(daily, x="Date", y="Mean_Bouts_Per_Duck", markers=True)
            fig.update_traces(line_color=PAL_RPBG[1])
            render_plotly(fig, "rhythm_daily_plot", "rhythm_daily.png")

        st.subheader("访饲节律（24h）")
        hourly = p["rhythm"]["hourly"]
        if len(hourly):
            fig = px.line(hourly, x="Hour_Block", y="Mean_Bouts_Per_Duck",
                          color="Week_Number", markers=True,
                          color_discrete_sequence=PAL_RPBG * 5)
            fig.update_xaxes(tickmode="array", tickvals=list(range(24)))
            render_plotly(fig, "rhythm_hourly_plot", "rhythm_hourly.png")

        st.subheader("每周 FCR")
        wfcr = p["weekly_fcr"]
        if len(wfcr):
            df_show(wfcr.round(3).head(200), 300)
            wfcr["Week_Label"] = "Week " + wfcr["Week_Number"].astype(int).astype(str)
            fig = px.violin(wfcr, x="Week_Label", y="Weekly_FCR", color="Week_Label",
                            box=True, points=False, color_discrete_sequence=PAL_RPBG * 5)
            fig.update_layout(showlegend=False)
            render_plotly(fig, "weekly_fcr_plot", "weekly_fcr.png")

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            if "Feed_Frequency_Group" in p["prod"].columns:
                p["prod"].groupby("Feed_Frequency_Group").agg(
                    N=("Animal_ID", "size")).reset_index().to_excel(w, sheet_name="分组", index=False)
            p["weekly_fcr"].to_excel(w, sheet_name="每周FCR", index=False)
            p["prod"].to_excel(w, sheet_name="个体表_含分组", index=False)
        buf.seek(0)
        st.download_button("下载本板块 Excel", data=buf,
                           file_name=f"模块3_组间时序_{date.today()}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# 模块④ 创新行为指标
# ============================================================
with tabs[3]:
    st.header("④ 创新行为指标")
    if st.session_state.pipe is None:
        st.warning("请先在模块①运行数据清洗。")
    else:
        p = st.session_state.pipe
        c1, c2 = st.columns(2)
        with c1:
            day_start = st.text_input("白天开始(HH:MM)", "06:00")
        with c2:
            day_end = st.text_input("白天结束(HH:MM)", "18:00")

        if st.button("计算创新指标", type="primary"):
            with st.spinner("计算中…"):
                daynight = compute_daynight_v5(p["bout"], day_start, day_end)
                cv = compute_behavior_cv_v5(p["bout"])
                cosinor = compute_cosinor_v5(p["bout"])
                fano = compute_fano_v5(p["bout"])
                for df in [daynight, cv, cosinor, fano]:
                    cols_to_drop = [c for c in df.columns if c in p["prod"].columns and c != "Animal_ID"]
                    if cols_to_drop:
                        p["prod"] = p["prod"].drop(columns=cols_to_drop)
                    p["prod"] = p["prod"].merge(df, on="Animal_ID", how="left")
                st.session_state.pipe = p
            st.success("创新指标计算完成")

        prod = p["prod"]
        if "Day_FI_Ratio" in prod.columns:
            st.subheader("昼夜分配")
            dn_cols = ["Animal_ID", "Day_FI_g", "Night_FI_g", "Day_Bouts", "Night_Bouts",
                       "Day_FI_Ratio", "Day_Bout_Ratio"]
            df_show(prod[[c for c in dn_cols if c in prod.columns]].round(3).head(200), 300)

            fig = px.scatter(prod, x="Day_FI_Ratio", y="FCR", trendline="ols",
                             color_discrete_sequence=[PAL_RPBG[3]])
            render_plotly(fig, "daynight_scatter", "daynight.png")

        if "CV_Duration" in prod.columns:
            st.subheader("行为变异性 CV")
            cv_cols = ["Animal_ID", "CV_Duration", "CV_FR", "Robust_CV_IMI",
                       "CV_Daily_Bouts", "CV_Daily_FI", "CV_Daily_TFD"]
            df_show(prod[[c for c in cv_cols if c in prod.columns]].round(3).head(200), 300)
            plot_df = prod.melt(
                id_vars=["Animal_ID", "FCR"],
                value_vars=[c for c in ["CV_Duration", "CV_FR", "Robust_CV_IMI",
                                         "CV_Daily_Bouts", "CV_Daily_FI", "CV_Daily_TFD"]
                             if c in prod.columns],
                var_name="CV_Type", value_name="CV_Value").dropna()
            fig = px.scatter(plot_df, x="CV_Value", y="FCR", facet_col="CV_Type",
                             facet_col_wrap=3, trendline="ols",
                             color_discrete_sequence=[PAL_RPBG[3]])
            render_plotly(fig, "cv_fcr_scatter", "cv_fcr.png")

        if "Cosinor_A" in prod.columns:
            st.subheader("余弦节律拟合")
            co_cols = ["Animal_ID", "Cosinor_M", "Cosinor_A", "Peak_Hour", "Cosinor_R2"]
            df_show(prod[[c for c in co_cols if c in prod.columns]].round(3).head(200), 300)
            fig = px.scatter(prod, x="Cosinor_A", y="FCR", trendline="ols",
                             color_discrete_sequence=[PAL_RPBG[2]])
            render_plotly(fig, "cosinor_scatter", "cosinor.png")

            fig = px.histogram(prod, x="Peak_Hour", nbins=24,
                               color_discrete_sequence=[PAL_RPBG[4]])
            render_plotly(fig, "phase_plot", "phase.png")

        if "Fano" in prod.columns:
            st.subheader("Fano 指数")
            df_show(prod[["Animal_ID", "Fano"]].round(3).head(200), 300)

# ============================================================
# 模块⑤ 遗传评估与留种（核心）
# ============================================================
with tabs[4]:
    st.header("⑤ 遗传评估与留种（核心）")
    if st.session_state.pipe is None:
        st.warning("请先在模块①运行数据清洗。")
    else:
        p = st.session_state.pipe

        st.subheader("① 育种目标")
        target_options = ["FCR", "RFI", "ADG_g", "FBW_kg", "FI_Day_g", "FR_g_sec",
                           "TFB_Day", "AMS_g", "ADFI_g", "IBW_kg", "Gain_kg",
                           "Day_FI_Ratio", "SkinFat_Rate", "AbFat_Rate", "BMP"]
        target_trait = st.multiselect("目标性状（育种目标，可多选）", target_options, default=["FCR"])

        c1, c2 = st.columns(2)
        with c1:
            indicator_min_r = st.number_input("指示性状阈值(|r|≥)", value=0.3, min_value=0.0)
        with c2:
            indicator_max_n = st.number_input("指示性状上限", value=5, min_value=1, max_value=10)

        st.subheader("② 遗传评估路线")
        ped_present = st.session_state.pipe["ped"] is not None and len(st.session_state.pipe["ped"]) > 0
        use_ped = st.radio("评估路线",
                           ["路线A：PBLUP动物模型（有系谱）", "路线B：文献遗传参数选择指数（无系谱）"],
                           index=0 if ped_present else 1)
        fixed_effects = st.multiselect("固定效应", ["Sex"], default=["Sex"])
        use_reml_h2 = st.checkbox("路线A运行时自动实测本场 h²（失败回退输入值）", value=True)

        st.subheader("③ 选择指数性状与权重")
        st.info("指数 = Σ(各性状EBV × 方向 × 权重)。方向已自动处理：FCR/RFI/采食量等越低越好，"
                "增重/终重等越高越好，权重只需填正数。")
        efficiency = ["FCR", "RFI", "ADG_g", "FBW_kg", "Gain_kg", "IBW_kg", "MBW"]
        feeding_grp = ["FI_Day_g", "ADFI_g", "TFB_Day", "AMS_g", "FR_g_sec",
                       "TFD_sec", "AFBD_sec", "IMI_sec"]
        rhythm_grp = ["Day_FI_Ratio", "Day_Bout_Ratio", "CV_Duration", "CV_FR",
                      "CV_Daily_Bouts", "CV_Daily_FI", "Cosinor_A", "Fano"]
        c1, c2, c3 = st.columns(3)
        with c1:
            eff_sel = st.multiselect("效率与增重", efficiency, default=["FCR", "ADG_g", "FBW_kg"])
        with c2:
            feed_sel = st.multiselect("采食行为", feeding_grp, default=[])
        with c3:
            rhythm_sel = st.multiselect("节律与稳定性", rhythm_grp, default=[])
        index_traits = eff_sel + feed_sel + rhythm_sel

        lit = literature_params()
        h2_map = dict(zip(lit["Trait"], lit["h2"]))
        w_defaults = {"FCR": 0.5, "RFI": 0.3, "ADG_g": 0.3, "FBW_kg": 0.2, "FI_Day_g": 0.2,
                      "FR_g_sec": 0.2, "TFB_Day": 0.2, "AMS_g": 0.2, "IBW_kg": 0.2}
        weights = {}
        h2_by_trait = {}
        if index_traits:
            tm = trait_meta()
            dir_map = dict(zip(tm["Trait"], tm["Dir"]))
            lab_map = dict(zip(tm["Trait"], tm["Label"]))
            st.markdown("**权重 / h² 设置**")
            for tr in index_traits:
                c1, c2 = st.columns(2)
                with c1:
                    weights[tr] = st.number_input(f"{tr} 权重",
                                                    value=float(w_defaults.get(tr, 0.2)),
                                                    min_value=0.0, step=0.05, key=f"w_{tr}")
                with c2:
                    h2v = float(h2_map.get(tr, 0.3))
                    h2_by_trait[tr] = st.number_input(
                        f"{tr} h² (文献参考 {h2v})", value=round(h2v, 2),
                        min_value=0.05, max_value=0.95, step=0.01, key=f"h2_{tr}")
                d = dir_map.get(tr, 1)
                st.caption(f"{lab_map.get(tr, tr)} · 方向：{'↓ 越低越好' if d < 0 else '↑ 越高越好'}")

        st.subheader("④ 留种设置")
        c1, c2 = st.columns(2)
        with c1:
            retention_ratio = st.slider("留种比例（前 %）", 1, 90, 30)
        with c2:
            sex_balance = st.checkbox("留种时尽量性别均衡", value=True)

        if st.button("运行遗传评估与留种", type="primary", use_container_width=True):
            try:
                with st.spinner("筛选指示性状…"):
                    ind = select_indicator_traits(p["prod"], target_trait,
                                                   indicator_min_r, indicator_max_n)
                    p["indicator_traits"] = ind

                link = p["link"]
                n_ped = int(link["Has_Pedigree"].sum()) if "Has_Pedigree" in link.columns else 0
                actual_use_ped = (use_ped.startswith("路线A")) and (n_ped >= 20)

                if weights:
                    total = sum(weights.values())
                    weights = {k: v / total for k, v in weights.items()} if total > 0 else weights

                genetic = {"mode": "PBLUP" if actual_use_ped else "文献参数"}
                retention = None
                h2_est = None

                with st.spinner("遗传评估计算中…"):
                    if actual_use_ped:
                        ped_animals = link[link["Has_Pedigree"]][
                            ["ID", "Sire_Cage", "Dam_Cage", "Sex"]].copy()
                        id_map_df = link[["eID", "ID"]].drop_duplicates("eID")
                        prod_gen = p["prod"].merge(id_map_df, left_on="Animal_ID",
                                                     right_on="eID", how="left")
                        prod_gen["Wing_ID"] = prod_gen["ID"].fillna(prod_gen["Animal_ID"])
                        prod_gen["Wing_ID"] = prod_gen["Wing_ID"].astype(str)
                        ped_animals["ID"] = ped_animals["ID"].astype(str)

                        if use_reml_h2:
                            h2_rows = []
                            for tr in index_traits:
                                if tr not in prod_gen.columns:
                                    continue
                                try:
                                    r = estimate_h2_reml(prod_gen, ped_animals, tr,
                                                          fixed_effects, "Wing_ID")
                                    h2_rows.append(r)
                                except Exception:
                                    pass
                            if h2_rows:
                                h2_est = pd.DataFrame(h2_rows)
                                h2_by_trait = dict(zip(h2_est["trait"], h2_est["h2"]))

                        ebv_rows = []
                        for tr in index_traits:
                            if tr not in prod_gen.columns:
                                continue
                            try:
                                r = run_pblup(prod_gen, ped_animals, tr,
                                              h2=h2_by_trait.get(tr, 0.3),
                                              fixed_effects=fixed_effects, id_col="Wing_ID")
                                ebv_rows.append(r["result"].assign(Trait=tr))
                            except Exception as e:
                                st.warning(f"{tr} PBLUP 失败：{e}")
                        if ebv_rows:
                            ebv_df = pd.concat(ebv_rows, ignore_index=True)
                            genetic["ebv_tbl"] = ebv_df
                            idx = selection_index(ebv_df, weights)
                            retention = make_retention_list(idx, p["prod"], retention_ratio / 100,
                                                             sex_balance)
                        else:
                            actual_use_ped = False
                            genetic["mode"] = "文献参数"

                    if not actual_use_ped:
                        li = run_literature_index(p["prod"], target_trait, lit, weights,
                                                   traits_subset=index_traits or None)
                        genetic["ebv_tbl"] = li["ebv"]
                        genetic["literature"] = li
                        idx = selection_index(li["ebv"], weights)
                        retention = make_retention_list(idx, p["prod"], retention_ratio / 100,
                                                         sex_balance)

                genetic["index"] = idx
                genetic["index_traits"] = index_traits
                p["genetic"] = genetic
                p["retention"] = retention
                p["indicator_traits"] = ind
                if h2_est is not None:
                    p["h2_est"] = h2_est
                st.session_state.pipe = p
                st.success(f"评估完成：模式={genetic['mode']}，"
                           f"参与个体={len(idx)}，留种={retention['n_keep']} 只")
            except Exception as e:
                st.error(f"遗传评估失败：{e}")

        if p.get("genetic"):
            g = p["genetic"]
            r = p.get("retention")

            st.subheader("关键指标卡")
            metrics = {
                "目标性状": " + ".join(target_trait),
                "参与指数性状": " + ".join(g.get("index_traits", [])),
                "评估模式": g["mode"],
                "留种数": r["n_keep"] if r else "-",
                "留种比例": f"{retention_ratio}%",
            }
            st.json(metrics)

            st.subheader("① 指示性状筛选")
            ind = p.get("indicator_traits")
            if ind is not None and len(ind):
                df_show(ind.round(3), 300)
                if "Spearman_r" in ind.columns and ind["Spearman_r"].notna().any():
                    fig = px.bar(ind, x="Indicator", y="Spearman_r",
                                 color="Abs_r", color_continuous_scale="RdBu_r")
                    render_plotly(fig, "indicator_plot", "indicator.png")
                elif "Genetic_r" in ind.columns:
                    fig = px.bar(ind, x="Indicator", y="Genetic_r",
                                 color="Abs_r", color_continuous_scale="RdBu_r")
                    render_plotly(fig, "indicator_plot", "indicator.png")

            if st.session_state.h2_est is not None:
                st.subheader("② 遗传力估计（REML 实测）")
                df_show(st.session_state.h2_est.round(4), 200)

            st.subheader("③ 育种值（EBV）")
            if "ebv_tbl" in g:
                df_show(g["ebv_tbl"].round(4).head(500), 400)

            st.subheader("④ 选择指数分布")
            idx_df = g["index"]
            keep_ids = r["retained"]["Animal_ID"].tolist() if r else []
            idx_df = idx_df.copy()
            idx_df["Is_Retained"] = idx_df["Animal_ID"].isin(keep_ids)
            fig = px.histogram(idx_df, x="Index", color="Is_Retained", nbins=40,
                               color_discrete_map={True: PAL_RPBG[0], False: PAL_RPBG[1]})
            render_plotly(fig, "index_plot", "index.png")

            st.subheader("⑤ 留种名单（最终结果）")
            if r:
                df_show(r["retained"].round(2), 400)

            st.subheader("附表·文献参数表")
            df_show(literature_params(), 400)

            st.subheader("附表·文献遗传相关 rG")
            df_show(literature_corr(), 400)

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                if ind is not None and len(ind):
                    ind.to_excel(w, sheet_name="指示性状", index=False)
                if st.session_state.h2_est is not None:
                    st.session_state.h2_est.to_excel(w, sheet_name="遗传力实测REML", index=False)
                if "ebv_tbl" in g:
                    g["ebv_tbl"].to_excel(w, sheet_name="育种值EBV", index=False)
                if r:
                    r["retained"].to_excel(w, sheet_name="留种名单", index=False)
                literature_params().to_excel(w, sheet_name="文献参数表", index=False)
                literature_corr().to_excel(w, sheet_name="文献遗传相关", index=False)
            buf.seek(0)
            st.download_button("下载本板块 Excel", data=buf,
                               file_name=f"模块5_遗传评估与留种_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# 模块⑥ AI 智能选种探索
# ============================================================
with tabs[5]:
    st.header("⑥ AI 智能选种探索")
    if st.session_state.pipe is None:
        st.warning("请先在模块①运行数据清洗。")
    else:
        p = st.session_state.pipe

        st.info("用机器学习为『指示性状筛选』补充证据：识别非线性关联、量化特征重要性、验证稳定性。")
        ml_targets = st.multiselect("预测目标性状", ["ADG_g", "FCR", "RFI"],
                                     default=["ADG_g", "FCR", "RFI"])
        c1, c2, c3 = st.columns(3)
        with c1:
            w_adg = st.number_input("ADG 权重", 0.33, min_value=0.0, max_value=1.0, step=0.05)
        with c2:
            w_fcr = st.number_input("FCR 权重", 0.33, min_value=0.0, max_value=1.0, step=0.05)
        with c3:
            w_rfi = st.number_input("RFI 权重", 0.34, min_value=0.0, max_value=1.0, step=0.05)

        c1, c2 = st.columns(2)
        with c1:
            w_imp = st.number_input("重要性得分权重", 0.7, min_value=0.0, max_value=1.0, step=0.05)
        with c2:
            w_stab = st.number_input("稳定性得分权重", 0.3, min_value=0.0, max_value=1.0, step=0.05)

        if st.button("运行 AI 智能选种分析", type="primary", use_container_width=True):
            try:
                w_target = {"ADG_g": w_adg, "FCR": w_fcr, "RFI": w_rfi}
                w_list = [w_target[t] for t in ml_targets]
                with st.spinner("AI 智能选种分析运行中…"):
                    st.session_state.ml = ml_run_all(p["prod"], ml_targets, w_list,
                                                       w_imp, w_stab)
                if "error" in st.session_state.ml:
                    st.error(st.session_state.ml["error"])
                else:
                    st.success("AI 智能选种分析完成")
            except Exception as e:
                st.error(f"运行失败：{e}")

        ml = st.session_state.ml
        if ml and "error" not in ml:
            st.subheader("采食模式聚类")
            if "result" in ml["cluster"]:
                df_show(ml["cluster"]["result"].round(2).head(200), 300)
                df_show(ml["cluster"]["compare"].round(2), 200)

            st.subheader("特征重要性")
            if isinstance(ml["fusion"], pd.DataFrame):
                df_show(ml["fusion"].round(4), 400)

            st.subheader("模型评价")
            evals = [t["evaluation"] for t in ml["targets"].values()
                     if t.get("evaluation") is not None]
            if evals:
                df_show(pd.concat(evals, ignore_index=True), 200)

            st.subheader("综合评分 Top10")
            df_show(ml["top10"].round(3), 300)

            st.subheader("稳定性验证")
            if isinstance(ml["stability"], pd.DataFrame):
                df_show(ml["stability"].round(3), 300)

            st.subheader("育种报告")
            rep = ml["report"]
            df_show(rep, 500)

            fig = px.bar(rep.sort_values("FinalScore"), x="FinalScore", y="Feature",
                         orientation="h", color="FinalScore",
                         color_continuous_scale=[PAL_RPBG[3], PAL_RPBG[0]])
            render_plotly(fig, "ml_report_score", "ml_report.png")

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                rep.to_excel(w, sheet_name="育种报告", index=False)
                ml["top10"].to_excel(w, sheet_name="Top10", index=False)
                if isinstance(ml["stability"], pd.DataFrame):
                    ml["stability"].to_excel(w, sheet_name="稳定性", index=False)
                if "result" in ml["cluster"]:
                    ml["cluster"]["result"].to_excel(w, sheet_name="聚类结果", index=False)
            buf.seek(0)
            st.download_button("下载育种报告 Excel", data=buf,
                               file_name=f"duck_ai_breeding_report_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
