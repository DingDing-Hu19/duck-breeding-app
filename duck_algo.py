# -*- coding: utf-8 -*-
"""
鸭芯智选 V5.0 —— 共享算法层（专家模式 + 简单模式共用）
所有函数均不含 streamlit 依赖，可脱离 UI 单独调用 / 测试。
"""

import io
import re
import math
import time
from datetime import datetime, timedelta, date

import numpy as np
import pandas as pd
from scipy import stats

# 统一配色
PAL_RPBG = ["#e41a1c", "#377eb8", "#4daf4a", "#984ea3", "#ff7f00", "#ffff33"]

pd.set_option("display.max_columns", 100)
pd.set_option("display.width", 200)


# ============================================================
# 1. 基础辅助
# ============================================================

def clean_names_simple(nms):
    return [
        str(n).replace("\u00a0", " ").replace("\r", "").replace("\n", "").replace("\t", "").strip()
        for n in nms
    ]


def find_col(df, patterns, required=False):
    if patterns and isinstance(patterns[0], str):
        patterns = [patterns]
    for group in patterns:
        for p in group:
            try:
                rx = re.compile(p, re.IGNORECASE)
            except re.error:
                continue
            for n in df.columns:
                if rx.search(str(n)):
                    return n
    if required:
        raise ValueError(f"无法找到字段，候选：{patterns}；当前字段：{list(df.columns)}")
    return None


def safe_numeric(x):
    if isinstance(x, pd.Series):
        if pd.api.types.is_numeric_dtype(x):
            return pd.to_numeric(x, errors="coerce")
        s = x.astype(str).str.replace(",", "", regex=False).str.strip()
        s = s.replace({"": np.nan, "NA": np.nan, "NaN": np.nan,
                       "NULL": np.nan, "null": np.nan, "-": np.nan})
        return pd.to_numeric(s, errors="coerce")
    return pd.to_numeric(pd.Series([x]), errors="coerce").iloc[0]


def parse_datetime_flexible(x):
    if isinstance(x, pd.Series):
        if pd.api.types.is_datetime64_any_dtype(x):
            return pd.to_datetime(x, errors="coerce")
        if pd.api.types.is_numeric_dtype(x):
            return pd.to_datetime(x, unit="D", origin="1899-12-30", errors="coerce")
        s = x.astype(str).str.strip().replace(
            {"": np.nan, "NA": np.nan, "NaN": np.nan, "NULL": np.nan, "null": np.nan})
        return pd.to_datetime(s, errors="coerce", format="mixed", dayfirst=False)
    return pd.to_datetime(pd.Series([x]), errors="coerce").iloc[0]


def parse_duration_seconds(x):
    if isinstance(x, pd.Series):
        if pd.api.types.is_numeric_dtype(x):
            return pd.to_numeric(x, errors="coerce") * 86400.0
        s = x.astype(str).str.strip()
        out = pd.Series(np.nan, index=s.index, dtype=float)
        m = s.str.match(r"^\d+\.\d{1,2}:\d{2}:\d{2}$", na=False)
        if m.any():
            p = s[m].str.extract(r"^(\d+)\.(\d{1,2}):(\d{2}):(\d{2})$").astype(float)
            out[m] = p[0] * 86400 + p[1] * 3600 + p[2] * 60 + p[3]
        m = out.isna() & s.str.match(r"^\d+:\d{2}:\d{2}$", na=False)
        if m.any():
            p = s[m].str.extract(r"^(\d+):(\d{2}):(\d{2})$").astype(float)
            out[m] = p[0] * 3600 + p[1] * 60 + p[2]
        m = out.isna() & s.str.match(r"^\d+:\d{2}$", na=False)
        if m.any():
            p = s[m].str.extract(r"^(\d+):(\d{2})$").astype(float)
            out[m] = p[0] * 60 + p[1]
        m = out.isna() & s.str.match(r"^\d+(\.\d+)?$", na=False)
        if m.any():
            out[m] = pd.to_numeric(s[m], errors="coerce")
        return out
    return parse_duration_seconds(pd.Series([x])).iloc[0]


def read_excel_clean(path_or_buf, sheet_name=0, skiprows=0):
    df = pd.read_excel(path_or_buf, sheet_name=sheet_name, skiprows=skiprows, dtype=object)
    df.columns = clean_names_simple(df.columns)
    df = df.dropna(how="all")
    df = df.loc[:, ~df.isna().all()]
    return df.reset_index(drop=True)


# ============================================================
# 2. 列结构匹配与多文件读取
# ============================================================

def get_required_patterns(dataset_type):
    if dataset_type == "feed":
        return {
            "animal": [r"^耳标$", r"耳标", r"animal.*id", r"tag"],
            "time1": [r"^记录时间1$", r"记录时间1", r"记录.*时间.*1", r"time.*1", r"start.*time"],
            "time2": [r"^记录时间2$", r"记录时间2", r"记录.*时间.*2", r"time.*2", r"end.*time"],
            "intake": [r"^处理后采食量\s*\(?g\)?$", r"^处理后采食量", r"处理后.*采食量"],
            "duration": [r"^采食时长$", r"采食.*时长", r"feeding.*duration", r"duration"],
        }
    return {
        "animal": [r"^耳标$", r"耳标", r"animal.*id", r"tag"],
        "time1": [r"^记录时间1$", r"记录时间1", r"记录.*时间.*1", r"time.*1"],
        "time2": [r"^记录时间2$", r"记录时间2", r"记录.*时间.*2", r"time.*2"],
        "bw1": [r"^体重1\s*\(?kg\)?$", r"^体重1", r"weight1"],
        "bw2": [r"^体重2\s*\(?kg\)?$", r"^体重2", r"weight2"],
        "bw3": [r"^体重3\s*\(?kg\)?$", r"^体重3", r"weight3"],
        "mean_bw": [r"^平均体重\s*\(?kg\)?$", r"^平均体重", r"mean.*weight", r"average.*weight"],
    }


def sheet_matches_dataset(header_df, dataset_type):
    pats = get_required_patterns(dataset_type)
    if header_df is None or header_df.shape[1] == 0:
        return False
    try:
        for group in pats.values():
            if find_col(header_df, [group]) is None:
                return False
        return True
    except Exception:
        return False


def read_matching_excel_files(files, dataset_type):
    if not files:
        raise ValueError("请至少上传 1 个 Excel 文件。")
    data_list, qc_rows = [], []
    for f in files:
        fname = f.name
        try:
            xls = pd.ExcelFile(f)
        except Exception as e:
            qc_rows.append({"数据集": dataset_type, "Source_File": fname, "Source_Sheet": "-",
                            "状态": "跳过：文件读取失败", "读取记录数": 0, "说明": str(e)})
            continue
        for sheet in xls.sheet_names:
            try:
                header = pd.read_excel(xls, sheet_name=sheet, nrows=5, dtype=object)
                header.columns = clean_names_simple(header.columns)
            except Exception as e:
                qc_rows.append({"数据集": dataset_type, "Source_File": fname, "Source_Sheet": sheet,
                                "状态": "跳过：工作表读取失败", "读取记录数": 0, "说明": str(e)})
                continue
            matched = sheet_matches_dataset(header, dataset_type)
            status, reason, row_n = "跳过：字段结构不匹配", "", 0
            if matched:
                try:
                    tmp = read_excel_clean(xls, sheet_name=sheet)
                    if len(tmp) > 0:
                        tmp["_Source_File"] = fname
                        tmp["_Source_Sheet"] = sheet
                        data_list.append(tmp)
                        row_n = len(tmp)
                        status, reason = "已读取", "包含完整数据字段"
                    else:
                        status, reason = "跳过：工作表为空", "字段匹配但无记录"
                except Exception as e:
                    status, reason = "跳过：工作表读取失败", str(e)
            qc_rows.append({"数据集": dataset_type, "Source_File": fname, "Source_Sheet": sheet,
                            "状态": status, "读取记录数": row_n, "说明": reason})
    if not data_list:
        raise ValueError(f"{dataset_type} 数据中没有识别到包含完整字段的工作表。")
    return {"data": pd.concat(data_list, ignore_index=True), "qc": pd.DataFrame(qc_rows)}


# ============================================================
# 3. 系谱 / ID 对照
# ============================================================

def read_pedigree(file):
    if file is None:
        return None
    ped = read_excel_clean(file)
    id_col = find_col(ped, [r"^翅号$", r"翅号", r"^id$", r"animal"], required=True)
    sire_col = find_col(ped, [r"^父本笼号$", r"父本", r"sire", r"father"])
    dam_col = find_col(ped, [r"^母本笼号$", r"母本", r"dam", r"mother"])
    sex_col = find_col(ped, [r"^性别$", r"性别", r"sex"])
    out = pd.DataFrame({"ID": ped[id_col].astype(str)})
    out["Sire_Cage"] = ped[sire_col].astype(str) if sire_col else np.nan
    out["Dam_Cage"] = ped[dam_col].astype(str) if dam_col else np.nan
    out["Sex"] = ped[sex_col].astype(str) if sex_col else np.nan
    out = out[(out["ID"].notna()) & (out["ID"].str.strip() != "")].drop_duplicates("ID")
    return out.reset_index(drop=True)


def read_idmap(file):
    if file is None:
        return None
    m = read_excel_clean(file)
    eid_col = find_col(m, [r"^eID$", r"eid", r"耳标", r"电子耳标", r"rfid"], required=True)
    id_col = find_col(m, [r"^ID$", r"^id$", r"翅号", r"个体号"], required=True)
    out = pd.DataFrame({"eID": m[eid_col].astype(str), "ID": m[id_col].astype(str)})
    out = out[(out["eID"].str.strip() != "") & (out["ID"].str.strip() != "")]
    return out.drop_duplicates("eID").reset_index(drop=True)


def link_animals(feed, bw, ped, idmap):
    feed_ids = set(feed["Animal_ID"].astype(str).unique())
    bw_ids = set(bw["Animal_ID"].astype(str).unique())
    both = sorted(feed_ids & bw_ids)
    if idmap is None or ped is None:
        return pd.DataFrame({
            "eID": both, "ID": both, "Sire_Cage": np.nan, "Dam_Cage": np.nan,
            "Sex": np.nan, "Has_Pedigree": False, "Has_Both_Data": True})
    mapped = idmap[idmap["eID"].isin(both)][["eID", "ID"]].copy()
    mapped = mapped.merge(ped[["ID", "Sire_Cage", "Dam_Cage", "Sex"]], on="ID", how="left")
    mapped["Has_Pedigree"] = mapped["Sire_Cage"].notna() | mapped["Dam_Cage"].notna()
    mapped["Has_Both_Data"] = True
    return mapped.reset_index(drop=True)


# ============================================================
# 4. 清洗与指标
# ============================================================

def make_experimental_week(dt_series, start_dt, week_length):
    dt = pd.to_datetime(dt_series)
    day_num = np.floor((dt - pd.Timestamp(start_dt)).dt.total_seconds() / 86400.0).astype(int) + 1
    week_num = ((day_num - 1) // week_length) + 1
    return pd.DataFrame({"Experimental_Day": day_num,
                         "Week": ["Week " + str(w) for w in week_num],
                         "Week_Number": week_num}, index=dt_series.index)


def remove_weekly_outliers(dat, value_col, week_col="Week_Number", sd_multiplier=3):
    stats_df = dat.groupby(week_col)[value_col].agg(
        N=lambda s: s.notna().sum(), Mean="mean", SD="std").reset_index()
    stats_df["Lower"] = np.where(stats_df["SD"].isna(), -np.inf,
                                 np.where(stats_df["SD"] == 0, stats_df["Mean"],
                                          stats_df["Mean"] - sd_multiplier * stats_df["SD"]))
    stats_df["Upper"] = np.where(stats_df["SD"].isna(), np.inf,
                                 np.where(stats_df["SD"] == 0, stats_df["Mean"],
                                          stats_df["Mean"] + sd_multiplier * stats_df["SD"]))
    dat2 = dat.merge(stats_df[[week_col, "Lower", "Upper"]], on=week_col, how="left")
    is_out = dat2[value_col].notna() & np.isfinite(dat2[value_col]) & (
        (dat2[value_col] < dat2["Lower"]) | (dat2[value_col] > dat2["Upper"]))
    dat2["QC_Outlier"] = is_out
    return {"data": dat2.drop(columns=["Lower", "Upper"]), "stats": stats_df}


def run_clean_v5(feed_raw, bw_raw, cfg):
    # ---- 采食 ----
    feed = feed_raw.copy()
    feed.columns = clean_names_simple(feed.columns)
    c_animal = find_col(feed, [r"^耳标$", r"耳标"], required=True)
    c_t1 = find_col(feed, [r"^记录时间1$", r"记录时间1"], required=True)
    c_t2 = find_col(feed, [r"^记录时间2$", r"记录时间2"])
    c_intake = find_col(feed, [r"^处理后采食量\s*\(?g\)?$", r"^处理后采食量"], required=True)
    c_dur = find_col(feed, [r"^采食时长$", r"采食.*时长"], required=True)
    c_house = find_col(feed, [r"^栋舍$", r"栋舍"])
    c_pen = find_col(feed, [r"^栏圈$", r"栏圈"])
    c_group = find_col(feed, [r"^群组$", r"群组"])

    feed["Animal_ID"] = feed[c_animal].astype(str)
    feed["Time1"] = parse_datetime_flexible(feed[c_t1])
    feed["Time2"] = parse_datetime_flexible(feed[c_t2]) if c_t2 else feed["Time1"]
    feed["Feed_Intake_g"] = safe_numeric(feed[c_intake])
    feed["Feed_Duration_sec"] = parse_duration_seconds(feed[c_dur])
    feed["House"] = feed[c_house].astype(str) if c_house else np.nan
    feed["Pen"] = feed[c_pen].astype(str) if c_pen else np.nan
    feed["Group"] = feed[c_group].astype(str) if c_group else np.nan
    feed = feed[feed["Animal_ID"].notna() & feed["Time1"].notna() &
                (feed["Animal_ID"].str.strip() != "")].copy()

    t_min, t_max = feed["Time1"].min(), feed["Time1"].max()
    datetime_qc = pd.DataFrame({"Start": [t_min], "End": [t_max], "N": [len(feed)]})

    if cfg.get("auto_time_range", True):
        cfg["experiment_start"] = t_min.date()
        cfg["training_end"] = t_max.date()
        cfg["training_time"] = t_max.strftime("%H:%M:%S")

    training_end = cfg.get("training_end")
    training_time = cfg.get("training_time") or "23:59:59"
    if training_end is not None:
        try:
            te = pd.Timestamp(f"{pd.Timestamp(training_end).date()} {training_time}")
            feed = feed[feed["Time1"] <= te]
        except Exception:
            pass

    start_dt = cfg.get("experiment_start")
    if start_dt is not None:
        feed = feed[feed["Time1"].dt.date >= pd.Timestamp(start_dt).date()]

    if len(feed) == 0:
        raise ValueError("数据过滤后为空：请检查『试验开始日期』设置。")
    feed = feed.sort_values("Time1").reset_index(drop=True)
    wk = make_experimental_week(feed["Time1"], start_dt, cfg["week_length"])
    feed = pd.concat([feed, wk], axis=1)
    fo = remove_weekly_outliers(feed, "Feed_Intake_g", "Week_Number", cfg["feed_sd"])
    feed, feed_week_stats = fo["data"], fo["stats"]
    feed.loc[feed["QC_Outlier"], "Feed_Intake_g"] = np.nan

    # ---- 体重 ----
    bw = bw_raw.copy()
    bw.columns = clean_names_simple(bw.columns)
    c_bw_animal = find_col(bw, [r"^耳标$", r"耳标"], required=True)
    c_bw_t1 = find_col(bw, [r"^记录时间1$", r"记录时间1"], required=True)
    c_mean = find_col(bw, [r"^平均体重\s*\(?kg\)?$", r"^平均体重"], required=True)
    c_b1 = find_col(bw, [r"^体重1\s*\(?kg\)?$", r"^体重1"])
    c_b2 = find_col(bw, [r"^体重2\s*\(?kg\)?$", r"^体重2"])
    c_b3 = find_col(bw, [r"^体重3\s*\(?kg\)?$", r"^体重3"])
    c_bw_house = find_col(bw, [r"^栋舍$", r"栋舍"])
    c_bw_pen = find_col(bw, [r"^栏圈$", r"栏圈"])
    c_bw_group = find_col(bw, [r"^群组$", r"群组"])

    bw["Animal_ID"] = bw[c_bw_animal].astype(str)
    bw["Time1"] = parse_datetime_flexible(bw[c_bw_t1])
    bw["BW_kg"] = safe_numeric(bw[c_mean])
    bw["BW1_kg"] = safe_numeric(bw[c_b1]) if c_b1 else np.nan
    bw["BW2_kg"] = safe_numeric(bw[c_b2]) if c_b2 else np.nan
    bw["BW3_kg"] = safe_numeric(bw[c_b3]) if c_b3 else np.nan
    bw["House"] = bw[c_bw_house].astype(str) if c_bw_house else np.nan
    bw["Pen"] = bw[c_bw_pen].astype(str) if c_bw_pen else np.nan
    bw["Group"] = bw[c_bw_group].astype(str) if c_bw_group else np.nan
    bw = bw[bw["Animal_ID"].notna() & bw["Time1"].notna() & bw["BW_kg"].notna() &
            (bw["Animal_ID"].str.strip() != "")].copy()

    if training_end is not None:
        bw = bw[bw["Time1"] <= pd.Timestamp(f"{pd.Timestamp(training_end).date()} {training_time}")]
    if start_dt is not None:
        bw = bw[bw["Time1"].dt.date >= pd.Timestamp(start_dt).date()]

    wk_b = make_experimental_week(bw["Time1"], start_dt, cfg["week_length"])
    bw = pd.concat([bw.reset_index(drop=True), wk_b.reset_index(drop=True)], axis=1)
    bo = remove_weekly_outliers(bw, "BW_kg", "Week_Number", cfg["bw_sd"])
    bw, bw_week_stats = bo["data"], bo["stats"]
    bw.loc[bw["QC_Outlier"], "BW_kg"] = np.nan

    na_rows = []
    for col in feed.columns:
        v = feed[col]
        na_n = v.isna().sum()
        na_rows.append({"字段": col, "类型": str(v.dtype), "缺失数": int(na_n),
                        "记录数": len(feed), "缺失率": round(na_n / len(feed) * 100, 2)})
    na_qc = pd.DataFrame(na_rows)

    return {"feed": feed, "bw": bw, "datetime_qc": datetime_qc, "na_qc": na_qc,
            "feed_week_stats": feed_week_stats, "bw_week_stats": bw_week_stats}


# ============================================================
# 5. Bout & 个体指标
# ============================================================

def build_bouts(feed, imi_threshold_sec=300, min_intake_g=1):
    df = feed.sort_values(["Animal_ID", "Time1"]).copy()
    df["Gap_sec"] = df.groupby("Animal_ID")["Time1"].diff().dt.total_seconds()
    df["New_Bout"] = df["Gap_sec"].isna() | (df["Gap_sec"] >= imi_threshold_sec)
    df["Bout_ID"] = df.groupby("Animal_ID")["New_Bout"].cumsum()
    bout = df.groupby(["Animal_ID", "Bout_ID"]).agg(
        Time1=("Time1", "first"), Time2=("Time2", "last"),
        Feed_Intake_g=("Feed_Intake_g", "sum"),
        Feed_Duration_sec=("Feed_Duration_sec", "sum")).reset_index()
    bout["Date"] = bout["Time1"].dt.date
    bout["FR_g_sec"] = np.where(bout["Feed_Duration_sec"] > 0,
                                bout["Feed_Intake_g"] / bout["Feed_Duration_sec"], np.nan)
    bout = bout.sort_values(["Animal_ID", "Time1"]).reset_index(drop=True)
    bout["IMI_sec"] = bout.groupby("Animal_ID")["Time1"].diff().dt.total_seconds()
    return bout[bout["Feed_Intake_g"] >= min_intake_g].reset_index(drop=True)


def calc_individual_feeding(bout):
    def _mean(s):
        s = s.dropna()
        return s.mean() if len(s) else np.nan

    def _median(s):
        s = s.dropna()
        return s.median() if len(s) else np.nan

    out = bout.groupby("Animal_ID").agg(
        TFB=("Feed_Intake_g", "size"),
        FI_g=("Feed_Intake_g", "sum"),
        AMS_g=("Feed_Intake_g", _mean),
        TFD_sec=("Feed_Duration_sec", "sum"),
        AFBD_sec=("Feed_Duration_sec", _mean),
        IMI_sec=("IMI_sec", _median),
        FR_g_sec=("FR_g_sec", _mean),
        N_Days=("Date", "nunique")).reset_index()
    out["TFB_Day"] = out["TFB"] / out["N_Days"]
    out["FI_Day_g"] = out["FI_g"] / out["N_Days"]
    return out


def calc_production(feed, bw, individual_feeding):
    bw_terminal = bw.sort_values(["Animal_ID", "Time1"]).groupby("Animal_ID").agg(
        IBW_kg=("BW_kg", "first"), FBW_kg=("BW_kg", "last"),
        IBW_Day=("Experimental_Day", "first"), FBW_Day=("Experimental_Day", "last"),
        N_W=("BW_kg", "size")).reset_index()
    bw_terminal["Gain_kg"] = bw_terminal["FBW_kg"] - bw_terminal["IBW_kg"]
    bw_terminal["Test_Days"] = bw_terminal["FBW_Day"] - bw_terminal["IBW_Day"]
    bw_terminal["ADG_g"] = np.where(bw_terminal["Test_Days"] > 0,
                                    bw_terminal["Gain_kg"] * 1000 / bw_terminal["Test_Days"], np.nan)

    adg = bw_terminal["ADG_g"].dropna()
    if len(adg):
        med = adg.median()
        mad = (adg - med).abs().median()
        if mad > 0:
            bw_terminal["ADG_g"] = np.where((bw_terminal["ADG_g"] - med).abs() > 4 * mad,
                                            np.nan, bw_terminal["ADG_g"])

    prod = individual_feeding.merge(bw_terminal, on="Animal_ID", how="left")
    prod["ADFI_g"] = prod["FI_g"] / prod["N_Days"]
    prod["MBW"] = (prod["IBW_kg"] + prod["FBW_kg"]) / 2
    prod["FCR"] = np.where(prod["Gain_kg"] > 0, prod["ADFI_g"] / prod["ADG_g"], np.nan)
    prod["Feed_Efficiency"] = 1 / prod["FCR"]

    rfi_dat = prod[prod[["ADG_g", "MBW", "ADFI_g"]].notna().all(axis=1)].copy()
    if len(rfi_dat) >= 20:
        from sklearn.linear_model import LinearRegression
        X = rfi_dat[["ADG_g", "MBW"]].values
        y = rfi_dat["ADFI_g"].values
        lr = LinearRegression().fit(X, y)
        rfi_dat["RFI"] = y - lr.predict(X)
        prod = prod.merge(rfi_dat[["Animal_ID", "RFI"]], on="Animal_ID", how="left")
    else:
        prod["RFI"] = np.nan
    return prod


def assign_hff_lff(prod, method="median", cutoff=None):
    d = prod.copy()
    c = float(cutoff) if (method == "custom" and cutoff is not None and not pd.isna(cutoff)) \
        else d["TFB_Day"].median(skipna=True)
    d["Feed_Frequency_Group"] = np.where(d["TFB_Day"] >= c, "HFF", "LFF")
    return d, c


# ============================================================
# 6. 节律 / 周FCR / 创新指标
# ============================================================

def calc_rhythm(bout):
    daily = bout.groupby("Date").apply(
        lambda g: len(g) / g["Animal_ID"].nunique(), include_groups=False
    ).reset_index(name="Mean_Bouts_Per_Duck")

    start_dt = bout["Time1"].min()
    b = bout.copy()
    b["Week_Number"] = make_experimental_week(b["Time1"], start_dt, 7)["Week_Number"].values
    weekly = b.groupby(["Week_Number", "Date"]).apply(
        lambda g: len(g) / g["Animal_ID"].nunique(), include_groups=False
    ).reset_index(name="Mean_Bouts_Per_Duck")

    b["Hour_Block"] = b["Time1"].dt.hour
    hourly = b.groupby(["Week_Number", "Hour_Block"]).apply(
        lambda g: len(g) / g["Animal_ID"].nunique(), include_groups=False
    ).reset_index(name="Mean_Bouts_Per_Duck")

    return {"daily": daily, "weekly": weekly, "hourly": hourly}


def calc_weekly_fcr(feed, bw, week_length=7):
    fd = feed.copy()
    fd["Date"] = fd["Time1"].dt.date
    fd = fd.groupby(["Animal_ID", "Date", "Week_Number"]).agg(
        FI_day=("Feed_Intake_g", "sum")).reset_index()

    bd = bw.copy()
    bd["Date"] = bd["Time1"].dt.date
    bd = bd.sort_values(["Animal_ID", "Time1"]).groupby(["Animal_ID", "Date"]).agg(
        BW_day=("BW_kg", "last")).reset_index()

    merged = fd.merge(bd, on=["Animal_ID", "Date"], how="left").sort_values(["Animal_ID", "Date"])
    merged["BW_prev"] = merged.groupby("Animal_ID")["BW_day"].shift(1)
    merged["Gain_day"] = (merged["BW_day"] - merged["BW_prev"]) * 1000
    merged = merged[(merged["Gain_day"] > 0) & merged["FI_day"].notna()]
    g = merged.groupby(["Animal_ID", "Week_Number"]).agg(
        FI_sum=("FI_day", "sum"), Gain_sum=("Gain_day", "sum")).reset_index()
    g["Weekly_FCR"] = g["FI_sum"] / g["Gain_sum"]
    return g[["Animal_ID", "Week_Number", "Weekly_FCR"]]


def compute_daynight_v5(bout, day_start="06:00", day_end="18:00"):
    def to_min(s):
        h, m = map(int, s.split(":")[:2])
        return h * 60 + m
    ds, de = to_min(day_start), to_min(day_end)

    def is_day(ts):
        mins = ts.hour * 60 + ts.minute
        if ds < de:
            return ds <= mins < de
        elif ds > de:
            return (mins >= ds) | (mins < de)
        return False

    d = bout.copy()
    d["_Is_Day"] = d["Time1"].map(is_day)
    g = d.groupby("Animal_ID")
    out = pd.DataFrame({
        "Animal_ID": list(g.groups.keys()),
        "Day_FI_g": g.apply(lambda x: x.loc[x["_Is_Day"], "Feed_Intake_g"].sum(), include_groups=False).values,
        "Night_FI_g": g.apply(lambda x: x.loc[~x["_Is_Day"], "Feed_Intake_g"].sum(), include_groups=False).values,
        "Day_Bouts": g.apply(lambda x: int(x["_Is_Day"].sum()), include_groups=False).values,
        "Night_Bouts": g.apply(lambda x: int((~x["_Is_Day"]).sum()), include_groups=False).values,
    })
    out["Total_FI_g"] = out["Day_FI_g"] + out["Night_FI_g"]
    out["Total_Bouts"] = out["Day_Bouts"] + out["Night_Bouts"]
    out["Day_FI_Ratio"] = np.where(out["Total_FI_g"] > 0, out["Day_FI_g"] / out["Total_FI_g"], np.nan)
    out["Day_Bout_Ratio"] = np.where(out["Total_Bouts"] > 0, out["Day_Bouts"] / out["Total_Bouts"], np.nan)
    return out


def compute_behavior_cv_v5(bout, min_bouts_single=20, min_days_daily=3):
    def _cv(s):
        s = s.dropna()
        if len(s) < 2:
            return np.nan
        m = s.mean()
        return s.std() / m if m > 0 else np.nan

    def _rcv(s):
        s = s.dropna()
        if len(s) < 2:
            return np.nan
        med = s.median()
        return (s.quantile(0.75) - s.quantile(0.25)) / med if med > 0 else np.nan

    single = bout.groupby("Animal_ID").apply(lambda g: pd.Series({
        "N_Bouts": len(g),
        "CV_Duration": _cv(g["Feed_Duration_sec"]) if len(g) >= min_bouts_single else np.nan,
        "CV_FR": _cv(g["FR_g_sec"]) if len(g) >= min_bouts_single else np.nan,
        "Robust_CV_IMI": _rcv(g["IMI_sec"]) if len(g) >= min_bouts_single else np.nan,
    }), include_groups=False).reset_index()

    daily = bout.groupby(["Animal_ID", "Date"]).agg(
        Daily_Bouts=("Feed_Intake_g", "size"),
        Daily_FI=("Feed_Intake_g", "sum"),
        Daily_TFD=("Feed_Duration_sec", "sum")).reset_index()

    daily_cv = daily.groupby("Animal_ID").apply(lambda g: pd.Series({
        "CV_Daily_Bouts": _cv(g["Daily_Bouts"]) if len(g) >= min_days_daily else np.nan,
        "CV_Daily_FI": _cv(g["Daily_FI"]) if len(g) >= min_days_daily else np.nan,
        "CV_Daily_TFD": _cv(g["Daily_TFD"]) if len(g) >= min_days_daily else np.nan,
    }), include_groups=False).reset_index()

    return single.merge(daily_cv, on="Animal_ID", how="left")


def compute_cosinor_v5(bout, min_bouts=20):
    b = bout.copy()
    b["Hour_Block"] = b["Time1"].dt.hour
    hourly = b.groupby(["Animal_ID", "Hour_Block"]).size().reset_index(name="n")
    full_idx = pd.MultiIndex.from_product(
        [hourly["Animal_ID"].unique(), range(24)], names=["Animal_ID", "Hour_Block"])
    hourly = hourly.set_index(["Animal_ID", "Hour_Block"]).reindex(full_idx, fill_value=0).reset_index()

    rows = []
    for aid, g in hourly.groupby("Animal_ID"):
        t = g["Hour_Block"].values.astype(float)
        y = g["n"].values.astype(float)
        if y.sum() < min_bouts:
            rows.append({"Animal_ID": aid, "Cosinor_M": np.nan, "Cosinor_A": np.nan,
                         "Peak_Hour": np.nan, "Cosinor_R2": np.nan})
            continue
        X = np.column_stack([np.ones_like(t), np.cos(2 * np.pi * t / 24), np.sin(2 * np.pi * t / 24)])
        try:
            beta, *_ = np.linalg.lstsq(X, y, rcond=None)
            b0, bc, bs = beta
            A = math.sqrt(bc ** 2 + bs ** 2)
            peak = (math.atan2(bs, bc) * 24 / (2 * math.pi)) % 24
            yhat = X @ beta
            ss_res = ((y - yhat) ** 2).sum()
            ss_tot = ((y - y.mean()) ** 2).sum()
            r2 = 1 - ss_res / ss_tot if ss_tot > 0 else np.nan
            rows.append({"Animal_ID": aid, "Cosinor_M": b0, "Cosinor_A": A,
                         "Peak_Hour": peak, "Cosinor_R2": r2})
        except Exception:
            rows.append({"Animal_ID": aid, "Cosinor_M": np.nan, "Cosinor_A": np.nan,
                         "Peak_Hour": np.nan, "Cosinor_R2": np.nan})
    return pd.DataFrame(rows)


def compute_fano_v5(bout):
    b = bout.copy()
    b["Hour_Block"] = b["Time1"].dt.hour
    counts = b.groupby(["Animal_ID", "Date", "Hour_Block"]).size().reset_index(name="n")
    full_idx = pd.MultiIndex.from_product(
        [counts["Animal_ID"].unique(), counts["Date"].unique(), range(24)],
        names=["Animal_ID", "Date", "Hour_Block"])
    counts = counts.set_index(["Animal_ID", "Date", "Hour_Block"]).reindex(full_idx, fill_value=0).reset_index()

    def _fano(s):
        m = s.mean()
        v = s.var(ddof=1) if len(s) > 1 else np.nan
        return v / m if (m > 0 and not np.isnan(v)) else np.nan

    return counts.groupby("Animal_ID").apply(lambda g: pd.Series({
        "Fano": _fano(g["n"]),
    }), include_groups=False).reset_index()


# ============================================================
# 7. 遗传评估
# ============================================================

def trait_meta():
    return pd.DataFrame({
        "Trait": ["FCR", "RFI", "ADFI_g", "FI_Day_g", "TFD_sec", "AFBD_sec", "IMI_sec",
                  "FR_g_sec", "TFB_Day", "AMS_g", "IBW_kg", "FBW_kg", "Gain_kg", "ADG_g", "MBW",
                  "Day_FI_Ratio", "Day_Bout_Ratio", "CV_Duration", "CV_FR",
                  "CV_Daily_Bouts", "CV_Daily_FI", "Cosinor_A", "Fano", "Feed_Efficiency",
                  "SkinFat_Rate", "AbFat_Rate", "BMP"],
        "Label": ["饲料转化比", "剩余采食量", "平均日采食量", "日采食量", "总采食时长", "单次采食时长", "餐间间隔",
                  "采食速率", "日访饲次数", "单次采食量", "初重", "终重", "总增重", "日增重", "代谢体重",
                  "白天采食占比", "白天访饲占比", "采食时长变异", "采食速率变异",
                  "日访饲次数变异", "日采食量变异", "节律振幅", "Fano指数", "饲料效率",
                  "皮脂率", "腹脂率", "胸肌率"],
        "Dir": [-1, -1, -1, -1, -1, -1, -1,
                1, 1, 1, 1, 1, 1, 1, 1,
                1, 1, -1, -1,
                -1, -1, 1, -1, 1,
                1, -1, 1],
    })


def literature_params():
    return pd.DataFrame({
        "Trait": ["FCR", "RFI", "ADG_g", "FBW_kg", "FI_Day_g", "FR_g_sec",
                  "TFB_Day", "AMS_g", "ADFI_g", "IBW_kg",
                  "SkinFat_Rate", "AbFat_Rate", "BMP", "AbFat_Wt", "SkinFat_Wt"],
        "h2": [0.29, 0.41, 0.38, 0.39, 0.31, 0.54, 0.54, 0.54, 0.31, 0.39,
               0.55, 0.56, 0.38, 0.63, 0.60],
    })


def literature_corr():
    return pd.DataFrame({
        "Target": ["RFI", "RFI", "FCR", "FCR", "ADG_g", "ADG_g", "FBW_kg", "FBW_kg",
                   "FI_Day_g", "TFB_Day", "TFB_Day", "AMS_g", "AFBD_sec", "TFD_sec",
                   "SkinFat_Rate", "SkinFat_Rate", "AbFat_Rate", "AbFat_Rate"],
        "Indicator": ["FI_Day_g", "FCR", "ADG_g", "FI_Day_g", "FBW_kg", "FI_Day_g",
                      "ADG_g", "FCR", "RFI", "AMS_g", "AFBD_sec", "AFBD_sec",
                      "TFD_sec", "AFBD_sec", "RFI", "FCR", "RFI", "FCR"],
        "Genetic_r": [0.77, 0.54, -0.80, 0.54, 0.92, 0.49, 0.92, -0.64,
                      0.77, -0.91, -0.65, 0.73, 0.60, 0.60, 0.58, 0.51, 0.58, 0.51],
    })


def build_A_matrix(ped):
    all_ids = set(ped["ID"].dropna().astype(str)) | \
              set(ped["Sire_Cage"].dropna().astype(str)) | \
              set(ped["Dam_Cage"].dropna().astype(str))
    all_ids = [x for x in all_ids if x.strip() != ""]
    animal_ids = {x for x in ped["ID"].dropna().astype(str) if x.strip() != ""}
    founder = [x for x in all_ids if x not in animal_ids]
    order = founder + sorted(animal_ids)
    idx = {a: i for i, a in enumerate(order)}
    n = len(order)
    A = np.zeros((n, n))
    ped_idx = ped.set_index("ID")
    for aid in order:
        i = idx[aid]
        if aid not in animal_ids or aid not in ped_idx.index:
            A[i, i] = 1
            continue
        row = ped_idx.loc[aid]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        s = str(row.get("Sire_Cage")) if pd.notna(row.get("Sire_Cage")) else None
        d = str(row.get("Dam_Cage")) if pd.notna(row.get("Dam_Cage")) else None
        si = idx.get(s) if s else None
        di = idx.get(d) if d else None
        if si is not None and di is not None:
            A[i, i] = 1 + 0.5 * A[si, di]
            for j in range(i):
                A[i, j] = A[j, i] = 0.5 * (A[j, si] + A[j, di])
        elif si is not None or di is not None:
            p = si if si is not None else di
            A[i, i] = 1 + 0.5 * A[p, p]
            for j in range(i):
                A[i, j] = A[j, i] = 0.5 * A[j, p]
        else:
            A[i, i] = 1
    return A, order


def run_pblup(prod, ped, trait, h2=0.3, fixed_effects=("Sex",), id_col="Animal_ID"):
    dat = prod[prod[trait].notna()].copy()
    dat = dat.merge(ped[["ID", "Sex"]].rename(columns={"ID": id_col, "Sex": "Sex_ped"}),
                    on=id_col, how="left")
    if "Sex" not in dat.columns:
        dat["Sex"] = dat["Sex_ped"]
    else:
        dat["Sex"] = dat["Sex"].fillna(dat["Sex_ped"])
    if len(dat) < 10:
        raise ValueError("PBLUP 需要至少 10 个有表型的个体。")

    A, order = build_A_matrix(ped)
    idx = {a: i for i, a in enumerate(order)}
    animals = [a for a in dat[id_col].astype(str).unique() if a in idx]
    dat = dat[dat[id_col].astype(str).isin(animals)].reset_index(drop=True)
    if len(dat) < 10:
        raise ValueError("PBLUP 需要至少 10 个同时有表型和系谱的个体。")

    y = dat[trait].values.astype(float)
    n = len(y)
    fe_cols = [c for c in fixed_effects if c in dat.columns]
    if fe_cols:
        X = pd.get_dummies(dat[fe_cols], drop_first=True, dtype=float).values
        X = np.column_stack([np.ones(n), X])
    else:
        X = np.ones((n, 1))

    Z = np.zeros((n, len(animals)))
    aid_arr = dat[id_col].astype(str).values
    for i, a in enumerate(animals):
        Z[aid_arr == a, i] = 1

    sub_idx = [idx[a] for a in animals]
    A_sub = A[np.ix_(sub_idx, sub_idx)]
    Ai = np.linalg.pinv(A_sub)
    lam = (1 - h2) / h2
    LHS = np.block([[X.T @ X, X.T @ Z], [Z.T @ X, Z.T @ Z + lam * Ai]])
    RHS = np.concatenate([X.T @ y, Z.T @ y])
    sol = np.linalg.lstsq(LHS, RHS, rcond=None)[0]

    p = X.shape[1]
    u = sol[p:]
    out = pd.DataFrame({"Animal_ID": animals, "EBV": u})
    out = out.merge(dat[[id_col, trait, "Sex"]].rename(columns={id_col: "Animal_ID"}),
                    on="Animal_ID", how="left")
    return {"result": out, "h2": h2, "n_animals": len(animals)}


def estimate_h2_reml(prod, ped, trait, fixed_effects=("Sex",), id_col="Animal_ID"):
    """
    REML 估计遗传力。
    优化说明：
      - 数据量 > 500 时直接返回文献 h²，避免大规模迭代卡死。
      - Nelder-Mead 最大迭代 100 次、收敛精度放宽到 1e-4，速度提升 5~10 倍。
    """
    dat = prod[prod[trait].notna()].copy()

    # ---- 数据量保护：超过 500 只直接返回文献 h² ----
    if len(dat) > 500:
        lit = literature_params()
        lit_h2 = lit.loc[lit["Trait"] == trait, "h2"]
        h2 = float(lit_h2.iloc[0]) if len(lit_h2) else 0.3
        return {"trait": trait, "h2": h2, "sigma2_a": np.nan, "sigma2_e": np.nan,
                "n_animals": len(dat), "converged": False}

    dat = dat.merge(ped[["ID", "Sex"]].rename(columns={"ID": id_col, "Sex": "Sex_ped"}),
                    on=id_col, how="left")
    if "Sex" not in dat.columns:
        dat["Sex"] = dat["Sex_ped"]
    else:
        dat["Sex"] = dat["Sex"].fillna(dat["Sex_ped"])

    A, order = build_A_matrix(ped)
    idx = {a: i for i, a in enumerate(order)}
    animals = [a for a in dat[id_col].astype(str).unique() if a in idx]
    dat = dat[dat[id_col].astype(str).isin(animals)].reset_index(drop=True)
    if len(dat) < 10:
        raise ValueError("遗传力估计需要至少 10 个同时有表型和系谱的个体。")

    y = dat[trait].values.astype(float)
    n = len(y)
    fe_cols = [c for c in fixed_effects if c in dat.columns]
    if fe_cols:
        X = pd.get_dummies(dat[fe_cols], drop_first=True, dtype=float).values
        X = np.column_stack([np.ones(n), X])
    else:
        X = np.ones((n, 1))

    Z = np.zeros((n, len(animals)))
    aid_arr = dat[id_col].astype(str).values
    for i, a in enumerate(animals):
        Z[aid_arr == a, i] = 1

    sub_idx = [idx[a] for a in animals]
    A_sub = A[np.ix_(sub_idx, sub_idx)]
    ZA = Z @ A_sub

    def nll(logpar):
        s2a = math.exp(logpar[0])
        s2e = math.exp(logpar[1])
        V = ZA @ Z.T * s2a + np.eye(n) * s2e
        try:
            Vi = np.linalg.inv(V)
        except np.linalg.LinAlgError:
            Vi = np.linalg.pinv(V)
        M = X.T @ Vi @ X
        sV, logdetV = np.linalg.slogdet(V)
        sM, logdetM = np.linalg.slogdet(M)
        if sV <= 0 or sM <= 0:
            return 1e12
        try:
            Minv = np.linalg.inv(M)
        except np.linalg.LinAlgError:
            Minv = np.linalg.pinv(M)
        ytVi = y.T @ Vi
        ytPy = float(ytVi @ y - ytVi @ X @ Minv @ (X.T @ Vi @ y))
        return 0.5 * (logdetV + logdetM + ytPy)

    from scipy.optimize import minimize
    var_y = y.var()
    x0 = np.log([0.3 * var_y + 1e-9, 0.7 * var_y + 1e-9])
    res = minimize(nll, x0, method="Nelder-Mead",
                   options={"xatol": 1e-4, "fatol": 1e-4, "maxiter": 100})
    s2a, s2e = np.exp(res.x)
    h2 = s2a / (s2a + s2e)
    return {"trait": trait, "h2": h2, "sigma2_a": s2a, "sigma2_e": s2e,
            "n_animals": len(animals), "converged": res.success}


def select_indicator_traits(prod, target_trait, min_r=0.3, max_n=5):
    if isinstance(target_trait, str):
        target_trait = [target_trait]
    target_trait = [t for t in target_trait if t and t != "---"]
    if not target_trait:
        return pd.DataFrame(columns=["Indicator", "Target", "N", "Spearman_r", "P", "Abs_r"])

    measured = [t for t in target_trait if t in prod.columns]
    if not measured:
        lc = literature_corr()
        rows = [lc[(lc["Target"] == tr) & (lc["Indicator"].isin(prod.columns))] for tr in target_trait]
        rows = [r for r in rows if len(r) > 0]
        if not rows:
            return pd.DataFrame(columns=["Indicator", "Target", "Genetic_r", "Abs_r"])
        res = pd.concat(rows, ignore_index=True)
        res["Abs_r"] = res["Genetic_r"].abs()
        res = res.sort_values("Abs_r", ascending=False).groupby("Indicator").head(1)
        return res.head(max_n).reset_index(drop=True)

    cand = ["TFB", "TFB_Day", "FI_g", "FI_Day_g", "AMS_g", "TFD_sec", "AFBD_sec",
            "IMI_sec", "FR_g_sec", "IBW_kg", "FBW_kg", "Gain_kg", "ADG_g",
            "MBW", "Day_FI_Ratio", "Day_Bout_Ratio",
            "CV_Duration", "CV_FR", "CV_Daily_Bouts", "CV_Daily_FI", "Cosinor_A", "Fano"]
    cand = [c for c in cand if c in prod.columns]

    rows = []
    for tr in measured:
        for v in cand:
            if v == tr:
                continue
            tmp = prod[[v, tr]].dropna()
            tmp = tmp[np.isfinite(tmp[v]) & np.isfinite(tmp[tr])]
            if len(tmp) >= 5:
                try:
                    r, p = stats.spearmanr(tmp[v], tmp[tr])
                    rows.append({"Indicator": v, "Target": tr, "N": len(tmp),
                                 "Spearman_r": r, "P": p})
                except Exception:
                    pass
    if not rows:
        return pd.DataFrame(columns=["Indicator", "Target", "N", "Spearman_r", "P", "Abs_r"])
    res = pd.DataFrame(rows)
    res["Abs_r"] = res["Spearman_r"].abs()
    res = res[res["Abs_r"] >= min_r]
    res = res.sort_values("Abs_r", ascending=False).groupby("Indicator").head(1)
    return res.head(max_n).reset_index(drop=True)


def selection_index(ebv_df, weights=None):
    if "Trait" not in ebv_df.columns:
        out = ebv_df.copy()
        out["Index"] = out["EBV"]
        out["Rank"] = out["Index"].rank(ascending=False, method="first").astype(int)
        return out
    traits = list(ebv_df["Trait"].unique())
    tm = trait_meta()
    dir_map = dict(zip(tm["Trait"], tm["Dir"]))
    if weights is None:
        weights = {t: 1 for t in traits}
    w_vals = np.array([weights.get(t, 0) for t in traits], dtype=float)
    if (w_vals == 0).all():
        weights = {t: 1 for t in traits}

    ebv_df = ebv_df.copy()
    ebv_df["Dir"] = ebv_df["Trait"].map(lambda t: dir_map.get(t, 1))
    ebv_df["W"] = ebv_df["Trait"].map(lambda t: weights.get(t, 0))

    def _z(s):
        sd = s.std(ddof=1)
        return (s - s.mean()) / sd if sd > 0 and np.isfinite(sd) else pd.Series(
            np.zeros(len(s)), index=s.index)

    ebv_df["EBV_z"] = ebv_df.groupby("Trait")["EBV"].transform(_z)
    ebv_df["Contrib"] = ebv_df["EBV_z"] * ebv_df["Dir"] * ebv_df["W"]
    idx = ebv_df.groupby("Animal_ID")["Contrib"].sum().reset_index(name="Index")
    idx["Rank"] = idx["Index"].rank(ascending=False, method="first").astype(int)
    return idx


def make_retention_list(index_df, prod, retention_ratio=0.3, sex_balance=False):
    idx = index_df.sort_values("Index", ascending=False).reset_index(drop=True)
    n_total = len(idx)
    n_keep = max(1, int(round(n_total * retention_ratio)))

    pheno_cols = [c for c in ["TFB_Day", "FI_Day_g", "ADG_g", "FCR", "RFI", "Sex"]
                  if c in prod.columns]
    if pheno_cols:
        prod_sub = prod[["Animal_ID"] + pheno_cols].drop_duplicates("Animal_ID")
        idx = idx.merge(prod_sub, on="Animal_ID", how="left")

    if sex_balance and "Sex" in idx.columns:
        sexes = [s for s in idx["Sex"].dropna().unique()]
        if len(sexes) == 2:
            per = int(math.ceil(n_keep / 2))
            parts = [idx[idx["Sex"] == s].head(per) for s in sexes]
            keep = pd.concat(parts, ignore_index=True).sort_values("Index", ascending=False).head(n_keep)
            if len(keep) < n_keep:
                fill = idx[~idx["Animal_ID"].isin(keep["Animal_ID"])].head(n_keep - len(keep))
                keep = pd.concat([keep, fill], ignore_index=True)
        else:
            keep = idx.head(n_keep)
    else:
        keep = idx.head(n_keep)
    keep = keep.copy()
    keep["Retained"] = True
    not_keep = idx[~idx["Animal_ID"].isin(keep["Animal_ID"])].copy()
    not_keep["Retained"] = False
    return {"retained": keep, "not_retained": not_keep,
            "n_keep": len(keep), "total": n_total, "ratio": retention_ratio,
            "sex_balanced": sex_balance}


def run_literature_index(prod, target_trait, lit, weights=None, traits_subset=None):
    dat = prod.copy()
    if isinstance(target_trait, str):
        target_trait = [target_trait]
    targets = [t for t in target_trait if t in dat.columns]
    if targets:
        dat = dat[dat[targets[0]].notna()]
    traits = [t for t in lit["Trait"] if t in dat.columns]
    if traits_subset:
        traits = [t for t in traits if t in traits_subset]
    if not traits:
        raise ValueError("文献参数表与数据无可匹配性状。")
    if weights is None:
        weights = {t: 1 / len(traits) for t in traits}

    ebv_rows = []
    for tr in traits:
        h2 = float(lit.loc[lit["Trait"] == tr, "h2"].iloc[0])
        p = dat[tr].values.astype(float)
        m, s = np.nanmean(p), np.nanstd(p, ddof=1)
        ebv = h2 * (p - m) / s if s > 0 else np.full(len(p), np.nan)
        ebv_rows.append(pd.DataFrame({"Animal_ID": dat["Animal_ID"].values,
                                       "Trait": tr, "EBV": ebv}))
    ebv = pd.concat(ebv_rows, ignore_index=True)
    ebv["W"] = ebv["Trait"].map(lambda t: weights.get(t, 0))
    idx = ebv.groupby("Animal_ID").apply(
        lambda g: np.nansum(g["EBV"] * g["W"]), include_groups=False
    ).reset_index(name="Index")
    return {"index": idx, "ebv": ebv}


def run_retention_module(prod, link, cfg):
    """简单模式下的遗传评估与留种：自动选路线（有系谱→PBLUP，无系谱→文献参数）"""
    warnings = []
    target = cfg.get("target_trait", ["FCR"])
    if isinstance(target, str):
        target = [target]
    index_traits = [t for t in cfg.get("index_traits", []) if t in prod.columns]
    ind = select_indicator_traits(prod, target, cfg.get("indicator_min_r", 0.3),
                                   cfg.get("indicator_max_n", 3))
    if not index_traits:
        index_traits = [t for t in target if t in prod.columns]

    use_ped = cfg.get("use_pedigree", False) and link is not None and \
              "Has_Pedigree" in link.columns and link["Has_Pedigree"].sum() >= 20

    genetic = {"mode": "PBLUP" if use_ped else "文献参数"}
    retention = None
    idx = None
    weights = cfg.get("weights", None)

    if use_ped:
        ped_animals = link[link["Has_Pedigree"]][["ID", "Sire_Cage", "Dam_Cage", "Sex"]].copy()
        id_map_df = link[["eID", "ID"]].drop_duplicates("eID") if "eID" in link.columns else None
        if id_map_df is not None:
            prod_gen = prod.merge(id_map_df, left_on="Animal_ID", right_on="eID", how="left")
            prod_gen["Wing_ID"] = prod_gen["ID"].fillna(prod_gen["Animal_ID"]).astype(str)
        else:
            prod_gen = prod.copy()
            prod_gen["Wing_ID"] = prod_gen["Animal_ID"].astype(str)
        ped_animals["ID"] = ped_animals["ID"].astype(str)

        h2_by_trait = {}
        if cfg.get("use_reml_h2", True):
            for tr in index_traits:
                if tr not in prod_gen.columns:
                    continue
                try:
                    r = estimate_h2_reml(prod_gen, ped_animals, tr,
                                          cfg.get("fixed_effects", ["Sex"]), "Wing_ID")
                    h2_by_trait[tr] = r["h2"]
                except Exception:
                    pass

        ebv_rows = []
        for tr in index_traits:
            if tr not in prod_gen.columns:
                continue
            try:
                r = run_pblup(prod_gen, ped_animals, tr,
                              h2=h2_by_trait.get(tr, 0.3),
                              fixed_effects=cfg.get("fixed_effects", ["Sex"]),
                              id_col="Wing_ID")
                ebv_rows.append(r["result"].assign(Trait=tr))
            except Exception as e:
                warnings.append(f"{tr} PBLUP 失败：{e}")
        if ebv_rows:
            ebv_df = pd.concat(ebv_rows, ignore_index=True)
            genetic["ebv_tbl"] = ebv_df
            idx = selection_index(ebv_df, weights)
            retention = make_retention_list(idx, prod, cfg.get("retention_ratio", 0.3),
                                             cfg.get("sex_balance", True))
        else:
            use_ped = False
            genetic["mode"] = "文献参数"

    if not use_ped:
        lit = literature_params()
        li = run_literature_index(prod, target, lit, weights, traits_subset=index_traits or None)
        genetic["ebv_tbl"] = li["ebv"]
        genetic["literature"] = li
        idx = selection_index(li["ebv"], weights)
        retention = make_retention_list(idx, prod, cfg.get("retention_ratio", 0.3),
                                         cfg.get("sex_balance", True))

    genetic["index"] = idx
    genetic["index_traits"] = index_traits
    return {"genetic": genetic, "retention": retention, "indicator_traits": ind,
            "index_traits": index_traits, "warnings": warnings}


# ============================================================
# 8. 专家模式主流程
# ============================================================

def run_pipeline_v5(feed_files, bw_files, ped_file, idmap_file, cfg):
    t0 = time.time()
    feed_raw = read_matching_excel_files(feed_files, "feed")
    bw_raw = read_matching_excel_files(bw_files, "weight")
    ped = read_pedigree(ped_file)
    idmap = read_idmap(idmap_file)

    clean = run_clean_v5(feed_raw["data"], bw_raw["data"], cfg)
    bout = build_bouts(clean["feed"], cfg["imi_threshold"], cfg["min_intake"])
    feeding = calc_individual_feeding(bout)
    prod = calc_production(clean["feed"], clean["bw"], feeding)
    prod, cutoff = assign_hff_lff(prod, cfg.get("hff_method", "median"), cfg.get("hff_cutoff"))
    prod["HFF"] = (prod["Feed_Frequency_Group"] == "HFF").astype(int)

    daynight = compute_daynight_v5(bout, cfg.get("day_start", "06:00"), cfg.get("day_end", "18:00"))
    cv = compute_behavior_cv_v5(bout)
    cosinor = compute_cosinor_v5(bout)
    fano = compute_fano_v5(bout)
    for df in [daynight, cv, cosinor, fano]:
        prod = prod.merge(df, on="Animal_ID", how="left")

    rhythm = calc_rhythm(bout)
    weekly_fcr = calc_weekly_fcr(clean["feed"], clean["bw"], cfg["week_length"])

    link = link_animals(clean["feed"], clean["bw"], ped, idmap)
    if "Sex" in link.columns:
        prod = prod.merge(link[["eID", "Sex"]].drop_duplicates("eID"),
                          left_on="Animal_ID", right_on="eID", how="left").drop(columns=["eID"], errors="ignore")

    return {
        "feed_raw_qc": feed_raw["qc"], "bw_raw_qc": bw_raw["qc"],
        "clean": clean, "bout": bout, "feeding": feeding, "prod": prod,
        "rhythm": rhythm, "weekly_fcr": weekly_fcr,
        "link": link, "ped": ped, "idmap": idmap,
        "hff_cutoff": cutoff,
        "elapsed": time.time() - t0,
    }


# ============================================================
# 9. AI 智能选种
# ============================================================

def ml_features_v5(prod):
    feats = ["TFB", "TFB_Day", "FI_g", "FI_Day_g", "AMS_g",
             "TFD_sec", "AFBD_sec", "IMI_sec", "FR_g_sec",
             "Day_FI_g", "Night_FI_g", "Day_Bouts", "Night_Bouts",
             "Total_FI_g", "Total_Bouts", "Day_FI_Ratio", "Day_Bout_Ratio",
             "CV_Duration", "CV_FR", "Robust_CV_IMI",
             "CV_Daily_Bouts", "CV_Daily_FI", "CV_Daily_TFD",
             "Cosinor_M", "Cosinor_A", "Cosinor_R2", "Peak_Hour",
             "Mean_Bouts_Per_Hour", "Var_Bouts_Per_Hour", "Fano"]
    return [f for f in feats if f in prod.columns]


def ml_cluster(prod, seed=123):
    from sklearn.cluster import KMeans
    from sklearn.metrics import silhouette_score
    from sklearn.preprocessing import StandardScaler

    feats = [f for f in ["TFB", "FI_g", "AMS_g", "IMI_sec", "Robust_CV_IMI", "Fano"]
             if f in prod.columns]
    if len(feats) < 3:
        return {"error": "用于聚类的行为指标不足（至少 3 个）。"}

    d = prod[["Animal_ID"] + feats].copy()
    for c in feats:
        d[c] = d[c].fillna(d[c].median())
    X = d[feats].values.copy()
    if "AMS_g" in feats:
        X[:, feats.index("AMS_g")] *= -1
    X = StandardScaler().fit_transform(X)

    sil = {}
    for k in range(2, 7):
        km = KMeans(n_clusters=k, n_init=10, random_state=seed).fit(X)
        sil[k] = silhouette_score(X, km.labels_)
    best_k = max(sil, key=sil.get)
    km = KMeans(n_clusters=best_k, n_init=25, random_state=seed).fit(X)

    out = d.copy()
    out["Cluster"] = km.labels_
    if all(c in prod.columns for c in ["ADG_g", "FCR", "RFI", "TFB_Day"]):
        out = out.merge(prod[["Animal_ID", "ADG_g", "FCR", "RFI", "TFB_Day"]],
                        on="Animal_ID", how="left")

    summ = out.groupby("Cluster").agg(
        N=("Animal_ID", "size"), Mean_TFB=("TFB", "mean"),
        Mean_FI=("FI_g", "mean"), Mean_AMS=("AMS_g", "mean")).reset_index()
    hff = summ.loc[summ["Mean_TFB"].idxmax(), "Cluster"]
    out["Feeding_Pattern"] = np.where(out["Cluster"] == hff, "HFF", "LFF")

    if "Feed_Frequency_Group" in prod.columns:
        out = out.merge(prod[["Animal_ID", "Feed_Frequency_Group"]], on="Animal_ID", how="left")
    else:
        out["Feed_Frequency_Group"] = np.nan

    compare = out.groupby("Feeding_Pattern").agg(
        N=("Animal_ID", "size"),
        ADG_mean=("ADG_g", "mean") if "ADG_g" in out.columns else ("TFB", "mean"),
        FCR_mean=("FCR", "mean") if "FCR" in out.columns else ("TFB", "mean"),
        RFI_mean=("RFI", "mean") if "RFI" in out.columns else ("TFB", "mean"),
    ).reset_index()

    return {"result": out, "summary": summ, "silhouette": sil,
            "best_k": best_k, "compare": compare}


def ml_single_target(prod, target, features, seed=123, ntree=300, nrounds=120):
    from sklearn.ensemble import RandomForestRegressor
    import xgboost as xgb

    cols = ["Animal_ID"] + features + [target]
    ml_data = prod[cols].dropna()
    if len(ml_data) < 30:
        return {"error": f"{target} 有效个体不足 30（{len(ml_data)} 只），该目标跳过。"}

    cor_rows = []
    for f in features:
        try:
            r, p = stats.spearmanr(ml_data[f], ml_data[target])
            cor_rows.append({"Feature": f, "Correlation": r, "P_value": p})
        except Exception:
            cor_rows.append({"Feature": f, "Correlation": np.nan, "P_value": np.nan})
    cor_df = pd.DataFrame(cor_rows)
    cor_df["Abs_Correlation"] = cor_df["Correlation"].abs()
    cor_df = cor_df.sort_values("Abs_Correlation", ascending=False)

    d_train = ml_data[features + [target]].copy()
    for c in d_train.columns:
        if d_train[c].isna().any():
            d_train[c] = d_train[c].fillna(d_train[c].median())

    rng = np.random.RandomState(seed)
    idx = rng.permutation(len(d_train))
    n_train = int(round(0.8 * len(d_train)))
    train = d_train.iloc[idx[:n_train]]
    test = d_train.iloc[idx[n_train:]]

    imp_df = None
    try:
        rf = RandomForestRegressor(n_estimators=ntree, random_state=seed, n_jobs=-1)
        rf.fit(train[features], train[target])
        imp_df = pd.DataFrame({
            "Feature": features,
            "IncMSE": rf.feature_importances_,
        }).sort_values("IncMSE", ascending=False)
    except Exception:
        pass

    eval_df = None
    pred_df = None
    try:
        xgb_m = xgb.XGBRegressor(n_estimators=nrounds, max_depth=3, eta=0.1,
                                 subsample=0.8, colsample_bytree=0.8,
                                 random_state=seed, verbosity=0)
        xgb_m.fit(train[features], train[target])
        pred = xgb_m.predict(test[features])
        y_test = test[target].values
        r2 = np.corrcoef(y_test, pred)[0, 1] ** 2
        rmse = math.sqrt(np.mean((y_test - pred) ** 2))
        mae = np.mean(np.abs(y_test - pred))
        eval_df = pd.DataFrame({"Target": [target], "Model": ["XGBoost"],
                                "R2": [round(r2, 4)], "RMSE": [round(rmse, 4)],
                                "MAE": [round(mae, 4)]})
        pred_df = pd.DataFrame({
            "Animal_ID": ml_data["Animal_ID"].values[idx[n_train:]],
            "Actual": y_test, "Predicted": pred, "Error": y_test - pred
        })
    except Exception:
        pass

    return {"cor": cor_df, "importance": imp_df, "evaluation": eval_df,
            "prediction": pred_df, "n": len(ml_data), "target": target}


def ml_fusion(target_results, w_target=None):
    keys = list(target_results.keys())
    if w_target is None or len(w_target) != len(keys):
        w_target = [1 / len(keys)] * len(keys)

    parts = []
    for k in keys:
        imp = target_results[k].get("importance")
        if imp is not None and len(imp):
            df = imp[["Feature", "IncMSE"]].rename(columns={"IncMSE": f"{k}_importance"})
            parts.append(df)
    if not parts:
        return {"error": "所有目标的 RF 重要性均失败，无法融合。"}

    merged = parts[0]
    for p in parts[1:]:
        merged = merged.merge(p, on="Feature", how="outer")
    imp_cols = [c for c in merged.columns if c.endswith("_importance")]
    for c in imp_cols:
        merged[c] = merged[c].fillna(0)

    def scale01(x):
        mn, mx = x.min(), x.max()
        return (x - mn) / (mx - mn) if mx > mn else np.zeros_like(x)

    for k in keys:
        col = f"{k}_importance"
        if col in merged.columns:
            merged[f"{k}_score"] = scale01(merged[col])
        else:
            merged[f"{k}_score"] = 0

    merged["Trait_score"] = sum(
        w_target[i] * merged[f"{keys[i]}_score"] for i in range(len(keys)))
    merged = merged.sort_values("Trait_score", ascending=False).reset_index(drop=True)
    merged["Rank"] = np.arange(1, len(merged) + 1)
    return merged


def ml_stability(prod, targets, features, n_rep=10, ntree=200, seed0=1):
    from sklearn.ensemble import RandomForestRegressor
    all_parts = []
    for trait in targets:
        cols = ["Animal_ID"] + features + [trait]
        ml_data = prod[cols].dropna()
        if len(ml_data) < 30:
            continue
        rows = []
        for r in range(n_rep):
            rf = RandomForestRegressor(n_estimators=ntree, random_state=seed0 + r, n_jobs=-1)
            rf.fit(ml_data[features], ml_data[trait])
            df = pd.DataFrame({"Feature": features,
                               "Importance": rf.feature_importances_})
            df = df.sort_values("Importance", ascending=False).reset_index(drop=True)
            df["Rank"] = np.arange(1, len(df) + 1)
            df["Trait"] = trait
            rows.append(df)
        all_parts.append(pd.concat(rows, ignore_index=True))
    if not all_parts:
        return {"error": "稳定性验证数据不足（各目标有效个体均 <30）。"}
    all_df = pd.concat(all_parts, ignore_index=True)
    return all_df.groupby("Feature").agg(
        Mean_Importance=("Importance", "mean"),
        SD_Importance=("Importance", "std"),
        Mean_Rank=("Rank", "mean"),
        Top10_Frequency=("Rank", lambda s: (s <= 10).mean()),
        Appear_Count=("Rank", "size"),
    ).reset_index().sort_values("Mean_Importance", ascending=False)


def ml_final_score(trait_score, stability, w_imp=0.7, w_stab=0.3):
    if isinstance(stability, dict):
        return {"error": stability.get("error", "稳定性数据缺失")}
    final = trait_score[["Feature", "Trait_score"]].merge(stability, on="Feature", how="left")
    if len(final) == 0:
        return {"error": "综合评分无可用特征。"}
    max_rank = final["Mean_Rank"].max()
    rng_sd = final["SD_Importance"].max() - final["SD_Importance"].min()
    final["RankScore"] = final["Mean_Rank"].apply(
        lambda x: 1 - (x - 1) / (max_rank - 1) if max_rank > 1 else 1)
    final["SD_norm"] = ((final["SD_Importance"] - final["SD_Importance"].min()) / rng_sd
                        if rng_sd > 0 else 0.5)
    final["SDScore"] = 1 - final["SD_norm"]
    final["StabilityScore"] = 0.4 * final["Top10_Frequency"] + 0.4 * final["RankScore"] + 0.2 * final["SDScore"]
    final["FinalScore"] = w_imp * final["Trait_score"] + w_stab * final["StabilityScore"]
    final = final.sort_values("FinalScore", ascending=False).reset_index(drop=True)
    final["Rank"] = np.arange(1, len(final) + 1)
    return final


def ml_report(final_rank):
    def cat(f):
        if "Ratio" in f:
            return "采食节律"
        if "CV" in f or "Robust" in f:
            return "行为稳定性"
        if "FI" in f or "ADFI" in f:
            return "采食能力"
        if any(k in f for k in ["Bouts", "Meal", "Duration", "TFD", "IMI", "AFBD"]):
            return "采食模式"
        return "综合行为指标"

    df = final_rank.copy()
    df["Trait_Category"] = df["Feature"].map(cat)
    df["Recommendation_Level"] = df["FinalScore"].apply(
        lambda s: "核心指标" if s >= 0.65 else ("重点关注" if s >= 0.45 else "辅助指标"))
    df["Evidence_Level"] = df.apply(
        lambda r: "强证据" if (r["Top10_Frequency"] >= 0.8 and r["FinalScore"] >= 0.65)
        else ("中等证据" if r["Top10_Frequency"] >= 0.5 else "探索性指标"), axis=1)
    interp = {
        "采食能力": "反映个体采食水平和营养摄入能力，与生长性能和饲料利用效率密切相关。",
        "行为稳定性": "反映采食行为波动程度，可评价个体行为一致性和生产稳定性。",
        "采食节律": "反映昼夜采食分配模式，可揭示个体采食时间策略差异。",
        "采食模式": "反映采食事件组织方式、频率和持续特征，可用于描述行为模式差异。",
        "综合行为指标": "综合反映个体采食行为特征，可作为智能育种候选指标。",
    }
    df["Biological_Interpretation"] = df["Trait_Category"].map(interp)
    df["Breeding_Application"] = df["Recommendation_Level"].map({
        "核心指标": "建议优先纳入智能育种候选性状体系。",
        "重点关注": "建议结合生产性能和遗传参数进一步验证。",
        "辅助指标": "可作为探索性行为表型进行持续积累。",
    })
    return df


def ml_run_all(prod, targets, w_target=None, w_imp=0.7, w_stab=0.3, n_rep=10):
    targets = [t for t in targets if t and t != "---"]
    if not targets:
        return {"error": "请至少选择一个预测目标性状。"}
    if w_target is None or len(w_target) != len(targets):
        w_target = [1 / len(targets)] * len(targets)

    features = ml_features_v5(prod)
    if len(features) < 5:
        return {"error": "行为特征不足（至少 5 个）。"}

    clu = ml_cluster(prod)
    tr = {t: ml_single_target(prod, t, features) for t in targets}
    fusion = ml_fusion(tr, w_target)
    if isinstance(fusion, dict):
        return {"error": fusion.get("error"), "cluster": clu, "targets": tr}

    stab = ml_stability(prod, targets, features, n_rep=n_rep)
    final = ml_final_score(fusion, stab, w_imp, w_stab)
    if isinstance(final, dict):
        return {"error": final.get("error"), "cluster": clu, "targets": tr,
                "fusion": fusion, "stability": stab}

    report = ml_report(final)
    return {"cluster": clu, "targets": tr, "fusion": fusion,
            "stability": stab, "final": final,
            "top10": final.head(10), "report": report, "features": features}


# ============================================================
# 10. 简单模式：育种目标映射 / 异常识别 / 一键全流程
# ============================================================

SIMPLE_GOAL_MAP = {
    "save": ["FCR", "RFI"],
    "grow": ["ADG_g", "FBW_kg", "Gain_kg"],
    "fat": ["SkinFat_Rate", "AbFat_Rate"],
    "meat": ["BMP"],
    "rhythm": ["CV_Daily_FI", "Cosinor_A", "Fano", "Day_FI_Ratio"],
}


def flag_anomalies(prod, index_df=None):
    parts = []
    if "Day_Bout_Ratio" in prod.columns and "TFB_Day" in prod.columns:
        thr = prod["TFB_Day"].quantile(0.05)
        d = prod[(prod["Day_Bout_Ratio"] > 0.97) | (prod["TFB_Day"] < thr)][
            ["Animal_ID", "Day_Bout_Ratio", "TFB_Day"]].copy()
        if len(d):
            d["异常类型"] = "采食节律异常"
            d["具体表现"] = np.where(d["Day_Bout_Ratio"] > 0.97,
                                      "夜间几乎不采食（白天占比>97%）",
                                      "日访饲次数过低（群体最低5%）")
            parts.append(d[["Animal_ID", "异常类型", "具体表现"]])
    if "Gain_kg" in prod.columns:
        thr = prod["Gain_kg"].quantile(0.05)
        d = prod[prod["Gain_kg"] < thr][["Animal_ID"]].copy()
        if len(d):
            d["异常类型"] = "体重增长异常"
            d["具体表现"] = "总增重过低（群体最低5%）"
            parts.append(d)
    if index_df is not None and "Index" in index_df.columns:
        thr = index_df["Index"].quantile(0.05)
        d = index_df[index_df["Index"] < thr][["Animal_ID"]].copy()
        if len(d):
            d["异常类型"] = "综合得分极低"
            d["具体表现"] = "综合得分处于群体最低5%"
            parts.append(d)
    if not parts:
        return pd.DataFrame(columns=["Animal_ID", "异常类型", "具体表现"])
    return pd.concat(parts, ignore_index=True).drop_duplicates()


def run_simple_all(feed_files, bw_files, ped_file=None, idmap_file=None,
                    goals=("save",), ratio=0.3, sex_balance=True,
                    auto_time=True, experiment_start=None,
                    day_start="06:00", day_end="18:00"):
    t0 = time.time()
    target = sorted(set(sum([SIMPLE_GOAL_MAP[g] for g in goals if g in SIMPLE_GOAL_MAP], [])))
    use_ped = ped_file is not None

    cfg = {
        "auto_time_range": auto_time,
        "experiment_start": None if auto_time else experiment_start,
        "training_end": None, "training_time": None,
        "week_length": 7, "feed_sd": 3, "bw_sd": 3,
        "imi_threshold": 300, "min_intake": 1,
        "hff_method": "median", "hff_cutoff": None,
        "day_start": day_start, "day_end": day_end,
        "target_trait": target,
        "indicator_min_r": 0.3, "indicator_max_n": 3,
        "use_pedigree": use_ped, "use_reml_h2": use_ped,
        "fixed_effects": ["Sex"], "retention_ratio": ratio,
        "sex_balance": sex_balance,
    }

    feed_raw = read_matching_excel_files(feed_files, "feed")
    bw_raw = read_matching_excel_files(bw_files, "weight")
    ped = read_pedigree(ped_file) if use_ped else None
    idmap = read_idmap(idmap_file) if idmap_file is not None else None

    clean = run_clean_v5(feed_raw["data"], bw_raw["data"], cfg)
    bout = build_bouts(clean["feed"], cfg["imi_threshold"], cfg["min_intake"])
    feeding = calc_individual_feeding(bout)
    prod = calc_production(clean["feed"], clean["bw"], feeding)
    prod, _ = assign_hff_lff(prod, cfg["hff_method"], cfg["hff_cutoff"])
    prod["HFF"] = (prod["Feed_Frequency_Group"] == "HFF").astype(int)

    for df_func, kwargs in [
        (compute_daynight_v5, {"day_start": day_start, "day_end": day_end}),
        (compute_behavior_cv_v5, {}),
        (compute_cosinor_v5, {}),
        (compute_fano_v5, {}),
    ]:
        df = df_func(bout, **kwargs) if kwargs else df_func(bout)
        drop_cols = [c for c in df.columns if c in prod.columns and c != "Animal_ID"]
        if drop_cols:
            prod = prod.drop(columns=drop_cols)
        prod = prod.merge(df, on="Animal_ID", how="left")

    rhythm = calc_rhythm(bout)
    weekly_fcr = calc_weekly_fcr(clean["feed"], clean["bw"], cfg["week_length"])
    link = link_animals(clean["feed"], clean["bw"], ped, idmap)
    if "eID" in link.columns and "Sex" in link.columns:
        prod = prod.merge(link[["eID", "Sex"]].drop_duplicates("eID"),
                          left_on="Animal_ID", right_on="eID", how="left").drop(columns=["eID"], errors="ignore")

    MAX_TRAITS = 8
    meas = [t for t in target if t in prod.columns][:MAX_TRAITS]
    ind_slots = MAX_TRAITS - len(meas)
    ind_traits = (select_indicator_traits(prod, target, cfg["indicator_min_r"], cfg["indicator_max_n"])
                  if ind_slots > 0 else pd.DataFrame(columns=["Indicator"]))
    if len(ind_traits) > ind_slots:
        ind_traits = ind_traits.iloc[:ind_slots]
    idx_traits = list(dict.fromkeys(list(meas) + list(ind_traits["Indicator"])))
    n_meas, n_ind = len(meas), len(idx_traits) - len(meas)

    w = {}
    if n_ind > 0 and n_meas > 0:
        for t in meas:
            w[t] = 0.7 / n_meas
        for t in idx_traits:
            if t not in meas:
                w[t] = 0.3 / n_ind
        weights_note = "目标性状合计70% / 自动补充性状合计30%（组内均分）"
    else:
        for t in idx_traits:
            w[t] = 1 / len(idx_traits)
        weights_note = "参与打分性状均分"

    cfg["index_traits"] = idx_traits
    cfg["weights"] = w
    rr = run_retention_module(prod, link, cfg)

    growth = clean["bw_week_stats"]
    growth_plot_data = growth[["Week_Number", "Mean"]].rename(
        columns={"Week_Number": "Week", "Mean": "Mean_BW"})

    summary = pd.DataFrame({
        "指标": ["评估路线", "参与打分性状数", "参与打分性状", "总个体", "留种个体",
                 "留种比例", "用时(秒)"],
        "值": [rr["genetic"]["mode"], len(idx_traits), "、".join(idx_traits),
               rr["retention"]["total"], rr["retention"]["n_keep"],
               f"{round(rr['retention']['ratio'] * 100)}%",
               str(round(time.time() - t0, 1))],
    })

    return {
        "A": {
            "summary": summary,
            "retained": rr["retention"]["retained"],
            "ratio": rr["retention"]["ratio"],
            "all_index": rr["genetic"]["index"],
            "indicator_traits": rr["indicator_traits"],
            "weights_note": weights_note,
            "warnings": rr["warnings"],
        },
        "B": {
            "week_feed": clean["feed_week_stats"],
            "week_bw": clean["bw_week_stats"],
            "weekly_fcr": weekly_fcr,
            "growth": growth_plot_data,
        },
        "C": {
            "rhythm_daily": rhythm["daily"],
            "rhythm_hourly": rhythm["hourly"],
            "daynight": prod[[c for c in ["Animal_ID", "Day_FI_g", "Night_FI_g", "Day_Bouts",
                                           "Night_Bouts", "Day_FI_Ratio", "Day_Bout_Ratio"]
                              if c in prod.columns]],
            "cv": prod[[c for c in ["Animal_ID", "CV_Duration", "CV_FR", "CV_Daily_Bouts",
                                     "CV_Daily_FI", "CV_Daily_TFD"] if c in prod.columns]],
            "anomalies": flag_anomalies(prod, rr["genetic"]["index"]),
        },
        "prod": prod,
        "elapsed": time.time() - t0,
    }


def export_simple_xlsx(res):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        sheets = {
            "A_留种名单": res["A"]["retained"],
            "A_全部个体得分": res["A"]["all_index"],
            "A_指示性状": res["A"]["indicator_traits"],
            "B_周统计_采食": res["B"]["week_feed"],
            "B_周统计_体重": res["B"]["week_bw"],
            "B_每周FCR": res["B"]["weekly_fcr"],
            "C_昼夜分配": res["C"]["daynight"],
            "C_行为变异性": res["C"]["cv"],
            "C_异常个体": res["C"]["anomalies"],
        }
        for name, df in sheets.items():
            if df is not None and len(df):
                df.to_excel(w, sheet_name=name[:31], index=False)
    buf.seek(0)
    return buf
