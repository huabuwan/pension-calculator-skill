#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
养老金测算工作簿生成脚本
======================
读取 inputs.json（由 agent 根据对话/解析文件后填写）+ data/社平工资表.xlsx（预缓存），
生成一个 Excel 工作簿：
  Sheet1 测算说明  —— 计算规则、政策依据、参数含义、数据来源与免责
  Sheet2 缴纳明细  —— 逐段/逐年缴纳记录（含预估待缴纳），顶部参数区，底部计算结果
  Sheet3 退休金预估 —— 从退休年起逐年（到 99 岁）月养老金、年养老金、累计领取

设计要点
--------
1. 指数分母一律用「待遇领取地社平」（人社部2009转移接续意见第二条口径）。
2. 历史社平用缓存实际值（静态）；未来社平 = 实际值×(1+社平增幅)^年差。
3. 未来（预估）段的月缴费基数 = 当年待遇领取地社平 × 灵活档位。
4. 指数封顶3.0 默认不封顶（FALSE），可手动改为 TRUE。

输出模式（重要）
----------------
默认 OUTPUT_VALUES=True：所有计算列直接写成「计算好的数值」，任何查看器（含腾讯文档预览）
打开即见正确数字，不会出现活公式未计算而显示 0 的问题。
若需要 Excel 内「改黄色参数即时重算」的活公式，请设置（二选一）：
  - 环境变量  PENSION_OUTPUT=formula
  - inputs.json 顶层字段  "output_mode": "formula"
活公式模式下本脚本仍写出引用参数的公式；数值模式则写出已算好的值（改参数需重跑本脚本）。

输入契约（inputs.json）
----------------------
{
  "retire_city": "北京",        # 待遇领取地（指数分母城市），必填
  "retire_age": 60,             # 退休年龄，必填
  "retire_year": 2040,          # 退休（开始领取）年份，必填
  "manual_sheping": {},         # 可选：{城市:{年份:月社平}} 覆盖/补充缓存（缓存缺城市时用）
  "output_mode": "values",      # 可选："values"(默认，写数值) / "formula"(写活公式)
  "params": {
    "sheping_growth": 0.02,     # 社平增幅（年化），默认 2%
    "pension_growth": 0.02,     # 养老增幅（年度调整），默认 2%
    "interest_rate": 0.025,     # 记账利率默认（历史段可逐年不同，见 yearly）
    "flex_ratio": 0.60,         # 灵活就业档位（未来段用，0.6=60%）
    "cap_index": false          # 指数封顶3.0（默认不封顶）
  },
  # 两种提供方式任选其一：
    "years": [                    # 方式A：agent 已归一化的逐年记录（推荐，最灵活）
    {"year":2010,"city":"上海","type":"单位职工","months":12,
     "base_monthly":9000, "rate":0.03, "is_estimated":false},
    {"year":2040,"city":"北京灵活","type":"灵活就业","months":12,
     "is_estimated":true},        # 预估段：base_monthly 留空，脚本按 社平×灵活档位 算
    ...
  ],
  "segments": [                 # 方式B：分段（脚本展开为逐年）
    {"year_start":2006,"month_start":7,"year_end":2014,"month_end":6,
     "city":"北京","type":"单位职工","base_monthly":8202,"rate":0.03,
     "is_estimated":false},
    ...
  ]
}
说明：
- 单位职工 / 灵活就业 均可；灵活就业 20% 中 8% 进个人账户（脚本统一按 月基数×8%×月数 入个账）。
- 若某年只知「个人账户年入账 account_entry」不知月基数，可在 yearly 记录里给 account_entry，
  脚本反推 月基数 = account_entry / (0.08 × months)。
- 预估段（is_estimated=true）：base_monthly 不填，脚本用 当年待遇领取地社平 × 灵活档位 计算月基数。
"""

import sys
import os
import json
import math

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(HERE, "..", "data", "社平工资表.xlsx")

# 计发月数（按退休年龄）
FAMING = {40: 233, 45: 216, 50: 195, 55: 170, 60: 139, 65: 101, 70: 56}

# 社保官方最新公布社平年（缓存截止年）；该年之后的年份一律由"社平增幅"参数推算，
# 即使 data/社平工资表.xlsx 里填了展示值也忽略，以保证与参数一致。
OFFICIAL_LAST_YEAR = 2023

# 城市别名映射：用户输入的常见写法 -> 缓存表标准城市名（因为缓存以省全口径为单位，
# 成都执行四川、广州执行广东等）。未知城市保持原样（由 manual_sheping 或联网补充）。
CITY_ALIAS = {
    '成都': '四川', '广州': '广东', '杭州': '浙江', '宁波': '浙江',
    '南京': '江苏', '苏州': '江苏', '无锡': '江苏', '武汉': '湖北',
    '西安': '陕西', '沈阳': '沈阳', '大连': '大连', '深圳': '深圳',
    '重庆': '重庆', '天津': '天津', '上海': '上海', '北京': '北京',
}


def norm_city(c):
    return CITY_ALIAS.get(c, c)


thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FILL = PatternFill("solid", fgColor="44546A")
HEAD_FONT = Font(bold=True, color="FFFFFF")
PARAM_FILL = PatternFill("solid", fgColor="FFF2CC")
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
SUB_FONT = Font(bold=True, size=12, color="1F4E78")
NOTE_FONT = Font(italic=True, size=9, color="808080")
EST_FILL = PatternFill("solid", fgColor="FCE4D6")  # 预估段浅橙标记


def load_sheping():
    """返回 {city: {year: val}} 以及每城市最后一个实际年。"""
    wb = openpyxl.load_workbook(DATA_PATH)
    ws = wb["社平"]
    cities = []
    for c in range(2, ws.max_column + 1):
        cities.append(ws.cell(row=1, column=c).value)
    data = {ci: {} for ci in cities}
    for r in range(2, ws.max_row + 1):
        yr = ws.cell(row=r, column=1).value
        if yr is None:
            continue
        for c, ci in enumerate(cities, start=2):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (int, float)):
                data[ci][int(yr)] = float(v)
    return data


def last_actual_year(data, city):
    if city in data and data[city]:
        return max(data[city].keys())
    return None


def resolve_sheping_value(data, manual, city, year, growth):
    """返回某个城市某年的社平：实际值，或基于最近实际年按 growth 推算的静态值。"""
    if manual and city in manual and year in manual[city]:
        return float(manual[city][year])
    if city in data and year in data[city]:
        return data[city][year]
    la = last_actual_year(data, city)
    if la is None:
        return None
    base = data[city][la]
    return base * ((1 + growth) ** (year - la))


def build():
    if len(sys.argv) >= 2:
        inputs_path = sys.argv[1]
    else:
        inputs_path = "inputs.json"
    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
    else:
        out_path = None

    with open(inputs_path, "r", encoding="utf-8") as f:
        inp = json.load(f)

    # 必填字段校验（缺失时给出清晰中文报错，避免 KeyError 难以定位）
    _required = ["retire_city", "retire_age", "retire_year"]
    _missing = [k for k in _required if k not in inp or inp.get(k) in (None, "")]
    if _missing:
        raise ValueError(
            "inputs.json 缺少必填字段：%s。\n"
            "retire_city=待遇领取地城市（指数分母）、retire_age=退休年龄、retire_year=开始领取年份，"
            "这三项必须从用户处采集，不能为空。" % "、".join(_missing)
        )

    retire_city = norm_city(inp["retire_city"])
    retire_age = int(inp["retire_age"])
    retire_year = int(inp["retire_year"])
    retire_base_year = retire_year - 1  # 退休时社平取退休前一年（最新公布年）
    manual = inp.get("manual_sheping", {}) or {}
    p = inp.get("params", {})
    sheping_growth = float(p.get("sheping_growth", 0.02))
    pension_growth = float(p.get("pension_growth", 0.02))
    interest_default = float(p.get("interest_rate", 0.025))
    flex_ratio = float(p.get("flex_ratio", 0.60))
    cap_index = bool(p.get("cap_index", False))
    fa = FAMING.get(retire_age, 139)

    # 输出模式：默认数值（values）；formula 走活公式
    output_mode = (inp.get("output_mode") or os.environ.get("PENSION_OUTPUT", "values")).lower()
    OUTPUT_VALUES = output_mode != "formula"

    def W(cell, formula, value):
        """写单元格：数值模式写计算好的值，公式模式写活公式。"""
        cell.value = value if OUTPUT_VALUES else formula

    # 合并数据（手动覆盖叠加到缓存，key 也做城市名归一）
    data = load_sheping()
    for ci, ymap in (manual or {}).items():
        nci = norm_city(ci)
        data.setdefault(nci, {})
        for y, v in ymap.items():
            data[nci][int(y)] = float(v)

    # 归一化逐年记录
    yearly = []
    if inp.get("years"):
        yearly = inp["years"]
    elif inp.get("segments"):
        for seg in inp["segments"]:
            ys, ms = int(seg["year_start"]), int(seg["month_start"])
            ye, me = int(seg["year_end"]), int(seg["month_end"])
            city = seg["city"]
            typ = seg.get("type", "单位职工")
            is_est = bool(seg.get("is_estimated", False))
            seg_rate = seg.get("rate", None)
            seg_flex = seg.get("flex_ratio", None)
            base = seg.get("base_monthly", None)
            for yr in range(ys, ye + 1):
                if yr == ys and yr == ye:
                    months = me - ms + 1
                elif yr == ys:
                    months = 12 - ms + 1
                elif yr == ye:
                    months = me
                else:
                    months = 12
                rec = {
                    "year": yr, "city": city, "type": typ,
                    "months": months, "is_estimated": is_est,
                    "rate": seg_rate, "flex_ratio": seg_flex,
                }
                if base is not None:
                    rec["base_monthly"] = base
                if "account_entry" in seg:
                    rec["account_entry"] = seg["account_entry"]
                yearly.append(rec)
    else:
        raise ValueError("inputs.json 必须提供 years 或 segments")

    # 逐年记录城市名归一（成都->四川 等）；保留原始城市名用于显示，_norm 用于社平查询
    for rec in yearly:
        rec["_norm"] = norm_city(rec.get("city"))

    # 校验 retire_city 是否在缓存（或手动）
    if retire_city not in data or not data[retire_city]:
        raise ValueError(
            "待遇领取地「%s」不在社平缓存且未提供 manual_sheping。"
            "请联网查询该城市“全口径城镇单位就业人员月平均工资”并在 inputs.json 的 "
            "manual_sheping 中按 {城市:{年份:月社平}} 补充，或用已有缓存城市。" % retire_city
        )

    # 该城市"官方截止年"（最近 ≤ OFFICIAL_LAST_YEAR 的实际年），用于未来年公式基准
    la_city = max(y for y in data[retire_city] if y <= OFFICIAL_LAST_YEAR)
    la_val = data[retire_city][la_city]

    # ---------------- 创建工作簿 ----------------
    wb = openpyxl.Workbook()

    # ===== Sheet1 测算说明 =====
    sh = wb.active
    sh.title = "测算说明"
    sh.column_dimensions["A"].width = 22
    sh.column_dimensions["B"].width = 96
    rows = [
        ("养老金测算 · 测算说明", "title"),
        ("", ""),
        ("一、本表能做什么", "sub"),
        ("根据您的社保缴纳历史与未来规划，估算退休后的月基本养老金，并从退休年起逐年（到 99 岁）"
         "推演月/年养老金与累计领取总额。所有关键假设都已固化到表中，打开即见结果。", ""),
        ("", ""),
        ("二、关键参数（见「缴纳明细」顶部参数区，黄色单元格为展示用）", "sub"),
        ("退休年龄 / 计发月数", "退休年龄决定计发月数（55→170、60→139、50→195、65→101）。计发月数越大每月越少。"),
        ("社平增幅", "各地社会平均工资每年的涨幅假设，用于推算“未公布的未来年份社平”和“退休时社平”。默认 2%。"),
        ("养老增幅", "国家每年统一上调养老金的幅度假设，用于终身领取推演。默认 2%（近年实际约 2%–3.8%）。"),
        ("记账利率", "个人账户历年结余的计息利率，默认 2.5%（近年实际约 2%–3%）。历史段可按实际年填写。"),
        ("灵活就业档位", "未来若以灵活就业身份缴纳，缴费基数 = 当年待遇领取地社平 × 该档位（60%–300%）。"),
        ("指数封顶3.0", "跨省到较低社平城市退休时，单年缴费指数是否按 3.0 封顶。默认 FALSE（不封顶），可改 TRUE。"),
        ("", ""),
        ("三、计算规则与公式", "sub"),
        ("① 月缴费指数", "= 当年月缴费基数 ÷ 待遇领取地对应年度社平工资（分母统一用“待遇领取地”，即归集退休地；"
         "跨省归集时北京/上海段高基数会被放大）。若开启封顶，单年指数超过 3.0 按 3.0 计。"),
        ("② 平均缴费指数", "= Σ(每年月缴费指数 × 当年缴费月数) ÷ 总缴费月数。"),
        ("③ 指数化工资", "= 退休时社平工资 × 平均缴费指数。"),
        ("④ 基础养老金", "= (退休时社平 + 指数化工资) ÷ 2 × 缴费年限(年) × 1%。"),
        ("⑤ 个人账户养老金", "= 个人账户累计额 ÷ 计发月数。"),
        ("⑥ 月基本养老金", "= 基础养老金 + 个人账户养老金。"),
        ("⑦ 灵活就业缴费", "按 20% 比例，其中 8% 进入个人账户（与单位职工个人部分一致）。"),
        ("⑧ 终身领取推演", "退休后每年月养老金 = 退休首年月养老金 ×(1+养老增幅)^年数；累计领取逐年累加。"),
        ("", ""),
        ("四、政策依据", "sub"),
        ("", "缴费工资指数口径：人社部《关于城镇企业职工基本养老保险关系转移接续若干问题的意见》"
         "（2009）第二条——跨省转移接续人员退休时，缴费工资指数 = 本人在各参保地的缴费工资 ÷ "
         "待遇领取地对应年度职工平均工资。故本表指数分母统一取“待遇领取地社平”。"),
        ("", "基本养老金计发：国发〔2005〕38 号。个人账户记账利率：人社部公布（近年约 2%–3%）。"),
        ("", ""),
        ("五、数据来源与免责", "sub"),
        ("", "社平工资来自 skill 内置缓存（data/社平工资表.xlsx，共 13 城、1989–2033 逐年补全，1989–2005 为历史回溯估算）："
         "北京/上海/四川为官方实际序列；其余城市锚点年为官方值、其余年份按约5%（历史）/2%（未来）估算。"
         "2024 年起一律按“社平增幅”参数推算，与缓存展示值无关。"),
        ("", "⚠ 缓存为通用参考，可能与您本人社保局系统精确值有出入。请以参保地社保局公布的“全口径社平工资”"
         "核实；如在 Excel 中用活公式模式，可直接修改对应年份社平，结果会即时重算。"),
        ("", "本表为个人规划辅助工具，非官方核算结果，不构成任何投保/退休决策建议。"),
    ]
    r = 1
    for a, b in rows:
        ca = sh.cell(row=r, column=1, value=a)
        if b in ("title", "sub"):
            # 样式标记：仅用于给 A 列设字体，B 列留空，不要把标记文字写进单元格
            ca.font = TITLE_FONT if b == "title" else SUB_FONT
        else:
            # 以 "=" 开头的描述串加单引号前缀，强制按文本显示（Excel 不会当公式求值出 #NAME?）
            bval = b
            if isinstance(bval, str) and bval.startswith("="):
                bval = "'" + bval
            cb = sh.cell(row=r, column=2, value=bval)
            ca.font = Font(bold=True, size=10)
            cb.font = Font(size=10)
            cb.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    # ===== Sheet2 缴纳明细 =====
    det = wb.create_sheet("缴纳明细")
    # 参数区
    det["A1"] = "参数区（黄色单元格为关键假设，仅供查看；改假设请重跑脚本或切 formula 模式）"
    det["A1"].font = SUB_FONT
    params = [
        ("退休年龄", retire_age, "岁"),
        ("计发月数", fa, "月"),
        ("社平增幅(年化)", sheping_growth, "如 0.02=2%"),
        ("养老增幅(年化)", pension_growth, "如 0.02=2%"),
        ("记账利率(默认)", interest_default, "历史段可逐年不同"),
        ("灵活就业档位(未来段)", flex_ratio, "0.6=60%，0.6–3.0"),
        ("指数封顶3.0(TRUE/FALSE)", cap_index, "默认不封顶"),
        ("待遇领取地", retire_city, "指数分母城市"),
        ("退休基准年(退休时社平年)", retire_base_year, "取退休前一年"),
        ("退休起始年(开始领取年)", retire_year, "退休当年/次年起领"),
    ]
    for i, (lab, val, unit) in enumerate(params):
        rr = 3 + i
        det.cell(row=rr, column=1, value=lab).font = Font(bold=True)
        c = det.cell(row=rr, column=2, value=val)
        c.fill = PARAM_FILL
        c.font = Font(bold=True)
        det.cell(row=rr, column=3, value=unit).font = NOTE_FONT
    # 参数单元格引用
    P_AGE, P_FA, P_SG, P_PG, P_IR, P_FLEX, P_CAP, P_CITY, P_BASE, P_START = (
        "$B$3", "$B$4", "$B$5", "$B$6", "$B$7", "$B$8", "$B$9", "$B$10", "$B$11", "$B$12"
    )

    # 表头
    hdr_row = 13
    headers = ["年份", "参保地", "类型", "性质", "月数", "月缴费基数",
               "年缴费基数", "个人年缴费(8%)", "当地社平(参考)", "待遇领取地社平(分母)",
               "月缴费指数", "指数加权(指数×月数)", "记账利率", "个账累计", "备注"]
    for ci, h in enumerate(headers, start=1):
        c = det.cell(row=hdr_row, column=ci, value=h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

    # 数据行
    data_start = hdr_row + 1
    rr = data_start
    # 累计器（数值模式使用）
    sum_E = 0.0
    sum_L = 0.0
    prev_n = 0.0
    retire_base_jval = None  # 退休基准年对应行的待遇地社平
    last_jval = None
    for rec in yearly:
        yr = int(rec["year"])
        city = rec.get("_norm", norm_city(rec.get("city")))  # 社平查询用归一名
        city_disp = rec.get("city")                          # 显示用原始城市名
        typ = rec.get("type", "单位职工")
        months = int(rec["months"])
        is_est = bool(rec.get("is_estimated", False))

        # 待遇领取地社平（retire_city）
        if yr > OFFICIAL_LAST_YEAR:
            j_val = la_val * ((1 + sheping_growth) ** (yr - la_city))
            j_formula = "=%.2f*(1+%s)^%d" % (round(la_val, 2), P_SG, yr - la_city)
        else:
            j_val = resolve_sheping_value(data, manual, retire_city, yr, sheping_growth)
            j_formula = round(j_val, 2)
        last_jval = j_val

        # 当地社平（row city）
        if city in data and data[city]:
            la_c = max(y for y in data[city] if y <= OFFICIAL_LAST_YEAR)
            if yr > OFFICIAL_LAST_YEAR:
                i_val = data[city][la_c] * ((1 + sheping_growth) ** (yr - la_c))
                i_formula = "=%.2f*(1+%s)^%d" % (round(data[city][la_c], 2), P_SG, yr - la_c)
            else:
                i_val = resolve_sheping_value(data, manual, city, yr, sheping_growth)
                i_formula = round(i_val, 2)
        else:
            i_val = None
            i_formula = None

        # 月缴费基数 F：base_monthly 优先（历史段或预估段中用户指定的基数）；
        # 仅当未给定 base_monthly 的预估段，才按 待遇地社平 × 灵活档位 计算
        if rec.get("base_monthly") is not None:
            f_val = float(rec["base_monthly"])
            f_formula = f_val
        elif is_est:
            f_val = j_val * flex_ratio
            f_formula = "=J%d*%s" % (rr, P_FLEX)
        elif rec.get("account_entry") is not None:
            ae = float(rec["account_entry"])
            f_val = ae / (0.08 * months) if months else 0
            f_formula = round(f_val, 2)
        else:
            f_val = None
            f_formula = None

        # 记账利率 M
        if is_est:
            m_val = interest_default
            m_formula = "=%s" % P_IR
        else:
            m_val = float(rec["rate"]) if rec.get("rate") is not None else interest_default
            m_formula = m_val

        # 派生量
        k_val = (min(f_val / j_val, 3.0) if cap_index else f_val / j_val) if (f_val is not None and j_val) else None
        g_val = f_val * months if f_val is not None else None
        h_val = f_val * 0.08 * months if f_val is not None else None
        l_val = k_val * months if k_val is not None else None

        # 个账累计 N（链式）
        if h_val is not None:
            prev_n = prev_n * (1 + m_val) + h_val
        n_val = prev_n
        n_formula = "=H%d" % rr if rr == data_start else "=N%d*(1+M%d)+H%d" % (rr - 1, rr, rr)

        # 累计器
        sum_E += months
        sum_L += (l_val if l_val is not None else 0.0)
        if int(rec["year"]) == retire_base_year:
            retire_base_jval = j_val

        # 写单元格
        det.cell(row=rr, column=1, value=yr)
        det.cell(row=rr, column=2, value=city_disp)
        det.cell(row=rr, column=3, value=typ)
        det.cell(row=rr, column=4, value="预估待缴" if is_est else "历史")
        det.cell(row=rr, column=5, value=months)  # E 月数为直接数值
        W(det.cell(row=rr, column=6), f_formula, f_val)
        W(det.cell(row=rr, column=7), "=F%d*E%d" % (rr, rr), g_val)
        W(det.cell(row=rr, column=8), "=F%d*0.08*E%d" % (rr, rr), h_val)
        W(det.cell(row=rr, column=9), i_formula, i_val)
        W(det.cell(row=rr, column=10), j_formula, j_val)
        W(det.cell(row=rr, column=11),
          "=IF(%s,MIN(F%d/J%d,3),F%d/J%d)" % (P_CAP, rr, rr, rr, rr), k_val)
        W(det.cell(row=rr, column=12), "=K%d*E%d" % (rr, rr), l_val)
        W(det.cell(row=rr, column=13), m_formula, m_val)
        W(det.cell(row=rr, column=14), n_formula, n_val)
        det.cell(row=rr, column=15, value="预估" if is_est else "")
        # 预估段着色
        if is_est:
            for cc in range(1, 16):
                det.cell(row=rr, column=cc).fill = EST_FILL
        # 数字格式
        for cc in (6, 7, 8, 9, 10, 14):
            det.cell(row=rr, column=cc).number_format = "#,##0"
        det.cell(row=rr, column=11).number_format = "0.0000"
        det.cell(row=rr, column=12).number_format = "0.00"
        det.cell(row=rr, column=13).number_format = "0.0%"
        for cc in range(1, 16):
            det.cell(row=rr, column=cc).border = border
        rr += 1
    data_end = rr - 1

    # 合计行
    tot = data_end + 1
    det.cell(row=tot, column=1, value="合计").font = Font(bold=True)
    det.cell(row=tot, column=4, value="——").font = Font(bold=True)
    W(det.cell(row=tot, column=5), "=SUM(E%d:E%d)" % (data_start, data_end), sum_E)
    W(det.cell(row=tot, column=12), "=SUM(L%d:L%d)" % (data_start, data_end), sum_L)
    det.cell(row=tot, column=5).number_format = "#,##0"
    det.cell(row=tot, column=12).number_format = "0.00"
    det.cell(row=tot, column=5).font = Font(bold=True)
    det.cell(row=tot, column=12).font = Font(bold=True)

    # 计算结果块
    a0 = tot + 2
    # 退休时社平：优先退休基准年对应行；否则取最后一段
    retire_j = retire_base_jval if retire_base_jval is not None else last_jval
    avg_idx = (sum_L / sum_E) if sum_E else 0.0
    idx_sal = retire_j * avg_idx
    years_n = sum_E / 12.0
    base_p = (retire_j + idx_sal) / 2 * years_n * 0.01
    acct_total = prev_n
    acct_p = acct_total / fa if fa else 0.0
    monthly_pension = base_p + acct_p
    annual_pension = monthly_pension * 12

    det.cell(row=a0, column=1, value="—— 计算结果（退休首年） ——").font = SUB_FONT
    agg = [
        ("总缴费月数", "=SUM(E%d:E%d)" % (data_start, data_end), "#,##0", "月", sum_E),
        ("平均缴费指数", "=SUM(L%d:L%d)/SUM(E%d:E%d)" % (data_start, data_end, data_start, data_end),
         "0.0000", "", avg_idx),
        ("退休时社平(元/月)", "=J%d" % _retire_base_row(yearly, retire_base_year, data_start),
         "#,##0", "", retire_j),
        ("指数化工资(元/月)", "=B%d*B%d" % (a0 + 3, a0 + 2), "#,##0", "", idx_sal),
        ("缴费年限(年)", "=B%d/12" % (a0 + 1), "0.00", "年", years_n),
        ("基础养老金(元/月)", "=(B%d+B%d)/2*B%d*0.01" % (a0 + 3, a0 + 4, a0 + 5),
         "#,##0", "", base_p),
        ("个人账户累计(元)", "=N%d" % data_end, "#,##0", "", acct_total),
        ("个人账户养老金(元/月)", "=B%d/%s" % (a0 + 7, P_FA), "#,##0", "", acct_p),
        ("月基本养老金(元/月)", "=B%d+B%d" % (a0 + 6, a0 + 8), "#,##0", "", monthly_pension),
        ("年基本养老金(元/年)", "=B%d*12" % (a0 + 9), "#,##0", "", annual_pension),
    ]
    for i, (lab, formula, fmt, unit, val) in enumerate(agg):
        rr = a0 + 1 + i
        det.cell(row=rr, column=1, value=lab).font = Font(bold=True)
        c = det.cell(row=rr, column=2)
        W(c, formula, val)
        c.number_format = fmt
        c.font = Font(bold=True, color="C00000")
        det.cell(row=rr, column=3, value=unit).font = NOTE_FONT
    monthly_pension_row = a0 + 9  # 月基本养老金 所在行（B列）

    # 列宽
    widths = [8, 14, 10, 10, 7, 12, 12, 14, 14, 16, 12, 18, 10, 14, 8]
    for ci, w in enumerate(widths, start=1):
        det.column_dimensions[get_column_letter(ci)].width = w
    det.freeze_panes = "A%d" % (hdr_row + 1)

    # ===== Sheet3 退休金预估 =====
    proj = wb.create_sheet("退休金预估")
    proj["A1"] = "退休金终身领取推演（从 %d 岁到 99 岁）" % retire_age
    proj["A1"].font = TITLE_FONT
    proj["A2"] = ("月基本养老金(退休首年) 取自「缴纳明细」计算结果（%.2f 元/月）；逐年按“养老增幅”复利增长；"
                  "累计领取逐年累加。" % monthly_pension)
    proj["A2"].font = NOTE_FONT
    ph = 3
    phdr = ["年龄", "对应年度", "月养老金(元)", "年养老金(元)", "累计领取(元)"]
    for ci, h in enumerate(phdr, start=1):
        c = proj.cell(row=ph, column=ci, value=h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = border
    MONTHLY_REF = "缴纳明细!B%d" % monthly_pension_row
    PG_REF = P_PG  # 养老增幅
    first = ph + 1
    cum = 0.0
    for off, age in enumerate(range(retire_age, 100)):
        r = first + off
        cm = monthly_pension * ((1 + pension_growth) ** off)  # 当年月养老金
        cy = cm * 12
        cum += cy
        proj.cell(row=r, column=1, value=age)
        W(proj.cell(row=r, column=2), "=%s+(A%d-%d)" % (P_START, r, retire_age), retire_year + off)
        W(proj.cell(row=r, column=3), "=%s*(1+%s)^(A%d-%d)" % (MONTHLY_REF, PG_REF, r, retire_age), cm)
        W(proj.cell(row=r, column=4), "=C%d*12" % r, cy)
        if off == 0:
            W(proj.cell(row=r, column=5), "=D%d" % r, cum)
        else:
            W(proj.cell(row=r, column=5), "=E%d+D%d" % (r - 1, r), cum)
        for ci in (3, 4, 5):
            proj.cell(row=r, column=ci).number_format = "#,##0"
        proj.cell(row=r, column=2).alignment = Alignment(horizontal="center")
        for ci in range(1, 6):
            proj.cell(row=r, column=ci).border = border
    pw = [8, 10, 16, 16, 18]
    for ci, w in enumerate(pw, start=1):
        proj.column_dimensions[get_column_letter(ci)].width = w
    proj.freeze_panes = "C%d" % first

    wb.calculation.fullCalcOnLoad = True

    if out_path is None:
        out_path = "养老金测算_%s.xlsx" % retire_city
    wb.save(out_path)
    print("SAVED:", out_path, "(模式: %s)" % ("values" if OUTPUT_VALUES else "formula"))
    print("  退休地=%s 退休年龄=%d 退休年=%d 计发月数=%d" % (retire_city, retire_age, retire_year, fa))
    print("  数据行 %d–%d（%d 段/年）；预估段已标记。" % (data_start, data_end, len(yearly)))
    print("  月基本养老金(退休首年) ≈ %.2f 元/月" % monthly_pension)


def _retire_base_row(yearly, base_year, data_start):
    """找到退休基准年对应的数据行号；找不到则取最后数据行。"""
    for i, rec in enumerate(yearly):
        if int(rec["year"]) == base_year:
            return data_start + i
    return data_start + len(yearly) - 1


if __name__ == "__main__":
    build()
