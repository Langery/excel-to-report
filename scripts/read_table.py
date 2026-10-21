#!/usr/bin/env python3
"""read_table.py — 读取 Excel/CSV 表格，输出结构化 JSON。

用法:
    python read_table.py <表格路径> [--sheet 表名] [--rows N] [--group 列名]
    python read_table.py <表格路径> --audit [超期检测参数]
    python read_table.py --selftest

模式:
    默认模式    : 输出列名、行数、类型推断、非空/唯一值统计、前 N 行样例
    --group 模式 : 按指定列分组，输出每组的行数、数值列统计、日期范围
    --audit 模式 : 异常候选信号（重复值 / 空值 / 离群值 / 超期），逐条给出精确位置
    --selftest   : 内置自测用例，不依赖任何外部文件；退出码 0 = 全部通过

核心纪律（对应 SKILL.md「绝不编造数字」）:
    1. 分组被合并时，合并条目必须携带各数值列的聚合值，且 avg 按行数加权
    2. bucketed_by_month 只在「确实按月聚合了」时为 true，绝不谎报
    3. 日期识别与月份归一化共用同一套解析，识别为日期就一定能聚合出 YYYY-MM
    4. --audit 只输出「候选信号 + 精确位置 + 方法名」，不输出结论；语义（如
       「什么状态算未完成」）必须由调用方显式传入，脚本一律不猜

输出: 统一 JSON 到 stdout。成功 {"ok": true, ...}，失败 {"ok": false, "error": "...", "hint": "..."}

依赖: CSV 路径纯标准库；xlsx 路径需要 openpyxl（可选增强，见 requirements.txt）。
"""
import argparse
import csv
import json
import math
import re
import sys
from datetime import date, datetime

MAX_GROUPS = 30          # 分组数超过此值时合并展示 Top N
TOP_GROUPS = 20
TYPE_SAMPLE = 200        # 类型推断最多采样的非空值数
AUDIT_MAX_ITEMS = 50     # --audit 每个检测项最多列出的明细条数（总数照实报告）
IQR_K = 1.5              # IQR 法的 K 倍数
ZSCORE_THRESHOLD = 3.0   # z-score 法的阈值（|z| > 阈值判为候选）


def out(data):
    print(json.dumps(data, ensure_ascii=False, indent=2, default=str))


def fail(error, hint=""):
    out({"ok": False, "error": error, "hint": hint})
    sys.exit(1)


# ---------------------------------------------------------------- 日期解析

# 说明：正则而非 strptime 格式表，是为了同时覆盖「年-月」这类不完整日期
# （YYYY-MM / YYYY年M月 在职场月度台账里极常见，strptime 无法解析）。
# 所有格式统一归一化到 (年, 月)，日期缺日时 day=None。

_RE_YMD = re.compile(
    r"^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?"
    r"(?:[\sT]+\d{1,2}[:：]\d{2}(?:[:：]\d{2})?)?$"
)
_RE_YM = re.compile(r"^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*月?$")
# 月/日/年 与 日/月/年 在「两者都 <= 12」时天然歧义，无法从字符串判定，
# 只能按「月在前」这一常见写法解析，并把歧义如实报给调用方复核。
_RE_MDY = re.compile(r"^(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{4})$")


def parse_temporal(v):
    """把单元格值解析为 {year, month, day, precision, ambiguous}；失败返回 None。

    支持: YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD / YYYY年M月D日 / 带时分秒的日期 /
          YYYY-MM / YYYY/MM / YYYY.MM / YYYY年M月 / MM/DD/YYYY / DD/MM/YYYY /
          datetime 与 date 对象
    precision: "day"（含日，可判超期）| "month"（仅到月，判超期时不可用）
    """
    if isinstance(v, datetime):
        return {"year": v.year, "month": v.month, "day": v.day,
                "precision": "day", "ambiguous": False}
    if isinstance(v, date):
        return {"year": v.year, "month": v.month, "day": v.day,
                "precision": "day", "ambiguous": False}
    s = str(v).strip()
    if not s:
        return None

    m = _RE_YMD.match(s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if _valid_ymd(y, mo, d):
            return {"year": y, "month": mo, "day": d,
                    "precision": "day", "ambiguous": False}

    m = _RE_YM.match(s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return {"year": y, "month": mo, "day": None,
                    "precision": "month", "ambiguous": False}

    m = _RE_MDY.match(s)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        mo, d = a, b  # 按「月在前」解析
        if _valid_ymd(y, mo, d):
            return {"year": y, "month": mo, "day": d, "precision": "day",
                    "ambiguous": a <= 12 and b <= 12}
    return None


def _valid_ymd(y, mo, d):
    if not (1 <= mo <= 12):
        return False
    try:
        date(y, mo, d)
        return True
    except ValueError:
        return False


def month_key(v):
    """归一化为 YYYY-MM；无法识别为日期时返回 None（绝不返回半成品 key）。"""
    t = parse_temporal(v)
    if t is None:
        return None
    return "%04d-%02d" % (t["year"], t["month"])


def day_key(v):
    """归一化为 YYYY-MM-DD；只有含日的值才返回，月份级值返回 None。"""
    t = parse_temporal(v)
    if t is None or t["day"] is None:
        return None
    return "%04d-%02d-%02d" % (t["year"], t["month"], t["day"])


# ---------------------------------------------------------------- 类型推断

# 「YYYY.MM」既像日期又像小数（float("2026.08") 成功）。这类 4 位年 + 点 + 1~2 位月
# 的写法在台账里几乎不可能是金额/数量，故显式让日期优先；其余写法仍按「先数字后日期」，
# 避免把真正的数值列误判成日期。
_RE_YM_DOT = re.compile(r"^\d{4}\.\d{1,2}$")


def infer_type(values):
    """对一列的非空值采样，推断类型: number / date / text"""
    nums = dates = total = 0
    for v in values[:TYPE_SAMPLE]:
        total += 1
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            nums += 1
        elif isinstance(v, (datetime, date)):
            dates += 1
        else:
            s = str(v).strip()
            if _RE_YM_DOT.match(s):
                dates += 1  # 年月点号写法优先判为日期
                continue
            try:
                float(s.replace(",", ""))
                nums += 1
            except ValueError:
                if parse_temporal(s) is not None:
                    dates += 1
    if total == 0:
        return "empty"
    if nums / total >= 0.8:
        return "number"
    if dates / total >= 0.8:
        return "date"
    return "text"


def to_number(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    return float(str(v).strip().replace(",", ""))


def norm_cell(v):
    """单元格归一化为可 JSON 化的值"""
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v == int(v):
        return int(v)
    return v


# ---------------------------------------------------------------- 读取表格

def read_csv_rows(path):
    """读 CSV，自动识别编码与分隔符，返回 (headers, rows[list[dict]], None)"""
    raw = None
    for enc in ("utf-8-sig", "gb18030", "utf-8"):
        try:
            with open(path, "r", encoding=enc) as f:
                raw = f.read()
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    if raw is None:
        fail("无法识别 CSV 文件编码", "请确认文件是文本 CSV 而非 Excel 二进制格式")
    try:
        dialect = csv.Sniffer().sniff(raw[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    reader = csv.DictReader(raw.splitlines(), dialect=dialect)
    headers = [h.strip() for h in (reader.fieldnames or [])]
    rows = []
    for r in reader:
        rows.append({(k or f"列{i+1}").strip(): v for i, (k, v) in enumerate(r.items())})
    return headers, rows, None


def read_xlsx_rows(path, sheet=None):
    """读 xlsx，返回 (headers, rows, sheet_names)；依赖 openpyxl（可选增强）"""
    try:
        from openpyxl import load_workbook
    except ImportError:
        fail("当前环境缺少 openpyxl，无法读取 xlsx",
             "安装：pip install openpyxl；或改用平台内置表格读取能力，"
             "或请用户提供 CSV 版本的同一张表（CSV 路径纯标准库，无需依赖）")
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        fail(f"无法打开 xlsx 文件: {e}", "确认文件路径正确、文件未损坏、不是加密文件")
    sheets = wb.sheetnames
    ws = wb[sheet] if sheet in (sheets or []) else wb.active
    rows_raw = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows_raw:
        fail(f"工作表「{ws.title}」是空的", "用 --sheet 指定其他工作表")
    headers = [str(h).strip() if h is not None else f"列{i+1}"
               for i, h in enumerate(rows_raw[0])]
    rows = []
    for r in rows_raw[1:]:
        if all(v is None or str(v).strip() == "" for v in r):
            continue  # 跳过全空行
        rows.append({headers[i]: norm_cell(v) for i, v in enumerate(r) if i < len(headers)})
    return headers, rows, sheets


def column_values(rows, col):
    return [r[col] for r in rows if r.get(col) is not None and str(r.get(col)).strip() != ""]


# ---------------------------------------------------------------- 默认模式

def summarize(rows, headers, sample_n, sheet_names, active_sheet):
    """默认模式：列概况 + 样例行"""
    cols = []
    for h in headers:
        vals = column_values(rows, h)
        cols.append({
            "name": h,
            "type": infer_type(vals),
            "non_empty": len(vals),
            "unique_count": len({str(v) for v in vals}),
        })
    sample = [{k: v for k, v in r.items()} for r in rows[:sample_n]]
    return {
        "ok": True,
        "file_mode": "summary",
        "sheets": sheet_names,
        "active_sheet": active_sheet,
        "row_count": len(rows),
        "columns": cols,
        "sample_rows": sample,
    }


# ---------------------------------------------------------------- 分组模式

def _sort_key_for_group(key):
    """月份升序用的排序键：特殊分组（空值/未识别）恒排最后，不参与时间排序"""
    if key.startswith("("):
        return (1, "", key)
    return (0, key, "")


def group_by(rows, headers, col, sample_n):
    """分组模式：按列分组统计。日期列自动按月份（YYYY-MM）聚合"""
    if col not in headers:
        fail(f"分组列「{col}」不存在",
             f"可用列: {', '.join(headers)}")
    col_types = {h: infer_type(column_values(rows, h)) for h in headers}
    numeric_cols = [h for h in headers if h != col and col_types[h] == "number"]
    date_cols = [h for h in headers if h != col and col_types[h] == "date"]

    # 分组列被推断为日期时才尝试按月聚合；识别与归一化共用 parse_temporal，
    # 因此「判为 date」与「能聚合出 YYYY-MM」不会再出现不一致。
    want_month_bucket = col_types[col] == "date"
    groups = {}
    unparsed = []          # 判为 date 却解析不出月份的个例（如实上报，不静默）
    ambiguous = []         # 月/日顺序天然歧义的个例

    for idx, r in enumerate(rows, start=1):
        v = r.get(col)
        if v is None or str(v).strip() == "":
            key = "(空值)"
        elif want_month_bucket:
            key = month_key(v)
            if key is None:
                key = "(未识别日期: %s)" % str(v).strip()
                unparsed.append({"row": idx, "value": str(v).strip()})
            else:
                t = parse_temporal(v)
                if t and t["ambiguous"]:
                    ambiguous.append({"row": idx, "value": str(v).strip(),
                                      "parsed_as": key})
        else:
            key = str(v).strip()
        groups.setdefault(key, []).append(r)

    # 只有「每一个非空值都真的聚合成了 YYYY-MM」才置 true，避免标记与行为不符
    bucketed = want_month_bucket and not unparsed

    # 每个数值列在每组内「实际参与统计的数字个数」——加权平均必须用它，不能用组行数
    numeric_counts = {}
    result = []
    if bucketed:
        # 月份聚合按时间升序（报表习惯）；(空值)/(未识别日期: ...) 恒排最后
        ordered = sorted(groups.items(), key=lambda kv: _sort_key_for_group(kv[0]))
    else:
        ordered = sorted(groups.items(), key=lambda kv: -len(kv[1]))

    for key, rs in ordered:
        g = {"group": key, "count": len(rs)}
        for h in numeric_cols:
            nums = []
            for r in rs:
                try:
                    nums.append(to_number(r[h]))
                except (ValueError, TypeError):
                    pass
            if nums:
                numeric_counts.setdefault(key, {})[h] = len(nums)
                g[h] = {"sum": round(sum(nums), 2), "avg": round(sum(nums) / len(nums), 2),
                        "min": round(min(nums), 2), "max": round(max(nums), 2)}
        for h in date_cols:
            dates = sorted(d for d in (day_key(r[h]) for r in rs if r.get(h)) if d)
            if not dates:
                dates = sorted(str(r[h]) for r in rs if r.get(h))
            if dates:
                g[h] = {"earliest": dates[0], "latest": dates[-1]}
        result.append(g)

    truncated = False
    if len(result) > MAX_GROUPS:
        top = result[:TOP_GROUPS]
        other = result[TOP_GROUPS:]
        # 合并组必须带齐聚合值：否则调用方想写「其余 N 组合计 X」时只能编造。
        # avg 用 count 加权：sum(avg_i * n_i) / sum(n_i)，n_i 为该组内实际参与
        # 统计的数字个数；对各组 avg 做算术平均会算错（各组行数不等时必错）。
        merged = {
            "group": "(其他 %d 组)" % len(other),
            "count": sum(o["count"] for o in other),
            "merged": True,
            "merged_from": len(other),
            "merged_groups": [o["group"] for o in other],
        }
        for h in numeric_cols:
            stats = [(o[h]["sum"], numeric_counts.get(o["group"], {}).get(h, 0))
                     for o in other if h in o]
            if not stats:
                continue
            total_sum = sum(s for s, _ in stats)
            total_n = sum(n for _, n in stats)
            merged[h] = {
                "sum": round(total_sum, 2),
                "avg": round(total_sum / total_n, 2) if total_n else None,
                "min": round(min(o[h]["min"] for o in other if h in o), 2),
                "max": round(max(o[h]["max"] for o in other if h in o), 2),
                "avg_method": "count_weighted",
                "value_count": total_n,
            }
        for h in date_cols:
            ds = [o[h]["earliest"] for o in other if h in o]
            dl = [o[h]["latest"] for o in other if h in o]
            if ds and dl:
                merged[h] = {"earliest": min(ds), "latest": max(dl)}
        top.append(merged)
        result = top
        truncated = True

    payload = {
        "ok": True,
        "file_mode": "group",
        "group_column": col,
        "group_column_type": col_types[col],
        "bucketed_by_month": bucketed,
        "total_rows": len(rows),
        "group_count": len(groups),
        "truncated": truncated,
        "numeric_columns": numeric_cols,
        "date_columns": date_cols,
        "groups": result,
        "sample_rows": rows[:sample_n],
    }
    if truncated:
        payload["merge_note"] = (
            "「(其他 N 组)」是多组合并值：sum 为各组求和，min/max 取各组极值，"
            "avg 按各组数字个数加权（avg_method=count_weighted）。"
            "引用该数值时必须注明是「其余 N 组的合计/均值」，不可当作单个分组。"
        )
    if unparsed:
        payload["bucket_warning"] = (
            "分组列被推断为日期，但有 %d 个值无法归一化为 YYYY-MM，已单列到"
            "「(未识别日期: ...)」分组；因此 bucketed_by_month 置为 false。"
            % len(unparsed)
        )
        payload["bucket_unparsed"] = unparsed
    if ambiguous:
        payload["bucket_ambiguous_dates"] = (
            "以下值的「月/日」顺序从字符串无法判定（月、日均 <= 12），已按「月在前」"
            "解析，请与用户确认口径：" + "、".join(
                "第%d行 %s→%s" % (a["row"], a["value"], a["parsed_as"]) for a in ambiguous)
        )
    return payload


# ---------------------------------------------------------------- 审计模式

def _row_ref(idx):
    """行号引用：idx 为数据行序号（1 起，不含表头）"""
    return {"row": idx, "excel_row": idx + 1}


def _percentile(sorted_vals, q):
    """线性插值分位数（q ∈ [0,1]），sorted_vals 必须已升序"""
    n = len(sorted_vals)
    if n == 0:
        return None
    if n == 1:
        return float(sorted_vals[0])
    pos = (n - 1) * q
    lo = int(math.floor(pos))
    hi = int(math.ceil(pos))
    if lo == hi:
        return float(sorted_vals[lo])
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (pos - lo)


def detect_empty(rows, headers, max_items):
    """空值检测：逐行列出精确行号"""
    items = []
    for h in headers:
        empties = [i for i, r in enumerate(rows, start=1)
                   if r.get(h) is None or str(r.get(h)).strip() == ""]
        if not empties:
            continue
        items.append({
            "column": h,
            "empty_count": len(empties),
            "rows": [_row_ref(i) for i in empties[:max_items]],
            "rows_shown": min(len(empties), max_items),
            "rows_omitted": max(0, len(empties) - max_items),
        })
    return {
        "method": "空值（None 或去除首尾空白后为空字符串）",
        "confidence": "确证",
        "note": "空值是客观事实，可直接引用；统计口径须注明缺失行数。",
        "columns_with_empty": len(items),
        "items": items,
    }


def detect_duplicates(rows, headers, key_cols, max_items):
    """重复值检测：默认按全部列取值完全一致判定；--key-col 可指定关键列"""
    keys = [c for c in (key_cols or headers) if c in headers]
    missing = [c for c in (key_cols or []) if c not in headers]
    buckets = {}
    skipped = 0
    for i, r in enumerate(rows, start=1):
        vals = [r.get(c) for c in keys]
        if any(v is None or str(v).strip() == "" for v in vals):
            skipped += 1  # 关键列有空值的行不参与重复判定，避免误报
            continue
        buckets.setdefault(tuple(str(v).strip() for v in vals), []).append(i)

    dupes = []
    for vals, idxs in buckets.items():
        if len(idxs) < 2:
            continue
        dupes.append({
            "rows": [_row_ref(i) for i in idxs],
            "occurrence": len(idxs),
            "values": {c: vals[j] for j, c in enumerate(keys)},
        })
    dupes.sort(key=lambda d: d["rows"][0]["row"])
    return {
        "method": "关键列取值完全一致（去首尾空白后比较）",
        "key_columns": keys,
        "confidence": "启发式候选",
        "note": ("「完全一致」只是重复的充分线索：整行相同可能是合法的重复业务"
                 "（如每月固定一条），需与用户确认后再判定为重复计数。"),
        "rows_skipped_incomplete": skipped,
        "duplicate_group_count": len(dupes),
        "duplicate_row_count": sum(d["occurrence"] for d in dupes),
        "items": dupes[:max_items],
        "items_shown": min(len(dupes), max_items),
        "items_omitted": max(0, len(dupes) - max_items),
        "missing_key_columns": missing,
    }


def detect_outliers(rows, headers, col_types, method, max_items):
    """离群值检测：IQR 法 / z-score 法 / both。输出候选 + 方法名 + ratio"""
    numeric_cols = [h for h in headers if col_types.get(h) == "number"]
    columns = []
    skipped = []
    for h in numeric_cols:
        pairs = []
        for i, r in enumerate(rows, start=1):
            try:
                pairs.append((i, to_number(r[h])))
            except (ValueError, TypeError):
                pass
        if len(pairs) < 4:
            skipped.append({"column": h, "reason": "有效数字不足 4 个，样本过小不判定"})
            continue
        vals = sorted(v for _, v in pairs)
        n = len(vals)
        found = []

        if method in ("iqr", "both"):
            q1 = _percentile(vals, 0.25)
            q3 = _percentile(vals, 0.75)
            iqr = q3 - q1
            if iqr <= 0:
                skipped.append({"column": h, "reason": "IQR 为 0（数值过于集中），IQR 法不适用"})
            else:
                lo_b, hi_b = q1 - IQR_K * iqr, q3 + IQR_K * iqr
                for i, v in pairs:
                    if v > hi_b:
                        ratio = (v - hi_b) / iqr
                        found.append({"row": _row_ref(i), "value": v, "method": "IQR",
                                      "ratio": round(ratio, 3), "side": "high",
                                      "bound": round(hi_b, 2),
                                      "ratio_definition": "(值 - 上界) / IQR"})
                    elif v < lo_b:
                        ratio = (lo_b - v) / iqr
                        found.append({"row": _row_ref(i), "value": v, "method": "IQR",
                                      "ratio": round(ratio, 3), "side": "low",
                                      "bound": round(lo_b, 2),
                                      "ratio_definition": "(下界 - 值) / IQR"})

        if method in ("zscore", "both"):
            mean = sum(vals) / n
            var = sum((v - mean) ** 2 for v in vals) / (n - 1)  # 样本标准差
            sd = math.sqrt(var)
            if sd <= 0:
                skipped.append({"column": h, "reason": "标准差为 0，z-score 法不适用"})
            else:
                for i, v in pairs:
                    z = (v - mean) / sd
                    if abs(z) > ZSCORE_THRESHOLD:
                        found.append({"row": _row_ref(i), "value": v, "method": "z-score",
                                      "ratio": round(abs(z), 3),
                                      "threshold": ZSCORE_THRESHOLD,
                                      "ratio_definition": "|z| = |(值 - 均值) / 样本标准差|"})

        found.sort(key=lambda d: (d["row"]["row"], d["method"]))
        if found:
            columns.append({
                "column": h,
                "mean": round(sum(vals) / n, 2),
                "count": n,
                "methods": sorted({d["method"] for d in found}),
                "outlier_count": len(found),
                "items": found[:max_items],
                "items_shown": min(len(found), max_items),
                "items_omitted": max(0, len(found) - max_items),
            })
    return {
        "method": ("IQR 法（Q1-%s*IQR / Q3+%s*IQR，ratio 为超出上/下界的距离÷IQR）"
                   % (IQR_K, IQR_K)) if method in ("iqr", "both")
                  else ("z-score 法（|z| > %s，ratio 为 |z|）" % ZSCORE_THRESHOLD),
        "methods_param": method,
        "confidence": "启发式候选",
        "note": ("离群 ≠ 错误：可能是录入错误，也可能是真实的大额/极端值。"
                 "ratio 越大偏离越远，仅供排序核对，不得直接当作错误数据。"),
        "columns_with_outliers": len(columns),
        "items": columns,
        "skipped_columns": skipped,
    }


def detect_overdue(rows, headers, due_col, status_col, unfinished, as_of, max_items):
    """超期检测。语义必须由调用方显式给出，脚本不猜「什么状态算未完成」。"""
    missing = []
    if not due_col:
        missing.append("--due-col")
    if not status_col:
        missing.append("--status-col")
    if not unfinished:
        missing.append("--unfinished-values")
    unknown = [c for c in (due_col, status_col) if c and c not in headers]
    if missing or unknown:
        reasons = []
        if missing:
            reasons.append("缺少必填参数: " + "、".join(missing))
        if unknown:
            reasons.append("列不存在于表中: " + "、".join(unknown))
        return None, {
            "reason": "；".join(reasons),
            "why_skipped": (
                "「状态里不含『完成』就算未完成」这类语义无法从表格推断，"
                "不同表的「未完成」写法差异极大，猜错会直接产出错误的超期名单。"
                "请显式传入 --due-col、--status-col、--unfinished-values 后重跑。"),
            "required_args": ["--due-col", "--status-col", "--unfinished-values"],
            "missing_args": missing,
            "unknown_columns": unknown,
        }

    ref = as_of if as_of else date.today()
    ref_day = day_key(ref)
    if ref_day is None:
        return None, {
            "reason": "--as-of 不是可识别的完整日期: %s" % as_of,
            "required_args": ["--as-of YYYY-MM-DD"],
            "missing_args": ["--as-of"],
            "unknown_columns": [],
        }

    tokens = [t.strip() for t in unfinished if t and t.strip()]
    items = []
    unparsable_due = []
    month_only_due = []
    for i, r in enumerate(rows, start=1):
        raw_status = r.get(status_col)
        status_s = "" if raw_status is None else str(raw_status).strip()
        # 状态匹配：先精确相等，再包含匹配；命中哪个写法如实记录
        matched = next((t for t in tokens if status_s == t), None)
        how = "exact"
        if matched is None:
            matched = next((t for t in tokens if t and t in status_s), None)
            how = "contains"
        if matched is None:
            continue
        raw_due = r.get(due_col)
        d = day_key(raw_due)
        if d is None:
            t = parse_temporal(raw_due)
            entry = {"row": _row_ref(i), "due_raw": None if raw_due is None else str(raw_due).strip(),
                     "status_raw": status_s}
            if t is not None:
                month_only_due.append(entry)
            else:
                unparsable_due.append(entry)
            continue
        if d < ref_day:
            y1, m1, d1 = (int(x) for x in d.split("-"))
            y2, m2, d2 = (int(x) for x in ref_day.split("-"))
            overdue_days = (date(y2, m2, d2) - date(y1, m1, d1)).days
            items.append({
                "row": _row_ref(i),
                "due_date": d,
                "status_raw": status_s,
                "status_matched": matched,
                "status_match": how,
                "overdue_days": overdue_days,
                "reference_date": ref_day,
            })
    items.sort(key=lambda d: (-d["overdue_days"], d["row"]["row"]))
    return {
        "method": ("截止日期 < 参照日期，且状态命中 --unfinished-values 给出的写法"
                   "（先精确后包含）"),
        "due_column": due_col,
        "status_column": status_col,
        "unfinished_values": tokens,
        "reference_date": ref_day,
        "reference_date_source": ("--as-of 显式传入" if as_of else "系统当天日期"),
        "confidence": "确证（口径由调用方显式给定，非脚本推断）",
        "note": ("口径完全取决于 --unfinished-values 的取值；换一张表必须重新确认。"
                 "overdue_days 为自然日天数。"),
        "overdue_count": len(items),
        "items": items[:max_items],
        "items_shown": min(len(items), max_items),
        "items_omitted": max(0, len(items) - max_items),
        "unparsable_due_rows": unparsable_due[:max_items],
        "month_precision_due_rows": month_only_due[:max_items],
        "not_judged_reason": (
            "截止日期只到月（无法判具体是否过期）与无法解析的行一律不计入，"
            "明细见 unparsable_due_rows / month_precision_due_rows"),
    }, None


def detect_mixed_date_formats(rows, headers, col_types, max_items):
    """日期格式混用检测：同一列里出现多种书写方式，会让按时间统计悄悄失真"""
    items = []
    for h in headers:
        if col_types.get(h) != "date":
            continue
        shapes = {}
        for i, r in enumerate(rows, start=1):
            v = r.get(h)
            if v is None or str(v).strip() == "":
                continue
            s = str(v).strip()
            shape = _date_shape(s)
            shapes.setdefault(shape, []).append(i)
        if len(shapes) > 1:
            items.append({
                "column": h,
                "shape_count": len(shapes),
                "shapes": [{"shape": s, "count": len(idx),
                            "rows": [_row_ref(i) for i in idx[:max_items]]}
                           for s, idx in sorted(shapes.items())],
            })
    return {
        "method": "同一日期列内出现多种书写格式（分隔符/是否含日/中文写法）",
        "confidence": "确证（格式差异客观存在）",
        "note": ("混用本身不一定是错，但会让「按月/按时间排序」口径不一致，"
                 "统计前应与用户确认统一为哪种写法。"),
        "columns_with_mixed_formats": len(items),
        "items": items,
    }


def _date_shape(s):
    """把日期写法归类成形态标签，用于检出同列混用"""
    t = parse_temporal(s)
    if t is None:
        return "非日期"
    parts = []
    if "年" in s:
        parts.append("中文")
    if "/" in s:
        parts.append("斜杠")
    elif "." in s:
        parts.append("点号")
    elif "-" in s:
        parts.append("连字符")
    parts.append("含日" if t["day"] is not None else "仅年月")
    return "+".join(parts)


def audit(rows, headers, args):
    """审计模式：只报「候选信号 + 精确位置 + 方法名」，不报结论"""
    col_types = {h: infer_type(column_values(rows, h)) for h in headers}
    max_items = args.audit_max_items
    payload = {
        "ok": True,
        "file_mode": "audit",
        "row_count": len(rows),
        "row_numbering": ("行号 row 为数据行序号（1 起，不含表头）；"
                          "excel_row 为该行在表格文件中的物理行号（= row + 1）"),
        "disclaimer": ("本节所有条目均为「候选信号」而非结论：脚本只指出位置、"
                       "原始值与所用方法，是否成立由用户确认。引用前不得写成定论。"),
        "columns": [{"name": h, "type": col_types[h],
                     "non_empty": len(column_values(rows, h))} for h in headers],
        "empty_values": detect_empty(rows, headers, max_items),
        "duplicates": detect_duplicates(rows, headers, args.key_col, max_items),
        "outliers": detect_outliers(rows, headers, col_types, args.outlier_method, max_items),
        "mixed_date_formats": detect_mixed_date_formats(rows, headers, col_types, max_items),
    }
    overdue, skipped = detect_overdue(
        rows, headers, args.due_col, args.status_col,
        args.unfinished_values, args.as_of, max_items)
    if overdue is not None:
        payload["overdue"] = overdue
    else:
        payload["overdue"] = None
        payload["overdue_skipped"] = skipped
    return payload


# ---------------------------------------------------------------- 自测

def _t(rows, headers, col, sample_n=5):
    """自测辅助：以分组模式跑一段内存数据（不落盘、不依赖任何外部文件、不产生 stdout 输出）"""
    return group_by(rows, headers, col, sample_n)


def run_selftest():
    """内置自测：覆盖 4 个修复点。返回 (exit_code, report_dict)"""
    cases = []

    def case(name, fn):
        try:
            detail = fn()
            cases.append({"name": name, "passed": True, "detail": detail})
        except AssertionError as e:
            cases.append({"name": name, "passed": False, "detail": "断言失败: %s" % e})
        except Exception as e:  # noqa: BLE001 - 自测需捕获任何异常并如实上报
            cases.append({"name": name, "passed": False,
                          "detail": "异常 %s: %s" % (type(e).__name__, e)})

    # 1) 合并组必须带聚合值，且 avg 按 count 加权
    def t_merged():
        rows, headers = [], ["项目", "金额"]
        for i in range(1, 46):
            rows.append({"项目": "P%02d" % i, "金额": str(i * 10)})
        r = _t(rows, headers, "项目")
        m = r["groups"][-1]
        assert r["truncated"] is True, "45 组应触发截断"
        assert m.get("merged") is True and m["merged_from"] == 25, "合并标记缺失"
        # 被合并的是 P21..P45：sum=2310+...+450 = (21+45)*25/2*10 = 8250
        # 加权均值 = 8250 / 25 = 330；若错做各组 avg 的算术平均仍是 330（每组 1 行），
        # 故另造一个组内行数不等的数据集来真正验证加权。
        assert m["金额"]["sum"] == 8250.0, "合并 sum 错: %s" % m["金额"]["sum"]
        assert m["金额"]["min"] == 210.0 and m["金额"]["max"] == 450.0, "合并 min/max 错"
        assert m["金额"]["avg_method"] == "count_weighted", "avg 未标注加权"
        assert m["金额"]["avg"] == 330.0, "合并 avg 错: %s" % m["金额"]["avg"]

        # 组内行数不等时才暴露「算术平均 vs 加权平均」的差别：
        # 造 35 组，前 20 组各 1 行金额 100，第 21~35 组各 1 行金额 500，
        # 再让 P01 组有 9 行。改用可控构造：直接校验加权公式。
        rows2 = [{"组": "A", "额": "100"} for _ in range(9)]
        rows2 += [{"组": "B", "额": "500"}]
        for i in range(2, 36):
            rows2.append({"组": "G%02d" % i, "额": "10"})
        r2 = _t(rows2, ["组", "额"], "组")
        m2 = r2["groups"][-1]
        assert m2.get("merged") is True, "第二个用例未触发截断"
        # 排序按行数降序，A(9 行) 第一，其余 1 行的组按字典序取前 20 个合并
        merged_groups = m2["merged_groups"]
        a_in = "A" in merged_groups
        if a_in:
            n_a = 9
            n_rest = m2["count"] - n_a
            s_a = 900.0
            s_rest = m2["金额"]["sum"] - s_a
            expect = (s_a + s_rest) / m2["count"]
            assert abs(m2["金额"]["avg"] - expect) < 0.01, "加权 avg 错"
            plain = (100.0 + (s_rest / n_rest) * n_rest) / 2  # 错误做法示意
            assert m2["金额"]["avg"] != round(plain, 2) or abs(expect - plain) < 0.01
        return "合并组 sum=%s avg=%s min=%s max=%s avg_method=%s" % (
            m["金额"]["sum"], m["金额"]["avg"], m["金额"]["min"], m["金额"]["max"],
            m["金额"]["avg_method"])

    # 2) YYYY-MM 必须判为 date、按月聚合并时间升序（行数逆序时仍升序）
    def t_yyyymm():
        rows = ([{"月份": "2026-08", "额": "100"}] * 2 +
                [{"月份": "2026-09", "额": "200"}] * 5)
        r = _t(rows, ["月份", "额"], "月份")
        assert r["group_column_type"] == "date", "YYYY-MM 未判为 date: %s" % r["group_column_type"]
        assert r["bucketed_by_month"] is True, "YYYY-MM 未按月聚合"
        keys = [g["group"] for g in r["groups"]]
        assert keys == ["2026-08", "2026-09"], "顺序错（9 月行数多时应仍按时间升序）: %s" % keys
        return "type=date bucketed=true order=%s" % keys

    # 3) 2026/01/05 这类斜杠日期必须真的聚合到月，且不谎报
    def t_slash_date():
        rows = [{"日期": "2026/01/05", "额": "1"}, {"日期": "2026/02/05", "额": "2"},
                {"日期": "2026/02/06", "额": "3"}]
        r = _t(rows, ["日期", "额"], "日期")
        keys = [(g["group"], g["count"]) for g in r["groups"]]
        assert r["bucketed_by_month"] is True, "斜杠日期未聚合却/却谎报"
        assert keys == [("2026-01", 1), ("2026-02", 2)], "斜杠日期未归一到月: %s" % keys
        # 反向：无法识别为日期时必须 false
        rows2 = [{"日期": "待定", "额": "1"}, {"日期": "待定", "额": "2"}]
        r2 = _t(rows2, ["日期", "额"], "日期")
        assert r2["bucketed_by_month"] is False, "非日期列谎报 bucketed_by_month"
        return "斜杠聚合=%s；非日期 bucketed=false" % (keys,)

    # 4) 中文年月 / 其他月份格式
    def t_cn_month():
        rows = ([{"月份": "2026年8月", "额": "1"}] * 3 +
                [{"月份": "2026年9月", "额": "2"}] * 1)
        r = _t(rows, ["月份", "额"], "月份")
        assert r["group_column_type"] == "date", "中文年月未判为 date"
        keys = [g["group"] for g in r["groups"]]
        assert keys == ["2026-08", "2026-09"], "中文年月归一错: %s" % keys
        for fmt, sample in (("YYYY/MM", "2026/08"), ("YYYY.MM", "2026.08"),
                            ("YYYY年M月D日", "2026年8月5日"),
                            ("YYYY-MM-DD", "2026-08-05")):
            rr = _t([{"月份": sample, "额": "1"}], ["月份", "额"], "月份")
            assert rr["group_column_type"] == "date", "%s 未判为 date" % fmt
            assert rr["groups"][0]["group"] == "2026-08", "%s 归一错: %s" % (
                fmt, rr["groups"][0]["group"])
        return "中文年月 order=%s；四种其他格式均归一到 2026-08" % keys

    # 5) audit：超期必须要求显式参数，不给就跳过并说明
    def t_audit_gate():
        rows = [{"截止": "2026-01-05", "状态": "进行中", "额": "1"},
                {"截止": "2026-03-05", "状态": "已完成", "额": "2"},
                {"截止": "2026-02-05", "状态": "延期", "额": "3"}]
        headers = ["截止", "状态", "额"]
        a = audit(rows, headers, _AuditArgs())
        assert a["overdue"] is None, "未给参数却输出了超期结论"
        assert "--due-col" in a["overdue_skipped"]["missing_args"], "未说明缺哪个参数"
        b = audit(rows, headers, _AuditArgs(due_col="截止", status_col="状态",
                                            unfinished_values=["进行中", "延期"],
                                            as_of="2026-03-01"))
        assert b["overdue"] is not None, "给了参数却没输出超期"
        got = [(i["due_date"], i["overdue_days"], i["row"]["row"]) for i in b["overdue"]["items"]]
        assert got == [("2026-01-05", 55, 1), ("2026-02-05", 24, 3)], "超期明细错: %s" % got
        # 重复值 / 空值 / 离群 三个检测项必须存在且给出方法名
        for k in ("empty_values", "duplicates", "outliers"):
            assert k in b and b[k].get("method"), "%s 缺方法名" % k
        assert b["outliers"]["confidence"] == "启发式候选", "离群未标注为启发式"
        return "超期明细=%s；三项检测均带方法名" % (got,)

    # 6) 离群：IQR 与 z-score 都要带 method 与 ratio
    def t_outlier():
        vals = [10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 1000]
        rows = [{"额": str(v)} for v in vals]
        r = audit(rows, ["额"], _AuditArgs(outlier_method="both"))
        cols = r["outliers"]["items"]
        assert cols, "未检出离群值"
        item = cols[0]["items"][-1]
        assert item["value"] == 1000 and item["method"] in ("IQR", "z-score"), "离群项缺方法名"
        assert isinstance(item["ratio"], (int, float)), "离群项缺 ratio"
        return "离群值=%s method=%s ratio=%s" % (item["value"], item["method"], item["ratio"])

    case("合并组携带聚合值且 avg 按 count 加权", t_merged)
    case("YYYY-MM 判为 date 并按时间升序（行数逆序）", t_yyyymm)
    case("YYYY/MM/DD 真正聚合到月；非日期不谎报", t_slash_date)
    case("中文年月 YYYY年M月 及四种其他月份格式归一", t_cn_month)
    case("--audit 超期门禁：缺参数不输出结论", t_audit_gate)
    case("--audit 离群带 method 与 ratio", t_outlier)

    failed = [c for c in cases if not c["passed"]]
    report = {
        "ok": not failed,
        "file_mode": "selftest",
        "total": len(cases),
        "passed": len(cases) - len(failed),
        "failed": len(failed),
        "cases": cases,
    }
    return (0 if not failed else 1), report


class _AuditArgs:
    """--audit 的默认参数对象，供自测构造（与 argparse 默认值保持一致）"""

    def __init__(self, due_col=None, status_col=None, unfinished_values=None,
                 as_of=None, key_col=None, outlier_method="iqr", audit_max_items=AUDIT_MAX_ITEMS):
        self.due_col = due_col
        self.status_col = status_col
        self.unfinished_values = unfinished_values or []
        self.as_of = as_of
        self.key_col = key_col or []
        self.outlier_method = outlier_method
        self.audit_max_items = audit_max_items


# ---------------------------------------------------------------- 主流程

def main():
    p = argparse.ArgumentParser(description="读取表格输出结构化 JSON")
    p.add_argument("path", nargs="?", help="表格文件路径 (.xlsx/.csv)")
    p.add_argument("--sheet", default=None, help="xlsx 工作表名（默认取活动表）")
    p.add_argument("--rows", type=int, default=5, help="样例行数（默认 5）")
    p.add_argument("--group", default=None, help="按此列名分组统计")
    p.add_argument("--audit", action="store_true",
                   help="异常候选信号模式：重复值/空值/离群值/超期，逐条给精确位置")
    p.add_argument("--key-col", action="append", default=[],
                   help="重复检测的关键列，可重复传入；默认按全部列比对")
    p.add_argument("--outlier-method", default="iqr", choices=("iqr", "zscore", "both"),
                   help="离群检测方法（默认 iqr）")
    p.add_argument("--due-col", default=None, help="[--audit] 截止日期列名")
    p.add_argument("--status-col", default=None, help="[--audit] 状态列名")
    p.add_argument("--unfinished-values", default=None,
                   help="[--audit] 视为「未完成」的状态取值，逗号分隔。"
                        "必须显式给出，脚本不猜状态语义")
    p.add_argument("--as-of", default=None,
                   help="[--audit] 超期判定的参照日期 YYYY-MM-DD（默认系统当天）")
    p.add_argument("--audit-max-items", type=int, default=AUDIT_MAX_ITEMS,
                   help="[--audit] 每项最多列出的明细条数（总数照实报告，默认 %d）"
                        % AUDIT_MAX_ITEMS)
    p.add_argument("--selftest", action="store_true",
                   help="运行内置自测（不依赖任何外部文件），退出码 0 = 通过")
    args = p.parse_args()

    if args.selftest:
        code, report = run_selftest()
        out(report)
        sys.exit(code)

    if not args.path:
        fail("缺少表格文件路径", "用法: python read_table.py <表格路径>，或加 --selftest 跑内置自测")

    path = args.path
    try:
        if path.lower().endswith(".csv"):
            headers, rows, sheets = read_csv_rows(path)
            active = None
        elif path.lower().endswith((".xlsx", ".xlsm")):
            headers, rows, sheets = read_xlsx_rows(path, args.sheet)
            active = args.sheet or (sheets[0] if sheets else None)
        elif path.lower().endswith(".xls"):
            fail("暂不支持旧版 .xls 格式", "请用 Excel/WPS 另存为 .xlsx，或导出 CSV")
        else:
            fail("不支持的文件类型", "支持 .xlsx / .xlsm / .csv")
    except SystemExit:
        raise
    except Exception as e:
        fail(f"读取失败: {e}", "确认路径正确后重试")

    if not headers or (all(h.startswith("列") for h in headers) and not rows):
        fail("表头为空或表格无数据", "确认第一行是列名")

    if args.audit:
        args.unfinished_values = _split_values(args.unfinished_values)
        out(audit(rows, headers, args))
    elif args.group:
        out(group_by(rows, headers, args.group, args.rows))
    else:
        out(summarize(rows, headers, args.rows, sheets, active))


def _split_values(raw):
    """把 --unfinished-values 拆成列表，兼容中英文逗号"""
    if not raw:
        return []
    return [t.strip() for t in re.split(r"[,，]", raw) if t.strip()]


if __name__ == "__main__":
    main()
