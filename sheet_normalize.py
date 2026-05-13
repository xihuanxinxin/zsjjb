# -*- coding: utf-8 -*-
"""
将「从 Excel 拆出的、含多段表头/合计行/列名变体」的 sheet 清洗为与资产管理系统一致的标准列，
并按 sheet（及可选的数据块）汇总「本月核算金额」「合计费用」。

Excel 在 GUI/流程中会先写入项目下 excel_normalize_cache/<子目录>/ 各 sheet 的原始网格 CSV，
再经 read_csv 解析，与用户另存的 .csv 走同一读取路径；计费仍在 gui.calculate_all，不在此模块重算。
"""
from __future__ import annotations

import argparse
import json
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from data_manager import (
    BILLING_DAYS_COLUMN,
    parse_money_decimal,
    round_money_half_up,
    sum_parsed_money,
)

# Excel 先落盘为「原始网格 CSV」再读入，与用户另存为 CSV 后走同一套 read_csv 逻辑一致
_MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_GRID_CACHE_ROOT = os.path.join(_MODULE_DIR, "excel_normalize_cache")

# 与 data_manager / gui 中「数据库资产」列顺序一致
STANDARD_COLUMNS: List[str] = [
    "唯一编号",
    "服务目录编号",
    "服务类型",
    "服务名称",
    "数量",
    "单位",
    "单价",
    "备注",
    "开通时间",
    "关停时间",
    "计费开始时间",
    "本期计费开始时间",
    "本期计费截至日期",
    "本月核算天数",
    "使用天数",
    "本月核算金额",
    BILLING_DAYS_COLUMN,
    "合计费用",
]

# 表头别名 -> 标准列名（源列名经 strip 后匹配）
HEADER_ALIASES: Dict[str, str] = {
    "内部IP": "内部IP",  # 合并进备注
    "使用数量": "数量",  # 总账统计表等用「使用数量」，解析时并入「数量」
    "征用数量": "数量",
    "服务类别": "服务类型",  # 总览统计表等用「服务类别」
    "开通日期": "开通时间",  # 部分子表不写「开通时间」而写「开通日期」
    "本期计费开始日期": "本期计费开始时间",
    "本期计费截止日期": "本期计费截至日期",
    "本期计费结束日期": "本期计费截至日期",
    "计费截至日期": "本期计费截至日期",
    "本期合计费用": "合计费用",  # 非急救转运等子表用「本期合计费用」列名
    # 「使用天数」等不合并进表头索引键名，由 parse_data_row 在「本月核算天数」为空时再取，避免两列争一个键丢列
}

INVALID_SERVICE_TYPE = frozenset(
    {"开通用户数", "关闭用户数", "合计数量", "总计金额", "合计金额"}
)


def _header_looks_like_detail_block(header_idx: Dict[str, int]) -> bool:
    """是否按「明细块」解析。无「开通时间」时多为汇总表；但部分子表不写开通列、仍带单价与目录/名称，须继续解析。"""
    if "开通时间" in header_idx:
        return True
    return (
        "单价" in header_idx
        and "服务名称" in header_idx
        and "服务目录编号" in header_idx
    )


def _strip_cell(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def normalize_header_name(name: str) -> str:
    n = _strip_cell(name)
    # Excel 单元格自动换行另存为 CSV 时，表头常带 \n，与标准列名不一致
    n = re.sub(r"[\r\n]+", "", n)
    n = _strip_cell(n)
    if n in HEADER_ALIASES:
        return HEADER_ALIASES[n]
    return n


def parse_money(val: Any) -> float:
    """单格金额：解析后四舍五入到分。多格求和请用 sum_parsed_money，避免 float 累加误差。"""
    return round_money_half_up(parse_money_decimal(val))


def parse_int_safe(val: Any) -> int:
    s = _strip_cell(val).replace(",", "")
    if not s:
        return 0
    try:
        return int(float(s))
    except ValueError:
        return 0


def _row_as_list(row: pd.Series, n_cols: int) -> List[str]:
    raw = [_strip_cell(row.iloc[j]) if j < len(row) else "" for j in range(max(len(row), n_cols))]
    return raw[: max(len(row), n_cols)] if raw else [""] * n_cols


def is_empty_row(vals: List[str]) -> bool:
    return all(not v for v in vals)


def is_header_row(vals: List[str]) -> bool:
    return bool(vals) and normalize_header_name(vals[0]) == "服务目录编号"


def is_title_row(vals: List[str]) -> bool:
    if not vals or not vals[0]:
        return False
    if vals[0] and "费用总合计" in vals[0]:
        return False
    if is_header_row(vals):
        return False
    if vals[0] in ("合计数量", "总计金额"):
        return False
    if vals[0].isdigit() and len(vals) > 1 and vals[1] not in INVALID_SERVICE_TYPE:
        return False
    if vals[0].isdigit() and len(vals) > 1 and vals[1] in INVALID_SERVICE_TYPE:
        return False
    # 标题行：首格有字，且通常第二格为空或首格含「统计」等
    if "统计" in vals[0] or "单项" in vals[0]:
        return True
    if len(vals) > 1 and not vals[1] and not vals[0].isdigit():
        return True
    return False


def is_summary_or_junk_row(vals: List[str]) -> bool:
    if not vals:
        return True
    a, b = vals[0], vals[1] if len(vals) > 1 else ""
    if a in ("合计数量", "总计金额"):
        return True
    if b in INVALID_SERVICE_TYPE:
        return True
    if "合计金额" in a or "合计数量" in a:
        return True
    joined = ",".join(vals)
    if "合计金额" in joined and "合计数量" in joined:
        return True
    return False


def _normalize_directory_id_token(a: str) -> str:
    """Excel/CSV 中服务目录编号可能被存成浮点，读成 '12345.0'，需与纯数字行一致。

    首页首行常见：全角数字、不间断空格等，导致解析后与资产管理表「1」对不上，回退匹配时误取 hits[0] 另一条目录的金额。
    """
    s = _strip_cell(a)
    if not s:
        return ""
    for ch in ("\u00a0", "\u2007", "\u202f", "\u3000", "\ufeff"):
        s = s.replace(ch, "")
    s = s.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    s = s.strip()
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".")[0]
    return s


def is_data_row(vals: List[str]) -> bool:
    if len(vals) < 3:
        return False
    a, b = _normalize_directory_id_token(vals[0]), vals[1]
    if not a or a in ("服务目录编号", "合计数量", "总计金额"):
        return False
    if b in INVALID_SERVICE_TYPE or "用户" in b:
        return False
    if not b:
        return False
    if a.isdigit():
        return True
    return bool(re.match(r"^\d+$", a))


def build_header_index(header_vals: List[str]) -> Dict[str, int]:
    idx: Dict[str, int] = {}
    for i, cell in enumerate(header_vals):
        key = normalize_header_name(_strip_cell(cell))
        if key and key not in idx:
            idx[key] = i
    return idx


def try_get_month_fee_summary_header_idx(header_vals: List[str]) -> Optional[Dict[str, int]]:
    """识别「本月合计费用」类汇总表表头：有目录/名称/合计金额列，且无「开通时间」（与明细块区分）。"""
    hi = build_header_index(header_vals)
    if "开通时间" in hi:
        return None
    if "服务目录编号" not in hi or "服务名称" not in hi:
        return None
    if "合计费用" not in hi and "本期合计费用" not in hi:
        return None
    return hi


def try_get_ledger_cost_table_header_idx(header_vals: List[str]) -> Optional[Dict[str, int]]:
    """首页「服务开通情况总账统计表」：有目录、名称、合计费用；无开通时间列；且与「本月核算」并列表区分（同窗格内不含本月核算金额）。"""
    hi = build_header_index(header_vals)
    if "开通时间" in hi:
        return None
    if "服务目录编号" not in hi or "服务名称" not in hi:
        return None
    if "本月核算金额" in hi:
        return None
    jc = hi.get("合计费用")
    if jc is None:
        jc = hi.get("本期合计费用")
    if jc is None:
        return None
    out = dict(hi)
    out["合计费用"] = jc
    return out


def try_get_ledger_month_table_header_idx(header_vals: List[str]) -> Optional[Dict[str, int]]:
    """首页「本月核算情况总账统计表」：有目录、名称、本月核算金额；无合计费用列（与开通总账并列区分）。"""
    hi = build_header_index(header_vals)
    if "开通时间" in hi:
        return None
    if "服务目录编号" not in hi or "服务名称" not in hi:
        return None
    if "本月核算金额" not in hi:
        return None
    if "合计费用" in hi or "本期合计费用" in hi:
        return None
    return hi


# 供应商各月表头「截至YYYY年M月D日应计费/使用天数」与程序内标准列名不同，按列序取最左匹配
_VENDOR_BILL_DAYS_HEADER_RE = re.compile(r"^截至\d{4}年\d{1,2}月\d{1,2}日应计费天数$")
_VENDOR_USE_DAYS_HEADER_RE = re.compile(r"^截至\d{4}年\d{1,2}月\d{1,2}日使用天数$")
_VENDOR_REQUISITION_DAYS_HEADER_RE = re.compile(r"^截至\d{4}年\d{1,2}月\d{1,2}日征用天数$")


def _ledger_header_has_usage_or_requisition_days_column(hi: Dict[str, int]) -> bool:
    """总览表含「截至…使用天数」或「截至…征用天数」，与仅有「单位」的「本月合计费用」小块区分。"""
    for k in hi:
        if _VENDOR_USE_DAYS_HEADER_RE.fullmatch(k) or _VENDOR_REQUISITION_DAYS_HEADER_RE.fullmatch(k):
            return True
    return False


def try_get_ledger_unified_overview_header_idx(header_vals: List[str]) -> Optional[Dict[str, int]]:
    """首页总览并列两表：两表列名一致——目录/类型/名称/数量(或使用数量)/截至…使用或征用天数/合计费用。

    第二张表最后一列在 Excel 中也叫「合计费用」，程序按标题区分后写入「本月核算金额」的跨表汇总。
    """
    hi = build_header_index(header_vals)
    if "开通时间" in hi:
        return None
    if "服务目录编号" not in hi or "服务名称" not in hi or "服务类型" not in hi:
        return None
    if "数量" not in hi:
        return None
    jc = hi.get("合计费用")
    if jc is None:
        jc = hi.get("本期合计费用")
    if jc is None:
        return None
    if not _ledger_header_has_usage_or_requisition_days_column(hi):
        return None
    out = dict(hi)
    out["合计费用"] = jc
    return out


def _cell_fixed_headers_else_vendor_pattern(
    header_idx: Dict[str, int],
    row_vals: List[str],
    fixed_names: Tuple[str, ...],
    vendor_pat: re.Pattern,
) -> str:
    for n in fixed_names:
        if n in header_idx:
            j = header_idx[n]
            if j < len(row_vals):
                return _strip_cell(row_vals[j])
            return ""
    ordered: List[Tuple[int, str]] = []
    for k, j in header_idx.items():
        if vendor_pat.fullmatch(k):
            ordered.append((j, k))
    ordered.sort(key=lambda x: x[0])
    for j, _k in ordered:
        if j < len(row_vals):
            return _strip_cell(row_vals[j])
    return ""


def parse_data_row(header_idx: Dict[str, int], row_vals: List[str]) -> Dict[str, str]:
    def g(*names: str) -> str:
        for n in names:
            if n in header_idx:
                j = header_idx[n]
                if j < len(row_vals):
                    return _strip_cell(row_vals[j])
        return ""

    internal_ip = g("内部IP")
    remark = g("备注")
    remark_merged = internal_ip
    if remark:
        remark_merged = f"{internal_ip} {remark}".strip() if internal_ip else remark

    out: Dict[str, str] = {}
    out["服务目录编号"] = g("服务目录编号")
    out["服务类型"] = g("服务类型")
    out["服务名称"] = g("服务名称")
    out["数量"] = g("数量")
    out["单位"] = g("单位")
    out["单价"] = g("单价")
    out["备注"] = remark_merged
    out["开通时间"] = g("开通时间")
    out["关停时间"] = g("关停时间")
    out["计费开始时间"] = g("计费开始时间")
    out["本期计费开始时间"] = g("本期计费开始时间", " 本期计费开始时间")
    out["本期计费截至日期"] = g("本期计费截至日期")
    # 使用类列：固定名或「截至…使用天数」「截至…征用天数」
    _usage = _cell_fixed_headers_else_vendor_pattern(
        header_idx,
        row_vals,
        ("使用天数", "使用点数", "本月使用天数", "实际使用天数"),
        _VENDOR_USE_DAYS_HEADER_RE,
    )
    if not _usage:
        _usage = _cell_fixed_headers_else_vendor_pattern(
            header_idx, row_vals, (), _VENDOR_REQUISITION_DAYS_HEADER_RE
        )
    out["使用天数"] = _usage
    # 本月核算天数：主列空时再用使用类列（与上同套别名）
    _md = g("本月核算天数")
    if not _md:
        _md = _usage
    out["本月核算天数"] = _md
    out["本月核算金额"] = g("本月核算金额")
    out[BILLING_DAYS_COLUMN] = _cell_fixed_headers_else_vendor_pattern(
        header_idx,
        row_vals,
        (BILLING_DAYS_COLUMN, "截至2026年4月30日应计费天数"),
        _VENDOR_BILL_DAYS_HEADER_RE,
    )
    out["合计费用"] = g("合计费用", "本期合计费用")
    return out


@dataclass
class BlockResult:
    title: str
    rows: List[Dict[str, str]] = field(default_factory=list)

    def totals(self) -> Tuple[float, float]:
        """对块内每行：读取该行「本月核算金额」「合计费用」单元格数值后相加（不重算公式）。"""
        m = sum_parsed_money([r.get("本月核算金额", "") for r in self.rows])
        c = sum_parsed_money([r.get("合计费用", "") for r in self.rows])
        return m, c


def extract_blocks_from_grid(df: pd.DataFrame) -> List[BlockResult]:
    """从无表头的原始网格中解析多个「表头+数据」块。"""
    blocks: List[BlockResult] = []
    n = len(df)
    max_cols = int(df.shape[1]) if df.shape[1] else 0
    i = 0
    pending_title = ""

    while i < n:
        row = df.iloc[i]
        vals = _row_as_list(row, max_cols)
        max_cols = max(max_cols, len(vals))

        if is_title_row(vals):
            pending_title = vals[0]
            i += 1
            continue

        if is_header_row(vals):
            header_vals = [_strip_cell(row.iloc[j]) if j < len(row) else "" for j in range(len(row))]
            header_idx = build_header_index(header_vals)
            i += 1
            # 无「开通时间」多为汇总表；见 _header_looks_like_detail_block
            if not _header_looks_like_detail_block(header_idx):
                while i < n:
                    row2 = df.iloc[i]
                    vals2 = _row_as_list(row2, max_cols)
                    if is_header_row(vals2) or is_title_row(vals2):
                        break
                    i += 1
                continue

            br = BlockResult(title=pending_title or "")
            pending_title = ""

            while i < n:
                row2 = df.iloc[i]
                vals2 = _row_as_list(row2, max_cols)
                if is_header_row(vals2):
                    break
                if is_title_row(vals2):
                    pending_title = vals2[0]
                    i += 1
                    break
                if is_summary_or_junk_row(vals2):
                    i += 1
                    continue
                if is_empty_row(vals2):
                    i += 1
                    continue
                if is_data_row(vals2):
                    br.rows.append(parse_data_row(header_idx, vals2))
                    i += 1
                    continue
                i += 1

            if br.rows:
                blocks.append(br)
            continue

        i += 1

    return blocks


@dataclass
class SheetMatchRecord:
    """一条可写回原表 Excel 的明细行（与 extract_blocks 识别规则一致）。"""

    excel_row_1based: int
    header_excel_row_1based: int
    header_idx: Dict[str, int]
    parsed: Dict[str, str]


def extract_sheet_match_records(df: pd.DataFrame) -> List[SheetMatchRecord]:
    """遍历与 extract_blocks_from_grid 相同的网格逻辑，输出每条明细的 Excel 行号与解析结果（用于回填）。"""
    out: List[SheetMatchRecord] = []
    n = len(df)
    max_cols = int(df.shape[1]) if df.shape[1] else 0
    i = 0
    pending_title = ""

    while i < n:
        row = df.iloc[i]
        vals = _row_as_list(row, max_cols)
        max_cols = max(max_cols, len(vals))

        if is_title_row(vals):
            pending_title = vals[0]
            i += 1
            continue

        if is_header_row(vals):
            header_vals = [_strip_cell(row.iloc[j]) if j < len(row) else "" for j in range(len(row))]
            header_idx = build_header_index(header_vals)
            header_excel_row = i + 1
            i += 1
            if not _header_looks_like_detail_block(header_idx):
                while i < n:
                    row2 = df.iloc[i]
                    vals2 = _row_as_list(row2, max_cols)
                    if is_header_row(vals2) or is_title_row(vals2):
                        break
                    i += 1
                continue

            pending_title = ""

            while i < n:
                row2 = df.iloc[i]
                vals2 = _row_as_list(row2, max_cols)
                if is_header_row(vals2):
                    break
                if is_title_row(vals2):
                    pending_title = vals2[0]
                    i += 1
                    break
                if is_summary_or_junk_row(vals2):
                    i += 1
                    continue
                if is_empty_row(vals2):
                    i += 1
                    continue
                if is_data_row(vals2):
                    out.append(
                        SheetMatchRecord(
                            excel_row_1based=i + 1,
                            header_excel_row_1based=header_excel_row,
                            header_idx=header_idx,
                            parsed=parse_data_row(header_idx, vals2),
                        )
                    )
                    i += 1
                    continue
                i += 1

            continue

        i += 1

    return out


def _read_excel_as_raw_grid(path: str, sheet_name: Optional[Any]) -> pd.DataFrame:
    """直接从 xlsx 读成无表头字符串网格（仅在不走磁盘缓存时使用）。"""
    df0 = pd.read_excel(
        path, sheet_name=sheet_name, header=None, dtype=str, keep_default_na=False
    )
    return df0.fillna("")


def new_excel_grid_cache_session(workbook_path: str) -> str:
    """在 excel_normalize_cache 下为本次处理新建子目录（按工作簿名 + 时间戳）。"""
    os.makedirs(EXCEL_GRID_CACHE_ROOT, exist_ok=True)
    wb_base = os.path.splitext(os.path.basename(workbook_path))[0]
    safe = re.sub(r'[\\/:*?"<>|]', "_", wb_base) or "workbook"
    sub = f"{safe}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    d = os.path.join(EXCEL_GRID_CACHE_ROOT, sub)
    os.makedirs(d, exist_ok=True)
    return d


def _sheet_grid_cache_filename(sheet_index: int, sheet_name: str) -> str:
    safe = re.sub(r'[\\/:*?"<>|]', "_", sheet_name) or "sheet"
    return f"{sheet_index:02d}_{safe}.csv"


def export_excel_workbook_grid_csvs(path: str, cache_dir: str) -> None:
    """把整本工作簿每个 sheet 写成无表头 UTF-8 CSV（与 process_csv_file 所读格式一致）。"""
    xls = pd.ExcelFile(path)
    for idx, sn in enumerate(xls.sheet_names):
        csv_path = os.path.join(cache_dir, _sheet_grid_cache_filename(idx, sn))
        df0 = pd.read_excel(
            path, sheet_name=sn, header=None, dtype=str, keep_default_na=False
        ).fillna("")
        df0.to_csv(csv_path, index=False, header=False, encoding="utf-8-sig")


def read_raw_grid(path: str, sheet_name: Optional[Any] = 0) -> pd.DataFrame:
    if path.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "utf-8", "gbk", "gb2312"):
            try:
                return pd.read_csv(path, header=None, encoding=enc, dtype=str, keep_default_na=False)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(path, header=None, dtype=str, keep_default_na=False)
    return _read_excel_as_raw_grid(path, sheet_name)


def assign_unique_ids(all_rows: List[Dict[str, str]], prefix: str = "CL") -> None:
    for k, r in enumerate(all_rows, start=1):
        r["唯一编号"] = f"{prefix}{str(k).zfill(5)}"


def blocks_to_standard_dataframe(blocks: List[BlockResult]) -> pd.DataFrame:
    flat: List[Dict[str, str]] = []
    for b in blocks:
        for r in b.rows:
            row = {c: r.get(c, "") for c in STANDARD_COLUMNS if c != "唯一编号"}
            row["数据块"] = b.title
            flat.append(row)
    assign_unique_ids(flat)
    cols = ["唯一编号", "数据块"] + [c for c in STANDARD_COLUMNS if c != "唯一编号"]
    return pd.DataFrame(flat, columns=cols)


@dataclass
class SheetSummary:
    sheet_name: str
    row_count: int
    total_month_amount: float
    total_cost: float
    blocks: List[Dict[str, Any]] = field(default_factory=list)
    # 由 GUI 在写出标准明细后填入，便于界面列出对应 CSV
    detail_csv_path: str = ""


def summarize_blocks(blocks: List[BlockResult]) -> Tuple[int, float, float, List[Dict[str, Any]]]:
    detail = []
    ms: List[float] = []
    cs: List[float] = []
    rc = 0
    for b in blocks:
        m, c = b.totals()
        ms.append(m)
        cs.append(c)
        rc += len(b.rows)
        detail.append(
            {
                "数据块标题": b.title or "(未命名块)",
                "行数": len(b.rows),
                "本月核算金额合计": m,
                "合计费用合计": c,
            }
        )
    return rc, sum_parsed_money(ms), sum_parsed_money(cs), detail


def process_one_sheet_raw(df: pd.DataFrame, sheet_name: str) -> Tuple[pd.DataFrame, SheetSummary]:
    blocks = extract_blocks_from_grid(df.fillna(""))
    clean_df = blocks_to_standard_dataframe(blocks)
    rc, ta, tc, detail = summarize_blocks(blocks)
    return clean_df, SheetSummary(
        sheet_name=sheet_name,
        row_count=rc,
        total_month_amount=ta,
        total_cost=tc,
        blocks=detail,
    )


def process_csv_file(
    path: str, sheet_label: Optional[str] = None
) -> Tuple[pd.DataFrame, SheetSummary]:
    raw = read_raw_grid(path, sheet_name=0)
    base = (
        sheet_label
        if sheet_label is not None
        else os.path.splitext(os.path.basename(path))[0]
    )
    return process_one_sheet_raw(raw, base)


def process_excel_workbook(
    path: str,
    sheet_names: Optional[List[str]] = None,
    grid_cache_dir: Optional[str] = None,
    skip_export: bool = False,
) -> List[Tuple[pd.DataFrame, SheetSummary]]:
    """先写原始网格 CSV 到磁盘再解析，与用户「另存 CSV 再上传」同一 read_csv 路径。

    grid_cache_dir 由调用方提供且 skip_export=True 时，假定该目录下已有 export_excel_workbook_grid_csvs 结果。
    """
    if not skip_export:
        if grid_cache_dir is None:
            grid_cache_dir = new_excel_grid_cache_session(path)
        export_excel_workbook_grid_csvs(path, grid_cache_dir)
    elif grid_cache_dir is None:
        raise ValueError("skip_export=True 时必须提供 grid_cache_dir")

    xls = pd.ExcelFile(path)
    if sheet_names is not None:
        ordered_sn = [sn for sn in sheet_names if sn in xls.sheet_names]
    else:
        ordered_sn = list(xls.sheet_names)

    out: List[Tuple[pd.DataFrame, SheetSummary]] = []
    for sn in ordered_sn:
        idx = xls.sheet_names.index(sn)
        csv_path = os.path.join(grid_cache_dir, _sheet_grid_cache_filename(idx, sn))
        raw = read_raw_grid(csv_path, sheet_name=0)
        clean_df, summary = process_one_sheet_raw(raw.fillna(""), sn)
        summary.detail_csv_path = os.path.abspath(csv_path)
        out.append((clean_df, summary))
    return out


def save_clean_csv(df: pd.DataFrame, out_path: str) -> None:
    df.to_csv(out_path, index=False, encoding="utf-8-sig")


def format_summary_lines(summaries: List[SheetSummary]) -> str:
    lines: List[str] = []
    for s in summaries:
        lines.append(f"【{s.sheet_name}】")
        lines.append(f"  明细行数: {s.row_count}")
        lines.append(f"  本月核算金额总计: ￥{s.total_month_amount:,.2f}")
        lines.append(f"  合计费用总计: ￥{s.total_cost:,.2f}")
        for b in s.blocks:
            lines.append(
                f"    · {b['数据块标题']}: 行{b['行数']}, "
                f"本月核算 {b['本月核算金额合计']}, 合计费用 {b['合计费用合计']}"
            )
        lines.append("")
    return "\n".join(lines).strip()


def run_cli():
    p = argparse.ArgumentParser(description="清洗 sheet 并汇总金额")
    p.add_argument("input", help="输入 .csv 或 .xlsx")
    p.add_argument("--out-dir", default=".", help="输出目录")
    args = p.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)
    inp = args.input

    if inp.lower().endswith(".csv"):
        clean_df, summary = process_csv_file(inp)
        base = os.path.splitext(os.path.basename(inp))[0]
        csv_path = os.path.join(args.out_dir, f"{base}_标准明细.csv")
        save_clean_csv(clean_df, csv_path)
        summ_path = os.path.join(args.out_dir, f"{base}_汇总.json")
        with open(summ_path, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "sheet": summary.sheet_name,
                    "row_count": summary.row_count,
                    "total_month_amount": summary.total_month_amount,
                    "total_cost": summary.total_cost,
                    "blocks": summary.blocks,
                },
                f,
                ensure_ascii=False,
                indent=2,
            )
        txt_path = os.path.join(args.out_dir, f"{base}_汇总.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(format_summary_lines([summary]))
        print(csv_path)
        print(summ_path)
        print(txt_path)
        print(format_summary_lines([summary]))
    else:
        pairs = process_excel_workbook(inp)
        base = os.path.splitext(os.path.basename(inp))[0]
        summaries: List[SheetSummary] = []
        for clean_df, summary in pairs:
            summaries.append(summary)
            safe = re.sub(r'[\\/:*?"<>|]', "_", summary.sheet_name) or "sheet"
            csv_path = os.path.join(args.out_dir, f"{base}_{safe}_标准明细.csv")
            save_clean_csv(clean_df, csv_path)
            print(csv_path)
        summ_path = os.path.join(args.out_dir, f"{base}_各sheet汇总.json")
        with open(summ_path, "w", encoding="utf-8") as f:
            json.dump(
                [
                    {
                        "sheet": s.sheet_name,
                        "row_count": s.row_count,
                        "total_month_amount": s.total_month_amount,
                        "total_cost": s.total_cost,
                        "blocks": s.blocks,
                    }
                    for s in summaries
                ],
                f,
                ensure_ascii=False,
                indent=2,
            )
        txt_path = os.path.join(args.out_dir, f"{base}_各sheet汇总.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(format_summary_lines(summaries))
        print(summ_path)
        print(txt_path)
        print(format_summary_lines(summaries))


if __name__ == "__main__":
    run_cli()
