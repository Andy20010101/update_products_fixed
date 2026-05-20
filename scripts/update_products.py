#!/usr/bin/env python3
from __future__ import annotations

r"""
Products 数据更新脚本
根据 Products 后台导出数据，更新「产品数据记录分析优化表」

数据源: Products-YYYY-MM-DD.xls (后台导出，标题行在第5行)
目标表: 一：产品数据记录分析优化表.xlsx -> Sheet「良友效果好的产品变化」
匹配键: 产品ID

策略: ID匹配成功->覆盖映射字段; 匹配失败->追加新行

用法:
  python update_products.py
  python update_products.py --source "D:\Downloads\Products-2026-05-02.xls"
  python update_products.py --source "D:\Downloads\Products-2026-05-02.xls" --target "D:\Downloads\...分析表.xlsx" --dry-run
"""

import argparse
import os
import re
import sys
import glob as glob_mod
from datetime import datetime
from typing import Optional

import xlrd
import openpyxl
from openpyxl.styles import PatternFill

RED_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

# ============================================================
# 配置 —— 按实际环境修改
# ============================================================
SOURCE_DIR = r"G:\阿里巴巴数据管家\阿里巴巴数据分析汇总\业务考核数据源和考核统计\产品数据表"
TARGET_DIR = r"\\192.168.1.10\zjh\社媒表汇总\社媒重要数据报表"
TARGET_FILENAME = "一：产品数据记录分析优化表_工作中.xlsx"
TARGET_SHEET = "良友效果好的产品变化"
SOURCE_HEADER_ROW = 6          # Products 表标题行在 Excel 中的行号 (1-based)
TARGET_HEADER_ROW = 2          # 分析表标题行在 Excel 中的行号 (1-based)
TARGET_DATA_START_ROW = 3      # 分析表数据起始行号 (1-based)

# Products 表列索引 (0-based, 对应第6行标题)
SRC = {
    "产品ID": 0,
    "产品名称": 1,
    "是否橱窗": 2,
    "是否顶展": 3,
    "是否P4P": 4,
    "搜索曝光次数": 5,
    "搜索点击次数": 6,
    "搜索点击率": 7,
    "访问人数": 8,
    "询盘个数": 9,
    "询盘率": 12,
    "TM咨询人数": 17,
}

# 分析表列索引 (0-based → Excel 列号 1-based)
TGT = {
    "ID": 3,                     # D列
    "曝光数": 5,                 # F列
    "点击数": 6,                 # G列
    "点击率": 7,                 # H列
    "访问人数": 9,               # J列
    "询盘数": 10,                # K列
    "询盘率": 11,                # L列
    "TM数": 13,                  # N列
    "是否橱窗": 15,              # P列
    "是否顶展": 16,              # Q列
    "是否P4P": 17,               # R列
    "2026反馈数量合计": 18,      # S列
}

# 受保护字段 (列索引 0-based)，人工维护，永不覆盖
PROTECTED_COLS = {0, 1, 2, 4, 8, 12, 14, 19, 20}

# 字段映射: 源列 → 目标列
FIELD_MAP = [
    (SRC["搜索曝光次数"], TGT["曝光数"]),
    (SRC["搜索点击次数"], TGT["点击数"]),
    (SRC["搜索点击率"], TGT["点击率"]),
    (SRC["访问人数"], TGT["访问人数"]),
    (SRC["询盘个数"], TGT["询盘数"]),
    (SRC["询盘率"], TGT["询盘率"]),
    (SRC["TM咨询人数"], TGT["TM数"]),
    (SRC["是否橱窗"], TGT["是否橱窗"]),
    (SRC["是否顶展"], TGT["是否顶展"]),
    (SRC["是否P4P"], TGT["是否P4P"]),
]

# 值格式转换器：确保源数据格式与目标表一致
def _convert_pct(val):
    """百分比字符串转小数: '1.15%' -> 0.0115"""
    if isinstance(val, str) and val.strip().endswith('%'):
        try:
            return float(val.strip().replace('%', '')) / 100.0
        except (ValueError, TypeError):
            return val
    return val


def _normalize_yn(val):
    """Y/N 统一为小写"""
    if isinstance(val, str) and val.strip().upper() in ('Y', 'N'):
        return val.strip().lower()
    return val


VALUE_CONVERTERS = {
    TGT["点击率"]: _convert_pct,
    TGT["询盘率"]: _convert_pct,
    TGT["是否橱窗"]: _normalize_yn,
    TGT["是否顶展"]: _normalize_yn,
    TGT["是否P4P"]: _normalize_yn,
}

# 标红的列：只有值真正变动时才标记
RED_HIGHLIGHT_COLS = {
    TGT["曝光数"],       # F列
    TGT["点击数"],       # G列
    TGT["访问人数"],     # J列
    TGT["询盘数"],       # K列
    TGT["TM数"],         # N列
}

# 新产品追加时写入的列 (0-based) 及获取值的函数
NEW_PRODUCT_COLS = [
    TGT["ID"],
    TGT["曝光数"],
    TGT["点击数"],
    TGT["点击率"],
    TGT["访问人数"],
    TGT["询盘数"],
    TGT["询盘率"],
    TGT["TM数"],
    TGT["是否橱窗"],
    TGT["是否顶展"],
    TGT["是否P4P"],
    TGT["2026反馈数量合计"],
]

# 启动时校验：所有写入列不能落在保护列中
_writable_cols = {tgt for _, tgt in FIELD_MAP} | {TGT["2026反馈数量合计"]} | set(NEW_PRODUCT_COLS)
_overlap = _writable_cols & PROTECTED_COLS
if _overlap:
    raise SystemExit(
        f"配置错误: 以下列同时在写入列和保护列中 (0-based): {_overlap}"
    )


def find_latest_products_file(source_dir: str) -> Optional[str]:
    """在 source_dir 中找到最新的 Products-*.xls 文件。

    优先按文件名中的日期匹配，其次按文件修改时间。
    """
    patterns = ["Products-*.xls", "Products-*.xlsx", "products-*.xls", "products-*.xlsx"]
    candidates = []
    for pat in patterns:
        candidates.extend(glob_mod.glob(os.path.join(source_dir, pat)))

    if not candidates:
        return None

    # 按文件名日期排序（降序）
    dated = []
    undated = []
    date_re = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
    for f in candidates:
        m = date_re.search(os.path.basename(f))
        if m:
            dated.append((f, f"{m.group(1)}{m.group(2)}{m.group(3)}"))
        else:
            undated.append(f)

    if dated:
        dated.sort(key=lambda x: x[1], reverse=True)
        return dated[0][0]
    elif undated:
        undated.sort(key=lambda f: os.path.getmtime(f), reverse=True)
        return undated[0]
    return None


def read_products_source(filepath: str) -> list[dict]:
    """读取 Products 导出表，返回数据行列表。

    每行为 dict: {列索引: 值}，保留原始类型。
    先尝试按 SOURCE_HEADER_ROW 定位标题行，若失败则自动扫描包含「产品ID」的行。
    """
    wb = xlrd.open_workbook(filepath)
    sheet = wb.sheet_by_index(0)

    # 自动定位标题行
    header_row_idx = None
    # 1) 先试配置值
    config_idx = SOURCE_HEADER_ROW - 1
    if config_idx < sheet.nrows:
        cell_val = str(sheet.cell(config_idx, 0).value).strip()
        if "产品ID" in cell_val:
            header_row_idx = config_idx
    # 2) 否则扫描前20行
    if header_row_idx is None:
        for rx in range(min(20, sheet.nrows)):
            cell_val = str(sheet.cell(rx, 0).value).strip()
            if "产品ID" in cell_val:
                header_row_idx = rx
                print(f"[检测] 标题行自动定位: Excel 第{rx + 1}行")
                break
    if header_row_idx is None:
        print("[错误] 无法定位 Products 表标题行 (未找到「产品ID」列)")
        sys.exit(1)

    rows = []
    for rx in range(header_row_idx + 1, sheet.nrows):
        row = {}
        for cx in range(sheet.ncols):
            cell = sheet.cell(rx, cx)
            if cell.ctype == xlrd.XL_CELL_EMPTY:
                row[cx] = ""
            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                # 整数保持为 int，浮点保持 float
                if cell.value == int(cell.value):
                    row[cx] = int(cell.value)
                else:
                    row[cx] = cell.value
            else:
                row[cx] = str(cell.value).strip()
        # 跳过空行
        if any(v != "" and v is not None for v in row.values()):
            rows.append(row)

    print(f"[读取] Products 源文件: {filepath}")
    print(f"       数据行数: {len(rows)}")
    print(f"       列数: {sheet.ncols}")
    return rows


def _values_differ(old, new) -> bool:
    """比较两个值是否不同，统一转字符串比较。"""
    old_str = str(old).strip() if old is not None else ""
    new_str = str(new).strip() if new is not None else ""
    return old_str != new_str


def safe_numeric(val, default=0):
    """将值安全转为数字，失败则返回 default。"""
    if val is None or val == "" or val == "-":
        return default
    try:
        v = float(str(val).replace("%", "").replace(",", "").strip())
        return v if v == v else default  # NaN check
    except (ValueError, TypeError):
        return default


def get_source_value(src_row: dict, src_col: int):
    """从源行中取值，处理空值。"""
    val = src_row.get(src_col, "")
    if val is None:
        return ""
    return val


def build_target_id_map(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict:
    """构建分析表中 产品ID → 行号 的映射。

    返回 {产品ID: Excel行号(1-based)}
    产品ID 统一转为可比较的键 (去掉小数部分，转字符串)。
    """
    id_map = {}
    id_col = TGT["ID"] + 1  # 转为 1-based 列号
    for rx in range(TARGET_DATA_START_ROW, ws.max_row + 1):
        cell = ws.cell(row=rx, column=id_col)
        val = cell.value
        if val is not None and str(val).strip() != "":
            key = normalize_id(val)
            if key:
                id_map[key] = rx
    print(f"[索引] 分析表现有产品ID数: {len(id_map)}")
    return id_map


def normalize_id(val) -> str:
    """将产品ID统一为字符串键。

    Products 表中的 ID 是数字字符串如 "1601562216792"，
    分析表中的 ID 可能是数字或字符串，浮点数（如 218534968.0）需要去小数。
    """
    try:
        f = float(str(val).strip())
        if f == int(f):
            return str(int(f))
        return str(f)
    except (ValueError, TypeError):
        return str(val).strip()


def update_analysis(
    target_path: str,
    source_rows: list[dict],
    dry_run: bool = False,
) -> dict:
    """核心更新逻辑。

    返回统计信息 dict。
    """
    wb = openpyxl.load_workbook(target_path)
    if TARGET_SHEET not in wb.sheetnames:
        print(f"[错误] 找不到 Sheet「{TARGET_SHEET}」，可用的 Sheet: {wb.sheetnames}")
        sys.exit(1)

    ws = wb[TARGET_SHEET]
    id_map = build_target_id_map(ws)

    stats = {"updated": 0, "appended": 0, "skipped": 0, "errors": []}
    new_rows = []

    id_col_src = SRC["产品ID"]

    for i, src_row in enumerate(source_rows):
        src_id_raw = get_source_value(src_row, id_col_src)
        src_id = normalize_id(src_id_raw)
        if not src_id:
            stats["skipped"] += 1
            continue

        if src_id in id_map:
            # ---- 更新已有行 ----
            target_row = id_map[src_id]
            for src_col, tgt_col in FIELD_MAP:
                val = get_source_value(src_row, src_col)
                converter = VALUE_CONVERTERS.get(tgt_col)
                if converter:
                    val = converter(val)
                cell = ws.cell(row=target_row, column=tgt_col + 1)
                # 空值不覆盖、不标红，保留表格原有数据
                if val != "" and val is not None:
                    if tgt_col in RED_HIGHLIGHT_COLS and _values_differ(cell.value, val):
                        cell.fill = RED_FILL
                    cell.value = val

            # 2026反馈数量合计 = 原值 + 本期询盘个数 + 本期TM咨询人数
            inquiry = safe_numeric(get_source_value(src_row, SRC["询盘个数"]))
            tm = safe_numeric(get_source_value(src_row, SRC["TM咨询人数"]))
            existing_feedback = safe_numeric(ws.cell(row=target_row, column=TGT["2026反馈数量合计"] + 1).value)
            feedback_cell = ws.cell(row=target_row, column=TGT["2026反馈数量合计"] + 1)
            feedback_cell.value = existing_feedback + inquiry + tm

            stats["updated"] += 1
        else:
            # ---- 新ID，追加到末尾 ----
            new_row = {}
            new_row[TGT["ID"]] = get_source_value(src_row, SRC["产品ID"])
            new_row[TGT["曝光数"]] = get_source_value(src_row, SRC["搜索曝光次数"])
            new_row[TGT["点击数"]] = get_source_value(src_row, SRC["搜索点击次数"])
            new_row[TGT["点击率"]] = _convert_pct(get_source_value(src_row, SRC["搜索点击率"]))
            new_row[TGT["访问人数"]] = get_source_value(src_row, SRC["访问人数"])
            new_row[TGT["询盘数"]] = get_source_value(src_row, SRC["询盘个数"])
            new_row[TGT["询盘率"]] = _convert_pct(get_source_value(src_row, SRC["询盘率"]))
            new_row[TGT["TM数"]] = get_source_value(src_row, SRC["TM咨询人数"])
            new_row[TGT["是否橱窗"]] = _normalize_yn(get_source_value(src_row, SRC["是否橱窗"]))
            new_row[TGT["是否顶展"]] = _normalize_yn(get_source_value(src_row, SRC["是否顶展"]))
            new_row[TGT["是否P4P"]] = _normalize_yn(get_source_value(src_row, SRC["是否P4P"]))
            inquiry = safe_numeric(get_source_value(src_row, SRC["询盘个数"]))
            tm = safe_numeric(get_source_value(src_row, SRC["TM咨询人数"]))
            new_row[TGT["2026反馈数量合计"]] = inquiry + tm
            new_rows.append(new_row)
            stats["appended"] += 1

    # ---- 追加新行 ----
    if new_rows:
        next_row = ws.max_row + 1
        for new_row in new_rows:
            for tgt_col, val in new_row.items():
                ws.cell(row=next_row, column=tgt_col + 1).value = val
            next_row += 1

    # ---- 保存 ----
    if dry_run:
        print("\n[DRY RUN] 不保存文件")
    else:
        wb.save(target_path)
        print(f"[保存] {target_path}")

    return stats


def main():
    parser = argparse.ArgumentParser(description="Products 数据更新脚本")
    parser.add_argument(
        "--source",
        help="Products 导出文件路径。不指定则在 SOURCE_DIR 中自动查找最新文件。",
    )
    parser.add_argument(
        "--target",
        help="目标分析表路径。不指定则使用 TARGET_DIR + TARGET_FILENAME。",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="试运行，不实际修改文件",
    )
    parser.add_argument(
        "--source-dir",
        default=SOURCE_DIR,
        help=f"Products 文件所在目录 (默认: {SOURCE_DIR})",
    )
    parser.add_argument(
        "--target-dir",
        default=TARGET_DIR,
        help=f"目标分析表所在目录 (默认: {TARGET_DIR})",
    )
    args = parser.parse_args()

    # 1. 确定源文件
    if args.source:
        source_path = args.source
    else:
        source_path = find_latest_products_file(args.source_dir)
        if not source_path:
            print(f"[错误] 在 {args.source_dir} 中找不到 Products-*.xls 文件")
            sys.exit(1)
        print(f"[自动选择] 最新 Products 文件: {source_path}")

    if not os.path.exists(source_path):
        print(f"[错误] 源文件不存在: {source_path}")
        sys.exit(1)

    # 2. 确定目标文件
    if args.target:
        target_path = args.target
    else:
        target_path = os.path.join(args.target_dir, TARGET_FILENAME)

    if not os.path.exists(target_path):
        print(f"[错误] 目标文件不存在: {target_path}")
        sys.exit(1)

    print(f"[目标] {target_path}")
    print(f"[源文件] {source_path}")
    print()

    # 3. 读取源数据
    source_rows = read_products_source(source_path)

    # 4. 执行更新
    stats = update_analysis(target_path, source_rows, dry_run=args.dry_run)

    # 5. 处理完成后将源文件移入 已处理 文件夹，防止重复累加
    if not args.dry_run:
        processed_dir = os.path.join(args.source_dir, "已处理")
        os.makedirs(processed_dir, exist_ok=True)
        dest = os.path.join(processed_dir, os.path.basename(source_path))
        # 如果目标位置已有同名文件，加时间戳避免覆盖
        if os.path.exists(dest):
            stem, ext = os.path.splitext(os.path.basename(source_path))
            dest = os.path.join(processed_dir, f"{stem}_{datetime.now():%Y%m%d_%H%M%S}{ext}")
        os.rename(source_path, dest)
        print(f"[归档] {source_path} -> {dest}")

    # 6. 报告
    print(f"\n{'='*50}")
    print(f"更新完成统计")
    print(f"{'='*50}")
    print(f"  源数据行数:    {len(source_rows)}")
    print(f"  更新已有行:    {stats['updated']}")
    print(f"  追加新行:      {stats['appended']}")
    print(f"  跳过 (无ID):   {stats['skipped']}")

    if stats["errors"]:
        print(f"\n错误明细:")
        for e in stats["errors"]:
            print(f"  - {e}")

    if args.dry_run:
        print(f"\n*** 试运行模式，文件未被修改 ***")


if __name__ == "__main__":
    main()
