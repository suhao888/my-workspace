#!/usr/bin/env python3
"""
扫描PDF表格提取 → Excel
========================================
基于Docling(RapidOCR)的扫描件表格提取管线。
支持中文扫描件，自动识别表格结构、去表头、过滤空行。

Usage:
    python pdf_to_excel.py <pdf_path> [--output <excel_path>] [--sheet <name>]

示例:
    python pdf_to_excel.py 2024合计.pdf --output 明细表.xlsx --sheet 2024年
    python pdf_to_excel.py 2024合计.pdf 2025合计.pdf --output 合并明细.xlsx

依赖:
    pip install docling pandas openpyxl
"""

import argparse
import re
import sys
import time
from pathlib import Path

import pandas as pd
from docling.document_converter import DocumentConverter

# ============================================================
# 配置
# ============================================================
HEADER_KEYWORDS = {
    "序号",
    "线路名称",
    "杆塔号",
    "塔号",
    "树种类别",
    "树种",
    "数量",
    "棵数",
    "树高",
    "净空",
    "处理方式",
    "所属地域",
    "省（市）",
    "地市",
    "县（市",
    "行政区",
    "所在县区",
    "行政区域",
    "树种类别",
}

COLUMN_NAMES_DEFAULT = [
    "序号",
    "线路名称",
    "杆塔号",
    "树种",
    "数量(棵)",
    "树高/净空(米)",
    "省(市)",
    "地市",
    "县(市、区)",
    "处理方式",
]


# ============================================================
# 表格清洗
# ============================================================
def is_header_row(values):
    """判断是否为表头行（含序号/线路名称等关键词 ≥2 个）"""
    text = " ".join(str(v) for v in values if v is not None)
    if not text.strip():
        return True
    return sum(1 for kw in HEADER_KEYWORDS if kw in text) >= 2


def is_empty_row(values):
    """判断是否为无数据行（≤1 个非空格）"""
    return sum(1 for v in values if v is not None and str(v).strip()) <= 1


def is_summary_row(values):
    """判断是否为合计/汇总行"""
    first = str(values[0]).strip() if values and values[0] else ""
    if not first:
        return sum(1 for v in values[1:] if v is not None and str(v).strip()) <= 2
    text = " ".join(str(v) for v in values if v is not None)
    return "合计" in text or "MAT" in text


def clean_text(val):
    """清洗OCR文本瑕疵"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    s = re.sub(r"扫描全能王.*|CamScanner.*|3亿人都在用.*", "", s)
    s = re.sub(r"https?://\S+", "", s)
    s = re.sub(r"\bMAT\b", "", s)
    s = s.replace(" ", "").replace("\u3000", "")  # 去OCR断字空格
    s = s.replace("|", "").replace("。", ".")
    s = s.strip("·,，.。、")
    return s


def clean_number(val):
    """尝试转数值，失败返回原文"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().replace(",", "").replace("，", "").replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return val


def clean_table_df(df):
    """
    清洗单页表格DataFrame：
    1. 跳过表头行（含HEADER_KEYWORDS的行）
    2. 过滤空行和合计行
    3. 清洗文本
    """
    rows = df.values.tolist()
    ncols = len(df.columns)

    # 定位数据起始行（连续表头 ≥2 关键词的行跳过）
    data_start = 0
    for i in range(min(4, len(rows))):
        vals = [str(v) if v is not None else "" for v in rows[i]]
        if is_header_row(vals):
            data_start = i + 1
        else:
            break
    # 额外检查：如果第1行还有表头特征，继续跳过
    for i in range(data_start, min(data_start + 2, len(rows))):
        if i >= len(rows):
            break
        vals = [str(v) if v is not None else "" for v in rows[i]]
        if is_header_row(vals):
            data_start = i + 1
        else:
            break

    # 收集数据行
    cleaned = []
    for i in range(data_start, len(rows)):
        vals = list(rows[i])
        vals_str = [str(v) if v is not None else "" for v in vals]
        if is_empty_row(vals_str) or is_summary_row(vals_str):
            continue
        cleaned.append(vals)

    if not cleaned:
        return pd.DataFrame()

    df_out = pd.DataFrame(cleaned)
    # 标准化列名
    rename = {}
    for i in range(min(len(COLUMN_NAMES_DEFAULT), df_out.shape[1])):
        rename[i] = COLUMN_NAMES_DEFAULT[i]
    df_out = df_out.rename(columns=rename)

    # 清洗各列
    num_cols = {"数量(棵)", "树高/净空(米)"}
    for col in df_out.columns:
        if col in num_cols:
            df_out[col] = df_out[col].apply(clean_number)
        else:
            df_out[col] = df_out[col].apply(clean_text)

    # 去空行（所有文本列均为空）
    text_cols = [c for c in df_out.columns if c not in num_cols]
    if text_cols:
        mask = ~df_out[text_cols].apply(lambda r: all(v == "" for v in r), axis=1)
        df_out = df_out[mask]

    return df_out.reset_index(drop=True)


# ============================================================
# 主管线
# ============================================================
def extract_tables(pdf_path: str, verbose: bool = True) -> pd.DataFrame:
    """
    用Docling提取PDF中所有表格并清洗，返回合并的DataFrame。
    返回的DataFrame包含COLUMN_NAMES_DEFAULT中的所有列。
    """
    pdf = Path(pdf_path)
    if not pdf.exists():
        raise FileNotFoundError(f"PDF文件不存在: {pdf}")

    t0 = time.time()
    if verbose:
        print(f"  解析: {pdf.name} ...", end=" ", flush=True)

    converter = DocumentConverter()
    result = converter.convert(str(pdf))
    doc = result.document
    tables = list(doc.tables)

    if verbose:
        print(f"{len(tables)}个表格, {time.time() - t0:.0f}s")

    all_dfs = []
    for ti, table in enumerate(tables):
        df_raw = table.export_to_dataframe(doc=doc)
        df_clean = clean_table_df(df_raw)
        if len(df_clean) > 0:
            all_dfs.append(df_clean)
        if verbose:
            print(f"    页{table.page_no} 表格{ti}: {len(df_raw)}行→{len(df_clean)}行")

    if not all_dfs:
        print("  ⚠ 未提取到有效数据")
        return pd.DataFrame(columns=COLUMN_NAMES_DEFAULT)

    combined = pd.concat(all_dfs, ignore_index=True)
    return combined


def save_excel(data: dict, output_path: str):
    """
    保存多Sheet Excel。
    data: { sheet_name: DataFrame }
    """
    output = Path(output_path)
    if output.suffix.lower() not in (".xlsx", ".xls"):
        output = output.with_suffix(".xlsx")

    with pd.ExcelWriter(str(output), engine="openpyxl") as writer:
        for sheet_name, df in data.items():
            if len(df) > 0:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  Sheet [{sheet_name}]: {len(df)}行 {len(df.columns)}列")
            else:
                print(f"  Sheet [{sheet_name}]: 空")

    print(f"\n输出: {output.resolve()}")


# ============================================================
# CLI
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="扫描PDF表格提取 → Excel (基于Docling)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("pdfs", nargs="+", help="扫描PDF文件路径（可多个）")
    parser.add_argument(
        "--output", "-o", help="输出Excel路径（默认: <第1个PDF文件名>_提取.xlsx）"
    )
    parser.add_argument(
        "--sheet", "-s", help="当只有一个PDF时指定Sheet名称（默认用PDF文件名）"
    )
    parser.add_argument(
        "--quiet", "-q", action="store_true", help="静默模式，不输出进度"
    )

    args = parser.parse_args()

    # 确定输出路径
    if args.output:
        output_path = args.output
    elif len(args.pdfs) == 1:
        stem = Path(args.pdfs[0]).stem
        output_path = f"{stem}_提取.xlsx"
    else:
        output_path = f"合并提取_{int(time.time())}.xlsx"

    # 逐个处理PDF
    all_data = {}
    for pdf_path in args.pdfs:
        pdf = Path(pdf_path)
        if not pdf.exists():
            print(f"⚠ 跳过（文件不存在）: {pdf_path}", file=sys.stderr)
            continue

        if not args.quiet:
            print(f"\n{'=' * 50}")

        try:
            df = extract_tables(pdf_path, verbose=not args.quiet)
        except Exception as e:
            print(f"⚠ 处理失败 {pdf.name}: {e}", file=sys.stderr)
            continue

        # Sheet命名
        if len(args.pdfs) == 1 and args.sheet:
            sheet_name = args.sheet
        else:
            sheet_name = pdf.stem[:31]  # Excel Sheet 名上限31字符

        all_data[sheet_name] = df

    if not all_data:
        print("没有成功处理的PDF文件")
        sys.exit(1)

    save_excel(all_data, output_path)

    # 汇总
    total_rows = sum(len(df) for df in all_data.values())
    print(f"\n总计: {len(all_data)}个PDF, {total_rows}行数据")


if __name__ == "__main__":
    main()
