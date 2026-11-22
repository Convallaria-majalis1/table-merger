# -*- coding: utf-8 -*-
"""
Excel / CSV 多表合并清洗器 —— 核心逻辑（不依赖界面，可单独运行、可单独测试）

设计原则：
1. 只管"把多个零散表格合成一张干净的宽表"，不做复杂业务。
2. 所有列名、单元格两端的空格都清掉，中文全角空格也清。
3. 类似"订单号 / 身份证 / 手机号"这类列，强制按文本处理，避免变成科学计数法。
4. 出错时给出人话提示，而不是抛一堆 traceback 给客户看。

命令行用法（方便你自己在终端测试，不需要打开界面）：
    python core.py 文件1.xlsx 文件2.csv -o 结果.xlsx --dedupe
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import pandas as pd

# CSV 常见编码：带 BOM 的 UTF-8、国标 GB18030（兼容 GBK）、纯 UTF-8
CSV_ENCODINGS = ("utf-8-sig", "gb18030", "utf-8")

# 命中这些关键词的列，按文本处理，避免长数字被 Excel 变成 1.23E+17
TEXT_COLUMN_HINTS = (
    "订单号", "单号", "编号", "编码", "身份证", "证件", "手机", "电话",
    "邮编", "卡号", "账号", "工号", "学号", "批次号", "合同号", "发票号",
)

# 合并时自动加上的来源列，去重时要排除掉，否则"同一行来自不同文件"永远不算重复
SOURCE_COLUMNS = ("来源文件", "来源Sheet")


def is_text_column(col_name: str) -> bool:
    """判断某一列是否应该强制当作文本处理"""
    name = str(col_name)
    return any(hint in name for hint in TEXT_COLUMN_HINTS)


def _clean_cell(value):
    """清理单个单元格：去两端空格（含全角空格、制表符）"""
    if isinstance(value, str):
        return value.replace("\u3000", " ").strip()
    return value


def _is_blank(value) -> bool:
    """判断单元格是否为空：""、纯空格、None、NaN、pandas 的空值都算空"""
    if value is None:
        return True
    if isinstance(value, str):
        return value.replace("\u3000", " ").strip() == ""
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return False


def read_csv_file(path: str) -> list[pd.DataFrame]:
    """读 CSV，自动尝试多种编码"""
    last_error = None
    for encoding in CSV_ENCODINGS:
        try:
            df = pd.read_csv(path, encoding=encoding, dtype=object, keep_default_na=False)
            return [df]
        except UnicodeDecodeError as e:
            last_error = e
            continue
        except pd.errors.EmptyDataError:
            # 空文件，直接跳过
            return []
    raise ValueError(f"读取 CSV 失败（编码无法识别）：{Path(path).name}\n详细信息：{last_error}")


def read_excel_file(path: str, all_sheets: bool = False) -> list[pd.DataFrame]:
    """读 Excel，可选择一个文件里的所有 Sheet"""
    if all_sheets:
        sheets = pd.read_excel(path, sheet_name=None, dtype=object, keep_default_na=False)
        frames = []
        for sheet_name, df in sheets.items():
            if df is None or df.empty:
                continue
            df = df.copy()
            df["来源Sheet"] = str(sheet_name)
            frames.append(df)
        return frames

    df = pd.read_excel(path, dtype=object, keep_default_na=False)
    return [] if df.empty else [df]


def read_any_file(path: str, all_sheets: bool = False) -> list[pd.DataFrame]:
    """按后缀分发读取"""
    suffix = Path(path).suffix.lower()
    if suffix == ".csv":
        return read_csv_file(path)
    if suffix in (".xlsx", ".xlsm", ".xls"):
        return read_excel_file(path, all_sheets=all_sheets)
    raise ValueError(f"不支持的文件格式：{Path(path).name}（仅支持 .xlsx / .xls / .csv）")


def normalize_frame(df: pd.DataFrame) -> pd.DataFrame:
    """规范化单张表：列名去空格、单元格去空格、长数字列转文本"""
    df = df.copy()

    # 列名去空格，顺带把"未命名: 0"这种空列名改成人能看懂的名字
    new_columns = []
    for i, col in enumerate(df.columns):
        name = str(col).replace("\u3000", " ").strip()
        if not name or name.lower().startswith("unnamed"):
            name = f"未命名列{i + 1}"
        new_columns.append(name)
    df.columns = new_columns

    # 单元格清理
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(_clean_cell)

    # 长数字列强制文本
    for col in df.columns:
        if is_text_column(col):
            df[col] = df[col].map(lambda v: "" if v is None else str(v).strip())

    return df


def merge_files(
    paths: list[str],
    all_sheets: bool = False,
    add_source: bool = True,
    log=None,
) -> pd.DataFrame:
    """
    把多个文件（或一个文件的多个 Sheet）纵向合并成一张表。
    列不一致时自动取并集，缺的格子留空。
    """
    if not paths:
        raise ValueError("没有选择任何文件")

    frames = []
    for path in paths:
        for df in read_any_file(path, all_sheets=all_sheets):
            df = normalize_frame(df)
            if add_source:
                df.insert(0, "来源文件", Path(path).name)
            frames.append(df)
            if log:
                log(f"已读取：{Path(path).name}（{len(df)} 行 / {len(df.columns)} 列）")

    if not frames:
        raise ValueError("所选文件里没有任何有效数据（可能全是空表）")

    merged = pd.concat(frames, ignore_index=True, sort=False)

    # 合并后统一列顺序：来源列放最前，其余按首次出现顺序
    columns = list(merged.columns)
    ordered = [c for c in SOURCE_COLUMNS if c in columns]
    ordered += [c for c in columns if c not in ordered]
    return merged[ordered]


def clean_frame(
    df: pd.DataFrame,
    drop_empty_rows: bool = True,
    drop_empty_cols: bool = True,
    drop_duplicates: bool = False,
    dedupe_keep_source: bool = True,
) -> pd.DataFrame:
    """对合并后的表做清洗"""
    df = df.copy()

    if drop_empty_cols:
        keep_cols = []
        for col in df.columns:
            # 来源列永远保留
            if col in SOURCE_COLUMNS or not all(_is_blank(v) for v in df[col].tolist()):
                keep_cols.append(col)
        df = df[keep_cols]

    if drop_empty_rows:
        # 判断"空行"时不能把来源列算进去，否则每行都因为来源文件非空而被判定为有内容
        check_cols = [c for c in df.columns if c not in SOURCE_COLUMNS]
        if check_cols:
            row_non_empty = df[check_cols].apply(
                lambda row: not all(_is_blank(v) for v in row.values), axis=1
            )
            df = df[row_non_empty]

    if drop_duplicates:
        subset = None
        if dedupe_keep_source:
            subset = [c for c in df.columns if c not in SOURCE_COLUMNS]
        df = df.drop_duplicates(subset=subset, keep="first")

    return df.reset_index(drop=True)


def write_result(df: pd.DataFrame, out_path: str) -> str:
    """写出结果文件"""
    out_path = str(out_path)
    suffix = Path(out_path).suffix.lower()

    # 输出目录不存在就建上
    parent = Path(out_path).parent
    if str(parent) and not parent.exists():
        parent.mkdir(parents=True, exist_ok=True)

    if suffix == ".csv":
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
        return out_path

    if suffix not in (".xlsx", ".xlsm"):
        # 用户没写后缀时默认 xlsx
        out_path = out_path + ".xlsx"

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="合并结果")
        worksheet = writer.sheets["合并结果"]

        # 列宽自适应（中文按 2 个字符宽度估算），并给文本列设置文本格式
        from openpyxl.utils import get_column_letter

        for idx, col in enumerate(df.columns, start=1):
            series = df[col].astype(str)
            max_len = max([len(str(col))] + [len(s) for s in series.head(200)])
            width = min(max(max_len * 2 + 2, 8), 50)
            letter = get_column_letter(idx)
            worksheet.column_dimensions[letter].width = width
            if is_text_column(col):
                for cell in worksheet[letter][1:]:
                    cell.number_format = "@"

        worksheet.freeze_panes = "A2"

    return out_path


def run(
    paths: list[str],
    out_path: str,
    all_sheets: bool = False,
    drop_duplicates: bool = False,
    drop_empty_rows: bool = True,
    drop_empty_cols: bool = True,
    add_source: bool = True,
    log=None,
) -> dict:
    """完整流程：读取 -> 合并 -> 清洗 -> 写出，返回统计信息"""
    merged = merge_files(paths, all_sheets=all_sheets, add_source=add_source, log=log)
    before = len(merged)
    cleaned = clean_frame(
        merged,
        drop_empty_rows=drop_empty_rows,
        drop_empty_cols=drop_empty_cols,
        drop_duplicates=drop_duplicates,
    )
    final_path = write_result(cleaned, out_path)
    stats = {
        "文件数": len(paths),
        "合并后行数": before,
        "清洗后行数": len(cleaned),
        "列数": len(cleaned.columns),
        "输出路径": final_path,
    }
    if log:
        log(f"完成：{before} 行 -> {len(cleaned)} 行，共 {len(cleaned.columns)} 列")
        log(f"已保存：{final_path}")
    return stats


def main():
    parser = argparse.ArgumentParser(description="把多个 Excel/CSV 合并成一张干净的表")
    parser.add_argument("files", nargs="+", help="要合并的文件路径")
    parser.add_argument("-o", "--output", default="合并结果.xlsx", help="输出文件路径")
    parser.add_argument("--all-sheets", action="store_true", help="读取每个文件的全部 Sheet")
    parser.add_argument("--dedupe", action="store_true", help="去掉重复行")
    parser.add_argument("--keep-empty-rows", action="store_true", help="保留空行")
    args = parser.parse_args()

    try:
        stats = run(
            args.files,
            args.output,
            all_sheets=args.all_sheets,
            drop_duplicates=args.dedupe,
            drop_empty_rows=not args.keep_empty_rows,
            log=lambda msg: print(msg),
        )
    except Exception as e:
        print(f"处理失败：{e}", file=sys.stderr)
        sys.exit(1)

    for key, value in stats.items():
        print(f"{key}：{value}")


if __name__ == "__main__":
    main()
