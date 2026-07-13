# -*- coding: utf-8 -*-
"""
template_filler.py — 税审底稿填充调度层

调用各专用填充器（sh_filler / tb_filler / dep_filler / fuzhu_filler / a_decl_filler）
按顺序执行5种模板填充，处理 xls→xlsx 转换
"""

import sys, os, shutil

sys.stdout.reconfigure(encoding="utf-8")

from pathlib import Path

from .sh_filler import fill_sh_sheets
from .tb_filler import fill_trial_balance
from .dep_filler import fill_depreciation
from .fuzhu_filler import fill_fuzhu_digao
from .a_decl_filler import fill_A_declaration


def _xls_to_xlsx_com(xls_path: str, xlsx_path: str):
    """
    用 Excel COM 将 xls 转换为 xlsx，完整保留格式
    FileFormat=51 → xlOpenXMLWorkbook
    """
    import win32com.client as win32

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(xls_path))
    wb.SaveAs(os.path.abspath(xlsx_path), FileFormat=51)
    wb.Close(SaveChanges=False)
    excel.Quit()


def fill_all_templates(result, template_dir, output_dir=None, assets=None):
    """
    一键填充所有底稿模板

    Parameters
    ----------
    result : CalculationResult
        税审计算结果
    template_dir : str
        模板目录路径
    output_dir : str, optional
        输出目录，默认桌面
    assets : list[AssetItem], optional
        固定资产卡片列表

    Returns
    -------
    dict
        {文件类型: 输出路径}
    """
    outputs = {}
    base = output_dir or "D:/Users/12844/Desktop"

    # 模板路径
    fp0 = os.path.join(template_dir, "税审工作底稿-0.xlsx")
    fp1 = os.path.join(template_dir, "税审工作底稿-1.xls")
    fp2 = os.path.join(template_dir, "税审工作底稿-2.xls")
    fp_fuzhu = os.path.join(
        template_dir, "2026年企业所得税汇算清缴纳税调整汇总表-辅助底稿-选做.xlsx"
    )
    fp_a100 = os.path.join(template_dir, "企业所得税年度纳税申报表A类2017版.xlsx")

    # 1. SH审定表
    if os.path.exists(fp0):
        out0 = os.path.join(base, "税审底稿_SH审定表.xlsx")
        fill_sh_sheets(fp0, out0, result)
        outputs["SH审定表"] = out0
        print(f"  SH审定表 → {out0}")

    # 2. 2-00会计账簿（Excel COM 转换）
    if os.path.exists(fp1):
        temp_xlsx = os.path.join(base, "_temp_2-00.xlsx")
        _xls_to_xlsx_com(fp1, temp_xlsx)
        out1 = os.path.join(base, "税审底稿_2-00会计账簿.xlsx")
        fill_trial_balance(temp_xlsx, out1, result)
        outputs["2-00会计账簿"] = out1
        print(f"  2-00会计账簿 → {out1}")
        if os.path.exists(temp_xlsx):
            os.remove(temp_xlsx)

    # 3. 折旧审核表（Excel COM 转换）
    if os.path.exists(fp2):
        temp_xlsx2 = os.path.join(base, "_temp_3-03-01.xlsx")
        _xls_to_xlsx_com(fp2, temp_xlsx2)
        out2 = os.path.join(base, "税审底稿_折旧审核表.xlsx")
        fill_depreciation(temp_xlsx2, out2, result, assets)
        outputs["折旧审核表"] = out2
        print(f"  折旧审核表 → {out2}")
        if os.path.exists(temp_xlsx2):
            os.remove(temp_xlsx2)

    # 4. 辅助底稿
    if os.path.exists(fp_fuzhu):
        out4 = os.path.join(base, "税审底稿_辅助底稿.xlsx")
        fill_fuzhu_digao(fp_fuzhu, out4, result)
        outputs["辅助底稿"] = out4
        print(f"  辅助底稿 → {out4}")

    # 5. A类申报表
    if os.path.exists(fp_a100):
        out5 = os.path.join(base, "税审底稿_A类申报表.xlsx")
        fill_A_declaration(fp_a100, out5, result, assets)
        outputs["A类申报表"] = out5
        print(f"  A类申报表 → {out5}")

    return outputs
