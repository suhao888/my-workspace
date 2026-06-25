# -*- coding: utf-8 -*-
"""分析13家油田合同结算流程调查表，提取核心指标"""

import os
import sys
import glob
import warnings
import json

warnings.filterwarnings("ignore")

BASE_DIR = r"D:/Users/12844/Desktop/10-梳理-----0528/基础数据"

# 文件列表（按编号顺序）
FILE_LIST = [
    "1、大港油田-合同结算流程调查表（大港油田）.xls",
    "2、辽河油田-辽河油田合同结算流程调查表0515.xls",
    "3、西南油田-附件1：合同结算流程调查表(西南）5.20.xls",
    "4、塔里木油田-附件1：合同结算流程调查表（塔里木油田） - 副本.xlsx",
    "5、浙江油田-附件1：合同结算流程调查表（浙江油田分公司）.et",
    "6、玉门油田-附件1：合同结算流程调查表-上报.xlsx",
    "7、青海油田-附件1：合同结算流程调查表--青海油田.xls",
    "8、大庆油田-附件1：合同结算流程调查表（大庆油田）5.21.xls",
    "9、吉林油田-附件1：合同结算流程调查表（吉林油田分公司）5.12.xls",
    "10、长庆油田-附件1：合同结算流程调查表 - 汇总（长庆）.xls",
    "11、冀东油田-附件1：合同结算流程调查表-冀东527.xls",
    "12、华北油田-一厂合同结算流程调查表(更新5.21-华北油田）.xls",
    "13、新疆油田-修改-合同结算流程调查表（新疆油田公司）.xlsx",
]


def read_file(filepath):
    """读取Excel/et文件，返回所有行的二维列表。失败返回None。"""
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".xlsx":
            import openpyxl
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            ws = wb.active
            rows = []
            for row in ws.iter_rows(values_only=True):
                rows.append(list(row))
            wb.close()
            return rows
        elif ext in (".xls", ".et"):
            import xlrd
            wb = xlrd.open_workbook(filepath)
            ws = wb.sheet_by_index(0)
            rows = []
            for i in range(ws.nrows):
                rows.append([ws.cell_value(i, j) for j in range(ws.ncols)])
            return rows
        else:
            return None
    except Exception as e:
        print(f"  [读取失败] {e}")
        return None


def safe_str(val):
    """安全转为字符串，处理None/NaN"""
    if val is None:
        return ""
    s = str(val).strip()
    if s in ("nan", "NaN", "NONE", "None", ""):
        return ""
    return s


def get_oilfield_name(filename):
    """从文件名提取油田名称"""
    # 取第一个、号后面的中文名到第一个-
    name = filename
    if "、" in name:
        name = name.split("、", 1)[1]
    if "-" in name:
        name = name.split("-", 1)[0]
    return name.strip()


def extract_node_count(text):
    """分析审批流程文本中的节点数量"""
    text = safe_str(text)
    if not text:
        return None

    # 方法1：通过箭头/->/→分割
    import re
    # 统一分隔符
    normalized = re.sub(r'[→➜➡＞＞]', '->', text)
    normalized = re.sub(r'—\s*>', '->', normalized)
    normalized = re.sub(r'——>', '->', normalized)

    parts = re.split(r'[-—>＞>]+', normalized)
    parts = [p.strip() for p in parts if p.strip() and len(p.strip()) > 1]
    if len(parts) >= 2:
        return len(parts)

    # 方法2：通过序号 1. 2. 3. 或 ①②③ 或 （1）（2）
    numbered = re.findall(r'(?:\d+[.、）)\s]|①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩)', text)
    if len(numbered) >= 2:
        return len(numbered)

    # 方法3：通过"-"分割（步骤列表）
    parts2 = re.split(r'[-\u2014\u2015]', text)
    parts2 = [p.strip() for p in parts2 if p.strip() and len(p.strip()) > 2]
    if len(parts2) >= 2:
        return len(parts2)

    # 方法4：计算换行数
    lines = [l.strip() for l in text.split('\n') if l.strip() and len(l.strip()) > 2]
    if len(lines) >= 2:
        return len(lines)

    return None


def analyze_file(filepath, filename):
    """分析单个文件，提取8项核心指标"""
    result = {
        "油田": get_oilfield_name(filename),
        "文件名": filename,
        "合同类别覆盖": [],
        "数据行总数": 0,
        "履约确认平台": [],
        "付款金额档次": [],
        "付款审批层级数": [],
        "合同签订是否并行": False,
        "关联交易特殊审批": False,
        "关联交易描述": [],
        "线上线下结合": False,
        "线上线下描述": [],
    }

    rows = read_file(filepath)
    if rows is None or len(rows) < 6:
        result["读取状态"] = "失败"
        return result

    result["读取状态"] = "成功"
    result["总行数"] = len(rows)
    result["总列数"] = len(rows[0]) if rows else 0

    data_rows = rows[5:]  # 行5起是数据行
    result["数据行总数"] = len(data_rows)

    col_categories1 = 1   # 一级类别
    col_platform = 9       # 操作平台
    col_amount = 22        # 结算金额
    col_approval = 23      # 具体审批流程（付款）
    col_contract_approval = 5   # 合同签订具体审批流程
    col_perm_class = 16    # 权限分类（可能含关联交易）
    col_internal_approval = 17  # 具体审批流程（内部/关联）
    col_permission_4 = 4   # 权限分类（第一个）
    col_resources_24 = 24  # 所需资料（付款相关）

    # 收集所有文本用于关联交易和线上线下检查
    all_texts = []

    for row in data_rows:
        # 确保行有足够列
        while len(row) < 30:
            row.append("")

        # 1. 一级类别
        cat1 = safe_str(row[col_categories1])
        if cat1 and cat1 not in result["合同类别覆盖"]:
            result["合同类别覆盖"].append(cat1)

        # 3. 操作平台
        platform = safe_str(row[col_platform])
        if platform and platform not in result["履约确认平台"]:
            result["履约确认平台"].append(platform)

        # 4. 结算金额档次
        amount = safe_str(row[col_amount])
        if amount and amount not in result["付款金额档次"]:
            result["付款金额档次"].append(amount)

        # 5. 付款审批层级数
        approval = safe_str(row[col_approval])
        if approval:
            node_count = extract_node_count(approval)
            if node_count is not None and node_count not in result["付款审批层级数"]:
                result["付款审批层级数"].append(node_count)
            all_texts.append(approval)

        # 6. 合同签订是否并行
        contract = safe_str(row[col_contract_approval])
        if contract:
            all_texts.append(contract)
            if any(kw in contract for kw in ["并行", "同时", "同步", "会签"]):
                result["合同签订是否并行"] = True

        # 7. 关联交易/内部交易特殊审批
        perm = safe_str(row[col_perm_class])
        internal = safe_str(row[col_internal_approval])
        perm4 = safe_str(row[col_permission_4])

        for text in [perm, internal, perm4]:
            if text:
                all_texts.append(text)
                if any(kw in text for kw in ["关联交易", "内部交易", "内部结算", "关联方"]):
                    result["关联交易特殊审批"] = True
                    if text not in result["关联交易描述"]:
                        result["关联交易描述"].append(text)

        # 8. 线上线下结合检查
        for col_idx in [col_approval, col_contract_approval, 5, 23, 17]:
            t = safe_str(row[col_idx]) if col_idx < len(row) else ""
            if t:
                all_texts.append(t)

    # 线上线下结合：在所有文本中查找
    for text in all_texts:
        if any(kw in text for kw in ["线上", "线下", "纸质", "纸面", "纸质单据", "系统", "OA", "ERP", "FMIS", "纸质版", "扫描", "传签"]):
            # 需要同时出现线上+线下相关词才算"结合"
            has_online = any(kw in text for kw in ["系统", "OA", "ERP", "FMIS", "线上", "网上", "电子", "线上审批", "平台"])
            has_offline = any(kw in text for kw in ["线下", "纸质", "纸面", "纸质单据", "纸质版", "扫描", "传签", "人工", "现场"])
            if has_online and has_offline:
                result["线上线下结合"] = True
                snippet = text[:100] if len(text) > 100 else text
                if snippet not in result["线上线下描述"]:
                    result["线上线下描述"].append(snippet)

    # 清理空列表
    result["关联交易描述"] = result["关联交易描述"][:3]  # 最多保留3条
    result["线上线下描述"] = result["线上线下描述"][:3]

    return result


def print_separator():
    print("=" * 80)


def main():
    print_separator()
    print("13家油田合同结算流程调查表 - 核心指标分析报告")
    print_separator()
    print()

    all_results = []
    failed_files = []

    for i, filename in enumerate(FILE_LIST, 1):
        filepath = os.path.join(BASE_DIR, filename)
        print(f"[{i}/13] 正在分析: {filename}")

        if not os.path.exists(filepath):
            print(f"  [跳过] 文件不存在")
            failed_files.append((filename, "文件不存在"))
            continue

        result = analyze_file(filepath, filename)
        all_results.append(result)

        if result.get("读取状态") != "成功":
            print(f"  [失败] {result.get('读取状态', '未知错误')}")
            failed_files.append((filename, result.get("读取状态", "读取失败")))
            continue

        print(f"  数据行: {result['数据行总数']} | 类别: {len(result['合同类别覆盖'])}个 | 平台: {len(result['履约确认平台'])}个")
        print()

    print_separator()
    print("详细分析结果")
    print_separator()

    for result in all_results:
        if result.get("读取状态") != "成功":
            continue

        name = result["油田"]
        print(f"\n{'─' * 60}")
        print(f"  【{name}】")
        print(f"{'─' * 60}")

        print(f"  1. 合同类别覆盖 ({len(result['合同类别覆盖'])}个):")
        for cat in result["合同类别覆盖"]:
            print(f"     - {cat}")

        print(f"\n  2. 数据行总数: {result['数据行总数']}")

        print(f"\n  3. 履约确认平台 ({len(result['履约确认平台'])}个):")
        for p in result["履约确认平台"]:
            print(f"     - {p[:80]}")

        print(f"\n  4. 付款金额档次 ({len(result['付款金额档次'])}个):")
        for a in result["付款金额档次"]:
            print(f"     - {a[:80]}")

        print(f"\n  5. 付款审批层级数:")
        if result["付款审批层级数"]:
            for n in sorted(result["付款审批层级数"]):
                print(f"     - {n}层")
            print(f"     范围: {min(result['付款审批层级数'])}~{max(result['付款审批层级数'])}层")
        else:
            print("     未能提取")

        print(f"\n  6. 合同签订是否并行: {'是' if result['合同签订是否并行'] else '否'}")

        print(f"\n  7. 关联交易特殊审批: {'是' if result['关联交易特殊审批'] else '否'}")
        if result["关联交易描述"]:
            for d in result["关联交易描述"][:2]:
                print(f"     > {d[:80]}")

        print(f"\n  8. 线上线下结合: {'是' if result['线上线下结合'] else '否'}")
        if result["线上线下描述"]:
            for d in result["线上线下描述"][:2]:
                print(f"     > {d[:80]}")

    # 汇总对比表
    print(f"\n\n{'=' * 80}")
    print("汇总对比表")
    print(f"{'=' * 80}")
    print(f"{'油田':<10} {'数据行':<8} {'类别数':<8} {'平台数':<8} {'金额档次':<10} {'审批层级':<12} {'并行':<6} {'关联交易':<8} {'线上+线下':<10}")
    print("-" * 80)

    for result in all_results:
        if result.get("读取状态") != "成功":
            continue
        levels = ""
        if result["付款审批层级数"]:
            levels = f"{min(result['付款审批层级数'])}~{max(result['付款审批层级数'])}"
        else:
            levels = "N/A"

        name = result["油田"][:8]
        print(f"{name:<10} {result['数据行总数']:<8} {len(result['合同类别覆盖']):<8} "
              f"{len(result['履约确认平台']):<8} {len(result['付款金额档次']):<10} "
              f"{levels:<12} "
              f"{'是' if result['合同签订是否并行'] else '否':<6} "
              f"{'是' if result['关联交易特殊审批'] else '否':<8} "
              f"{'是' if result['线上线下结合'] else '否':<10}")

    # 失败文件
    if failed_files:
        print(f"\n\n[未能读取的文件]")
        for fn, reason in failed_files:
            print(f"  - {fn}: {reason}")

    print()

    # 输出JSON供后续使用
    output_json = os.path.join(BASE_DIR, "analysis_result.json")
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"JSON结果已保存: {output_json}")


if __name__ == "__main__":
    main()
