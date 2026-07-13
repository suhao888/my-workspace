# -*- coding: utf-8 -*-
"""
业务类型自动检测器

根据输入数据（试算平衡表、企业信息、调整项、模板目录等）
自动判断当前业务类型，选择对应模板和规则集。

检测流程：
1. 用户显式指定 → 直接使用
2. 模板目录名匹配 → 根据目录名关键词判断
3. 数据特征匹配 → TB科目/调整项特征判断
4. 模糊时提示 → 列出候选让用户选择
"""

import sys, os, re

sys.stdout.reconfigure(encoding="utf-8")

from typing import Optional, List, TYPE_CHECKING

if TYPE_CHECKING:
    from .registry import BusinessDefinition


# ============================================================
# 预定义业务类型指纹
# 用于在没有 registry 时快速判断
# ============================================================

BUSINESS_FINGERPRINTS = {
    "corporate_income_tax": {
        "name": "企业所得税汇算清缴",
        "template_dir_patterns": [
            r"企业所得税.*汇算清缴",
            r"汇算清缴.*底稿",
            r"税审.*底稿.*模板",
            r"中税网.*企业所得税",
        ],
        "required_tb_keys": ["主营业务收入", "利润总额", "所得税费用"],
        "typical_adj_items": [
            "业务招待费",
            "广告费和业务宣传费",
            "捐赠支出",
            "职工福利费",
            "工资薪金",
        ],
    },
    "high_tech_enterprise": {
        "name": "高新技术企业专项审计",
        "template_dir_patterns": [
            r"高新技术企业",
            r"高企.*认定",
            r"高企.*复审",
        ],
        "required_tb_keys": ["研发费用", "高新技术产品收入"],
        "typical_adj_items": ["研发费用加计扣除"],
        "enterprise_attrs": ["is_high_tech"],
    },
    "rd_expense_special": {
        "name": "研发费用加计扣除专项",
        "template_dir_patterns": [
            r"研发费用.*加计扣除",
            r"加计扣除.*专项",
            r"研发费用.*审核",
        ],
        "required_tb_keys": ["研发费用"],
        "typical_adj_items": ["研发费用加计扣除"],
    },
    "loss_carryforward": {
        "name": "亏损弥补审核",
        "template_dir_patterns": [
            r"亏损.*弥补",
            r"税前.*弥补.*亏损",
        ],
        "required_tb_keys": ["利润总额"],
        "typical_adj_items": ["亏损弥补"],
    },
}


def detect_business_type(
    template_dir: str = None,
    tb=None,
    enterprise=None,
    adjustments=None,
) -> Optional[str]:
    """
    自动检测业务类型

    返回业务 ID（如 "corporate_income_tax"），
    无法确定时返回 None。
    """
    scores = {}

    for biz_id, fingerprint in BUSINESS_FINGERPRINTS.items():
        score = 0

        # 模板目录名匹配（权重 3）
        if template_dir and fingerprint.get("template_dir_patterns"):
            for pattern in fingerprint["template_dir_patterns"]:
                if re.search(pattern, template_dir):
                    score += 3
                    break

        # TB 必须科目（权重 2）
        if tb and fingerprint.get("required_tb_keys"):
            matches = 0
            for key in fingerprint["required_tb_keys"]:
                if tb.get(key) and abs(tb.get(key)) > 0.01:
                    matches += 1
            if matches == len(fingerprint["required_tb_keys"]):
                score += 2  # 全部命中
            elif matches > 0:
                score += 1  # 部分命中

        # 调整项匹配（权重 1）
        if adjustments and fingerprint.get("typical_adj_items"):
            adj_names = {a.item_name for a in adjustments}
            for item in fingerprint["typical_adj_items"]:
                if item in adj_names:
                    score += 1
                    break  # 任一项命中即可

        # 企业属性匹配（权重 2）
        if enterprise and fingerprint.get("enterprise_attrs"):
            for attr in fingerprint["enterprise_attrs"]:
                if getattr(enterprise, attr, None):
                    score += 2

        if score > 0:
            scores[biz_id] = score

    if not scores:
        return None

    # 取最高分
    best = max(scores, key=scores.get)
    best_score = scores[best]

    # 如果有并列最高分，模糊
    ties = [k for k, v in scores.items() if v == best_score]
    if len(ties) > 1:
        return None  # 无法确定

    return best


def list_matching_types(
    template_dir: str = None,
    tb=None,
    enterprise=None,
    adjustments=None,
) -> List[tuple]:
    """
    列出所有匹配的业务类型及其得分

    Returns
    -------
    list of (business_id, business_name, score)
    """
    results = []
    for biz_id, fingerprint in BUSINESS_FINGERPRINTS.items():
        score = 0

        if template_dir and fingerprint.get("template_dir_patterns"):
            for pattern in fingerprint["template_dir_patterns"]:
                if re.search(pattern, template_dir):
                    score += 3
                    break

        if tb and fingerprint.get("required_tb_keys"):
            for key in fingerprint["required_tb_keys"]:
                if tb.get(key) and abs(tb.get(key)) > 0.01:
                    score += 1

        if adjustments and fingerprint.get("typical_adj_items"):
            adj_names = {a.item_name for a in adjustments}
            for item in fingerprint["typical_adj_items"]:
                if item in adj_names:
                    score += 1
                    break

        if score > 0:
            results.append((biz_id, fingerprint["name"], score))

    results.sort(key=lambda x: -x[2])
    return results
