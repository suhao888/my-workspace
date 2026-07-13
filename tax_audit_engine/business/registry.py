# -*- coding: utf-8 -*-
"""
业务注册表 — 管理所有税务审计业务类型及其模板/规则配置

每个业务类型由以下定义：
  - 唯一 ID（如 corporate_income_tax）
  - 显示名称（如 "企业所得税汇算清缴"）
  - 检测规则（用于自动识别业务类型）
  - 模板清单（文件+配置）
  - 校验规则（勾稽关系）

配置存放于 configs/{business_id}/ 目录
"""

import sys, os, glob, yaml

sys.stdout.reconfigure(encoding="utf-8")

from pathlib import Path
from typing import Dict, List, Optional


# ============================================================
# 业务定义数据结构
# ============================================================


class BusinessDefinition:
    """一个税务审计业务类型的完整定义"""

    def __init__(self, business_id: str, manifest: dict):
        self.id = business_id
        self.name = manifest.get("business", {}).get("name", business_id)
        self.description = manifest.get("business", {}).get("description", "")
        self.detect = manifest.get("detect", {})
        self.templates = manifest.get("templates", [])
        self.checks = manifest.get("checks", [])
        self.raw_manifest = manifest

    def __repr__(self):
        return f"<Business {self.id}: {self.name}>"


# ============================================================
# 业务注册处
# ============================================================


class BusinessRegistry:
    """
    业务注册处

    负责：
    1. 扫描 configs/ 目录注册所有业务
    2. 按 ID 查找业务
    3. 按检测规则查找业务
    4. 加载业务配置
    """

    def __init__(self, configs_root: str = None):
        self._businesses: Dict[str, BusinessDefinition] = {}
        self._configs_root = configs_root

    # ---- 注册 ----

    def scan(self, configs_root: str = None):
        """扫描 configs/ 目录注册所有业务"""
        root = Path(configs_root or self._configs_root or self._default_configs_root())
        if not root.exists():
            print(f"  ⚠ 配置目录不存在: {root}")
            return

        for biz_dir in sorted(root.iterdir()):
            if not biz_dir.is_dir():
                continue
            manifest_path = biz_dir / "manifest.yaml"
            if manifest_path.exists():
                self._register_from_file(biz_dir.name, str(manifest_path))

    def _default_configs_root(self):
        """获取默认配置目录"""
        pkg_dir = Path(__file__).parent.parent
        return str(pkg_dir / "configs")

    def _register_from_file(self, business_id: str, manifest_path: str):
        """从 manifest.yaml 注册业务"""
        try:
            with open(manifest_path, "r", encoding="utf-8") as f:
                manifest = yaml.safe_load(f)
            biz = BusinessDefinition(business_id, manifest)
            self._businesses[business_id] = biz
        except Exception as e:
            print(f"  ⚠ 加载业务 {business_id} 失败: {e}")

    def register(self, business_id: str, manifest: dict):
        """直接注册业务（内存中）"""
        biz = BusinessDefinition(business_id, manifest)
        self._businesses[business_id] = biz

    # ---- 查询 ----

    def get(self, business_id: str) -> Optional[BusinessDefinition]:
        """按 ID 获取业务定义"""
        return self._businesses.get(business_id)

    def list_all(self) -> List[BusinessDefinition]:
        """列出所有已注册业务"""
        return list(self._businesses.values())

    def list_ids(self) -> List[str]:
        """列出所有已注册业务 ID"""
        return list(self._businesses.keys())

    # ---- 业务检测 ----

    def detect(
        self, tb=None, enterprise=None, template_dir=None, adjustments=None
    ) -> Optional[BusinessDefinition]:
        """
        根据输入数据自动检测业务类型

        检测策略（按优先级）：
        1. 模板目录名匹配
        2. TB 科目特征匹配
        3. 调整项特征匹配
        """
        candidates = []

        for biz in self._businesses.values():
            score = self._score_match(biz, tb, enterprise, template_dir, adjustments)
            if score > 0:
                candidates.append((score, biz))

        if not candidates:
            return None

        candidates.sort(key=lambda x: -x[0])
        best = candidates[0]

        # 如果有多个候选得分相同且 > 0，模糊
        if len(candidates) > 1 and candidates[1][0] == best[0]:
            return None  # 模糊，需要用户指定

        return best[1]

    def _score_match(
        self,
        biz: BusinessDefinition,
        tb=None,
        enterprise=None,
        template_dir=None,
        adjustments=None,
    ) -> int:
        """
        计算匹配得分
        命中检测规则中的一项得1分
        """
        score = 0
        detect = biz.detect

        # 模板目录名匹配
        if template_dir and detect.get("template_dir_pattern"):
            import re

            for pattern in detect["template_dir_pattern"]:
                if re.search(pattern, template_dir):
                    score += 2  # 目录匹配权重更高

        # TB 科目匹配
        if tb and detect.get("has_tb_keys"):
            for key in detect["has_tb_keys"]:
                if tb.get(key) and abs(tb.get(key)) > 0.01:
                    score += 1

        # 企业信息匹配
        if enterprise and detect.get("has_enterprise_attrs"):
            for attr in detect["has_enterprise_attrs"]:
                if getattr(enterprise, attr, None):
                    score += 1

        # 调整项匹配
        if adjustments and detect.get("has_adj_items"):
            adj_names = {a.item_name for a in adjustments}
            for item in detect["has_adj_items"]:
                if item in adj_names:
                    score += 1

        return score

    # ---- 加载模板配置 ----

    def load_template_config(
        self, business_id: str, template_id: str
    ) -> Optional[dict]:
        """
        加载某个模板的 YAML 配置

        配置路径: configs/{business_id}/templates/{template_id}.yaml
        """
        biz = self.get(business_id)
        if not biz:
            return None

        # 先找模板定义
        tmpl_def = None
        for tmpl in biz.templates:
            if tmpl.get("id") == template_id:
                tmpl_def = tmpl
                break
        if not tmpl_def:
            return None

        # 加载 YAML 配置
        config_file = tmpl_def.get("config")
        if not config_file:
            return tmpl_def  # 无独立配置，直接用模板定义

        configs_root = Path(self._configs_root or self._default_configs_root())
        config_path = configs_root / business_id / config_file
        if not config_path.exists():
            print(f"  ⚠ 配置不存在: {config_path}")
            return tmpl_def

        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)

        # 合并：模板定义属性 → 加载的配置
        if config:
            config.setdefault("convert", tmpl_def.get("convert", False))
            config.setdefault("output", tmpl_def.get("output", ""))

        return config or tmpl_def
