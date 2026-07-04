"""
模板填充器 — 将计算结果写入模板文件

策略：保留原始模板所有格式，只填充数据单元格。
模板中已有公式（如 H=G-E），引擎只需填入 E(账载金额) 和 G(税收金额) 列。

支持模板：
  1. 辅助底稿-选做（汇总调整表）
  2. 税审工作底稿-0（封面 + SH系列）
  3. A类申报表（A100000 + A105000）
"""

import shutil
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .models import CalculationResult, TaxAdjustment, AdjustmentCategory


class TemplateFiller:
    """
    模板填充器
    """

    def __init__(self, result: CalculationResult):
        self.result = result
        self._enterprise = result.enterprise

    @staticmethod
    def _safe_set(ws, row, col, value):
        """安全写入单元格，跳过合并格的不可写区域"""
        try:
            ws.cell(row=row, column=col).value = value
        except AttributeError:
            for mr in ws.merged_cells.ranges:
                if ws.cell(row=row, column=col).coordinate in mr:
                    ws.cell(row=mr.min_row, column=mr.min_col).value = value
                    break

    # ============================================================
    # 辅助底稿-选做 → 汇总调整表
    # ============================================================

    # 行映射规则：(模板行号, 匹配关键词列表, 类别处理方式)
    # 类别处理方式: 'income'=H=G-E, 'deduction'=H=E-G, 'incentive'=H=G-E
    ROW_MAP = {
        # 收入类 (R3: 一、收入类调整项目, R4-R8)
        4: (["视同销售收入"], "income"),
        5: (["投资收益", "权益法", "股息", "红利"], "income"),
        6: (["公允价值变动"], "income"),
        7: (["其他"], "income"),  # 收入-其他
        # 扣除类 (R9: 二、扣除类调整项目, R10-R29)
        10: (["视同销售成本"], "deduction"),
        11: (["工资薪金"], "deduction"),
        12: (["职工福利费"], "deduction"),
        13: (["职工教育经费"], "deduction"),
        14: (["工会经费"], "deduction"),
        15: (["基本社会保险", "社会保"], "deduction"),
        16: (["住房公积金"], "deduction"),
        17: (["补充养老保险"], "deduction"),
        18: (["补充医疗保险"], "deduction"),
        20: (["业务招待费"], "deduction"),
        21: (["广告费", "宣传费"], "deduction"),
        22: (["公益性捐赠", "捐赠"], "deduction"),
        23: (["税收滞纳金", "加收利息"], "deduction"),
        24: (["罚金", "罚款", "没收"], "deduction"),
        27: (["党组织", "党"], "deduction"),
        28: (["赞助", "无关"], "deduction"),
        # 资产类 (R30: 三、资产类调整项目, R31-R35)
        31: (["折旧"], "asset"),
        32: (["摊销"], "asset"),
        33: (["资产减值", "减值准备"], "asset"),
        # 特殊事项 (R36: 四、特殊事项, R37-R38)
        37: (["重组"], "special"),
        # 免税/加计扣除 (R39: 五、免税、减计收入及加计扣除, R40-R43)
        40: (["免税", "国债利息", "股息红利-免税"], "incentive"),
        41: (["研发费用加计扣除", "加计扣除"], "incentive"),
        42: (["残疾人", "残疾"], "incentive"),
    }

    def fill_fuzhu_digao(self, template_path: str, output_path: str) -> str:
        """
        填充辅助底稿-选做

        Parameters
        ----------
        template_path : str
            辅助底稿-选做模板文件路径
        output_path : str
            输出路径

        Returns
        -------
        str
            输出路径
        """
        src = Path(template_path)
        dst = Path(output_path)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))

        import openpyxl

        wb = openpyxl.load_workbook(str(dst))

        # 找到汇总调整表sheet
        sheet_name = None
        for name in wb.sheetnames:
            if "调整" in name or "汇总" in name:
                sheet_name = name
                break
        if not sheet_name:
            sheet_name = wb.sheetnames[1]  # fallback

        ws = wb[sheet_name]

        # 1. 更新标题中的企业名称
        enterprise_name = self._enterprise.name or "被审计单位"
        ws["A1"] = f"2025年所得税纳税调整明细表-{enterprise_name}"

        # 2. 分类汇总各调整项
        cat_totals = {}  # category -> (E_total, G_total)
        # 将调整项按映射分配到模板行
        row_items: Dict[int, List[TaxAdjustment]] = {}
        for adj in self.result.adjustments:
            if adj.increase == 0 and adj.decrease == 0:
                continue
            matched_row = self._match_row(adj)
            if matched_row:
                row_items.setdefault(matched_row, []).append(adj)

        # 3. 填入每个行的 E(账载金额) 和 G(税收金额)
        income_total_inc = 0  # 收入类调增合计
        income_total_dec = 0  # 收入类调减合计
        deduction_total = 0
        asset_total = 0
        special_total = 0
        incentive_total = 0

        for row_num, items in row_items.items():
            if row_num not in self.ROW_MAP:
                continue
            _, adj_type = self.ROW_MAP[row_num]

            # 汇总归到同一行的调整项
            total_book = sum(a.book_amount for a in items)
            total_tax = sum(a.tax_base for a in items)

            # 如果没有book_amount/tax_base，用increase/decrease反推
            if total_book == 0 and total_tax == 0:
                total_inc = sum(a.increase for a in items)
                total_dec = sum(a.decrease for a in items)
                if adj_type in ("income", "incentive", "special"):
                    # H=G-E: positive = increase
                    if total_inc > 0:
                        total_tax = total_inc
                    elif total_dec > 0:
                        total_book = total_dec
                elif adj_type in ("deduction", "asset"):
                    # H=E-G: positive = increase
                    if total_inc > 0:
                        total_book = total_inc
                    elif total_dec > 0:
                        total_tax = total_dec

            # 填入E列(账载金额)和G列(税收金额)
            # 根据调整类型决定谁大谁小，让公式自动算出正确的调整额
            if adj_type in ("income", "incentive", "special"):
                # 调增时：G>E, 调减时：E>G
                net = sum(a.increase for a in items) - sum(a.decrease for a in items)
                if net >= 0:
                    ws.cell(row=row_num, column=7, value=total_book + net)  # G=税收
                    ws.cell(row=row_num, column=5, value=total_book)  # E=账面
                else:
                    ws.cell(row=row_num, column=5, value=total_book - net)  # E=账面
                    ws.cell(row=row_num, column=7, value=total_book)  # G=税收
            else:
                # 扣除/资产类：调增时 E>G, 调减时 G>E
                net = sum(a.increase for a in items) - sum(a.decrease for a in items)
                if net >= 0:
                    ws.cell(row=row_num, column=5, value=total_book + net)  # E=账面
                    ws.cell(row=row_num, column=7, value=total_book)  # G=税收
                else:
                    ws.cell(row=row_num, column=7, value=total_book - net)  # G=税收
                    ws.cell(row=row_num, column=5, value=total_book)  # E=账面

        # 4. 填入利润总额
        profit_row = self._find_row_by_label(ws, "利润总额")
        if profit_row:
            ws.cell(row=profit_row, column=7, value=self.result.accounting_profit)

        # 5. 填入税率
        rate_row = self._find_row_by_label(ws, "税率")
        if rate_row:
            ws.cell(row=rate_row, column=8, value=self.result.tax_rate)

        # 6. 填入弥补亏损（如果有）
        loss_row = self._find_row_by_label(ws, "亏损")
        if loss_row:
            ws.cell(row=loss_row, column=8, value=0)

        # 7. 填入所得税减免（如果有）
        deduction_row = self._find_row_by_label(ws, "减免")
        if deduction_row:
            ws.cell(row=deduction_row, column=8, value=self.result.deducted_tax)

        # 8. 填入预缴税额（如果有）
        prepay_row = self._find_row_by_label(ws, "预缴")
        if prepay_row:
            ws.cell(row=prepay_row, column=8, value=0)

        wb.save(str(dst))
        return str(dst)

    def _match_row(self, adj: TaxAdjustment) -> Optional[int]:
        """匹配调整项到模板行号"""
        name = adj.item_name
        for row_num, (keywords, _) in self.ROW_MAP.items():
            for kw in keywords:
                if kw in name:
                    return row_num
        return None

    def _find_row_by_label(self, ws, keyword: str) -> Optional[int]:
        """按关键字查找行号"""
        for row in range(1, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=7).value
            if cell_val and keyword in str(cell_val):
                return row
            cell_val = ws.cell(row=row, column=2).value
            if cell_val and keyword in str(cell_val):
                return row
        return None

    # ============================================================
    # 税审工作底稿-0 → 封面
    # ============================================================

    def fill_cover(self, template_path: str, output_path: str) -> str:
        """填充封面"""
        src = Path(template_path)
        dst = Path(output_path)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))

        import openpyxl

        wb = openpyxl.load_workbook(str(dst))

        # 找封面sheet（通常在第一个或名称含"封面"）
        sheet_name = None
        for name in wb.sheetnames:
            if "封面" in name:
                sheet_name = name
                break
        if not sheet_name:
            sheet_name = wb.sheetnames[0]
        ws = wb[sheet_name]

        # 根据模板分析结果填入各字段
        # 封面中字段通过相邻标签单元格定位
        field_map = {
            "被审核单位": self._enterprise.name,
            "纳税人识别号": self._enterprise.uscc,
            "纳税年度": str(self._enterprise.tax_year),
            "所属行业": self._enterprise.industry,
            "法定代表人": self._enterprise.legal_rep,
            "资产总额": self._enterprise.total_assets,
            "从业人数": self._enterprise.employee_count,
        }

        # 收集合并单元格范围信息
        merged_ranges = list(ws.merged_cells.ranges)

        def find_writable_cell(row, col_start):
            """在合并单元格中找到可写的左上角单元格"""
            for _col in range(col_start, ws.max_column + 1):
                cell = ws.cell(row=row, column=_col)
                # 检查是否是合并单元格的一部分
                is_merged = False
                for mr in merged_ranges:
                    if cell.coordinate in mr:
                        # 返回合并范围的左上角
                        is_merged = True
                        tl = ws.cell(row=mr.min_row, column=mr.min_col)
                        if tl.value is None or str(tl.value).strip() == "":
                            return tl
                        break
                if not is_merged:
                    if cell.value is None or str(cell.value).strip() == "":
                        return cell
            return None

        for row in range(1, min(ws.max_row + 1, 50)):
            for col in range(1, min(ws.max_column + 1, 20)):
                val = ws.cell(row=row, column=col).value
                if val and isinstance(val, str):
                    for field_name, field_value in field_map.items():
                        if field_name in val and field_value:
                            target = find_writable_cell(row, col + 1)
                            if target:
                                try:
                                    target.value = field_value
                                except AttributeError:
                                    pass
                            break
                            break

        wb.save(str(dst))
        return str(dst)

    # ============================================================
    # A类申报表 → A100000 主表
    # ============================================================

    def fill_A100000(self, template_path: str, output_path: str) -> str:
        """
        填充A类申报表主表

        A100000 关键行：
          R6(D)=营业收入
          R23(D)=利润总额
          R25(D)=纳税调整增加额
          R26(D)=纳税调整减少额
          R27(D)=免税、减计收入及加计扣除
          R32(D)=应纳税所得额
          R33(D)=税率
          R34(D)=应纳所得税额
          R35(D)=减免所得税额
          R36(D)=实际应纳所得税额
        """
        src = Path(template_path)
        dst = Path(output_path)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))

        import openpyxl

        wb = openpyxl.load_workbook(str(dst))

        # 找A100000 sheet（实际名称为"A100000年度纳税申报表"）
        sheet_name = None
        for name in wb.sheetnames:
            if "100000" in name:
                sheet_name = name
                break
        if not sheet_name:
            for name in wb.sheetnames:
                if "年度纳税申报" in name:
                    sheet_name = name
                    break
        if not sheet_name:
            print("  [警告] 未找到A100000 sheet")
            return str(dst)
        ws = wb[sheet_name]

        # 用行标签定位
        fill = {
            "利润总额": self.result.accounting_profit,
            "纳税调整增加": self.result.total_increase,
            "纳税调整减少": self.result.total_decrease,
            "应纳税所得额": self.result.taxable_income,
            "税率": self.result.tax_rate,
            "应纳所得税额": self.result.tax_payable,
            "减免所得税": self.result.deducted_tax,
            "实际应纳所得税": self.result.final_tax,
        }

        for row in range(1, ws.max_row + 1):
            c_val = ws.cell(row=row, column=3).value  # C列=项目名称
            if c_val and isinstance(c_val, str):
                for keyword, value in fill.items():
                    if keyword in c_val:
                        ws.cell(row=row, column=4, value=value)
                        break

        wb.save(str(dst))
        return str(dst)

    # ============================================================
    # A类申报表 → A105000 纳税调整明细表
    # ============================================================

    def fill_A105000(self, template_path: str, output_path: str) -> str:
        """
        填充A105000纳税调整明细表

        A105000 调增/调减列：
          E列 = 调增金额
          F列 = 调减金额
        """
        src = Path(template_path)
        dst = Path(output_path)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))

        import openpyxl

        wb = openpyxl.load_workbook(str(dst))

        # 找A105000 sheet（实际名称为"A105000纳税调整项目表"）
        sheet_name = None
        for name in wb.sheetnames:
            if "105000" in name:
                sheet_name = name
                break
        if not sheet_name:
            for name in wb.sheetnames:
                if "纳税调整" in name and "项目" in name:
                    sheet_name = name
                    break
        if not sheet_name:
            print("  [警告] 未找到A105000 sheet")
            return str(dst)
        ws = wb[sheet_name]

        # A105000 行映射: (行号, 关键词, 是否间接计算)
        # 直接填入调增金额(E列)和调减金额(F列)
        a105000_map = {
            8: ["视同销售收入"],
            9: ["未按权责发生制"],
            10: ["投资收益"],
            11: ["权益法", "初始投资"],
            13: ["公允价值变动"],
            15: ["不征税收入"],
            16: ["折扣折让"],
            # 扣除类
            20: ["职工薪酬"],
            21: ["业务招待费"],
            22: ["广告", "宣传费"],
            23: ["捐赠支出"],
            24: ["利息支出"],
            25: ["罚金", "罚款", "没收"],
            26: ["税收滞纳金", "加收利息"],
            27: ["赞助支出"],
            29: ["佣金"],
            33: ["无关", "无关支出"],
            35: ["党组织"],
            # 资产类
            37: ["折旧", "摊销"],
            38: ["减值准备"],
            39: ["资产损失"],
            # 特殊事项
            42: ["重组"],
            # 免税
            47: ["免税", "国债利息"],
            48: ["研发费用加计扣除"],
            49: ["残疾人"],
        }

        # 先汇总所有调整
        inc_map: Dict[str, float] = {}
        dec_map: Dict[str, float] = {}
        for adj in self.result.adjustments:
            name = adj.item_name
            if adj.increase > 0:
                inc_map[name] = inc_map.get(name, 0) + adj.increase
            if adj.decrease > 0:
                dec_map[name] = dec_map.get(name, 0) + adj.decrease

        for row_num, keywords in a105000_map.items():
            total_inc = 0
            total_dec = 0
            for inc_name, inc_val in inc_map.items():
                for kw in keywords:
                    if kw in inc_name:
                        total_inc += inc_val
                        break
            for dec_name, dec_val in dec_map.items():
                for kw in keywords:
                    if kw in dec_name:
                        total_dec += dec_val
                        break
            if total_inc > 0:
                ws.cell(row=row_num, column=5, value=total_inc)
            if total_dec > 0:
                ws.cell(row=row_num, column=6, value=total_dec)

        # 汇总行（手工汇总公式，不需要改，会自算）

        wb.save(str(dst))
        return str(dst)

    # ============================================================
    # 研发费用加计扣除 → 审核汇总表(3-01) + 优惠审核表(3-02)
    # ============================================================

    def fill_rd_workpaper(
        self,
        template_path: str,
        output_path: str,
        rd_result=None,
        rd_input_projects=None,
    ) -> str:
        """
        填充研发费用加计扣除工作底稿

        主要填充 3-01(审核汇总表) 和 3-02(优惠审核表)
        """
        src = Path(template_path)
        dst = Path(output_path)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))

        import openpyxl

        wb = openpyxl.load_workbook(str(dst))

        # --- 3-02 优惠审核表 ---
        sheet_302 = None
        for name in wb.sheetnames:
            if "3-02" in name:
                sheet_302 = wb[name]
                break
        if sheet_302 is not None and rd_result is not None:
            # 行映射（基于模板B列=项目名称）
            rd_map_302 = {
                6: "一、自主研发、合作研发、集中研发",
                7: "（一）人员人工费用",
                8: "1.直接从事研发活动人员工资薪金",
                11: "（二）直接投入费用",
                16: "（三）折旧费用",
                19: "（四）无形资产摊销",
                23: "（五）设计费用",
                26: "（八）其他相关费用",
                27: "（八）其他相关费用限额",
                36: "三、年度可加计扣除研发费用小计",
                37: "四、计入本年损益的加计扣除额",
            }
            # 填入基础信息
            for row in range(1, min(sheet_302.max_row + 1, 5)):
                for col in range(1, sheet_302.max_column + 1):
                    v = sheet_302.cell(row=row, column=col).value
                    if v and isinstance(v, str):
                        if "被审核单位" in v and self._enterprise.name:
                            self._safe_set(
                                sheet_302, row, col + 1, self._enterprise.name
                            )
                        if "年  度" in v:
                            self._safe_set(
                                sheet_302,
                                row,
                                col + 1,
                                f"{self._enterprise.tax_year}年度",
                            )

            # 填入F列（金额）
            total_deduction = rd_result.deduction_amount
            for row_num, label in rd_map_302.items():
                if row_num == 36:
                    self._safe_set(sheet_302, row_num, 6, total_deduction)
                elif row_num == 37 and rd_result.qualifying_amount > 0:
                    self._safe_set(sheet_302, row_num, 6, rd_result.deduction_amount)
                elif row_num == 6 and total_deduction > 0:
                    self._safe_set(sheet_302, row_num, 6, rd_result.qualifying_amount)
                elif row_num == 6:  # 一、自主研发...
                    if total_deduction > 0:
                        sheet_302.cell(
                            row=row_num, column=6, value=rd_result.qualifying_amount
                        )

        # --- 3-01 审核汇总表 ---
        sheet_301 = None
        for name in wb.sheetnames:
            if "3-01" in name:
                sheet_301 = wb[name]
                break
        if sheet_301 is not None and rd_result is not None:
            # 填入费用分类合计到J列
            cat_map = {
                7: "人员人工费用",
                11: "直接投入费用",
                20: "折旧费用",
            }
            for row_num, label in cat_map.items():
                self._safe_set(sheet_301, row_num, 10, 0)

        wb.save(str(dst))
        return str(dst)

    # ============================================================
    # 全税种核查 → 各税种汇总表(2-01) + 应交税金审核表(2-05)
    # ============================================================

    def fill_fulltax(
        self, template_path: str, output_path: str, fulltax_result=None
    ) -> str:
        """填充全税种核查测算底稿"""
        src = Path(template_path)
        dst = Path(output_path)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))

        import openpyxl

        wb = openpyxl.load_workbook(str(dst))

        # 2-01 各税种汇总表
        sheet = None
        for name in wb.sheetnames:
            if "2-01" in name:
                sheet = wb[name]
                break
        if sheet is not None and fulltax_result is not None:
            tax_map = {}
            for row in range(5, sheet.max_row + 1):
                tax_name = sheet.cell(row=row, column=2).value
                if tax_name and isinstance(tax_name, str) and len(tax_name.strip()) > 0:
                    tax_map[tax_name.strip()] = row
            for item in fulltax_result.item_details:
                if hasattr(item, "tax_type"):
                    tax_name = item.tax_type
                else:
                    continue
                if tax_name in tax_map:
                    rn = tax_map[tax_name]
                    if hasattr(item, "declared_amount"):
                        self._safe_set(sheet, rn, 5, item.declared_amount)
                    if hasattr(item, "actual_amount"):
                        self._safe_set(sheet, rn, 6, item.actual_amount)

        wb.save(str(dst))
        return str(dst)

    # ============================================================
    # 财产损失 → 损失汇总表
    # ============================================================

    def fill_loss(self, template_path: str, output_path: str, loss_result=None) -> str:
        """填充财产损失审核表"""
        src = Path(template_path)
        dst = Path(output_path)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))

        import openpyxl

        wb = openpyxl.load_workbook(str(dst))

        # 找损失汇总表sheet
        sheet = None
        for name in wb.sheetnames:
            if "汇总" in name or "损失" in name:
                sheet = wb[name]
                break
        if sheet is None:
            sheet = wb.active

        if loss_result is not None:
            for row in range(1, min(sheet.max_row + 1, 10)):
                for col in range(1, sheet.max_column + 1):
                    v = sheet.cell(row=row, column=col).value
                    if v and isinstance(v, str) and ("损失合计" in v or "合计" in v):
                        for cc in range(1, sheet.max_column + 1):
                            try:
                                sv = sheet.cell(row=row, column=cc).value
                                if sv is None:
                                    sheet.cell(
                                        row=row, column=cc
                                    ).value = loss_result.qualifying_loss
                                    break
                            except AttributeError:
                                pass

        wb.save(str(dst))
        return str(dst)

    # ============================================================
    # 高新鉴证 → 研发费用审定表 + 结构明细表 + 收入审定表
    # ============================================================

    def fill_hightech(
        self, template_path: str, output_path: str, hightech_result=None
    ) -> str:
        """填充高新鉴证底稿"""
        src = Path(template_path)
        dst = Path(output_path)
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))

        import openpyxl

        wb = openpyxl.load_workbook(str(dst))

        if hightech_result is None:
            wb.save(str(dst))
            return str(dst)

        # 研发费用审定表
        sheet_rd = None
        for name in wb.sheetnames:
            if "研发费用" in name and "审定" in name:
                sheet_rd = wb[name]
                break
        if sheet_rd is not None:
            for row in range(1, sheet_rd.max_row + 1):
                c2 = sheet_rd.cell(row=row, column=2).value
                if c2 and isinstance(c2, str) and ("合计" in c2 or "总计" in c2):
                    self._safe_set(
                        sheet_rd, row, 6, hightech_result.input.rd_expense_last_year
                    )
                    break

        # 研发费用结构明细表
        sheet_struct = None
        for name in wb.sheetnames:
            if "结构明细" in name and "汇总" not in name:
                sheet_struct = wb[name]
                break
        if sheet_struct is not None:
            rd_in = hightech_result.input
            self._safe_set(sheet_struct, 4, 2, rd_in.rd_expense_3year_total)
            self._safe_set(sheet_struct, 14, 2, rd_in.rd_expense_3year_total)
            self._safe_set(sheet_struct, 15, 2, rd_in.total_revenue_3year)
            if rd_in.total_revenue_3year > 0:
                ratio = rd_in.rd_expense_3year_total / rd_in.total_revenue_3year
                self._safe_set(sheet_struct, 16, 2, ratio)

        # 收入审定表
        sheet_inc = None
        for name in wb.sheetnames:
            if "收入" in name and "审定" in name:
                sheet_inc = wb[name]
                break
        if sheet_inc is not None:
            for row in range(1, sheet_inc.max_row + 1):
                c2 = sheet_inc.cell(row=row, column=2).value
                if c2 and isinstance(c2, str) and ("合计" in c2 or "总计" in c2):
                    self._safe_set(
                        sheet_inc, row, 6, hightech_result.input.hi_product_revenue
                    )
                    break

        wb.save(str(dst))
        return str(dst)


# ============================================================
# 快捷函数
# ============================================================


def fill_all_templates(
    result: CalculationResult,
    fuzhu_template: str,
    cover_template: str,
    a100000_template: str,
    a105000_template: str,
    output_dir: str = None,
) -> Dict[str, str]:
    """
    填充所有模板

    Returns
    -------
    dict
        {模板名称: 输出路径}
    """
    filler = TemplateFiller(result)
    base = output_dir or "D:/Users/12844/Desktop"

    outputs = {}

    if fuzhu_template and Path(fuzhu_template).exists():
        out = filler.fill_fuzhu_digao(
            fuzhu_template, f"{base}/税审底稿_辅助底稿_已填.xlsx"
        )
        outputs["辅助底稿"] = out

    if cover_template and Path(cover_template).exists():
        out = filler.fill_cover(cover_template, f"{base}/税审底稿_封面_已填.xlsx")
        outputs["封面"] = out

    if a100000_template and Path(a100000_template).exists():
        out = filler.fill_A100000(
            a100000_template, f"{base}/税审底稿_A100000_已填.xlsx"
        )
        outputs["A100000"] = out

    if a105000_template and Path(a105000_template).exists():
        out = filler.fill_A105000(
            a105000_template, f"{base}/税审底稿_A105000_已填.xlsx"
        )
        outputs["A105000"] = out

    return outputs
