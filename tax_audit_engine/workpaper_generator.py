"""
底稿生成器 — 将计算结果输出为结构化Excel工作底稿
不复制原模板格式，按审计底稿逻辑生成清晰表格
"""

from pathlib import Path
from typing import List, Optional
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

from .models import (
    CalculationResult,
    TaxAdjustment,
    AdjustmentCategory,
    AssetItem,
    EnterpriseInfo,
    TrialBalance,
)


# ============================================================
# 样式定义
# ============================================================


class Styles:
    """Excel样式常量"""

    # 颜色
    HEADER_FILL = PatternFill(
        start_color="4472C4", end_color="4472C4", fill_type="solid"
    )
    HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    TITLE_FONT = Font(name="微软雅黑", size=14, bold=True)
    SUBTITLE_FONT = Font(name="微软雅黑", size=10, bold=True)
    NORMAL_FONT = Font(name="微软雅黑", size=10)
    BOLD_FONT = Font(name="微软雅黑", size=10, bold=True)
    SMALL_FONT = Font(name="微软雅黑", size=9, color="666666")

    # 填充
    LIGHT_BLUE_FILL = PatternFill(
        start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"
    )
    LIGHT_GRAY_FILL = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    LIGHT_YELLOW_FILL = PatternFill(
        start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
    )
    WHITE_FILL = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )

    # 边框
    THIN_BORDER = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    BOTTOM_BORDER = Border(
        bottom=Side(style="medium"),
    )

    # 对齐
    CENTER = Alignment(horizontal="center", vertical="center")
    LEFT = Alignment(horizontal="left", vertical="center")
    RIGHT = Alignment(horizontal="right", vertical="center")
    WRAP = Alignment(wrap_text=True, vertical="top")

    # 数字格式
    AMOUNT_FORMAT = "#,##0.00"
    INT_FORMAT = "#,##0"
    PCT_FORMAT = "0.00%"


class WorkpaperGenerator:
    """
    税审工作底稿生成器
    基于计算结果，生成结构化Excel底稿
    """

    def __init__(self, result: CalculationResult):
        self.result = result
        self.wb = Workbook()
        self._sheet_index = 0

    def generate(self, output_path: str):
        """生成完整工作底稿Excel"""
        # 1. 封面
        self._create_cover_sheet()

        # 2. 利润表审定表
        self._create_pl_audit_sheet()

        # 3. 纳税调整明细表（收入类）
        if self._has_adjustments(AdjustmentCategory.INCOME):
            self._create_adjustment_sheet(
                "3-01 收入类纳税调整明细表",
                AdjustmentCategory.INCOME,
                "收入类纳税调整项目",
            )

        # 4. 纳税调整明细表（扣除类）
        if self._has_adjustments(AdjustmentCategory.DEDUCTION):
            self._create_adjustment_sheet(
                "3-02 扣除类纳税调整明细表",
                AdjustmentCategory.DEDUCTION,
                "扣除类纳税调整项目",
            )

        # 5. 纳税调整明细表（资产类）
        if self._has_adjustments(AdjustmentCategory.ASSET):
            self._create_adjustment_sheet(
                "3-03 资产类纳税调整明细表",
                AdjustmentCategory.ASSET,
                "资产类纳税调整项目",
            )

        # 6. 税收优惠明细表
        if self._has_adjustments(AdjustmentCategory.TAX_INCENTIVE):
            self._create_adjustment_sheet(
                "5 税收优惠明细表", AdjustmentCategory.TAX_INCENTIVE, "税收优惠明细项目"
            )

        # 7. 资产折旧测算表（如有资产数据）
        if (
            self.result.enterprise
            and hasattr(self.result, "_assets")
            and self.result._assets
        ):
            self._create_depreciation_sheet()

        # 8. 纳税调整汇总表
        self._create_summary_sheet()

        # 9. 应纳税所得额计算表
        self._create_tax_calculation_sheet()

        # 保存
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(output_path)

    # ============================================================
    # 封面
    # ============================================================

    def _create_cover_sheet(self):
        """生成封面sheet"""
        ws = self.wb.active
        ws.title = "封面"
        self._set_column_widths(ws, [3, 20, 30, 20, 3])
        ent = self.result.enterprise

        row = 2
        ws.merge_cells("B2:D2")
        ws["B2"] = "企业所得税汇算清缴纳税申报审核工作底稿"
        ws["B2"].font = Styles.TITLE_FONT
        ws["B2"].alignment = Styles.CENTER

        row = 4
        info_items = [
            ("被审核单位名称", ent.name),
            ("统一社会信用代码", ent.uscc),
            ("所属行业", ent.industry),
            ("纳税年度", f"{ent.tax_year}年度"),
            ("法定代表人", ent.legal_rep),
            ("注册地址", ent.address),
            ("从业人数", f"{ent.employee_count}人"),
            ("资产总额", f"{ent.total_assets:,.2f}元"),
            ("申报类型", "独立纳税企业"),
            ("编制日期", ""),
        ]
        for label, value in info_items:
            ws[f"B{row}"] = label
            ws[f"B{row}"].font = Styles.BOLD_FONT
            ws[f"C{row}"] = value
            ws[f"C{row}"].font = Styles.NORMAL_FONT
            ws[f"B{row}"].alignment = Styles.LEFT
            row += 1

        row += 2
        ws[f"B{row}"] = "计算结果摘要"
        ws[f"B{row}"].font = Styles.SUBTITLE_FONT

        row += 1
        summary = self.result.summary()
        for label, value in summary.items():
            ws[f"B{row}"] = label
            ws[f"B{row}"].font = Styles.BOLD_FONT
            ws[f"C{row}"] = value
            ws[f"C{row}"].font = Styles.NORMAL_FONT
            if isinstance(value, float):
                ws[f"C{row}"].number_format = Styles.AMOUNT_FORMAT
            row += 1

    # ============================================================
    # 利润表审定表
    # ============================================================

    def _create_pl_audit_sheet(self):
        """生成利润表审定表"""
        ws = self.wb.create_sheet("2-00 利润表审定表")
        self._set_column_widths(ws, [5, 25, 15, 15, 15, 15])

        tb = self.result.tb
        ent = self.result.enterprise

        # 标题
        ws.merge_cells("A1:E1")
        ws["A1"] = "利润表及纳税申报审核表"
        ws["A1"].font = Styles.TITLE_FONT
        ws["A1"].alignment = Styles.CENTER

        # 表头信息
        ws["A2"] = f"被审核单位: {ent.name}" if ent.name else ""
        ws["A2"].font = Styles.SMALL_FONT
        ws["D2"] = f"所属年度: {ent.tax_year}年度" if ent.tax_year else ""
        ws["D2"].font = Styles.SMALL_FONT

        # 列头
        headers = ["行次", "项目", "账载金额", "税收金额", "调增金额", "调减金额"]
        row = 4
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Styles.HEADER_FONT
            cell.fill = Styles.HEADER_FILL
            cell.alignment = Styles.CENTER
            cell.border = Styles.THIN_BORDER

        # 利润表项目（按申报表格式）
        pl_items = [
            ("一、营业收入", tb.revenue_total, None),
            ("  主营业务收入", tb.revenue_main, None),
            ("  其他业务收入", tb.revenue_other, None),
            ("减：营业成本", tb.cost_main + tb.cost_other, None),
            ("  主营业务成本", tb.cost_main, None),
            ("  其他业务成本", tb.cost_other, None),
            ("  税金及附加", tb.tax_surcharge, None),
            ("  销售费用", tb.selling_expense, None),
            ("  管理费用", tb.admin_expense, None),
            ("  财务费用", tb.finance_expense, None),
            ("  资产减值损失", tb.asset_impairment, None),
            ("加：公允价值变动收益", tb.fair_value_change, None),
            ("  投资收益", tb.investment_income, None),
            ("  资产处置收益", tb.asset_disposal_income, None),
            ("  其他收益", tb.get("其他收益", 0), None),
            ("  营业外收入", tb.non_operating_income, None),
            ("减：营业外支出", tb.get("营业外支出", 0), None),
            ("二、利润总额", tb.accounting_profit, None),
        ]

        row = 5
        for i, (name, amount, _) in enumerate(pl_items, 1):
            is_total = name.startswith("一") or name.startswith("二")
            ws.cell(row=row, column=1, value=i).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=2, value=name).font = (
                Styles.BOLD_FONT if is_total else Styles.NORMAL_FONT
            )
            ws.cell(row=row, column=3, value=amount).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=4).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=5).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=6).font = Styles.NORMAL_FONT
            for col in range(1, 7):
                ws.cell(row=row, column=col).border = Styles.THIN_BORDER
                if col >= 3:
                    ws.cell(row=row, column=col).number_format = Styles.AMOUNT_FORMAT
                    ws.cell(row=row, column=col).alignment = Styles.RIGHT
            if is_total:
                for col in range(1, 7):
                    ws.cell(row=row, column=col).fill = Styles.LIGHT_BLUE_FILL
            row += 1

    # ============================================================
    # 纳税调整明细表（通用模板）
    # ============================================================

    def _create_adjustment_sheet(
        self, title: str, category: AdjustmentCategory, subtitle: str
    ):
        """生成某个类别的纳税调整明细表"""
        ws = self.wb.create_sheet(title[:31])  # sheet名最多31字符
        self._set_column_widths(ws, [5, 6, 28, 18, 18, 18, 18, 40])

        adjustments = self._get_adjustments(category)

        # 标题
        ws.merge_cells("A1:G1")
        ws["A1"] = title
        ws["A1"].font = Styles.TITLE_FONT
        ws["A1"].alignment = Styles.CENTER

        ws["A2"] = f"{subtitle} — 税法依据及计算过程"
        ws["A2"].font = Styles.SMALL_FONT

        # 列头
        headers = [
            "行次",
            "类别",
            "项目",
            "账载金额",
            "计税基础",
            "纳税调增",
            "纳税调减",
            "税法依据/计算说明",
        ]
        row = 4
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Styles.HEADER_FONT
            cell.fill = Styles.HEADER_FILL
            cell.alignment = Styles.CENTER
            cell.border = Styles.THIN_BORDER

        # 数据行
        row = 5
        for i, adj in enumerate(adjustments, 1):
            ws.cell(row=row, column=1, value=i).font = Styles.NORMAL_FONT
            ws.cell(
                row=row, column=2, value=adj.category.value
            ).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=3, value=adj.item_name).font = Styles.NORMAL_FONT

            # 数值列
            amt_cols = [4, 5, 6, 7]
            vals = [adj.book_amount, adj.tax_base, adj.increase, adj.decrease]
            for c, v in zip(amt_cols, vals):
                cell = ws.cell(row=row, column=c, value=v)
                cell.font = Styles.NORMAL_FONT
                cell.number_format = Styles.AMOUNT_FORMAT
                cell.alignment = Styles.RIGHT

            # 税法依据
            ref = adj.tax_law_ref
            calc = adj.calculation
            remark = adj.remark
            note_parts = [ref]
            if calc:
                note_parts.append(calc)
            if remark:
                note_parts.append(remark)

            ws.cell(row=row, column=8, value="\n".join(note_parts))
            ws.cell(row=row, column=8).font = Styles.SMALL_FONT
            ws.cell(row=row, column=8).alignment = Styles.WRAP

            # 边框 + 斑马纹
            fill = Styles.WHITE_FILL if i % 2 == 1 else Styles.LIGHT_GRAY_FILL
            for col in range(1, 9):
                ws.cell(row=row, column=col).border = Styles.THIN_BORDER
                ws.cell(row=row, column=col).fill = fill

            row += 1

        # 合计行
        ws.cell(row=row, column=1).font = Styles.BOLD_FONT
        ws.cell(row=row, column=3, value="合计").font = Styles.BOLD_FONT
        total_book = sum(a.book_amount for a in adjustments)
        total_increase = sum(a.increase for a in adjustments)
        total_decrease = sum(a.decrease for a in adjustments)
        ws.cell(row=row, column=4, value=total_book).font = Styles.BOLD_FONT
        ws.cell(row=row, column=4).number_format = Styles.AMOUNT_FORMAT
        ws.cell(row=row, column=5).font = Styles.BOLD_FONT
        ws.cell(row=row, column=5).number_format = Styles.AMOUNT_FORMAT
        ws.cell(row=row, column=6, value=total_increase).font = Styles.BOLD_FONT
        ws.cell(row=row, column=6).number_format = Styles.AMOUNT_FORMAT
        ws.cell(row=row, column=6).fill = Styles.LIGHT_YELLOW_FILL
        ws.cell(row=row, column=7, value=total_decrease).font = Styles.BOLD_FONT
        ws.cell(row=row, column=7).number_format = Styles.AMOUNT_FORMAT
        ws.cell(row=row, column=7).fill = Styles.LIGHT_YELLOW_FILL

        for col in range(1, 9):
            ws.cell(row=row, column=col).border = Styles.THIN_BORDER

    # ============================================================
    # 资产折旧测算表
    # ============================================================

    def _create_depreciation_sheet(self):
        """生成资产折旧测算表"""
        ws = self.wb.create_sheet("3-03-01 固定资产折旧测算")
        self._set_column_widths(ws, [5, 15, 12, 12, 12, 12, 12, 12, 12])

        assets = self.result._assets  # 从计算器传入

        ws.merge_cells("A1:H1")
        ws["A1"] = "固定资产折旧及纳税调整审核表"
        ws["A1"].font = Styles.TITLE_FONT
        ws["A1"].alignment = Styles.CENTER

        headers = [
            "行次",
            "资产类别",
            "资产原值",
            "会计折旧",
            "税收折旧",
            "差异(税收-会计)",
            "税法年限",
            "加速折旧",
        ]
        row = 3
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Styles.HEADER_FONT
            cell.fill = Styles.HEADER_FILL
            cell.alignment = Styles.CENTER
            cell.border = Styles.THIN_BORDER

        row = 4
        total_orig = total_acct = total_tax = total_diff = 0
        for i, asset in enumerate(assets, 1):
            diff = asset.current_tax_depr - asset.current_accounting_depr
            total_orig += asset.original_value
            total_acct += asset.current_accounting_depr
            total_tax += asset.current_tax_depr
            total_diff += diff

            ws.cell(row=row, column=1, value=i)
            ws.cell(row=row, column=2, value=f"{asset.category}—{asset.name}")
            ws.cell(row=row, column=3, value=asset.original_value)
            ws.cell(row=row, column=4, value=asset.current_accounting_depr)
            ws.cell(row=row, column=5, value=asset.current_tax_depr)
            ws.cell(row=row, column=6, value=diff)
            ws.cell(
                row=row,
                column=7,
                value=f"会计{asset.accounting_life_years}年/税法{asset.tax_life_years}年",
            )
            ws.cell(row=row, column=8, value="是" if asset.is_accelerated else "否")

            for col in range(1, 9):
                ws.cell(row=row, column=col).font = Styles.NORMAL_FONT
                ws.cell(row=row, column=col).border = Styles.THIN_BORDER
                if col in (3, 4, 5, 6):
                    ws.cell(row=row, column=col).number_format = Styles.AMOUNT_FORMAT
                    ws.cell(row=row, column=col).alignment = Styles.RIGHT
            row += 1

        # 合计
        ws.cell(row=row, column=1).font = Styles.BOLD_FONT
        ws.cell(row=row, column=2, value="合计").font = Styles.BOLD_FONT
        ws.cell(row=row, column=3, value=total_orig).font = Styles.BOLD_FONT
        ws.cell(row=row, column=3).number_format = Styles.AMOUNT_FORMAT
        ws.cell(row=row, column=4, value=total_acct).font = Styles.BOLD_FONT
        ws.cell(row=row, column=4).number_format = Styles.AMOUNT_FORMAT
        ws.cell(row=row, column=5, value=total_tax).font = Styles.BOLD_FONT
        ws.cell(row=row, column=5).number_format = Styles.AMOUNT_FORMAT
        ws.cell(row=row, column=6, value=total_diff).font = Styles.BOLD_FONT
        ws.cell(row=row, column=6).number_format = Styles.AMOUNT_FORMAT
        for col in range(1, 9):
            ws.cell(row=row, column=col).border = Styles.THIN_BORDER
            ws.cell(row=row, column=col).fill = Styles.LIGHT_BLUE_FILL

    # ============================================================
    # 纳税调整汇总表
    # ============================================================

    def _create_summary_sheet(self):
        """生成纳税调整汇总表"""
        ws = self.wb.create_sheet("纳税调整汇总表")
        self._set_column_widths(ws, [5, 30, 18, 18, 18])

        ws.merge_cells("A1:D1")
        ws["A1"] = "企业所得税纳税调整项目汇总表"
        ws["A1"].font = Styles.TITLE_FONT
        ws["A1"].alignment = Styles.CENTER

        headers = ["行次", "调整类别", "纳税调增金额", "纳税调减金额", "调整项数"]
        row = 3
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Styles.HEADER_FONT
            cell.fill = Styles.HEADER_FILL
            cell.alignment = Styles.CENTER
            cell.border = Styles.THIN_BORDER

        categories = [
            (AdjustmentCategory.INCOME, "一、收入类调整项目"),
            (AdjustmentCategory.DEDUCTION, "二、扣除类调整项目"),
            (AdjustmentCategory.ASSET, "三、资产类调整项目"),
            (AdjustmentCategory.TAX_INCENTIVE, "四、税收优惠调整项目"),
        ]

        row = 4
        grand_inc = grand_dec = grand_count = 0
        for cat, label in categories:
            adjustments = self._get_adjustments(cat)
            inc = sum(a.increase for a in adjustments)
            dec = sum(a.decrease for a in adjustments)
            cnt = len(adjustments)
            grand_inc += inc
            grand_dec += dec
            grand_count += cnt

            ws.cell(row=row, column=1).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=2, value=label).font = Styles.BOLD_FONT
            ws.cell(row=row, column=3, value=inc).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=3).number_format = Styles.AMOUNT_FORMAT
            ws.cell(row=row, column=4, value=dec).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=4).number_format = Styles.AMOUNT_FORMAT
            ws.cell(row=row, column=5, value=cnt).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=5).alignment = Styles.CENTER
            for col in range(1, 6):
                ws.cell(row=row, column=col).border = Styles.THIN_BORDER
            row += 1

        # 合计行
        ws.cell(row=row, column=2, value="合  计").font = Styles.BOLD_FONT
        ws.cell(row=row, column=3, value=grand_inc).font = Styles.BOLD_FONT
        ws.cell(row=row, column=3).number_format = Styles.AMOUNT_FORMAT
        ws.cell(row=row, column=3).fill = Styles.LIGHT_YELLOW_FILL
        ws.cell(row=row, column=4, value=grand_dec).font = Styles.BOLD_FONT
        ws.cell(row=row, column=4).number_format = Styles.AMOUNT_FORMAT
        ws.cell(row=row, column=4).fill = Styles.LIGHT_YELLOW_FILL
        ws.cell(row=row, column=5, value=grand_count).font = Styles.BOLD_FONT
        ws.cell(row=row, column=5).alignment = Styles.CENTER
        for col in range(1, 6):
            ws.cell(row=row, column=col).border = Styles.THIN_BORDER
            ws.cell(row=row, column=col).fill = Styles.LIGHT_BLUE_FILL

    # ============================================================
    # 应纳税所得额计算表
    # ============================================================

    def _create_tax_calculation_sheet(self):
        """生成应纳税所得额及税额计算表"""
        ws = self.wb.create_sheet("应纳税所得额计算")
        self._set_column_widths(ws, [5, 35, 20, 20])

        ws.merge_cells("A1:C1")
        ws["A1"] = "应纳税所得额及应纳所得税额计算表"
        ws["A1"].font = Styles.TITLE_FONT
        ws["A1"].alignment = Styles.CENTER

        r = self.result

        items = [
            ("一、会计利润总额", r.accounting_profit, None),
            ("加：纳税调增金额合计", r.total_increase, None),
            ("减：纳税调减金额合计", r.total_decrease, None),
            ("二、应纳税所得额", r.taxable_income, "会计利润+调增-调减"),
            ("三、适用税率", r.tax_rate, "详见税率说明"),
            ("四、应纳所得税额", r.tax_payable, "应纳税所得额×税率"),
            ("减：减免所得税额", r.deducted_tax, None),
            ("五、实际应纳所得税额", r.final_tax, None),
        ]

        headers = ["行次", "项目", "金额", "说明"]
        row = 3
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Styles.HEADER_FONT
            cell.fill = Styles.HEADER_FILL
            cell.alignment = Styles.CENTER
            cell.border = Styles.THIN_BORDER

        row = 4
        for i, (label, amount, note) in enumerate(items, 1):
            is_result = label.startswith("二") or label.startswith("五")
            ws.cell(row=row, column=1, value=i).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=2, value=label).font = (
                Styles.BOLD_FONT if is_result else Styles.NORMAL_FONT
            )
            if isinstance(amount, float):
                ws.cell(row=row, column=3, value=amount).font = (
                    Styles.BOLD_FONT if is_result else Styles.NORMAL_FONT
                )
                ws.cell(row=row, column=3).number_format = Styles.AMOUNT_FORMAT
                ws.cell(row=row, column=3).alignment = Styles.RIGHT
            else:
                ws.cell(row=row, column=3, value=amount).font = Styles.NORMAL_FONT
            ws.cell(row=row, column=4, value=note or "").font = Styles.SMALL_FONT
            for col in range(1, 5):
                ws.cell(row=row, column=col).border = Styles.THIN_BORDER
            if is_result:
                for col in range(1, 5):
                    ws.cell(row=row, column=col).fill = Styles.LIGHT_YELLOW_FILL
            row += 1

    # ============================================================
    # 辅助方法
    # ============================================================

    def _has_adjustments(self, category: AdjustmentCategory) -> bool:
        """检查某类别是否有调整项"""
        return len(self._get_adjustments(category)) > 0

    def _get_adjustments(self, category: AdjustmentCategory) -> List[TaxAdjustment]:
        """获取某类别的调整项（含0值但非None的）"""
        return [a for a in self.result.adjustments if a.category == category]

    @staticmethod
    def _set_column_widths(ws, widths: List[int]):
        """设置列宽"""
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def set_assets(self, assets: List[AssetItem]):
        """设置资产数据（用于折旧测算表）"""
        self.result._assets = assets
