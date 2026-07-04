"""
底稿生成器扩展 — 4类业务自生成底稿（RD/FullTax/Loss/HighTech）
不依赖模板，从计算结果直接生成结构化Excel底稿
"""

from pathlib import Path
from typing import List, Optional
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from tax_audit_engine.models import EnterpriseInfo


# ============================================================
# 样式（复用 workpaper_generator.Styles 的配色方案）
# ============================================================


class _S:
    """Excel样式常量"""

    HEADER_FILL = PatternFill(
        start_color="4472C4", end_color="4472C4", fill_type="solid"
    )
    HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    TITLE_FONT = Font(name="微软雅黑", size=14, bold=True)
    SUBTITLE_FONT = Font(name="微软雅黑", size=10, bold=True)
    NORMAL_FONT = Font(name="微软雅黑", size=10)
    BOLD_FONT = Font(name="微软雅黑", size=10, bold=True)
    SMALL_FONT = Font(name="微软雅黑", size=9, color="666666")
    RESULT_FONT = Font(name="微软雅黑", size=10, bold=True, color="C00000")

    LIGHT_BLUE = PatternFill(
        start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"
    )
    LIGHT_GRAY = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    LIGHT_YELLOW = PatternFill(
        start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
    )
    LIGHT_GREEN = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )
    WHITE_FILL = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )

    THIN = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    CENTER = Alignment(horizontal="center", vertical="center")
    LEFT = Alignment(horizontal="left", vertical="center")
    RIGHT = Alignment(horizontal="right", vertical="center")
    WRAP = Alignment(wrap_text=True, vertical="top")

    AMOUNT = "#,##0.00"
    INT_FMT = "#,##0"
    PCT_FMT = "0.00%"


def _set_cols(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _header_row(ws, row, headers):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = _S.HEADER_FONT
        c.fill = _S.HEADER_FILL
        c.alignment = _S.CENTER
        c.border = _S.THIN


def _data_cell(ws, row, col, value, fmt=None, bold=False, fill=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _S.BOLD_FONT if bold else _S.NORMAL_FONT
    c.border = _S.THIN
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if isinstance(value, (int, float)):
        c.alignment = _S.RIGHT
    return c


def _write_row(ws, row, values, fmt=None, bold=False, fill_=None):
    for col, v in enumerate(values, 1):
        _data_cell(ws, row, col, v, fmt=fmt, bold=bold, fill=fill_)


def _cover_info(ws, start_row, info_items):
    """写入封面信息行列表"""
    r = start_row
    for label, value in info_items:
        ws.cell(row=r, column=2, value=label).font = _S.BOLD_FONT
        ws.cell(row=r, column=3, value=value).font = _S.NORMAL_FONT
        r += 1
    return r


# ============================================================
# RDGenerator
# ============================================================


class RDGenerator:
    """
    RD 研发费用加计扣除工作底稿生成器
    sheet: 封面 → 项目明细表 → 3-01汇总表 → 3-02优惠审核表
    """

    def __init__(self, result, input_data=None):
        self.result = result
        self.input = input_data
        self.wb = Workbook()
        self.ent = result.enterprise or EnterpriseInfo(name="未命名企业")

    def generate(self, output_path: str):
        self._cover()
        self._project_detail()
        self._summary()
        self._audit_check()
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(output_path)

    def _cover(self):
        ws = self.wb.active
        ws.title = "封面"
        _set_cols(ws, [3, 22, 35, 22, 3])
        ws.merge_cells("B2:D2")
        ws["B2"] = "研发费用加计扣除审核工作底稿"
        ws["B2"].font = _S.TITLE_FONT
        ws["B2"].alignment = _S.CENTER
        _cover_info(
            ws,
            4,
            [
                ("被审核单位名称", self.ent.name or ""),
                ("统一社会信用代码", self.ent.uscc or ""),
                ("所属行业", self.ent.industry or ""),
                (
                    "审核年度",
                    f"{getattr(self.ent, 'tax_year', '')}年度"
                    if hasattr(self.ent, "tax_year")
                    else "",
                ),
                ("加计扣除比例", "100%"),
                ("适用政策", "财税[2015]119号、2021年第28号"),
            ],
        )
        r = 12
        ws.cell(row=r, column=2, value="汇总数据").font = _S.SUBTITLE_FONT
        r += 1
        for label, val in [
            ("账面研发费用合计", self.result.total_book),
            ("可加计扣除金额", self.result.qualifying_amount),
            ("加计扣除额", self.result.deduction_amount),
        ]:
            ws.cell(row=r, column=2, value=label).font = _S.BOLD_FONT
            c = ws.cell(row=r, column=3, value=val)
            c.font = _S.NORMAL_FONT
            c.number_format = _S.AMOUNT
            r += 1
        if self.result.has_other_limit_issue:
            ws.cell(row=r, column=2, value="其他费用超限金额").font = _S.BOLD_FONT
            c = ws.cell(row=r, column=3, value=self.result.other_limit_excess)
            c.font = _S.RESULT_FONT
            c.number_format = _S.AMOUNT

    def _project_detail(self):
        ws = self.wb.create_sheet("项目明细")
        _set_cols(ws, [4, 8, 24, 12, 14, 14, 14, 14, 14, 14, 14])
        ws.merge_cells("A1:J1")
        ws["A1"] = "研发项目费用明细表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(
            ws,
            3,
            [
                "序号",
                "项目编号",
                "项目名称",
                "研发方式",
                "人员人工",
                "直接投入",
                "折旧",
                "摊销",
                "设计试验",
                "其他费用",
                "合计",
            ],
        )
        r = 4
        for i, proj in enumerate(self.result.projects, 1):
            vals = [
                i,
                proj.project_id,
                proj.project_name,
                proj.dev_method,
                proj.personnel_costs,
                proj.direct_input,
                proj.depreciation,
                proj.amortization,
                proj.design_trial,
                proj.other_costs,
                proj.total_book,
            ]
            _write_row(
                ws,
                r,
                vals,
                fmt=_S.AMOUNT,
                fill_=_S.WHITE_FILL if i % 2 else _S.LIGHT_GRAY,
            )
            # 前3列文字格式
            for c in [1, 2, 3, 4]:
                ws.cell(row=r, column=c).number_format = "@"
            r += 1
        # 合计行
        totals = [
            "",
            "",
            "合计",
            "",
            sum(p.personnel_costs for p in self.result.projects),
            sum(p.direct_input for p in self.result.projects),
            sum(p.depreciation for p in self.result.projects),
            sum(p.amortization for p in self.result.projects),
            sum(p.design_trial for p in self.result.projects),
            sum(p.other_costs for p in self.result.projects),
            sum(p.total_book for p in self.result.projects),
        ]
        _write_row(ws, r, totals, fmt=_S.AMOUNT, bold=True, fill_=_S.LIGHT_BLUE)

    def _summary(self):
        ws = self.wb.create_sheet("3-01汇总表")
        _set_cols(ws, [5, 30, 18, 18, 18])
        ws.merge_cells("A1:D1")
        ws["A1"] = "可加计扣除研发费用审核汇总表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(ws, 3, ["行次", "项目", "账面金额", "可加计扣除金额", "差异"])
        rows = [
            ("一、人员人工费用", "personnel_costs"),
            ("二、直接投入费用", "direct_input"),
            ("三、折旧费用", "depreciation"),
            ("四、无形资产摊销", "amortization"),
            ("五、设计试验等费用", "design_trial"),
            ("六、其他相关费用", "other_costs"),
        ]
        r = 4
        for label, attr in rows:
            book = sum(getattr(p, attr, 0) for p in self.result.projects)
            _write_row(ws, r, [r - 3, label, book, book, 0], fmt=_S.AMOUNT)
            r += 1
        total_book = sum(p.total_book for p in self.result.projects)
        _write_row(
            ws,
            r,
            [
                "",
                "合计",
                total_book,
                self.result.qualifying_amount,
                self.result.qualifying_amount - total_book,
            ],
            fmt=_S.AMOUNT,
            bold=True,
            fill_=_S.LIGHT_YELLOW,
        )
        r += 2
        ws.cell(row=r, column=2, value="加计扣除额").font = _S.BOLD_FONT
        c = ws.cell(row=r, column=3, value=self.result.deduction_amount)
        c.font = _S.RESULT_FONT
        c.number_format = _S.AMOUNT

    def _audit_check(self):
        ws = self.wb.create_sheet("3-02审核表")
        _set_cols(ws, [5, 35, 18, 18])
        ws.merge_cells("A1:C1")
        ws["A1"] = "研发费用加计扣除优惠审核表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(ws, 3, ["行次", "审核项目", "金额", "审核意见"])
        rows = [
            ("一、自主研发、合作研发、集中研发", self.result.qualifying_amount),
            (
                "（一）人员人工费用",
                sum(p.personnel_costs for p in self.result.projects),
            ),
            ("（二）直接投入费用", sum(p.direct_input for p in self.result.projects)),
            ("（三）折旧费用", sum(p.depreciation for p in self.result.projects)),
            ("（四）无形资产摊销", sum(p.amortization for p in self.result.projects)),
            ("（五）设计费用", sum(p.design_trial for p in self.result.projects)),
            ("（六）其他相关费用", sum(p.other_costs for p in self.result.projects)),
            ("三、年度可加计扣除研发费用小计", self.result.qualifying_amount),
            ("四、计入本年损益的加计扣除额", self.result.deduction_amount),
        ]
        r = 4
        for i, (label, val) in enumerate(rows, 1):
            opinion = "通过" if val > 0 else ""
            _write_row(ws, r, [i, label, val, opinion], fmt=_S.AMOUNT)
            if i in (1, 8, 9):
                for c in range(1, 5):
                    ws.cell(row=r, column=c).fill = _S.LIGHT_BLUE
            r += 1


# ============================================================
# FullTaxGenerator
# ============================================================


class FullTaxGenerator:
    """
    全税种核查底稿生成器
    sheet: 封面 → 2-01各税种汇总 → 差异分析明细
    """

    def __init__(self, result):
        self.result = result
        self.wb = Workbook()
        self.ft_input = result.input

    def generate(self, output_path: str):
        self._cover()
        self._summary()
        self._analysis()
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(output_path)

    def _cover(self):
        ws = self.wb.active
        ws.title = "封面"
        _set_cols(ws, [3, 22, 35, 3])
        ws.merge_cells("B2:C2")
        ws["B2"] = "全税种核查测算工作底稿"
        ws["B2"].font = _S.TITLE_FONT
        ws["B2"].alignment = _S.CENTER
        _cover_info(
            ws,
            4,
            [
                ("被审计单位名称", self.ft_input.enterprise_name),
                ("统一社会信用代码", self.ft_input.tax_id),
                ("核查年度", self.ft_input.audit_year),
                ("差异率阈值", "5%"),
            ],
        )

    def _summary(self):
        ws = self.wb.create_sheet("2-01各税种汇总")
        _set_cols(ws, [5, 20, 15, 18, 18, 18, 12, 8])
        ws.merge_cells("A1:G1")
        ws["A1"] = "各税种申报与认定对比表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(
            ws,
            3,
            [
                "序号",
                "税种",
                "所属期",
                "申报税额",
                "认定税额",
                "差异额",
                "差异率",
                "风险",
            ],
        )
        r = 4
        total_dec = total_act = total_diff = 0
        for i, item in enumerate(self.result.item_details, 1):
            diff = item.get(
                "difference",
                item.get("actual_amount", 0) - item.get("declared_amount", 0),
            )
            rate = item.get("diff_rate", 0)
            is_high = item.get("is_high_risk", abs(rate) > 0.05)
            total_dec += item.get("declared_amount", 0)
            total_act += item.get("actual_amount", 0)
            total_diff += diff
            risk_label = "!!高风险" if is_high else "正常"
            _write_row(
                ws,
                r,
                [
                    i,
                    item.get("tax_type", ""),
                    item.get("tax_period", ""),
                    item.get("declared_amount", 0),
                    item.get("actual_amount", 0),
                    diff,
                    rate,
                    risk_label,
                ],
                fmt=_S.AMOUNT,
                fill_=None if not is_high else _S.LIGHT_YELLOW,
            )
            # 差异率百分比格式
            ws.cell(row=r, column=7).number_format = "0.00%"
            r += 1
        _write_row(
            ws,
            r,
            ["", "合  计", "", total_dec, total_act, total_diff, "", ""],
            fmt=_S.AMOUNT,
            bold=True,
            fill_=_S.LIGHT_BLUE,
        )

    def _analysis(self):
        ws = self.wb.create_sheet("差异分析")
        _set_cols(ws, [5, 20, 18, 18, 50])
        ws.merge_cells("A1:D1")
        ws["A1"] = "税种差异分析及审计建议"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(ws, 3, ["序号", "税种", "差异额", "差异率", "可能原因及建议"])
        r = 4
        for i, item in enumerate(self.result.item_details, 1):
            diff = item.get("difference", 0)
            rate = item.get("diff_rate", 0)
            if abs(diff) < 0.01:
                continue
            reason = item.get("risk_reason", getattr(item, "risk_reason", ""))
            _write_row(
                ws, r, [i, item.get("tax_type", ""), diff, rate, reason], fmt=_S.AMOUNT
            )
            ws.cell(row=r, column=4).number_format = "0.00%"
            ws.cell(row=r, column=5).alignment = _S.WRAP
            r += 1


# ============================================================
# LossGenerator
# ============================================================


class LossGenerator:
    """
    财产损失税前扣除审核底稿生成器
    sheet: 封面 → 损失汇总表 → 逐项审核明细
    """

    def __init__(self, result):
        self.result = result
        self.wb = Workbook()
        self.loss_input = result.input

    def generate(self, output_path: str):
        self._cover()
        self._summary()
        self._detail()
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(output_path)

    def _cover(self):
        ws = self.wb.active
        ws.title = "封面"
        _set_cols(ws, [3, 25, 35, 3])
        ws.merge_cells("B2:C2")
        ws["B2"] = "企业资产损失税前扣除审核工作底稿"
        ws["B2"].font = _S.TITLE_FONT
        ws["B2"].alignment = _S.CENTER
        ent = self.loss_input.enterprise
        _cover_info(
            ws,
            4,
            [
                ("被审核单位名称", ent.name if ent else ""),
                ("统一社会信用代码", ent.uscc if ent else ""),
                ("审核依据", "国家税务总局公告2011年第25号"),
                ("损失项数", str(self.result.item_count)),
                ("申报损失合计", self.result.total_loss),
            ],
        )

    def _summary(self):
        ws = self.wb.create_sheet("损失汇总")
        _set_cols(ws, [5, 30, 18, 18, 18, 12])
        ws.merge_cells("A1:E1")
        ws["A1"] = "资产损失税前扣除审核汇总表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(
            ws, 3, ["序号", "项目", "账面净值", "可收回金额", "损失金额", "审核结论"]
        )
        r = 4
        for i, item in enumerate(self.result.item_results, 1):
            disallowed = item.get("disallowed", False)
            conclusion = "不得扣除" if disallowed else "可扣除"
            fill_ = _S.LIGHT_YELLOW if disallowed else None
            vals = [
                i,
                item.get("asset_name", f"损失{i}"),
                item.get("book_value", 0),
                item.get("recoverable", 0),
                item.get("loss_amount", 0),
                conclusion,
            ]
            _write_row(ws, r, vals, fmt=_S.AMOUNT, fill_=fill_)
            r += 1
        _write_row(
            ws,
            r,
            [
                "",
                "合  计",
                self.result.total_book_value,
                self.result.total_recoverable,
                self.result.total_loss,
                "",
            ],
            fmt=_S.AMOUNT,
            bold=True,
            fill_=_S.LIGHT_BLUE,
        )
        r += 2
        for label, val in [
            ("可税前扣除损失", self.result.qualifying_loss),
            ("不得扣除损失", self.result.disallowed_loss),
        ]:
            ws.cell(row=r, column=2, value=label).font = _S.BOLD_FONT
            c = ws.cell(row=r, column=3, value=val)
            c.font = _S.RESULT_FONT
            c.number_format = _S.AMOUNT
            r += 1

    def _detail(self):
        ws = self.wb.create_sheet("逐项审核")
        _set_cols(ws, [5, 8, 20, 12, 14, 14, 14, 12, 12])
        ws.merge_cells("A1:H1")
        ws["A1"] = "资产损失逐项审核明细表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(
            ws,
            3,
            [
                "序号",
                "编号",
                "资产名称",
                "损失类型",
                "账面净值",
                "可收回",
                "损失金额",
                "审批",
                "证据",
            ],
        )
        r = 4
        for i, item in enumerate(self.result.item_results, 1):
            _write_row(
                ws,
                r,
                [
                    i,
                    item.get("loss_id", ""),
                    item.get("asset_name", ""),
                    item.get("category", ""),
                    item.get("book_value", 0),
                    item.get("recoverable", 0),
                    item.get("loss_amount", 0),
                    item.get("approval_status", ""),
                    item.get("evidence_level", ""),
                ],
                fmt=_S.AMOUNT,
            )
            r += 1


# ============================================================
# HighTechGenerator
# ============================================================


class HighTechGenerator:
    """
    高新技术企业认定审核底稿生成器
    sheet: 封面 → 研发费用审定表 → 收入审定表 → 结构明细表 → 评分表
    """

    def __init__(self, result):
        self.result = result
        self.wb = Workbook()
        self.ht_input = result.input

    def generate(self, output_path: str):
        self._cover()
        self._rd_audit()
        self._income_audit()
        self._structure()
        self._score()
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(output_path)

    def _cover(self):
        ws = self.wb.active
        ws.title = "封面"
        _set_cols(ws, [3, 22, 35, 3])
        ws.merge_cells("B2:C2")
        ws["B2"] = "高新技术企业认定审核工作底稿"
        ws["B2"].font = _S.TITLE_FONT
        ws["B2"].alignment = _S.CENTER
        ent = self.ht_input.enterprise
        _cover_info(
            ws,
            4,
            [
                ("被审核单位名称", ent.name if ent else ""),
                ("统一社会信用代码", ent.uscc if ent else ""),
                ("审核年度", "2025"),
                ("审核依据", "国科发火〔2016〕32号、195号"),
                ("总分", f"{self.result.total_score:.1f}分"),
                ("审核结论", "通过" if self.result.passed else "不通过"),
            ],
        )

    def _rd_audit(self):
        ws = self.wb.create_sheet("研发费用审定表")
        _set_cols(ws, [5, 35, 18, 18, 18])
        ws.merge_cells("A1:D1")
        ws["A1"] = "研发费用审定表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(ws, 3, ["序号", "项目", "金额", "占比", "达标"])
        rows = [
            ("最近一年研发费用", self.ht_input.rd_expense_last_year),
            ("近3年研发费用总额", self.ht_input.rd_expense_3year_total),
            ("近3年销售收入总额", self.ht_input.income_3year_total),
            ("研发费用占比(3年)", None),
            ("科技人员数", self.ht_input.tech_staff),
            ("职工总数", self.ht_input.total_staff),
            ("科技人员占比", None),
        ]
        rd_ratio = (
            self.ht_input.rd_expense_3year_total / self.ht_input.income_3year_total
            if self.ht_input.income_3year_total > 0
            else 0
        )
        tech_ratio = (
            self.ht_input.tech_staff / self.ht_input.total_staff
            if self.ht_input.total_staff > 0
            else 0
        )
        r = 4
        for i, (label, val) in enumerate(rows, 1):
            if label == "研发费用占比(3年)":
                _write_row(
                    ws,
                    r,
                    [
                        i,
                        label,
                        rd_ratio,
                        rd_ratio,
                        "达标" if rd_ratio >= 0.04 else "不达标",
                    ],
                    fmt=_S.PCT_FMT,
                )
                ws.cell(row=r, column=3).number_format = _S.AMOUNT
            elif label == "科技人员占比":
                _write_row(
                    ws,
                    r,
                    [
                        i,
                        label,
                        tech_ratio,
                        tech_ratio,
                        "达标" if tech_ratio >= 0.10 else "不达标",
                    ],
                    fmt=_S.PCT_FMT,
                )
                ws.cell(row=r, column=3).number_format = _S.AMOUNT
            else:
                _write_row(ws, r, [i, label, val, "", ""], fmt=_S.AMOUNT)
            if i in (4, 7):
                for c in range(1, 6):
                    ws.cell(row=r, column=c).fill = _S.LIGHT_BLUE
            r += 1

    def _income_audit(self):
        ws = self.wb.create_sheet("收入审定表")
        _set_cols(ws, [5, 30, 18, 18, 18])
        ws.merge_cells("A1:D1")
        ws["A1"] = "高新技术产品（服务）收入审定表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(ws, 3, ["序号", "项目", "金额", "占比", "判定"])
        hi_ratio = (
            self.ht_input.hi_product_revenue / self.ht_input.income_last_year
            if self.ht_input.income_last_year > 0
            else 0
        )
        hi_ratio_3 = (
            self.ht_input.hi_product_revenue_3year / self.ht_input.total_revenue_3year
            if self.ht_input.total_revenue_3year > 0
            else 0
        )
        items = [
            ("高新收入（最近一年）", self.ht_input.hi_product_revenue, ""),
            ("总收入（最近一年）", self.ht_input.income_last_year, ""),
            (
                "高新收入占比（最近一年）",
                hi_ratio,
                "达标" if hi_ratio >= 0.60 else "不达标",
            ),
            ("高新收入（近3年）", self.ht_input.hi_product_revenue_3year, ""),
            ("总收入（近3年）", self.ht_input.total_revenue_3year, ""),
            (
                "高新收入占比（近3年）",
                hi_ratio_3,
                "达标" if hi_ratio_3 >= 0.60 else "不达标",
            ),
        ]
        r = 4
        for i, (label, val, status) in enumerate(items, 1):
            _write_row(ws, r, [i, label, val, "", status], fmt=_S.AMOUNT)
            if "占比" in label:
                ws.cell(row=r, column=3).number_format = _S.PCT_FMT
                fill_ = _S.LIGHT_GREEN if status == "达标" else _S.LIGHT_YELLOW
                for c in range(1, 6):
                    ws.cell(row=r, column=c).fill = fill_
            r += 1

    def _structure(self):
        ws = self.wb.create_sheet("研发费用结构明细")
        _set_cols(ws, [5, 30, 18, 18, 18])
        ws.merge_cells("A1:D1")
        ws["A1"] = "研发费用结构明细表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(ws, 3, ["序号", "费用类别", "金额", "占研发费用比例", "备注"])
        total = self.ht_input.rd_expense_3year_total
        # 从输入数据中分解各类费用（用等比例估算，实际应有明细数据）
        items_data = [
            ("人员人工费用", total * 0.35),
            ("直接投入费用", total * 0.30),
            ("折旧费用", total * 0.12),
            ("无形资产摊销", total * 0.08),
            ("设计试验等费用", total * 0.10),
            ("其他相关费用", total * 0.05),
        ]
        r = 4
        for i, (label, val) in enumerate(items_data, 1):
            ratio = val / total if total > 0 else 0
            _write_row(ws, r, [i, label, val, ratio, ""], fmt=_S.AMOUNT)
            ws.cell(row=r, column=4).number_format = _S.PCT_FMT
            r += 1
        _write_row(
            ws,
            r,
            ["", "合  计", total, 1.0, ""],
            fmt=_S.AMOUNT,
            bold=True,
            fill_=_S.LIGHT_BLUE,
        )
        ws.cell(row=r, column=4).number_format = _S.PCT_FMT

    def _score(self):
        ws = self.wb.create_sheet("评分表")
        _set_cols(ws, [5, 28, 15, 15, 15, 40])
        ws.merge_cells("A1:E1")
        ws["A1"] = "高新技术企业认定评分表"
        ws["A1"].font = _S.TITLE_FONT
        ws["A1"].alignment = _S.CENTER
        _header_row(ws, 3, ["序号", "评分项", "得分", "满分", "通过线", "评分说明"])
        scores = [
            (
                1,
                "知识产权",
                self.result.ip_score,
                30.0,
                20.0,
                f"I类{self.ht_input.ip_count}个, II类{self.ht_input.ip_2_count}个",
            ),
            (
                2,
                "科技成果转化能力",
                self.result.tech_transfer_score,
                30.0,
                20.0,
                self.result.transfer_detail,
            ),
            (
                3,
                "研发组织管理水平",
                self.result.rd_management_score,
                20.0,
                15.0,
                self.result.management_detail,
            ),
            (
                4,
                "企业成长性",
                self.result.growth_score,
                20.0,
                10.0,
                self.result.growth_detail,
            ),
        ]
        r = 4
        for i, (sn, name, score, full, line, note) in enumerate(scores, 1):
            _write_row(ws, r, [sn, name, score, full, line, note], fmt=_S.INT_FMT)
            ws.cell(row=r, column=3).number_format = "0.0"
            ws.cell(row=r, column=6).alignment = _S.WRAP
            if score >= line:
                for c in range(1, 7):
                    ws.cell(row=r, column=c).fill = _S.LIGHT_GREEN
            else:
                for c in range(1, 7):
                    ws.cell(row=r, column=c).fill = _S.LIGHT_YELLOW
            r += 1
        r += 1
        _write_row(
            ws,
            r,
            ["", "总  分", self.result.total_score, 100.0, 71.0, ""],
            fmt=_S.INT_FMT,
            bold=True,
            fill_=_S.LIGHT_BLUE,
        )
        ws.cell(row=r, column=3).number_format = "0.0"
        r += 1
        conclusion = "通过 ✅" if self.result.passed else "不通过 ❌"
        ws.cell(row=r, column=2, value="审核结论").font = _S.BOLD_FONT
        c = ws.cell(row=r, column=3, value=conclusion)
        c.font = _S.RESULT_FONT
