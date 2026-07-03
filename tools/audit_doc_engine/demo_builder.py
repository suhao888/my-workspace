# -*- coding: utf-8 -*-
"""
DocBuilder 使用演示
===================
展示通用排版引擎在不同文档类型中的应用。

运行: python demo_builder.py
输出: D:/Users/12844/Desktop/演示文档排版.docx
"""

import sys, os

_PROJECT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _PROJECT not in sys.path:
    sys.path.insert(0, _PROJECT)
from tools.audit_doc_engine.doc_builder import DocBuilder


def demo_meeting_minutes():
    """会议纪要"""
    b = DocBuilder()
    b.setup_page().setup_styles(body_size=12)

    b.add_cover(
        title="项目进度周例会会议纪要",
        subtitle="华北油田公司 — 数智化转型项目组",
        date_text="2026年7月3日",
        confidential="内部资料",
    )
    b.setup_header("华北油田数智化转型项目组")
    b.setup_footer(left="会议纪要", right="page_number")

    b.add_heading_1("一、会议基本信息")
    b.add_kv_table(
        [
            ("会议主题", "数智化转型项目第23次周例会"),
            ("会议时间", "2026年7月3日 14:00-15:30"),
            ("会议地点", "油田公司A座3楼会议室"),
            ("主持人", "张明"),
            ("参会人员", "李华、王强、赵丽、陈晓东、刘洋"),
            ("记录人", "陈晓东"),
        ]
    )

    b.add_heading_1("二、议题讨论")
    b.add_heading_2("议题一：商旅平台对接进度")
    b.add_body(
        "商旅平台开发组汇报了本周工作进展。目前已完成与财务系统的基础数据对接，"
        "包括员工信息同步、部门架构映射等基础功能。预计下周启动报销流程联调测试。"
    )
    b.add_body(
        "存在问题：部分历史数据迁移时发现了编码不一致的情况，涉及约300条记录，"
        "需要与财务部核对后统一处理。"
    )

    b.add_heading_2("议题二：差旅费管理办法执行情况")
    b.add_body(
        "审计组通报了近期专项审查的结果。13家单位中，5家存在不同程度的违规情况，"
        "主要体现在报销附件不合规、审批流程倒置等方面。建议在商旅平台中增加审批"
        "流程控制，从系统层面杜绝先出差后审批的问题。"
    )

    b.add_heading_1("三、行动计划")
    b.add_table(
        headers=["序号", "行动项", "责任人", "完成时限"],
        rows=[
            ["1", "完成数据编码统一", "王强", "7月10日"],
            ["2", "启动联调测试", "李华", "7月12日"],
            ["3", "下发整改通知", "赵丽", "7月8日"],
            ["4", "更新管理办法实施细则", "刘洋", "7月15日"],
        ],
        col_widths=[1, 4, 2, 2],
        caption="表：本周行动计划",
    )

    b.add_heading_1("四、下次会议安排")
    b.add_body("下次例会时间：2026年7月10日 14:00，地点不变。")
    b.add_body("请各负责人提前准备本周工作进展汇报材料。")

    return b


def demo_resume():
    """个人简历"""
    b = DocBuilder()
    b.setup_page(margins={"left": 2.5, "right": 2.5, "top": 2.0, "bottom": 2.0})
    b.setup_styles(body_size=11, line_spacing=1.3)

    b.add_cover(title="个人简历", subtitle="应聘：高级审计经理", date_text="2026年7月")

    b.add_heading_1("基本信息")
    b.add_kv_table(
        [
            ("姓名", "张三"),
            ("电话", "138-0013-8000"),
            ("邮箱", "zhangsan@email.com"),
            ("出生年月", "1990年1月"),
            ("政治面貌", "中共党员"),
        ]
    )

    b.add_heading_1("教育背景")
    b.add_table(
        headers=["时间", "学校", "专业", "学位"],
        rows=[
            ["2015.09-2018.06", "上海财经大学", "会计学", "硕士"],
            ["2011.09-2015.06", "上海财经大学", "会计学", "学士"],
        ],
        col_widths=[3, 4, 3, 2],
    )

    b.add_heading_1("工作经历")
    b.add_heading_2("普华永道（中国）  |  高级审计师  |  2020.07-2024.06")
    b.add_body(
        "主导并执行3个大型IPO上市审计项目，包括前期尽职调查、财务报表审阅及上市申报材料复核。"
    )
    b.add_body(
        "负责5家上市公司及2家大型民营企业的年度财务报表审计工作，严格遵循审计底稿编制规范。"
    )
    b.add_body(
        "深入分析客户内控体系，识别关键风险点并提出改进建议，帮助客户优化内控流程。"
    )

    b.add_heading_2("德勤华永会计师事务所  |  审计师  |  2018.07-2020.06")
    b.add_body(
        "参与8家跨国企业和国内知名企业的年度审计项目，负责底稿编制、往来款项函证、存货监盘等基础审计工作。"
    )
    b.add_body(
        "协助高级审计师进行风险评估程序和内控测试，撰写相关工作底稿和内部控制缺陷报告。"
    )

    b.add_heading_1("专业技能")
    b.add_list_item("审计技能：IPO上市审计、财务报表审计、内部控制审计、风险评估")
    b.add_list_item("会计准则：中国会计准则（CAS）、国际财务报告准则（IFRS）")
    b.add_list_item("审计工具：Auditor Assistant、ACL、Excel高级分析、SAP、用友")
    b.add_list_item("语言能力：英语（商务流利）、普通话（母语）")

    b.add_heading_1("证书资质")
    b.add_table(
        headers=["证书名称", "颁发机构", "取得时间"],
        rows=[
            ["中国注册会计师（CPA）", "中国注册会计师协会", "2019.08"],
            ["ACCA", "特许公认会计师公会", "2017.06"],
        ],
        col_widths=[4, 4, 2],
    )

    b.add_signature(
        [
            ("候选人", "张三"),
            ("日期", "2026年7月3日"),
        ]
    )

    return b


def demo_data_report():
    """数据统计报告"""
    b = DocBuilder()
    b.setup_page().setup_styles()

    b.add_cover(
        title="2026年上半年度差旅费统计分析报告",
        subtitle="华北油田公司财务部",
        date_text="2026年7月",
    )
    b.setup_header("华北油田公司 — 差旅费统计分析")
    b.setup_footer(left="财务部", right="page_x_of_y")

    b.add_heading_1("一、总体情况")
    b.add_body(
        "2026年上半年度，华北油田公司下属13家单位累计发生差旅费3,486,295.68元，"
        "其中交通费1,245,678.50元，住宿费1,892,345.18元，出差补助348,272.00元。"
        "与2025年同期相比，差旅费总额下降12.3%，主要得益于商旅平台的推广使用。"
    )

    b.add_heading_1("二、各单位差旅费支出")
    b.add_table(
        headers=[
            "单位名称",
            "差旅费总额(元)",
            "交通费(元)",
            "住宿费(元)",
            "商旅使用率",
        ],
        rows=[
            ["勘探开发研究院", "823,456.00", "312,456.00", "435,678.00", "78.5%"],
            ["经济技术研究院", "654,321.00", "245,678.00", "345,678.00", "62.2%"],
            ["数智技术公司", "523,456.00", "198,765.00", "287,654.00", "80.7%"],
            ["雄安分公司", "456,789.00", "176,543.00", "245,678.00", "74.8%"],
            ["公共事务中心", "345,678.00", "132,456.00", "187,654.00", "85.3%"],
            ["其他8家单位", "682,595.68", "179,780.50", "389,903.18", "82.1%"],
        ],
        col_widths=[3.5, 2.5, 2.5, 2.5, 2],
        caption="表1：各单位差旅费支出明细",
    )

    b.add_heading_1("三、商旅平台使用分析")
    b.add_body(
        "整体商旅平台使用率为76.8%，较2025年下半年的65.2%提高了11.6个百分点。"
        "其中住宿费的商旅平台使用率最高，达到82.3%；交通费为71.2%。"
    )
    b.add_body(
        "从各单位来看，公共事务中心的商旅平台使用率最高（85.3%），经济技术研究院"
        "最低（62.2%）。建议对使用率低于80%的单位进行重点督导。"
    )

    b.add_heading_1("四、问题与建议")
    b.add_body(
        "统计数据显示，部分单位仍存在以下问题：一是部分差旅费报销附件不完整，"
        "涉及金额约7.4万元；二是存在审批流程倒置现象；三是少数报销存在超标准情况。"
    )
    b.add_note_box(
        "建议：各单位应严格执行《差旅费管理办法》，财务部将建立月度通报机制，"
        "对商旅平台使用率连续三个月低于80%的单位进行约谈。"
    )

    return b


def main():
    output_dir = r"D:\Users\12844\Desktop"

    # 生成三种文档
    reports = [
        ("会议纪要_数智化项目周例会.docx", demo_meeting_minutes()),
        ("个人简历_张三.docx", demo_resume()),
        ("差旅费统计分析报告_2026上半年.docx", demo_data_report()),
    ]

    for fname, builder in reports:
        path = os.path.join(output_dir, fname)
        builder.save(path)
        size = os.path.getsize(path) / 1024
        print(f"  -> {fname}  ({size:.1f} KB)")

    print(f"\n全部生成完毕，文件保存在: {output_dir}")


if __name__ == "__main__":
    main()
