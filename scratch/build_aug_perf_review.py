"""生成 8 月杭州运营绩效核对版（不修改原文件）。"""
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill

DIR = Path(__file__).resolve().parents[1] / "resource/2026/8月绩效"
SRC = DIR / "杭州运营绩效-8月.xlsx"
DST = DIR / "杭州运营绩效-8月-核对版.xlsx"

RED = PatternFill("solid", fgColor="FFC7CE")
YELLOW = PatternFill("solid", fgColor="FFEB9C")
GREEN = PatternFill("solid", fgColor="C6EFCE")
HEAD = PatternFill("solid", fgColor="D9E1F2")
BOLD = Font(bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")
AUTHOR = "核对"

# 方案：千橙经销组 坎级（毛利率下限%, 提成比例），按下限从高到低匹配，>25% 视同最高档
TIERS = {
    "美妆": [(18, .015), (14, .012), (10, .01), (9, .008), (8, .007), (4, .005), (0, .002)],
    "非奶食品": [(16, .015), (12, .012), (10, .01), (8, .008), (6, .007), (4, .005), (0, .002)],
    "低温奶": [(16, .015), (12, .012), (10, .01), (8, .008), (5, .007), (3, .005), (0, .002)],
}
RESOURCE_HALF = {"88VIP"}


def tier(category, margin):
    if category not in TIERS or margin is None or margin <= 0:
        return None
    for low, rate in TIERS[category]:
        if margin * 100 >= low:
            return rate
    return None


def bounded_rate(actual, target):
    if not target:
        return None
    return min(1.5, max(0.6, actual / target))


def header(ws, row, titles, widths=None):
    for i, t in enumerate(titles, 1):
        c = ws.cell(row, i, t)
        c.font, c.fill, c.alignment = BOLD, HEAD, WRAP
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def mark(ws, ref, fill, note):
    ws[ref].fill = fill
    ws[ref].comment = Comment(note, AUTHOR, width=320, height=140)


def main():
    values = openpyxl.load_workbook(SRC, data_only=True)["销售数据"]
    wb = openpyxl.load_workbook(SRC)
    ws = wb["销售数据"]

    # ---- 1. 修正公式错误 ----
    unique_store_rows = [2, 3, 4, 5, 6, 7, 9, 13, 15, 16, 17, 18, 21, 22, 23, 24, 25, 26]
    ws["G33"] = "=" + "+".join(f"G{r}" for r in unique_store_rows)
    mark(ws, "G33", RED, "原公式 =SUM(G2:G26) 重复计入同店铺共用毛利（天猫食品×2、淘工厂×4、淘宝买菜×3、橙理×2）。已改为每店只计一次，应为 113,269.85。")
    ws["G35"] = "=G33+G34"
    mark(ws, "G35", RED, "原公式 =G33-G34，代运营毛利为负，相减反而抬高合计。已改为 =G33+G34。")

    # ---- 2. 标注待确认项（不改金额）----
    mark(ws, "M21", YELLOW, "王敏管理层提成：完成率用 5 店完成率简单平均(1.342)；方案为“经销项目总月度经营目标完成率”。按汇总口径(封顶150%)约 4,761.37。另需确认计入范围及“业务线整体盈利”条件。见【方案口径测算】。")
    mark(ws, "L11", YELLOW, "淘工厂-简爱完成率 114%、店铺盈利，未发基础奖励 1000。需确认基础奖励按店铺还是按品牌目标发放。")
    mark(ws, "L19", YELLOW, "淘宝买菜-简爱完成率 255%、店铺盈利，未发基础奖励 1000。需确认基础奖励按店铺还是按品牌目标发放。")
    mark(ws, "I25", YELLOW, "万橙汇京东自营：若属“自营供货”类，方案第4条要求提成减半，应为 0.35%，提成 203.97。")
    mark(ws, "I3", YELLOW, "毛利率 33.7% 超出美妆最高档 (18,25)，方案未定义 >25%，现按 1.5% 计。")
    mark(ws, "K29", YELLOW, "缺少公式。店铺毛利为负，按方案为 0，不影响金额。")
    mark(ws, "K32", YELLOW, "乔克蒂姆在方案中无对应规则；京东店毛利为正(89.9%)，目前为 0，需确认。")
    mark(ws, "M2", YELLOW, "方案【管理层提成机制】中但可（运营主管）0.1% 管理层提成本表未计算。见【方案口径测算】。")

    # ---- 3. 坎级复核 ----
    ck = wb.create_sheet("坎级复核", 0)
    header(ck, 1, ["行号", "考核分类", "运营", "店铺", "经营毛利率", "表内坎级", "复核坎级",
                   "表内完成率", "复核完成率(60%-150%)", "表内计算绩效", "复核计算绩效", "结果"],
           [6, 10, 12, 30, 11, 10, 10, 11, 14, 13, 13, 30])
    out = 2
    for r in range(2, 27):
        cat, op, shop, target, sales, _, margin, rate, j, k = (values.cell(r, c).value for c in range(2, 12))
        exp_rate = tier(cat, margin)
        if exp_rate and shop in RESOURCE_HALF:
            exp_rate /= 2
        exp_j = bounded_rate(sales, target)
        exp_k = round(sales * (exp_rate or 0) * (exp_j or 1), 2)
        ok_rate = (rate or 0) == (exp_rate or 0) or abs((rate or 0) - (exp_rate or 0)) < 1e-9
        ok_k = abs((k or 0) - exp_k) < 0.01
        note = "一致" if ok_rate and ok_k else "不一致"
        if shop == "淘宝-一个小木登子":
            note = "一致（毛利率>25%，按最高档 1.5%，待确认）"
        if shop == "京东-万橙汇京东自营":
            note = "一致（若属自营供货应减半，待确认）"
        row = [r, cat, op, shop, margin, rate, exp_rate, j, exp_j, k, exp_k, note]
        for i, v in enumerate(row, 1):
            ck.cell(out, i, v)
        for col in (5, 6, 7, 8, 9):
            ck.cell(out, col).number_format = "0.00%"
        for col in (10, 11):
            ck.cell(out, col).number_format = "#,##0.00"
        ck.cell(out, 12).fill = GREEN if note == "一致" else (YELLOW if note.startswith("一致") else RED)
        out += 1
    ck.cell(out + 1, 1, "说明：复核坎级按《千橙经销组》各分类区间下限匹配；88VIP 按资源型减半；无目标店铺完成率按 100% 计。代运营（三元/光明/乔克蒂姆）另见【核对说明】。")
    ck.freeze_panes = "A2"

    # ---- 4. 方案口径测算（公式联动原表）----
    cs = wb.create_sheet("方案口径测算", 1)
    header(cs, 1, ["项目", "计算项", "数值", "说明"], [26, 30, 16, 70])
    S = "销售数据!"
    rows = [
        ("但可 运营主管管理层提成", "盈利经销店铺销售额", f"=SUMIFS({S}F2:F26,{S}G2:G26,\">0\")", "方案：基于盈利店铺实际完成业绩"),
        ("", "提成比例", 0.001, "方案【管理层提成机制】"),
        ("", "公司总目标完成率(原始)", f"=SUMIFS({S}F2:F32,{S}E2:E32,\">0\")/SUM({S}E2:E32)", "假设：有目标店铺（含代运营）实际合计 / 目标合计，口径需确认"),
        ("", "计提完成率(60%-150%)", "=MIN(1.5,MAX(0.6,C4))", "假设沿用经销组上下限"),
        ("", "测算管理层提成", "=ROUND(C2*C3*C5,2)", "表内未计算"),
        ("", "", None, ""),
        ("王敏 自运营小组长管理层提成", "计入店铺销售额(沿用表内范围)", f"=SUM({S}F8,{S}F10:F11,{S}F18:F19)", "范围：吾岛/简爱于天猫食品、淘工厂、淘宝买菜 5 行，需确认"),
        ("", "对应目标合计", f"=SUM({S}E8,{S}E10:E11,{S}E18:E19)", ""),
        ("", "汇总完成率(原始)", "=C8/C9", "方案：按经销项目总月度经营目标完成率"),
        ("", "计提完成率(60%-150%)", "=MIN(1.5,MAX(0.6,C10))", "假设沿用 150% 上限"),
        ("", "方案口径测算", "=ROUND(C8*0.003*C11,2)", ""),
        ("", "表内现值", f"={S}M21", "表内用 5 店完成率简单平均"),
        ("", "差额", "=C12-C13", ""),
        ("", "", None, ""),
        ("万橙汇京东自营（若减半）", "测算提成", f"=ROUND({S}F25*0.007*50%,2)", "方案第4条：自营供货类减半"),
        ("", "差额", f"=C16-{S}K25", "王敏/尹慧分配相应减少"),
        ("", "", None, ""),
        ("基础奖励（若按品牌目标发放）", "淘工厂-简爱", f"=IF(AND({S}F11>={S}E11,{S}G11>0),1000,0)", "表内未发"),
        ("", "淘宝买菜-简爱", f"=IF(AND({S}F19>={S}E19,{S}G19>0),1000,0)", "表内未发"),
    ]
    for i, (a, b, c, d) in enumerate(rows, 2):
        cs.cell(i, 1, a).font = BOLD
        cs.cell(i, 2, b)
        cs.cell(i, 3, c)
        cs.cell(i, 4, d).alignment = WRAP
    for ref in ("C3", "C4", "C5", "C10", "C11"):
        cs[ref].number_format = "0.00%"
    for ref in ("C2", "C6", "C8", "C9", "C12", "C13", "C14", "C16", "C17", "C19", "C20"):
        cs[ref].number_format = "#,##0.00"

    # ---- 5. 核对说明 ----
    nt = wb.create_sheet("核对说明", 0)
    header(nt, 1, ["序号", "类型", "位置", "问题", "原值", "建议/测算值", "影响"],
           [6, 10, 22, 60, 16, 16, 30])
    issues = [
        ("公式错误", "销售数据!G33", "经营毛利 SUM 重复计入同店铺共用毛利", 320976.38, 113269.85, "仅汇总行，已修正"),
        ("公式错误", "销售数据!G35", "合计毛利公式为 G33-G34，应相加", 374094.45, 60151.78, "仅汇总行，已修正"),
        ("待确认", "销售数据!M2 / 方案口径测算", "但可运营主管 0.1% 管理层提成未计算", 0, "见测算表 C6", "但可"),
        ("待确认", "销售数据!M21 / 方案口径测算", "王敏管理层提成完成率用简单平均，方案为总完成率；计入范围及业务线盈利条件需确认", 4260.83, 4761.37, "王敏 +500.54"),
        ("待确认", "销售数据!L11、L19", "简爱品牌在淘工厂/淘宝买菜完成目标且盈利，未发基础奖励", 0, 2000, "吴垠寰、但可 各 +1000"),
        ("待确认", "销售数据!I25", "万橙汇京东自营若属自营供货，提成应减半", 407.94, 203.97, "王敏/尹慧 合计 -203.97"),
        ("待确认", "销售数据!I3", "毛利率 33.7% 超出方案最高档上限 25%", "1.5%", "方案未定义", "金额 14.57，影响小"),
        ("待确认", "销售数据!K30:K32", "乔克蒂姆无对应方案；京东店毛利为正", 0, "方案未定义", "章姝杭"),
        ("提示", "销售数据!K29", "光明代运营缺公式，毛利为负结果为 0", None, 0, "无"),
        ("提示", "Sheet1/2/3", "三张表内容相同（谢巧巧），目标与 Q3 目标表不符，拼多多简爱完成率按 100% 封顶（应 143%），疑似旧数据残留", None, None, "建议删除或单独核对"),
        ("提示", "绩效方案文件", "标题写“第一季度”、管理层适用范围写“第二季度”、审核小组“五人/三人”不一致", None, None, "文字修订"),
        ("已核对", "销售数据 各行", "坎级、完成率(150%封顶)、计算绩效、8月目标、88VIP减半、三元0.3%、人员汇总 均正确", None, None, "见【坎级复核】"),
    ]
    fills = {"公式错误": RED, "待确认": YELLOW, "提示": None, "已核对": GREEN}
    for i, item in enumerate(issues, 2):
        nt.cell(i, 1, i - 1)
        for j, v in enumerate(item, 2):
            c = nt.cell(i, j, v)
            c.alignment = WRAP
            if isinstance(v, float):
                c.number_format = "#,##0.00"
        if fills[item[0]]:
            nt.cell(i, 2).fill = fills[item[0]]
    nt.freeze_panes = "A2"

    wb.active = 0
    wb.save(DST)
    print("saved", DST)


if __name__ == "__main__":
    main()
