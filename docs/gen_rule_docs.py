"""根据 src/rules/*.json 生成各平台的分类规则文档。

用法（仓库根目录）：python docs/gen_rule_docs.py
输出：
- docs/rule/<platform>.md
- docs/skill/alipay-bill-classifier/references/<platform>-rules.md
同时在终端打印被前序规则遮蔽、永远不会命中的条件。
"""
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RULES_DIR = os.path.join(ROOT, "src", "rules")
DOC_DIR = os.path.join(ROOT, "docs", "rule")
REF_DIR = os.path.join(ROOT, "docs", "skill", "alipay-bill-classifier", "references")

ALIPAY_STEPS = [
    "读取第一个工作表，以第 5 行为表头（`header=4`），跳过末尾 4 行汇总。",
    "清理列名两侧空格。",
    "新增 `净值 = 收入金额（+元） + 支出金额（-元）`（支出在源表中已是负数）。",
    "按本文规则表从上到下匹配，新增 `分类` 列；未命中为 `其它-未分类`。",
    "按商品名称新增 `税率` 列（规则见总览文档的“税率与透视细分”）。",
    "把 `净值` 移到 `支出金额（-元）` 与 `账户余额（元）` 之间。",
    "生成透视表：按分类汇总 `净值` 与 `行数`，并追加 `总和` 行。",
]

PLATFORMS = {
    "taobao": {
        "title": "淘宝店",
        "file": "文件名以 `淘宝` 开头的 `.xlsx`（`淘宝买菜` 开头的除外）",
        "module": "src/taobao.py",
        "steps": ALIPAY_STEPS,
    },
    "taobao_factory": {
        "title": "淘工厂店",
        "file": "文件名以 `淘工厂` 开头的 `.xlsx`",
        "module": "src/taobao_factory.py",
        "steps": ALIPAY_STEPS[:2]
        + ["`业务类型` 为 `在线支付` 且 `备注` 为空（含空白/制表符）时，用 `商品名称` 填充 `备注`，以便规则匹配。"]
        + ALIPAY_STEPS[2:],
    },
    "tmall": {
        "title": "天猫店",
        "file": "文件名以 `天猫` 开头的 `.xlsx`",
        "module": "src/tmall.py",
        "steps": ALIPAY_STEPS,
    },
    "taocc": {
        "title": "淘宝买菜（淘菜菜）",
        "file": "文件名以 `淘宝买菜` 开头的 `.xlsx`",
        "module": "src/taocc.py",
        "steps": ALIPAY_STEPS,
    },
    "group": {
        "title": "聚合账单",
        "file": "文件名包含 `聚合` 的 `.xlsx`（优先级最高）",
        "module": "src/group.py",
        "steps": [
            "读取第一个工作表；首行含 `商家昵称` 或 `支付流水号` 时视为表头，否则按固定 14 列命名："
            "商家昵称、入帐日期、入帐时间、支付流水号、主订单id、子订单id、入帐类型、收入金额(元)、支出金额(元)、业务描述、备注、收/付渠道、数据创建时间、数据修改时间。",
            "金额列转为数值，空值按 0 处理。",
            "新增 `净值 = 收入金额(元) - 支出金额(元)`（支出在源表中为正数）。",
            "按本文规则表匹配 `分类`，未命中为 `其它-未分类`；按商品名称新增 `税率` 列。",
            "把 `净值` 移到 `支出金额(元)` 之后。",
            "生成透视表：按分类汇总 `净值` 与 `行数`，并追加 `总和` 行。",
        ],
    },
    "pdd": {
        "title": "拼多多",
        "file": "文件名以 `拼多多` 开头的 `.xlsx`",
        "module": "src/pdd.py（`run.py` 当前使用）；src/pdd1.py（使用本规则，未接入 `run.py`）",
        "steps": [
            "自动查找含 `商户订单号` 的行作为表头（找不到时默认第 5 行）。",
            "金额列转为数值，新增 `净值 = 收入金额（+元） + 支出金额（-元）`。",
            "**`run.py` 当前调用的 `src/pdd.py` 不读取本规则文件**，直接把 `业务描述` 作为 `分类`。",
            "`src/pdd1.py` 才会使用本规则：按规则表匹配，未命中时以原始 `业务描述` 作为分类，并把 `净值`、`分类`、`税率` 放到 `支出金额（-元）` 之后。",
            "生成透视表：按分类汇总 `净值` 与 `行数`，并追加 `总和` 行。",
        ],
    },
    "jd": {
        "title": "京东",
        "file": "文件名以 `京东` 开头的 `.xlsx`",
        "module": "src/jd.py",
        "steps": [
            "读取第一个工作表，首行为表头。",
            "**`src/jd.py` 不读取本规则文件**，直接把 `费用项` 作为 `分类`；缺少该列时跳过该文件。",
            "`金额` 转为数值；`收支方向` 含 `支出` 且金额为正时按负数统计。",
            "生成透视表：按分类汇总 `金额` 与 `行数`，并追加 `总和` 行（不计算净值和税率）。",
        ],
    },
}

OP_TEXT = {
    "contains": "包含",
    "startswith": "开头为",
    "endswith": "结尾为",
    "eq": "等于", "=": "等于", "==": "等于", "equals": "等于",
    "ne": "不等于", "!=": "不等于", "<>": "不等于",
    "gt": "大于", ">": "大于",
    "gte": "大于等于", ">=": "大于等于",
    "lt": "小于", "<": "小于",
    "lte": "小于等于", "<=": "小于等于",
    "in": "属于", "not_in": "不属于",
}


def esc(text):
    return str(text).replace("|", "\\|")


def render(cond):
    if "all" in cond:
        return "，且".join(render(c) for c in cond["all"])
    if "any" in cond:
        return "（" + " 或 ".join(render(c) for c in cond["any"]) + "）"
    if "not" in cond:
        return "不满足（" + render(cond["not"]) + "）"
    op = str(cond.get("op", "eq")).lower()
    value = cond.get("value")
    value_text = str(value) if isinstance(value, (int, float)) else f"“{esc(value)}”"
    return f"“{cond.get('field')}”{OP_TEXT.get(op, op)}{value_text}"


def top_conditions(rule):
    return rule.get("conditions") or [rule.get("condition", {})]


def rule_table(rules):
    lines = ["| 序号 | 分类 | 匹配条件（任一行命中即可） |", "| --- | --- | --- |"]
    for i, rule in enumerate(rules, 1):
        cell = "<br>".join(render(c) for c in top_conditions(rule))
        lines.append(f"| {i} | {esc(rule['category'])} | {cell} |")
    return "\n".join(lines)


def shadow_warnings(rules):
    """前序 contains 条件是后续 contains 条件的子串时，后续条件永远不会命中。"""
    seen = []
    warnings = []
    for i, rule in enumerate(rules, 1):
        leaves = [c for c in top_conditions(rule) if c.get("op") == "contains" and "field" in c]
        for cond in leaves:
            value = str(cond["value"]).strip()
            for j, category, field, prev in seen:
                if field == cond["field"] and prev and prev in value and category != rule["category"]:
                    warnings.append(
                        f"第 {i} 条 `{esc(rule['category'])}` 的“{esc(value)}”"
                        f"会先被第 {j} 条 `{esc(category)}` 的“{esc(prev)}”命中"
                    )
                    break
        seen.extend((i, rule["category"], c["field"], str(c["value"]).strip()) for c in leaves)
    return warnings


def build_doc(name, meta, rules):
    warnings = shadow_warnings(rules)
    parts = [
        f"# {meta['title']}账单分类规则",
        "",
        f"> 本文档由 `docs/gen_rule_docs.py` 根据 `src/rules/{name}_rules.json` 自动生成，请勿手动修改规则表；"
        "修改 JSON 后重新运行脚本即可。",
        "",
        "## 适用范围",
        "",
        "| 项目 | 说明 |",
        "| --- | --- |",
        f"| 账单文件 | {meta['file']} |",
        f"| 处理模块 | {meta['module']} |",
        f"| 规则文件 | `src/rules/{name}_rules.json` |",
        f"| 规则条数 | {len(rules)} |",
        "",
        "## 处理流程",
        "",
    ]
    parts += [f"{i}. {step}" for i, step in enumerate(meta["steps"], 1)]
    parts += [
        "",
        "## 分类规则",
        "",
        "- 规则按序号从小到大匹配，**第一条命中的规则决定分类**，序号越小优先级越高。",
        "- 同一分类下有多行条件时，任意一行命中即可；一行内用“且”连接的条件需同时满足。",
        "- “包含”为子串匹配，“等于”会先去掉首尾空格再比较。",
        "",
        rule_table(rules),
        "",
    ]
    if warnings:
        parts += [
            "## 优先级提示",
            "",
            "以下条件的关键字包含了更靠前规则的关键字，按当前顺序**永远不会命中**。"
            "如需生效，应把对应规则移到前面，或收窄前序规则的关键字：",
            "",
        ]
        parts += [f"- {w}" for w in warnings]
        parts.append("")
    return "\n".join(parts), warnings


if __name__ == "__main__":
    os.makedirs(REF_DIR, exist_ok=True)
    for name, meta in PLATFORMS.items():
        with open(os.path.join(RULES_DIR, f"{name}_rules.json"), encoding="utf-8") as f:
            rules = json.load(f)
        doc, warnings = build_doc(name, meta, rules)
        for path in (os.path.join(DOC_DIR, f"{name}.md"), os.path.join(REF_DIR, f"{name}-rules.md")):
            with open(path, "w", encoding="utf-8") as f:
                f.write(doc)
        print(f"{name}: {len(rules)} 条规则，{len(warnings)} 条优先级提示")
