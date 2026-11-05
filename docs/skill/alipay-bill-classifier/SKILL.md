---
name: alipay-bill-classifier
description: Use when processing, classifying, or troubleshooting e-commerce bill Excel files in this repository (淘宝店、淘工厂、天猫、淘宝买菜、聚合、京东、拼多多), or when adding/changing classification rules in src/rules/*.json. Covers running src/run.py, file routing, the JSON rule engine, tax-rate pivot splitting, and where each platform's ordered rule table lives.
---

# Alipay Bill Classifier

Use this skill for 账单整理、分类、透视汇总、规则排查或规则新增.

## Repository Map

| Path | Role |
| --- | --- |
| `src/run.py` | Batch entry point; routes files by name |
| `src/taobao.py` / `taobao_factory.py` / `tmall.py` / `taocc.py` | Alipay bill processors (淘宝店 / 淘工厂 / 天猫 / 淘宝买菜) |
| `src/group.py` | 聚合账单 processor |
| `src/jd.py`, `src/pdd.py` | 京东 / 拼多多 processors (no rule file; see below) |
| `src/category_rules.py` | Shared rule engine and tax-rate logic |
| `src/rules/<platform>_rules.json` | Ordered classification rules — the source of truth |
| `docs/gen_rule_docs.py` | Regenerates rule docs from the JSON files |
| `docs/rule/README.md` | Human-facing overview (Chinese) |

## Run Workflow

From the repository root (paths are resolved from the current working directory):

```bash
source .venv/bin/activate
pip install -r requirements.txt   # first run only
python src/run.py
```

The input directory is `target_directory` in `src/run.py` (currently `resource/2026/支付宝-10月`). Change it when the user points at another month or folder.

## File Routing

`src/run.py` checks each `.xlsx` file name in this order; the first match wins and other files are ignored:

| Order | File name | Processor | Rules |
| --- | --- | --- | --- |
| 1 | contains `聚合` | `group.py` | `group_rules.json` |
| 2 | starts with `淘宝买菜` | `taocc.py` | `taocc_rules.json` |
| 3 | starts with `淘工厂` | `taobao_factory.py` | `taobao_factory_rules.json` |
| 4 | starts with `淘宝` | `taobao.py` | `taobao_rules.json` |
| 5 | starts with `天猫` | `tmall.py` | `tmall_rules.json` |
| 6 | starts with `京东` | `jd.py` | none — `费用项` is the category |
| 7 | starts with `拼多多` | `pdd.py` | none — `业务描述` is the category |

`pdd_rules.json` is only used by `src/pdd1.py`, which `run.py` does not call. `jd_rules.json` is not used by any processor.

## Processing Behavior (Alipay bills)

1. Read the first sheet with `header=4` and `skipfooter=4`; strip column names.
2. 淘工厂 only: when `业务类型 == 在线支付` and `备注` is blank, copy `商品名称` into `备注`.
3. `净值 = 收入金额（+元） + 支出金额（-元）`.
4. `分类` from the ordered rules; default `其它-未分类`.
5. `税率` from `商品名称` (13% / 9% / blank) for categories containing `交易收款`, `限时红包`, `多买多省`.
6. Move `净值` between `支出金额（-元）` and `账户余额（元）`.
7. Pivot `净值` sum and `行数` by category. Categories containing `交易收款`, `限时红包`, `消费券` with a non-empty tax rate get a suffix (e.g. `淘宝-交易收款9%`). Append a `总和` row.

聚合 uses `净值 = 收入金额(元) - 支出金额(元)` and fixed 14-column naming when the header row is missing. Read `docs/rule/README.md` for full details, including the tax keyword lists.

## Output

`output/整理/YYYYMMDD_HHMMSS/<original-name>_整理.xlsx` with sheets `整理` (rows + computed columns) and `透视` (pivot + `总和`).

## Rule Engine

Each rule file is an ordered JSON array. **The first matching rule wins.**

```json
{ "category": "淘工厂-淘客佣金",
  "condition": { "field": "备注", "op": "contains", "value": "C2M-正向扣佣-精选淘客" } }
```

- `condition`: one condition. `conditions`: array, matches if **any** item matches.
- Nesting: `{"all": [...]}`, `{"any": [...]}`, `{"not": {...}}`.
- Operators: `contains`, `startswith`, `endswith`, `eq`/`=`/`==`, `ne`/`!=`, `>`, `>=`, `<`, `<=`, `in`, `not_in`.
- Text is compared after `strip()`; numeric operators return false when a value is not numeric; blank cells are empty strings.

## Classification References

Read only the reference for the platform involved. Each one lists rules in priority order and ends with a list of conditions that can never match:

| Platform | Reference |
| --- | --- |
| 淘宝店 | `references/taobao-rules.md` |
| 淘工厂店 | `references/taobao_factory-rules.md` |
| 天猫店 | `references/tmall-rules.md` |
| 淘宝买菜 | `references/taocc-rules.md` |
| 聚合账单 | `references/group-rules.md` |
| 拼多多 | `references/pdd-rules.md` |
| 京东 | `references/jd-rules.md` |

## Adding or Changing Rules

1. Identify which platform files the user means (e.g. "淘宝店和淘工厂店" means `taobao_rules.json` and `taobao_factory_rules.json`). The same category often exists in several files; update every requested file.
2. **Order matters.** If the new keyword contains an existing rule's keyword (e.g. `C2M-正向扣佣-精选淘客` contains `正向扣佣-精选淘客-`), insert the new rule *before* that existing rule. Otherwise the existing rule swallows it.
3. To add a keyword to an existing category, append to its `conditions` array; convert `condition` to `conditions` when it needs more than one entry.
4. Keep the JSON strict: no trailing commas, no comments (`json.load` will fail and the whole platform breaks).
5. Verify with a sample row:

   ```bash
   python -c "
   import json, sys, pandas as pd
   sys.path.insert(0, 'src')
   from category_rules import assign_category_by_rules
   rules = json.load(open('src/rules/taobao_factory_rules.json'))
   print(assign_category_by_rules(pd.Series({'备注': '示例备注'}), rules))
   "
   ```

6. Regenerate the docs with `python docs/gen_rule_docs.py`. This rewrites `docs/rule/<platform>.md` and `references/<platform>-rules.md`, and prints how many unreachable conditions each platform has.
