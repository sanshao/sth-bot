---
name: alipay-bill-classifier
description: Use this skill when processing, organizing, classifying, or troubleshooting Taobao and Tmall Alipay Excel bills in this repository. It covers how to run src/run.py, how files are selected, where outputs are written, and where to read the detailed Taobao/Tmall classification rules.
---

# Alipay Bill Classifier

Use this skill for Taobao/Tmall Alipay bill整理、分类、透视汇总、规则排查或规则新增。

## Repository Context

The implementation lives in:

- `src/run.py`: batch entry point.
- `src/taobao.py`: Taobao bill parser and classifier.
- `src/tmall.py`: Tmall bill parser and classifier.
- `src/category_rules.py`: shared rule engine.

Dependencies are listed in `requirements.txt`; the key runtime packages are `pandas` and `openpyxl`.

## Input Selection

`src/run.py` processes Excel files from a target directory. By default, it uses:

```python
target_directory = os.path.join(current_dir, "resource/支付宝-3月")
```

Inside that directory:

- Files whose name starts with `淘` and ends with `.xlsx` are processed by `process_taobao_file`.
- Files whose name starts with `天猫` and ends with `.xlsx` are processed by `process_tmall_file`.
- Other files are ignored.

If the user provides another input directory, update the `target_directory` value in `src/run.py` or invoke `process_taobao_files_in_directory(directory)` directly from Python.

## Run Workflow

From the repository root:

```bash
uv venv
source .venv/bin/activate
pip install -r requirements.txt
python src/run.py
```

If dependencies are already installed, only run:

```bash
python src/run.py
```

## Processing Behavior

For each matched workbook:

1. Read the first sheet.
2. Use row 5 as the header (`header=4`) and skip the last 4 footer rows.
3. Strip whitespace from column names.
4. Add `净值 = 收入金额（+元） + 支出金额（-元）`.
5. Add `分类` by matching ordered rules.
6. Move `净值` between `支出金额（-元）` and `账户余额（元）`.
7. Create a pivot table grouped by `分类`, summing `净值`.
8. Append a `总和` row to the pivot table.

Rules are evaluated in order. The first matching rule wins. If no rule matches, the category is `其它-未分类`.

## Output

Each run creates a timestamped output directory:

```text
output/整理/YYYYMMDD_HHMMSS/
```

Each processed workbook is written as:

```text
<original-name>_整理.xlsx
```

The output workbook has two sheets:

- `整理`: original rows plus computed `净值` and `分类`.
- `透视`: sum of `净值` by `分类`, with a final `总和` row.

## Classification References

Read only the relevant reference file when changing or explaining rules:

- Taobao rules: `references/taobao-rules.md`
- Tmall rules: `references/tmall-rules.md`

The reference files preserve the current rule order. Keep that order meaningful because it controls priority.

## Rule Maintenance Guidance

When adding or changing rules:

- Prefer editing `explicit_rules` for compound conditions, account-specific conditions, non-`备注` fields, numeric comparisons, or exclusions.
- Prefer editing `category_mapping` for simple keyword matching against `备注`.
- Add high-priority or narrow rules before broad rules.
- Preserve the default behavior of returning `其它-未分类` when nothing matches.
- After changes, run the script against a representative workbook and inspect both `整理` and `透视`.

The shared rule condition structure supports:

- Text operations: `contains`, `startswith`, `endswith`, `eq`, `=`, `==`, `ne`, `!=`.
- Numeric operations: `>`, `>=`, `<`, `<=`.
- Logical composition: `all`, `any`, `not`.

Example compound rule:

```python
{
    "category": "示例分类",
    "condition": {
        "all": [
            {"field": "备注", "op": "contains", "value": "关键字"},
            {"field": "支出金额（-元）", "op": "<", "value": 0},
        ]
    },
}
```
