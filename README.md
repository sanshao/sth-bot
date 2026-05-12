# sth-bot

```
uv venv
```

```
source .venv/bin/activate
```


```bash
pip install -r requirements.txt
```

```
python src/run.py
```

## 分类规则使用说明

当前淘宝/天猫分类已从硬编码 `if/elif` 升级为规则引擎，核心文件：

- `src/category_rules.py`
- `src/taobao.py`
- `src/tmall.py`

分类规则按顺序匹配，命中第一条即返回分类；都不命中则返回默认值 `其它-未分类`。

### 1. 规则结构

每条规则结构：

```python
{
  "category": "分类名称",
  "condition": {
    "field": "备注",
    "op": "contains",
    "value": "关键字"
  }
}
```

叶子条件字段：

- `field`: Excel 列名（如 `备注`、`商品名称`、`业务类型`、`支出金额（-元）`）
- `op`: 操作符
- `value`: 比较值

### 2. 支持的操作符

文本匹配：

- `contains` 包含匹配
- `startswith` 开头匹配
- `endswith` 结尾匹配
- `eq` / `=` / `==` 全匹配
- `ne` / `!=` 不等于

数值比较（自动尝试转数字）：

- `>`、`>=`、`<`、`<=`

### 3. 逻辑组合（SQL 风格）

- `all`: 且（AND）
- `any`: 或（OR）
- `not`: 非（NOT）

示例：`(备注包含A 或 备注包含B) 且 支出金额<=-100`

```python
{
  "category": "示例分类",
  "condition": {
    "all": [
      {
        "any": [
          {"field": "备注", "op": "contains", "value": "A"},
          {"field": "备注", "op": "contains", "value": "B"}
        ]
      },
      {"field": "支出金额（-元）", "op": "<=", "value": -100}
    ]
  }
}
```

### 4. 在代码中的使用方式

`taobao.py` / `tmall.py` 中规则由两部分组成：

1. `explicit_rules`: 显式组合规则（复杂条件）
2. `build_keyword_rules(category_mapping, field="备注")`: 把原关键词映射自动转为“备注 contains 关键字”

最终分类调用：

```python
rules = explicit_rules + build_keyword_rules(category_mapping, field="备注")
df["分类"] = df.apply(lambda row: assign_category_by_rules(row, rules), axis=1)
```

### 5. 给后续 Web 配置的建议

- 前端可直接维护 `rules` 的 JSON 结构（与当前 Python 结构一致）。
- 后端只需按顺序执行规则即可，保证和本地脚本行为一致。
- 新规则建议插入到 `explicit_rules` 前部（优先级更高）。
