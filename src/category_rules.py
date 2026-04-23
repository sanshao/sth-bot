from __future__ import annotations

from typing import Any, Dict, Iterable, List

import pandas as pd


Rule = Dict[str, Any]
Condition = Dict[str, Any]

SUPPORTED_OPERATORS = [
    "eq",
    "ne",
    "contains",
    "startswith",
    "endswith",
    "gt",
    "gte",
    "lt",
    "lte",
    "=",
    "==",
    "!=",
    ">",
    ">=",
    "<",
    "<=",
]


def build_keyword_rules(
    category_mapping: Dict[str, Iterable[str]],
    field: str = "备注",
) -> List[Rule]:
    rules: List[Rule] = []
    for category, keywords in category_mapping.items():
        rules.append(
            {
                "category": category,
                "condition": {
                    "any": [
                        {"field": field, "op": "contains", "value": keyword}
                        for keyword in keywords
                    ]
                },
            }
        )
    return rules


def assign_category_by_rules(
    row: pd.Series,
    rules: Iterable[Rule],
    default: str = "其它-未分类",
) -> str:
    row_data = row.to_dict()
    for rule in rules:
        condition = rule.get("condition", {})
        if evaluate_condition(condition, row_data):
            return str(rule["category"])
    return default


def get_supported_operators() -> List[str]:
    return SUPPORTED_OPERATORS.copy()


def evaluate_condition(condition: Condition, row_data: Dict[str, Any]) -> bool:
    if "all" in condition:
        return all(evaluate_condition(item, row_data) for item in condition["all"])
    if "any" in condition:
        return any(evaluate_condition(item, row_data) for item in condition["any"])
    if "not" in condition:
        return not evaluate_condition(condition["not"], row_data)

    field = condition.get("field")
    op = str(condition.get("op", "eq")).lower()
    expected = condition.get("value")
    actual = row_data.get(field)
    return evaluate_leaf(actual, op, expected)


def evaluate_leaf(actual: Any, op: str, expected: Any) -> bool:
    op_alias = {
        "=": "eq",
        "==": "eq",
        "equals": "eq",
        "!=": "ne",
        "<>": "ne",
        ">": "gt",
        ">=": "gte",
        "<": "lt",
        "<=": "lte",
    }
    normalized_op = op_alias.get(op, op)

    if normalized_op in ("contains", "startswith", "endswith"):
        actual_text = normalize_text(actual)
        expected_text = normalize_text(expected)
        if normalized_op == "contains":
            return expected_text in actual_text
        if normalized_op == "startswith":
            return actual_text.startswith(expected_text)
        return actual_text.endswith(expected_text)

    if normalized_op in ("gt", "gte", "lt", "lte"):
        actual_num = to_number(actual)
        expected_num = to_number(expected)
        if actual_num is None or expected_num is None:
            return False
        if normalized_op == "gt":
            return actual_num > expected_num
        if normalized_op == "gte":
            return actual_num >= expected_num
        if normalized_op == "lt":
            return actual_num < expected_num
        return actual_num <= expected_num

    if normalized_op == "ne":
        return not compare_equal(actual, expected)

    if normalized_op == "in":
        expected_list = expected if isinstance(expected, list) else []
        return normalize_text(actual) in [normalize_text(item) for item in expected_list]

    if normalized_op == "not_in":
        expected_list = expected if isinstance(expected, list) else []
        return normalize_text(actual) not in [normalize_text(item) for item in expected_list]

    return compare_equal(actual, expected)


def compare_equal(actual: Any, expected: Any) -> bool:
    actual_num = to_number(actual)
    expected_num = to_number(expected)
    if actual_num is not None and expected_num is not None:
        return actual_num == expected_num
    return normalize_text(actual) == normalize_text(expected)


def normalize_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_number(value: Any) -> float | None:
    if pd.isna(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None
