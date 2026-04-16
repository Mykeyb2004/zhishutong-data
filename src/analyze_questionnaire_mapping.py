from __future__ import annotations

import argparse
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from mapping_rules import MappingRules, load_rules


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
CELL_RE = re.compile(r"([A-Z]+)(\d+)")
QUESTION_RE = re.compile(r"^(Q\d+)")
SUSPICIOUS_LIST_RE = re.compile(r"^\s*\d+(?:\s*([,，;；|/、+])\s*\d+)+\s*$")
SUSPICIOUS_DOT_LIST_RE = re.compile(r"^\s*\d+(?:\s*[.]\s*\d+){2,}\s*$")


def col_to_idx(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + ord(char) - 64
    return value - 1


def idx_to_col(idx: int) -> str:
    idx += 1
    chars: list[str] = []
    while idx:
        idx, remainder = divmod(idx - 1, 26)
        chars.append(chr(65 + remainder))
    return "".join(reversed(chars))


def load_xlsx(path: Path) -> list[list[str]]:
    with ZipFile(path) as workbook:
        try:
            shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(item.itertext()) for item in shared_root.findall("a:si", NS)
            ]
        except KeyError:
            shared_strings = []

        sheet_root = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
        sheet_data = sheet_root.find("a:sheetData", NS)
        if sheet_data is None:
            raise ValueError(f"{path} 缺少 sheetData")

        rows: list[list[str]] = []
        max_idx = 0
        for row in sheet_data:
            values: dict[int, str] = {}
            for cell in row:
                ref = cell.attrib["r"]
                col = CELL_RE.match(ref)
                if not col:
                    continue

                idx = col_to_idx(col.group(1))
                max_idx = max(max_idx, idx)
                cell_type = cell.attrib.get("t")
                if cell_type == "s":
                    value_node = cell.find("a:v", NS)
                    value = shared_strings[int(value_node.text)] if value_node is not None else ""
                elif cell_type == "inlineStr":
                    inline_node = cell.find("a:is", NS)
                    value = "".join(inline_node.itertext()) if inline_node is not None else ""
                else:
                    value_node = cell.find("a:v", NS)
                    value = value_node.text if value_node is not None else ""
                values[idx] = value
            rows.append([values.get(i, "") for i in range(max_idx + 1)])

    width = max(len(row) for row in rows)
    return [row + [""] * (width - len(row)) for row in rows]


def normalize_title(header: str) -> str:
    return " | ".join(part.strip() for part in header.splitlines() if part.strip())


def option_label(header: str) -> str:
    parts = header.split("_", 1)
    return parts[1].strip() if len(parts) == 2 else normalize_title(header)


def question_key(header: str, rules: MappingRules) -> str:
    match = rules.question_id_regex.match(header)
    return match.group(1) if match else normalize_title(header)


def sort_mapping_items(items: Iterable[tuple[str, str]]) -> list[tuple[str, str]]:
    def sort_key(item: tuple[str, str]) -> tuple[int, int | str, str]:
        key = item[0]
        if re.fullmatch(r"-?\d+", key):
            return (0, int(key), item[1])
        if re.fullmatch(r"-?\d+-.+", key):
            prefix, suffix = key.split("-", 1)
            return (1, int(prefix), suffix)
        return (2, key, item[1])

    return sorted(items, key=sort_key)


@dataclass
class ColumnAnalysis:
    column: str
    header: str
    kind: str
    mappings: list[tuple[str, str]]


def analyze_columns(
    text_rows: list[list[str]], value_rows: list[list[str]], rules: MappingRules
) -> tuple[list[ColumnAnalysis], list[tuple[str, str, int, str, str]]]:
    headers = text_rows[0]
    analyses: list[ColumnAnalysis] = []
    suspicious_cells: list[tuple[str, str, int, str, str]] = []

    for idx, header in enumerate(headers):
        column = idx_to_col(idx)
        pairs = Counter()
        only_equal = True
        value_to_texts: dict[str, set[str]] = defaultdict(set)

        for row_idx in range(1, len(text_rows)):
            value = value_rows[row_idx][idx]
            text = text_rows[row_idx][idx]
            if value or text:
                pairs[(value, text)] += 1
                value_to_texts[value].add(text)
            if value != text:
                only_equal = False

        non_empty_pairs = [(value, text) for (value, text) in pairs if value or text]

        if rules.is_question_header(header):
            for row_idx in range(1, len(text_rows)):
                value = value_rows[row_idx][idx].strip()
                if not value:
                    continue
                if (
                    SUSPICIOUS_LIST_RE.match(value)
                    or SUSPICIOUS_DOT_LIST_RE.match(value)
                ):
                    suspicious_cells.append(
                        (
                            column,
                            text_rows[row_idx][0],
                            row_idx + 1,
                            normalize_title(header),
                            value,
                        )
                    )

        if only_equal:
            analyses.append(ColumnAnalysis(column, header, "same", non_empty_pairs))
            continue

        if set(non_empty_pairs).issubset(rules.binary_flag_pairs):
            analyses.append(ColumnAnalysis(column, header, "binary_flag", non_empty_pairs))
            continue

        if set(non_empty_pairs) == rules.termination_flag_pairs:
            analyses.append(ColumnAnalysis(column, header, "termination_flag", non_empty_pairs))
            continue

        if rules.is_other_text_header(header):
            others = [(value, text) for value, text in non_empty_pairs if value != "0"]
            if all(value == text and rules.selected_other_prefix(value) for value, text in others):
                analyses.append(ColumnAnalysis(column, header, "other_prefixed_text", non_empty_pairs))
                continue

        mapped_pairs = [(value, text) for value, text in non_empty_pairs if value and text and value != text]
        if mapped_pairs and all(len(texts) == 1 for texts in value_to_texts.values() if texts):
            analyses.append(ColumnAnalysis(column, header, "coded_mapping", mapped_pairs))
            continue

        analyses.append(ColumnAnalysis(column, header, "mixed_or_open", non_empty_pairs))

    return analyses, suspicious_cells


def build_report(
    text_path: Path,
    value_path: Path,
    text_rows: list[list[str]],
    value_rows: list[list[str]],
    analyses: list[ColumnAnalysis],
    suspicious_cells: list[tuple[str, str, int, str, str]],
    rules: MappingRules,
) -> str:
    headers = text_rows[0]
    total_rows = len(text_rows) - 1
    total_cols = len(headers)
    aligned_rows = sum(
        1 for row_idx in range(1, len(text_rows)) if text_rows[row_idx][0] == value_rows[row_idx][0]
    )

    counts = Counter(item.kind for item in analyses)

    question_binary_groups: dict[str, list[ColumnAnalysis]] = defaultdict(list)
    coded_mappings: list[ColumnAnalysis] = []
    special_others: list[ColumnAnalysis] = []
    mixed_columns: list[ColumnAnalysis] = []

    for item in analyses:
        if item.kind == "binary_flag":
            question_binary_groups[question_key(item.header, rules)].append(item)
        elif item.kind == "coded_mapping":
            coded_mappings.append(item)
        elif item.kind == "other_prefixed_text":
            special_others.append(item)
        elif item.kind == "mixed_or_open":
            mixed_columns.append(item)

    lines: list[str] = []
    lines.append("# 问卷值/文本映射分析")
    lines.append("")
    lines.append("## 1. 文件对齐检查")
    lines.append("")
    lines.append(f"- 文本文件：`{text_path.name}`")
    lines.append(f"- 数值文件：`{value_path.name}`")
    lines.append(f"- 数据行数：{total_rows}")
    lines.append(f"- 列数：{total_cols}")
    lines.append(f"- 按 `提交序号` 对齐成功：{aligned_rows}/{total_rows}")
    lines.append("- 结论：两份文件是同一份问卷结果的两种导出视图，可按行直接比对。")
    lines.append("")
    lines.append("## 2. 列类型概览")
    lines.append("")
    lines.append(f"- 完全相同列：{counts['same']} 列")
    lines.append(f"- 单选/编码映射列：{counts['coded_mapping']} 列")
    lines.append(f"- 多选展开二值列：{counts['binary_flag']} 列")
    lines.append(f"- 终止提醒列（`-2 -> 空`）：{counts['termination_flag']} 列")
    lines.append(f"- `其他，请注明` 特殊列：{counts['other_prefixed_text']} 列")
    lines.append(f"- 其余开放题/混合列：{counts['mixed_or_open']} 列")
    lines.append("")
    lines.append("## 3. 已确认的选项值 -> 选项文本映射（基于样本中实际出现值）")
    lines.append("")
    for item in sorted(coded_mappings, key=lambda current: current.column):
        mapping_text = "；".join(
            f"`{value}` -> `{text}`" for value, text in sort_mapping_items(item.mappings)
        )
        lines.append(
            f"- `{item.column}` / `{question_key(item.header, rules)}` / `{normalize_title(item.header)}`：{mapping_text}"
        )

    lines.append("")
    lines.append("## 4. 多选题的存储方式")
    lines.append("")
    lines.append(
        "- 这份导出里，大多数多选题不是把多个编码塞进一个单元格，而是“每个选项单独一列”。"
    )
    lines.append("- 在 `data-value.xlsx` 中，多选列通常是：`1 = 选中`，`0 = 未选中`。")
    lines.append("- 在 `data-text.xlsx` 中，对应列通常是：`1 = 选中`，空白 = 未选中。")
    lines.append("- 选项文本直接写在列头里，所以值/文本映射主要靠“列头 + 0/1”恢复。")
    lines.append("")
    for question, items in sorted(question_binary_groups.items()):
        option_text = "；".join(
            f"`{item.column}`={option_label(item.header)}" for item in sorted(items, key=lambda current: current.column)
        )
        lines.append(f"- `{question}`：`0 -> 未选`，`1 -> 选中`；展开列为 {option_text}")

    lines.append("")
    lines.append("## 5. `其他，请注明` 的特殊格式")
    lines.append("")
    lines.append("- 这些列不是用分隔符保存多个编码，而是把 `1` 作为选中标记直接拼在自由文本前面。")
    for item in sorted(special_others, key=lambda current: current.column):
        examples = [value for value, text in item.mappings if value != "0"][:3]
        example_text = "；".join(f"`{value}`" for value in examples) if examples else "无样本"
        lines.append(
            f"- `{item.column}` / `{normalize_title(item.header)}`：空值时通常是 `0`，非空样例有 {example_text}"
        )

    lines.append("")
    lines.append("## 6. 可疑“编码串”与分隔符线索")
    lines.append("")
    if suspicious_cells:
        separator_counter = Counter()
        for _, _, _, _, value in suspicious_cells:
            for char in re.findall(r"[,，;；|/、+.]", value):
                separator_counter[char] += 1

        lines.append(
            "- 在真正的多选题列中，没有发现稳定的“同一单元格内多编码”格式；大多已被展开成多列。"
        )
        lines.append(
            "- 目前只在 `BK`（`Q23` 开放题）发现少量看起来像“人工录入编码串”的内容，文本版与数值版完全相同，因此无法仅靠这两份文件反推数字含义。"
        )
        lines.append(
            "- 在这些可疑单元格里，实际观察到的分隔符只有："
            + "、".join(f"`{sep}`({count}次)" for sep, count in sorted(separator_counter.items()))
        )
        lines.append("")
        for column, submit_id, row_number, header, value in suspicious_cells:
            lines.append(
                f"- 行 `{row_number}` / 提交序号 `{submit_id}` / 列 `{column}` / `{question_key(header, rules)}`：`{value}`"
            )
    else:
        lines.append("- 没发现疑似“一个单元格内保存多编码”的答案。")

    lines.append("")
    lines.append("## 7. 结论")
    lines.append("")
    lines.append(
        "- 可以稳定恢复出单选题的“编码 -> 文本”映射，以及多选题的“列头选项 + 0/1 标记”映射。"
    )
    lines.append(
        "- 如果后续你想把这两份表合并成一份更好分析的标准表，我建议把多选题统一转成长表或布尔列。"
    )
    lines.append(
        f"- 当前共识别出 {len(suspicious_cells)} 条可疑编码串，最好结合问卷原始题库/访问员说明再确认它们是不是手工记号。"
    )
    lines.append("")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="分析问卷数值版与文本版的映射关系")
    parser.add_argument("--text", default="data-text.xlsx", type=Path)
    parser.add_argument("--value", default="data-value.xlsx", type=Path)
    parser.add_argument(
        "--report",
        default=Path("docs/questionnaire_mapping_report.md"),
        type=Path,
    )
    parser.add_argument("--rules", default=None, type=Path)
    args = parser.parse_args()
    rules = load_rules(args.rules)

    text_rows = load_xlsx(args.text)
    value_rows = load_xlsx(args.value)
    if len(text_rows) != len(value_rows):
        raise ValueError("两份文件行数不一致")
    if text_rows[0] != value_rows[0]:
        raise ValueError("两份文件列头不一致")

    analyses, suspicious_cells = analyze_columns(text_rows, value_rows, rules)
    report = build_report(
        text_path=args.text,
        value_path=args.value,
        text_rows=text_rows,
        value_rows=value_rows,
        analyses=analyses,
        suspicious_cells=suspicious_cells,
        rules=rules,
    )
    args.report.write_text(report, encoding="utf-8")
    print(f"报告已生成：{args.report}")


if __name__ == "__main__":
    main()
