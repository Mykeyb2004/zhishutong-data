from __future__ import annotations

import argparse
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pandas as pd
import pyreadstat


NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
CELL_RE = re.compile(r"([A-Z]+)(\d+)")
NUMERIC_RE = re.compile(r"-?\d+(?:\.\d+)?")


@dataclass
class VariableSpec:
    source_col: str
    source_header: str
    question_id: str
    question_label: str
    option_label: str
    spss_name: str
    variable_label: str
    var_type: str
    measure: str
    role: str
    keep: bool
    missing_rule: str
    transform_rule: str
    notes: str


def col_to_idx(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + ord(char) - 64
    return value - 1


def read_xlsx_sheet(path: Path, sheet_name: str | None = None) -> list[list[str]]:
    with ZipFile(path) as workbook:
        try:
            shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(item.itertext()) for item in shared_root.findall("a:si", NS)
            ]
        except KeyError:
            shared_strings = []

        workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
        workbook_rels = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in workbook_rels.findall(
                "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
            )
        }

        sheets = list(workbook_root.find("a:sheets", NS))
        if sheet_name is None:
            if not sheets:
                raise ValueError(f"{path} 中没有可读取的 sheet")
            target = rel_map[
                sheets[0].attrib[
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                ]
            ]
        else:
            target = None
            for sheet in sheets:
                if sheet.attrib["name"] == sheet_name:
                    rel_id = sheet.attrib[
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                    ]
                    target = rel_map[rel_id]
                    break
        if target is None:
            raise ValueError(f"{path} 中不存在 sheet: {sheet_name}")

        sheet_root = ET.fromstring(workbook.read(f"xl/{target}"))
        sheet_data = sheet_root.find("a:sheetData", NS)
        if sheet_data is None:
            return []

        rows: list[list[str]] = []
        max_idx = 0
        for row in sheet_data:
            values: dict[int, str] = {}
            for cell in row:
                ref = cell.attrib["r"]
                match = CELL_RE.match(ref)
                if not match:
                    continue
                idx = col_to_idx(match.group(1))
                max_idx = max(max_idx, idx)
                cell_type = cell.attrib.get("t")
                if cell_type == "s":
                    value_node = cell.find("a:v", NS)
                    value = (
                        shared_strings[int(value_node.text)]
                        if value_node is not None
                        else ""
                    )
                elif cell_type == "inlineStr":
                    inline_node = cell.find("a:is", NS)
                    value = "".join(inline_node.itertext()) if inline_node is not None else ""
                else:
                    value_node = cell.find("a:v", NS)
                    value = value_node.text if value_node is not None else ""
                values[idx] = value
            rows.append([values.get(i, "") for i in range(max_idx + 1)])

    width = max(len(row) for row in rows) if rows else 0
    return [row + [""] * (width - len(row)) for row in rows]


def load_variable_specs(path: Path) -> list[VariableSpec]:
    rows = read_xlsx_sheet(path, "variables")
    header = rows[0]
    index = {name: position for position, name in enumerate(header)}
    specs: list[VariableSpec] = []
    for row in rows[1:]:
        if not any(row):
            continue
        specs.append(
            VariableSpec(
                source_col=row[index["source_col"]],
                source_header=row[index["source_header"]],
                question_id=row[index["question_id"]],
                question_label=row[index["question_label"]],
                option_label=row[index["option_label"]],
                spss_name=row[index["spss_name"]],
                variable_label=row[index["variable_label"]],
                var_type=row[index["var_type"]],
                measure=row[index["measure"]],
                role=row[index["role"]],
                keep=row[index["keep"]] == "1",
                missing_rule=row[index["missing_rule"]],
                transform_rule=row[index["transform_rule"]],
                notes=row[index["notes"]],
            )
        )
    return specs


def load_value_labels(path: Path) -> dict[str, list[tuple[str, str]]]:
    rows = read_xlsx_sheet(path, "value_labels")
    header = rows[0]
    index = {name: position for position, name in enumerate(header)}
    labels: dict[str, list[tuple[str, str]]] = {}
    for row in rows[1:]:
        if not any(row):
            continue
        labels.setdefault(row[index["spss_name"]], []).append(
            (row[index["code"]], row[index["label"]])
        )
    return labels


def parse_number(value: str) -> float | int | None:
    text = value.strip()
    if not text:
        return None
    if not NUMERIC_RE.fullmatch(text):
        return None
    number = float(text)
    if number.is_integer():
        return int(number)
    return number


def transform_value(raw_value: str, spec: VariableSpec) -> Any:
    text = raw_value.strip()
    if spec.transform_rule == "copy_as_numeric_0_1":
        if text == "":
            return None
        number = parse_number(text)
        return int(number) if number is not None else None
    if spec.transform_rule == "derive_1_if_value_not_in['','0']_else_0":
        return 0 if text in {"", "0"} else 1
    if spec.transform_rule == "strip_leading_1_when_present":
        if text in {"", "0"}:
            return ""
        return text[1:].strip() if text.startswith("1") else text
    if spec.transform_rule == "copy_raw_value":
        return text
    return text


def is_code_list_numeric(labels: list[tuple[str, str]] | None) -> bool:
    if not labels:
        return False
    return all(parse_number(code) is not None for code, _ in labels)


def build_dataframe(
    source_rows: list[list[str]],
    specs: list[VariableSpec],
    value_labels: dict[str, list[tuple[str, str]]],
) -> tuple[pd.DataFrame, dict[str, list[str]]]:
    headers = source_rows[0]
    data_rows = source_rows[1:]
    submit_id_idx = col_to_idx("A")
    warnings: dict[str, list[str]] = {}
    columns: dict[str, pd.Series] = {}

    for spec in specs:
        if not spec.keep:
            continue

        source_idx = col_to_idx(spec.source_col)
        raw_values = [
            row[source_idx] if source_idx < len(row) else ""
            for row in data_rows
        ]
        transformed = [transform_value(raw, spec) for raw in raw_values]

        if spec.role in {"multi_binary", "multi_other_flag"}:
            columns[spec.spss_name] = pd.Series(
                [None if value is None else int(value) for value in transformed],
                dtype="Float64",
            )
            continue

        if spec.role == "open_numeric":
            numeric_values = []
            invalid_examples: list[str] = []
            for row, value in zip(data_rows, transformed):
                if value in {"", None}:
                    numeric_values.append(pd.NA)
                    continue
                parsed = parse_number(str(value))
                if parsed is None:
                    numeric_values.append(pd.NA)
                    submit_id = row[submit_id_idx] if submit_id_idx < len(row) else "?"
                    invalid_examples.append(f"提交序号 {submit_id}: {value}")
                else:
                    numeric_values.append(parsed)
            if invalid_examples:
                warnings[spec.spss_name] = invalid_examples[:10]
            columns[spec.spss_name] = pd.Series(numeric_values, dtype="Float64")
            continue

        label_pairs = value_labels.get(spec.spss_name)
        if spec.var_type == "numeric":
            numeric_values = []
            invalid_examples: list[str] = []
            numeric_ok = is_code_list_numeric(label_pairs)
            for row, value in zip(data_rows, transformed):
                if value in {"", None}:
                    numeric_values.append(pd.NA)
                    continue
                parsed = parse_number(str(value))
                if parsed is None:
                    numeric_ok = False
                    submit_id = row[submit_id_idx] if submit_id_idx < len(row) else "?"
                    invalid_examples.append(f"提交序号 {submit_id}: {value}")
                    numeric_values.append(value)
                else:
                    numeric_values.append(parsed)
            if numeric_ok:
                columns[spec.spss_name] = pd.Series(numeric_values, dtype="Float64")
            else:
                if invalid_examples:
                    warnings[spec.spss_name] = invalid_examples[:10]
                columns[spec.spss_name] = pd.Series(
                    ["" if value is pd.NA else ("" if value is None else str(value)) for value in numeric_values],
                    dtype="string",
                )
            continue

        columns[spec.spss_name] = pd.Series(
            ["" if value is None else str(value) for value in transformed],
            dtype="string",
        )

    df = pd.DataFrame(columns)
    return df, warnings


def normalize_value_label_key(code: str, series: pd.Series) -> Any:
    if pd.api.types.is_string_dtype(series.dtype):
        return code
    number = parse_number(code)
    if number is None:
        return code
    return float(number)


def build_pyreadstat_metadata(
    df: pd.DataFrame,
    specs: list[VariableSpec],
    value_labels: dict[str, list[tuple[str, str]]],
) -> tuple[dict[str, str], dict[str, dict[Any, str]], dict[str, str], dict[str, str]]:
    kept_specs = [spec for spec in specs if spec.keep and spec.spss_name in df.columns]
    column_labels = {spec.spss_name: spec.variable_label for spec in kept_specs}
    variable_measure = {
        spec.spss_name: spec.measure if spec.measure in {"nominal", "ordinal", "scale"} else "nominal"
        for spec in kept_specs
    }
    variable_format: dict[str, str] = {}
    variable_value_labels: dict[str, dict[Any, str]] = {}

    for spec in kept_specs:
        series = df[spec.spss_name]
        if pd.api.types.is_string_dtype(series.dtype):
            variable_format[spec.spss_name] = f"A{max(8, int(series.astype(str).map(len).max() or 1))}"
        elif spec.role in {"multi_binary", "multi_other_flag"}:
            variable_format[spec.spss_name] = "F1.0"
        elif spec.role == "open_numeric":
            variable_format[spec.spss_name] = "F8.2"
        else:
            variable_format[spec.spss_name] = "F8.2"

        pairs = value_labels.get(spec.spss_name)
        if not pairs:
            continue
        variable_value_labels[spec.spss_name] = {
            normalize_value_label_key(code, series): label for code, label in pairs
        }

    return column_labels, variable_value_labels, variable_measure, variable_format


def preview_rows(df: pd.DataFrame, columns: list[str], limit: int = 5) -> list[list[str]]:
    records: list[list[str]] = []
    subset = df[columns].head(limit)
    for _, row in subset.iterrows():
        records.append(
            [
                "" if (value is None or (isinstance(value, float) and math.isnan(value))) else str(value)
                for value in row.tolist()
            ]
        )
    return records


def choose_preview_columns(specs: list[VariableSpec], df: pd.DataFrame) -> list[str]:
    role_priority = [
        "meta",
        "single",
        "multi_binary",
        "multi_other_flag",
        "multi_other_text",
        "open_text",
        "open_numeric",
        "flow",
    ]
    chosen: list[str] = []
    seen: set[str] = set()
    kept_specs = [spec for spec in specs if spec.keep and spec.spss_name in df.columns]

    for role in role_priority:
        for spec in kept_specs:
            if spec.role != role or spec.spss_name in seen:
                continue
            chosen.append(spec.spss_name)
            seen.add(spec.spss_name)
            break

    for spec in kept_specs:
        if spec.spss_name in seen:
            continue
        chosen.append(spec.spss_name)
        seen.add(spec.spss_name)
        if len(chosen) >= 8:
            break

    return chosen[:8]


def write_preview_report(
    path: Path,
    sav_path: Path,
    df: pd.DataFrame,
    specs: list[VariableSpec],
    value_labels: dict[str, list[tuple[str, str]]],
    warnings: dict[str, list[str]],
    metadata: Any,
) -> None:
    kept_specs = [spec for spec in specs if spec.keep and spec.spss_name in df.columns]
    key_names = choose_preview_columns(specs, df)
    preview = preview_rows(df, key_names) if key_names else []

    lines = [
        "# SAV 转换预览",
        "",
        f"- 输出文件：`{sav_path.name}`",
        f"- 样本数：{len(df)}",
        f"- 变量数：{len(df.columns)}",
        f"- `pyreadstat` 读回变量数：{len(metadata.column_names)}",
        f"- 变量标签数：{len(metadata.column_labels)}",
        f"- 值标签变量数：{len(metadata.variable_value_labels)}",
        "",
        "## 关键变量预览",
        "",
    ]
    if key_names:
        lines.append("| " + " | ".join(key_names) + " |")
        lines.append("| " + " | ".join(["---"] * len(key_names)) + " |")
        for row in preview:
            lines.append("| " + " | ".join(item.replace("\n", " ") for item in row) + " |")
    else:
        lines.append("- 无可预览变量")

    lines.extend(["", "## 变量标签示例", ""])
    for spec in kept_specs[:12]:
        lines.append(f"- `{spec.spss_name}` -> `{spec.variable_label}`")

    lines.extend(["", "## 值标签示例", ""])
    for name in [spec.spss_name for spec in kept_specs if spec.spss_name in value_labels][:8]:
        pairs = "；".join(f"`{code}` -> `{label}`" for code, label in value_labels[name][:10])
        lines.append(f"- `{name}`：{pairs}")

    lines.extend(["", "## 转换警告", ""])
    if warnings:
        for name, items in warnings.items():
            lines.append(f"- `{name}`：")
            for item in items:
                lines.append(f"  - {item}")
    else:
        lines.append("- 无")

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="将现有问卷 Excel 转换为 SAV")
    parser.add_argument("--data", type=Path, default=Path("data-value.xlsx"))
    parser.add_argument("--sheet", type=str, default=None)
    parser.add_argument("--mapping", type=Path, default=Path("docs/sav_mapping_template.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("data-value.converted.sav"))
    parser.add_argument("--preview", type=Path, default=Path("docs/sav_conversion_preview.md"))
    args = parser.parse_args()

    source_rows = read_xlsx_sheet(args.data, args.sheet)
    specs = load_variable_specs(args.mapping)
    value_labels = load_value_labels(args.mapping)

    df, warnings = build_dataframe(source_rows, specs, value_labels)
    column_labels, variable_value_labels, variable_measure, variable_format = build_pyreadstat_metadata(
        df,
        specs,
        value_labels,
    )

    pyreadstat.write_sav(
        df,
        args.output,
        file_label="Converted from questionnaire Excel",
        column_labels=column_labels,
        variable_value_labels=variable_value_labels,
        variable_measure=variable_measure,
        variable_format=variable_format,
    )

    _, metadata = pyreadstat.read_sav(args.output, apply_value_formats=False)
    args.preview.parent.mkdir(parents=True, exist_ok=True)
    write_preview_report(
        args.preview,
        args.output,
        df,
        specs,
        value_labels,
        warnings,
        metadata,
    )
    print(f"SAV 已生成：{args.output}")
    print(f"预览报告：{args.preview}")


if __name__ == "__main__":
    main()
