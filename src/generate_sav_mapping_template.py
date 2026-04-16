from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

from analyze_questionnaire_mapping import (
    ColumnAnalysis,
    analyze_columns,
    idx_to_col,
    load_xlsx,
    normalize_title,
    option_label,
    question_key,
)
from mapping_rules import MappingRules, load_rules


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NUMERIC_RE = re.compile(r"-?\d+(?:\.\d+)?")


@dataclass
class VariableRow:
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
    keep: str
    missing_rule: str
    transform_rule: str
    notes: str


def build_question_label(header: str) -> str:
    normalized = normalize_title(header)
    if "_" in normalized:
        return normalized.split("_", 1)[0].strip()
    return normalized


def parse_number(text: str) -> float | int | None:
    value = text.strip()
    if not value or not NUMERIC_RE.fullmatch(value):
        return None
    number = float(value)
    return int(number) if number.is_integer() else number


def numeric_ratio(values: list[str]) -> float:
    cleaned = [value.strip() for value in values if value.strip()]
    if not cleaned:
        return 0.0
    numeric_count = sum(1 for value in cleaned if parse_number(value) is not None)
    return numeric_count / len(cleaned)


def has_mixed_numeric_and_text(values: list[str]) -> bool:
    cleaned = [value.strip() for value in values if value.strip()]
    if not cleaned:
        return False
    has_numeric = any(parse_number(value) is not None for value in cleaned)
    has_text = any(parse_number(value) is None for value in cleaned)
    return has_numeric and has_text


def codes_are_numeric(pairs: list[tuple[str, str]]) -> bool:
    if not pairs:
        return False
    return all(parse_number(code) is not None for code, _ in pairs)


def split_other_prefixed(value: str) -> str:
    if value.startswith("1"):
        return value[1:].strip()
    return value.strip()


def extract_subquestion_token(header: str, rules: MappingRules) -> str:
    matches = rules.extract_subquestion_tokens(header)
    if matches:
        return matches[-1].lower()
    return ""


def ensure_unique_name(base: str, used_names: set[str]) -> str:
    candidate = base
    counter = 2
    while candidate in used_names:
        candidate = f"{base}_{counter}"
        counter += 1
    used_names.add(candidate)
    return candidate


def infer_name(
    question_id: str,
    role: str,
    column: str,
    header: str,
    used_names: set[str],
    rules: MappingRules,
) -> str:
    if question_id == "META":
        return ensure_unique_name(f"meta_{column.lower()}", used_names)

    base = question_id.lower()
    if role == "flow":
        base = f"{base}_flow"
    elif role == "open_text":
        base = f"{base}_text"
    elif role == "open_numeric":
        token = extract_subquestion_token(header, rules)
        base = f"{base}_{token}" if token else f"{base}_num"
    return ensure_unique_name(base, used_names)


def infer_role_and_type(
    analysis: ColumnAnalysis,
    header: str,
    raw_values: list[str],
    is_question: bool,
    rules: MappingRules,
) -> tuple[str, str]:
    number_ratio = numeric_ratio(raw_values)
    mixed_numeric_text = has_mixed_numeric_and_text(raw_values)

    if not is_question:
        return "meta", ("string" if mixed_numeric_text else ("numeric" if number_ratio == 1.0 and raw_values else "string"))

    if analysis.kind == "termination_flag":
        return "flow", "numeric"

    if analysis.kind == "coded_mapping":
        return "single", ("numeric" if codes_are_numeric(analysis.mappings) else "string")

    if analysis.kind == "same":
        if rules.is_open_text_header(header):
            return ("open_text", "string") if mixed_numeric_text else (("open_numeric", "numeric") if number_ratio >= 0.7 else ("open_text", "string"))
        if mixed_numeric_text:
            return "open_text", "string"
        return ("single", "numeric") if number_ratio == 1.0 and raw_values else ("open_text", "string")

    if analysis.kind == "mixed_or_open":
        return ("open_text", "string") if mixed_numeric_text else (("open_numeric", "numeric") if number_ratio >= 0.7 else ("open_text", "string"))

    if mixed_numeric_text:
        return "single", "string"
    return "single", ("numeric" if number_ratio == 1.0 and raw_values else "string")


def multi_option_base_name(
    question_id: str,
    option_idx: int,
    option_text: str,
    rules: MappingRules,
) -> str:
    prefix = question_id.lower()
    if rules.is_other_text_header(option_text):
        return f"{prefix}_99"
    if rules.is_none_of_above(option_text):
        return f"{prefix}_98"
    return f"{prefix}_{option_idx:02d}"


def suggest_measure(role: str) -> str:
    if role == "open_numeric":
        return "scale"
    return "nominal"


def suggest_keep(role: str) -> str:
    return "0" if role == "flow" else "1"


def build_variable_rows(
    analyses: list[ColumnAnalysis],
    headers: list[str],
    value_rows: list[list[str]],
    rules: MappingRules,
) -> list[VariableRow]:
    analysis_by_col = {item.column: item for item in analyses}
    multi_positions: dict[str, int] = {}
    used_names: set[str] = set()
    rows: list[VariableRow] = []

    for idx, header in enumerate(headers):
        column = idx_to_col(idx)
        analysis = analysis_by_col[column]
        normalized = normalize_title(header)
        qid = question_key(header, rules) if rules.is_question_header(header) else "META"
        qlabel = build_question_label(header)
        raw_values = [
            row[idx] if idx < len(row) else ""
            for row in value_rows[1:]
        ]

        if analysis.kind == "binary_flag":
            option_text = option_label(header)
            multi_positions[qid] = multi_positions.get(qid, 0) + 1
            base_name = ensure_unique_name(
                multi_option_base_name(qid, multi_positions[qid], option_text, rules),
                used_names,
            )
            rows.append(
                VariableRow(
                    source_col=column,
                    source_header=normalized,
                    question_id=qid,
                    question_label=qlabel,
                    option_label=option_text,
                    spss_name=base_name,
                    variable_label=normalized,
                    var_type="numeric",
                    measure=suggest_measure("multi_binary"),
                    role="multi_binary",
                    keep=suggest_keep("multi_binary"),
                    missing_rule="",
                    transform_rule="copy_as_numeric_0_1",
                    notes="多选题拆列；0=未选，1=选中",
                )
            )
            continue

        if analysis.kind == "other_prefixed_text":
            option_text = option_label(header)
            multi_positions[qid] = multi_positions.get(qid, 0) + 1
            base_name = ensure_unique_name(
                multi_option_base_name(qid, multi_positions[qid], option_text, rules),
                used_names,
            )
            rows.append(
                VariableRow(
                    source_col=column,
                    source_header=normalized,
                    question_id=qid,
                    question_label=qlabel,
                    option_label=option_text,
                    spss_name=base_name,
                    variable_label=f"{normalized}_是否填写",
                    var_type="numeric",
                    measure=suggest_measure("multi_other_flag"),
                    role="multi_other_flag",
                    keep=suggest_keep("multi_other_flag"),
                    missing_rule="",
                    transform_rule="derive_1_if_value_not_in['','0']_else_0",
                    notes="由原始的 `1文本` 形式派生为二值变量",
                )
            )
            rows.append(
                VariableRow(
                    source_col=column,
                    source_header=normalized,
                    question_id=qid,
                    question_label=qlabel,
                    option_label=f"{option_text}_文本",
                    spss_name=ensure_unique_name(f"{base_name}_text", used_names),
                    variable_label=f"{normalized}_填写文本",
                    var_type="string",
                    measure=suggest_measure("multi_other_text"),
                    role="multi_other_text",
                    keep=suggest_keep("multi_other_text"),
                    missing_rule="",
                    transform_rule="strip_leading_1_when_present",
                    notes=f"示例：`1KTV` -> `{split_other_prefixed('1KTV')}`",
                )
            )
            continue

        role, var_type = infer_role_and_type(
            analysis=analysis,
            header=normalized,
            raw_values=raw_values,
            is_question=rules.is_question_header(header),
            rules=rules,
        )
        rows.append(
            VariableRow(
                source_col=column,
                source_header=normalized,
                question_id=qid,
                question_label=qlabel,
                option_label="",
                spss_name=infer_name(qid, role, column, normalized, used_names, rules),
                variable_label=normalized,
                var_type=var_type,
                measure=suggest_measure(role),
                role=role,
                keep=suggest_keep(role),
                missing_rule="-2=user_missing" if analysis.kind == "termination_flag" else "",
                transform_rule="copy_raw_value",
                notes="",
            )
        )

    return rows


def build_value_label_rows(analyses: list[ColumnAnalysis], variable_rows: list[VariableRow]) -> list[list[str]]:
    analysis_by_col = {item.column: item for item in analyses}
    rows: list[list[str]] = []

    for variable in variable_rows:
        analysis = analysis_by_col[variable.source_col]
        if variable.role in {"multi_binary", "multi_other_flag"}:
            rows.append([variable.spss_name, "0", "未选", variable.source_col, variable.question_id])
            rows.append([variable.spss_name, "1", "选中", variable.source_col, variable.question_id])
            continue

        if analysis.kind == "coded_mapping":
            for value, label in analysis.mappings:
                rows.append([variable.spss_name, value, label, variable.source_col, variable.question_id])
            continue

        if analysis.kind == "termination_flag":
            rows.append([variable.spss_name, "-2", "流程终止/空白", variable.source_col, variable.question_id])

    return rows


def build_mrsets(variable_rows: list[VariableRow]) -> list[list[str]]:
    grouped: dict[str, list[VariableRow]] = {}
    for row in variable_rows:
        if row.role in {"multi_binary", "multi_other_flag"}:
            grouped.setdefault(row.question_id, []).append(row)

    mrsets: list[list[str]] = []
    for question_id, items in sorted(grouped.items()):
        items.sort(key=lambda item: item.spss_name)
        mrsets.append(
            [
                f"{question_id.lower()}_mr",
                items[0].question_label,
                question_id,
                "multiple_dichotomy",
                "1",
                ",".join(item.spss_name for item in items),
                "建议用于 SPSS Multiple Response Set",
            ]
        )
    return mrsets


def readme_rows() -> list[list[str]]:
    return [
        ["sheet", "说明"],
        ["README", "模板说明与使用建议"],
        ["variables", "一行代表一个输出到 SAV 的变量；可由同一 source_col 派生多个变量"],
        ["value_labels", "为 categorical 变量补充值标签"],
        ["mrsets", "为多选题预留 Multiple Response Set 定义"],
        ["", ""],
        ["原则", "说明"],
        ["自动命名", "默认按 question_id / 列角色推断，例如 q13、q23_text、q20_01"],
        ["自动类型", "默认根据列值和映射关系推断 numeric/string，可人工调整"],
        ["自动保留", "除 flow 变量外默认 keep=1，可人工修改"],
        ["", ""],
        ["字段", "说明"],
        ["source_col", "原始 Excel 列号"],
        ["source_header", "原始列头"],
        ["question_id", "问卷题号，如 Q20"],
        ["spss_name", "建议的 SAV 变量名，可人工调整"],
        ["variable_label", "建议的 SAV 变量标签"],
        ["var_type", "numeric/string"],
        ["measure", "nominal/ordinal/scale"],
        ["role", "single/multi_binary/open_text/meta 等"],
        ["keep", "1=输出到最终 SAV；0=默认不输出"],
        ["transform_rule", "脚本执行字段转换时使用的规则标记"],
        ["notes", "补充说明"],
    ]


def col_letter(index: int) -> str:
    result = []
    current = index + 1
    while current:
        current, remainder = divmod(current - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def xml_cell(ref: str, value: object) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"><v>{value}</v></c>'
    text = escape(str(value)).replace("\n", "&#10;")
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def worksheet_xml(rows: list[list[object]]) -> str:
    row_xml: list[str] = []
    for row_idx, row in enumerate(rows, start=1):
        cells = []
        for col_idx, value in enumerate(row):
            cell = xml_cell(f"{col_letter(col_idx)}{row_idx}", value)
            if cell:
                cells.append(cell)
        row_xml.append(f'<row r="{row_idx}">{"".join(cells)}</row>')
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS_MAIN}"><sheetData>{"".join(row_xml)}</sheetData></worksheet>'
    )


def content_types_xml(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides)
        + "</Types>"
    )


def root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def workbook_xml(sheet_names: Iterable[str]) -> str:
    sheets = []
    for index, name in enumerate(sheet_names, start=1):
        sheets.append(f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}"><sheets>'
        + "".join(sheets)
        + "</sheets></workbook>"
    )


def workbook_rels_xml(sheet_count: int) -> str:
    rels = []
    for index in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{index}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels)
        + "</Relationships>"
    )


def app_xml(sheet_names: Iterable[str]) -> str:
    names = list(sheet_names)
    titles = "".join(f"<vt:lpstr>{escape(name)}</vt:lpstr>" for name in names)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>Codex</Application>"
        f"<TitlesOfParts><vt:vector size=\"{len(names)}\" baseType=\"lpstr\">{titles}</vt:vector></TitlesOfParts>"
        f"<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(names)}</vt:i4></vt:variant></vt:vector></HeadingPairs>"
        "</Properties>"
    )


def core_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:title>SAV Mapping Template</dc:title>"
        "<dc:creator>Codex</dc:creator>"
        "</cp:coreProperties>"
    )


def write_xlsx(path: Path, sheets: list[tuple[str, list[list[object]]]]) -> None:
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        workbook.writestr("_rels/.rels", root_rels_xml())
        workbook.writestr("xl/workbook.xml", workbook_xml(name for name, _ in sheets))
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        workbook.writestr("docProps/app.xml", app_xml(name for name, _ in sheets))
        workbook.writestr("docProps/core.xml", core_xml())
        for index, (_, rows) in enumerate(sheets, start=1):
            workbook.writestr(f"xl/worksheets/sheet{index}.xml", worksheet_xml(rows))


def main() -> None:
    parser = argparse.ArgumentParser(description="生成 SAV 映射 Excel 模板")
    parser.add_argument("--text", type=Path, default=Path("data-text.xlsx"))
    parser.add_argument("--value", type=Path, default=Path("data-value.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("docs/sav_mapping_template.xlsx"))
    parser.add_argument("--rules", default=None, type=Path)
    args = parser.parse_args()
    rules = load_rules(args.rules)

    text_rows = load_xlsx(args.text)
    value_rows = load_xlsx(args.value)
    if text_rows[0] != value_rows[0]:
        raise ValueError("文本版与数值版列头不一致")

    analyses, _ = analyze_columns(text_rows, value_rows, rules)
    headers = text_rows[0]
    variable_rows = build_variable_rows(analyses, headers, value_rows, rules)

    variables_sheet = [[
        "source_col",
        "source_header",
        "question_id",
        "question_label",
        "option_label",
        "spss_name",
        "variable_label",
        "var_type",
        "measure",
        "role",
        "keep",
        "missing_rule",
        "transform_rule",
        "notes",
    ]]
    variables_sheet.extend([
        [
            row.source_col,
            row.source_header,
            row.question_id,
            row.question_label,
            row.option_label,
            row.spss_name,
            row.variable_label,
            row.var_type,
            row.measure,
            row.role,
            row.keep,
            row.missing_rule,
            row.transform_rule,
            row.notes,
        ]
        for row in variable_rows
    ])

    value_labels_sheet = [["spss_name", "code", "label", "source_col", "question_id"]]
    value_labels_sheet.extend(build_value_label_rows(analyses, variable_rows))

    mrsets_sheet = [["set_name", "set_label", "question_id", "set_type", "counted_value", "variables", "notes"]]
    mrsets_sheet.extend(build_mrsets(variable_rows))

    sheets = [
        ("README", readme_rows()),
        ("variables", variables_sheet),
        ("value_labels", value_labels_sheet),
        ("mrsets", mrsets_sheet),
    ]
    args.output.parent.mkdir(parents=True, exist_ok=True)
    write_xlsx(args.output, sheets)
    print(f"模板已生成：{args.output}")


if __name__ == "__main__":
    main()
