# 问卷 Excel 转 SAV 工具

本项目用于把问卷平台导出的 Excel 数据转换为 SPSS 可读取的 `.sav` 文件。

当前工具面向的是**同一类问卷导出格式**，不是任意 Excel 表格转换器。已确认并接受的业务前提是：原始数据文件中的题号体系和列结构稳定，题目列以 `Q数字` 形式开头，多选题在 Excel 中已按“每个选项一列”展开。

## 1. 核心业务前提

- `data-value.xlsx`：数值版问卷结果，是生成 `.sav` 的主数据源。
- `data-text.xlsx`：文本版问卷结果，用于反推选项值、选项文本、变量标签。
- 两个 Excel 文件应来自同一份问卷结果，行顺序一致，列头一致。
- 问题列默认以 `Q数字` 开头，例如 `Q6`、`Q20`、`Q38`。
- 多选题不是存成“一个单元格多个编码”，而是已经拆成多个二值列。
- 多选题在数值版中通常是 `1=选中`、`0=未选中`；文本版中通常是 `1=选中`、空白=未选中。
- 如果某一列同时出现数值和文本，整列按文本变量处理，不强行转数值。
- `其他，请注明` 类字段可能出现 `1KTV` 这类格式，默认拆成“是否填写”二值变量和“填写文本”字符串变量。

## 2. 目录约定

```text
.
├── src/
│   ├── analyze_questionnaire_mapping.py      # 分析文本版/数值版映射关系
│   ├── generate_sav_mapping_template.py      # 生成 SAV 映射模板
│   ├── convert_excel_to_sav.py               # 按映射模板转换为 SAV
│   ├── mapping_rules.py                      # 规则加载代码
│   └── mapping_rules.json                    # 问卷结构识别规则
├── docs/
│   ├── questionnaire_mapping_report.md       # 映射分析报告
│   ├── sav_mapping_template.xlsx             # SAV 映射模板
│   ├── sav_conversion_preview.md             # SAV 转换预览
│   ├── sav_conversion_plan.md                # 开发方案
│   └── generic_conversion_audit.md           # 通用化审计
├── tests/                                    # 测试脚本
├── logs/                                     # 运行日志
├── data-text.xlsx                            # 示例：文本版结果
├── data-value.xlsx                           # 示例：数值版结果
└── data-value.converted.sav                  # 示例：转换输出
```

新增脚本放在 `src/`，测试脚本放在 `tests/`，技术文档放在 `docs/`，日志文件放在 `logs/`。

## 3. 快速使用

所有命令都在项目根目录运行，并使用 `uv run`。

### 3.1 分析文本版与数值版映射关系

```bash
uv run python src/analyze_questionnaire_mapping.py \
  --text data-text.xlsx \
  --value data-value.xlsx \
  --rules src/mapping_rules.json \
  --report docs/questionnaire_mapping_report.md
```

输出：

- `docs/questionnaire_mapping_report.md`

这一步会检查两份文件是否能按行、按列对应，并分析：

- 单选题的 `编码 -> 文本` 映射
- 多选题是否已拆成二值列
- `其他，请注明` 类特殊字段
- 疑似“一个单元格内多个编码”的可疑值

### 3.2 生成 SAV 映射模板

```bash
uv run python src/generate_sav_mapping_template.py \
  --text data-text.xlsx \
  --value data-value.xlsx \
  --rules src/mapping_rules.json \
  --output docs/sav_mapping_template.xlsx
```

输出：

- `docs/sav_mapping_template.xlsx`

这一步会生成可人工审核和修改的映射模板。正式转换前，建议重点检查 `variables`、`value_labels`、`mrsets` 三个工作表。

### 3.3 按映射模板转换为 SAV

```bash
uv run --with pandas --with pyreadstat python src/convert_excel_to_sav.py \
  --data data-value.xlsx \
  --mapping docs/sav_mapping_template.xlsx \
  --output data-value.converted.sav \
  --preview docs/sav_conversion_preview.md
```

输出：

- `data-value.converted.sav`
- `docs/sav_conversion_preview.md`

转换脚本会写出 `.sav`，再用 `pyreadstat` 读回校验，并生成预览报告。预览报告中如果出现“转换警告”，通常代表某个应该是数值的字段里出现了非数值文本。

### 3.4 指定 Excel 工作表

`convert_excel_to_sav.py` 支持通过 `--sheet` 指定数据源工作表：

```bash
uv run --with pandas --with pyreadstat python src/convert_excel_to_sav.py \
  --data data-value.xlsx \
  --sheet Sheet1 \
  --mapping docs/sav_mapping_template.xlsx \
  --output data-value.converted.sav \
  --preview docs/sav_conversion_preview.md
```

当前分析脚本和模板生成脚本默认读取工作簿中的第一个工作表。由于当前原始数据文件结构稳定，一般不需要额外指定。

## 4. 脚本说明

| 脚本 | 作用 | 主要输入 | 主要输出 |
| --- | --- | --- | --- |
| `src/analyze_questionnaire_mapping.py` | 比对文本版和数值版，反推值标签与多选结构 | `--text`、`--value`、`--rules` | `--report` |
| `src/generate_sav_mapping_template.py` | 生成可人工审核的 SAV 映射模板 | `--text`、`--value`、`--rules` | `--output` |
| `src/convert_excel_to_sav.py` | 按映射模板生成 `.sav` | `--data`、`--mapping` | `--output`、`--preview` |

## 5. 规则配置：`src/mapping_rules.json`

`mapping_rules.json` 控制“如何识别问卷结构”。它用于分析和模板生成阶段，不直接控制最终 `.sav` 中的变量字典。最终变量字典以 `docs/sav_mapping_template.xlsx` 为准。

当前配置如下：

```json
{
  "question_id_regex": "^(Q\\d+)",
  "question_header_patterns": [
    "^(Q\\d+)"
  ],
  "subquestion_token_regex": "([0-9]+[A-Za-z])[:：]",
  "binary_flag_pairs": [
    ["0", ""],
    ["1", "1"]
  ],
  "termination_flag_pairs": [
    ["-2", ""]
  ],
  "open_text_keywords": [
    "开放题"
  ],
  "other_text_keywords": [
    "其他，请注明"
  ],
  "none_of_above_keywords": [
    "以上均无",
    "都没有喝过"
  ],
  "other_text_selected_prefixes": [
    "1"
  ]
}
```

### 5.1 配置项解释

| 配置项 | 业务含义 | 当前规则 |
| --- | --- | --- |
| `question_id_regex` | 从列头中提取题号 | `Q数字`，例如 `Q20` |
| `question_header_patterns` | 判断哪些列属于问卷题目列 | 以 `Q数字` 开头 |
| `subquestion_token_regex` | 从列头提取子题编号，用于开放数值题命名 | 例如 `14a:` |
| `binary_flag_pairs` | 识别多选题拆列的值/文本组合 | 数值版 `0/1`，文本版空白/`1` |
| `termination_flag_pairs` | 识别流程终止或跳题标记 | `-2 -> 空白` |
| `open_text_keywords` | 判断开放题列 | 列头包含 `开放题` |
| `other_text_keywords` | 判断“其他，请注明”列 | 列头包含 `其他，请注明` |
| `none_of_above_keywords` | 判断“以上均无”类选项 | 用于变量名后缀 `_98` |
| `other_text_selected_prefixes` | 判断“其他”文本前的选中前缀 | 默认前缀为 `1` |

### 5.2 常见修改方式

如果新的同类问卷导出仍然是 `Q数字` 题号，一般不需要修改 `question_id_regex` 和 `question_header_patterns`。

如果多选题的未选/选中编码发生变化，例如平台改成 `0/1` 与 `否/是`，则调整：

```json
"binary_flag_pairs": [
  ["0", "否"],
  ["1", "是"]
]
```

如果“其他，请注明”的选项文案变化，例如列头写成 `其他（请说明）`，则调整：

```json
"other_text_keywords": [
  "其他，请注明",
  "其他（请说明）"
]
```

如果“以上均无”的选项文案变化，例如列头写成 `无以上选项`，则调整：

```json
"none_of_above_keywords": [
  "以上均无",
  "都没有喝过",
  "无以上选项"
]
```

## 6. 映射模板：`docs/sav_mapping_template.xlsx`

映射模板是最终转换的业务字典。推荐流程是：先自动生成，再人工审核，最后按模板转换。

### 6.1 `variables` 工作表

一行代表一个输出变量，部分原始列可能派生多个 SAV 变量，例如 `其他，请注明` 会派生“是否填写”和“填写文本”两个变量。

| 字段 | 说明 | 常见值 |
| --- | --- | --- |
| `source_col` | 原始 Excel 列号 | `A`、`R`、`AV` |
| `source_header` | 原始 Excel 列头 | 中文题干或元数据列名 |
| `question_id` | 问卷题号 | `Q6`、`Q20`、`META` |
| `question_label` | 题目标签 | 原始题干 |
| `option_label` | 多选题选项标签 | `广告/公共关系公司` |
| `spss_name` | SAV 变量名 | `q20_01`、`q38_99_text` |
| `variable_label` | SAV 变量标签 | SPSS 中显示的中文说明 |
| `var_type` | 变量类型 | `numeric`、`string` |
| `measure` | SPSS 测量水平 | `nominal`、`ordinal`、`scale` |
| `role` | 变量角色 | `single`、`multi_binary`、`open_text`、`meta` 等 |
| `keep` | 是否输出 | `1` 输出，`0` 不输出 |
| `missing_rule` | 缺失值规则 | 例如 `-2=user_missing` |
| `transform_rule` | 转换规则 | 见下节 |
| `notes` | 备注 | 自动推断说明或人工备注 |

审核建议：

- `spss_name` 使用短英文、数字、下划线，避免中文、空格和特殊符号。
- `variable_label` 保留完整中文含义，便于 SPSS 使用者理解。
- `var_type` 需要谨慎确认；如果同一列混合数值和文本，保持 `string`。
- `keep=0` 的变量不会进入最终 `.sav`。
- `role` 会影响转换逻辑，不建议随意修改。

### 6.2 `value_labels` 工作表

用于定义分类变量的值标签。

| 字段 | 说明 |
| --- | --- |
| `spss_name` | 对应 `variables.spss_name` |
| `code` | 数据中的编码值 |
| `label` | 编码对应的文本含义 |
| `source_col` | 来源列号 |
| `question_id` | 来源题号 |

典型规则：

- 单选题写入样本中观察到的 `编码 -> 文本`。
- 多选拆列统一写入 `0 -> 未选`、`1 -> 选中`。
- 开放文本题通常不写值标签。

### 6.3 `mrsets` 工作表

用于记录多选题的 Multiple Response Set 定义。

| 字段 | 说明 |
| --- | --- |
| `set_name` | 多响应集合名称 |
| `set_label` | 多响应集合标签 |
| `question_id` | 对应题号 |
| `set_type` | 当前为 `multiple_dichotomy` |
| `counted_value` | 当前为 `1` |
| `variables` | 参与集合的变量名列表 |
| `notes` | 备注 |

当前 `.sav` 转换脚本基于 `pyreadstat` 写入变量、变量标签和值标签。`mrsets` 工作表会保留多选题集合定义，但不保证直接写入 `.sav` 内部的 SPSS Multiple Response Set 元数据；如需在 SPSS 内完整注册 MRSETS，可基于该工作表后续生成 SPSS syntax 或接入更强的 SAV 写入工具。

## 7. 转换规则：`transform_rule`

`transform_rule` 位于映射模板的 `variables` 工作表中，用于说明每个变量如何从原始单元格转换。

| 规则 | 含义 |
| --- | --- |
| `copy_raw_value` | 原样复制，按 `var_type` 输出 |
| `copy_as_numeric_0_1` | 多选拆列使用，将 `0/1` 转为数值 |
| `derive_1_if_value_not_in['','0']_else_0` | 只要原始值非空且不是 `0`，派生为 `1`，否则为 `0` |
| `strip_leading_1_when_present` | 用于 `1文本` 格式，去掉前导 `1` 后保留填写文本 |

一般不需要手工修改 `transform_rule`。如果确实要新增业务清洗规则，需要同步修改 `src/convert_excel_to_sav.py` 中的转换逻辑。

## 8. 推荐工作流

### 8.1 首次处理一批新问卷数据

1. 运行映射分析命令，查看 `docs/questionnaire_mapping_report.md`。
2. 运行模板生成命令，得到 `docs/sav_mapping_template.xlsx`。
3. 人工审核模板中的变量名、变量类型、值标签和 `keep`。
4. 运行 SAV 转换命令，得到 `.sav` 和预览报告。
5. 打开 `docs/sav_conversion_preview.md`，确认变量数、值标签数和转换警告。
6. 用 SPSS 或兼容工具打开 `.sav` 做最终校验。

### 8.2 只替换同结构的新数据

如果问卷结构和列头不变，只是样本数据更新，可以直接复用已有 `docs/sav_mapping_template.xlsx`：

```bash
uv run --with pandas --with pyreadstat python src/convert_excel_to_sav.py \
  --data new-data-value.xlsx \
  --mapping docs/sav_mapping_template.xlsx \
  --output new-data-value.converted.sav \
  --preview docs/new-data-value.preview.md
```

### 8.3 修改规则后重新生成模板

如果修改了 `src/mapping_rules.json`，建议重新执行：

```bash
uv run python src/analyze_questionnaire_mapping.py \
  --text data-text.xlsx \
  --value data-value.xlsx \
  --rules src/mapping_rules.json \
  --report docs/questionnaire_mapping_report.md

uv run python src/generate_sav_mapping_template.py \
  --text data-text.xlsx \
  --value data-value.xlsx \
  --rules src/mapping_rules.json \
  --output docs/sav_mapping_template.xlsx
```

注意：重新生成模板会覆盖原模板。若模板中有人工修改，请先另存备份。

## 9. 日志与排错

如需保留运行日志，请写入 `logs/`：

```bash
mkdir -p logs

uv run --with pandas --with pyreadstat python src/convert_excel_to_sav.py \
  --data data-value.xlsx \
  --mapping docs/sav_mapping_template.xlsx \
  --output data-value.converted.sav \
  --preview docs/sav_conversion_preview.md \
  > logs/convert.log 2>&1
```

常见问题：

- `pyreadstat` 或 `pandas` 缺失：转换命令使用 `uv run --with pandas --with pyreadstat ...`。
- 预览报告出现转换警告：检查对应变量的 `var_type`，混合数值和文本的列应按 `string` 处理。
- 值标签不完整：当前值标签来自样本中实际出现的值；若某些选项样本中没出现，需要在 `value_labels` 工作表手工补充。
- 多选题分析异常：检查 `mapping_rules.json` 中的 `binary_flag_pairs` 是否符合当前平台导出格式。
- `其他，请注明` 文本不符合预期：检查 `other_text_keywords` 和 `other_text_selected_prefixes`。

## 10. 当前边界

- 本项目默认处理同类问卷导出格式，不以任意 Excel 表格为目标。
- 题号体系 `Q数字` 是已接受的稳定前提。
- 多选题按“拆列二值变量”处理。
- `.sav` 当前重点保证变量名、变量标签、值标签和数据类型；Multiple Response Set 定义目前保留在模板中，必要时再做增强写入。
