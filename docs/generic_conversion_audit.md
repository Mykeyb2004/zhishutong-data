# 通用化审计：现有 Excel -> SAV 转换中的定制逻辑

本文档用于盘点当前代码里，哪些逻辑仍然偏向当前这组问卷 Excel，而不是“任意同类数据都可复用”的通用转换器。

## 1. 直接针对当前文件/当前导出格式的逻辑

### 1.1 固定读取 `sheet1.xml`
- 文件：`src/analyze_questionnaire_mapping.py`
- 问题：`load_xlsx()` 直接读取 `xl/worksheets/sheet1.xml`
- 影响：如果 Excel 的有效数据不在第一个 sheet，分析脚本会读错
- 结论：这是当前文件结构假设，不够通用

### 1.2 默认输入/输出文件名写死
- 文件：`src/analyze_questionnaire_mapping.py`
- 默认值：`data-text.xlsx`、`data-value.xlsx`、`docs/questionnaire_mapping_report.md`
- 文件：`src/generate_sav_mapping_template.py`
- 默认值：`data-text.xlsx`、`data-value.xlsx`、`docs/sav_mapping_template.xlsx`
- 文件：`src/convert_excel_to_sav.py`
- 默认值：`data-value.xlsx`、`docs/sav_mapping_template.xlsx`、`data-value.converted.sav`、`docs/sav_conversion_preview.md`
- 影响：虽然可通过参数覆盖，但默认工作流仍然强绑定当前项目文件名

### 1.3 报告文案写入了当前数据结论
- 文件：`src/analyze_questionnaire_mapping.py`
- 问题：
  - 报告正文直接写“这份导出里，大多数多选题……”
  - 写死 `data-value.xlsx` / `data-text.xlsx`
  - 写死 `BK` / `Q23` 的当前样本结论
- 影响：分析报告生成逻辑不是通用模板，而是掺杂了当前工作簿观察结果

### 1.4 默认把第 1 列当作“提交序号/样本标识”
- 文件：`src/analyze_questionnaire_mapping.py`
- 问题：对齐统计和可疑值记录都使用第 1 列作为样本标识
- 文件：`src/convert_excel_to_sav.py`
- 问题：转换警告中固定使用 A 列作为样本定位信息
- 影响：如果别的 Excel 第一列不是提交序号，这些提示会失真

### 1.5 列头解析依赖当前问卷导出格式
- 文件：`src/analyze_questionnaire_mapping.py`
- 问题：`option_label()` 用第一个 `_` 之后的内容当选项文本
- 文件：`src/generate_sav_mapping_template.py`
- 问题：`build_question_label()` 也按 `_` 截断题干
- 影响：依赖当前列头形如“题干_选项”的结构；换一种导出列头就可能失效

### 1.6 `1文本` 这种“其他，请注明”格式仍按当前样本规则处理
- 文件：`src/generate_sav_mapping_template.py`
- 文件：`src/convert_excel_to_sav.py`
- 问题：
  - 默认把 `1文本` 拆成“是否填写=1 + 文本内容”
  - 默认清洗规则是“去掉前导 `1`”
- 影响：这不是通用 Excel 规律，而是当前问卷导出约定

### 1.7 示例文本仍写死为 `1KTV`
- 文件：`src/generate_sav_mapping_template.py`
- 问题：模板备注里直接写示例 `1KTV -> KTV`
- 影响：虽不影响执行，但暴露当前样本特征

## 2. 已经外置，但默认规则仍明显偏向当前问卷

### 2.1 题号识别默认只认 `Q\d+`
- 文件：`src/mapping_rules.json`
- 当前默认：`question_id_regex = ^(Q\\d+)`
- 影响：适合当前问卷，不适合没有 `Qxx` 编号、或题号在别的位置的表

### 2.2 问题列识别默认只认以 `Q\d+` 开头的列头
- 文件：`src/mapping_rules.json`
- 当前默认：`question_header_patterns = ["^(Q\\d+)"]`
- 影响：非此格式的问卷列头会被误判为 `meta`

### 2.3 “开放题 / 其他，请注明 / 以上均无”关键词是中文且偏当前业务
- 文件：`src/mapping_rules.json`
- 当前默认：
  - `开放题`
  - `其他，请注明`
  - `以上均无`
  - `都没有喝过`
- 影响：虽然已可配置，但默认仍是当前问卷家族规则

### 2.4 子题 token 规则偏向当前 `14a:` 这类写法
- 文件：`src/mapping_rules.json`
- 当前默认：`([0-9]+[A-Za-z])[:：]`
- 影响：适合当前 Q27 一类子题命名，不适合其它命名体系

### 2.5 二值列与终止列的值对规则是当前样本归纳
- 文件：`src/mapping_rules.json`
- 当前默认：
  - 多选展开：`0 -> 空`、`1 -> 1`
  - 终止列：`-2 -> 空`
- 影响：这是当前问卷常见格式，不是所有导出都如此

## 3. 不是“当前文件专属”，但仍不是通用表格转换

这一组更准确地说是“问卷域假设”，不是当前这份 Excel 独有，但如果目标是“通用型 Excel -> SAV”，它们仍然算限制。

### 3.1 整套模型默认是“问卷题目”而不是“普通表格”
- 文件：`src/generate_sav_mapping_template.py`
- 体现：
  - `question_id`
  - `question_label`
  - `option_label`
  - `role = single / multi_binary / open_text / flow`
  - `mrsets`
- 影响：这更像“问卷 Excel -> SPSS 问卷数据集”转换器，而不是任意 Excel 转 SAV

### 3.2 多选题默认按 SPSS `multiple_dichotomy` 设计
- 文件：`src/generate_sav_mapping_template.py`
- 当前行为：
  - 生成 `mrsets`
  - 固定 `set_type = multiple_dichotomy`
  - 固定 `counted_value = 1`
- 影响：适合当前“多列 0/1”问卷结构，不适合“单列多编码”或其它多响应表示法

### 3.3 变量命名规则默认围绕问卷题号
- 文件：`src/generate_sav_mapping_template.py`
- 当前行为：
  - `q6`
  - `q20_01`
  - `q23_text`
  - `meta_a`
- 影响：对问卷很合理，但对普通业务表格并不通用

### 3.4 `98/99` 命名保留位是假定的问卷习惯
- 文件：`src/generate_sav_mapping_template.py`
- 当前行为：
  - `其他` -> `_99`
  - `以上均无` -> `_98`
- 影响：这是常见问卷编码习惯，但并不是通用规则

### 3.5 类型推断仍是启发式，不是数据字典驱动
- 文件：`src/generate_sav_mapping_template.py`
- 当前行为：
  - 用数值占比推断 numeric/string
  - `open_numeric` 采用 `0.7` 阈值
  - 混合数值/文本时整列按文本
- 影响：这解决了当前 Q27 问题，但本质仍是启发式

### 3.6 数值解析能力仍偏窄
- 文件：`src/generate_sav_mapping_template.py`
- 文件：`src/convert_excel_to_sav.py`
- 当前仅支持：纯整数/小数
- 不直接支持：
  - 千分位
  - 货币符号
  - 百分比
  - 科学计数法
  - 日期/时间型 Excel 单元格
- 影响：不是当前样本硬编码，但会限制通用性

## 4. 哪些逻辑已经算“够通用”

以下部分我认为已经不属于“只为当前 Excel 写死”的逻辑：

- 通过文本版 / 数值版逐列比对，反推值标签映射
- 单选题 `code -> label` 的样本内推断
- 识别“多选题已拆成多列二值变量”的思路
- “同一列混合数值与文本时整列按文本处理”的类型保护规则
- 基于映射模板再生成 SAV，而不是直接把当前列名硬塞进 SAV

## 5. 按优先级排序的去定制化建议

### P1：先去掉当前文件专属逻辑
1. 让 `analyze_questionnaire_mapping.py` 支持按 sheet 名或首个可见 sheet 读取，而不是固定 `sheet1.xml`
2. 把第 1 列样本标识改为可配置，或仅在存在显式 ID 列时使用
3. 把 `option_label()` / `build_question_label()` 的列头拆解规则外置到配置
4. 把 `1文本` 类清洗规则完全外置到规则文件，不在代码里假设前导 `1`
5. 移除报告生成里的 `BK` / `Q23` 等当前数据叙述

### P2：把“当前问卷默认规则”升级为“可切换 profile”
1. 为 `src/mapping_rules.json` 增加 profile 概念
2. 当前这套规则作为 `questionnaire_cn_default`
3. 未来可再增加：
   - `questionnaire_en_default`
   - `plain_tabular_default`
   - `single_sheet_codelist_default`

### P3：真正走向通用型转换器
1. 把“问卷专用模式”和“普通表格模式”分开
2. 问卷模式保留：
   - 题号
   - 值标签
   - 多选题
   - MRSETS
3. 普通表格模式提供：
   - 列名直映射
   - 类型推断
   - 可选值标签表
   - 不强依赖 `Qxx`

## 6. 总结

结论分两层：

1. **如果你问“还有没有只为当前 Excel 写的逻辑”**：有，主要集中在 `sheet1`、首列样本 ID、列头 `_` 结构、`1文本` 清洗、默认文件名、报告文案。
2. **如果你问“现在是不是通用型转换器”**：还不是。现在更准确地说是一个“可配置但仍以问卷数据为中心的 Excel -> SAV 转换器”。

下一步若要继续，我建议先做 **P1 去硬编码**，再做 **P2 profile 化**。这样不会一下子把代码改得过重，但能明显提升复用性。
