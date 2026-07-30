# PRD 完成度审计

审计日期：2026-07-15（Asia/Shanghai）

## 审计口径

权威需求来源为 `K12.md` 第 17 章的 10 组、96 项验收标准。`/api/acceptance-matrix` 是逐项证据索引，但只有成功产物、完整性校验、真实外部作业或原生执行报告才能证明对应能力完成。合同自检、dry-run 和空历史下没有发现错误，都不等于真实执行完成。

## 当前基线

- 验收项：96。
- 已覆盖：88。
- 合同覆盖或需真实环境验证：8。
- 规划中：0。
- 实现阻断：0。
- 顶层 `uncovered_risks` 必须包含全部 8 项未完成真实验收的要求，不得只统计 Mathpix 和桌面端两类显式状态。

## 未完成真实验收项

| PRD | 要求 | 已有证据 | 尚缺证据 |
| --- | --- | --- | --- |
| 17.3.3 | OMML 转 MathType | 依赖复制、动作队列和桌面计划 | 同平台 MathType 对象写回 |
| 17.3.6 | 复制依赖后重试转换 | `retry_conversion` 和 `omml.retry_conversion` 合同 | 依赖补齐后的真实 MathType 重试成功产物 |
| 17.4.5 | 按顺序执行 Word 宏 | 宏队列、顺序和 dry-run | Word VBA 真实执行报告 |
| 17.6.2 | 扫描 PDF OCR | Mathpix 请求与输出合同 | 授权上传、轮询、DOCX 下载与哈希 |
| 17.6.5 | PDF 数学公式识别 | 公式 OCR 和 `tex.zip` 请求合同 | 真实 Mathpix 作业、`tex.zip` 与公式报告 |
| 17.7.2 | Word 原生公式转 MathType | OMML 交接合同 | 同平台 MathType 对象写回 |
| 17.7.4 | PDF 公式识别并转 MathType | Mathpix 合同和 `pdf_formula_mathtype` 桌面计划 | Mathpix 识别证据与同平台 MathType 写回证据 |
| 17.10.7 | 本地端执行 Office、MathType、OMML、Word 宏 | macOS Office 真实规范化和旧 PPT 二段转换；服务端闭环回归 | MathType、OMML 写回和 Word 宏原生执行 |

## 动态证据规则

当任务保存经过服务端脱敏和校验的 `k12.localNativeExecutionReport.v1` 时，矩阵只升级对应动作。只有 `office_conversion` 证据时，17.10.7 必须显示“部分实测”，并继续列出 MathType、OMML 和 Word 宏缺口。只有 Office 动作不能证明该复合要求已覆盖。

## 完成判定

当前代码范围内没有“规划中”或“实现阻断”项，但项目整体目标仍不能标记完成：上述 8 项缺少与要求范围相匹配的真实环境证据。只有矩阵在可复核的真实任务历史下达到 96/96，且全量回归、产物完整性和隐私门禁均通过，才可宣布 PRD 完成。
