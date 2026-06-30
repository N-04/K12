# K12 架构说明

## 分层

K12 首版采用本地单机架构：

- `static/`：浏览器工作台，负责拖拽文件、任务创建、筛选、设置和报告查看。
- 工作台处理流程支持拖动左侧线条图标或上下移动排序，创建任务时会把 `workflowOrder` 与展示标签写入任务 `options`，本地客户端和报告可复核用户当时的流程顺序。
- `k12/server.py`：本地 HTTP API，默认监听 `127.0.0.1`，只暴露 JSON 与静态资源。
- `localSecurityToken` 为空时 API 不拦截；设置后除健康检查外的 API 都需要 `X-K12-Token`、Bearer token 或下载查询参数。
- `k12/store.py`：SQLite 存储，保存文件、任务、报告、日志、设置。
- 上传文件删除只清理 `.k12-data/uploads/` 内的缓存文件和 ZIP 解压目录，不删除外部 `file_path` 指向的原始文档。
- `/api/files/{file_id}/replace` 用新上传内容替换文件记录，保留 `file_id` 并清理旧上传缓存；若原文件是 ZIP 父文件，会清理旧解压子文件记录。
- 加密文件密码只保存在当前 `AppStore` 进程内存中，SQLite 文件记录只保存会话状态；服务重启后会重置为需要重新输入密码。
- `duplicateFileStrategy` 设置控制重复文件导入策略，支持自动重命名、跳过和覆盖。
- 宏顺序模板保存在 SQLite `macro_templates` 表中，供网页端保存、加载和删除常用宏编排。
- `k12/processor.py`：处理管线与能力路由。当前为可替换模拟实现，后续可接 FastAPI worker、Office COM、MathType、LibreOffice、OCR 等真实引擎。
- `k12/install_profiles.py`：Windows/macOS 安装画像、组件清单、任务分流和 MathType 兼容策略。
- `k12/mathpix.py`：Mathpix PDF OCR 客户端封装，用于 PDF 转 Word 的提交、状态轮询、DOCX 下载和 `tex.zip` 公式结果下载。
- PDF 上传分析会固化 `pdfType`、文本层、图片对象、公式线索、表格线索和 OCR 建议；PDF 转 Word 引擎在设置保存和读取时都会归一为 Mathpix，避免 API 写入其他引擎绕过 Mathpix 合同。Mathpix PDF 请求会把 DOCX、`tex.zip` 等真实输出格式放入 `conversion_formats`，并把 K12 OCR 语言、文字/公式/表格开关、精度和速度模式作为审计元数据记录在作业中。文字、公式、表格 OCR 全部关闭时，PDF 转 Word 会阻止提交并在报告中记录 `ocr_disabled`。
- 转换页的 PDF 工作区直接读取文件 `content_summary`、当前 OCR 设置、PDF 来源公式报告和 Mathpix artifacts，提供选择 PDF、跳转 OCR 设置和创建 `pdf_to_word` 导出任务的前端入口。
- 每个 Mathpix PDF 作业会生成 `recognition_plan`、`retention_plan` 和脱敏 `request_options`，记录外部上传授权、凭证环境变量名、是否允许提交、图片对象、表格线索、公式线索、OCR 开关、实际 Mathpix 输出格式和图片/表格/公式保留状态。该计划在未授权外部上传或 OCR 全部关闭时也会进入报告，但不会触发实际上传。
- `/api/mathpix-jobs` 从报告中的 Mathpix 作业记录派生队列，展示授权、凭证、OCR 关闭、提交、完成、识别计划和输出摘要状态；接口不调用 Mathpix，也不返回本地输出路径。
- Mathpix 预检会在外部上传授权后检查 APP ID / APP KEY 环境变量是否存在，报告只展示环境变量名并隐藏实际凭证值。
- Mathpix 完成后的 DOCX 和公式结果包会转为标准 `artifacts` 项，复用 `/api/artifacts/{task_id}/{file}` 下载和任务结果包。
- `TaskProcessor` 会从 Mathpix `tex.zip` 中提取 `.tex` 公式，覆盖行内公式、展示公式、equation/align/gather/multline 环境和纯 TeX 公式行，作为 `source_type=PDF` 的公式项进入报告、公式页和公式 ZIP 导出；PDF 公式后处理通过 `pdf_formula_mathtype` 本地动作进入 `local-payload` 和 `desktop_execution_plan`，由本地客户端按 Windows/macOS MathType 合同写回对象或保留 MathML/LaTeX/图片兜底。
- `/api/enhancement-plan` 汇总 V3.0 AI 排版修复、智能模板套用、私有化部署、开放 API 和质量反馈闭环规划，前端只展示输入、授权和风控边界，真实 AI 或私有化引擎后续通过本地/私有化服务接入。
- 公式预检与 MathType 格式化任务会在每个公式项写入 `format_comparison`，包含格式化前、格式化后、变化项和摘要，用于报告导出与后续本地客户端复核。任务级 `formulaFormatScope` / `formatScope` 可覆盖设置范围，支持全文、当前章节、选中区域和单个公式。LaTeX 来源公式会写入 MathML 兜底和 MathType 预览字段，验收矩阵会用公式项自检证明这些链路。
- MathType 格式化任务会写入 `format_failure_policy` 和 `preserve_original_formula`。格式化关闭、已有 MathType 原样保留或格式化失败时，报告会明确记录兜底格式、原公式引用和是否可重试；验收矩阵会用公式项和对象保留清单自检已有 MathType 公式的原样保留策略。
- 公式任务会复用微小图片提取结果；疑似公式图片没有识别出可用 LaTeX 时，会生成 `source_type=图片公式` 的待确认公式项，保留原图资源引用、失败说明和 `fallback_action=保留原图`。
- `k12/previews.py`：按文件类型生成结构化预览，供工作台展示页面、对象和告警摘要；搜索、告警/对象点击定位、命中高亮、翻页和缩放在前端预览面板完成。
- PDF 微小图片扫描会解析图片 XObject 的宽高、Filter 和可导出流，优先使用真实对象元数据生成图片报告，缺少对象尺寸时才退回为 PDF 图片对象占位记录。
- `/api/reports/{report_id}/comparison` 复用结构化预览生成转换前后对比，比较源文件和成功转换 artifact 的页段、对象数量和告警。
- `k12/reports.py`：生成 JSON、HTML、PDF、XLSX 与 TXT 报告。
- `k12/converters.py`：使用标准库提取 docx/pptx 文本、PPT 备注和图片/图表/表格对象摘要，并生成最小可用的 pptx/docx 转换产物；Word 转 PPT 会额外统计源 DOCX 图片、表格、OMML 和 MathType/嵌入对象，写入 PPTX 对象保留清单页，验收矩阵会用临时 DOCX 自检该清单。
- PPT 解析会统计 OMML、MathType、Equation Native 和 LaTeX 线索，PPT 转 Word 产物、预览对象列表和公式报告都会显示公式数量；验收矩阵会用临时 PPTX 自检同一套公式线索识别逻辑，但复杂 MathType 对象复核仍交给本地客户端。转换页 PPT 工作区复用结构化预览展示缩略结构和内容摘要，并把模式、目录、备注提取、图片提取和公式提取写入任务 `ppt` options；PPT 转 Word DOCX 会写入目录段落和当前 Word 模板名，模板名、模板路径和目录开关会进入转换设置快照。
- `k12/converters.py` 也负责提取 xlsx 工作表、单元格、公式、计算结果、图表、图片和批注摘要，并生成标准 PDF、Word 和 PPT 产物；`TaskProcessor` 会在生成前按 Excel 转换范围、公式保留模式、拆分工作表开关或任务 options 过滤 sheet 与选区单元格，并在拆分开启时为每个 sheet 生成独立 artifact。
- OMML 依赖检索位于 `TaskProcessor`，在 Word 公式预检与 OMML 转换类任务中复用。
- `analysis.ommlConversionPrompts` 记录 Word 自带公式转 MathType 的用户确认合同，包含 OMML 数量、转换决策、决策来源、原公式保留策略和依赖状态；本地任务载荷中的 `omml_mathtype` action 会携带同一组 prompts，未确认时状态为 `awaiting_confirmation`，人工校正标记 `retry_conversion` 后状态为 `retry_queued`，桌面执行计划追加 `omml.retry_conversion` 步骤。
- 设置中的 `manualOmmlPath` 可手动指定 OMML 依赖文件，优先于自动目录扫描，适合自动检索失败后的人工修复。
- Word 宏处理位于 `TaskProcessor`，负责宏检测、宏选择队列、执行顺序、备份和本地客户端待执行状态；验收矩阵会用临时 `.docm` / `.dotm` 文件自检 Word 文档和模板宏启用容器检测，并用临时 `.docm` 文件自检宏顺序进入本地动作队列、`macro.run_ordered` 桌面计划、dry-run 同步和多文件复用，真实 VBA 模块枚举和执行仍由本地客户端承担。
- 宏安全设置支持关闭宏检测、关闭执行队列、按来源授权、白名单模式、执行超时和批量执行控制；被禁用或未授权的宏只进入报告，不会创建备份或进入本地执行队列。
- 宏库接口会从本地报告记录推导 `usage_count`、`last_used_at` 和 `recently_used`，前端可按名称、来源、用途和最近使用筛选宏。
- 宏模板 API 由 `/api/macro-templates` 提供，保存的模板会记录宏 ID、顺序、失败策略、执行时机和确认状态。
- 宏队列会把失败策略展开为 `failure_policy`，本地客户端可据此停止队列、跳过当前宏、继续后续宏或暂停等待用户决策。
- 前端宏执行报告卡片可用原报告中的宏 ID、文件 ID、失败策略和执行时机重新创建单个宏的 `macro_sequence` 本地任务，用于重新执行指定宏。
- 宏页可将最新宏报告中的宏顺序、文件集合、失败策略和执行时机载入当前执行顺序区，用户可通过拖拽手柄或上/下移图标调整顺序后通过原任务创建接口重新执行。
- `/api/reports/{report_id}/macro-failures.csv` 会动态导出失败、未授权和已禁用宏，默认脱敏备份路径，用于宏执行校正。
- 宏备份恢复由 `/api/tasks/{task_id}/restore-backups` 提供，只恢复 `.k12-data/backups/` 下由任务报告记录的备份。
- 微小图片检索从 OOXML 媒体目录提取图片，保存到 `.k12-data/images/`，通过受限 `/api/assets/images/...` 资源接口展示。
- 图片检索设置控制最大宽高/面积、页眉页脚图片、水印、透明图、重复图和导出格式；图片 ZIP 清单会包含这些标记，并分列记录 area、image_format、suspected_type 和 confirmed_type。
- 图片人工校正记录保存在 SQLite `image_annotations` 表中，支持类型已确认、误判、删除待处理和替换待处理；前端图片详情面板复用报告中的 smallImages 和人工标注记录，展示图片预览、位置、尺寸、面积、类型判断和标注状态，并可用图片 `file_id` 和 `location` 跳回来源文件预览；图片包由 `/api/reports/{report_id}/images.zip` 动态生成，微小图片清单由 `/api/reports/{report_id}/images.xlsx` 动态生成。
- 公式人工校正记录保存在 SQLite `formula_annotations` 表中，由 `/api/formula-annotations` 提供确认、批量确认、跳过、重识别和 LaTeX 修正记录；公式页详情校正工作区复用报告中的公式项和人工校正记录，集中渲染原始公式截图引用、LaTeX 编辑、MathType 预览、MathML/OMML 结果、格式化参数和置信度。
- 公式识别结果由 `/api/reports/{report_id}/formulas.zip` 打包导出 JSON、TEX、MML 与逐条公式文件，也可通过 `/api/reports/{report_id}/formulas.xlsx` 导出表格清单。
- OMML 人工校正记录保存在 SQLite `omml_annotations` 表中，由 `/api/omml-annotations` 提供转换失败、保留 OMML、重新转换、已修复和手动指定依赖路径记录；前端 OMML 依赖卡片会读取这些记录并提供保留、重新转换、已修复、指定依赖文件和导出 OMML 转换失败清单 CSV 的快捷操作，`manual_omml_path` 默认按本地路径脱敏返回。
- 排版人工校正记录保存在 SQLite `layout_annotations` 表中，由 `/api/layout-annotations` 提供页码、标题层级、表格结构、图片位置、公式位置和公式编号校正记录；报告页排版校正表单会写入修复前后描述、建议和处理状态。公式识别项缺少位置时，报告生成器会写入异常位置并同步到失败清单、质量检查和公式导出清单。
- 公式导出由 `/api/reports/{report_id}/formulas.zip` 动态生成，会合并人工校正结果并输出 `.json`、`.tex`、`.mml` 和清单。
- 源文件下载由 `/api/files/{file_id}/download` 提供，只返回文件名和二进制内容，不把本地路径写入下载响应。
- 批量源文件下载由 `/api/files/download` 提供，支持 `ids` 或 `type` 查询参数，并输出带 `manifest.csv` 的 ZIP。
- 文件列表的下载、替换和删除按钮直接映射到源文件下载、`/api/files/{file_id}/replace` 和 `DELETE /api/files/{file_id}`；替换和删除受 `files.manage` 权限控制。
- 转换输出保存到 `.k12-data/outputs/{task_id}/`，通过受限 `/api/artifacts/{task_id}/{file}` 下载。
- `outputDirectory` 为空或相对路径时解析到 `.k12-data/` 下，绝对路径则写入用户指定目录；报告会记录实际产物路径，避免之后改设置影响旧下载。
- 转换 artifact 会写入 `conversion_settings` 设置快照，覆盖模板、模式、PDF 精度、Excel 范围、输出同名策略、Word 转 PPT 分页/目录/备注/美化参数、PPT 转 Word 目录、Word 模板、备注提取、图片提取和公式提取偏好，以及保留图片/表格/页眉页脚/脚注尾注/批注/修订标记偏好；Word 转 PPT artifact 还会写入 `object_preservation`，用于同步到 HTML/TXT/XLSX 报告和本地客户端复核。
- `autoOpenOutputDirectory` 开启后，任务完成时只写入 `output_directory_action` 和日志，由后续本地客户端执行打开目录动作；网页端不直接打开本地路径。
- 报告生成阶段会写入 `qualityChecks`，覆盖 PRD 的文件校验、结构数量、内容保留、转换前后预览一致性、公式置信度、OMML、宏、微小图片、转换输出和 Mathpix 状态检查；预览一致性复用结构化预览比较源文件与成功转换产物的页段数、对象数量、对象减少项和告警数量。
- 报告导出会包含公式明细表，记录文件、位置、来源、原始截图引用、LaTeX、MathML、MathType 预览、格式化状态、置信度和处理状态。
- 公式置信度阈值、输出格式、字体、字号、格式化范围、对齐、变量/函数样式、上下标比例、分式/根式/矩阵/希腊字母样式、行内基线、独立公式间距、编号方式和低置信度处理策略来自设置页，任务选项可覆盖格式化范围，并写入公式明细和质量检查。
- 报告生成阶段会固化 `failureRows` 并写入 `{report_id}-failures.csv`，通过 `/api/reports/{report_id}/failures.csv` 下载。
- `/api/reports/{report_id}` 支持 DELETE，删除报告记录、报告导出文件、报告图片缓存和关联人工标注，不删除源文件、任务记录或转换输出。
- 批量任务会写入 `analysis.batchResults`，为每个文件记录状态、失败原因、建议任务、可重试标记和跳过状态；同时写入 `analysis.batchPlan` 固化工作台传入的最大并发、失败策略、是否遇错继续、是否需要用户决策、跳过数量和分片计划，并进入 HTML/TXT/PDF/XLSX 报告。
- `/api/tasks/{task_id}/skip-file` 支持将批量任务中的失败文件标记为跳过，写入 `skipped_file_ids` 和任务日志，并重建报告，使失败清单、质量检查和最近报告保持一致。
- 任务存储会根据 `start_time` 与 `end_time` 派生 `duration_seconds` 和 `duration_label`，并在任务完成后保存成功/失败数量、失败清单数量、可重试数量和任务级失败摘要，任务中心用这些字段展示耗时与失败原因。
- 任务中心详情面板复用 `/api/tasks` 的脱敏任务对象和 `/api/reports` 的报告摘要，展示任务参数、关联文件、报告入口、失败摘要和本地动作状态。
- 批量任务详情会读取报告中的 `analysis.batchResults`，逐文件展示状态、进度、失败原因和建议。
- 任务级结果包由 `/api/tasks/{task_id}/download` 动态生成，包含任务 JSON、日志、报告和转换输出。
- 服务启动时会将遗留的 `处理中` 任务标记为 `已中断` 和 `recoverable=true`；`/api/tasks/recovery-summary` 返回中断、暂停、失败和可重试任务的脱敏恢复清单；`/api/tasks/{task_id}/resume` 可从 `已中断` 或 `已暂停` 状态重新运行。
- `enableTaskCompletionNotice` 开启时，任务完成后会写入 `completion_notice`，前端会在工作台显示待提示任务；`/api/tasks/{task_id}/completion-notice/ack` 会把提醒标记为 `已提示` 并写入任务日志。
- `/api/logs?task_id={id}` 可筛选单个任务日志；`/api/logs/download` 可按 `format=txt|log|csv|json` 导出全局或指定任务日志，默认格式来自 `logExportFormat` 设置。任务中心的日志面板会在全局日志和任务日志之间切换。
- 日志表包含 `category` 字段，上传、转换、公式、OMML、宏执行、图片和错误日志受设置页分类开关控制；系统控制日志默认保留。
- `/api/maintenance/cleanup` 按设置或请求中的保留天数清理旧任务、报告、日志、报告文件、输出、备份和图片缓存，保留文件库和上传源文件。
- 服务启动时会读取 `autoCleanTemp`、`cleanupIntervalDays` 和 `lastCleanupAt`，只在开启且超过周期后执行一次自动清理。
- `saveHistory=false` 时，任务完成后立即删除任务历史、报告、日志和运行缓存；同步响应仍返回本次任务状态，文件库不受影响。

## 本地优先原则

PRD 中的敏感能力默认由本地客户端完成：

- Word 宏执行。
- OMML 转 MathType。
- MathType 格式化。
- 本机 OMML 文件检索和复制。
- 本地 Office 自动化。
- PDF 转 Word 走 Mathpix OCR，需要用户授权外部上传并配置凭证。
- OMML 文件只复制副本到当前 Word 文档目录，不删除原始依赖文件。
- 网页端不直接执行 Word 宏；宏任务必须由本地客户端授权执行。
- 本地 API 默认限制在 `127.0.0.1`，启用令牌后上传、设置、资源下载和报告导出都走同一套令牌校验。
- `capabilities.localConnection` 暴露本地客户端启用状态、网页唤起授权、云端同步授权、任务状态同步授权和敏感文档本地优先策略。
- `/api/local-client/manifest` 为桌面端提供运行清单，包含本地 API 地址、`k12-local://` 唤起协议、载荷接口、状态同步接口、心跳接口、支持动作、待本地任务数量和安装画像；`/api/local-client/heartbeat` 只保存平台、版本、状态、白名单能力摘要和脱敏组件预检，不保存令牌或本地路径。
- `python3 -m k12.local_client` / `k12-local-client` 是标准库本地伴随 CLI，可独立读取 manifest、发送心跳、领取本地任务载荷并执行 dry-run 状态同步；心跳预检会检测 Office、MathType、LibreOffice 和 OMML 依赖是否可见，并仅回传组件状态和能力位。dry-run 会生成 `k12.localDryRunExecution.v1`，逐个动作记录 gate、所需能力、步骤数量、输出类型和等待/阻塞原因。`--native-plan` 会生成 `k12.localNativeExecutionRequest.v1` 原生执行请求合同，列出 Windows 同平台、能力门槛、操作步骤、输出类型和阻断原因，供后续 pywin32/Office COM 执行器接管。`--execute-file-actions` 会生成 `k12.localFileActionExecution.v1` 并执行 OMML 依赖复制等安全本地文件动作，目标必须是当前文档所在目录。当前 CLI 不执行 Office、MathType、OMML 写回或 Word 宏，也不会在摘要中输出本地路径或令牌。
- 设置页异常处理面板会把系统异常预检拆成磁盘空间、文件权限、用户权限、本地客户端、客户端平台心跳、Office/MathType 组件、OCR 和宏授权明细，便于对应 PRD 12.7 的修复动作。任务预检会读取最近心跳中的脱敏组件能力，已连接客户端缺少当前任务必需能力时进入质量检查和失败清单。
- `/api/tasks/{task_id}/local-launch` 在网页端点击启动本地客户端时生成可审计的启动请求，写回任务的 `local_launch_request`，返回不含令牌的 `k12-local://` 协议 URL；前端跳转前再把当前受令牌保护的载荷地址放入协议参数。
- `/api/tasks/{task_id}/local-payload` 为本地客户端提供任务参数包，包含任务选项、输入文件本地路径、预检项、本地动作队列、桌面执行计划、客户端就绪判断、输出目录、公式交付合同和云端同步策略；`/api/tasks/{task_id}/local-readiness` 为网页端提供脱敏就绪状态，只返回任务摘要、握手状态、动作能力要求、能力缺口和桌面执行计划摘要，不返回输入路径、输出目录、组件路径或令牌。客户端就绪判断读取最近心跳的脱敏预检结果，按当前动作核对 Office 自动化、MathType 自动化、宏执行、OMML 依赖检索能力以及安装目标平台；当设置目标平台与心跳平台都是 Windows/macOS 且不一致时，状态为 `platform_mismatch`，避免平台专属 MathType/Office 对象跨系统交付。公式交付合同会把 Windows/macOS 平台、MathType 对象格式、跨平台不通用风险、输出优先级、兜底格式和是否允许平台专属 MathType 对象写入交给桌面端。桌面执行计划使用 `k12.desktopExecutionPlan.v1`，列出每个本地动作的能力门槛、授权门槛、平台门槛、步骤、输出合同和网页端不执行原生文档动作的边界。`/api/tasks/{task_id}/local-sync` 接收本地客户端回传的任务状态、进度、输出摘要、动作级 dry-run 校验摘要和结果上传意图。这些接口涉及本地任务交接，敏感载荷必须先配置本地安全令牌，状态同步还需要开启任务状态同步授权。
- `/api/local-client/uploads` 的 GET 从本地同步记录派生结果上传队列，展示本地客户端是否请求上传、云端同步是否授权、输出数量和脱敏输出摘要；POST 需要本地安全令牌和云端同步授权，可只登记上传包清单、输出哈希和大小，也可接收 `files` 中的 Base64 结果内容并写入受控 `cloud_uploads/{upload_id}/` 接收区，把状态更新为“已登记待云端接收”。`/api/local-client/uploads/{upload_id}/manifest` 返回云端接收清单，包含 upload_id、任务摘要、包哈希、输出名、大小、输出哈希、授权状态、内容传输模式和接收合同，设置页可从上传队列打开该清单。队列、登记响应和接收清单不暴露本地输出路径，也不回显文件内容。
- `requireLogin=true` 时，本地 API 会检查当前用户是否已通过 `/api/session` 登录且允许登录；未登录时仅放行健康检查、用户列表和登录接口。
- 文件、报告、任务控制和管理写接口会在处理器层检查角色权限：文件上传、替换、密码登记和删除需要 `files.manage`，报告删除需要 `reports.manage`，任务暂停、继续、取消、重试、失败文件跳过和备份恢复需要 `tasks.control`，用户管理需要 `users.manage`，通用模板和宏顺序模板需要 `templates.manage`，授权切换需要 `authorizations.manage`，设置保存需要 `settings.manage`；访客默认为只读，前端会读取当前用户有效权限并禁用无权限的文件密码、报告删除和任务控制按钮，`/api/api-catalog` 会暴露对应 `permission` 字段。
- API 列表和详情响应默认脱敏本地路径，下载、恢复和打包仍由后端使用原始 store 数据；`exposeLocalPaths` 仅在用户明确开启时返回真实路径。
- Windows 和 macOS 本地客户端使用不同安装画像；MathType 对象跨平台不通用，预检会对 OMML/MathType 相关任务提示平台专属对象风险，并支持 MathML/LaTeX/图片兜底策略。
- `/api/install-profile` 可按设置或查询参数返回目标平台画像，用于安装前预检和前端实时预览。
- `/api/install-plan` 返回平台安装计划，Windows 目标为 `.msi`，macOS 目标为 `.pkg`，并包含安装包状态、SHA256、步骤、令牌建议、客户端心跳平台匹配状态、`formula_compatibility` 公式交付合同和 MathType 跨平台兜底提示；公式交付合同固定声明 Windows/macOS MathType 原生对象不跨平台兼容、同平台原生对象要求和 MathML/LaTeX/图片兜底格式。只有 `.k12-data/installers/` 中存在真实安装包时，`/api/installers/{file_name}` 才会开放受控下载。
- `/api/architecture` 返回 PRD 第 16 章技术架构蓝图，包含当前 Python 本地 API、静态前端、SQLite/文件目录运行时，桌面端/网页端/本地服务/云端服务/任务队列/文档引擎/OCR/公式宏各层状态，以及本地任务载荷、状态同步、安装画像、能力探测和 Mathpix PDF 的接口合同。
- `/api/api-catalog` 返回 V3.0 开放 API 目录，列出健康检查、文件、任务、报告、日志、设置、能力、安装画像、技术架构、本地任务载荷和本地状态同步等接口的方法、路径、鉴权要求和敏感标记；本地任务载荷固定标记为必须配置安全令牌。
- `/api/acceptance-matrix` 返回 PRD 第 17 章逐条验收矩阵，按 10 个验收组和 96 个验收项标记已覆盖、合同覆盖、需本地/Mathpix 实测或缺口，并给出证据、当前状态和下一步；Word 转 PPT 与 PPT 转 Word 会运行临时 OOXML 产物自检来证明最小转换能力，Word 对象保留和已有 MathType 原样保留会运行临时 DOCX/公式项自检，公式批量/范围格式化和 LaTeX 转 MathType 预览会运行公式项自检，PPT 公式识别会运行临时 PPTX 公式线索自检，OMML 依赖检索/复制会运行临时目录自检，宏顺序执行交接和宏批量顺序会运行临时 `.docm` 队列自检，PDF 转 Word、扫描件 OCR 和公式 OCR 会运行不上传文件的 Mathpix 请求/输出合同自检，PDF 图片/表格保留会运行临时 PDF 图片 XObject 与表格线索自检，PDF 公式 MathType 后处理会运行临时 Mathpix `tex.zip` 自检，均不依赖历史任务样本。
- `/api/product-summary` 返回 PRD 第 21 章产品总结，包含 13 条核心能力和四阶段产品路线，并把真实桌面执行、Mathpix 外部识别和私有化/AI 增强等能力标记为合同覆盖或规划中，避免误报为已完成。

网页端和本地 API 只传递任务摘要、参数、状态和报告，不默认保存或上传完整正文内容。

## 后续演进

1. 将 `processor.py` 的接口替换为 FastAPI + 后台任务队列。
2. 增加 Electron 壳，把当前静态工作台打包为桌面端。
3. 在 Windows 本地服务中接入 pywin32、Office COM、MathType 自动化，并为 macOS 保持受限自动化与兜底格式链路。
4. 将 SQLite 数据模型迁移或同步到云端 Django/FastAPI 管理服务。
