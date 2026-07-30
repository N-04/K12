# K12 文档智能处理工具

最后维护时间：2026-07-03 CST

K12 是一个本地优先的文档处理工作台，用于集中处理 Word、Excel、PPT、PDF 等教学与办公文档。项目提供文件上传、格式识别、任务队列、转换报告、公式与图片检查、本地客户端交接和网页管理能力。

当前版本使用 Python 标准库后端和静态前端，默认在本机运行，不依赖外网服务。涉及 Office、MathType、OMML、Word 宏、PDF OCR 等强本机或第三方能力时，系统会先记录能力位、授权状态和处理计划，再交给本地客户端或外部服务执行。

## 核心能力

- 本地文件上传、分类、预览和任务管理。
- ZIP 批量导入包含路径穿越清理、展开大小、条目数量、异常压缩比和加密成员校验。
- ZIP 重新上传会保留父文件 ID，同时清理旧子条目并重建新压缩包内容。
- Word、PPT、Excel、PDF 转换流程与结构化报告。
- 公式、OMML、MathType、宏、微小图片等对象的检测与处理记录；Word 宏按 OOXML 的实际 VBA 项目部件检测，不仅依赖扩展名。
- 批量任务、失败清单、日志、报告导出和任务结果打包。
- 本地客户端运行清单、心跳、任务载荷、dry-run、脱敏原生执行报告和同步接口。
- 权限、隐私、安全令牌、外部上传授权和本地路径脱敏控制。

## 能力边界

- 默认不上传敏感文件到外部服务。
- 加密 Office/PDF 会在分析阶段提示密码；密码仅保存在当前本地服务进程，不写入数据库或报告。
- 未通过文件校验的内容不会进入转换、对象提取或 Mathpix 外部上传流程。
- 不兼容输入、转换异常、Mathpix 阻断或图片提取失败会将任务标记为失败；报告按文件去重统计成功和失败数。
- 本地转换和 Mathpix 下载产物在登记成功前会校验 OOXML ZIP 结构、PDF 头尾或公式 ZIP 安全边界并登记大小、SHA-256；每次下载会复算内容，登记后即使被替换成另一份格式有效文件也不会放行；重试任务只从最新报告开放输出。
- 产物下载和任务结果包会再次校验成功状态、任务输出目录边界与文件完整性；报告下载只允许访问受管报告目录。
- 上传与 ZIP 解压条目的下载路径必须位于受管上传缓存目录；用户显式登记的外部本地文件仍按原路径合同处理。
- 旧版 `.doc/.dot/.xls/.ppt` 上传会验证 OLE/CFB 签名、字节序和扇区头；损坏容器会阻止任务，加密容器会提示密码，深度解析与转换明确交给本地 Office 客户端。
- 通过预检的旧版 Office 转换不会被 OOXML ZIP 转换器误报失败；任务以 95% “待处理”进入本地队列，报告单独统计待本地文件，直到客户端回传真实结果。
- 旧格式任务只在 `local-sync` 同时提供原生 `office_conversion` 成功证据和受管、可打开的输出文件后升级为成功；空状态、损坏输出或任务目录外路径不能消除 pending。
- 本地 Office 失败回传会把 pending artifact 终结为 `local_execution_failed`，并同步文件失败计数、质量检查、失败 CSV 和任务错误信息。
- 本地 Office 成功回传还会校验输出文件修改时间不早于当前任务轮次，防止重试时把仍留在任务目录中的旧文件冒充新结果。
- 待本地 Office 输出必须回传合法文件大小和 SHA-256，并与任务目录内文件的实际字节数、内容哈希一致；缺失、格式错误或内容不匹配都会阻止成功登记。
- JSON、HTML、PDF、XLSX、TXT 和失败 CSV 报告生成后统一登记大小、SHA-256；报告下载与任务 ZIP 会使用复验后的字节快照，受管目录内文件被替换也不会放行。
- 上传缓存、ZIP 解压条目和注册时存在的外部源文件会登记 `source_sha256`；单个源文件下载与批量 ZIP 都复验大小、哈希和读取快照，分析后被替换的源文件不会继续下载。
- 每次任务运行在分析、OCR 或转换前重新核对源文件登记大小与 SHA-256；登记后被替换的文件会进入校验失败、失败报告和失败清单，不生成派生产物。
- 复验通过的输入会复制为当前任务 `.inputs` 快照，服务端处理和本地客户端载荷都使用该快照及其 `source_sha256`；原文件在本轮处理中变化不会改变转换输入。
- 用户取消或本地客户端回传取消会把待处理 artifact 终结为 `local_execution_cancelled`，单独统计取消文件并重写报告；取消既不冒充成功，也不污染失败清单。
- Mathpix PDF OCR 需要显式授权并配置 `MATHPIX_APP_ID`、`MATHPIX_APP_KEY`。
- 纯文本层 PDF 会在本地解析普通或 Flate 压缩内容流中的 `Tj` / `TJ` 文本并生成 DOCX，支持 PDF 转义、UTF-8、带 BOM 的 UTF-16 十六进制字符串和连续 `TJ` 字形片段。本地解析限制源文件 50 MiB、单流展开 10 MiB、累计流展开 25 MiB 和 1000 个流，超限内容不会进入本地转换。扫描型、混合型、含图片、公式或表格的 PDF 仍走 Mathpix 识别合同。每个 Mathpix 作业会记录 `task_option_audit`（`k12.mathpixTaskOptionAudit.v1`），审计任务参数中的 `allowExternalMathpixUpload`、`externalUploadAuthorized` 等伪授权键，并声明外部上传授权只能来自 `settings.allowExternalMathpixUpload`。
- Office COM、MathType 原生对象写回、真实宏执行等能力依赖后续本地客户端或平台适配器。
- macOS 已接入 Word/PowerPoint AppleScript 任务适配器：只有同时传入 `--allow-native-execution --execute-native-office` 才会执行旧 Word/PPT 转换；输入使用任务快照，本地路径不插入脚本，最终产物写入任务受管目录并通过 OOXML、ZIP CRC、大小和 SHA-256 校验，再携带脱敏原生报告与输出摘要同步。
- macOS Office 输入解析真实路径后必须是任务输出目录下 `.inputs/` 的直接文件；外部路径、嵌套路径和指向目录外的符号链接即使哈希正确也不会启动 Office。
- macOS Office 批量执行采用原子提交：本轮任一文件失败会删除已生成的本轮最终产物，只有全部文件成功才向 `local-sync` 提交输出；回滚只允许删除任务输出根目录的直接文件。
- 当前机器的 Word、PowerPoint 和真实旧 PPT 二段转换证据记录在 `docs/native-verification.md`；文档同时保留真实旧 Word 输入尚未闭环的限制。
- 运行数据写入 `.k12-data/`，该目录用于本地 SQLite、上传缓存、报告和输出文件。

## 运行

```bash
make run
```

打开：

```text
http://127.0.0.1:8765
```

## 测试

```bash
make test
```

## 项目结构

```text
k12/              本地 API、数据模型、任务处理与报告生成
static/           无构建前端工作台
tests/            标准库 unittest 测试
docs/             架构与范围说明
.k12-data/        运行时 SQLite、报告与输出目录（不入库）
```

## 本地客户端

本地伴随 CLI 用于验证网页与本地能力交接：

本地客户端下载清单通过 `platform.installer` 暴露 `k12.localInstallerManifest.v1` 合同，按 Windows 和 macOS 区分安装包、可用状态、SHA256 和 `MathType 原生对象边界`。安装包下载入口为 `/api/installers/{file_name}?platform=...`，接口只返回受控文件名和校验信息，且不暴露 `installers/` 本地目录；本地伴随 CLI 运行结果会输出脱敏的 `k12.localClientManifestSummary.v1`，保留安装包平台、下载 platform 查询要求和 MathType 同平台边界。

```bash
K12_LOCAL_TOKEN="你的本地安全令牌" python3 -m k12.local_client --origin http://127.0.0.1:8765 --sync-dry-run
```

生成原生执行请求合同：

```bash
K12_LOCAL_TOKEN="你的本地安全令牌" python3 -m k12.local_client --origin http://127.0.0.1:8765 --native-plan --allow-native-execution
```

在 macOS 上显式执行受支持的旧 Word/PPT Office 转换并同步结果：

```bash
K12_LOCAL_TOKEN="你的本地安全令牌" python3 -m k12.local_client --origin http://127.0.0.1:8765 --allow-native-execution --execute-native-office
```

真实桌面客户端完成 Office、MathType、OMML 或 Word 宏动作后，可通过 `local-sync` 回传 `k12.localNativeExecutionReport.v1` 脱敏执行报告；验收矩阵只读取动作类型、平台、能力、步骤数量、输出类型和成功状态，不保存本地路径或令牌。已有外部原生执行器结果 JSON 时，可用本地伴随 CLI 负责脱敏和同步：

```bash
K12_LOCAL_TOKEN="你的本地安全令牌" python3 -m k12.local_client --origin http://127.0.0.1:8765 --native-report-json ./native-report.json
```

## 安全说明

服务默认监听 `127.0.0.1`。设置本地安全令牌后，API、上传、报告下载、图片资源和任务打包下载都会要求携带令牌。API 响应默认隐藏本地文件路径，只有在设置中允许显示本地路径时才返回真实路径。
