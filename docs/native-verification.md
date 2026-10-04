# 本地原生执行验证

验证日期：2026-07-15（Asia/Shanghai）

## macOS 组件

- Microsoft Word：已安装，AppleScript 自动化可调用。
- Microsoft PowerPoint：已安装，AppleScript 自动化可调用。
- MathType：未发现。
- LibreOffice：未发现。

## Word for Mac

使用受管 DOCX 副本调用 `normalize_macos_office_document`，真实启动 Word 并另存为临时 DOCX。首次实测发现 `open` 不返回可直接绑定的文档对象，改为读取 `active document` 后通过。

验证结果：

- 状态：成功。
- 输出大小：35,975 字节。
- SHA-256：`0834ca10d9200993808495cb46a45edd7b7715a4877ff07d77917b64ff3bce37`。
- OOXML 关键部件、ZIP CRC、大小与哈希校验：通过。
- Word for Mac 不接受测试脚本另存为 `format document97`，因此当前环境未制造出真实 `.doc` 测试夹具；旧 Word 完整链路仍保留此实测缺口。

## PowerPoint for Mac

实测确认 PowerPoint 的保存目标必须使用 HFS 文本路径，即 `(POSIX file path) as text`；直接传 POSIX 文本会只修改演示文稿名称而不按目标路径落盘，传文件对象会产生零字节占位文件。修正后通过真实保存验证。

规范化验证结果：

- 状态：成功。
- 输出大小：35,498 字节。
- SHA-256：`b88b333a3ccf265cbd989ad3f0f8c26f9e56fcf57d0ee71041dfdb42274d427e`。
- OOXML 关键部件、ZIP CRC、大小与哈希校验：通过。

真实旧 PPT 任务链验证结果：

- PowerPoint 生成真实旧版 `.ppt`：46,592 字节。
- 执行链：`.ppt` → PowerPoint 规范化 `.pptx` → K12 生成 `.docx`。
- `k12.macosOfficeTaskExecution.v1` 状态：`success`。
- 最终 DOCX：1,073 字节。
- 最终 SHA-256：`30b72696a18fbe60c88f335f91ffd80d97c500ae1fd7336ae01e26ae1cd6a476`。
- 中间 `.native-*` 文件：已清理。

## 结论边界

PowerPoint 旧格式转换与任务二段处理已有当前机器真实证据；Word 的 Office 自动化和 DOCX 规范化已有真实证据，但真实 `.doc` 输入的完整任务链仍未在当前环境验证。MathType、OMML 对象写回和 Word 宏执行仍不得据此升级为已覆盖。

## 自动化闭环回归

集成测试从真实服务端任务合同出发，覆盖旧 `.ppt` 上传分析、macOS 心跳与 `officeAutomation` 能力门禁、受管 `.inputs` 快照、大小与 SHA-256 校验、原生执行报告脱敏、`local-sync` 对账、报告和质量状态重写，以及 artifact 下载内容一致性。测试中的 Office 规范化边界使用可控 PPTX 夹具，用于验证跨模块合同，不替代上述真实 PowerPoint GUI 实测。

## 2026-10-04：Word for Mac 宏接口复核

直接读取当前安装的 /Applications/Microsoft Word.app/Contents/Resources/Word.sdef，确认命令 run VB macro（sWRD1149）及参数 macro name（5112）存在，说明原生应用提供已有宏执行接口。sdef CLI 因未安装完整 Xcode 不可用，使用应用自带字典完成只读核验。

该证据不证明宏实际执行成功。字典未发现 VBProject 或导入源代码入口，当前导入的 .bas 不能据此自动加载；客户端还需实现受管文档/模板的已有宏执行、备份、所有权、顺序和失败策略，并以真实产物验证。macroExecution 继续保持 False。
