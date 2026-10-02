# Windows 使用说明

需要 Python 3.11 或更高版本。基础网页功能仅使用 Python 标准库，无需 Make、Bash 或 pywin32。

## 启动

双击项目根目录的 `start-windows.cmd`。它先检查 `.venv`，再检查 Python 启动器和 PATH 中的 Python，启动后打开默认浏览器。

在 PowerShell 中也可直接运行：

```powershell
.\.venv\Scripts\python.exe -m k12 --open-browser
```

工作台地址为 http://127.0.0.1:8765。关闭服务终端或按 Ctrl+C 停止服务。运行数据默认保存在项目的 `.k12-data`，可用 `--data-dir` 指定其他目录。

PyCharm 的项目解释器请选择 `.venv\Scripts\python.exe`，运行配置选择模块 `k12`，工作目录选择项目根目录。若解释器安装中断，应先安装或修复 Python，而不是使用不可用的虚拟环境。

## 功能范围

基础功能包括上传、ZIP 批量导入、文件分析、结构预览、任务管理、DOCX/PPTX/XLSX 的轻量转换、文本层 PDF 转 DOCX，以及报告和任务结果下载。

这些轻量转换以文档结构和摘要为主，不能保证 Office 原始版式完整复现。扫描 PDF 和复杂 PDF 的 Mathpix OCR 仍需要配置凭据与明确的外部上传授权。

旧 `.doc/.dot` 的 Word 转 PPT，以及旧 `.ppt` 的 PPT 转 Word，可通过本机桌面版 Word/PowerPoint 进行 OOXML 规范化，再由 K12 生成目标文档。需要 Windows PowerShell 和对应的 Office COM 注册；转换过程中禁止执行文档宏。

MathType 原生对象写回、OMML 转 MathType、Word VBA 宏实际执行仍需要独立适配器；现有检测、计划和模拟执行不能证明这些功能已完成。旧 `.xls` 的业务转换流程仍未接入原生执行。

## 执行旧 Word/PPT 任务

1. 在网页设置中保存本地安全令牌，平台选择 Windows。
2. 在 PowerShell 中设置 `K12_LOCAL_TOKEN`，运行客户端发送心跳。
3. 上传旧 Word/PPT 文件并创建对应转换任务。
4. 明确授权执行该任务：

```powershell
$env:K12_LOCAL_TOKEN = Read-Host '输入网页设置中的本地安全令牌'
.\.venv\Scripts\python.exe -m k12.local_client
.\.venv\Scripts\python.exe -m k12.local_client --task-id 任务ID --allow-native-execution --execute-native-office
Remove-Item Env:K12_LOCAL_TOKEN
```

任务输入使用受管 `.inputs` 快照，执行前复核大小和 SHA-256。最终产物检查 OOXML 结构和 ZIP CRC，再回传真实执行报告、大小与哈希；服务端重新校验后才允许下载。客户端只处理 `pending_files` 指定的待执行文件；同批已完成的 DOCX/PPTX 产物保持不变。待执行批次内任何文件失败只回滚本轮新建产物。空清单直接跳过执行和回传。

## 测试

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests
```

Windows 创建符号链接需要开发者模式或相应权限。仅涉及真实符号链接的安全测试在缺少权限时明确跳过，其他边界测试仍执行。

安装桌面 Word 和 PowerPoint 后，可单独运行真实集成测试。该测试生成自己的无宏文档、启动独立本地服务、校验转换与下载，结束时清理测试数据：

```powershell
$env:PYTHONPATH = (Get-Location).Path
.\.venv\Scripts\python.exe tests\windows_smoke.py
```

Office 首次启动、激活、恢复或加载项弹窗可能阻止自动化。首次使用前请手动打开相应应用，完成激活并关闭弹窗后再测试。为保护未保存文档，转换前必须关闭对应 Office 应用。COM 转换有超时限制。客户端登记本次实例的应用类型、进程 ID 和创建时间，先尝试关闭文档并退出，必要时仅终止身份核对通过的实例。若初始化阶段无法确认归属，将明确提示并保留未知进程；请人工检查并关闭 Office 后重试。不会按进程名称批量结束 Office。

COM 参数参考：[Word SaveAs2](https://learn.microsoft.com/en-us/office/vba/api/word.saveas2)、[PowerPoint SaveAs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas)、[Office AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/office.msoautomationsecurity)。
