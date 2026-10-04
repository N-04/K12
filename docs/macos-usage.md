# macOS 使用说明

安装 Python 3.11 或更高版本后，双击项目根目录 `start-macos.command`。启动器优先使用 `.venv/bin/python`，然后检查 Homebrew 和系统路径中的 Python，启动本地服务并打开浏览器。

也可从终端运行：

```bash
./start-macos.command
```

指定端口：

```bash
./start-macos.command --port 8766
```

关闭服务终端中的进程即可停止服务。文档与任务数据默认保存在项目 `.k12-data/`。

## 版本能力

macOS 与 Windows 共用工作台、文档上传、任务管理、报告与轻量 OOXML 转换。Windows 启动入口为 `start-windows.cmd`。

macOS 原生旧 Word/PPT/Excel 转换需要安装对应的 Microsoft Word、PowerPoint 或 Excel，并按 README 配置本地安全令牌、客户端心跳及任务授权。执行时同时指定 `--allow-native-execution --execute-native-office`。macOS 已接入旧 XLS 规范化及 PDF/Word/PPT 转换，脚本已通过本机 Excel 字典编译，实际打开保存仍需 Office 授权与实机验收。旧 XLS 的拆分工作表输出、MathType/OMML 写回和 Word 宏仍需后续适配；Windows COM 转换仍需 Windows 实机验收。

当前启动器运行源码版本，需要 Python；它不是独立 `.app` 或安装包。完整安装包交付仍待完成。
