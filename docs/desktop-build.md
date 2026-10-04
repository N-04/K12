# macOS 与 Windows 桌面构建

在对应系统安装 Python 3.11 或更新版本后运行：

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r tools/requirements-desktop.txt
.venv/bin/python tools/build_desktop.py
```

Windows 使用 `.venv\Scripts\python.exe` 替代 `.venv/bin/python`，创建虚拟环境可使用 `py -3 -m venv .venv`。

macOS 输出 `dist/K12.app`；Windows 输出 `dist/K12/K12.exe`，发布时需要保留整个 `dist/K12` 文件夹。构建包含 Python 运行时、静态页面和 Office PowerShell 脚本，用户运行时无需另外安装 Python。

双击程序会启动本地 HTTP 工作台并打开默认浏览器。此版本使用浏览器作为界面；原生 Office 转换仍需要用户安装相应 Office 并显式授权。

数据目录：macOS 为 `~/Library/Application Support/K12`；Windows 为 `%LOCALAPPDATA%\K12`。支持 `--data-dir` 指定其他目录，源码版本的数据不会自动迁移。

本地伴随客户端也包含在桌面程序中：

```bash
K12 --local-client --help
```

构建需在目标系统执行。macOS APP 已完成本机构建、启动与转换下载验证；Windows EXE、正式签名以及 MSI/PKG 安装包仍需对应系统验证。

构建方式依据 [PyInstaller 打包文档](https://pyinstaller.org/en/stable/spec-files.html)；资源路径依据 [运行时文档](https://pyinstaller.org/en/stable/runtime-information.html)。

## 本机验证（2026-10-03）

已在 macOS 15.8 arm64、Python 3.11.9 下生成 `dist/K12.app`，实际验证帮助入口、HTTP 首页、设置 API 和临时独立数据目录。完整回归 299 项通过，1 项 Windows 实机测试跳过。

若打包签名提示资源扩展属性，可对新生成的应用包执行：

```bash
xattr -cr dist/K12.app
codesign --force --deep --sign - dist/K12.app
```

本机包使用临时签名，尚未完成 Apple 开发者签名和公证。Windows 构建、MSI/PKG 安装及分发验收仍待完成。

## 可重复验证与双平台工作流

构建后执行：

```bash
python tools/smoke_desktop.py dist/K12.app/Contents/MacOS/K12
```

Windows 改用 `dist/K12/K12.exe`。验证器实际检查客户端帮助入口、首页、JS/CSS、设置 API 和独立 SQLite 目录；使用临时目录与空闲端口，结束后关闭本轮进程。桌面程序新增 `--no-open-browser`，用于后台运行及验证。

`.github/workflows/desktop.yml` 在 macOS、Windows runner 分别运行全量测试、打包及验证，成功后上传对应压缩包。macOS 使用 `ditto` 保留应用权限，Windows 保留完整运行目录。配置已加入本地仓库，尚未上传或执行 GitHub 工作流，不能作为 Windows 构建成功证据。

本机最新结果：300 项测试通过，1 项 Windows 实机测试跳过；重新构建的 macOS APP 通过上述验证器。本机扩展属性清理和临时签名已加入构建脚本。

工作流参考 GitHub 官方 [checkout](https://github.com/actions/checkout)、[setup-python](https://github.com/actions/setup-python) 和 [upload-artifact](https://github.com/actions/upload-artifact) 配置。

## 最新包内转换验证

已重新构建 macOS APP，纳入最新 Word 样式、PPT 布局、页面顺序和桌面启动异常处理。验证器通过包内 API 注册临时 DOCX、创建 Word 转 PPT 任务、读取报告、下载 PPTX，确认正文保留且 SHA-256 与报告一致。

交付文件：`dist/K12-macOS.zip`（macOS arm64，临时签名）。SHA-256：`0f6406915643fc52fd453497990ab93edb29be0332b051bc0ee8b2c0f749471b`。工作流同样执行这项包内转换检查。Windows EXE 与正式签名/公证仍未验证。

最新重建纳入中文 PDF、PPT/Excel 原生 Word 表格、表格保留开关和 Excel 任务设置快照修复。316 项回归通过（1 项 Windows 实机测试跳过），实际应用包再次通过资源、API、转换及下载校验。

最新包进一步纳入 Excel 合并区域、矩形选区和转换参数界面；320 项测试运行成功（319 项通过、1 项 Windows 实机跳过），macOS 包内验证通过。

最新包已纳入图表缓存及源数据恢复、原生柱状/条形/折线/饼图和内嵌工作簿。桌面验证器实际调用 Excel 转 PPT API，检查下载哈希、原生折线图及内嵌 XLSX；本机 macOS 包通过。325 项源码回归运行成功（324 项通过，1 项 Windows 实机跳过）。

最新包纳入传统/线程批注及去重、图表提示、macOS 旧 Excel 适配器和 XLSX 结构校验。包内实际 Excel 原生图表与批注正文、Word 转 PPT、API、资源及下载哈希验证通过。最新完整源码回归为 336 项（335 项通过，1 项 Windows 实机跳过）。 macOS 旧 XLS 实际打开保存及 Windows 实机仍未验证。


本轮交付包纳入第四十一至四十四轮修复：PDF 图表数据提示、Word 原生 PPT 表格、横向/纵向合并和分页避让。验证器增加合并起点、覆盖单元格以及表格后正文检查。工作目录文件提供器会重新附加 Finder 元数据，因此在 /private/tmp 临时目录清理并重新临时签名后压缩；对实际 ZIP 解压的应用执行严格签名校验和包内 API 转换验证。最新源码完整回归 340 项（339 项通过、1 项 Windows 实机跳过）。压缩包仍为 macOS arm64 临时签名，未经开发者签名或公证；Windows 构建和旧 Office 实机验收尚未完成。


构建完成后运行 `python tools/package_desktop.py` 生成交付 ZIP。工具在临时目录处理 macOS 元数据和临时签名，或压缩 Windows 完整运行目录；解压 ZIP 后运行包内验证，通过才覆盖已有交付包，并打印 SHA-256。双平台工作流已使用该命令。工具在本机 macOS 实际通过；Windows 分支尚未实机运行。本轮重新压缩的是第四十五轮构建的应用，包含至第四十四轮转换修复；第四十六轮 PPT 转 Word 合并表格源码仍未进入该应用。


第四十八轮已重新构建最新源码，第四十六轮 PPT 转 Word 合并结构进入交付包。实际 ZIP 解压后的验证增加 Word → PPT → Word 合并表格往返：确认横向跨度、纵向起点/续接、唯一正文及反向下载哈希。macOS 签名、包内资源、API、图表和批注检查均通过，当前压缩包哈希见上方交付文件记录。Windows 实机验证与正式签名、公证仍待完成。


第五十三轮交付已纳入列宽、显式/混合自动行高及合并格高度分摊。包内验证增加 1:2 列宽与行高比例检查，实际 ZIP 解压运行通过。构建工具不再原地重签名；macOS 签名和启动验证统一由 `package_desktop.py` 在临时目录完成，工作流移除签名前的重复启动验证。普通构建输出不作为最终可分发包，请使用经过交付工具验证的 ZIP。


第五十七轮交付重新构建至第五十六轮源码，纳入首表后的正文归类、单元格空段落/空格与双向显式对齐。实际 ZIP 解压验证新增居中、空段落往返检查，签名、运行、双向表格、行列比例及下载验证全部通过。共享源码最新全量回归为 349 项（348 项通过、1 项 Windows 实机跳过）。Windows 实机及完整视觉验收仍未完成。


第六十二轮交付重建至第六十一轮源码，纳入对象验收误报修复及双向实际图片媒体、裁剪/旋转/翻转。验证器新增包内 Word → PPT → Word 图片字节、原生形状/绘图及引用检查；实际 ZIP 解压后的签名、运行、图片及已有表格/图表/批注检查通过。完整图片布局和 Windows 实机仍待验收，不将此包内样本验证作为全 PRD 完成证据。


第六十六轮交付已重建至第六十五轮源码，包含 VML/光栅尺寸回退、页眉页脚图片提取和图片布局限制报告。实际 ZIP 解压验证增加双向任务产物布局提示检查，签名、资源、API、图片往返及既有表格/图表/批注检查通过。页眉页脚与 VML 的独立验证证据仍为源码专项，未将包内正文 PNG 样本检查扩大为全部格式实机验收。最近完整回归 358 项（357 项通过、1 项 Windows 实机跳过）。
