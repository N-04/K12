const tokenStorageKey = "k12LocalSecurityToken";

const state = {
  files: [],
  tasks: [],
  reports: [],
  logs: [],
  taskLogs: [],
  settings: {},
  capabilities: null,
  installProfile: null,
  installPlan: null,
  localClientManifest: null,
  localUploadQueue: null,
  localUploadManifest: null,
  localReadiness: {},
  architecture: null,
  apiCatalog: null,
  dataDictionary: null,
  acceptanceMatrix: null,
  productSummary: null,
  enhancementPlan: null,
  macros: [],
  macroTemplates: [],
  users: [],
  currentUser: null,
  templates: [],
  authorizations: [],
  preflightChecks: [],
  mathpixJobs: null,
  imageAnnotations: [],
  formulaAnnotations: [],
  ommlAnnotations: [],
  layoutAnnotations: [],
  previewComparison: null,
  filePreview: null,
  selectedFiles: new Set(),
  macroSequence: [],
  plannerOrder: [],
  previewSearch: "",
  previewZoom: 1,
  previewPageOffset: 0,
  filter: "all",
  view: "workspace",
  activeFileId: "",
  activeTaskDetailId: "",
  activeReportId: "",
  activeLogTaskId: "",
  activeFormulaId: "",
  activeImageId: "",
  activeMacroId: "",
  localToken: readStoredToken(),
};

let plannerPointerDrag = null;

const taskSpecs = {
  word_to_ppt: {
    label: "Word 转 PPT",
    mode: "hybrid",
    inputs: ["Word"],
    output: "PPT",
    description: "标题层级、表格、图片、公式和备注进入课件生成流程",
  },
  ppt_to_word: {
    label: "PPT 转 Word",
    mode: "hybrid",
    inputs: ["PPT"],
    output: "Word",
    description: "抽取幻灯片标题、正文和备注，生成讲义或汇报文档",
  },
  pdf_to_word: {
    label: "PDF 转 Word",
    mode: "hybrid",
    inputs: ["PDF"],
    output: "Word",
    description: "Mathpix 识别 PDF，公式 OCR 后进入本地 MathType 后处理",
  },
  excel_to_pdf: {
    label: "Excel 转 PDF",
    mode: "web",
    inputs: ["Excel"],
    output: "PDF",
    description: "保留工作表、公式摘要、图表和分页输出",
  },
  excel_to_word: {
    label: "Excel 转 Word",
    mode: "web",
    inputs: ["Excel"],
    output: "Word",
    description: "提取工作表、单元格和公式摘要，生成 Word 表格文档",
  },
  excel_to_ppt: {
    label: "Excel 转 PPT",
    mode: "web",
    inputs: ["Excel"],
    output: "PPT",
    description: "按工作表生成图表页摘要，生成 PPT 演示文稿",
  },
  formula_precheck: {
    label: "Word 公式预检",
    mode: "local",
    inputs: ["Word"],
    output: "公式报告",
    description: "检测 OMML、MathType、LaTeX、图片公式和低置信度对象",
  },
  omml_to_mathtype: {
    label: "OMML 转 MathType",
    mode: "local",
    inputs: ["Word"],
    output: "Word",
    description: "检索 OMML 依赖，复制到文档目录后转换为 MathType",
  },
  mathtype_format: {
    label: "MathType 格式化",
    mode: "local",
    inputs: ["Word", "PPT"],
    output: "格式化报告",
    description: "统一公式对象格式，保留失败兜底和数量校验",
  },
  macro_sequence: {
    label: "Word 宏顺序执行",
    mode: "local",
    inputs: ["Word"],
    output: "宏执行报告",
    description: "按用户选择的顺序执行宏，支持备份和失败策略",
  },
  small_image_scan: {
    label: "微小图片检索",
    mode: "hybrid",
    inputs: ["Word", "Excel", "PPT", "PDF"],
    output: "图片报告",
    description: "定位公式截图、图标、二维码、印章和签名等小对象",
  },
  batch_process: {
    label: "批量处理",
    mode: "hybrid",
    inputs: ["Word", "Excel", "PPT", "PDF"],
    output: "批量报告",
    description: "多文件排队、并发调度、失败重试和结果打包",
  },
};

const taskLabels = Object.fromEntries(Object.entries(taskSpecs).map(([key, value]) => [key, value.label]));

const viewMeta = {
  workspace: ["工作台", "文件导入、任务分流、能力状态和报告入口"],
  menu: ["菜单结构", "按 PRD 菜单树快速进入文档处理功能"],
  word: ["转换", "Word、PPT、PDF、Excel 转换任务和 Word 转 PPT 参数"],
  formula: ["公式处理", "按来源、置信度和格式化状态查看公式对象"],
  macro: ["宏处理", "宏库选择、执行顺序、备份和风险确认"],
  images: ["微小图片", "公式截图、图标、二维码、印章、签名检索"],
  tasks: ["任务中心", "进度、日志、重试、取消和报告下载"],
  admin: ["网页管理", "用户、权限、模板、授权和运行预检"],
  settings: ["设置", "用户、模板、授权、预检、转换、公式、OMML、宏、OCR 和本地连接"],
};

const menuStructure = [
  {
    title: "首页",
    view: "workspace",
    items: [{ label: "工作台总览", view: "workspace" }],
  },
  {
    title: "文件处理",
    view: "workspace",
    items: [
      { label: "Word 转 PPT", task: "word_to_ppt", view: "word" },
      { label: "PPT 转 Word", task: "ppt_to_word", view: "word" },
      { label: "PDF 转 Word", task: "pdf_to_word", view: "word" },
      { label: "Excel 处理", task: "excel_to_pdf", view: "word" },
    ],
  },
  {
    title: "Word 处理",
    view: "word",
    items: [
      { label: "Word 公式预检", task: "formula_precheck", view: "word" },
      { label: "Word 自带公式检测", task: "formula_precheck", view: "word" },
      { label: "OMML 转 MathType", task: "omml_to_mathtype", view: "word" },
      { label: "OMML 文件检索", task: "omml_to_mathtype", view: "word" },
      { label: "Word 宏顺序执行", task: "macro_sequence", view: "macro" },
      { label: "Word 公式格式化", task: "mathtype_format", view: "formula" },
      { label: "Word 转 PPT", task: "word_to_ppt", view: "word" },
    ],
  },
  {
    title: "宏处理",
    view: "macro",
    items: [
      { label: "宏检测", task: "macro_sequence", view: "macro" },
      { label: "宏选择", view: "macro" },
      { label: "宏排序", view: "macro" },
      { label: "宏执行", task: "macro_sequence", view: "macro" },
      { label: "宏执行报告", view: "tasks" },
      { label: "宏顺序模板", view: "macro" },
    ],
  },
  {
    title: "公式处理",
    view: "formula",
    items: [
      { label: "MathType 识别", task: "formula_precheck", view: "formula" },
      { label: "Word 公式格式化", task: "mathtype_format", view: "formula" },
      { label: "PDF 公式识别", task: "pdf_to_word", view: "word" },
      { label: "图片公式识别", task: "formula_precheck", view: "formula" },
      { label: "公式人工校正", view: "formula" },
      { label: "公式报告", view: "tasks" },
    ],
  },
  {
    title: "图片检索",
    view: "images",
    items: [
      { label: "微小图片检索", task: "small_image_scan", view: "images" },
      { label: "图片提取", task: "small_image_scan", view: "images" },
      { label: "图片去重", view: "images" },
      { label: "图片报告", view: "tasks" },
    ],
  },
  {
    title: "批量任务",
    view: "tasks",
    items: [
      { label: "任务列表", view: "tasks" },
      { label: "任务进度", view: "tasks" },
      { label: "失败重试", view: "tasks" },
      { label: "结果下载", view: "tasks" },
    ],
  },
  {
    title: "本地客户端",
    view: "settings",
    anchor: "settings-local",
    items: [
      { label: "本地服务状态", view: "settings", anchor: "settings-local" },
      { label: "本地任务", view: "tasks" },
      { label: "本地文件处理", view: "workspace" },
      { label: "本地 Office 检测", view: "settings", anchor: "settings-local" },
      { label: "MathType 检测", view: "settings", anchor: "settings-mathtype" },
      { label: "OMML 依赖检测", view: "settings", anchor: "settings-omml" },
      { label: "宏权限检测", view: "settings", anchor: "settings-macro" },
    ],
  },
  {
    title: "网页管理",
    view: "admin",
    items: [
      { label: "用户管理", view: "admin" },
      { label: "权限管理", view: "admin" },
      { label: "模板管理", view: "admin" },
      { label: "任务管理", view: "tasks" },
      { label: "报告管理", view: "admin" },
      { label: "授权管理", view: "admin" },
    ],
  },
  {
    title: "文档预览",
    view: "workspace",
    items: [{ label: "文件详情与结构预览", view: "workspace" }],
  },
  {
    title: "处理报告",
    view: "tasks",
    items: [{ label: "报告中心", view: "tasks" }],
  },
  {
    title: "系统设置",
    view: "settings",
    items: [
      { label: "基础设置", view: "settings", anchor: "settings-basic" },
      { label: "转换设置", view: "settings", anchor: "settings-conversion" },
      { label: "OCR 设置", view: "settings", anchor: "settings-ocr" },
      { label: "公式设置", view: "settings", anchor: "settings-formula" },
      { label: "OMML 设置", view: "settings", anchor: "settings-omml" },
      { label: "MathType 格式化设置", view: "settings", anchor: "settings-mathtype" },
      { label: "Word 宏设置", view: "settings", anchor: "settings-macro" },
      { label: "图片设置", view: "settings", anchor: "settings-image" },
      { label: "本地客户端设置", view: "settings", anchor: "settings-local" },
      { label: "网页同步设置", view: "settings", anchor: "settings-cloud" },
      { label: "日志设置", view: "settings", anchor: "settings-log" },
    ],
  },
];

const riskControls = [
  {
    key: "mathtype_compatibility",
    title: "MathType 兼容性复杂",
    risk: "不同 Windows/macOS 和 MathType 版本的公式对象不通用",
    solution: "保留原对象，并使用 MathML/LaTeX 或图片兜底",
    anchor: "settings-mathtype",
  },
  {
    key: "omml_dependency",
    title: "OMML 转 MathType 不稳定",
    risk: "依赖 Office、MathType 和本地 OMML 文件",
    solution: "Windows 本地优先，自动检索依赖并支持手动指定",
    anchor: "settings-omml",
  },
  {
    key: "macro_security",
    title: "宏执行安全风险",
    risk: "宏可能修改文档或执行未知逻辑",
    solution: "默认风险确认、执行前备份、来源授权和白名单控制",
    anchor: "settings-macro",
  },
  {
    key: "pdf_formula_accuracy",
    title: "PDF 公式识别准确率不足",
    risk: "扫描质量和公式复杂度会影响 Mathpix 识别结果",
    solution: "启用公式 OCR、低置信度标记和人工校正",
    anchor: "settings-ocr",
  },
  {
    key: "conversion_layout",
    title: "Word/PPT 转换排版错乱",
    risk: "文档结构差异会导致标题、图片、表格和公式错位",
    solution: "多转换模式、对象保留清单和质量检测报告",
    anchor: "settings-conversion",
  },
  {
    key: "small_image_false_positive",
    title: "微小图片误判",
    risk: "标点、图标、公式截图、二维码和印章容易混淆",
    solution: "阈值配置、类型筛选、人工确认和误判标记",
    anchor: "settings-image",
  },
  {
    key: "large_file_queue",
    title: "文件过大导致卡顿",
    risk: "OCR、转换和批量处理耗时较长",
    solution: "单文件限制、任务队列、并发控制和失败重试",
    anchor: "settings-basic",
  },
  {
    key: "local_file_access",
    title: "网页端无法访问本地文件",
    risk: "浏览器权限限制导致 Office、MathType、宏和本地路径不可直接处理",
    solution: "本地客户端交接、本地任务载荷和状态同步",
    anchor: "settings-local",
  },
  {
    key: "privacy",
    title: "用户隐私风险",
    risk: "文档可能包含敏感内容，第三方上传需明确授权",
    solution: "默认本地处理，Mathpix 外部上传和云端同步均需授权",
    anchor: "settings-cloud",
  },
  {
    key: "local_communication",
    title: "本地客户端与网页通信风险",
    risk: "本地 API 可能被非授权调用",
    solution: "127.0.0.1 监听、本地安全令牌和下载/API 统一保护",
    anchor: "settings-local",
  },
];

const performanceTargets = [
  {
    key: "drag_response",
    metric: "文件拖拽响应",
    target: "3 秒内展示到文件列表",
    route: "前端导入后立即刷新文件队列",
  },
  {
    key: "single_file_size",
    metric: "单文件大小",
    target: "默认支持 500MB，可配置",
    route: "基础设置的单文件限制",
  },
  {
    key: "batch_queue",
    metric: "批量文件数量",
    target: "默认支持 20 个以上排队",
    route: "批量任务队列和当前文件列表",
  },
  {
    key: "concurrency",
    metric: "并发任务数",
    target: "默认 2-4 个，可配置",
    route: "基础设置最大并发",
  },
  {
    key: "word_parse",
    metric: "100 页 Word 解析",
    target: "30 秒内完成基础解析",
    route: "Word 结构预览和公式预检",
  },
  {
    key: "pdf_to_word",
    metric: "100 页 PDF 转 Word",
    target: "根据 OCR 模式动态处理",
    route: "Mathpix OCR 模式和轮询超时",
  },
  {
    key: "small_image_scan",
    metric: "微小图片检索",
    target: "100 页文档 1 分钟内完成基础检索",
    route: "图片阈值和报告任务",
  },
  {
    key: "omml_search",
    metric: "OMML 文件检索",
    target: "常见目录 1 分钟内完成基础检索",
    route: "OMML 扫描文件上限和自定义路径",
  },
  {
    key: "macro_detection",
    metric: "宏列表检测",
    target: "10 秒内完成基础检测",
    route: "宏检测开关和本地客户端队列",
  },
  {
    key: "formula_recognition",
    metric: "公式识别",
    target: "单公式 1-3 秒，复杂公式允许更长",
    route: "Mathpix 公式 OCR 与低置信度策略",
  },
  {
    key: "task_recovery",
    metric: "任务恢复",
    target: "程序异常退出后支持恢复",
    route: "中断任务状态和恢复提示",
  },
];

const exceptionPolicies = [
  {
    key: "file",
    category: "文件异常",
    exceptions: "损坏、加密、过大、格式不支持、文件名非法、重复文件",
    handling: "提示原因，密码由用户输入，重复文件按覆盖/跳过/重命名处理",
    anchor: "settings-basic",
  },
  {
    key: "conversion",
    category: "转换异常",
    exceptions: "Word/PPT/PDF/Excel 转换失败、输出打不开、排版严重错乱",
    handling: "记录错误，支持重试，报告质量异常和失败清单",
    anchor: "settings-conversion",
  },
  {
    key: "formula",
    category: "公式异常",
    exceptions: "MathType/OMML/LaTeX/图片公式读取、识别、格式化或位置异常",
    handling: "保留原公式或原图，标记低置信度，提供人工确认",
    anchor: "settings-formula",
  },
  {
    key: "omml",
    category: "OMML 文件异常",
    exceptions: "找不到依赖、复制失败、不可用、目录无权限、多个文件冲突",
    handling: "自动检索，手动指定，保留原公式并导出 OMML 失败清单",
    anchor: "settings-omml",
  },
  {
    key: "macro",
    category: "宏异常",
    exceptions: "不支持宏、宏被禁用、权限不足、执行失败、超时、依赖缺失",
    handling: "风险确认，按失败策略处理，执行前备份，支持恢复和失败清单",
    anchor: "settings-macro",
  },
  {
    key: "image",
    category: "图片异常",
    exceptions: "图片提取失败、误判、定位失败、导出失败、重复判断失败",
    handling: "记录错误，保留原始结果，支持人工标记、重新导出和替换待处理",
    anchor: "settings-image",
  },
  {
    key: "system",
    category: "系统异常",
    exceptions: "磁盘空间、权限、Office/MathType/OCR、本地客户端、API、任务中断",
    handling: "运行预检，提示修复，任务中断可恢复或重试，保留任务状态",
    anchor: "settings-local",
  },
];

const acceptanceGroups = [
  {
    key: "upload",
    section: "17.1",
    title: "文件上传验收",
    total: 9,
    evidence: "拖拽/文件夹导入、格式校验、加密提示、文件列表",
    view: "workspace",
  },
  {
    key: "word",
    section: "17.2",
    title: "Word 处理验收",
    total: 10,
    evidence: "Word 解析、Word 转 PPT、对象保留清单、MathType 格式化报告",
    view: "word",
  },
  {
    key: "omml",
    section: "17.3",
    title: "Word 自带公式与 OMML 验收",
    total: 10,
    evidence: "OMML 检测、转换确认、依赖检索/复制、OMML 失败清单",
    view: "word",
    anchor: "settings-omml",
  },
  {
    key: "macro",
    section: "17.4",
    title: "Word 宏验收",
    total: 10,
    evidence: "宏库、宏详情、顺序调整、备份、失败策略和本地队列",
    view: "macro",
    anchor: "settings-macro",
  },
  {
    key: "ppt",
    section: "17.5",
    title: "PPT 处理验收",
    total: 8,
    evidence: "PPT 解析、PPT 转 Word、备注/图片/公式提取和转换日志",
    view: "word",
  },
  {
    key: "pdf",
    section: "17.6",
    title: "PDF 处理验收",
    total: 9,
    evidence: "PDF 类型识别、Mathpix OCR 计划、低置信度公式和 PDF 转换报告",
    view: "word",
    anchor: "settings-ocr",
  },
  {
    key: "mathtype",
    section: "17.7",
    title: "MathType 公式验收",
    total: 10,
    evidence: "MathType 保留、OMML/LaTeX/PDF 公式、格式化对比和公式报告",
    view: "formula",
    anchor: "settings-mathtype",
  },
  {
    key: "images",
    section: "17.8",
    title: "微小图片验收",
    total: 12,
    evidence: "Word/PPT/Excel/PDF 小图检索、筛选、标注、导出和图片报告",
    view: "images",
  },
  {
    key: "batch",
    section: "17.9",
    title: "批量处理验收",
    total: 8,
    evidence: "排队、整体/单文件进度、重试、跳过、打包下载和日志导出",
    view: "tasks",
  },
  {
    key: "hybrid",
    section: "17.10",
    title: "本地 + 网页双模式验收",
    total: 10,
    evidence: "本地安全令牌、能力分流、本地任务载荷、状态同步和安装画像",
    view: "settings",
    anchor: "settings-local",
  },
];

const versionRoadmap = [
  {
    version: "V1.0",
    goal: "基础本地可用版",
    features: "拖拽上传、文件列表、Word 转 PPT、PPT 转 Word、PDF 转 Word、图片提取",
    acceptanceKeys: ["upload", "word", "ppt", "pdf", "images"],
  },
  {
    version: "V1.1",
    goal: "微小图片版",
    features: "微小图片检索、筛选、定位、导出、图片报告",
    acceptanceKeys: ["images"],
  },
  {
    version: "V1.2",
    goal: "公式增强版",
    features: "MathType 识别、MathType 保留、PDF 公式识别、Word 公式 MathType 格式化",
    acceptanceKeys: ["pdf", "mathtype"],
  },
  {
    version: "V1.3",
    goal: "OMML 增强版",
    features: "Word 自带公式检测、OMML 转 MathType、OMML 文件自动检索与复制",
    acceptanceKeys: ["omml", "mathtype"],
  },
  {
    version: "V1.4",
    goal: "宏处理增强版",
    features: "Word 宏检测、选择、排序、顺序执行、宏执行报告",
    acceptanceKeys: ["macro"],
  },
  {
    version: "V1.5",
    goal: "批量增强版",
    features: "批量转换、批量公式处理、批量宏处理、批量图片检索、失败重试、日志导出",
    acceptanceKeys: ["batch", "macro", "images", "mathtype"],
  },
  {
    version: "V1.6",
    goal: "网页管理版",
    features: "用户登录、任务管理、模板管理、报告管理、权限管理",
    acceptanceKeys: ["batch", "hybrid"],
  },
  {
    version: "V2.0",
    goal: "混合模式版",
    features: "本地客户端 + 网页端联动，本地处理复杂文档，网页管理任务",
    acceptanceKeys: ["hybrid", "omml", "macro", "mathtype"],
  },
  {
    version: "V2.1",
    goal: "质量检测版",
    features: "转换质量检查、低置信度标记、异常对象定位",
    acceptanceKeys: ["word", "pdf", "mathtype", "images"],
  },
  {
    version: "V3.0",
    goal: "智能增强版",
    features: "AI 排版修复、智能模板套用、私有化部署、API 接口开放",
    acceptanceKeys: [],
    planned: true,
  },
];

const dataDictionary = [
  {
    key: "files",
    name: "FileItem",
    title: "文件对象",
    privacy: "file_path、storage_path 默认脱敏，开启显示本地路径后才返回真实路径",
    fields: [
      ["id", "string", "文件 ID"],
      ["file_name", "string", "文件名"],
      ["file_type", "string", "文件类型"],
      ["file_size", "number", "文件大小"],
      ["file_path", "string", "文件路径"],
      ["page_count", "number", "页数"],
      ["slide_count", "number", "PPT 页数"],
      ["sheet_count", "number", "Excel Sheet 数"],
      ["status", "string", "状态"],
      ["has_formula", "boolean", "是否含公式"],
      ["has_mathtype", "boolean", "是否含 MathType"],
      ["has_omml", "boolean", "是否含 Word 自带公式"],
      ["has_macro", "boolean", "是否含宏"],
      ["has_small_image", "boolean", "是否含微小图片"],
      ["created_at", "datetime", "创建时间"],
    ],
  },
  {
    key: "tasks",
    name: "Task",
    title: "任务对象",
    privacy: "input_path、output_path 默认按本地路径隐私策略脱敏",
    fields: [
      ["id", "string", "任务 ID"],
      ["task_type", "string", "任务类型"],
      ["execute_mode", "string", "local / web / hybrid"],
      ["file_ids", "array", "文件 ID 列表"],
      ["options", "object", "任务参数，包含 k12.workflowPlan.v1 流程顺序"],
      ["status", "string", "任务状态"],
      ["progress", "number", "进度"],
      ["input_path", "string", "输入路径"],
      ["output_path", "string", "输出路径"],
      ["error_message", "string", "错误信息"],
      ["start_time", "datetime", "开始时间"],
      ["end_time", "datetime", "结束时间"],
    ],
  },
  {
    key: "formulas",
    name: "FormulaItem",
    title: "公式对象",
    privacy: "original_image_path 默认脱敏，报告里只展示引用或受控资源链接",
    fields: [
      ["id", "string", "公式 ID"],
      ["file_id", "string", "文件 ID"],
      ["page_index", "number", "页码"],
      ["position", "string", "位置信息"],
      ["position_status", "string", "已记录 / 异常位置"],
      ["position_issue", "string", "位置异常说明"],
      ["fallback_position", "string", "位置缺失时的异常位置记录"],
      ["source_type", "string", "MathType / OMML / LaTeX / PDF / 图片"],
      ["original_image_path", "string", "原始截图"],
      ["latex", "string", "LaTeX 结果"],
      ["mathml", "string", "MathML 结果"],
      ["mathtype_data", "string", "MathType 数据"],
      ["format_status", "string", "是否已格式化"],
      ["confidence", "number", "置信度"],
      ["status", "string", "成功 / 失败 / 待确认"],
    ],
  },
  {
    key: "formula_annotations",
    name: "FormulaAnnotation",
    title: "公式人工校正对象",
    privacy: "重识别请求只保存授权门槛、阻断原因和下一步，不保存 Mathpix 凭证值",
    fields: [
      ["id", "string", "校正记录 ID"],
      ["formula_id", "string", "公式 ID"],
      ["report_id", "string", "报告 ID"],
      ["status", "string", "已确认 / 已修正 / 重新识别 / 跳过"],
      ["latex", "string", "人工修正 LaTeX"],
      ["mathml", "string", "人工修正 MathML"],
      ["retry_recognition", "boolean", "是否请求重新识别"],
      ["recognition_status", "string", "重识别请求状态"],
      ["recognition_request", "object", "k12.formulaRecognitionRequest.v1 请求合同"],
      ["next_step", "string", "下一步处理建议"],
      ["note", "string", "人工备注"],
      ["updated_at", "datetime", "更新时间"],
    ],
  },
  {
    key: "omml",
    name: "OmmlDependencyItem",
    title: "OMML 依赖对象",
    privacy: "document_path、omml_source_path、omml_target_path 默认脱敏",
    fields: [
      ["id", "string", "依赖 ID"],
      ["file_id", "string", "文件 ID"],
      ["document_path", "string", "当前 Word 文档路径"],
      ["omml_file_name", "string", "OMML 文件名称"],
      ["omml_source_path", "string", "OMML 文件来源路径"],
      ["omml_target_path", "string", "复制到的目标路径"],
      ["found_status", "string", "已找到 / 未找到 / 手动选择"],
      ["copy_status", "string", "成功 / 失败 / 跳过"],
      ["error_message", "string", "错误信息"],
      ["created_at", "datetime", "创建时间"],
    ],
  },
  {
    key: "macros",
    name: "MacroItem",
    title: "宏对象",
    privacy: "backup_path 默认脱敏，恢复动作只接受当前任务记录的备份",
    fields: [
      ["id", "string", "宏 ID"],
      ["file_id", "string", "文件 ID"],
      ["macro_name", "string", "宏名称"],
      ["macro_source", "string", "当前文档 / 模板 / 本地宏库 / 系统内置"],
      ["macro_description", "string", "宏说明"],
      ["execute_order", "number", "执行顺序"],
      ["execute_timing", "string", "执行时机"],
      ["execute_status", "string", "成功 / 失败 / 跳过"],
      ["failure_strategy", "string", "停止 / 跳过 / 询问"],
      ["error_message", "string", "错误信息"],
      ["backup_path", "string", "执行前备份路径"],
    ],
  },
  {
    key: "images",
    name: "SmallImageItem",
    title: "微小图片对象",
    privacy: "image_path 默认脱敏，缩略图和导出使用受控资源地址",
    fields: [
      ["id", "string", "图片 ID"],
      ["file_id", "string", "文件 ID"],
      ["page_index", "number", "页码"],
      ["location", "string", "所在位置"],
      ["image_path", "string", "图片路径"],
      ["width", "number", "宽度"],
      ["height", "number", "高度"],
      ["area", "number", "面积"],
      ["image_type", "string", "图片类型"],
      ["is_formula_like", "boolean", "是否疑似公式"],
      ["is_icon_like", "boolean", "是否疑似图标"],
      ["is_qrcode_like", "boolean", "是否疑似二维码"],
      ["is_stamp_like", "boolean", "是否疑似印章"],
      ["is_signature_like", "boolean", "是否疑似签名"],
      ["is_duplicate", "boolean", "是否重复"],
      ["duplicate_check_status", "string", "重复判断状态"],
      ["duplicate_fallback", "string", "重复判断异常兜底"],
      ["export_status", "string", "可导出 / 已重新导出 / 重新导出失败 / 原始流不可用"],
      ["export_message", "string", "图片导出或重新导出说明"],
      ["reexported_at", "datetime", "最近重新导出时间"],
      ["confidence", "number", "置信度"],
    ],
  },
  {
    key: "reports",
    name: "ReportItem",
    title: "报告对象",
    privacy: "report_path 和导出路径默认脱敏，下载走受令牌保护的资源接口",
    fields: [
      ["id", "string", "报告 ID"],
      ["task_id", "string", "任务 ID"],
      ["file_id", "string", "文件 ID"],
      ["report_type", "string", "报告类型"],
      ["success_count", "number", "成功数量"],
      ["fail_count", "number", "失败数量"],
      ["formula_count", "number", "公式数量"],
      ["omml_count", "number", "OMML 公式数量"],
      ["omml_converted_count", "number", "OMML 转换成功数量"],
      ["macro_count", "number", "宏数量"],
      ["macro_success_count", "number", "宏执行成功数量"],
      ["macro_fail_count", "number", "宏执行失败数量"],
      ["formatted_formula_count", "number", "已格式化公式数量"],
      ["small_image_count", "number", "微小图片数量"],
      ["error_count", "number", "错误数量"],
      ["report_path", "string", "报告路径"],
      ["created_at", "datetime", "创建时间"],
    ],
  },
];

const securityPrivacyControls = [
  {
    key: "file_security",
    title: "文件安全",
    requirement: "默认本地处理、加密文件主动输入密码、第三方上传需授权、可清理历史和缓存",
    anchor: "settings-basic",
  },
  {
    key: "macro_security_prd",
    title: "宏安全",
    requirement: "不自动执行宏、风险确认、执行前备份、白名单和来源授权、本地客户端执行",
    anchor: "settings-macro",
  },
  {
    key: "local_web_security",
    title: "本地与网页通信安全",
    requirement: "127.0.0.1、本地客户端授权、安全令牌、状态同步授权、本地路径默认隐藏",
    anchor: "settings-local",
  },
  {
    key: "privacy_protection",
    title: "隐私保护",
    requirement: "日志不记录完整正文、报告不默认展示完整文档、可关闭历史、自动清理周期",
    anchor: "settings-log",
  },
  {
    key: "offline_private",
    title: "离线与私有化",
    requirement: "本地离线模式、私有化部署规划、云端默认仅摘要",
    anchor: "settings-cloud",
  },
];

const compatibilityMatrix = [
  {
    key: "operating_systems",
    title: "操作系统",
    requirement: "Windows 10/11 优先，macOS 和 Linux 可选支持",
    items: ["Windows 10", "Windows 11", "macOS", "Linux 可选"],
    anchor: "settings-local",
  },
  {
    key: "office_suite",
    title: "Office 兼容",
    requirement: "Microsoft Office、WPS Office、LibreOffice，以及 Word/PPT/Excel 2007 及以上",
    items: ["Microsoft Office", "WPS Office", "LibreOffice", "Word/PPT/Excel 2007+"],
    anchor: "settings-local",
  },
  {
    key: "formula_formats",
    title: "公式兼容",
    requirement: "MathType、OMML、MathML、LaTeX、图片公式和 PDF 公式均需可追踪",
    items: ["MathType", "Word 原生 OMML", "MathML", "LaTeX", "图片公式", "PDF 公式"],
    anchor: "settings-formula",
  },
  {
    key: "macro_formats",
    title: "宏兼容",
    requirement: "Word 文档宏、模板宏、.docm/.dotm、系统内置宏和本地宏库",
    items: ["Word 文档宏", "Word 模板宏", ".docm", ".dotm", "系统内置宏", "本地宏库"],
    anchor: "settings-macro",
  },
  {
    key: "mode_matrix",
    title: "模式兼容",
    requirement: "Windows 本地端、macOS 本地端和网页端按任务能力分流",
    items: ["Word/PPT 转换", "PDF 转 Word", "MathType 格式化", "OMML 转 MathType", "Word 宏执行", "微小图片检索", "批量任务管理"],
    anchor: "settings-local",
  },
];

const modeLabels = {
  local: "本地",
  web: "网页",
  hybrid: "混合",
  auto: "自动分流",
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));

async function api(path, options = {}) {
  const headers = {
    ...(options.body instanceof FormData ? {} : { "Content-Type": "application/json" }),
    ...authHeaders(),
    ...(options.headers || {}),
  };
  const response = await fetch(path, {
    ...options,
    headers,
  });
  const data = await response.json();
  if (!response.ok) {
    const error = new Error(data.error || "请求失败");
    error.status = response.status;
    throw error;
  }
  return data;
}

function authHeaders() {
  return state.localToken ? { "X-K12-Token": state.localToken } : {};
}

function readStoredToken() {
  try {
    return localStorage.getItem(tokenStorageKey) || "";
  } catch {
    return "";
  }
}

function rememberSecurityToken(token) {
  state.localToken = token || "";
  try {
    if (state.localToken) {
      localStorage.setItem(tokenStorageKey, state.localToken);
    } else {
      localStorage.removeItem(tokenStorageKey);
    }
  } catch {
    return;
  }
}

function rememberResponseSecurityToken(settings) {
  const token = String(settings?.localSecurityToken || "").trim();
  if (token) rememberSecurityToken(token);
}

function hasLocalSecurityCredential() {
  return Boolean(String(state.localToken || state.settings.localSecurityToken || "").trim());
}

function hasLocalSecurityToken() {
  return hasLocalSecurityCredential() || Boolean(state.settings.localSecurityTokenConfigured);
}

function withToken(path) {
  if (!state.localToken) return path;
  const url = new URL(path, window.location.origin);
  url.searchParams.set("token", state.localToken);
  return `${url.pathname}${url.search}${url.hash}`;
}

function absoluteUrl(path) {
  return new URL(path, window.location.origin).toString();
}

async function loadAll() {
  const health = await api("/api/health");
  let loaded;
  try {
    loaded = await Promise.all([
      api("/api/files"),
      api("/api/tasks"),
      api("/api/reports"),
      api("/api/logs"),
      api("/api/settings"),
      api("/api/capabilities"),
      api("/api/local-client/manifest"),
      api("/api/local-client/uploads"),
      api("/api/architecture"),
      api("/api/api-catalog"),
      api("/api/data-dictionary"),
      api("/api/acceptance-matrix"),
      api("/api/product-summary"),
      api("/api/enhancement-plan"),
      api("/api/install-profile"),
      api("/api/install-plan"),
      api("/api/macros"),
      api("/api/macro-templates"),
      api("/api/users"),
      api("/api/templates"),
      api("/api/authorizations"),
      api("/api/preflight"),
      api("/api/mathpix-jobs"),
      api("/api/image-annotations"),
      api("/api/formula-annotations"),
      api("/api/omml-annotations"),
      api("/api/layout-annotations"),
    ]);
  } catch (error) {
    if (error.status === 401) {
      const token = window.prompt("请输入本地 API 安全令牌", state.localToken || "");
      if (token !== null) {
        rememberSecurityToken(token.trim());
        return loadAll();
      }
    }
    throw error;
  }
  const [
    files,
    tasks,
    reports,
    logs,
    settings,
    capabilities,
    localClientManifest,
    localUploadQueue,
    architecture,
    apiCatalog,
    dataDictionary,
    acceptanceMatrix,
    productSummary,
    enhancementPlan,
    profile,
    installPlan,
    macros,
    macroTemplates,
    users,
    templates,
    authorizations,
    preflight,
    mathpixJobs,
    imageAnnotations,
    formulaAnnotations,
    ommlAnnotations,
    layoutAnnotations,
  ] = loaded;
  state.files = files.files;
  state.tasks = tasks.tasks;
  state.reports = reports.reports;
  state.logs = logs.logs;
  if (state.activeLogTaskId && state.tasks.some((task) => task.id === state.activeLogTaskId)) {
    state.taskLogs = (await api(`/api/logs?task_id=${encodeURIComponent(state.activeLogTaskId)}`)).logs;
  } else {
    state.activeLogTaskId = "";
    state.taskLogs = [];
  }
  state.settings = settings.settings;
  rememberResponseSecurityToken(state.settings);
  state.capabilities = capabilities.capabilities;
  state.localClientManifest = localClientManifest.localClientManifest;
  state.localUploadQueue = localUploadQueue.uploadQueue;
  state.architecture = architecture.architecture;
  state.apiCatalog = apiCatalog.apiCatalog;
  state.dataDictionary = dataDictionary.dataDictionary;
  state.acceptanceMatrix = acceptanceMatrix.acceptanceMatrix;
  state.productSummary = productSummary.productSummary;
  state.enhancementPlan = enhancementPlan.enhancementPlan;
  state.installProfile = profile.installProfile;
  state.installPlan = installPlan.installPlan;
  state.macros = macros.macros;
  state.macroTemplates = macroTemplates.templates;
  state.users = users.users;
  state.currentUser = users.currentUser;
  state.templates = templates.templates;
  state.authorizations = authorizations.authorizations;
  state.preflightChecks = preflight.checks;
  state.mathpixJobs = mathpixJobs.mathpixJobs;
  state.imageAnnotations = imageAnnotations.annotations;
  state.formulaAnnotations = formulaAnnotations.annotations;
  state.ommlAnnotations = ommlAnnotations.annotations;
  state.layoutAnnotations = layoutAnnotations.annotations;
  state.selectedFiles = new Set([...state.selectedFiles].filter((id) => state.files.some((file) => file.id === id)));
  if (state.activeTaskDetailId && !state.tasks.some((task) => task.id === state.activeTaskDetailId)) state.activeTaskDetailId = "";
  if (state.activeReportId && !state.reports.some((report) => report.id === state.activeReportId)) state.activeReportId = "";
  state.macroSequence = state.macroSequence.filter((id) => state.macros.some((macro) => macro.id === id));
  if (state.activeMacroId && !state.macros.some((macro) => macro.id === state.activeMacroId)) state.activeMacroId = "";
  await refreshLocalReadiness();
  $("#serviceStatus").textContent = health.status === "ok" ? "运行中" : "异常";
  $("#localStatus").textContent = `${state.capabilities.host}:${state.capabilities.port}`;
  $("#localDot").classList.add("ok");
  syncSettingsToForm();
  render();
}

function render() {
  renderMenuStructure();
  renderModeStrip();
  renderWorkspaceMetrics();
  renderWorkspaceProgress();
  renderWorkspaceRecentTasks();
  renderCompletionNotices();
  renderTaskPlanner();
  renderRouteAdvisor();
  renderQueueTabs();
  renderFiles();
  renderFileDetail();
  renderFilePreview();
  renderTasks();
  renderTaskDetail();
  renderReports();
  renderLogs();
  renderCapabilities();
  renderInstallProfile();
  renderLocalHandoff();
  renderRiskRegister();
  renderPerformanceTargets();
  renderExceptionPolicies();
  renderAcceptanceOverview();
  renderAcceptanceMatrix();
  renderVersionRoadmap();
  renderDataDictionary();
  renderSecurityPrivacy();
  renderCompatibilityMatrix();
  renderArchitectureBlueprint();
  renderApiCatalog();
  renderProductSummary();
  renderEnhancementPlan();
  renderAdmin();
  renderWordWorkspace();
  renderWordInsight();
  renderWordFlow();
  renderPdfWorkspace();
  renderPptWorkspace();
  renderOmmlDependencies();
  renderMacros();
  renderMacroDetail();
  renderMacroTemplates();
  renderMacroResults();
  renderMacroLogs();
  renderImages();
  renderSelectionBar();
}

function renderModeStrip() {
  const cap = state.capabilities;
  const selected = contextualFiles();
  const localNeeded = selected.filter((file) => file.has_omml || file.has_macro || file.has_mathtype).length;
  const webReady = selected.filter((file) => ["PDF", "Excel", "图片"].includes(file.file_type)).length;
  const hybridReady = selected.filter((file) => ["Word", "PPT"].includes(file.file_type) || file.has_small_image).length;
  $("#localModeBadge").textContent = cap?.officeAutomation?.available ? "可执行" : "待本地客户端";
  $("#webModeBadge").textContent = webReady ? `${webReady} 个可走网页` : "任务管理";
  $("#hybridModeBadge").textContent = hybridReady || localNeeded ? `${hybridReady + localNeeded} 个建议混合` : "自动分流";
  $("#localModeDetail").textContent = cap?.officeAutomation?.status || "等待能力检测";
  $("#webModeDetail").textContent = state.settings.allowCloudSync ? "已允许云端同步" : "默认本地元数据与任务管理";
  $("#hybridModeDetail").textContent = localNeeded ? "选中文件含 OMML、MathType 或宏，建议本地执行" : "网页编排，本地处理复杂文档";
}

function renderMenuStructure() {
  const target = $("#prdMenuStructure");
  if (!target) return;
  target.innerHTML = menuStructure
    .map((group) => {
      const items = group.items
        .map((item) => {
          const attrs = [
            `data-menu-view="${escapeHtml(item.view || group.view)}"`,
            item.task ? `data-menu-task="${escapeHtml(item.task)}"` : "",
            item.anchor ? `data-menu-anchor="${escapeHtml(item.anchor)}"` : "",
          ]
            .filter(Boolean)
            .join(" ");
          const mode = item.task ? modeLabels[taskSpecs[item.task]?.mode] || "任务" : viewMeta[item.view || group.view]?.[0] || "入口";
          return `<li><button class="menu-leaf" type="button" ${attrs}><span>${escapeHtml(item.label)}</span><small>${escapeHtml(mode)}</small></button></li>`;
        })
        .join("");
      const attrs = [
        `data-menu-view="${escapeHtml(group.view)}"`,
        group.anchor ? `data-menu-anchor="${escapeHtml(group.anchor)}"` : "",
      ]
        .filter(Boolean)
        .join(" ");
      return `<article class="menu-group">
        <button class="menu-group-head" type="button" ${attrs}>
          <span>${escapeHtml(group.title)}</span>
          ${lineIcon("arrow-right")}
        </button>
        <ul>${items}</ul>
      </article>`;
    })
    .join("");
}

function renderWorkspaceMetrics() {
  const counts = fileStats(state.files);
  const taskCounts = taskStats();
  const metrics = [
    ["公式", counts.formula, "识别对象"],
    ["OMML", counts.omml, "Word 自带公式"],
    ["宏", counts.macro, "需风险确认"],
    ["微小图", counts.smallImage, "可生成图片报告"],
    ["处理文件", state.files.length, "当前队列"],
  ];
  $("#workspaceMetrics").innerHTML = metrics.map(metricCard).join("");
  $("#taskMetrics").innerHTML = [
    ["待处理", taskCounts.pending, "等待队列"],
    ["处理中", taskCounts.running, "显示实时进度"],
    ["成功", taskCounts.success, "可下载报告"],
    ["失败", taskCounts.failed, "支持重试"],
  ].map(metricCard).join("");
}

function renderQueueTabs() {
  const counts = fileStats(state.files);
  const values = {
    all: state.files.length,
    Word: counts.Word,
    PDF: counts.PDF,
    PPT: counts.PPT,
    Excel: counts.Excel,
    异常: state.files.filter((file) => file.status !== "待处理" && file.status !== "成功").length,
  };
  $$(".queue-tabs .segment").forEach((button) => {
    const baseLabel = button.dataset.label || button.textContent.replace(/\s*\(\d+\)\s*$/, "").trim();
    button.dataset.label = baseLabel;
    button.textContent = `${baseLabel} (${values[button.dataset.filter] || 0})`;
    button.classList.toggle("active", state.filter === button.dataset.filter);
  });
}

function renderTaskPlanner() {
  const files = contextualFiles();
  const tasks = orderedPlannerTasks(files);
  $("#plannerScope").textContent = files.length ? `${files.length} 个文件` : "未选择文件";
  $("#taskPlanner").innerHTML = tasks
    .map((taskType) => {
      const spec = taskSpecs[taskType];
      const mode = resolveExecuteMode(taskType, files);
      return `<article class="planner-item" data-task="${taskType}">
        <button type="button" class="planner-drag-handle" draggable="true" data-task="${taskType}" title="拖拽或用方向键调整流程顺序" aria-label="拖拽或用方向键调整流程顺序">${lineIcon("drag-lines", "drag-line-icon")}</button>
        <span class="planner-enabled-check" title="已加入处理流程" aria-label="已加入处理流程" role="img">${lineIcon("check", "planner-check-icon")}</span>
        <div class="planner-copy">
          <strong>${spec.label}</strong>
          <small>${spec.description}</small>
          <span class="badge ${modeClass(mode)}">${modeLabels[mode]}</span>
          <span class="badge blue">${spec.output}</span>
        </div>
        <div class="planner-actions">
          <button class="mini-button planner-task" data-task="${taskType}">使用</button>
        </div>
      </article>`;
    })
    .join("");
}

function lineIcon(name, className = "") {
  const icons = {
    "drag-lines": '<path d="M7 7h10"></path><path d="M7 12h10"></path><path d="M7 17h10"></path>',
    grip: '<path d="M8 5h8M8 10h8M8 15h8"></path>',
    check: '<path d="m6 12 4 4 8-9"></path>',
    "arrow-right": '<path d="M5 12h14"></path><path d="m13 6 6 6-6 6"></path>',
    "arrow-up": '<path d="M12 19V5"></path><path d="m6 11 6-6 6 6"></path>',
    "arrow-down": '<path d="M12 5v14"></path><path d="m18 13-6 6-6-6"></path>',
  };
  const classes = ["line-icon", className].filter(Boolean).join(" ");
  return `<svg class="${classes}" viewBox="0 0 24 24" aria-hidden="true">${icons[name] || icons.grip}</svg>`;
}

function renderRouteAdvisor() {
  const files = contextualFiles();
  const taskType = $("#taskTypeSelect")?.value || defaultPlannerTask(files);
  const mode = resolveExecuteMode(taskType, files);
  const spec = taskSpecs[taskType];
  const blockers = compatibilityIssues(taskType, files);
  const details = [
    `任务：${spec.label}`,
    `推荐：${modeLabels[mode]}模式`,
    files.length ? `文件：${files.length} 个` : "文件：未选择",
  ];
  $("#routeAdvisor").innerHTML = `<div class="route-card ${modeClass(mode)}">
    <strong>${modeLabels[mode]}模式</strong>
    <small>${spec.description}</small>
  </div>
  <div class="route-list">
    ${details.map((line) => `<span>${escapeHtml(line)}</span>`).join("")}
  </div>
  ${
    blockers.length
      ? `<div class="notice-list">${blockers.map((item) => `<span>${escapeHtml(item)}</span>`).join("")}</div>`
      : `<div class="notice-list good"><span>当前选择可创建任务</span></div>`
  }`;
}

function renderSelectionBar() {
  const selected = selectedFileObjects();
  const bar = $("#selectionBar");
  const toolbarCount = $("#toolbarSelectionCount");
  const issues = selected.flatMap((file) => file.validation_errors || []);
  if (toolbarCount) toolbarCount.textContent = `已选择 ${selected.length} 个文件`;
  if (!selected.length) {
    bar.classList.add("muted");
    $("#selectionTitle").textContent = "未选择文件";
    $("#selectionHint").textContent = "选择文件后可创建转换、公式预检、宏执行或图片检索任务";
    $("#downloadSelectedButton").disabled = true;
    return;
  }
  bar.classList.remove("muted");
  $("#downloadSelectedButton").disabled = false;
  $("#selectionTitle").textContent = `已选择 ${selected.length} 个文件`;
  $("#selectionHint").textContent = issues.length
    ? `${issues.length} 条校验提示，建议先查看文件详情`
    : `${selected.filter((file) => file.has_formula).length} 个含公式，${selected.filter((file) => file.has_macro).length} 个含宏`;
}

function renderFiles() {
  const tbody = $("#fileTableBody");
  const files = filteredFiles();
  $("#fileCounter").textContent = `${state.files.length} 个文件`;
  syncSelectAll(files);
  if (!files.length) {
    tbody.innerHTML = $("#emptyTemplate").innerHTML;
    return;
  }
  tbody.innerHTML = files
    .map((file) => {
      const checked = state.selectedFiles.has(file.id) ? "checked" : "";
      const active = state.activeFileId === file.id ? "active-row" : "";
      const validationError = primaryValidationError(file);
      return `<tr class="${active}">
        <td><input type="checkbox" class="file-check" data-id="${file.id}" ${checked}></td>
        <td>
          <span class="file-name" title="${escapeHtml(file.file_name)}"><span class="file-type-icon ${fileTypeClass(file.file_type)}">${fileTypeShort(file.file_type)}</span>${escapeHtml(file.file_name)}</span>
          ${file.source_relative_path ? `<small class="table-note">${escapeHtml(file.source_relative_path)}</small>` : ""}
          ${file.encrypted ? '<small class="table-note">需要密码</small>' : ""}
          ${validationError ? `<small class="table-note error-note" title="${escapeHtml(validationError)}">${escapeHtml(validationError)}</small>` : ""}
        </td>
        <td>${escapeHtml(file.file_type)}</td>
        <td>${formatBytes(file.file_size)}</td>
        <td>${escapeHtml(fileMetric(file))}</td>
        <td>${fileCapabilityBadges(file)}</td>
        <td>${formulaCount(file)}</td>
        <td>${ommlCount(file)}</td>
        <td>${macroCount(file)}</td>
        <td><span class="badge ${statusClass(file.status)}">${escapeHtml(file.status)}</span></td>
        <td><div class="row-actions">
          <button class="mini-button file-inspect" data-id="${file.id}" title="查看详情">详情</button>
          ${file.encrypted ? `<button class="mini-button file-password" data-id="${file.id}" ${permissionButtonAttrs(true, "输入密码", "files.manage")}>密码</button>` : ""}
          <button class="mini-button file-preview" data-id="${file.id}" title="生成预览">预览</button>
          <button class="mini-button file-convert" data-id="${file.id}" ${permissionButtonAttrs(true, "创建转换任务", "tasks.create")}>转换</button>
          <button class="mini-button file-scan" data-id="${file.id}" ${permissionButtonAttrs(true, "创建检索任务", "tasks.create")}>检索</button>
          <button class="mini-button file-download" data-id="${file.id}" title="下载源文件">下载</button>
          <button class="mini-button file-reports" data-id="${file.id}" title="查看该文件关联报告">报告</button>
          <button class="mini-button file-replace" data-id="${file.id}" ${permissionButtonAttrs(true, "重新上传替换", "files.manage")}>替换</button>
          <button class="mini-button file-delete danger" data-id="${file.id}" ${permissionButtonAttrs(true, "删除文件", "files.manage")}>删除</button>
        </div></td>
      </tr>`;
    })
    .join("");
}

function renderFileDetail() {
  const selected = selectedFileObjects();
  const file = state.files.find((item) => item.id === state.activeFileId) || selected[0] || state.files[0];
  if (!file) {
    $("#fileDetail").innerHTML = `<div class="empty-mini">暂无文件</div>`;
    return;
  }
  const errors = file.validation_errors?.length ? file.validation_errors.join("；") : "无";
  const summary = file.content_summary || {};
  const summaryRows = Object.entries(summary).slice(0, 6);
  $("#fileDetail").innerHTML = `<div class="detail-head">
    <strong>${escapeHtml(file.file_name)}</strong>
    <span class="badge ${statusClass(file.status)}">${escapeHtml(file.status)}</span>
  </div>
  <dl>
    <dt>类型</dt><dd>${escapeHtml(file.file_type)} ${escapeHtml(file.extension || "")}</dd>
    <dt>大小</dt><dd>${formatBytes(file.file_size)}</dd>
    <dt>页/页签</dt><dd>${fileMetric(file)}</dd>
    <dt>校验</dt><dd>${escapeHtml(errors)}</dd>
    <dt>来源层级</dt><dd>${escapeHtml(file.source_relative_path || "单文件上传")}</dd>
    <dt>来源</dt><dd>${escapeHtml(file.file_path || file.storage_path || "上传缓存")}</dd>
  </dl>
  ${
    summaryRows.length
      ? `<div class="summary-tags">${summaryRows
          .map(([key, value]) => `<span>${escapeHtml(key)}: ${escapeHtml(value)}</span>`)
          .join("")}</div>`
      : ""
  }`;
}

function renderFilePreview() {
  const target = $("#filePreviewPanel");
  if (!target) return;
  const preview = state.filePreview;
  if (!preview) {
    target.innerHTML = `<div class="empty-mini">选择文件后点击预览</div>`;
    return;
  }
  const warnings = preview.warnings || [];
  const pages = preview.pages || [];
  const filteredPages = filterPreviewPages(pages);
  const visiblePages = filteredPages.slice(state.previewPageOffset, state.previewPageOffset + 4);
  const objects = preview.objects || [];
  target.innerHTML = `<div class="preview-head">
    <div>
      <strong>${escapeHtml(preview.file_name)}</strong>
      <small>${escapeHtml(preview.file_type)} · ${escapeHtml(preview.mode)} · ${escapeHtml(preview.status)}</small>
    </div>
    <span class="badge ${preview.status === "ready" ? "good" : "warn"}">${escapeHtml(preview.status)}</span>
  </div>
  ${
    warnings.length
      ? `<div class="notice-list preview-jump-list">${warnings.map((item) => previewJumpChip(item, item, "告警定位")).join("")}</div>`
      : ""
  }
  <div class="preview-objects">
    ${
      objects.length
        ? objects
            .map((item) => previewObjectChip(item))
            .join("")
        : `<span><b>0</b>对象<small>未检测到对象</small></span>`
    }
  </div>
  <div class="preview-tools">
    <input id="previewSearchInput" class="input" value="${escapeHtml(state.previewSearch)}" placeholder="搜索预览文本或对象" />
    <button class="mini-button preview-page-prev" ${state.previewPageOffset <= 0 ? "disabled" : ""}>上一页</button>
    <span>${filteredPages.length ? `${state.previewPageOffset + 1}-${Math.min(state.previewPageOffset + 4, filteredPages.length)} / ${filteredPages.length}` : "0 / 0"}</span>
    <button class="mini-button preview-page-next" ${state.previewPageOffset + 4 >= filteredPages.length ? "disabled" : ""}>下一页</button>
    <button class="mini-button preview-zoom-out" ${state.previewZoom <= 0.8 ? "disabled" : ""}>缩小</button>
    <span>${Math.round(state.previewZoom * 100)}%</span>
    <button class="mini-button preview-zoom-in" ${state.previewZoom >= 1.4 ? "disabled" : ""}>放大</button>
  </div>
  <div class="preview-pages" style="--preview-zoom:${state.previewZoom}">
    ${
      visiblePages.length
        ? visiblePages
            .map((page) => `<article>
              <span>${escapeHtml(page.kind || "页面")} ${escapeHtml(page.index || "")}</span>
              <strong>${escapeHtml(page.title || "预览")}</strong>
              ${previewMarkerChips(page.markers || [])}
              <p>${highlightPreviewText(page.text || "")}</p>
            </article>`)
            .join("")
        : `<div class="empty-mini">${state.previewSearch ? "没有匹配的预览内容" : "暂无可显示页面"}</div>`
    }
  </div>`;
}

function filterPreviewPages(pages) {
  const term = state.previewSearch.trim().toLowerCase();
  if (!term) return pages;
  return pages.filter((page) => previewPageSearchText(page).includes(term));
}

function previewObjectChip(item) {
  const label = item.label || item.type || "对象";
  const status = item.status || "";
  const term = [label, status].filter(Boolean).join(" ");
  return previewJumpChip(`<b>${escapeHtml(item.count || 0)}</b>${escapeHtml(label)}<small>${escapeHtml(status)}</small>`, term, "对象定位", true);
}

function previewJumpChip(content, term, title, trustedHtml = false) {
  return `<button class="preview-jump-chip" type="button" data-term="${escapeHtml(term || "")}" title="${escapeHtml(title || "定位到预览内容")}">${trustedHtml ? content : escapeHtml(content)}</button>`;
}

function previewMarkerChips(markers) {
  if (!markers.length) return "";
  return `<div class="preview-markers">${markers
    .map((marker) => `<button class="preview-jump-chip preview-marker ${escapeHtml(marker.level || "blue")}" type="button" data-term="${escapeHtml(marker.term || marker.label || "")}" title="${escapeHtml(marker.message || "预览对象标记")}">${escapeHtml(marker.label || marker.term || "标记")}</button>`)
    .join("")}</div>`;
}

function jumpPreviewTo(term) {
  const nextTerm = String(term || "").trim();
  state.previewSearch = nextTerm;
  const pages = state.filePreview?.pages || [];
  const matchIndex = nextTerm
    ? pages.findIndex((page) => previewPageSearchText(page).includes(nextTerm.toLowerCase()))
    : -1;
  state.previewPageOffset = matchIndex >= 0 ? Math.floor(matchIndex / 4) * 4 : 0;
  renderFilePreview();
  toast(matchIndex >= 0 ? "已定位到预览命中页" : "已应用定位关键词");
}

function previewPageSearchText(page) {
  const markerText = (page.markers || [])
    .map((marker) => `${marker.term || ""} ${marker.label || ""} ${marker.message || ""}`)
    .join(" ");
  return `${page.title || ""} ${page.text || ""} ${page.kind || ""} ${markerText}`.toLowerCase();
}

function highlightPreviewText(text) {
  const escaped = escapeHtml(text || "");
  const term = state.previewSearch.trim();
  if (!term) return escaped;
  const pattern = new RegExp(escapeRegExp(escapeHtml(term)), "gi");
  return escaped.replace(pattern, (match) => `<mark>${match}</mark>`);
}

function renderTasks() {
  const tbody = $("#taskTableBody");
  $("#taskCounter").textContent = `${state.tasks.length} 个任务`;
  if (!state.tasks.length) {
    tbody.innerHTML = `<tr><td colspan="13" class="empty-state">暂无任务</td></tr>`;
    return;
  }
  tbody.innerHTML = state.tasks
    .map((task) => `<tr>
      <td><strong>${escapeHtml(task.task_label || taskLabels[task.task_type] || task.task_type)}</strong><small class="table-note">${escapeHtml(task.id)}</small></td>
      <td><span class="badge ${modeClass(task.execute_mode)}">${modeLabels[task.execute_mode] || task.execute_mode}</span></td>
      <td>${task.file_ids.length}</td>
      <td><span class="badge ${statusClass(task.status)}">${escapeHtml(task.status)}</span>${taskResultSummary(task)}</td>
      <td>${toCount(task.success_count)}</td>
      <td>${toCount(task.fail_count)}</td>
      <td><div class="progress" title="${task.progress}%"><span style="width:${task.progress}%"></span></div></td>
      <td>${escapeHtml(formatDuration(task))}</td>
      <td>${taskFailureReason(task)}</td>
      <td>${formatTime(task.start_time)}</td>
      <td>${formatTime(task.end_time)}</td>
      <td>${reportButtons(task.id)}</td>
      <td>
        ${taskActionButtons(task)}
      </td>
    </tr>`)
    .join("");
}

function taskResultSummary(task) {
  const success = toCount(task.success_count);
  const failed = toCount(task.fail_count);
  const pending = toCount(task.pending_count);
  const cancelled = toCount(task.cancelled_count);
  const failureRows = toCount(task.failure_count);
  if (!success && !failed && !pending && !cancelled && !failureRows) return "";
  const retryable = toCount(task.retryable_count);
  const pendingHint = pending ? ` · 待本地 ${pending}` : "";
  const cancelledHint = cancelled ? ` · 已取消 ${cancelled}` : "";
  const retryHint = retryable ? ` · 可重试 ${retryable}` : "";
  return `<small class="table-note">成功 ${success} / 失败 ${failed}${pendingHint}${cancelledHint}${retryHint}</small>`;
}

function taskFailureReason(task) {
  const reason = String(task.error_message || "").trim();
  if (!reason) return '<span class="failure-reason muted">-</span>';
  return `<span class="failure-reason" title="${escapeHtml(reason)}">${escapeHtml(truncateText(reason, 44))}</span>`;
}

function taskActionButtons(task) {
  const canPause = ["待处理", "处理中"].includes(task.status);
  const canResume = ["已暂停", "已中断"].includes(task.status);
  const canRetry = ["失败", "已取消", "已中断"].includes(task.status);
  const canCancel = !["成功", "失败", "已取消"].includes(task.status);
  const canRestore = hasMacroBackup(task.id);
  return `<div class="row-actions">
    <button class="mini-button task-detail" data-id="${task.id}" title="查看任务详情">详情</button>
    <button class="mini-button task-retry" data-id="${task.id}" ${permissionButtonAttrs(canRetry, "重试", "tasks.control")}>重试</button>
    <button class="mini-button task-pause" data-id="${task.id}" ${permissionButtonAttrs(canPause, "暂停", "tasks.control")}>暂停</button>
    <button class="mini-button task-resume" data-id="${task.id}" ${permissionButtonAttrs(canResume, "继续", "tasks.control")}>继续</button>
    <button class="mini-button task-log" data-id="${task.id}" title="查看任务日志">日志</button>
    <button class="mini-button task-restore" data-id="${task.id}" ${permissionButtonAttrs(canRestore, "恢复宏备份", "tasks.control")}>恢复</button>
    <button class="mini-button task-download" data-id="${task.id}" title="打包下载">打包</button>
    <button class="mini-button task-cancel" data-id="${task.id}" ${permissionButtonAttrs(canCancel, "取消", "tasks.control")}>取消</button>
  </div>`;
}

function hasPermission(permission) {
  const user = state.currentUser || {};
  const permissions = [
    ...(Array.isArray(user.effective_permissions) ? user.effective_permissions : []),
    ...(Array.isArray(user.permissions) ? user.permissions : []),
    ...(Array.isArray(user.role_permissions) ? user.role_permissions : []),
  ];
  return permissions.includes(permission);
}

function permissionButtonAttrs(enabled, title, permission) {
  const allowed = hasPermission(permission);
  const label = allowed ? title : `${title}（缺少 ${permission} 权限）`;
  return `title="${escapeHtml(label)}" ${enabled && allowed ? "" : "disabled"}`;
}

function renderTaskDetail() {
  const target = $("#taskDetailPanel");
  if (!target) return;
  const task = state.tasks.find((item) => item.id === state.activeTaskDetailId) || state.tasks[0];
  if (!task) {
    $("#taskDetailScope").textContent = "未选择任务";
    target.innerHTML = `<div class="empty-mini">暂无任务详情</div>`;
    return;
  }
  state.activeTaskDetailId = task.id;
  const label = task.task_label || taskLabels[task.task_type] || task.task_type;
  const files = (task.file_ids || []).map((id) => state.files.find((file) => file.id === id)).filter(Boolean);
  const report = reportForTask(task.id);
  const options = Object.entries(task.options || {}).slice(0, 8);
  const outputAction = task.output_directory_action?.status ? `${task.output_directory_action.status}：${task.output_directory_action.message || ""}` : "未请求";
  $("#taskDetailScope").textContent = `${label} · ${task.id}`;
  target.innerHTML = `<div class="detail-head">
    <strong>${escapeHtml(label)}</strong>
    <span class="badge ${statusClass(task.status)}">${escapeHtml(task.status)}</span>
  </div>
  <dl>
    <dt>任务 ID</dt><dd>${escapeHtml(task.id)}</dd>
    <dt>执行模式</dt><dd>${escapeHtml(modeLabels[task.execute_mode] || task.execute_mode)}</dd>
    <dt>进度</dt><dd>${escapeHtml(task.progress || 0)}% · ${escapeHtml(formatDuration(task))}</dd>
    <dt>结果</dt><dd>成功 ${toCount(task.success_count)} / 失败 ${toCount(task.fail_count)} / 待本地 ${toCount(task.pending_count)} / 已取消 ${toCount(task.cancelled_count)} / 可重试 ${toCount(task.retryable_count)}</dd>
    <dt>时间</dt><dd>${formatTime(task.start_time)} -> ${formatTime(task.end_time)}</dd>
    <dt>输出</dt><dd>${escapeHtml(task.output_path || "未生成")}</dd>
    <dt>本地动作</dt><dd>${escapeHtml(outputAction)}</dd>
    <dt>失败摘要</dt><dd>${escapeHtml(task.error_message || "无")}</dd>
  </dl>
  <div class="summary-tags">
    ${files.length ? files.map((file) => `<span>${escapeHtml(file.file_name)} · ${escapeHtml(file.file_type)}</span>`).join("") : `<span>未关联文件</span>`}
    ${report ? `<span>报告 ${escapeHtml(report.report_type)} · 失败清单 ${escapeHtml(report.failure_count || 0)}</span>` : `<span>暂无报告</span>`}
    ${task.completion_notice?.message ? `<span>${escapeHtml(task.completion_notice.message)}</span>` : ""}
    ${options.map(([key, value]) => `<span>${escapeHtml(key)}: ${escapeHtml(JSON.stringify(value))}</span>`).join("")}
  </div>
  ${batchProgressRows(report)}
  <div class="row-actions">
    <button class="mini-button task-log" data-id="${task.id}" title="查看任务日志">查看日志</button>
    <button class="mini-button task-download" data-id="${task.id}" title="打包下载">打包下载</button>
    ${report ? downloadButtons(report.id) : ""}
  </div>`;
}

function batchProgressRows(report) {
  const results = report?.analysis?.batchResults || [];
  if (!results.length) return "";
  const taskId = report?.task_id || "";
  return `<section class="batch-progress-list">
    <strong>批量单文件进度</strong>
    ${results
      .map((item) => {
        const progress = Math.max(0, Math.min(100, toCount(item.progress)));
        const canSkip = taskId && item.file_id && item.retryable;
        return `<article>
          <div>
            <span>${escapeHtml(item.file_name || item.file_id || "文件")}</span>
            <small>${escapeHtml(item.file_type || "-")} · ${escapeHtml(item.error_message || item.recommendation || "无异常")}</small>
          </div>
          <span class="badge ${statusClass(item.status)}">${escapeHtml(item.status)}</span>
          <div class="progress" title="${progress}%"><span style="width:${progress}%"></span></div>
          <b>${progress}%</b>
          ${canSkip ? `<button class="mini-button batch-skip-file" data-task="${escapeHtml(taskId)}" data-file="${escapeHtml(item.file_id)}" ${permissionButtonAttrs(true, "跳过失败文件", "tasks.control")}>跳过</button>` : ""}
        </article>`;
      })
      .join("")}
  </section>`;
}

function renderReports() {
  renderFormulaList();
  renderReportCenter();
  renderLatestReportPanel();
  renderComparisonPanel();
  renderLayoutReview();
}

function renderFormulaList() {
  const latest = latestReport();
  const formulas = filteredFormulaItems(latest);
  const active = activeFormulaItem(formulas);
  $("#formulaStats").innerHTML = [
    ["公式", formulas.length, "当前筛选"],
    ["OMML", formulas.filter((item) => item.source_type === "OMML").length, "Word 自带公式"],
    ["待确认", formulas.filter((item) => item.status === "待确认").length, "低置信度"],
    ["位置异常", formulas.filter((item) => item.position_status === "异常位置").length, "需校正"],
  ].map(metricCard).join("");
  $("#formulaList").innerHTML = formulas.length
    ? formulas
        .map((item) => {
          const annotation = formulaAnnotation(item.id, latest?.id);
          const status = annotation?.status || item.status;
          const latex = annotation?.latex || item.latex;
          const note = annotation?.note || "";
          const originalRef = item.original_image_ref || item.original_image_path || "待真实引擎回填";
          const mathtypePreview = item.mathtype_preview || item.mathtype_data || "待真实引擎回填";
          const activeClass = item.id === active?.id ? " active" : "";
          return `<article class="object-card formula-card${activeClass}">
          <div class="object-title">
            <strong>${escapeHtml(item.source_type)} · ${escapeHtml(status)}</strong>
            <span class="badge ${formulaBadgeClass(status, item.confidence)}">${item.confidence}%</span>
          </div>
          <small>${escapeHtml(item.position)} · ${escapeHtml(item.position_status || "已记录")}</small>
          ${item.position_issue ? `<small>${escapeHtml(item.position_issue)}</small>` : ""}
          <small>原始截图：${escapeHtml(originalRef)}</small>
          <small>MathType：${escapeHtml(mathtypePreview)}</small>
          <textarea class="formula-latex-input" data-id="${escapeHtml(item.id)}" rows="3">${escapeHtml(latex)}</textarea>
          ${note ? `<small>${escapeHtml(note)}</small>` : ""}
          <span class="badge blue">${escapeHtml(item.format_status)}</span>
          <div class="row-actions formula-actions">
            <button class="mini-button formula-select" data-id="${escapeHtml(item.id)}">查看</button>
            <button class="mini-button formula-confirm" data-report="${escapeHtml(latest.id)}" data-id="${escapeHtml(item.id)}">确认</button>
            <button class="mini-button formula-skip" data-report="${escapeHtml(latest.id)}" data-id="${escapeHtml(item.id)}">跳过</button>
            <button class="mini-button formula-rerecognize" data-report="${escapeHtml(latest.id)}" data-id="${escapeHtml(item.id)}">重识别</button>
            <button class="mini-button formula-save" data-report="${escapeHtml(latest.id)}" data-id="${escapeHtml(item.id)}">保存LaTeX</button>
          </div>
        </article>`;
        })
        .join("")
    : `<div class="object-card"><small>暂无公式报告</small></div>`;
  renderFormulaDetail(latest, active);
}

function filteredFormulaItems(report) {
  const filter = $("#formulaSourceFilter")?.value || "all";
  let formulas = report?.analysis?.formulas || [];
  if (filter === "low") formulas = formulas.filter((item) => item.status === "待确认" || item.confidence < 80);
  else if (filter !== "all") formulas = formulas.filter((item) => item.source_type === filter);
  return formulas;
}

function activeFormulaItem(formulas) {
  if (!formulas.length) {
    state.activeFormulaId = "";
    return null;
  }
  const active = formulas.find((item) => item.id === state.activeFormulaId) || formulas[0];
  state.activeFormulaId = active.id;
  return active;
}

function renderFormulaDetail(report, item) {
  const target = $("#formulaDetailPanel");
  if (!target) return;
  if (!report || !item) {
    target.innerHTML = `<div class="empty-mini">选择公式后查看校正详情</div>`;
    return;
  }
  const annotation = formulaAnnotation(item.id, report.id);
  const status = annotation?.status || item.status || "待确认";
  const latex = annotation?.latex || item.latex || "";
  const mathml = annotation?.mathml || item.mathml || formulaMathmlFallback(latex);
  const mathtypePreview = item.mathtype_preview || item.mathtype_data || "等待本地 MathType 引擎生成预览";
  const originalRef = item.original_image_ref || item.original_image_path || "";
  const comparison = item.format_comparison || {};
  const policy = item.format_failure_policy || {};
  target.innerHTML = `<div class="formula-detail-head">
    <div>
      <h3>公式校正</h3>
      <small>${escapeHtml(item.position || "未记录位置")} · ${escapeHtml(item.source_type || "未知来源")}</small>
    </div>
    <span class="badge ${formulaBadgeClass(status, item.confidence)}">${escapeHtml(status)} · ${escapeHtml(item.confidence || 0)}%</span>
  </div>
  <section class="formula-preview-block">
    <h4>位置状态</h4>
    <small>${escapeHtml(item.position_status || "已记录")} · ${escapeHtml(item.position_issue || item.fallback_position || "位置已记录")}</small>
  </section>
  <section class="formula-preview-block">
    <h4>原始公式截图</h4>
    ${formulaOriginalPreview(originalRef)}
  </section>
  <section class="formula-preview-block">
    <h4>LaTeX 编辑</h4>
    <textarea class="formula-latex-input formula-detail-latex" data-id="${escapeHtml(item.id)}" rows="4">${escapeHtml(latex)}</textarea>
  </section>
  <section class="formula-preview-block">
    <h4>MathType 预览区域</h4>
    <div class="formula-preview-box">${escapeHtml(mathtypePreview)}</div>
  </section>
  <section class="formula-preview-block">
    <h4>MathML / OMML 转换结果</h4>
    <code>${escapeHtml(mathml)}</code>
    <small>${escapeHtml(policy.message || item.format_status || "等待格式化结果")}</small>
  </section>
  <section class="formula-preview-block">
    <h4>格式化参数</h4>
    <div class="formula-param-grid">${formulaParameterRows(item).join("")}</div>
  </section>
  <section class="formula-preview-block">
    <h4>格式化前后对比</h4>
    <small>${escapeHtml(comparison.summary || "暂无格式化对比")}</small>
  </section>
  <div class="row-actions formula-actions">
    <button class="secondary-button formula-confirm" data-report="${escapeHtml(report.id)}" data-id="${escapeHtml(item.id)}">确认</button>
    <button class="secondary-button formula-rerecognize" data-report="${escapeHtml(report.id)}" data-id="${escapeHtml(item.id)}">重新识别</button>
    <button class="secondary-button formula-save" data-report="${escapeHtml(report.id)}" data-id="${escapeHtml(item.id)}">保存 LaTeX</button>
    <button class="mini-button formula-skip" data-report="${escapeHtml(report.id)}" data-id="${escapeHtml(item.id)}">跳过</button>
  </div>`;
}

function formulaOriginalPreview(ref) {
  if (!ref) return `<div class="formula-preview-box">暂无原始截图引用</div>`;
  if (ref.startsWith("/")) {
    return `<div class="formula-original-image"><img src="${escapeHtml(withToken(ref))}" alt="原始公式截图"></div>`;
  }
  return `<div class="formula-preview-box">${escapeHtml(ref)}</div>`;
}

function formulaMathmlFallback(latex) {
  return latex ? `<math><mtext>${escapeHtml(latex)}</mtext></math>` : "等待识别结果";
}

function formulaParameterRows(item) {
  const rows = [
    ["输出格式", item.output_format || state.settings.formulaOutputFormat],
    ["字体", `${item.font || state.settings.formulaFont || "-"} ${item.font_size || state.settings.formulaFontSize || "-"}pt`],
    ["格式化范围", item.format_scope || state.settings.formulaFormatScope],
    ["对齐", item.alignment || state.settings.formulaAlignment],
    ["变量/函数", `${item.variable_style || state.settings.formulaVariableStyle || "-"} / ${item.function_style || state.settings.formulaFunctionStyle || "-"}`],
    ["上下标比例", `${item.script_scale || state.settings.formulaScriptScale || "-"}%`],
    ["编号", item.numbering || state.settings.formulaNumbering],
    ["低置信度策略", item.low_confidence_strategy || state.settings.lowConfidenceFormulaStrategy],
  ];
  return rows.map(([label, value]) => `<span><b>${escapeHtml(label)}</b><small>${escapeHtml(value || "-")}</small></span>`);
}

function renderReportCenter() {
  $("#reportCounter").textContent = `${state.reports.length} 份报告`;
  const tbody = $("#reportTableBody");
  if (!state.reports.length) {
    tbody.innerHTML = `<tr><td colspan="9" class="empty-state">暂无报告</td></tr>`;
    return;
  }
  tbody.innerHTML = state.reports
    .map((report) => {
      const active = report.id === state.activeReportId ? "active-row" : "";
      return `<tr class="${active}">
      <td><strong>${escapeHtml(report.report_type)}</strong><small class="table-note">${escapeHtml(report.id)}</small></td>
      <td>${report.success_count}</td>
      <td>${report.fail_count}</td>
      <td>${report.formula_count}</td>
      <td>${report.macro_count}</td>
      <td>${report.small_image_count}</td>
      <td>${report.omml_dependency_count || 0}</td>
      <td>${formatTime(report.created_at)}</td>
      <td>${downloadButtons(report.id)}<div class="row-actions"><button class="mini-button report-delete danger" data-id="${report.id}" ${permissionButtonAttrs(true, "删除报告", "reports.manage")}>删除</button></div></td>
    </tr>`;
    })
    .join("");
}

function renderLatestReportPanel() {
  const report = latestReport();
  const dateTarget = $("#latestReportDate");
  if (!report) {
    if (dateTarget) dateTarget.textContent = "";
    $("#latestReportPanel").innerHTML = `<div class="empty-mini">暂无报告</div>`;
    return;
  }
  if (dateTarget) dateTarget.textContent = `(${formatDate(report.created_at)})`;
  const qualityIssues = Number(report.quality_issue_count || 0);
  const batchText = report.batch_total_count ? `批量 ${report.batch_success_count || 0}/${report.batch_total_count}` : "单任务";
  $("#latestReportPanel").innerHTML = `<div class="latest-report-line">
    <strong>${escapeHtml(report.report_type)}</strong>
    <span>${escapeHtml(batchText)}</span>
  </div>
  <div class="latest-report-line">
    <small>质量问题 ${qualityIssues}</small>
    <small>失败 ${escapeHtml(report.fail_count || 0)} · ${formatTime(report.created_at)}</small>
  </div>
  <div class="row-actions latest-report-actions">${downloadButtons(report.id)}</div>`;
}

function renderComparisonPanel() {
  const target = $("#comparisonPanel");
  if (!target) return;
  const comparison = state.previewComparison;
  if (!comparison) {
    target.innerHTML = `<div class="empty-mini">选择报告后点击“对比”</div>`;
    return;
  }
  if (!comparison.items?.length) {
    target.innerHTML = `<div class="empty-mini">当前报告暂无可对比的转换输出</div>`;
    return;
  }
  target.innerHTML = comparison.items
    .map((item) => `<article class="comparison-card">
      <div class="object-title">
        <strong>${escapeHtml(item.source_file)} -> ${escapeHtml(item.output_file)}</strong>
        <span class="badge blue">${escapeHtml(item.output_type)}</span>
      </div>
      <div class="comparison-columns">
        ${previewMiniBlock("转换前", item.source_preview)}
        ${previewMiniBlock("转换后", item.output_preview)}
      </div>
      <div class="notice-list ${item.diff?.warnings?.length ? "" : "good"}">
        <span>页段变化 ${escapeHtml(item.diff?.page_delta ?? 0)}</span>
        <span>对象变化 ${escapeHtml(item.diff?.object_delta ?? 0)}</span>
        <span>告警 ${escapeHtml(item.diff?.warnings?.length || 0)}</span>
      </div>
      ${comparisonObjectChanges(item.diff?.object_changes || [])}
    </article>`)
    .join("");
}

function previewMiniBlock(label, preview) {
  const pages = preview?.pages || [];
  const firstPage = pages[0] || {};
  const objects = preview?.objects || [];
  return `<section class="preview-mini">
    <strong>${label}</strong>
    <small>${escapeHtml(preview?.file_type || "-")} · ${escapeHtml(preview?.status || "-")} · ${pages.length} 页段</small>
    <p>${escapeHtml(firstPage.text || firstPage.title || "暂无预览文本")}</p>
    <div class="summary-tags">${objects.slice(0, 4).map((item) => `<span>${escapeHtml(item.label)} ${escapeHtml(item.count || 0)}</span>`).join("")}</div>
  </section>`;
}

function comparisonObjectChanges(changes) {
  if (!changes.length) return "";
  return `<div class="comparison-changes">${changes
    .slice(0, 6)
    .map((item) => `<span>${escapeHtml(item.label)}: ${escapeHtml(item.source)} -> ${escapeHtml(item.output)} (${escapeHtml(item.delta)})</span>`)
    .join("")}</div>`;
}

function renderLayoutReview() {
  const target = $("#layoutAnnotationList");
  if (!target) return;
  const report = latestReport();
  if ($("#layoutReviewScope")) $("#layoutReviewScope").textContent = report ? `${report.report_type} · ${report.id}` : "暂无报告";
  if (!report) {
    target.innerHTML = `<div class="object-card"><small>暂无可校正报告</small></div>`;
    return;
  }
  const annotations = state.layoutAnnotations.filter((item) => item.report_id === report.id);
  target.innerHTML = annotations.length
    ? annotations
        .map((item) => `<article class="object-card">
          <div class="object-title">
            <strong>${escapeHtml(item.issue_type)} · ${escapeHtml(item.location)}</strong>
            <span class="badge ${layoutStatusClass(item.status)}">${escapeHtml(item.status)}</span>
          </div>
          <small>页码：${escapeHtml(item.page_index || "-")}</small>
          <p>${escapeHtml(layoutChangeSummary(item))}</p>
          ${item.recommendation ? `<small>建议：${escapeHtml(item.recommendation)}</small>` : ""}
          ${item.note ? `<small>备注：${escapeHtml(item.note)}</small>` : ""}
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无排版校正记录</small></div>`;
}

function renderLogs() {
  renderLogTarget($("#workspaceLogList"), state.logs, "暂无日志");
  const task = state.tasks.find((item) => item.id === state.activeLogTaskId);
  const scopedLogs = task ? state.taskLogs : state.logs;
  const scopeLabel = task ? `${task.task_label || taskLabels[task.task_type] || task.task_type} · ${task.id}` : "全部任务";
  if ($("#logScope")) $("#logScope").textContent = scopeLabel;
  renderLogTarget($("#logList"), scopedLogs, task ? "该任务暂无日志" : "暂无日志");
}

function renderLogTarget(target, logs, emptyText) {
  if (!target) return;
  if (!logs.length) {
    target.innerHTML = `<div class="log-item"><small>${escapeHtml(emptyText)}</small></div>`;
    return;
  }
  target.innerHTML = logs
    .slice(0, 36)
    .map((log) => {
      const levelClass = logLevelClass(log.level);
      return `<div class="log-item structured-log-item ${levelClass}">
      <span class="log-line-time">${formatTime(log.created_at)}</span>
      <span class="log-level-chip ${levelClass}">${escapeHtml(logLevelLabel(levelClass))}</span>
      <span class="log-line-message">${escapeHtml(log.message)}</span>
      <small>${escapeHtml(log.category || "system")}</small>
    </div>`;
    })
    .join("");
}

function logLevelClass(level) {
  const normalized = String(level || "info").toLowerCase();
  if (["error", "warning", "warn", "success", "done", "completed", "info"].includes(normalized)) {
    if (normalized === "warn") return "warning";
    if (["done", "completed"].includes(normalized)) return "success";
    return normalized;
  }
  return "info";
}

function logLevelLabel(level) {
  return {
    error: "错误",
    warning: "警告",
    success: "完成",
    info: "信息",
  }[level] || "信息";
}

function renderWorkspaceProgress() {
  const target = $("#workspaceProgress");
  if (!target) return;
  const tasks = state.tasks;
  const latest = tasks[0];
  const total = tasks.length || state.files.length || 0;
  const completed = tasks.filter((task) => task.status === "成功").length;
  const progress = latest ? Number(latest.progress || 0) : total ? Math.round((completed / total) * 100) : 0;
  target.innerHTML = `<div class="progress-head">
    <span>总体进度</span>
    <strong>${completed} / ${total || 0}</strong>
  </div>
  <div class="progress"><span style="width:${Math.max(0, Math.min(100, progress))}%"></span></div>
  <div class="progress-foot">
    <span>${latest ? escapeHtml(latest.task_label || taskLabels[latest.task_type] || latest.task_type) : "暂无任务"}</span>
    <strong>${progress}%</strong>
  </div>`;
}

function renderWorkspaceRecentTasks() {
  const target = $("#workspaceRecentTasks");
  if (!target) return;
  const tasks = state.tasks.slice(0, 5);
  target.innerHTML = `<div class="recent-task-head">
    <strong>最近任务</strong>
    <button class="mini-button" data-view-shortcut="tasks" type="button">任务中心</button>
  </div>${
    tasks.length
      ? tasks
          .map((task) => {
            const progress = Math.max(0, Math.min(100, Number(task.progress || 0)));
            return `<article class="recent-task-item">
              <div>
                <strong>${escapeHtml(task.task_label || taskLabels[task.task_type] || task.task_type)}</strong>
                <small>${escapeHtml(formatDuration(task))} · ${formatTime(task.created_at || task.start_time)}</small>
              </div>
              <span class="badge ${statusClass(task.status)}">${escapeHtml(task.status)}</span>
              <div class="progress"><span style="width:${progress}%"></span></div>
            </article>`;
          })
          .join("")
      : `<div class="empty-mini">暂无任务</div>`
  }`;
}

function renderCompletionNotices() {
  const target = $("#completionNoticePanel");
  if (!target) return;
  const notices = state.tasks.filter((task) => task.completion_notice?.status === "待前端提示");
  if (!notices.length) {
    target.innerHTML = "";
    return;
  }
  target.innerHTML = notices
    .slice(0, 4)
    .map((task) => {
      const notice = task.completion_notice || {};
      return `<div class="task-notice-item ${notice.level === "success" ? "good" : "warn"}">
        <div>
          <strong>${escapeHtml(notice.message || "任务已完成")}</strong>
          <small>${escapeHtml(task.id)} · ${formatTime(notice.created_at)}</small>
        </div>
        <div class="row-actions">
          <button class="mini-button task-detail" data-id="${task.id}" type="button">详情</button>
          <button class="mini-button task-notice-ack" data-id="${task.id}" type="button">知道了</button>
        </div>
      </div>`;
    })
    .join("");
}

function renderCapabilities() {
  const cap = state.capabilities;
  if (!cap) return;
  const profile = state.installProfile || cap.installProfile || {};
  const rows = [
    ["安装画像", { available: true, status: `${profile.platform || cap.platform || "自动检测"} · ${profile.recommendedInstaller || "按平台选择客户端"}` }],
    ["Office 自动化", cap.officeAutomation],
    ["MathType", cap.mathType],
    ["OMML 检索", cap.ommlSearch],
    ["宏执行", cap.macroExecution],
  ];
  $("#capabilityList").innerHTML = rows
    .map(([label, item]) => `<div class="capability-item">
      <strong>${label} ${item.available ? '<span class="badge good">可用</span>' : '<span class="badge warn">待接入</span>'}</strong>
      <small>${escapeHtml(item.status)}</small>
    </div>`)
    .join("");
}

function renderInstallProfile() {
  const target = $("#installProfilePanel");
  const profile = state.installProfile || state.capabilities?.installProfile;
  if (!target || !profile) return;
  const plan = state.installPlan || {};
  const packageInfo = plan.package || {};
  const readiness = plan.readiness || [];
  const steps = plan.steps || [];
  const formulaContract = plan.formula_compatibility || {};
  const contractBlockers = formulaContract.native_handoff_blocking_reasons || [];
  const manifestInstaller = state.localClientManifest?.platform?.installer || {};
  const manifestInstallerBlockers = manifestInstaller.native_handoff_blocking_reasons || [];
  const warnings = plan.warnings || [];
  const routes = Object.entries(profile.taskRouting || {}).slice(0, 6);
  target.innerHTML = `<div class="profile-summary">
    <div><span>平台</span><strong>${escapeHtml(profile.platform)}</strong></div>
    <div><span>安装包</span><strong>${escapeHtml(profile.recommendedInstaller)}</strong></div>
    <div><span>公式对象</span><strong>${escapeHtml(profile.mathtypeObjectFormat)}</strong></div>
    <div><span>兼容策略</span><strong>${escapeHtml(compatibilityLabel(profile.selectedCompatibilityMode))}</strong></div>
  </div>
  <p class="profile-warning">${escapeHtml(profile.formulaPortability)}</p>
  <section class="install-plan-card">
    <div class="object-title">
      <strong>安装计划</strong>
      <span class="badge ${packageInfo.download_url ? "good" : "blue"}">${escapeHtml(packageInfo.status || "待生成")}</span>
    </div>
    <div class="install-package-grid">
      <div><span>文件名</span><strong>${escapeHtml(packageInfo.file_name || "选择平台后生成")}</strong></div>
      <div><span>格式</span><strong>${escapeHtml(plan.installer_kind || profile.installerKind || "-")}</strong></div>
      <div><span>大小</span><strong>${packageInfo.size ? formatBytes(packageInfo.size) : "待打包"}</strong></div>
      <div><span>SHA256</span><strong>${escapeHtml(packageInfo.sha256 || "打包后生成")}</strong></div>
    </div>
    <div class="install-plan-actions">
      ${
        packageInfo.download_url
          ? `<a class="secondary-button" href="${escapeHtml(withToken(packageInfo.download_url))}">${escapeHtml(plan.download_label || profile.downloadLabel || "下载安装包")}</a>`
          : `<button class="secondary-button" type="button" disabled>${escapeHtml(plan.download_label || profile.downloadLabel || "下载安装包")}</button>`
      }
      <span>${escapeHtml(packageInfo.expected_location || "安装包尚未放入本地 installers 目录")}</span>
    </div>
  </section>
  <section class="install-contract-card">
    <div class="object-title">
      <strong>公式交付合同</strong>
      <span class="badge ${installContractBadgeClass(formulaContract)}">${escapeHtml(formulaContract.status || "跨平台兜底")}</span>
    </div>
    <div class="install-contract-grid">
      <div><span>目标平台</span><strong>${escapeHtml(formulaContract.target_platform || profile.platform || "-")}</strong></div>
      <div><span>心跳平台</span><strong>${escapeHtml(formulaContract.heartbeat_platform || "未连接")}</strong></div>
      <div><span>对象格式</span><strong>${escapeHtml(formulaContract.native_object_format || profile.mathtypeObjectFormat || "-")}</strong></div>
      <div><span>兜底格式</span><strong>${escapeHtml((formulaContract.fallback_formats || ["MathML", "LaTeX", "图片"]).join(" / "))}</strong></div>
    </div>
    <div class="install-boundary-grid">
      <div>
        <span>同平台要求</span>
        <strong>${formulaContract.same_platform_required_for_native_objects ? "必须同平台" : "使用兜底格式"}</strong>
      </div>
      <div>
        <span>原生交接</span>
        <strong>${formulaContract.native_handoff_allowed ? "允许" : "阻止"}</strong>
      </div>
      <div>
        <span>阻断原因</span>
        <strong>${escapeHtml(contractBlockers.length ? contractBlockers.join(" / ") : "无")}</strong>
      </div>
      <div>
        <span>安装包边界</span>
        <strong>${escapeHtml(formulaContract.package_boundary || "Windows .msi 与 macOS .pkg 不能混用")}</strong>
      </div>
    </div>
    <p>${escapeHtml(formulaContract.message || profile.formulaPortability || "")}</p>
    <p class="install-contract-action">${escapeHtml(formulaContract.recommended_action || "跨平台交付时保留 MathML、LaTeX 或图片兜底。")}</p>
  </section>
  <section class="installer-manifest-card">
    <div class="object-title">
      <strong>本地客户端安装清单</strong>
      <span class="badge ${manifestInstaller.download_url ? "good" : "blue"}">${escapeHtml(manifestInstaller.status || packageInfo.status || "待打包")}</span>
    </div>
    <div class="install-contract-grid">
      <div><span>清单版本</span><strong>${escapeHtml(manifestInstaller.schema_version || "k12.localInstallerManifest.v1")}</strong></div>
      <div><span>平台 query</span><strong>${manifestInstaller.download_requires_platform_query ? "必须携带" : "无可下载包"}</strong></div>
      <div><span>校验要求</span><strong>${manifestInstaller.checksum_required ? "必须校验 SHA256" : "等待安装包"}</strong></div>
      <div><span>路径策略</span><strong>${escapeHtml(manifestInstaller.path_policy || "安装包路径不写入 manifest")}</strong></div>
    </div>
    <div class="install-boundary-grid">
      <div><span>清单目标平台</span><strong>${escapeHtml(manifestInstaller.target_platform || manifestInstaller.platform || profile.platform || "-")}</strong></div>
      <div><span>清单心跳平台</span><strong>${escapeHtml(manifestInstaller.heartbeat_platform || "未连接")}</strong></div>
      <div><span>清单原生交接</span><strong>${manifestInstaller.native_handoff_allowed ? "允许" : "阻止"}</strong></div>
      <div><span>清单兜底格式</span><strong>${escapeHtml((manifestInstaller.fallback_formats || ["MathML", "LaTeX", "图片"]).join(" / "))}</strong></div>
    </div>
    <p class="install-contract-action">${escapeHtml(manifestInstallerBlockers.length ? `阻断原因：${manifestInstallerBlockers.join(" / ")}` : "manifest 已声明同平台交接边界；跨平台仍需保留兜底格式。")}</p>
    <p>${escapeHtml(manifestInstaller.formula_object_boundary || formulaContract.package_boundary || "Windows .msi 与 macOS .pkg、MathType 原生对象不能跨平台混用。")}</p>
  </section>
  <div class="install-warning-list">
    ${warnings.map((warning) => `<small>${escapeHtml(warning)}</small>`).join("")}
  </div>
  <div class="install-readiness-list">
    ${readiness
      .map((item) => `<article>
        <strong>${escapeHtml(item.title)}</strong>
        <span class="badge ${escapeHtml(item.level)}">${escapeHtml(item.status)}</span>
        <small>${escapeHtml(item.detail)}</small>
      </article>`)
      .join("")}
  </div>
  <div class="install-step-list">
    ${steps
      .map((step) => `<div>
        <b>${escapeHtml(step.order)}.</b>
        <strong>${escapeHtml(step.title)}</strong>
        <small>${escapeHtml(step.detail)}</small>
      </div>`)
      .join("")}
  </div>
  <div class="profile-columns">
    <section>
      <h3>必装组件</h3>
      ${profileList(profile.requiredComponents)}
    </section>
    <section>
      <h3>可选组件</h3>
      ${profileList(profile.optionalComponents)}
    </section>
    <section>
      <h3>安装前检查</h3>
      ${profileList(profile.preflightChecks)}
    </section>
  </div>
  <div class="profile-routes">
    <h3>任务分流</h3>
    ${routes.map(([task, note]) => `<div><strong>${escapeHtml(taskLabels[task] || task)}</strong><span>${escapeHtml(note)}</span></div>`).join("")}
  </div>`;
}

function renderLocalHandoff() {
  const target = $("#localHandoffPanel");
  if (!target) return;
  const tasks = localTaskCandidates();
  const task = activeLocalHandoffTask(tasks);
  const tokenReady = hasLocalSecurityToken();
  const credentialReady = hasLocalSecurityCredential();
  const clientEnabled = Boolean(state.settings.localClientEnabled ?? true);
  const launchAllowed = Boolean(state.settings.allowWebLaunchLocalClient);
  const syncAllowed = Boolean(state.settings.allowTaskStatusCloudSync);
  const cloudAllowed = Boolean(state.settings.allowCloudSync);
  const platform = state.installProfile?.platform || state.settings.localClientPlatform || "auto";
  const manifest = state.localClientManifest || {};
  const heartbeat = manifest.heartbeat || {};
  const queue = manifest.queue || {};
  const uploadQueue = state.localUploadQueue || {};
  const uploadManifest = state.localUploadManifest || null;
  const uploadSummary = uploadQueue.summary || {};
  const uploadItems = uploadQueue.items || [];
  const launch = manifest.launch || {};
  const security = manifest.security || {};
  const heartbeatReady = Boolean(heartbeat.received_at);
  $("#localHandoffScope").textContent = task ? `${task.task_label || taskLabels[task.task_type] || task.task_type} · ${task.status}` : "暂无本地任务";
  const statusCards = [
    ["本地客户端", clientEnabled ? "已启用" : "已关闭", clientEnabled ? "可接收本地任务参数" : "先启用本地客户端连接", clientEnabled ? "good" : "warn"],
    ["网页唤起", launchAllowed ? "已授权" : "需授权", launchAllowed ? "可发送 k12-local 唤起请求" : "开启后才显示启动客户端动作", launchAllowed ? "good" : "warn"],
    ["安全令牌", tokenReady ? "已配置" : "需配置", tokenReady ? "载荷接口受令牌保护" : "本地任务载荷包含本地路径，必须先配置令牌", tokenReady ? "good" : "bad"],
    ["状态同步", syncAllowed ? "已授权" : "未授权", syncAllowed ? "本地客户端可回传进度与输出摘要" : "仅网页端保留本地队列状态", syncAllowed ? "good" : "warn"],
    ["结果上传", cloudAllowed ? "允许" : "本地保存", cloudAllowed ? "本地结果可选择上传云端" : "默认仅保存在本机输出目录", cloudAllowed ? "blue" : "warn"],
    ["客户端心跳", heartbeatReady ? heartbeat.status || "在线" : "未连接", heartbeatReady ? `${heartbeat.platform || platform} · ${formatTime(heartbeat.received_at)}` : "等待本地客户端回传心跳", heartbeatReady ? "good" : "blue"],
  ];
  const endpoint = task ? absoluteUrl(withToken(`/api/tasks/${encodeURIComponent(task.id)}/local-payload`)) : "";
  const readinessPayload = task ? state.localReadiness[task.id] : null;
  const launchDisabled = !task || !clientEnabled || !launchAllowed || !credentialReady;
  const payloadDisabled = !task || !credentialReady;
  target.innerHTML = `<div class="local-handoff-status">
    ${statusCards
      .map(([label, value, note, badge]) => `<article>
        <strong>${escapeHtml(label)}</strong>
        <span class="badge ${badge}">${escapeHtml(value)}</span>
        <small>${escapeHtml(note)}</small>
      </article>`)
      .join("")}
  </div>
  <article class="local-client-manifest">
    <div class="object-title">
      <strong>本地客户端运行清单</strong>
      <span class="badge ${security.token_configured ? "good" : "warn"}">${security.token_configured ? "令牌已配置" : "需令牌"}</span>
    </div>
    <dl class="compact-detail-list">
      <div><dt>服务地址</dt><dd>${escapeHtml(manifest.service?.origin || "等待设置")}</dd></div>
      <div><dt>唤起协议</dt><dd>${escapeHtml(launch.url_template || "k12-local://open")}</dd></div>
      <div><dt>心跳接口</dt><dd>${escapeHtml(absoluteUrl("/api/local-client/heartbeat"))}</dd></div>
      <div><dt>待本地任务</dt><dd>${escapeHtml(String(queue.pending_local_task_count || 0))} / ${escapeHtml(String(queue.local_task_count || 0))}</dd></div>
    </dl>
    <div class="local-client-actions">
      ${(manifest.supported_actions || [])
        .map((item) => `<span>${escapeHtml(item.label)}</span>`)
        .join("")}
    </div>
  </article>
  <article class="local-upload-queue">
    <div class="object-title">
      <strong>本地结果上传队列</strong>
      <span class="badge ${uploadQueue.cloud_sync_allowed ? "blue" : "warn"}">${uploadQueue.cloud_sync_allowed ? "云端同步已授权" : "默认本地保存"}</span>
    </div>
    <div class="local-upload-summary">
      <div><span>总数</span><strong>${escapeHtml(String(uploadSummary.total || 0))}</strong></div>
      <div><span>等待上传</span><strong>${escapeHtml(String(uploadSummary.queued || 0))}</strong></div>
      <div><span>已登记</span><strong>${escapeHtml(String(uploadSummary.registered || 0))}</strong></div>
      <div><span>未授权</span><strong>${escapeHtml(String(uploadSummary.not_authorized || 0))}</strong></div>
      <div><span>输出</span><strong>${escapeHtml(String(uploadSummary.output_count || 0))}</strong></div>
    </div>
    <div class="local-upload-list">
      ${
        uploadItems.length
          ? uploadItems
              .slice(0, 4)
              .map((item) => `<div>
                <strong>${escapeHtml(item.task_label || item.task_id)}</strong>
                <span class="badge ${uploadBadge(item.upload_status)}">${escapeHtml(item.upload_status_label || item.upload_status)}</span>
                ${
                  item.manifest_endpoint
                    ? `<button class="mini-button upload-manifest" data-endpoint="${escapeHtml(item.manifest_endpoint)}" type="button">接收清单</button>`
                    : ""
                }
                <small>${escapeHtml(item.task_id)} · ${escapeHtml(item.synced_at ? formatTime(item.synced_at) : "尚未同步")} · 输出 ${escapeHtml(String(item.output_count || 0))} 个</small>
              </div>`)
              .join("")
          : `<div class="empty-mini">本地客户端回传输出后，会在这里显示本地保存或上传云端的处理状态。</div>`
      }
    </div>
    ${renderLocalUploadManifest(uploadManifest)}
    <small>${escapeHtml(uploadQueue.policy || "本地结果默认保存在本机输出目录。")}</small>
  </article>
  ${renderLocalReadinessCard(readinessPayload, task)}
  <article class="local-handoff-task">
    ${
      task
        ? `<div>
            <strong>${escapeHtml(task.task_label || taskLabels[task.task_type] || task.task_type)}</strong>
            <small>${escapeHtml(task.id)} · ${escapeHtml(modeLabels[task.execute_mode] || task.execute_mode || "自动")} · ${escapeHtml(platform)}</small>
          </div>
          <div class="local-handoff-progress">
            <span style="width:${Math.max(0, Math.min(100, Number(task.progress || 0)))}%"></span>
          </div>
          <dl class="compact-detail-list">
            <div><dt>载荷地址</dt><dd>${payloadDisabled ? "配置本地安全令牌后可用" : escapeHtml(endpoint)}</dd></div>
            <div><dt>本地状态</dt><dd>${escapeHtml(task.local_client_status || task.status || "待处理")}</dd></div>
            <div><dt>同步说明</dt><dd>${escapeHtml(task.local_client_message || task.error_message || "等待本地客户端接收任务")}</dd></div>
            <div><dt>启动请求</dt><dd>${escapeHtml(task.local_launch_request?.id ? `${task.local_launch_request.id} · ${formatTime(task.local_launch_request.created_at)}` : "尚未请求")}</dd></div>
            <div><dt>最近同步</dt><dd>${escapeHtml(task.local_synced_at ? formatTime(task.local_synced_at) : "尚未同步")}</dd></div>
          </dl>
          <div class="row-actions">
            <button class="secondary-button launch-local-client" data-id="${escapeHtml(task.id)}" ${launchDisabled ? "disabled" : ""} type="button">启动本地客户端</button>
            <button class="secondary-button copy-local-payload" data-id="${escapeHtml(task.id)}" ${payloadDisabled ? "disabled" : ""} type="button">复制载荷地址</button>
            <button class="mini-button task-detail" data-id="${escapeHtml(task.id)}" type="button">任务详情</button>
          </div>`
        : `<div class="empty-mini">创建 Word 宏、OMML、MathType 格式化或其他本地优先任务后，会在这里显示本地客户端交接状态。</div>`
    }
  </article>`;
}

function localTaskCandidates() {
  return state.tasks.filter((task) => taskNeedsLocalClient(task)).slice(0, 8);
}

function activeLocalHandoffTask(tasks = localTaskCandidates()) {
  return tasks.find((item) => ["待处理", "处理中", "待确认", "待本地客户端执行", "已暂停"].includes(item.status)) || tasks[0] || null;
}

async function refreshLocalReadiness(taskId = "") {
  const task = taskId ? state.tasks.find((item) => item.id === taskId) : activeLocalHandoffTask();
  state.localReadiness = {};
  if (!task) return;
  try {
    const result = await api(`/api/tasks/${encodeURIComponent(task.id)}/local-readiness`);
    state.localReadiness[task.id] = result.localReadiness;
  } catch (error) {
    state.localReadiness[task.id] = {
      error: error.message || "客户端就绪状态读取失败",
    };
  }
}

function renderLocalReadinessCard(payload, task) {
  if (!task) {
    return `<article class="local-readiness-card">
      <div class="object-title">
        <strong>客户端就绪</strong>
        <span class="badge blue">等待任务</span>
      </div>
      <div class="empty-mini">暂无需要本地客户端接手的任务。</div>
    </article>`;
  }
  if (payload?.error) {
    return `<article class="local-readiness-card">
      <div class="object-title">
        <strong>客户端就绪</strong>
        <span class="badge warn">读取失败</span>
      </div>
      <small>${escapeHtml(payload.error)}</small>
    </article>`;
  }
  const readiness = payload?.client_readiness || {};
  const required = readiness.required_capabilities || [];
  const missing = new Set(readiness.missing_capabilities || []);
  const actions = payload?.local_actions || [];
  const componentCount = Object.values(readiness.components || {}).filter((item) => item?.available).length;
  return `<article class="local-readiness-card">
    <div class="object-title">
      <strong>客户端就绪</strong>
      <span class="badge ${clientReadinessBadge(readiness.status)}">${escapeHtml(clientReadinessLabel(readiness.status))}</span>
    </div>
    <small>${escapeHtml(readiness.message || "等待本地客户端预检结果")}</small>
    <div class="local-readiness-capabilities">
      ${
        required.length
          ? required
              .map((item) => `<span class="${missing.has(item.key) ? "missing" : "ready"}">${escapeHtml(item.label)} · ${missing.has(item.key) ? "缺失" : "可用"}</span>`)
              .join("")
          : `<span class="ready">当前动作无额外桌面能力要求</span>`
      }
    </div>
    <dl class="compact-detail-list">
      <div><dt>平台</dt><dd>${escapeHtml(readiness.expected_platform && readiness.expected_platform !== "auto" ? `${readiness.platform || "Unknown"} / 目标 ${readiness.expected_platform}` : readiness.platform || "Unknown")}</dd></div>
      <div><dt>心跳</dt><dd>${escapeHtml(readiness.last_heartbeat_at ? formatTime(readiness.last_heartbeat_at) : "未收到")}</dd></div>
      <div><dt>可见组件</dt><dd>${escapeHtml(String(componentCount))}</dd></div>
      <div><dt>动作数</dt><dd>${escapeHtml(String(actions.length))}</dd></div>
    </dl>
    <div class="local-readiness-actions">
      ${
        actions.length
          ? actions.map((action) => `<span>${escapeHtml(action.label || action.type)} · ${escapeHtml(action.status || "待处理")}</span>`).join("")
          : `<span>暂无本地动作队列</span>`
      }
    </div>
    ${renderLocalExecutionPlan(payload?.desktop_execution_plan)}
  </article>`;
}

function renderLocalExecutionPlan(plan) {
  if (!plan || !plan.schema_version) {
    return `<section class="local-execution-plan">
      <div class="object-title">
        <strong>桌面执行计划</strong>
        <span class="badge blue">等待载荷</span>
      </div>
      <small>读取本地就绪状态后，会展示 Office、MathType、OMML 和宏动作的交接计划。</small>
    </section>`;
  }
  const actions = plan.actions || [];
  const platform = plan.platform || {};
  return `<section class="local-execution-plan">
    <div class="object-title">
      <strong>桌面执行计划</strong>
      <span class="badge ${executionPlanBadge(plan.status)}">${escapeHtml(executionPlanLabel(plan.status))}</span>
    </div>
    <dl class="compact-detail-list">
      <div><dt>原生动作</dt><dd>${escapeHtml(String(plan.native_action_count || 0))}</dd></div>
      <div><dt>网页执行</dt><dd>${plan.web_executes_native_documents ? "会执行" : "不执行"}</dd></div>
      <div><dt>平台</dt><dd>${escapeHtml(platform.actual || "Unknown")} / ${escapeHtml(platform.expected || "auto")}</dd></div>
      <div><dt>对象边界</dt><dd>${platform.same_platform_required_for_native_mathtype ? "需同平台" : "兜底格式"}</dd></div>
    </dl>
    <div class="local-execution-actions">
      ${
        actions.length
          ? actions
              .map((action) => `<span>${escapeHtml(action.label || action.type)} · ${escapeHtml(executionPlanLabel(action.gate_status || action.status))} · ${escapeHtml(String((action.steps || []).length))} 步</span>`)
              .join("")
          : `<span>暂无桌面动作</span>`
      }
    </div>
    <small>${escapeHtml(plan.path_policy || "就绪摘要不返回本地路径或令牌。")}</small>
  </section>`;
}

function clientReadinessBadge(status) {
  if (status === "ready_for_handoff" || status === "not_required") return "good";
  if (status === "missing_capability" || status === "needs_heartbeat" || status === "client_disabled" || status === "platform_mismatch") return "warn";
  return "blue";
}

function clientReadinessLabel(status) {
  const labels = {
    ready_for_handoff: "可交接",
    missing_capability: "缺少能力",
    platform_mismatch: "平台不符",
    needs_heartbeat: "待心跳",
    not_required: "无需接手",
    no_executable_action: "无动作",
  };
  return labels[status] || "待预检";
}

function executionPlanBadge(status) {
  if (status === "ready_for_native_client" || status === "ready" || status === "not_required") return "good";
  if (String(status || "").startsWith("blocked") || status === "waiting_for_heartbeat") return "warn";
  return "blue";
}

function executionPlanLabel(status) {
  const labels = {
    ready_for_native_client: "可交接执行",
    waiting_for_heartbeat: "待心跳",
    blocked_by_capability: "能力阻塞",
    blocked_by_platform: "平台阻塞",
    no_executable_action: "无桌面动作",
    not_required: "无需本地",
    ready: "可执行",
    pending: "待预检",
    pending_preflight: "待预检",
  };
  return labels[status] || status || "待预检";
}

function taskNeedsLocalClient(task) {
  if (!task) return false;
  if (task.execute_mode === "local" || task.local_client_status || task.local_sync) return true;
  return ["formula_precheck", "omml_to_mathtype", "mathtype_format", "macro_sequence"].includes(task.task_type);
}

function uploadBadge(status) {
  if (status === "queued") return "good";
  if (status === "not_authorized") return "warn";
  return "blue";
}

async function loadLocalUploadManifest(endpoint) {
  if (!endpoint) return;
  const result = await api(endpoint);
  state.localUploadManifest = result.uploadManifest;
  renderLocalHandoff();
  toast("已加载云端接收清单");
}

function renderLocalUploadManifest(manifest) {
  if (!manifest) {
    return `<div class="local-upload-manifest empty-mini">选择“接收清单”后显示 upload_id、包哈希、输出校验和云端接收合同。</div>`;
  }
  const pkg = manifest.package || {};
  const integrity = manifest.integrity || {};
  const contract = manifest.receive_contract || {};
  const outputs = manifest.outputs || [];
  return `<div class="local-upload-manifest">
    <div class="object-title">
      <strong>云端接收清单</strong>
      <span class="badge ${manifest.receive_state === "ready_for_cloud_receiver" ? "good" : "warn"}">${escapeHtml(manifest.receive_state || manifest.status || "未确认")}</span>
    </div>
    <dl class="compact-detail-list">
      <div><dt>upload_id</dt><dd>${escapeHtml(manifest.upload_id || "-")}</dd></div>
      <div><dt>包哈希</dt><dd>${escapeHtml(pkg.package_sha256 || "未提供")}</dd></div>
      <div><dt>输出数量</dt><dd>${escapeHtml(String(pkg.output_count || 0))}</dd></div>
      <div><dt>输出校验</dt><dd>${escapeHtml(String(integrity.output_sha256_count || 0))} / ${escapeHtml(String(pkg.output_count || 0))}</dd></div>
      <div><dt>文件传输</dt><dd>${escapeHtml(contract.content_transfer || "not_included")}</dd></div>
      <div><dt>路径策略</dt><dd>${escapeHtml(contract.path_policy || manifest.path_policy || "本地路径已隐藏")}</dd></div>
    </dl>
    <div class="local-upload-manifest-outputs">
      ${
        outputs.length
          ? outputs
              .slice(0, 4)
              .map((item) => `<span>${escapeHtml(item.name || "output")} · ${escapeHtml(item.output_type || "-")} · ${escapeHtml(String(item.size || 0))} B · ${escapeHtml(item.sha256 ? "有哈希" : "缺少哈希")}</span>`)
              .join("")
          : `<span>暂无输出摘要</span>`
      }
    </div>
  </div>`;
}

function renderRiskRegister() {
  const target = $("#riskRegisterPanel");
  if (!target) return;
  const evaluated = riskControls.map((item) => ({ ...item, ...riskControlState(item.key) }));
  const controlled = evaluated.filter((item) => item.level === "good").length;
  const warnings = evaluated.filter((item) => item.level === "warn").length;
  $("#riskRegisterScope").textContent = `${controlled} 项已控制，${warnings} 项需确认`;
  target.innerHTML = evaluated
    .map((item) => `<article class="risk-card">
      <div class="object-title">
        <strong>${escapeHtml(item.title)}</strong>
        <span class="badge ${item.level}">${escapeHtml(item.status)}</span>
      </div>
      <small>${escapeHtml(item.risk)}</small>
      <p>${escapeHtml(item.solution)}</p>
      <div class="risk-current">${escapeHtml(item.current)}</div>
      <button class="mini-button risk-jump" data-risk-anchor="${escapeHtml(item.anchor)}" type="button">查看设置</button>
    </article>`)
    .join("");
}

function riskControlState(key) {
  const settings = state.settings || {};
  const preflight = state.preflightChecks || [];
  const hasFailedPreflight = preflight.some((item) => item.status === "失败");
  const hasWarningPreflight = preflight.some((item) => item.status === "需确认");
  if (key === "mathtype_compatibility") {
    const mode = settings.mathtypeCompatibilityMode || "platform-specific";
    const platform = settings.localClientPlatform || "auto";
    if (mode === "mathml-latex" || mode === "image-fallback") {
      return { level: "good", status: "已控制", current: `${compatibilityLabel(mode)} 已启用跨系统兜底` };
    }
    return {
      level: platform === "auto" ? "warn" : "warn",
      status: "需确认",
      current: platform === "auto" ? "尚未选择 Windows 或 macOS 平台" : `${platform} 平台专属对象模式`,
    };
  }
  if (key === "omml_dependency") {
    if (settings.autoSearchOmml && settings.allowManualOmml) return { level: "good", status: "已控制", current: "自动检索和手动指定均已开启" };
    if (settings.autoSearchOmml || settings.allowManualOmml) return { level: "warn", status: "需确认", current: "仅开启一种 OMML 依赖补救方式" };
    return { level: "bad", status: "高风险", current: "OMML 自动检索和手动指定均未开启" };
  }
  if (key === "macro_security") {
    if (!settings.enableMacroExecution) return { level: "good", status: "已控制", current: "宏执行队列已关闭，网页端不会执行宏" };
    if (settings.macroBackup && settings.macroWhitelistOnly) return { level: "good", status: "已控制", current: "已启用执行前备份和白名单控制" };
    if (settings.macroBackup) return { level: "warn", status: "需确认", current: "已启用备份，但白名单未强制开启" };
    return { level: "bad", status: "高风险", current: "宏队列开启但缺少执行前备份" };
  }
  if (key === "pdf_formula_accuracy") {
    if (settings.pdfToWordEngine === "Mathpix" && settings.enableFormulaOcr && Number(settings.formulaConfidenceThreshold || 0) >= 1) {
      return { level: "good", status: "已控制", current: "PDF 转 Word 使用 Mathpix，公式 OCR 和低置信度阈值已配置" };
    }
    return { level: "warn", status: "需确认", current: "需要确认 Mathpix 公式 OCR 或低置信度策略" };
  }
  if (key === "conversion_layout") {
    if (settings.generateReport && settings.retainImages && settings.retainTables && settings.retainHeadersFooters) {
      return { level: "good", status: "已控制", current: "转换报告与主要对象保留偏好已开启" };
    }
    return { level: "warn", status: "需确认", current: "建议开启报告和图片/表格/页眉页脚保留偏好" };
  }
  if (key === "small_image_false_positive") {
    if (Number(settings.smallImageMaxArea || 0) > 0 && !settings.includeDuplicateImages) {
      return { level: "good", status: "已控制", current: "小图面积阈值已配置，重复图片默认过滤" };
    }
    return { level: "warn", status: "需确认", current: "建议确认小图阈值、重复图和人工标记策略" };
  }
  if (key === "large_file_queue") {
    if (Number(settings.singleFileLimitMb || 0) > 0 && Number(settings.maxConcurrentTasks || 0) > 0) {
      return { level: "good", status: "已控制", current: `单文件 ${settings.singleFileLimitMb} MB，并发 ${settings.maxConcurrentTasks}` };
    }
    return { level: "warn", status: "需确认", current: "需要配置单文件限制和并发数量" };
  }
  if (key === "local_file_access") {
    if (settings.localClientEnabled && settings.sensitiveFilesPreferLocal) return { level: "good", status: "已控制", current: "本地客户端启用，敏感文档本地优先" };
    if (settings.localClientEnabled) return { level: "warn", status: "需确认", current: "本地客户端已启用，但敏感文档本地优先未开启" };
    return { level: "bad", status: "高风险", current: "本地客户端连接已关闭" };
  }
  if (key === "privacy") {
    if (!settings.allowCloudSync && !settings.allowExternalMathpixUpload) return { level: "good", status: "已控制", current: "云端同步与 Mathpix 外部上传均未授权，默认本地处理" };
    if (settings.allowExternalMathpixUpload) return { level: "warn", status: "需确认", current: "已授权 Mathpix 外部上传，请确认文档可上传" };
    return { level: "warn", status: "需确认", current: "已允许云端同步，请确认数据范围" };
  }
  if (key === "local_communication") {
    const host = String(settings.localApiHost || "127.0.0.1");
    const tokenReady = hasLocalSecurityToken();
    if (host === "127.0.0.1" && tokenReady && !hasFailedPreflight) return { level: "good", status: "已控制", current: "本地 API 限制在 127.0.0.1 且安全令牌已配置" };
    if (hasFailedPreflight || hasWarningPreflight) return { level: "warn", status: "需确认", current: "运行预检存在需处理项，请查看预检列表" };
    return { level: "warn", status: "需确认", current: tokenReady ? `当前监听 ${host}` : "本地安全令牌未配置" };
  }
  return { level: "warn", status: "需确认", current: "等待配置" };
}

function renderPerformanceTargets() {
  const target = $("#performanceTargetPanel");
  if (!target) return;
  const evaluated = performanceTargets.map((item) => ({ ...item, ...performanceTargetState(item.key) }));
  const ready = evaluated.filter((item) => item.level === "good").length;
  const warnings = evaluated.filter((item) => item.level === "warn").length;
  $("#performanceTargetScope").textContent = `${ready} 项达标，${warnings} 项需确认`;
  target.innerHTML = evaluated
    .map((item) => `<article class="performance-card">
      <div class="object-title">
        <strong>${escapeHtml(item.metric)}</strong>
        <span class="badge ${item.level}">${escapeHtml(item.status)}</span>
      </div>
      <small>目标：${escapeHtml(item.target)}</small>
      <p>${escapeHtml(item.route)}</p>
      <div class="performance-current">${escapeHtml(item.current)}</div>
    </article>`)
    .join("");
}

function performanceTargetState(key) {
  const settings = state.settings || {};
  const completedTasks = state.tasks.filter((task) => Number(task.duration_seconds || 0) > 0);
  const recentDuration = completedTasks.length
    ? Math.round(completedTasks.slice(0, 8).reduce((sum, task) => sum + Number(task.duration_seconds || 0), 0) / Math.min(8, completedTasks.length))
    : 0;
  if (key === "drag_response") {
    return { level: "good", status: "已覆盖", current: "文件选择和拖拽导入后刷新队列，目标 3 秒内可感知" };
  }
  if (key === "single_file_size") {
    const limit = Number(settings.singleFileLimitMb || 0);
    return {
      level: limit >= 500 ? "good" : "warn",
      status: limit >= 500 ? "达标" : "需确认",
      current: limit ? `当前单文件限制 ${limit} MB` : "尚未配置单文件限制",
    };
  }
  if (key === "batch_queue") {
    const fileCount = state.files.length;
    return {
      level: fileCount >= 20 || Number(settings.maxConcurrentTasks || 0) > 0 ? "good" : "warn",
      status: "已覆盖",
      current: `当前队列 ${fileCount} 个文件，批量任务支持排队和失败重试`,
    };
  }
  if (key === "concurrency") {
    const count = Number(settings.maxConcurrentTasks || 0);
    const ok = count >= 2 && count <= 4;
    return { level: ok ? "good" : "warn", status: ok ? "达标" : "需确认", current: count ? `当前最大并发 ${count}` : "尚未配置最大并发" };
  }
  if (key === "word_parse") {
    const wordCount = state.files.filter((file) => file.file_type === "Word").length;
    return { level: "good", status: "已覆盖", current: `${wordCount} 个 Word 文件可进入结构预览、公式预检和报告` };
  }
  if (key === "pdf_to_word") {
    const enabledOcr = [settings.enableTextOcr, settings.enableFormulaOcr, settings.enableTableOcr].filter(Boolean).length;
    return {
      level: settings.pdfToWordEngine === "Mathpix" && enabledOcr ? "good" : "warn",
      status: settings.pdfToWordEngine === "Mathpix" ? "已配置" : "需确认",
      current: `引擎 ${settings.pdfToWordEngine || "未配置"}，OCR 开关 ${enabledOcr}/3，轮询超时 ${settings.mathpixPollTimeoutSeconds || 0} 秒`,
    };
  }
  if (key === "small_image_scan") {
    return {
      level: Number(settings.smallImageMaxArea || 0) > 0 ? "good" : "warn",
      status: Number(settings.smallImageMaxArea || 0) > 0 ? "已配置" : "需确认",
      current: `最大面积 ${settings.smallImageMaxArea || 0}，宽 ${settings.smallImageMaxWidth || 0}，高 ${settings.smallImageMaxHeight || 0}`,
    };
  }
  if (key === "omml_search") {
    const maxFiles = Number(settings.ommlSearchMaxFiles || 0);
    return {
      level: settings.autoSearchOmml && maxFiles > 0 ? "good" : "warn",
      status: settings.autoSearchOmml ? "已配置" : "需确认",
      current: `自动检索 ${settings.autoSearchOmml ? "开启" : "关闭"}，扫描文件上限 ${maxFiles}`,
    };
  }
  if (key === "macro_detection") {
    return {
      level: settings.enableMacroDetection ? "good" : "warn",
      status: settings.enableMacroDetection ? "已开启" : "需确认",
      current: `宏检测 ${settings.enableMacroDetection ? "开启" : "关闭"}，宏超时 ${settings.macroTimeoutSeconds || 0} 秒`,
    };
  }
  if (key === "formula_recognition") {
    const threshold = Number(settings.formulaConfidenceThreshold || 0);
    return {
      level: settings.enableFormulaOcr && threshold > 0 ? "good" : "warn",
      status: settings.enableFormulaOcr ? "已配置" : "需确认",
      current: `公式 OCR ${settings.enableFormulaOcr ? "开启" : "关闭"}，置信度阈值 ${threshold}`,
    };
  }
  if (key === "task_recovery") {
    const interrupted = state.tasks.filter((task) => task.status === "已中断").length;
    return {
      level: "good",
      status: "已覆盖",
      current: interrupted ? `${interrupted} 个中断任务可恢复或重试，最近平均耗时 ${recentDuration || 0} 秒` : `当前无中断任务，最近平均耗时 ${recentDuration || 0} 秒`,
    };
  }
  return { level: "warn", status: "需确认", current: "等待性能数据" };
}

function renderExceptionPolicies() {
  const target = $("#exceptionPolicyPanel");
  if (!target) return;
  const evaluated = exceptionPolicies.map((item) => ({ ...item, ...exceptionPolicyState(item.key) }));
  const normal = evaluated.filter((item) => item.level === "good").length;
  const active = evaluated.filter((item) => item.level !== "good").length;
  $("#exceptionPolicyScope").textContent = `${normal} 类正常，${active} 类需关注`;
  target.innerHTML = evaluated
    .map((item) => `<article class="exception-card">
      <div class="object-title">
        <strong>${escapeHtml(item.category)}</strong>
        <span class="badge ${item.level}">${escapeHtml(item.status)}</span>
      </div>
      <small>${escapeHtml(item.exceptions)}</small>
      <p>${escapeHtml(item.handling)}</p>
      <div class="exception-current">${escapeHtml(item.current)}</div>
      ${item.key === "system" ? renderSystemExceptionDetails() : ""}
      <button class="mini-button exception-jump" data-exception-anchor="${escapeHtml(item.anchor)}" type="button">查看配置</button>
    </article>`)
    .join("");
}

function renderSystemExceptionDetails() {
  const details = systemExceptionDetails();
  return `<div class="system-exception-list">
    ${details
      .map((item) => `<div class="system-exception-row">
        <strong>${escapeHtml(item.label)}</strong>
        <span class="badge ${item.level}">${escapeHtml(item.status)}</span>
        <small>${escapeHtml(item.message)}</small>
      </div>`)
      .join("")}
  </div>`;
}

function systemExceptionDetails() {
  const checks = new Map((state.preflightChecks || []).map((item) => [item.id, item]));
  return [
    ["disk_space", "磁盘空间不足", "检查输出目录所在磁盘可用空间和最小空间阈值"],
    ["output_directory_writable", "文件权限不足", "检查输出目录写入权限或切换输出目录"],
    ["user_permissions", "用户权限不足", "检查当前用户是否具备任务创建、宏执行或任务控制权限"],
    ["local_client", "本地客户端异常", "检查本地客户端开关和本地模式任务分流"],
    ["local_client_platform", "客户端平台心跳", "检查 Windows/macOS 客户端平台是否与安装计划一致"],
    ["local_client_components", "Office / MathType 组件", "检查 Office、MathType、OMML 依赖和宏执行组件能力"],
    ["mathpix_authorization", "OCR 引擎异常", "检查 PDF 转 Word 的 Mathpix 授权、OCR 开关和凭证环境变量"],
    ["macro_authorization", "宏授权异常", "检查宏执行队列、来源授权、白名单和备份策略"],
  ].map(([id, label, fallback]) => {
    const check = checks.get(id) || {};
    const status = check.status || "未检查";
    return {
      label,
      status,
      level: exceptionPreflightLevel(status),
      message: check.message || check.recommendation || fallback,
    };
  });
}

function exceptionPreflightLevel(status) {
  if (status === "通过") return "good";
  if (status === "失败") return "bad";
  if (status === "需确认") return "warn";
  return "blue";
}

function exceptionPolicyState(key) {
  const settings = state.settings || {};
  const reports = state.reports || [];
  const failureRows = reports.reduce((sum, report) => sum + Number(report.failure_count || 0), 0);
  const failedTasks = state.tasks.filter((task) => task.status === "失败").length;
  if (key === "file") {
    const invalidFiles = state.files.filter((file) => file.validation_errors?.length || file.status === "校验失败").length;
    const encryptedFiles = state.files.filter((file) => file.encrypted).length;
    const ok = invalidFiles === 0;
    return {
      level: ok ? "good" : "warn",
      status: ok ? "正常" : "需处理",
      current: `校验异常 ${invalidFiles} 个，加密文件 ${encryptedFiles} 个，重复策略 ${settings.duplicateFileStrategy || "未配置"}`,
    };
  }
  if (key === "conversion") {
    const conversionFailures = state.tasks.filter((task) => ["word_to_ppt", "ppt_to_word", "pdf_to_word", "excel_to_pdf", "excel_to_word", "excel_to_ppt"].includes(task.task_type) && task.status === "失败").length;
    return {
      level: conversionFailures || failureRows ? "warn" : "good",
      status: conversionFailures || failureRows ? "需关注" : "正常",
      current: `转换失败任务 ${conversionFailures} 个，失败清单 ${failureRows} 条，报告生成 ${settings.generateReport ? "开启" : "关闭"}`,
    };
  }
  if (key === "formula") {
    const formulaIssues = state.formulaAnnotations.filter((item) => ["重识别", "跳过", "已修正"].includes(item.status)).length;
    const lowConfidence = reports.reduce((sum, report) => sum + Number(report.low_confidence_formula_count || 0), 0);
    return {
      level: formulaIssues || lowConfidence ? "warn" : "good",
      status: formulaIssues || lowConfidence ? "需复核" : "正常",
      current: `人工校正 ${formulaIssues} 条，低置信度 ${lowConfidence} 条，原图保留 ${settings.keepFormulaImages ? "开启" : "关闭"}`,
    };
  }
  if (key === "omml") {
    const ommlIssues = state.ommlAnnotations.filter((item) => ["转换失败", "重新转换", "手动指定依赖"].includes(item.status)).length;
    const ommlFiles = state.files.filter((file) => file.has_omml).length;
    return {
      level: ommlIssues ? "warn" : "good",
      status: ommlIssues ? "需处理" : "正常",
      current: `OMML 文件 ${ommlFiles} 个，校正记录 ${ommlIssues} 条，自动检索 ${settings.autoSearchOmml ? "开启" : "关闭"}，手动选择 ${settings.allowManualOmml ? "开启" : "关闭"}`,
    };
  }
  if (key === "macro") {
    const macroFailures = reports.reduce((sum, report) => sum + Number(report.macro_fail_count || 0), 0);
    const macroFiles = state.files.filter((file) => file.has_macro).length;
    return {
      level: macroFailures ? "warn" : "good",
      status: macroFailures ? "需处理" : "正常",
      current: `含宏文件 ${macroFiles} 个，宏失败 ${macroFailures} 个，失败策略 ${settings.macroFailureStrategy || "未配置"}，备份 ${settings.macroBackup ? "开启" : "关闭"}`,
    };
  }
  if (key === "image") {
    const imageIssues = state.imageAnnotations.filter((item) => ["误判", "删除待处理", "替换待处理"].includes(item.status)).length;
    const imageExtractionErrors = reports.reduce((sum, report) => sum + (report.analysis?.imageExtractionErrors || []).length, 0);
    const smallImages = reports.reduce((sum, report) => sum + Number(report.small_image_count || 0), 0);
    return {
      level: imageIssues || imageExtractionErrors ? "warn" : "good",
      status: imageIssues || imageExtractionErrors ? "需复核" : "正常",
      current: `微小图片 ${smallImages} 个，提取失败 ${imageExtractionErrors} 条，人工标记 ${imageIssues} 条，导出格式 ${settings.imageExportFormat || "原格式"}`,
    };
  }
  if (key === "system") {
    const failedPreflight = state.preflightChecks.filter((item) => item.status === "失败").length;
    const warningPreflight = state.preflightChecks.filter((item) => item.status === "需确认").length;
    const interrupted = state.tasks.filter((task) => task.status === "已中断").length;
    const level = failedPreflight ? "bad" : warningPreflight || interrupted || failedTasks ? "warn" : "good";
    return {
      level,
      status: level === "good" ? "正常" : "需关注",
      current: `预检失败 ${failedPreflight} 项，需确认 ${warningPreflight} 项，中断任务 ${interrupted} 个，失败任务 ${failedTasks} 个`,
    };
  }
  return { level: "warn", status: "需确认", current: "等待异常数据" };
}

function renderAcceptanceOverview() {
  const target = $("#acceptanceOverviewPanel");
  if (!target) return;
  const evaluated = acceptanceGroups.map((group) => ({ ...group, ...acceptanceGroupState(group.key) }));
  const covered = evaluated.filter((group) => group.level === "good").length;
  const contract = evaluated.filter((group) => group.level === "warn").length;
  $("#acceptanceOverviewScope").textContent = `${covered} 组已覆盖，${contract} 组需本地或外部实测`;
  target.innerHTML = evaluated
    .map((group) => `<article class="acceptance-card">
      <div class="object-title">
        <strong>${escapeHtml(group.section)} ${escapeHtml(group.title)}</strong>
        <span class="badge ${group.level}">${escapeHtml(group.status)}</span>
      </div>
      <div class="acceptance-score"><span style="width:${Math.round((group.covered / group.total) * 100)}%"></span></div>
      <small>${escapeHtml(group.covered)} / ${escapeHtml(group.total)} 项 · ${escapeHtml(group.evidence)}</small>
      <p>${escapeHtml(group.current)}</p>
      <button class="mini-button acceptance-jump" data-acceptance-view="${escapeHtml(group.view)}" ${group.anchor ? `data-acceptance-anchor="${escapeHtml(group.anchor)}"` : ""} type="button">查看证据</button>
    </article>`)
    .join("");
}

function renderAcceptanceMatrix() {
  const target = $("#acceptanceMatrixPanel");
  if (!target) return;
  const riskTarget = $("#acceptanceRiskPanel");
  const matrix = state.acceptanceMatrix || {};
  const groups = matrix.groups || [];
  const summary = matrix.summary || {};
  const risks = matrix.uncovered_risks || [];
  $("#acceptanceMatrixScope").textContent = `${summary.covered || 0} / ${summary.items || 0} 项已覆盖，${summary.contract || 0} 项需实测，${risks.length || 0} 条未覆盖风险`;
  if (riskTarget) {
    riskTarget.innerHTML = risks.length
      ? risks
          .map((item) => `<article class="acceptance-risk-card">
            <div class="object-title">
              <strong>${escapeHtml(item.key)} ${escapeHtml(item.requirement)}</strong>
              <span class="badge warn">${escapeHtml(item.status)}</span>
            </div>
            <p>${escapeHtml(item.uncovered_risk || "未覆盖风险待补充")}</p>
            <small>阻断：${escapeHtml((item.blocking_reasons || []).join(" / ") || "待确认")}</small>
            <small>环境：${escapeHtml((item.required_environment || []).join(" / ") || "待确认")}</small>
            ${(item.verification_checklist || []).length ? `<ol class="acceptance-risk-checklist">${item.verification_checklist.map((step) => `<li>${escapeHtml(step)}</li>`).join("")}</ol>` : ""}
          </article>`)
          .join("")
      : "";
  }
  if (!groups.length) {
    target.innerHTML = '<div class="empty-state-mini">暂无验收证据矩阵</div>';
    return;
  }
  target.innerHTML = groups
    .map((group) => {
      const progress = group.total ? Math.round((Number(group.covered || 0) / Number(group.total || 1)) * 100) : 0;
      const rows = (group.items || [])
        .map((item) => `<div class="acceptance-item-row">
          <span class="badge ${escapeHtml(item.level)}">${escapeHtml(item.status)}</span>
          <strong>${escapeHtml(item.key)} ${escapeHtml(item.requirement)}</strong>
          <small>${escapeHtml(item.evidence)}</small>
          ${item.current ? `<em>${escapeHtml(item.current)}</em>` : ""}
          ${item.gap ? `<em>${escapeHtml(item.gap)}</em>` : ""}
          <small>${escapeHtml(item.next_step || "")}</small>
        </div>`)
        .join("");
      return `<article class="acceptance-matrix-card">
        <div class="object-title">
          <strong>${escapeHtml(group.section)} ${escapeHtml(group.title)}</strong>
          <span class="badge ${escapeHtml(group.level)}">${escapeHtml(group.status)}</span>
        </div>
        <div class="acceptance-score"><span style="width:${progress}%"></span></div>
        <small>${escapeHtml(group.covered)} 项已覆盖，${escapeHtml(group.contract)} 项需实测，${escapeHtml(group.blocked)} 项缺口</small>
        <details ${group.level !== "good" ? "open" : ""}>
          <summary>查看 ${escapeHtml(group.total)} 条验收证据</summary>
          <div class="acceptance-item-list">${rows}</div>
        </details>
      </article>`;
    })
    .join("");
}

function acceptanceGroupState(key) {
  const settings = state.settings || {};
  const reports = state.reports || [];
  const tasks = state.tasks || [];
  const files = state.files || [];
  const reportTypes = new Set(reports.map((report) => report.task_type));
  const taskTypes = new Set(tasks.map((task) => task.task_type));
  if (key === "upload") {
    const typeCount = new Set(files.map((file) => file.file_type)).size;
    const hasValidation = files.some((file) => file.validation_errors?.length || file.status === "校验失败");
    return {
      covered: 9,
      level: "good",
      status: "已覆盖",
      current: `当前文件 ${files.length} 个，类型 ${typeCount} 类，校验提示 ${hasValidation ? "已有异常样本" : "等待异常样本"}`,
    };
  }
  if (key === "word") {
    const hasWordOutput = reportTypes.has("word_to_ppt") || taskTypes.has("word_to_ppt");
    return {
      covered: hasWordOutput ? 9 : 8,
      level: "warn",
      status: "合同覆盖",
      current: "已生成最小可用 Word 转 PPT 产物和对象保留清单，高保真 Office 引擎仍需本地实测",
    };
  }
  if (key === "omml") {
    const ommlFiles = files.filter((file) => file.has_omml).length;
    return {
      covered: settings.autoSearchOmml && settings.allowManualOmml ? 9 : 8,
      level: "warn",
      status: "交接覆盖",
      current: `OMML 文件 ${ommlFiles} 个，自动检索 ${settings.autoSearchOmml ? "开启" : "关闭"}，手动选择 ${settings.allowManualOmml ? "开启" : "关闭"}；真实 OMML 转 MathType 需本地客户端实测`,
    };
  }
  if (key === "macro") {
    const macroReports = reports.reduce((sum, report) => sum + Number(report.macro_count || 0), 0);
    return {
      covered: settings.macroBackup ? 9 : 8,
      level: "warn",
      status: "交接覆盖",
      current: `宏记录 ${macroReports} 条，备份 ${settings.macroBackup ? "开启" : "关闭"}，失败策略 ${settings.macroFailureStrategy || "跳过"}；网页端只生成本地队列，真实宏执行需本地客户端实测`,
    };
  }
  if (key === "ppt") {
    const hasPptOutput = reportTypes.has("ppt_to_word") || taskTypes.has("ppt_to_word");
    return {
      covered: hasPptOutput ? 8 : 7,
      level: hasPptOutput ? "good" : "warn",
      status: hasPptOutput ? "已覆盖" : "需样本",
      current: "PPT 缩略结构、内容预览、备注/图片/公式提取和最小 DOCX 输出已接入",
    };
  }
  if (key === "pdf") {
    const pdfReports = reports.filter((report) => report.task_type === "pdf_to_word").length;
    return {
      covered: settings.pdfToWordEngine === "Mathpix" ? 8 : 6,
      level: "warn",
      status: "需 Mathpix 实测",
      current: `PDF 引擎 ${settings.pdfToWordEngine || "未配置"}，Mathpix 报告 ${pdfReports} 份；外部上传需用户授权和凭证`,
    };
  }
  if (key === "mathtype") {
    return {
      covered: settings.enableMathTypeFormatting ? 9 : 7,
      level: "warn",
      status: "需本地客户端实测",
      current: `格式化 ${settings.enableMathTypeFormatting ? "开启" : "关闭"}，兼容模式 ${compatibilityLabel(settings.mathtypeCompatibilityMode)}；真实 MathType 对象写回需本地客户端`,
    };
  }
  if (key === "images") {
    const imageReports = reports.reduce((sum, report) => sum + Number(report.small_image_count || 0), 0);
    return {
      covered: 12,
      level: "good",
      status: "已覆盖",
      current: `图片报告对象 ${imageReports} 个，支持筛选、详情、标注、导出图片包和清单`,
    };
  }
  if (key === "batch") {
    const batchTasks = tasks.filter((task) => task.task_type === "batch_process").length;
    return {
      covered: 8,
      level: "good",
      status: "已覆盖",
      current: `批量任务 ${batchTasks} 个，支持排队、进度、重试、跳过、打包下载和日志导出`,
    };
  }
  if (key === "hybrid") {
    const tokenReady = hasLocalSecurityToken();
    const syncReady = Boolean(settings.allowTaskStatusCloudSync);
    const covered = 7 + (tokenReady ? 1 : 0) + (syncReady ? 1 : 0);
    return {
      covered,
      level: "warn",
      status: "交接覆盖",
      current: `本地客户端 ${settings.localClientEnabled ? "启用" : "关闭"}，令牌 ${tokenReady ? "已配置" : "未配置"}，状态同步 ${syncReady ? "已授权" : "未授权"}；真实 Office/MathType/宏执行需桌面端实测`,
    };
  }
  return { covered: 0, level: "warn", status: "需确认", current: "等待验收证据" };
}

function renderVersionRoadmap() {
  const target = $("#versionRoadmapPanel");
  if (!target) return;
  const evaluated = versionRoadmap.map((item) => ({ ...item, ...versionRoadmapState(item) }));
  const covered = evaluated.filter((item) => item.level === "good").length;
  const contracts = evaluated.filter((item) => item.level === "warn").length;
  const planned = evaluated.filter((item) => item.level === "blue").length;
  $("#versionRoadmapScope").textContent = `${covered} 个版本已覆盖，${contracts} 个需实测，${planned} 个规划中`;
  target.innerHTML = evaluated
    .map((item) => `<article class="version-card">
      <div class="object-title">
        <strong>${escapeHtml(item.version)} ${escapeHtml(item.goal)}</strong>
        <span class="badge ${item.level}">${escapeHtml(item.status)}</span>
      </div>
      <div class="version-progress"><span style="width:${item.progress}%"></span></div>
      <small>${escapeHtml(item.features)}</small>
      <p>${escapeHtml(item.current)}</p>
    </article>`)
    .join("");
}

function versionRoadmapState(item) {
  if (item.planned) {
    return {
      level: "blue",
      status: "规划中",
      progress: 20,
      current: "AI 排版修复、智能模板套用、私有化部署和开放 API 尚未进入当前首版实现范围",
    };
  }
  const states = item.acceptanceKeys.map((key) => ({
    ...acceptanceGroupState(key),
    total: acceptanceGroups.find((group) => group.key === key)?.total || 0,
  }));
  const total = states.reduce((sum, state) => sum + Number(state.total || 0), 0);
  const covered = states.reduce((sum, state) => sum + Number(state.covered || 0), 0);
  const hasContract = states.some((state) => state.level === "warn");
  const progress = total ? Math.max(5, Math.min(100, Math.round((covered / total) * 100))) : 0;
  if (hasContract) {
    return {
      level: "warn",
      status: "需实测",
      progress,
      current: "界面、报告和本地交接合同已覆盖；Office、MathType、宏或 Mathpix 相关能力仍需真实环境验收",
    };
  }
  return {
    level: "good",
    status: "已覆盖",
    progress,
    current: "当前首版界面、任务、报告和导出链路已提供对应能力入口与本地验证覆盖",
  };
}

function renderDataDictionary() {
  const target = $("#dataDictionaryPanel");
  if (!target) return;
  const objects = dataDictionaryObjects();
  const totalFields = objects.reduce((sum, item) => sum + item.fields.length, 0);
  const summary = state.dataDictionary?.summary;
  $("#dataDictionaryScope").textContent = `${summary?.objects || objects.length} 个对象，${summary?.fields || totalFields} 个字段`;
  target.innerHTML = objects
    .map((item) => {
      const count = item.count ?? dataDictionaryCount(item.key);
      const fields = dataDictionaryFields(item);
      const previewFields = fields
        .slice(0, 6)
        .map(([field, type, note]) => `<li><code>${escapeHtml(field)}</code><span>${escapeHtml(type)}</span><small>${escapeHtml(note)}</small></li>`)
        .join("");
      return `<article class="data-card">
        <div class="object-title">
          <strong>${escapeHtml(item.name)}</strong>
          <span class="badge blue">${escapeHtml(count)} 条</span>
        </div>
        <small>${escapeHtml(item.title)} · ${escapeHtml(fields.length)} 个字段</small>
        <p>${escapeHtml(item.privacy)}</p>
        <ul class="data-field-list">${previewFields}</ul>
        <details>
          <summary>查看全部字段</summary>
          <div class="data-field-table">
            ${fields
              .map(([field, type, note]) => `<div><code>${escapeHtml(field)}</code><span>${escapeHtml(type)}</span><small>${escapeHtml(note)}</small></div>`)
              .join("")}
          </div>
        </details>
      </article>`;
    })
    .join("");
}

function dataDictionaryObjects() {
  if (state.dataDictionary?.schema_version === "k12.dataDictionary.v1" && state.dataDictionary.objects?.length) {
    return state.dataDictionary.objects;
  }
  return dataDictionary;
}

function dataDictionaryFields(item) {
  return (item.fields || []).map((field) => {
    if (Array.isArray(field)) return field;
    return [field.name, field.type, field.description];
  });
}

function dataDictionaryCount(key) {
  if (key === "files") return state.files.length;
  if (key === "tasks") return state.tasks.length;
  if (key === "formulas") return state.reports.reduce((sum, report) => sum + Number(report.formula_count || 0), 0);
  if (key === "omml") return state.reports.reduce((sum, report) => sum + Number(report.omml_dependency_count || report.omml_count || 0), 0);
  if (key === "macros") return state.reports.reduce((sum, report) => sum + Number(report.macro_count || 0), 0) || state.macros.length;
  if (key === "images") return state.reports.reduce((sum, report) => sum + Number(report.small_image_count || 0), 0);
  if (key === "reports") return state.reports.length;
  return 0;
}

function renderSecurityPrivacy() {
  const target = $("#securityPrivacyPanel");
  if (!target) return;
  const evaluated = securityPrivacyControls.map((item) => ({ ...item, ...securityPrivacyState(item.key) }));
  const ready = evaluated.filter((item) => item.level === "good").length;
  const warning = evaluated.filter((item) => item.level === "warn").length;
  const planned = evaluated.filter((item) => item.level === "blue").length;
  $("#securityPrivacyScope").textContent = `${ready} 项已控制，${warning} 项需确认，${planned} 项规划中`;
  target.innerHTML = evaluated
    .map((item) => `<article class="security-card">
      <div class="object-title">
        <strong>${escapeHtml(item.title)}</strong>
        <span class="badge ${item.level}">${escapeHtml(item.status)}</span>
      </div>
      <small>${escapeHtml(item.requirement)}</small>
      <div class="security-current">${escapeHtml(item.current)}</div>
      <button class="mini-button security-jump" data-security-anchor="${escapeHtml(item.anchor)}" type="button">查看设置</button>
    </article>`)
    .join("");
}

function securityPrivacyState(key) {
  const settings = state.settings || {};
  const files = state.files || [];
  const tokenReady = hasLocalSecurityToken();
  if (key === "file_security") {
    const encryptedFiles = files.filter((file) => file.encrypted).length;
    const externalBlocked = !settings.allowCloudSync && !settings.allowExternalMathpixUpload;
    const cleanupReady = Boolean(settings.autoCleanTemp || settings.saveHistory === false);
    if (externalBlocked && cleanupReady) {
      return {
        level: "good",
        status: "已控制",
        current: `外部上传未授权，云端同步关闭，加密文件 ${encryptedFiles} 个；历史或缓存已有清理策略`,
      };
    }
    if (!externalBlocked) {
      return {
        level: "warn",
        status: "需确认",
        current: `外部上传 ${settings.allowExternalMathpixUpload ? "已授权" : "未授权"}，云端同步 ${settings.allowCloudSync ? "已开启" : "已关闭"}，需确认敏感文件范围`,
      };
    }
    return {
      level: "warn",
      status: "需清理",
      current: `默认本地处理已满足，加密文件 ${encryptedFiles} 个；建议开启自动清理或关闭历史保存`,
    };
  }
  if (key === "macro_security_prd") {
    if (!settings.enableMacroExecution) {
      return { level: "good", status: "已控制", current: "宏执行队列已关闭，网页端只做检测和报告，不会自动执行宏" };
    }
    if (settings.macroBackup && settings.macroWhitelistOnly) {
      return { level: "good", status: "已控制", current: "宏执行队列开启，执行前备份和白名单控制均已启用" };
    }
    if (settings.macroBackup) {
      return { level: "warn", status: "需确认", current: "宏执行前备份已开启，但白名单未强制开启，需确认宏来源授权" };
    }
    return { level: "bad", status: "高风险", current: "宏执行队列开启但缺少执行前备份，建议先关闭执行或开启备份" };
  }
  if (key === "local_web_security") {
    const host = String(settings.localApiHost || "127.0.0.1").trim();
    const pathHidden = !settings.exposeLocalPaths;
    if (host === "127.0.0.1" && tokenReady && pathHidden) {
      return { level: "good", status: "已控制", current: "本地 API 限制在 127.0.0.1，安全令牌已配置，本地路径默认隐藏" };
    }
    if (host !== "127.0.0.1" && !tokenReady) {
      return { level: "bad", status: "高风险", current: `当前监听 ${host || "未配置"} 且缺少安全令牌，需要先收紧通信设置` };
    }
    return {
      level: "warn",
      status: "需确认",
      current: `监听 ${host || "未配置"}，令牌 ${tokenReady ? "已配置" : "未配置"}，本地路径 ${pathHidden ? "隐藏" : "显示"}`,
    };
  }
  if (key === "privacy_protection") {
    const historyClosed = settings.saveHistory === false;
    const cleanupReady = Boolean(settings.autoCleanTemp);
    const pathHidden = !settings.exposeLocalPaths;
    if (pathHidden && (historyClosed || cleanupReady)) {
      return {
        level: "good",
        status: "已控制",
        current: `本地路径默认隐藏，历史记录 ${historyClosed ? "关闭" : "开启"}，自动清理 ${cleanupReady ? "开启" : "关闭"}`,
      };
    }
    return {
      level: "warn",
      status: "需确认",
      current: `本地路径 ${pathHidden ? "隐藏" : "显示"}，历史记录 ${settings.saveHistory === false ? "关闭" : "开启"}，自动清理 ${cleanupReady ? "开启" : "关闭"}`,
    };
  }
  if (key === "offline_private") {
    if (settings.localClientEnabled && !settings.allowCloudSync && settings.sensitiveFilesPreferLocal) {
      return { level: "good", status: "本地优先", current: "本地客户端启用，云端同步关闭，敏感文档本地优先，可支撑离线处理流程" };
    }
    return {
      level: "blue",
      status: "规划中",
      current: `本地客户端 ${settings.localClientEnabled ? "启用" : "关闭"}，云端同步 ${settings.allowCloudSync ? "开启" : "关闭"}；私有化部署在 V3.0 路线图中`,
    };
  }
  return { level: "warn", status: "需确认", current: "等待配置" };
}

function renderCompatibilityMatrix() {
  const target = $("#compatibilityMatrixPanel");
  if (!target) return;
  const evaluated = compatibilityMatrix.map((item) => ({ ...item, ...compatibilityMatrixState(item.key) }));
  const ready = evaluated.filter((item) => item.level === "good").length;
  const limited = evaluated.filter((item) => item.level === "warn").length;
  const planned = evaluated.filter((item) => item.level === "blue").length;
  $("#compatibilityMatrixScope").textContent = `${ready} 项可用，${limited} 项需本地实测，${planned} 项可选或规划`;
  target.innerHTML = evaluated
    .map((item) => `<article class="compatibility-card">
      <div class="object-title">
        <strong>${escapeHtml(item.title)}</strong>
        <span class="badge ${item.level}">${escapeHtml(item.status)}</span>
      </div>
      <small>${escapeHtml(item.requirement)}</small>
      <div class="compatibility-item-list">
        ${item.items.map((label) => `<span>${escapeHtml(label)}</span>`).join("")}
      </div>
      <div class="compatibility-current">${escapeHtml(item.current)}</div>
      <button class="mini-button compatibility-jump" data-compatibility-anchor="${escapeHtml(item.anchor)}" type="button">查看设置</button>
    </article>`)
    .join("");
}

function compatibilityMatrixState(key) {
  const settings = state.settings || {};
  const profile = state.installProfile || state.capabilities?.installProfile || {};
  const platform = profile.platform || settings.localClientPlatform || "auto";
  const cap = state.capabilities || {};
  if (key === "operating_systems") {
    if (platform === "Windows") {
      return { level: "good", status: "Windows 优先", current: "当前安装画像为 Windows，本地端优先支持 Office COM、MathType/OLE、OMML 和宏队列" };
    }
    if (platform === "macOS") {
      return { level: "warn", status: "macOS 受限", current: "当前安装画像为 macOS，MathType/宏自动化受限，跨平台公式必须生成 MathML、LaTeX 或图片兜底" };
    }
    return { level: "blue", status: "需选择平台", current: "当前未明确 Windows 或 macOS，Linux 仅作为网页/本地服务可选支持，不承诺 Office/MathType 自动化" };
  }
  if (key === "office_suite") {
    const status = cap.officeAutomation?.status || profile.officeAutomation || "等待本地能力检测";
    if (profile.capabilities?.officeAutomation) {
      return { level: "warn", status: "需本地组件实测", current: `${status}；真实桌面 Office/WPS/LibreOffice 兼容仍由本地客户端预检确认` };
    }
    return { level: "warn", status: "受限", current: `${status}；当前网页端保留解析、转换合同和本地客户端交接，不直接声明桌面 Office 自动化已跑通` };
  }
  if (key === "formula_formats") {
    const mode = settings.mathtypeCompatibilityMode || "platform-specific";
    const objectFormat = cap.mathType?.objectFormat || profile.mathtypeObjectFormat || "MathML / LaTeX / 图片兜底";
    if (mode === "mathml-latex" || mode === "image-fallback") {
      return { level: "good", status: "兜底可用", current: `${compatibilityLabel(mode)} 已启用；当前对象格式 ${objectFormat}` };
    }
    return { level: "warn", status: "需跨平台兜底", current: `${objectFormat}；Windows 与 macOS MathType 对象不通用，交付前建议生成 MathML、LaTeX 或图片兜底` };
  }
  if (key === "macro_formats") {
    if (!settings.enableMacroExecution) {
      return { level: "good", status: "安全关闭", current: "宏检测和报告仍可用，宏执行队列关闭时不会把 .docm/.dotm 宏交给本地端执行" };
    }
    if (profile.capabilities?.macroExecution && settings.macroBackup) {
      return { level: "warn", status: "本地执行", current: "Windows 宏队列可交给本地客户端，执行前备份已开启；模板宏和本地宏库仍需来源授权" };
    }
    return { level: "warn", status: "受限", current: "宏队列需要本地客户端、执行前备份和来源授权；macOS 或未知平台不直接声明宏自动化支持" };
  }
  if (key === "mode_matrix") {
    const pdfReady = settings.pdfToWordEngine === "Mathpix";
    const localReady = Boolean(settings.localClientEnabled);
    const syncReady = Boolean(settings.allowTaskStatusCloudSync);
    if (pdfReady && localReady) {
      return {
        level: syncReady ? "good" : "warn",
        status: syncReady ? "分流可用" : "待同步授权",
        current: `PDF 转 Word 使用 Mathpix，Word/PPT/OMML/宏走本地或混合模式，状态同步 ${syncReady ? "已授权" : "未授权"}`,
      };
    }
    return { level: "warn", status: "需配置", current: "建议保持 PDF 转 Word 使用 Mathpix，并启用本地客户端以承接 Office、MathType、OMML 和宏任务" };
  }
  return { level: "warn", status: "需确认", current: "等待兼容性证据" };
}

function renderArchitectureBlueprint() {
  const target = $("#architectureBlueprintPanel");
  if (!target) return;
  const blueprint = state.architecture || {};
  const layers = blueprint.layers || [];
  const contracts = blueprint.contracts || [];
  const runtime = blueprint.runtime || {};
  const implemented = layers.filter((item) => item.level === "good").length;
  const contractsCount = layers.filter((item) => item.level === "warn").length;
  const planned = layers.filter((item) => item.level === "blue").length;
  $("#architectureBlueprintScope").textContent = `${implemented} 层已落地，${contractsCount} 层合同覆盖，${planned} 层后续接入`;
  const runtimeCard = `<article class="architecture-card architecture-runtime-card">
    <div class="object-title">
      <strong>当前运行时</strong>
      <span class="badge blue">${escapeHtml(runtime.platform || "未知平台")}</span>
    </div>
    <dl class="architecture-runtime">
      <div><dt>Python</dt><dd>${escapeHtml(runtime.python || "-")}</dd></div>
      <div><dt>服务</dt><dd>${escapeHtml(runtime.service || "-")}</dd></div>
      <div><dt>前端</dt><dd>${escapeHtml(runtime.frontend || "-")}</dd></div>
      <div><dt>存储</dt><dd>${escapeHtml(runtime.store || "-")}</dd></div>
    </dl>
  </article>`;
  const layerCards = layers
    .map((item) => `<article class="architecture-card">
      <div class="object-title">
        <strong>${escapeHtml(item.title)}</strong>
        <span class="badge ${escapeHtml(item.level)}">${escapeHtml(item.status)}</span>
      </div>
      <small>${escapeHtml(item.recommended_stack)}</small>
      <p>${escapeHtml(item.current)}</p>
    </article>`)
    .join("");
  const contractCard = `<article class="architecture-card architecture-contract-card">
    <div class="object-title">
      <strong>接口合同</strong>
      <span class="badge good">${escapeHtml(contracts.length)} 项</span>
    </div>
    <div class="architecture-contract-list">
      ${contracts
        .map((item) => `<div><strong>${escapeHtml(item.name)}</strong><code>${escapeHtml(item.endpoint)}</code><small>${escapeHtml(item.status)}</small></div>`)
        .join("")}
    </div>
  </article>`;
  target.innerHTML = `${runtimeCard}${layerCards}${contractCard}`;
}

function renderApiCatalog() {
  const target = $("#apiCatalogPanel");
  if (!target) return;
  const catalog = state.apiCatalog || {};
  const endpoints = catalog.endpoints || [];
  const sensitive = endpoints.filter((item) => item.sensitive).length;
  const tokenRequired = Boolean(catalog.auth?.token_required);
  $("#apiCatalogScope").textContent = `${endpoints.length} 个接口，${sensitive} 个敏感接口，令牌${tokenRequired ? "已启用" : "未启用"}`;
  const authCard = `<article class="api-card api-auth-card">
    <div class="object-title">
      <strong>鉴权边界</strong>
      <span class="badge ${tokenRequired ? "good" : "warn"}">${tokenRequired ? "令牌启用" : "本地模式"}</span>
    </div>
    <div class="api-auth-grid">
      <div><span>Header</span><strong>${escapeHtml(catalog.auth?.token_header || "X-K12-Token")}</strong></div>
      <div><span>Bearer</span><strong>${catalog.auth?.bearer_supported ? "支持" : "不支持"}</strong></div>
      <div><span>Query Token</span><strong>${catalog.auth?.query_token_supported ? "支持" : "不支持"}</strong></div>
      <div><span>本地载荷</span><strong>${catalog.auth?.local_payload_requires_configured_token ? "必须配置令牌" : "跟随全局"}</strong></div>
    </div>
  </article>`;
  const endpointCards = endpoints
    .map((item) => `<article class="api-card">
      <div class="api-endpoint-head">
        <span class="api-method">${escapeHtml(item.method)}</span>
        <strong>${escapeHtml(item.name)}</strong>
        <span class="badge ${item.sensitive ? "warn" : "blue"}">${item.sensitive ? "敏感" : "目录"}</span>
      </div>
      <code>${escapeHtml(item.path)}</code>
      <p>${escapeHtml(item.description)}</p>
      <small>鉴权：${escapeHtml(item.auth)}</small>
    </article>`)
    .join("");
  target.innerHTML = `${authCard}${endpointCards}`;
}

function renderProductSummary() {
  const target = $("#productSummaryPanel");
  if (!target) return;
  const summary = state.productSummary || {};
  const core = summary.core_capabilities || [];
  const phases = summary.phase_roadmap || [];
  const counts = summary.summary || {};
  $("#productSummaryScope").textContent = `${counts.implemented || 0} 项已落地，${counts.contract || 0} 项合同覆盖，${counts.planned || 0} 项规划中`;
  const positioningCard = `<article class="product-card product-position-card">
    <div class="object-title">
      <strong>产品定位</strong>
      <span class="badge blue">${escapeHtml(summary.schema_version || "k12.productSummary.v1")}</span>
    </div>
    <p>${escapeHtml(summary.positioning || "统一处理办公文档，并增强公式、宏和图片处理能力")}</p>
  </article>`;
  const capabilityCards = core
    .map((item, index) => `<article class="product-card">
      <div class="object-title">
        <strong>${index + 1}. ${escapeHtml(item.title)}</strong>
        <span class="badge ${escapeHtml(item.level)}">${escapeHtml(item.status)}</span>
      </div>
      <p>${escapeHtml(item.evidence)}</p>
    </article>`)
    .join("");
  const phaseCard = `<article class="product-card product-phase-card">
    <div class="object-title">
      <strong>产品路线</strong>
      <span class="badge good">${escapeHtml(phases.length)} 阶段</span>
    </div>
    <div class="product-phase-list">
      ${phases
        .map((item) => `<div>
          <strong>${escapeHtml(item.title)}</strong>
          <span class="badge ${escapeHtml(item.level)}">${escapeHtml(item.status)}</span>
          <small>${escapeHtml(item.evidence)}</small>
        </div>`)
        .join("")}
    </div>
  </article>`;
  target.innerHTML = `${positioningCard}${capabilityCards}${phaseCard}`;
}

function renderEnhancementPlan() {
  const target = $("#enhancementPlanPanel");
  if (!target) return;
  const plan = state.enhancementPlan || {};
  const items = plan.items || [];
  const readiness = plan.readiness || [];
  const counts = plan.summary || {};
  $("#enhancementPlanScope").textContent = `${counts.implemented || 0} 项已落地，${counts.contract || 0} 项合同覆盖，${counts.planned || 0} 项规划中`;
  const scopeCard = `<article class="enhancement-card enhancement-scope-card">
    <div class="object-title">
      <strong>${escapeHtml(plan.scope || "V3.0 增强能力规划")}</strong>
      <span class="badge blue">${escapeHtml(plan.schema_version || "k12.enhancementPlan.v1")}</span>
    </div>
    <p>AI 排版、模板套用和私有化部署先以输入、授权、风控和交接合同落地，真实执行由后续本地或私有化引擎接入。</p>
  </article>`;
  const readinessCard = `<article class="enhancement-card enhancement-readiness-card">
    <div class="object-title">
      <strong>接入前置条件</strong>
      <span class="badge good">${escapeHtml(readiness.length)} 项</span>
    </div>
    <div class="enhancement-readiness-list">
      ${readiness
        .map((item) => `<div>
          <strong>${escapeHtml(item.title)}</strong>
          <span class="badge ${escapeHtml(item.level)}">${escapeHtml(item.status)}</span>
          <small>${escapeHtml(item.detail)}</small>
        </div>`)
        .join("")}
    </div>
  </article>`;
  const itemCards = items
    .map((item) => `<article class="enhancement-card">
      <div class="object-title">
        <strong>${escapeHtml(item.title)}</strong>
        <span class="badge ${escapeHtml(item.level)}">${escapeHtml(item.status)}</span>
      </div>
      <p>${escapeHtml(item.evidence)}</p>
      <div class="enhancement-meta">
        <span>输入</span>
        <code>${escapeHtml((item.inputs || []).join(" / ") || "-")}</code>
      </div>
      <div class="enhancement-guardrails">
        ${(item.guardrails || []).map((guardrail) => `<small>${escapeHtml(guardrail)}</small>`).join("")}
      </div>
    </article>`)
    .join("");
  target.innerHTML = `${scopeCard}${readinessCard}${itemCards}`;
}

function renderAdmin() {
  renderUsers();
  renderTemplates();
  renderAuthorizations();
  renderAdminReports();
  renderPreflight();
}

function renderUsers() {
  const target = $("#userList");
  if (!target) return;
  $("#currentUserLabel").textContent = state.currentUser
    ? `${state.currentUser.name} · ${state.currentUser.role}`
    : "未选择";
  target.innerHTML = state.users.length
    ? state.users
        .map((user) => `<article class="object-card">
          <div class="object-title">
            <strong>${escapeHtml(user.name)}</strong>
            <span class="badge ${user.status === "启用" ? "good" : "warn"}">${escapeHtml(user.status)}</span>
          </div>
          <small>${escapeHtml(user.role)} · ${user.active ? "当前用户" : "可切换"}</small>
          <p>${escapeHtml((user.role_permissions || user.permissions || []).join("，"))}</p>
          <div class="row-actions">
            <button class="mini-button user-activate" data-id="${escapeHtml(user.id)}" ${user.active || user.status !== "启用" ? "disabled" : ""}>切换</button>
            <button class="mini-button user-delete danger" data-id="${escapeHtml(user.id)}">删除</button>
          </div>
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无用户</small></div>`;
}

function renderTemplates() {
  const target = $("#templateList");
  if (!target) return;
  $("#templateCount").textContent = `${state.templates.length} 个模板`;
  target.innerHTML = state.templates.length
    ? state.templates
        .map((template) => `<article class="object-card">
          <div class="object-title">
            <strong>${escapeHtml(template.name)}</strong>
            <span class="badge blue">${escapeHtml(template.template_type)}</span>
          </div>
          <small>${escapeHtml(taskLabels[template.applies_to] || template.applies_to)} · ${escapeHtml(template.status)}</small>
          <p>${escapeHtml(template.description || template.template_path || "模板参数已保存")}</p>
          <div class="row-actions">
            <button class="mini-button template-delete danger" data-id="${escapeHtml(template.id)}">删除</button>
          </div>
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无通用模板</small></div>`;
}

function renderAuthorizations() {
  const target = $("#authorizationList");
  if (!target) return;
  target.innerHTML = state.authorizations.length
    ? state.authorizations
        .map((item) => `<article class="object-card">
          <div class="object-title">
            <strong>${escapeHtml(item.name)}</strong>
            <span class="badge ${item.enabled ? "good" : "warn"}">${escapeHtml(item.status)}</span>
          </div>
          <small>${escapeHtml(item.scope)} · 风险 ${escapeHtml(item.risk_level)}</small>
          <p>${escapeHtml(item.description)}</p>
          <div class="row-actions">
            <button class="mini-button authorization-toggle" data-key="${escapeHtml(item.key)}">${item.enabled ? "撤销授权" : "授权"}</button>
          </div>
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无授权项</small></div>`;
}

function renderAdminReports() {
  const summaryTarget = $("#adminReportSummary");
  const listTarget = $("#adminReportList");
  if (!summaryTarget || !listTarget) return;
  const reports = state.reports || [];
  const failed = reports.filter((report) => Number(report.fail_count || 0) > 0 || report.status === "失败").length;
  const qualityIssues = reports.reduce((sum, report) => sum + Number(report.quality_issue_count || 0), 0);
  const artifacts = reports.reduce((sum, report) => sum + ((report.analysis?.artifacts || []).filter((item) => item.url).length), 0);
  summaryTarget.innerHTML = [
    ["报告", reports.length, "全部记录"],
    ["失败报告", failed, "需处理"],
    ["质量问题", qualityIssues, "检查项"],
    ["转换产物", artifacts, "可下载"],
  ].map(metricCard).join("");
  listTarget.innerHTML = reports.length
    ? reports
        .slice(0, 6)
        .map((report) => `<article class="object-card">
          <div class="object-title">
            <strong>${escapeHtml(report.report_type)}</strong>
            <span class="badge ${statusClass(report.status)}">${escapeHtml(report.status)}</span>
          </div>
          <small>${escapeHtml(taskLabels[report.task_type] || report.task_type)} · ${formatTime(report.created_at)}</small>
          <p>成功 ${escapeHtml(report.success_count || 0)} · 失败 ${escapeHtml(report.fail_count || 0)} · 待本地 ${escapeHtml(report.pending_count || 0)} · 已取消 ${escapeHtml(report.cancelled_count || 0)} · 公式 ${escapeHtml(report.formula_count || 0)} · 宏 ${escapeHtml(report.macro_count || 0)} · 微小图 ${escapeHtml(report.small_image_count || 0)}</p>
          <div class="row-actions">${downloadButtons(report.id)}<button class="mini-button report-delete danger" data-id="${escapeHtml(report.id)}" ${permissionButtonAttrs(true, "删除报告", "reports.manage")}>删除</button></div>
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无报告</small></div>`;
}

function renderPreflight() {
  const target = $("#preflightList");
  if (!target) return;
  const failed = state.preflightChecks.filter((item) => item.status === "失败").length;
  const warnings = state.preflightChecks.filter((item) => item.status === "需确认").length;
  $("#preflightSummary").textContent = failed ? `${failed} 项失败` : warnings ? `${warnings} 项需确认` : "全部通过";
  target.innerHTML = state.preflightChecks.length
    ? state.preflightChecks
        .map((item) => `<article class="object-card">
          <div class="object-title">
            <strong>${escapeHtml(item.label)}</strong>
            <span class="badge ${qualityBadgeClass(item.status)}">${escapeHtml(item.status)}</span>
          </div>
          <small>${escapeHtml(item.metric)}</small>
          <p>${escapeHtml(item.message)}</p>
          <small>${escapeHtml(item.recommendation)}</small>
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无预检结果</small></div>`;
}

function renderWordInsight() {
  const source = contextualFiles();
  const stats = [
    ["Word 文件", source.filter((file) => file.file_type === "Word").length, "参与 Word 流程"],
    ["含公式", source.filter((file) => file.has_formula).length, "需预检"],
    ["含 OMML", source.filter((file) => file.has_omml).length, "建议转 MathType"],
    ["含 MathType", source.filter((file) => file.has_mathtype).length, "可格式化"],
    ["含宏", source.filter((file) => file.has_macro).length, "需确认风险"],
    ["校验失败", source.filter((file) => file.status === "校验失败").length, "需处理"],
  ];
  $("#wordInsight").innerHTML = stats.map(metricCard).join("");
}

function renderWordWorkspace() {
  const files = state.files.filter((file) => file.file_type === "Word");
  const selectedWordCount = selectedFileObjects().filter((file) => file.file_type === "Word").length;
  if ($("#wordStats")) {
    $("#wordStats").innerHTML = [
      ["Word", files.length, selectedWordCount ? `已选择 ${selectedWordCount}` : "当前队列"],
      ["公式", files.filter((file) => file.has_formula).length, "预检对象"],
      ["OMML", files.filter((file) => file.has_omml).length, "转 MathType"],
      ["宏", files.filter((file) => file.has_macro).length, "需授权执行"],
    ].map(metricCard).join("");
  }
  renderWordFilePreviewList(files);
  renderWordFormulaDetection(files);
  renderWordMacroDetection(files);
  renderWordReportLinks();
}

function renderWordFilePreviewList(files) {
  const target = $("#wordFilePreviewList");
  if (!target) return;
  target.innerHTML = files.length
    ? files
        .map((file) => {
          const summary = file.content_summary || {};
          const active = state.selectedFiles.has(file.id) ? "已选择" : "未选择";
          return `<article class="object-card">
            <div class="object-title">
              <strong>${escapeHtml(file.file_name)}</strong>
              <span class="badge ${state.selectedFiles.has(file.id) ? "good" : "blue"}">${active}</span>
            </div>
            <small>页段 ${escapeHtml(fileMetric(file))} · 段落 ${escapeHtml(summary.paragraphs || 0)} · 标题 ${escapeHtml(summary.headings || 0)}</small>
            <p>图片 ${escapeHtml(summary.images || 0)} · 表格 ${escapeHtml(summary.tables || 0)} · 批注 ${escapeHtml(summary.comments || 0)} · 修订 ${escapeHtml(summary.revisions || 0)}</p>
            <div class="row-actions">
              <button class="mini-button word-select" data-id="${escapeHtml(file.id)}" type="button">选择</button>
              <button class="mini-button file-preview" data-id="${escapeHtml(file.id)}" type="button">预览</button>
            </div>
          </article>`;
        })
        .join("")
    : `<div class="object-card"><small>暂无 Word 文件</small></div>`;
}

function renderWordFormulaDetection(files) {
  const target = $("#wordFormulaDetectionList");
  if (!target) return;
  const rows = files.filter((file) => file.has_formula || file.has_omml || file.has_mathtype);
  target.innerHTML = rows.length
    ? rows
        .map((file) => {
          const summary = file.content_summary || {};
          const missing = summary.missingOmmlDependency;
          const status = missing ? "需修复" : file.has_omml ? "需确认" : "可处理";
          return `<article class="object-card">
            <div class="object-title">
              <strong>${escapeHtml(file.file_name)}</strong>
              <span class="badge ${missing ? "bad" : file.has_omml ? "warn" : "good"}">${status}</span>
            </div>
            <small>公式 ${escapeHtml(formulaCount(file))} · OMML ${escapeHtml(ommlCount(file))} · MathType ${file.has_mathtype ? "是" : "否"}</small>
            <p>${missing ? "缺少 OMML 依赖，需自动检索或手动指定后再转换。" : "可进入公式预检、OMML 转 MathType 或 MathType 格式化流程。"}</p>
          </article>`;
        })
        .join("")
    : `<div class="object-card"><small>暂无公式、OMML 或 MathType 检测结果</small></div>`;
}

function renderWordMacroDetection(files) {
  const target = $("#wordMacroDetectionList");
  if (!target) return;
  const rows = files.filter((file) => file.has_macro);
  target.innerHTML = rows.length
    ? rows
        .map((file) => `<article class="object-card">
          <div class="object-title">
            <strong>${escapeHtml(file.file_name)}</strong>
            <span class="badge warn">需授权</span>
          </div>
          <small>宏数量 ${escapeHtml(macroCount(file))} · 来源 当前文档或模板</small>
          <p>宏任务会进入安全队列，按宏页顺序和失败策略交给本地客户端执行。</p>
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无 Word 宏检测结果</small></div>`;
}

function renderWordReportLinks() {
  const target = $("#wordReportLinkList");
  if (!target) return;
  const reports = state.reports
    .filter((report) => ["formula_precheck", "omml_to_mathtype", "mathtype_format", "macro_sequence", "word_to_ppt"].includes(report.task_type))
    .slice(0, 5);
  target.innerHTML = reports.length
    ? reports
        .map((report) => `<article class="object-card compact-object">
          <div class="object-title">
            <strong>${escapeHtml(taskLabels[report.task_type] || report.task_type)}</strong>
            <span class="badge ${statusClass(report.status)}">${escapeHtml(report.status)}</span>
          </div>
          <small>${escapeHtml(formatDate(report.created_at))} · 公式 ${escapeHtml(report.formula_count || 0)} · 宏 ${escapeHtml(report.macro_count || 0)}</small>
          ${downloadButtons(report.id)}
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无 Word 处理报告</small></div>`;
}

function renderWordFlow() {
  const steps = [
    ["1", "公式预检", "识别 OMML、MathType、LaTeX 与图片公式"],
    ["2", "依赖修复", "缺少 OMML 文件时自动检索并复制到文档目录"],
    ["3", "宏编排", "按顺序执行已确认的 Word 宏并生成日志"],
    ["4", "转换输出", "Word 转 PPT 或进入后续格式化、图片检索流程"],
    ["5", "报告导出", "生成处理、公式、宏、图片和失败清单"],
  ];
  $("#wordFlow").innerHTML = steps
    .map(([num, title, detail]) => `<div class="flow-step"><span>${num}</span><strong>${title}</strong><small>${detail}</small></div>`)
    .join("");
}

function renderPdfWorkspace() {
  const files = state.files.filter((file) => file.file_type === "PDF");
  const selectedPdfCount = selectedFileObjects().filter((file) => file.file_type === "PDF").length;
  const scanned = files.filter((file) => (file.content_summary || {}).pdfType === "扫描型 PDF").length;
  const mixed = files.filter((file) => (file.content_summary || {}).pdfType === "混合型 PDF").length;
  const formulaHints = files.reduce((sum, file) => sum + toCount((file.content_summary || {}).formulaHints), 0);
  const tableHints = files.reduce((sum, file) => sum + toCount((file.content_summary || {}).tableHints), 0);
  const imageObjects = files.reduce((sum, file) => sum + toCount((file.content_summary || {}).imageObjects), 0);
  if ($("#pdfStats")) {
    $("#pdfStats").innerHTML = [
      ["PDF", files.length, "当前队列"],
      ["扫描/混合", scanned + mixed, "建议 Mathpix OCR"],
      ["公式线索", formulaHints, "公式 OCR"],
      ["表格线索", tableHints, "表格 OCR"],
    ].map(metricCard).join("");
  }
  const fileList = $("#pdfFileList");
  if (fileList) {
    fileList.innerHTML = files.length
      ? files
          .map((file) => {
            const summary = file.content_summary || {};
            const active = state.selectedFiles.has(file.id) ? "已选择" : "未选择";
            return `<article class="object-card">
              <div class="object-title">
                <strong>${escapeHtml(file.file_name)}</strong>
                <span class="badge ${state.selectedFiles.has(file.id) ? "good" : "blue"}">${active}</span>
              </div>
              <small>${escapeHtml(summary.pdfType || "未知 PDF")} · ${escapeHtml(summary.ocrRecommendation || "等待分析")}</small>
              <p>页段 ${escapeHtml(fileMetric(file))} · 图片 ${escapeHtml(summary.imageObjects || 0)} · 公式 ${escapeHtml(summary.formulaHints || 0)} · 表格 ${escapeHtml(summary.tableHints || 0)}</p>
              <div class="row-actions">
                <button class="mini-button pdf-select" data-id="${escapeHtml(file.id)}" type="button">选择</button>
                <button class="mini-button file-preview" data-id="${escapeHtml(file.id)}" type="button">预览</button>
              </div>
            </article>`;
          })
          .join("")
      : `<div class="object-card"><small>暂无 PDF 文件，可先在工作台添加 PDF</small></div>`;
  }
  renderPdfOcrSettings(selectedPdfCount, imageObjects, formulaHints, tableHints);
  renderPdfMathpixJobs();
  renderPdfLowConfidence();
  renderPdfArtifacts();
}

function renderPdfOcrSettings(selectedPdfCount, imageObjects, formulaHints, tableHints) {
  const target = $("#pdfOcrSettings");
  if (!target) return;
  const settings = state.settings || {};
  const authStatus = settings.allowExternalMathpixUpload ? "已授权" : "未授权";
  const checks = [
    ["Mathpix 引擎", settings.pdfToWordEngine || "Mathpix", "PDF 转 Word 固定使用 Mathpix"],
    ["外部上传", authStatus, settings.allowExternalMathpixUpload ? "可提交 OCR 作业" : "仅生成授权提示和保留计划"],
    ["文字 OCR", settings.enableTextOcr === false ? "关闭" : "开启", "扫描页文字识别"],
    ["公式 OCR", settings.enableFormulaOcr === false ? "关闭" : "开启", `公式线索 ${formulaHints} 个`],
    ["表格 OCR", settings.enableTableOcr === false ? "关闭" : "开启", `表格线索 ${tableHints} 个`],
    ["保留计划", `PDF ${selectedPdfCount || "未选"} 个`, `图片对象 ${imageObjects} 个`],
  ];
  target.innerHTML = checks
    .map(([label, value, detail]) => `<article class="object-card compact-object">
      <div class="object-title">
        <strong>${escapeHtml(label)}</strong>
        <span class="badge ${value === "关闭" || value === "未授权" ? "warn" : "good"}">${escapeHtml(value)}</span>
      </div>
      <small>${escapeHtml(detail)}</small>
    </article>`)
    .join("");
}

function renderPdfMathpixJobs() {
  const target = $("#pdfMathpixJobList");
  if (!target) return;
  const queue = state.mathpixJobs || {};
  const jobs = queue.jobs || [];
  const summary = queue.summary || {};
  const summaryCard = `<article class="object-card compact-object mathpix-summary-card">
    <div class="object-title">
      <strong>${escapeHtml(queue.engine || "Mathpix")}</strong>
      <span class="badge ${queue.external_upload_allowed ? "good" : "warn"}">${escapeHtml(queue.authorization || "未授权")}</span>
    </div>
    <small>总数 ${escapeHtml(summary.total || 0)} · 阻塞 ${escapeHtml(summary.blocked || 0)} · 已完成 ${escapeHtml(summary.completed || 0)}</small>
    <small>凭证变量：${escapeHtml((queue.credential_envs || []).join(" / ") || "MATHPIX_APP_ID / MATHPIX_APP_KEY")}</small>
  </article>`;
  const jobCards = jobs.length
    ? jobs
        .slice(0, 5)
        .map((job) => {
          const plan = job.recognition_plan || {};
          const gate = plan.upload_gate || {};
          const gateReasons = (gate.blocking_reasons || []).join(" / ");
          return `<article class="object-card compact-object">
            <div class="object-title">
              <strong>${escapeHtml(job.file_name || job.file_id)}</strong>
              <span class="badge ${mathpixJobBadge(job.status)}">${escapeHtml(job.status_label || job.status)}</span>
            </div>
            <small>${escapeHtml(job.message || "等待 Mathpix 状态")}</small>
            <small>${escapeHtml(job.retention?.pdf_type || "PDF")} · 公式 ${escapeHtml(job.retention?.formula_hints || 0)} · 表格 ${escapeHtml(job.retention?.table_hints || 0)} · 输出 ${escapeHtml((job.outputs || []).length)}</small>
            <small>识别计划：${escapeHtml(plan.status_label || plan.status || "待生成")} · 格式 ${escapeHtml((plan.conversion_formats || []).join(" / ") || "docx")} · ${plan.submit_allowed ? "可提交" : "等待授权或预检"}</small>
            <small>上传门禁：${escapeHtml(gate.status || "待生成")} · ${escapeHtml(gateReasons || "无阻塞")}</small>
          </article>`;
        })
        .join("")
    : `<div class="object-card"><small>创建 PDF 转 Word 任务后，会在这里显示 Mathpix 授权、提交、轮询和输出状态。</small></div>`;
  target.innerHTML = `${summaryCard}${jobCards}`;
}

function mathpixJobBadge(status) {
  if (status === "completed") return "good";
  if (["authorization_required", "missing_credentials", "ocr_disabled", "missing_local_file"].includes(status)) return "warn";
  if (["api_error", "download_failed"].includes(status)) return "bad";
  return "blue";
}

function renderPdfLowConfidence() {
  const target = $("#pdfLowConfidenceList");
  if (!target) return;
  const formulas = state.reports
    .flatMap((report) => (report.analysis?.formulas || []).map((formula) => ({ ...formula, report_id: report.id })))
    .filter((formula) => formula.source_type === "PDF" && (formula.status === "待确认" || Number(formula.confidence || 0) < 80))
    .slice(0, 6);
  target.innerHTML = formulas.length
    ? formulas
        .map((formula) => `<article class="object-card compact-object">
          <div class="object-title">
            <strong>${escapeHtml(formula.position || formula.id)}</strong>
            <span class="badge ${formulaBadgeClass(formula.status, formula.confidence)}">${escapeHtml(formula.confidence || 0)}%</span>
          </div>
          <small>${escapeHtml(formula.latex || "等待识别结果")}</small>
          <small>报告：${escapeHtml(formula.report_id)}</small>
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无 PDF 低置信度内容</small></div>`;
}

function renderPdfArtifacts() {
  const target = $("#pdfArtifactList");
  if (!target) return;
  const reports = state.reports.filter((report) => report.task_type === "pdf_to_word" || report.analysis?.mathpix?.length);
  const artifacts = reports
    .flatMap((report) => (report.analysis?.artifacts || []).map((artifact) => ({ ...artifact, report_id: report.id })))
    .filter((artifact) => artifact.output_type === "docx" || artifact.output_type === "tex.zip")
    .slice(0, 6);
  target.innerHTML = artifacts.length
    ? artifacts
        .map((artifact) => `<article class="object-card compact-object">
          <div class="object-title">
            <strong>${escapeHtml(artifact.file_name || artifact.output_type)}</strong>
            <span class="badge ${statusClass(artifact.status)}">${escapeHtml(artifact.status)}</span>
          </div>
          <small>${escapeHtml(artifact.message || "Mathpix 输出")}</small>
          ${artifact.url ? `<a class="mini-button" href="${escapeHtml(withToken(artifact.url))}" target="_blank" rel="noopener">下载</a>` : ""}
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无 PDF 转 Word 输出</small></div>`;
}

function renderPptWorkspace() {
  const files = pptFileObjects();
  const selectedPpts = selectedFileObjects().filter((file) => file.file_type === "PPT");
  const active = files.find((file) => file.id === state.activeFileId) || selectedPpts[0] || files[0];
  const totals = files.reduce(
    (acc, file) => {
      const summary = file.content_summary || {};
      acc.slides += toCount(summary.slides || file.slide_count);
      acc.notes += toCount(summary.notes);
      acc.images += toCount(summary.images);
      acc.formulas += toCount(summary.formulas);
      return acc;
    },
    { slides: 0, notes: 0, images: 0, formulas: 0 }
  );
  if ($("#pptStats")) {
    $("#pptStats").innerHTML = [
      ["PPT", files.length, "当前队列"],
      ["幻灯片", totals.slides, "可转讲义"],
      ["备注", totals.notes, "可提取"],
      ["公式", totals.formulas, "可进公式报告"],
    ].map(metricCard).join("");
  }
  const list = $("#pptFileList");
  if (list) {
    list.innerHTML = files.length
      ? files
          .map((file) => {
            const summary = file.content_summary || {};
            const selected = state.selectedFiles.has(file.id);
            const activeLabel = file.id === active?.id ? "预览对象" : (selected ? "已选择" : "未选择");
            return `<article class="object-card">
              <div class="object-title">
                <strong>${escapeHtml(file.file_name)}</strong>
                <span class="badge ${selected ? "good" : "blue"}">${escapeHtml(activeLabel)}</span>
              </div>
              <small>幻灯片 ${escapeHtml(summary.slides || file.slide_count || 0)} · 备注 ${escapeHtml(summary.notes || 0)}</small>
              <p>图片 ${escapeHtml(summary.images || 0)} · 表格 ${escapeHtml(summary.tables || 0)} · 图表 ${escapeHtml(summary.charts || 0)} · 公式 ${escapeHtml(summary.formulas || 0)}</p>
              <div class="row-actions">
                <button class="mini-button ppt-select" data-id="${escapeHtml(file.id)}" type="button">选择</button>
                <button class="mini-button file-preview" data-id="${escapeHtml(file.id)}" type="button">预览</button>
              </div>
            </article>`;
          })
          .join("")
      : `<div class="object-card"><small>暂无 PPT 文件，可先在工作台添加 PPT</small></div>`;
  }
  renderPptContentPreview(active);
  renderPptArtifacts();
}

function renderPptContentPreview(file) {
  const target = $("#pptContentPreview");
  if (!target) return;
  if (!file) {
    target.innerHTML = `<div class="object-card"><small>暂无 PPT 内容预览</small></div>`;
    return;
  }
  const preview = state.filePreview?.file_id === file.id ? state.filePreview : null;
  const pages = (preview?.pages || []).slice(0, 4);
  const objects = preview?.objects || pptSummaryObjects(file);
  const objectBadges = objects
    .filter((item) => Number(item.count || 0) > 0)
    .slice(0, 8)
    .map((item) => `<span class="badge blue">${escapeHtml(item.label || item.type)} ${escapeHtml(item.count || 0)}</span>`)
    .join("");
  target.innerHTML = `<article class="object-card compact-object">
    <div class="object-title">
      <strong>${escapeHtml(file.file_name)}</strong>
      <span class="badge ${preview ? "good" : "warn"}">${preview ? "已加载" : "待预览"}</span>
    </div>
    <small>${preview ? `${escapeHtml(preview.mode)} · ${escapeHtml(preview.status)}` : "点击预览后显示幻灯片标题、正文和对象统计"}</small>
    <div class="summary-tags">${objectBadges || "<span>暂无对象统计</span>"}</div>
  </article>${
    pages.length
      ? pages
          .map((page) => `<article class="object-card compact-object">
            <div class="object-title">
              <strong>${escapeHtml(page.title || `幻灯片 ${page.index || ""}`)}</strong>
              <span class="badge blue">${escapeHtml(page.kind || "幻灯片")} ${escapeHtml(page.index || "")}</span>
            </div>
            <p>${escapeHtml(page.text || "无正文")}</p>
          </article>`)
          .join("")
      : `<article class="object-card compact-object"><small>可点击“预览”加载前 4 页幻灯片正文。</small></article>`
  }`;
}

function pptSummaryObjects(file) {
  const summary = file?.content_summary || {};
  return [
    ["幻灯片", summary.slides || file?.slide_count || 0],
    ["备注", summary.notes || 0],
    ["图片", summary.images || 0],
    ["表格", summary.tables || 0],
    ["公式", summary.formulas || 0],
    ["图表", summary.charts || 0],
    ["文本框", summary.textBoxes || 0],
    ["形状", summary.shapes || 0],
  ].map(([label, count]) => ({ label, count }));
}

function renderPptArtifacts() {
  const target = $("#pptArtifactList");
  if (!target) return;
  const artifacts = state.reports
    .filter((report) => report.task_type === "ppt_to_word")
    .flatMap((report) => (report.analysis?.artifacts || []).map((artifact) => ({ ...artifact, report_id: report.id })))
    .filter((artifact) => artifact.output_type === "docx")
    .slice(0, 6);
  target.innerHTML = artifacts.length
    ? artifacts
        .map((artifact) => `<article class="object-card compact-object">
          <div class="object-title">
            <strong>${escapeHtml(artifact.file_name || "PPT 转 Word 输出")}</strong>
            <span class="badge ${statusClass(artifact.status)}">${escapeHtml(artifact.status)}</span>
          </div>
          <small>${escapeHtml(artifact.message || "已生成 Word 讲义")}</small>
          ${artifact.url ? `<a class="mini-button" href="${escapeHtml(withToken(artifact.url))}" target="_blank" rel="noopener">下载</a>` : ""}
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无 PPT 转 Word 输出</small></div>`;
}

function renderOmmlDependencies() {
  const latest = state.reports.find((report) => report.analysis?.ommlDependencies?.length);
  const dependencies = latest?.analysis?.ommlDependencies || [];
  const target = $("#ommlDependencyList");
  if (!target) return;
  target.innerHTML = dependencies.length
    ? dependencies
        .map((item) => {
          const annotation = ommlAnnotation(item.id, latest?.id);
          return `<article class="object-card">
          <strong>${escapeHtml(item.omml_file_name || "未找到")}</strong>
          <small>检索：${escapeHtml(item.found_status)} · 复制：${escapeHtml(item.copy_status)}</small>
          ${annotation ? `<span class="badge ${statusClass(annotation.status)}">${escapeHtml(annotation.status)}</span>` : ""}
          <p>${escapeHtml(item.error_message || item.omml_target_path || "依赖已记录")}</p>
          <div class="row-actions">
            <button class="mini-button omml-keep" data-report="${latest.id}" data-id="${item.id}" type="button">保留 OMML</button>
            <button class="mini-button omml-retry" data-report="${latest.id}" data-id="${item.id}" type="button">重新转换</button>
            <button class="mini-button omml-fixed" data-report="${latest.id}" data-id="${item.id}" type="button">已修复</button>
            <button class="mini-button omml-manual" data-report="${latest.id}" data-id="${item.id}" type="button">指定依赖</button>
          </div>
        </article>`;
        })
        .join("")
    : `<div class="object-card"><small>暂无 OMML 依赖报告</small></div>`;
}

function renderMacros() {
  const term = ($("#macroSearch")?.value || "").toLowerCase();
  const source = $("#macroSourceFilter")?.value || "all";
  const purpose = $("#macroPurposeFilter")?.value || "all";
  const recent = $("#macroRecentFilter")?.value || "all";
  const macros = state.macros.filter((macro) => {
    const text = `${macro.macro_name || ""} ${macro.macro_description || ""} ${macro.macro_purpose || ""}`.toLowerCase();
    if (term && !text.includes(term)) return false;
    if (source !== "all" && macro.macro_source !== source) return false;
    if (purpose !== "all" && macro.macro_purpose !== purpose) return false;
    if (recent === "recent" && !macro.recently_used) return false;
    if (recent === "unused" && macro.recently_used) return false;
    return true;
  });
  $("#macroLibrary").innerHTML = macros.length
    ? macros
        .map((macro) => {
          const selected = state.macroSequence.includes(macro.id);
          return `<div class="macro-item">
            <div>
              <strong>${escapeHtml(macro.macro_name)}</strong>
              <small>${escapeHtml(macro.macro_source)} · ${escapeHtml(macro.macro_purpose || "通用")} · ${escapeHtml(macro.macro_description)}</small>
              <span class="badge ${macro.local_only ? "warn" : "blue"}">${macro.local_only ? "本地" : "网页"}</span>
              ${macro.recently_used ? `<span class="badge good">最近使用 ${Number(macro.usage_count || 0)}</span>` : ""}
            </div>
            <div class="row-actions">
              <button class="mini-button macro-detail" data-id="${macro.id}">详情</button>
              <button class="mini-button macro-add" data-id="${macro.id}" ${selected ? "disabled" : ""}>加入</button>
            </div>
          </div>`;
        })
        .join("")
    : `<div class="empty-mini">没有匹配的宏</div>`;
  renderMacroOrder();
}

function renderMacroDetail() {
  const target = $("#macroDetailPanel");
  if (!target) return;
  const macro = state.macros.find((item) => item.id === state.activeMacroId) || state.macros[0];
  if (!macro) {
    if ($("#macroDetailScope")) $("#macroDetailScope").textContent = "未选择";
    target.innerHTML = `<div class="empty-mini">暂无宏详情</div>`;
    return;
  }
  state.activeMacroId = macro.id;
  if ($("#macroDetailScope")) $("#macroDetailScope").textContent = `${macro.macro_source} · ${macro.macro_purpose || "通用"}`;
  const selected = state.macroSequence.includes(macro.id);
  const rows = [
    ["宏 ID", macro.id],
    ["来源", macro.macro_source],
    ["用途", macro.macro_purpose || "通用"],
    ["执行顺序", macro.execute_order || "-"],
    ["最近使用", macro.recently_used ? `是，${Number(macro.usage_count || 0)} 次` : "否"],
    ["执行位置", macro.local_only ? "本地客户端" : "网页编排"],
    ["当前状态", selected ? "已加入执行顺序" : "未加入"],
  ];
  target.innerHTML = `<article class="macro-detail-card">
    <div class="object-title">
      <strong>${escapeHtml(macro.macro_name)}</strong>
      <span class="badge ${macro.local_only ? "warn" : "blue"}">${macro.local_only ? "本地" : "网页"}</span>
    </div>
    <p>${escapeHtml(macro.macro_description || "暂无说明")}</p>
    <div class="detail-list compact-detail-list">
      <dl>${rows.map(([label, value]) => `<dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd>`).join("")}</dl>
    </div>
    <div class="row-actions">
      <button class="secondary-button macro-add" data-id="${escapeHtml(macro.id)}" ${selected ? "disabled" : ""}>加入执行顺序</button>
      <button class="secondary-button quick-task" data-task="macro_sequence">创建宏任务</button>
    </div>
  </article>`;
}

function renderMacroOrder() {
  const order = state.macroSequence.map((id) => state.macros.find((macro) => macro.id === id)).filter(Boolean);
  $("#macroOrder").innerHTML = order.length
    ? order
        .map((macro, index) => `<li class="macro-order-item" draggable="true" data-macro="${escapeHtml(macro.id)}">
          <span class="macro-drag-handle" title="拖拽调整宏顺序" aria-hidden="true">${lineIcon("grip")}</span>
          <span class="order-index">${index + 1}</span>
          <div>
            <strong>${escapeHtml(macro.macro_name)}</strong>
            <small>${escapeHtml(macro.execute_timing)} · ${escapeHtml(macro.macro_source)}</small>
          </div>
          <div class="row-actions">
            <button class="icon-button planner-icon-button macro-up" data-id="${macro.id}" ${index === 0 ? "disabled" : ""} title="上移" aria-label="上移">${lineIcon("arrow-up")}</button>
            <button class="icon-button planner-icon-button macro-down" data-id="${macro.id}" ${index === order.length - 1 ? "disabled" : ""} title="下移" aria-label="下移">${lineIcon("arrow-down")}</button>
            <button class="mini-button macro-remove danger" data-id="${macro.id}">移除</button>
          </div>
        </li>`)
        .join("")
    : `<li class="empty-order">从宏库加入需要执行的宏</li>`;
}

function renderMacroTemplates() {
  const select = $("#macroTemplateSelect");
  const count = $("#macroTemplateCount");
  if (!select) return;
  const current = select.value;
  select.innerHTML = `<option value="">选择宏顺序模板</option>${state.macroTemplates
    .map((template) => `<option value="${escapeHtml(template.id)}">${escapeHtml(template.name)} · ${template.macro_sequence?.length || 0} 个宏</option>`)
    .join("")}`;
  select.value = state.macroTemplates.some((template) => template.id === current) ? current : "";
  if (count) count.textContent = `${state.macroTemplates.length} 个模板`;
}

function renderMacroResults() {
  const latest = latestReport();
  const macros = latest?.analysis?.macros || [];
  $("#macroResults").innerHTML = macros.length
    ? macros
        .slice(0, 6)
        .map((macro, index) => `<article class="object-card">
          <strong>${escapeHtml(macro.execute_order)}. ${escapeHtml(macro.macro_name)}</strong>
          <small>${escapeHtml(macro.macro_source)} · ${escapeHtml(macro.failure_strategy)}</small>
          <p>${escapeHtml(macro.error_message || macro.execute_status)}</p>
          <span class="badge ${statusClass(macro.execute_status)}">${escapeHtml(macro.execute_status)}</span>
          <div class="row-actions">
            <button class="mini-button macro-rerun" data-report="${escapeHtml(latest.id)}" data-id="${escapeHtml(macro.id)}" data-index="${index}" type="button">重跑</button>
          </div>
        </article>`)
        .join("")
    : `<div class="object-card"><small>暂无宏执行报告</small></div>`;
}

function renderMacroLogs() {
  const target = $("#macroLogList");
  if (!target) return;
  const latestMacroReport = state.reports.find((report) => report.task_type === "macro_sequence" || report.analysis?.macros?.length);
  const reportMacros = latestMacroReport?.analysis?.macros || [];
  const macroTaskIds = new Set(state.tasks.filter((task) => task.task_type === "macro_sequence").map((task) => task.id));
  const logs = state.logs.filter((log) => log.category === "macro" || macroTaskIds.has(log.task_id));
  if ($("#macroLogScope")) {
    $("#macroLogScope").textContent = latestMacroReport
      ? `${taskLabels[latestMacroReport.task_type] || latestMacroReport.task_type} · ${formatDate(latestMacroReport.created_at)}`
      : "最近宏任务";
  }
  if (logs.length) {
    renderLogTarget(target, logs, "暂无宏执行日志");
    return;
  }
  target.innerHTML = reportMacros.length
    ? reportMacros
        .slice(0, 8)
        .map((macro) => `<div class="log-item ${macro.execute_status === "失败" ? "error" : "info"}">
          <span>${escapeHtml(macro.macro_name)} · ${escapeHtml(macro.execute_status)}</span>
          <small>${escapeHtml(macro.failure_strategy || "跳过")} · ${escapeHtml(macro.error_message || macro.failure_policy?.description || "等待本地客户端回传")}</small>
        </div>`)
        .join("")
    : `<div class="log-item"><small>暂无宏执行日志</small></div>`;
}

function renderImages() {
  const latest = latestReport();
  const kind = $("#imageKindFilter")?.value || "all";
  const sourceId = $("#imageSourceFilter")?.value || "all";
  const duplicateFilter = $("#imageDuplicateFilter")?.value || "all";
  const locationTerm = ($("#imageLocationFilter")?.value || "").trim().toLowerCase();
  const maxArea = Number($("#imageMaxAreaFilter")?.value || 0);
  syncImageSourceFilter(latest);
  let images = latest?.analysis?.smallImages || [];
  if (kind !== "all") {
    const key = {
      formula: "is_formula_like",
      qrcode: "is_qrcode_like",
      stamp: "is_stamp_like",
      signature: "is_signature_like",
    }[kind];
    images = images.filter((item) => item[key]);
  }
  if (sourceId !== "all") images = images.filter((item) => item.file_id === sourceId);
  if (duplicateFilter === "unique") images = images.filter((item) => !item.is_duplicate);
  if (duplicateFilter === "duplicate") images = images.filter((item) => item.is_duplicate);
  if (locationTerm) {
    images = images.filter((item) => `${item.location || ""} ${item.source_name || ""}`.toLowerCase().includes(locationTerm));
  }
  if (maxArea > 0) images = images.filter((item) => Number(item.area || 0) <= maxArea);
  $("#imageStats").innerHTML = [
    ["图片", images.length, "当前筛选"],
    ["公式", images.filter((item) => item.is_formula_like).length, "公式截图"],
    ["二维码", images.filter((item) => item.is_qrcode_like).length, "可定位"],
    ["印章/签名", images.filter((item) => item.is_stamp_like || item.is_signature_like).length, "审核关注"],
  ].map(metricCard).join("");
  const active = activeImageItem(images);
  $("#smallImageGrid").innerHTML = images.length
    ? images
        .map((item) => {
          const annotation = imageAnnotation(item.id, latest?.id);
          const activeClass = item.id === active?.id ? " active" : "";
          return `<article class="image-card${activeClass} ${imageAnnotationMuted(annotation) ? "muted-card" : ""}">
          <div class="image-thumb">${imagePreview(item)}</div>
          <div class="object-title">
            <strong>${imageLabel(item)}</strong>
            ${annotation ? `<span class="badge ${imageAnnotationBadge(annotation.status)}">${escapeHtml(annotation.status)}</span>` : ""}
          </div>
          ${annotation?.label ? `<small>确认类型：${escapeHtml(annotation.label)}</small>` : ""}
          ${annotation?.replacement_file_name ? `<small>替换为：${escapeHtml(annotation.replacement_file_name)}</small>` : ""}
          ${annotation?.replacement_asset_url ? `<small>替换素材已登记 · ${formatBytes(annotation.replacement_file_size || 0)}</small>` : ""}
          <small>导出状态：${escapeHtml(item.export_status || "可导出")}</small>
          <small>${escapeHtml(item.location)} · ${item.confidence}% · ${item.area} px</small>
          <div class="row-actions image-actions">
            <button class="mini-button image-detail" data-id="${escapeHtml(item.id)}">详情</button>
            <button class="mini-button image-open" data-url="${escapeHtml(item.asset_url || "")}" ${item.asset_url ? "" : "disabled"}>查看</button>
            <button class="mini-button image-download" data-url="${escapeHtml(item.asset_url || "")}" ${item.asset_url ? "" : "disabled"}>导出</button>
            <button class="mini-button image-reexport" data-id="${escapeHtml(item.id)}" data-report="${escapeHtml(latest?.id || "")}">重新导出</button>
            <button class="mini-button image-locate" data-id="${escapeHtml(item.id)}" data-report="${escapeHtml(latest?.id || "")}" data-file="${escapeHtml(item.file_id || "")}" data-location="${escapeHtml(item.location || item.source_name || "")}">定位</button>
            <button class="mini-button image-confirm-type" data-id="${item.id}" data-report="${latest?.id || ""}" data-label="${escapeHtml(imageKindLabel(item))}">确认类型</button>
            <button class="mini-button image-false-positive" data-id="${item.id}" data-report="${latest?.id || ""}">标误判</button>
            <button class="mini-button image-delete-mark" data-id="${item.id}" data-report="${latest?.id || ""}">删图</button>
            <button class="mini-button image-replace-mark" data-id="${item.id}" data-report="${latest?.id || ""}">替换</button>
          </div>
        </article>`;
        })
        .join("")
    : `<div class="image-card"><small>暂无图片检索报告</small></div>`;
  renderImageDetail(latest, active);
}

function activeImageItem(images) {
  if (!images.length) {
    state.activeImageId = "";
    return null;
  }
  const active = images.find((item) => item.id === state.activeImageId) || images[0];
  state.activeImageId = active.id;
  return active;
}

function renderImageDetail(report, item) {
  const target = $("#imageDetailPanel");
  if (!target) return;
  if (!report || !item) {
    target.innerHTML = `<div class="empty-mini">选择图片后查看详情</div>`;
    return;
  }
  const annotation = imageAnnotation(item.id, report.id);
  const badges = imageDetailBadges(item)
    .map((label) => `<span class="badge blue">${escapeHtml(label)}</span>`)
    .join("");
  target.innerHTML = `<div class="image-detail-head">
    <div>
      <h3>图片详情</h3>
      <small>${escapeHtml(item.source_name || item.file_id || "未知来源")}</small>
    </div>
    ${annotation ? `<span class="badge ${imageAnnotationBadge(annotation.status)}">${escapeHtml(annotation.status)}</span>` : `<span class="badge warn">待判断</span>`}
  </div>
  <div class="image-detail-preview">${imagePreview(item)}</div>
  <div class="image-detail-meta">
    <span><b>位置</b><small>${escapeHtml(item.location || "未记录")}</small></span>
    <span><b>尺寸</b><small>${escapeHtml(item.width || 0)} × ${escapeHtml(item.height || 0)}</small></span>
    <span><b>面积</b><small>${escapeHtml(item.area || 0)} px</small></span>
    <span><b>置信度</b><small>${escapeHtml(item.confidence || 0)}%</small></span>
    <span><b>格式</b><small>${escapeHtml(item.image_format || item.format || "未知")}</small></span>
    <span><b>重复</b><small>${item.is_duplicate ? "是" : "否"}</small></span>
    <span><b>重复判断</b><small>${escapeHtml(item.duplicate_check_status || "未比较")}</small></span>
    <span><b>异常兜底</b><small>${escapeHtml(item.duplicate_fallback || "保留原始结果")}</small></span>
    <span><b>导出状态</b><small>${escapeHtml(item.export_status || "可导出")}</small></span>
    <span><b>导出信息</b><small>${escapeHtml(item.export_message || "可直接导出")}</small></span>
  </div>
  <div class="summary-tags">${badges || "<span>暂无类型判断</span>"}</div>
  ${annotation?.label ? `<small>人工确认类型：${escapeHtml(annotation.label)}</small>` : ""}
  ${annotation?.replacement_file_name ? `<small>替换目标：${escapeHtml(annotation.replacement_file_name)}</small>` : ""}
  ${annotation?.replacement_asset_url ? `<small>替换素材：${escapeHtml(annotation.replacement_file_name || "未命名图片")} · ${formatBytes(annotation.replacement_file_size || 0)} · ${escapeHtml((annotation.replacement_sha256 || "").slice(0, 12))}</small>` : ""}
  <div class="row-actions image-actions">
    <button class="secondary-button image-open" data-url="${escapeHtml(item.asset_url || "")}" ${item.asset_url ? "" : "disabled"}>查看</button>
    <button class="secondary-button image-download" data-url="${escapeHtml(item.asset_url || "")}" ${item.asset_url ? "" : "disabled"}>导出</button>
    <button class="secondary-button image-reexport" data-id="${escapeHtml(item.id)}" data-report="${escapeHtml(report.id)}">重新导出</button>
    ${annotation?.replacement_asset_url ? `<button class="secondary-button image-open" data-url="${escapeHtml(annotation.replacement_asset_url)}">查看替换素材</button>` : ""}
    <button class="secondary-button image-locate" data-id="${escapeHtml(item.id)}" data-report="${escapeHtml(report.id)}" data-file="${escapeHtml(item.file_id || "")}" data-location="${escapeHtml(item.location || item.source_name || "")}">定位</button>
    <button class="mini-button image-confirm-type" data-id="${escapeHtml(item.id)}" data-report="${escapeHtml(report.id)}" data-label="${escapeHtml(imageKindLabel(item))}">确认类型</button>
    <button class="mini-button image-false-positive" data-id="${escapeHtml(item.id)}" data-report="${escapeHtml(report.id)}">标误判</button>
    <button class="mini-button image-delete-mark" data-id="${escapeHtml(item.id)}" data-report="${escapeHtml(report.id)}">删图</button>
    <button class="mini-button image-replace-mark" data-id="${escapeHtml(item.id)}" data-report="${escapeHtml(report.id)}">替换</button>
  </div>`;
}

function imageDetailBadges(item) {
  const labels = [];
  if (item.is_formula_like) labels.push("疑似公式");
  if (item.is_qrcode_like) labels.push("疑似二维码");
  if (item.is_stamp_like) labels.push("疑似印章");
  if (item.is_signature_like) labels.push("疑似签名");
  if (item.is_header_footer) labels.push("页眉页脚");
  if (item.is_watermark) labels.push("水印");
  if (item.is_transparent) labels.push("透明图");
  if (item.is_duplicate) labels.push("重复图片");
  return labels;
}

function syncImageSourceFilter(report) {
  const select = $("#imageSourceFilter");
  if (!select || !report) return;
  const current = select.value || "all";
  const files = report.files || [];
  select.innerHTML = `<option value="all">全部文件</option>${files
    .map((file) => `<option value="${escapeHtml(file.id)}">${escapeHtml(file.file_name)}</option>`)
    .join("")}`;
  select.value = files.some((file) => file.id === current) ? current : "all";
}

function filteredFiles() {
  if (state.filter === "all") return state.files;
  if (state.filter === "异常") return state.files.filter((file) => file.status !== "待处理" && file.status !== "成功");
  return state.files.filter((file) => file.file_type === state.filter);
}

function selectedFileObjects() {
  return state.files.filter((file) => state.selectedFiles.has(file.id));
}

function pdfFileObjects() {
  return state.files.filter((file) => file.file_type === "PDF");
}

function wordFileObjects() {
  return state.files.filter((file) => file.file_type === "Word");
}

function pptFileObjects() {
  return state.files.filter((file) => file.file_type === "PPT");
}

function selectWordFiles() {
  const words = wordFileObjects();
  if (!words.length) {
    toast("请先添加 Word 文件");
    return [];
  }
  state.selectedFiles = new Set(words.map((file) => file.id));
  state.activeFileId = words[0].id;
  if ($("#taskTypeSelect")) $("#taskTypeSelect").value = "word_to_ppt";
  render();
  toast(`已选择 ${words.length} 个 Word`);
  return words;
}

async function previewSelectedWord() {
  let file = selectedFileObjects().find((item) => item.file_type === "Word");
  if (!file) file = wordFileObjects()[0];
  if (!file) {
    toast("请先添加 Word 文件");
    return;
  }
  state.activeFileId = file.id;
  await loadFilePreview(file.id);
  switchView("word");
}

async function startWordToPpt() {
  let words = selectedFileObjects().filter((file) => file.file_type === "Word");
  if (!words.length) words = selectWordFiles();
  if (!words.length) return;
  state.selectedFiles = new Set(words.map((file) => file.id));
  if ($("#taskTypeSelect")) $("#taskTypeSelect").value = "word_to_ppt";
  await createTask("word_to_ppt");
}

function selectPdfFiles() {
  const pdfs = pdfFileObjects();
  if (!pdfs.length) {
    toast("请先添加 PDF 文件");
    return [];
  }
  state.selectedFiles = new Set(pdfs.map((file) => file.id));
  state.activeFileId = pdfs[0].id;
  if ($("#taskTypeSelect")) $("#taskTypeSelect").value = "pdf_to_word";
  render();
  toast(`已选择 ${pdfs.length} 个 PDF`);
  return pdfs;
}

async function startPdfToWord() {
  let pdfs = selectedFileObjects().filter((file) => file.file_type === "PDF");
  if (!pdfs.length) pdfs = selectPdfFiles();
  if (!pdfs.length) return;
  state.selectedFiles = new Set(pdfs.map((file) => file.id));
  if ($("#taskTypeSelect")) $("#taskTypeSelect").value = "pdf_to_word";
  await createTask("pdf_to_word");
}

function openPdfSettings() {
  switchView("settings");
  const target = $("#allowExternalMathpixUpload") || $("#pdfToWordEngine");
  target?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function selectPptFiles() {
  const ppts = pptFileObjects();
  if (!ppts.length) {
    toast("请先添加 PPT 文件");
    return [];
  }
  state.selectedFiles = new Set(ppts.map((file) => file.id));
  state.activeFileId = ppts[0].id;
  if ($("#taskTypeSelect")) $("#taskTypeSelect").value = "ppt_to_word";
  render();
  toast(`已选择 ${ppts.length} 个 PPT`);
  return ppts;
}

async function previewSelectedPpt() {
  let file = selectedFileObjects().find((item) => item.file_type === "PPT");
  if (!file) file = pptFileObjects()[0];
  if (!file) {
    toast("请先添加 PPT 文件");
    return;
  }
  state.activeFileId = file.id;
  await loadFilePreview(file.id);
  switchView("word");
}

async function startPptToWord() {
  let ppts = selectedFileObjects().filter((file) => file.file_type === "PPT");
  if (!ppts.length) ppts = selectPptFiles();
  if (!ppts.length) return;
  state.selectedFiles = new Set(ppts.map((file) => file.id));
  if ($("#taskTypeSelect")) $("#taskTypeSelect").value = "ppt_to_word";
  await createTask("ppt_to_word");
}

function contextualFiles() {
  const selected = selectedFileObjects();
  return selected.length ? selected : state.files;
}

async function importFiles(fileList) {
  const files = Array.from(fileList || []);
  if (!files.length) return;
  const form = new FormData();
  files.forEach((file) => form.append("files", file, file.k12RelativePath || file.webkitRelativePath || file.name));
  const response = await fetch("/api/uploads", {
    method: "POST",
    headers: authHeaders(),
    body: form,
  });
  const result = await response.json();
  if (!response.ok) {
    toast(result.error || "上传失败");
    return;
  }
  result.files.forEach((file) => state.selectedFiles.add(file.id));
  $("#fileInput").value = "";
  $("#folderInput").value = "";
  await loadAll();
  toast(uploadValidationSummary(result.files));
}

async function importDroppedItems(dataTransfer) {
  const files = await droppedFiles(dataTransfer);
  await importFiles(files.length ? files : dataTransfer?.files);
}

async function droppedFiles(dataTransfer) {
  const items = Array.from(dataTransfer?.items || []);
  const entries = items.map((item) => item.webkitGetAsEntry?.()).filter(Boolean);
  if (!entries.length) return Array.from(dataTransfer?.files || []);
  const nested = await Promise.all(entries.map((entry) => collectDroppedEntryFiles(entry)));
  return nested.flat();
}

async function collectDroppedEntryFiles(entry, prefix = "") {
  if (entry.isFile) {
    return new Promise((resolve) => {
      entry.file(
        (file) => {
          file.k12RelativePath = `${prefix}${file.name}`;
          resolve([file]);
        },
        () => resolve([]),
      );
    });
  }
  if (!entry.isDirectory) return [];
  const children = await readAllDroppedDirectoryEntries(entry.createReader());
  const nested = await Promise.all(children.map((child) => collectDroppedEntryFiles(child, `${prefix}${entry.name}/`)));
  return nested.flat();
}

function readAllDroppedDirectoryEntries(reader) {
  return new Promise((resolve, reject) => {
    const entries = [];
    const readNext = () => {
      reader.readEntries((batch) => {
        if (!batch.length) {
          resolve(entries);
          return;
        }
        entries.push(...batch);
        readNext();
      }, reject);
    };
    readNext();
  });
}

async function loadFilePreview(fileId) {
  const result = await api(`/api/files/${fileId}/preview`);
  state.activeFileId = fileId;
  state.filePreview = result.preview;
  state.previewSearch = "";
  state.previewZoom = 1;
  state.previewPageOffset = 0;
  renderFiles();
  renderFileDetail();
  renderFilePreview();
}

async function markImageLocationUnknown(reportId, imageId, note) {
  if (!reportId || !imageId) return false;
  try {
    await annotateImage(reportId, imageId, "位置未知", note || "图片定位失败，等待人工复核", "", { silent: true });
    return true;
  } catch (error) {
    toast(error.message || "图片位置未知标记失败");
    return false;
  }
}

async function locateImageInSource(fileId, location, reportId = "", imageId = "") {
  if (!fileId) {
    const marked = await markImageLocationUnknown(reportId, imageId, "图片缺少来源文件信息，已标记位置未知");
    toast(marked ? "图片定位失败，已标记为位置未知" : "图片缺少来源文件信息");
    return;
  }
  try {
    await loadFilePreview(fileId);
  } catch (error) {
    const marked = await markImageLocationUnknown(reportId, imageId, `来源文件预览失败：${error.message || "未知错误"}`);
    toast(marked ? "图片定位失败，已标记为位置未知" : error.message || "来源文件预览失败");
    return;
  }
  const term = String(location || "").trim();
  const matched = !term || (state.filePreview?.pages || []).some((page) => previewPageSearchText(page).includes(term.toLowerCase()));
  state.previewSearch = term;
  state.previewPageOffset = 0;
  renderFilePreview();
  switchView("workspace");
  if (!matched) {
    const marked = await markImageLocationUnknown(reportId, imageId, "来源预览中未命中图片位置，已标记位置未知");
    toast(marked ? "图片定位失败，已标记为位置未知" : "未在来源预览中命中图片位置");
    return;
  }
  toast("已跳转到来源文件预览");
}

async function openFileReports(fileId) {
  const file = state.files.find((item) => item.id === fileId);
  if (!file) {
    toast("文件不存在");
    return;
  }
  try {
    const result = await api(`/api/files/${encodeURIComponent(fileId)}/reports`);
    const reports = result.reports || [];
    if (!reports.length) {
      toast("当前文件暂无报告");
      return;
    }
    const incomingIds = new Set(reports.map((report) => report.id));
    state.reports = [...reports, ...state.reports.filter((report) => !incomingIds.has(report.id))];
    state.activeReportId = reports[0].id;
    switchView("tasks");
    renderReportCenter();
    toast(`已打开 ${reports.length} 份关联报告`);
  } catch (error) {
    toast(error.message || "关联报告读取失败");
  }
}

async function setFilePassword(fileId) {
  const file = state.files.find((item) => item.id === fileId);
  const password = window.prompt(`请输入 ${file?.file_name || "加密文件"} 的打开密码`, "");
  if (password === null) return;
  if (!password) {
    toast("密码不能为空");
    return;
  }
  await api(`/api/files/${fileId}/password`, {
    method: "POST",
    body: JSON.stringify({ password }),
  });
  toast("密码已登记到当前本地会话");
  await loadAll();
}

function replaceFile(fileId) {
  const input = document.createElement("input");
  input.type = "file";
  input.addEventListener("change", async () => {
    const file = input.files?.[0];
    if (!file) return;
    const form = new FormData();
    form.append("files", file, file.name);
    try {
      const result = await api(`/api/files/${fileId}/replace`, {
        method: "POST",
        body: form,
      });
      state.selectedFiles.add(result.file.id);
      state.activeFileId = result.file.id;
      state.filePreview = null;
      const validationError = primaryValidationError(result.file);
      toast(validationError ? `文件已替换，需处理：${validationError}` : "文件已重新上传并替换");
      await loadAll();
    } catch (error) {
      toast(error.message || "重新上传失败");
    }
  });
  input.click();
}

function downloadSelectedFiles() {
  const ids = [...state.selectedFiles];
  if (!ids.length) {
    toast("请先选择文件");
    return;
  }
  window.open(withToken(`/api/files/download?ids=${encodeURIComponent(ids.join(","))}`), "_blank", "noopener");
}

function downloadFilteredFiles() {
  const type = state.filter || "all";
  window.open(withToken(`/api/files/download?type=${encodeURIComponent(type)}`), "_blank", "noopener");
}

function macroSequencePayload() {
  return state.macroSequence.map((id, index) => ({
    id,
    execute_order: index + 1,
    failure_strategy: $("#macroFailureStrategy")?.value || state.settings.macroFailureStrategy,
    execute_timing: $("#macroTimingSelect")?.value || "Word 处理前",
    confirmed: $("#confirmMacroRisk")?.checked || false,
  }));
}

async function saveMacroTemplate() {
  const sequence = macroSequencePayload();
  if (!sequence.length) {
    toast("请先把宏加入执行顺序");
    return;
  }
  const result = await api("/api/macro-templates", {
    method: "POST",
    body: JSON.stringify({
      name: $("#macroTemplateName")?.value || `宏顺序模板 ${state.macroTemplates.length + 1}`,
      macro_sequence: sequence,
      failureStrategy: $("#macroFailureStrategy")?.value || state.settings.macroFailureStrategy,
      executeTiming: $("#macroTimingSelect")?.value || "Word 处理前",
      confirmMacroRisk: $("#confirmMacroRisk")?.checked || false,
    }),
  });
  state.macroTemplates = [result.template, ...state.macroTemplates.filter((item) => item.id !== result.template.id)];
  $("#macroTemplateSelect").value = result.template.id;
  toast("宏顺序模板已保存");
  renderMacroTemplates();
}

function loadMacroTemplate() {
  const template = selectedMacroTemplate();
  if (!template) {
    toast("请选择宏顺序模板");
    return;
  }
  const ids = (template.macro_sequence || []).map((item) => item.id).filter((id) => state.macros.some((macro) => macro.id === id));
  state.macroSequence = [...new Set(ids)];
  const defaults = template.defaults || {};
  if ($("#macroFailureStrategy")) $("#macroFailureStrategy").value = defaults.failureStrategy || template.macro_sequence?.[0]?.failure_strategy || "跳过";
  if ($("#macroTimingSelect")) $("#macroTimingSelect").value = defaults.executeTiming || template.macro_sequence?.[0]?.execute_timing || "Word 处理前";
  if ($("#confirmMacroRisk")) $("#confirmMacroRisk").checked = Boolean(defaults.confirmMacroRisk);
  $("#macroTemplateName").value = template.name;
  toast("宏顺序模板已加载");
  renderMacros();
  renderMacroTemplates();
}

async function deleteMacroTemplate() {
  const template = selectedMacroTemplate();
  if (!template) {
    toast("请选择宏顺序模板");
    return;
  }
  await api(`/api/macro-templates/${template.id}`, { method: "DELETE" });
  state.macroTemplates = state.macroTemplates.filter((item) => item.id !== template.id);
  $("#macroTemplateSelect").value = "";
  toast("宏顺序模板已删除");
  renderMacroTemplates();
}

async function saveUser() {
  const name = $("#userName")?.value?.trim();
  if (!name) {
    toast("请输入用户名称");
    return;
  }
  await api("/api/users", {
    method: "POST",
    body: JSON.stringify({
      name,
      role: $("#userRole")?.value || "学生",
      status: $("#userStatus")?.value || "启用",
      active: $("#userActive")?.checked || false,
    }),
  });
  $("#userName").value = "";
  $("#userActive").checked = false;
  toast("用户已保存");
  await loadAll();
}

async function activateUser(userId) {
  await api("/api/session", {
    method: "POST",
    body: JSON.stringify({ user_id: userId }),
  });
  toast("当前用户已切换");
  await loadAll();
}

async function deleteUser(userId) {
  await api(`/api/users/${userId}`, { method: "DELETE" });
  toast("用户已删除");
  await loadAll();
}

async function saveTemplate() {
  const name = $("#templateName")?.value?.trim();
  if (!name) {
    toast("请输入模板名称");
    return;
  }
  await api("/api/templates", {
    method: "POST",
    body: JSON.stringify({
      name,
      template_type: $("#templateType")?.value || "PPT 模板",
      applies_to: $("#templateAppliesTo")?.value || "word_to_ppt",
      template_path: $("#templatePath")?.value || "",
      description: $("#templateDescription")?.value || "",
      settings: {
        wordToPptTemplate: name,
        mode: $("#pptToWordMode")?.value || state.settings.pptToWordMode,
      },
    }),
  });
  $("#templateName").value = "";
  $("#templatePath").value = "";
  $("#templateDescription").value = "";
  toast("模板已保存");
  await loadAll();
}

async function deleteTemplate(templateId) {
  await api(`/api/templates/${templateId}`, { method: "DELETE" });
  toast("模板已删除");
  await loadAll();
}

async function toggleAuthorization(key) {
  const item = state.authorizations.find((authorization) => authorization.key === key);
  if (!item) return;
  const enabled = !item.enabled;
  const confirmed = enabled || window.confirm(`撤销 ${item.name}？相关功能会回到未授权状态。`);
  if (!confirmed) return;
  await api("/api/authorizations", {
    method: "POST",
    body: JSON.stringify({
      key,
      enabled,
      note: enabled ? "用户在网页管理页授权" : "用户在网页管理页撤销授权",
      updated_by: state.currentUser?.name || "本地管理员",
    }),
  });
  toast(enabled ? "授权已开启" : "授权已撤销");
  await loadAll();
}

async function markImageFalsePositive(reportId, imageId) {
  return annotateImage(reportId, imageId, "误判", "用户在图片页标记为误判");
}

async function annotateImage(reportId, imageId, status, note, replacementFileName = "", options = {}) {
  if (!reportId || !imageId) {
    toast("没有可标记的图片报告");
    return;
  }
  const result = await api("/api/image-annotations", {
    method: "POST",
    body: JSON.stringify({
      report_id: reportId,
      image_id: imageId,
      status,
      note,
      replacement_file_name: replacementFileName,
    }),
  });
  state.imageAnnotations = [result.annotation, ...state.imageAnnotations.filter((item) => item.id !== result.annotation.id)];
  state.activeImageId = imageId;
  if (!options.silent) toast(`图片已标记为${status}`);
  renderImages();
}

async function confirmImageType(reportId, imageId, currentLabel) {
  const confirmedLabel = window.prompt("确认图片类型（公式/二维码/印章/签名/图标）", currentLabel || "图标");
  if (confirmedLabel === null) return;
  const label = confirmedLabel.trim();
  if (!label) {
    toast("请填写图片类型");
    return;
  }
  if (!["公式", "二维码", "印章", "签名", "图标"].includes(label)) {
    toast("图片类型需为公式、二维码、印章、签名或图标");
    return;
  }
  const result = await api("/api/image-annotations", {
    method: "POST",
    body: JSON.stringify({
      report_id: reportId,
      image_id: imageId,
      status: "类型已确认",
      label,
      note: `用户确认图片类型为${label}`,
    }),
  });
  state.imageAnnotations = [result.annotation, ...state.imageAnnotations.filter((item) => item.id !== result.annotation.id)];
  state.activeImageId = imageId;
  toast(`图片类型已确认为${label}`);
  renderImages();
}

async function markImageReplacement(reportId, imageId) {
  if (!reportId || !imageId) {
    toast("没有可替换的图片报告");
    return;
  }
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/png,image/jpeg,image/gif,image/webp,image/bmp";
  input.addEventListener("change", async () => {
    const file = input.files?.[0];
    if (!file) return;
    try {
      const contentBase64 = await readFileAsDataUrl(file);
      const result = await api("/api/image-annotations/replacement", {
        method: "POST",
        body: JSON.stringify({
          report_id: reportId,
          image_id: imageId,
          file_name: file.name,
          mime_type: file.type,
          content_base64: contentBase64,
          note: "用户选择替换图片，等待本地客户端改写源文档",
        }),
      });
      state.imageAnnotations = [result.annotation, ...state.imageAnnotations.filter((item) => item.id !== result.annotation.id)];
      state.activeImageId = imageId;
      toast(`替换图片已登记：${file.name}`);
      renderImages();
    } catch (error) {
      toast(error.message || "替换图片登记失败");
    }
  });
  input.click();
}

async function reexportImage(reportId, imageId) {
  if (!reportId || !imageId) {
    toast("没有可重新导出的图片报告");
    return;
  }
  const result = await api("/api/image-assets/reexport", {
    method: "POST",
    body: JSON.stringify({ report_id: reportId, image_id: imageId }),
  });
  if (result.report) {
    state.reports = [result.report, ...state.reports.filter((item) => item.id !== result.report.id)];
  }
  if (result.image?.id) state.activeImageId = result.image.id;
  toast(result.message || (result.reexported ? "图片已重新导出" : "图片重新导出失败"));
  renderImages();
  renderReports();
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(String(reader.result || "")));
    reader.addEventListener("error", () => reject(new Error("替换图片读取失败")));
    reader.readAsDataURL(file);
  });
}

async function annotateFormula(reportId, formulaId, status) {
  if (!reportId || !formulaId) {
    toast("没有可校正的公式报告");
    return;
  }
  const input = $$(".formula-detail-latex").find((node) => node.dataset.id === formulaId)
    || $$(".formula-latex-input").find((node) => node.dataset.id === formulaId);
  const latex = input?.value || "";
  const note = {
    已确认: "用户确认公式识别结果",
    跳过: "用户跳过该公式处理",
    重新识别: "用户要求重新识别该公式",
    已修正: "用户保存修正后的 LaTeX",
  }[status] || "";
  const result = await api("/api/formula-annotations", {
    method: "POST",
    body: JSON.stringify({
      report_id: reportId,
      formula_id: formulaId,
      status,
      latex,
      note,
    }),
  });
  state.formulaAnnotations = [result.annotation, ...state.formulaAnnotations.filter((item) => item.id !== result.annotation.id)];
  state.activeFormulaId = formulaId;
  toast(`公式已标记为${status}`);
  renderFormulaList();
}

async function annotateOmml(reportId, dependencyId, status, note = "", extra = {}) {
  if (!reportId || !dependencyId) {
    toast("没有可校正的 OMML 依赖报告");
    return;
  }
  const result = await api("/api/omml-annotations", {
    method: "POST",
    body: JSON.stringify({
      report_id: reportId,
      dependency_id: dependencyId,
      status,
      keep_omml: status === "保留OMML",
      retry_conversion: status === "重新转换",
      note,
      ...extra,
    }),
  });
  state.ommlAnnotations = [result.annotation, ...state.ommlAnnotations.filter((item) => item.id !== result.annotation.id)];
  toast(`OMML 校正已标记为${status}`);
  renderOmmlDependencies();
}

async function bulkConfirmFormulas() {
  const report = latestReport();
  if (!report?.analysis?.formulas?.length) {
    toast("暂无可批量确认的公式报告");
    return;
  }
  const result = await api("/api/formula-annotations/bulk-confirm", {
    method: "POST",
    body: JSON.stringify({
      report_id: report.id,
      min_confidence: 80,
      only_unannotated: true,
    }),
  });
  const incoming = result.annotations || [];
  state.formulaAnnotations = [
    ...incoming,
    ...state.formulaAnnotations.filter((item) => !incoming.some((annotation) => annotation.id === item.id)),
  ];
  toast(`已批量确认 ${result.confirmed_count || 0} 个公式`);
  renderFormulaList();
}

function exportFormulas() {
  const report = latestReport();
  if (!report?.analysis?.formulas?.length) {
    toast("暂无可导出的公式报告");
    return;
  }
  window.open(withToken(`/api/reports/${report.id}/formulas.zip`), "_blank", "noopener");
}

function exportFormulaSheet() {
  const report = latestReport();
  if (!report?.analysis?.formulas?.length) {
    toast("暂无可导出的公式表");
    return;
  }
  window.open(withToken(`/api/reports/${report.id}/formulas.xlsx`), "_blank", "noopener");
}

function exportOmmlFailures() {
  const report = state.reports.find((item) => item.analysis?.ommlDependencies?.length) || latestReport();
  if (!report) {
    toast("暂无可导出的 OMML 报告");
    return;
  }
  window.open(withToken(`/api/reports/${report.id}/omml-failures.csv`), "_blank", "noopener");
}

function exportMacroFailures() {
  const report = state.reports.find((item) => item.analysis?.macros?.length) || latestReport();
  if (!report) {
    toast("暂无可导出的宏执行报告");
    return;
  }
  window.open(withToken(`/api/reports/${report.id}/macro-failures.csv`), "_blank", "noopener");
}

async function saveLayoutAnnotation() {
  const report = latestReport();
  if (!report) {
    toast("暂无可校正报告");
    return;
  }
  const location = ($("#layoutLocation")?.value || "").trim();
  if (!location) {
    toast("请填写排版问题位置");
    return;
  }
  const result = await api("/api/layout-annotations", {
    method: "POST",
    body: JSON.stringify({
      report_id: report.id,
      issue_type: $("#layoutIssueType")?.value || "标题层级",
      status: $("#layoutStatus")?.value || "待校正",
      location,
      page_index: $("#layoutPageIndex")?.value || "",
      before: $("#layoutBefore")?.value || "",
      after: $("#layoutAfter")?.value || "",
      recommendation: $("#layoutRecommendation")?.value || "",
      note: $("#layoutNote")?.value || "",
    }),
  });
  state.layoutAnnotations = [result.annotation, ...state.layoutAnnotations.filter((item) => item.id !== result.annotation.id)];
  ["layoutLocation", "layoutPageIndex", "layoutBefore", "layoutAfter", "layoutRecommendation", "layoutNote"].forEach((id) => {
    const node = $(`#${id}`);
    if (node) node.value = "";
  });
  toast("排版校正已保存");
  renderLayoutReview();
}

function exportLatestImages() {
  const report = latestReport();
  if (!report?.analysis?.smallImages?.length) {
    toast("暂无可导出的图片报告");
    return;
  }
  window.open(withToken(`/api/reports/${report.id}/images.zip`), "_blank", "noopener");
}

function exportImageManifest() {
  const report = latestReport();
  if (!report?.analysis?.smallImages?.length) {
    toast("暂无可导出的图片清单");
    return;
  }
  window.open(withToken(`/api/reports/${report.id}/images.xlsx`), "_blank", "noopener");
}

function exportLogs() {
  const format = encodeURIComponent((state.settings.logExportFormat || "txt").toLowerCase());
  const taskPart = state.activeLogTaskId ? `&task_id=${encodeURIComponent(state.activeLogTaskId)}` : "";
  window.open(withToken(`/api/logs/download?format=${format}${taskPart}`), "_blank", "noopener");
}

async function viewTaskLogs(taskId) {
  state.activeLogTaskId = taskId;
  state.taskLogs = (await api(`/api/logs?task_id=${encodeURIComponent(taskId)}`)).logs;
  renderLogs();
  $("#logList")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function clearTaskLogFilter() {
  state.activeLogTaskId = "";
  state.taskLogs = [];
  renderLogs();
}

function activeTask() {
  return state.tasks.find((task) => !["成功", "失败", "已取消"].includes(task.status)) || state.tasks[0] || null;
}

async function pauseActiveTask() {
  const task = activeTask();
  if (!task || ["成功", "失败", "已取消", "已暂停", "已中断"].includes(task.status)) {
    toast("当前没有可暂停的任务");
    return;
  }
  await api(`/api/tasks/${task.id}/pause`, { method: "POST", body: "{}" });
  await loadAll();
}

async function stopActiveTask() {
  const task = activeTask();
  if (!task || ["成功", "失败", "已取消"].includes(task.status)) {
    toast("当前没有可停止的任务");
    return;
  }
  await api(`/api/tasks/${task.id}/cancel`, { method: "POST", body: "{}" });
  await loadAll();
}

function exportLatestReport() {
  const report = latestReport();
  if (!report) {
    toast("暂无可导出的报告");
    return;
  }
  window.open(withToken(`/api/reports/${report.id}/download?format=pdf`), "_blank", "noopener");
}

async function loadComparison(reportId) {
  const result = await api(`/api/reports/${reportId}/comparison`);
  state.previewComparison = result.comparison;
  renderComparisonPanel();
  switchView("tasks");
}

async function cleanupMaintenance() {
  const retentionDays = Math.max(1, Number($("#cleanupRetentionDays")?.value || state.settings.cleanupRetentionDays || 30));
  const logRetentionDays = Math.max(1, Number($("#logRetentionDays")?.value || state.settings.logRetentionDays || retentionDays));
  const confirmed = window.confirm(`清理 ${retentionDays} 天前的任务、报告、输出和 ${logRetentionDays} 天前日志？`);
  if (!confirmed) return;
  const result = await api("/api/maintenance/cleanup", {
    method: "POST",
    body: JSON.stringify({ retentionDays, logRetentionDays }),
  });
  await loadAll();
  toast(`清理完成：任务 ${result.tasks_deleted}，报告 ${result.reports_deleted}，日志 ${result.logs_deleted}`);
}

function selectedMacroTemplate() {
  const id = $("#macroTemplateSelect")?.value || "";
  return state.macroTemplates.find((template) => template.id === id);
}

async function createTask(taskType = $("#taskTypeSelect").value || defaultPlannerTask()) {
  const selected = selectedFileObjects();
  if (!selected.length) {
    $("#taskTypeSelect").value = taskType;
    renderRouteAdvisor();
    toast("请先选择文件");
    return;
  }
  const modeChoice = $("#executeModeSelect")?.value || "auto";
  const payload = {
    task_type: taskType,
    file_ids: selected.map((file) => file.id),
    options: collectTaskOptions(taskType),
  };
  if (modeChoice !== "auto") {
    payload.execute_mode = modeChoice;
  }
  const result = await api("/api/tasks", {
    method: "POST",
    body: JSON.stringify(payload),
  });
  await loadAll();
  if (result.task && result.task.history_saved === false) {
    toast(result.task.history_message || "历史记录已关闭，任务结果未保留");
  }
  if (result.task?.completion_notice) {
    toast(result.task.completion_notice.message || "任务已完成");
  }
  switchView("tasks");
}

async function createFileTask(fileId, taskType) {
  const file = state.files.find((item) => item.id === fileId);
  if (!file) {
    toast("文件不存在");
    return;
  }
  state.selectedFiles = new Set([file.id]);
  state.activeFileId = file.id;
  if ($("#taskTypeSelect")) $("#taskTypeSelect").value = taskType;
  await createTask(taskType);
}

async function rerunMacroFromReport(reportId, macroId, macroIndex) {
  const report = state.reports.find((item) => item.id === reportId) || latestReport();
  const macros = report?.analysis?.macros || [];
  const index = Number(macroIndex);
  const macro = Number.isInteger(index) && macros[index]?.id === macroId ? macros[index] : macros.find((item) => item.id === macroId);
  if (!macro?.file_id) {
    toast("没有可重新执行的宏记录");
    return;
  }
  const failureStrategy = macro.failure_strategy || $("#macroFailureStrategy")?.value || state.settings.macroFailureStrategy || "跳过";
  const executeTiming = macro.execute_timing || $("#macroTimingSelect")?.value || "Word 处理前";
  const payload = {
    task_type: "macro_sequence",
    execute_mode: "local",
    file_ids: [macro.file_id],
    options: {
      selectedMacros: [
        {
          id: macro.id,
          execute_order: 1,
          failure_strategy: failureStrategy,
          execute_timing: executeTiming,
          confirmed: true,
          selected: true,
        },
      ],
      failureStrategy,
      executeTiming,
      confirmMacroRisk: true,
      macroBackup: true,
      macroTimeoutSeconds: macro.timeout_seconds || state.settings.macroTimeoutSeconds,
    },
  };
  const result = await api("/api/tasks", { method: "POST", body: JSON.stringify(payload) });
  await loadAll();
  toast(result.task?.completion_notice?.message || `已重新创建宏任务：${macro.macro_name || macro.id}`);
  switchView("tasks");
}

function loadMacroSequenceFromLatestReport() {
  const report = state.reports.find((item) => item.analysis?.macros?.length);
  const reportMacros = report?.analysis?.macros || [];
  const seen = new Set();
  const ordered = reportMacros
    .slice()
    .sort((a, b) => Number(a.execute_order || 0) - Number(b.execute_order || 0))
    .filter((macro) => {
      if (!macro.id || seen.has(macro.id) || !state.macros.some((item) => item.id === macro.id)) return false;
      seen.add(macro.id);
      return true;
    });
  if (!ordered.length) {
    toast("暂无可载入的宏执行顺序");
    return;
  }
  state.macroSequence = ordered.map((macro) => macro.id);
  const fileIds = reportMacros.map((macro) => macro.file_id).filter(Boolean);
  if (fileIds.length) {
    state.selectedFiles = new Set(fileIds);
  }
  const first = ordered[0];
  setNodeValue($("#macroFailureStrategy"), first.failure_strategy || state.settings.macroFailureStrategy || "跳过");
  setSelectIfOptionExists($("#macroTimingSelect"), first.execute_timing || "Word 处理前");
  setNodeValue($("#macroBackup"), true);
  setNodeValue($("#confirmMacroRisk"), true);
  $("#taskTypeSelect").value = "macro_sequence";
  $("#executeModeSelect").value = "local";
  renderFiles();
  renderMacros();
  renderRouteAdvisor();
  renderTaskPlanner();
  toast(`已载入 ${ordered.length} 个宏，可调整顺序后重新执行`);
}

function collectTaskOptions(taskType = "") {
  const workspaceFailureStrategy = $("#workspaceFailureStrategy")?.value || "跳过";
  const workspaceConcurrency = clampNumber($("#workspaceConcurrency")?.value, 3, 1, 8);
  const workflowOrder = orderedPlannerTasks(contextualFiles());
  const batchOptions = taskType === "batch_process"
    ? {
        batchFailureStrategy: workspaceFailureStrategy,
        continueOnFailure: workspaceFailureStrategy === "跳过",
        maxConcurrentTasks: workspaceConcurrency,
      }
    : {};
  return {
    workflowOrder,
    workflowLabels: workflowOrder.map((item) => taskLabels[item] || item),
    failureStrategy: $("#macroFailureStrategy")?.value || state.settings.macroFailureStrategy,
    ...batchOptions,
    executeTiming: $("#macroTimingSelect")?.value || "Word 处理前",
    macroBackup: $("#macroBackup")?.checked ?? state.settings.macroBackup,
    confirmMacroRisk: $("#confirmMacroRisk")?.checked || false,
    selectedMacros: macroSequencePayload().map((item) => ({ ...item, selected: true })),
    autoSearchOmml: $("#autoSearchOmml")?.checked ?? state.settings.autoSearchOmml,
    enableMathTypeFormatting: $("#enableMathTypeFormatting")?.checked ?? state.settings.enableMathTypeFormatting,
    allowExternalMathpixUpload: $("#allowExternalMathpixUpload")?.checked ?? state.settings.allowExternalMathpixUpload,
    waitForMathpix: $("#waitForMathpix")?.checked ?? state.settings.waitForMathpix,
    word: {
      maxCharsPerSlide: Number($("#wordMaxCharsPerSlide")?.value || 320),
      autoPagination: $("#wordAutoPagination")?.checked ?? true,
      generateToc: $("#wordGenerateToc")?.checked ?? true,
      retainImages: $("#wordRetainImages")?.checked ?? state.settings.retainImages,
      retainTables: $("#wordRetainTables")?.checked ?? state.settings.retainTables,
      retainFormulas: $("#wordRetainFormulas")?.checked ?? true,
      convertOmmlFirst: $("#wordConvertOmmlFirst")?.checked ?? true,
      applyTemplate: $("#wordApplyTemplate")?.checked ?? true,
      generateNotes: $("#wordGenerateNotes")?.checked ?? false,
      autoBeautify: $("#wordAutoBeautify")?.checked ?? true,
    },
    ppt: {
      mode: $("#pptWorkspaceMode")?.value || $("#pptToWordMode")?.value || state.settings.pptToWordMode,
      generateToc: $("#pptWorkspaceGenerateToc")?.checked ?? state.settings.pptToWordGenerateToc,
      extractNotes: $("#pptWorkspaceExtractNotes")?.checked ?? true,
      retainImages: $("#pptWorkspaceExtractImages")?.checked ?? state.settings.retainImages,
      retainFormulas: $("#pptWorkspaceExtractFormulas")?.checked ?? true,
      templateName: state.settings.pptToWordTemplate || "",
      templatePath: state.settings.pptToWordTemplatePath || "",
    },
  };
}

async function saveSettings() {
  const patch = {};
  const ids = [
    "maxConcurrentTasks",
    "singleFileLimitMb",
    "minFreeDiskMb",
    "outputDirectory",
    "outputConflictStrategy",
    "duplicateFileStrategy",
    "keepOriginalFile",
    "autoOpenOutputDirectory",
    "generateReport",
    "enableTaskCompletionNotice",
    "saveHistory",
    "verboseLogging",
    "wordToPptTemplate",
    "pptToWordMode",
    "pdfPrecisionMode",
    "excelConversionRange",
    "excelFormulaMode",
    "excelSplitSheets",
    "retainImages",
    "retainTables",
    "retainHeadersFooters",
    "retainFootnotesEndnotes",
    "retainComments",
    "retainRevisions",
    "autoCleanTemp",
    "cleanupRetentionDays",
    "logRetentionDays",
    "cleanupIntervalDays",
    "logUploadEvents",
    "logConversionEvents",
    "logFormulaEvents",
    "logOmmlEvents",
    "logMacroEvents",
    "logImageEvents",
    "logErrorEvents",
    "logExportFormat",
    "allowCloudSync",
    "ocrLanguage",
    "enableTextOcr",
    "enableFormulaOcr",
    "enableTableOcr",
    "ocrPrecisionMode",
    "ocrSpeedMode",
    "smallImageMaxWidth",
    "smallImageMaxHeight",
    "smallImageMaxArea",
    "includeHeaderFooterImages",
    "includeWatermarkImages",
    "includeTransparentImages",
    "includeDuplicateImages",
    "imageExportFormat",
    "localClientEnabled",
    "localClientPlatform",
    "localApiHost",
    "localApiPort",
    "localSecurityToken",
    "allowWebLaunchLocalClient",
    "allowTaskStatusCloudSync",
    "exposeLocalPaths",
    "mathtypeCompatibilityMode",
    "sensitiveFilesPreferLocal",
    "pdfToWordEngine",
    "allowExternalMathpixUpload",
    "waitForMathpix",
    "mathpixAppIdEnv",
    "mathpixAppKeyEnv",
    "mathpixPollTimeoutSeconds",
    "enableOmmlPrecheck",
    "autoSearchOmml",
    "allowManualOmml",
    "manualOmmlPath",
    "ommlSearchPaths",
    "ommlSearchMaxFiles",
    "enableMacroDetection",
    "enableMacroExecution",
    "allowDocumentMacros",
    "allowTemplateMacros",
    "allowLocalMacroLibrary",
    "macroWhitelistOnly",
    "macroWhitelist",
    "macroTimeoutSeconds",
    "allowBatchMacroExecution",
    "enableMathTypeFormatting",
    "formulaOutputFormat",
    "formulaFont",
    "formulaFontSize",
    "formulaFormatScope",
    "formulaAlignment",
    "formulaVariableStyle",
    "formulaFunctionStyle",
    "formulaScriptScale",
    "formulaFractionStyle",
    "formulaRadicalStyle",
    "formulaMatrixSpacing",
    "formulaGreekStyle",
    "formulaInlineBaseline",
    "formulaDisplaySpacing",
    "formulaNumbering",
    "formulaConfidenceThreshold",
    "lowConfidenceFormulaStrategy",
    "keepFormulaImages",
    "generateFormulaReport",
    "ommlCopyStrategy",
  ];
  ids.forEach((id) => {
    const node = $(`#${id}`);
    if (!node) return;
    patch[id] = nodeValue(node);
  });
  const submittedSecurityToken = String(patch.localSecurityToken || "").trim();
  if (!submittedSecurityToken && state.settings.localSecurityTokenConfigured) {
    delete patch.localSecurityToken;
  }
  patch.macroBackup = $("#macroBackupSettings") ? nodeValue($("#macroBackupSettings")) : ($("#macroBackup")?.checked ?? state.settings.macroBackup);
  patch.macroFailureStrategy = $("#macroFailureStrategySettings")?.value || $("#macroFailureStrategy")?.value || state.settings.macroFailureStrategy;
  const result = await api("/api/settings", { method: "PUT", body: JSON.stringify(patch) });
  if (submittedSecurityToken) rememberSecurityToken(submittedSecurityToken);
  state.settings = result.settings;
  rememberResponseSecurityToken(state.settings);
  const [capabilities, profile, installPlan, localClientManifest, localUploadQueue] = await Promise.all([
    api("/api/capabilities"),
    api("/api/install-profile"),
    api("/api/install-plan"),
    api("/api/local-client/manifest"),
    api("/api/local-client/uploads"),
  ]);
  state.capabilities = capabilities.capabilities;
  state.installProfile = profile.installProfile;
  state.installPlan = installPlan.installPlan;
  state.localClientManifest = localClientManifest.localClientManifest;
  state.localUploadQueue = localUploadQueue.uploadQueue;
  await refreshLocalReadiness();
  $("#localStatus").textContent = `${state.capabilities.host}:${state.capabilities.port}`;
  toast("设置已保存");
  syncSettingsToForm();
  render();
}

function syncSettingsToForm() {
  Object.entries(state.settings).forEach(([key, value]) => {
    const node = $(`#${key}`);
    if (node) setNodeValue(node, value);
  });
  const tokenNode = $("#localSecurityToken");
  if (tokenNode) {
    tokenNode.placeholder = state.settings.localSecurityTokenConfigured ? "已配置，留空则保持不变" : "留空则不启用";
    if (state.localToken) setNodeValue(tokenNode, state.localToken);
  }
  if ($("#macroBackupSettings")) setNodeValue($("#macroBackupSettings"), state.settings.macroBackup);
  if ($("#macroFailureStrategySettings")) $("#macroFailureStrategySettings").value = state.settings.macroFailureStrategy || "跳过";
  if ($("#macroBackup")) setNodeValue($("#macroBackup"), state.settings.macroBackup);
  if ($("#macroFailureStrategy")) $("#macroFailureStrategy").value = state.settings.macroFailureStrategy || "跳过";
  if ($("#pptWorkspaceMode")) $("#pptWorkspaceMode").value = state.settings.pptToWordMode || "逐页讲义模式";
  if ($("#pptWorkspaceGenerateToc")) setNodeValue($("#pptWorkspaceGenerateToc"), state.settings.pptToWordGenerateToc ?? true);
  if ($("#pptWorkspaceExtractImages")) setNodeValue($("#pptWorkspaceExtractImages"), state.settings.retainImages ?? true);
}

async function previewInstallProfile() {
  const value = $("#localClientPlatform")?.value || "auto";
  try {
    const [profile, installPlan] = await Promise.all([
      api(`/api/install-profile?platform=${encodeURIComponent(value)}`),
      api(`/api/install-plan?platform=${encodeURIComponent(value)}`),
    ]);
    state.installProfile = profile.installProfile;
    state.installPlan = installPlan.installPlan;
    renderInstallProfile();
    renderCapabilities();
  } catch (error) {
    toast(error.message || "安装画像读取失败");
  }
}

function switchView(view) {
  state.view = view;
  $$(".nav-item").forEach((button) => button.classList.toggle("active", button.dataset.view === view));
  $$(".view").forEach((panel) => {
    const panelName = panel.dataset.panel;
    panel.classList.toggle("view-active", panelName === view || (view === "settings" && panelName === "admin"));
  });
  const [title, subtitle] = viewMeta[view] || viewMeta.workspace;
  $("#viewTitle").textContent = title;
  $("#viewSubtitle").textContent = subtitle;
}

function openMenuTarget(button) {
  const view = button.dataset.menuView || "workspace";
  const task = button.dataset.menuTask || "";
  const anchor = button.dataset.menuAnchor || "";
  switchView(view);
  if (task && $("#taskTypeSelect")) {
    $("#taskTypeSelect").value = task;
    renderRouteAdvisor();
    renderTaskPlanner();
  }
  if (anchor) {
    window.setTimeout(() => document.getElementById(anchor)?.scrollIntoView({ block: "start" }), 0);
  }
}

function localPayloadUrl(taskId) {
  return absoluteUrl(withToken(`/api/tasks/${encodeURIComponent(taskId)}/local-payload`));
}

function replaceTask(task) {
  if (!task?.id) return;
  const index = state.tasks.findIndex((item) => item.id === task.id);
  if (index >= 0) {
    state.tasks.splice(index, 1, task);
  } else {
    state.tasks.unshift(task);
  }
}

async function launchLocalClient(taskId) {
  if (!taskId) return;
  if (!state.settings.allowWebLaunchLocalClient) {
    toast("请先授权网页唤起本地客户端");
    return;
  }
  if (!hasLocalSecurityCredential()) {
    toast("请先配置本地安全令牌");
    return;
  }
  try {
    const result = await api(`/api/tasks/${encodeURIComponent(taskId)}/local-launch`, {
      method: "POST",
      body: JSON.stringify({ api_origin: window.location.origin }),
    });
    if (result.task) replaceTask(result.task);
    if (result.localClientManifest) state.localClientManifest = result.localClientManifest;
    const url = launchUrlWithToken(result.launchRequest, taskId);
    window.location.href = url;
    toast("已请求启动本地客户端");
    await refreshLocalReadiness(taskId);
    renderLocalHandoff();
  } catch (error) {
    toast(error.message || "启动请求失败");
  }
}

function launchUrlWithToken(launchRequest, taskId) {
  const fallback = `k12-local://open?task_id=${encodeURIComponent(taskId)}`;
  try {
    const url = new URL(launchRequest?.protocol_url || fallback);
    url.searchParams.set("task_id", taskId);
    if (launchRequest?.id) url.searchParams.set("request_id", launchRequest.id);
    url.searchParams.set("api", window.location.origin);
    url.searchParams.set("payload_url", localPayloadUrl(taskId));
    return url.toString();
  } catch {
    return `${fallback}&payload_url=${encodeURIComponent(localPayloadUrl(taskId))}`;
  }
}

async function copyLocalPayload(taskId) {
  if (!taskId) return;
  const url = localPayloadUrl(taskId);
  try {
    await navigator.clipboard.writeText(url);
    toast("载荷地址已复制");
  } catch {
    window.prompt("复制本地任务载荷地址", url);
  }
}

function bindEvents() {
  $("#selectFilesButton").addEventListener("click", () => $("#fileInput").click());
  $("#selectFolderButton").addEventListener("click", () => $("#folderInput").click());
  $("#fileInput").addEventListener("change", (event) => importFiles(event.target.files));
  $("#folderInput").addEventListener("change", (event) => importFiles(event.target.files));
  $("#refreshButton").addEventListener("click", loadAll);
  $("#createTaskButton").addEventListener("click", () => createTask());
  $("#createBatchButton").addEventListener("click", () => createTask("batch_process"));
  $("#pauseActiveTaskButton")?.addEventListener("click", () => pauseActiveTask());
  $("#stopActiveTaskButton")?.addEventListener("click", () => stopActiveTask());
  $("#exportLatestReportButton")?.addEventListener("click", exportLatestReport);
  $("#clearLogVisualButton")?.addEventListener("click", () => {
    const target = $("#workspaceLogList");
    if (target) target.innerHTML = `<div class="log-item"><small>日志显示已清空，刷新后可重新载入</small></div>`;
  });
  $("#downloadSelectedButton").addEventListener("click", downloadSelectedFiles);
  $("#downloadFilteredButton").addEventListener("click", downloadFilteredFiles);
  $("#clearSelectionButton").addEventListener("click", () => {
    state.selectedFiles.clear();
    render();
  });
  $("#saveSettingsButton").addEventListener("click", saveSettings);
  $("#localClientPlatform").addEventListener("change", previewInstallProfile);
  $("#saveMacroTemplateButton").addEventListener("click", saveMacroTemplate);
  $("#loadMacroReportSequenceButton")?.addEventListener("click", loadMacroSequenceFromLatestReport);
  $("#loadMacroTemplateButton").addEventListener("click", loadMacroTemplate);
  $("#deleteMacroTemplateButton").addEventListener("click", deleteMacroTemplate);
  $("#saveUserButton").addEventListener("click", saveUser);
  $("#saveTemplateButton").addEventListener("click", saveTemplate);
  $("#selectWordFilesButton")?.addEventListener("click", selectWordFiles);
  $("#wordPreviewSelectedButton")?.addEventListener("click", previewSelectedWord);
  $("#wordExportPptButton")?.addEventListener("click", startWordToPpt);
  $("#selectPdfFilesButton")?.addEventListener("click", selectPdfFiles);
  $("#pdfSettingsShortcutButton")?.addEventListener("click", openPdfSettings);
  $("#pdfExportWordButton")?.addEventListener("click", startPdfToWord);
  $("#selectPptFilesButton")?.addEventListener("click", selectPptFiles);
  $("#pptPreviewSelectedButton")?.addEventListener("click", previewSelectedPpt);
  $("#pptExportWordButton")?.addEventListener("click", startPptToWord);
  $("#macroSearch").addEventListener("input", renderMacros);
  $("#macroSourceFilter").addEventListener("change", renderMacros);
  $("#macroPurposeFilter").addEventListener("change", renderMacros);
  $("#macroRecentFilter").addEventListener("change", renderMacros);
  $("#imageKindFilter").addEventListener("change", renderImages);
  $("#imageSourceFilter").addEventListener("change", renderImages);
  $("#imageDuplicateFilter").addEventListener("change", renderImages);
  $("#imageLocationFilter").addEventListener("input", renderImages);
  $("#imageMaxAreaFilter").addEventListener("input", renderImages);
  $("#exportImagesButton").addEventListener("click", exportLatestImages);
  $("#exportImageManifestButton").addEventListener("click", exportImageManifest);
  $("#exportLogsButton").addEventListener("click", exportLogs);
  $("#clearTaskLogFilterButton")?.addEventListener("click", clearTaskLogFilter);
  $("#cleanupMaintenanceButton").addEventListener("click", cleanupMaintenance);
  $("#formulaSourceFilter").addEventListener("change", renderFormulaList);
  $("#formulaBulkConfirm").addEventListener("click", bulkConfirmFormulas);
  $("#exportFormulasButton").addEventListener("click", exportFormulas);
  $("#exportFormulaSheetButton").addEventListener("click", exportFormulaSheet);
  $("#exportOmmlFailuresButton")?.addEventListener("click", exportOmmlFailures);
  $("#exportMacroFailuresButton")?.addEventListener("click", exportMacroFailures);
  $("#saveLayoutAnnotationButton")?.addEventListener("click", saveLayoutAnnotation);
  $("#taskTypeSelect").addEventListener("change", () => {
    renderRouteAdvisor();
    renderTaskPlanner();
  });
  $("#executeModeSelect").addEventListener("change", renderRouteAdvisor);

  const dropZone = $("#dropZone");
  ["dragenter", "dragover"].forEach((name) => {
    dropZone.addEventListener(name, (event) => {
      event.preventDefault();
      dropZone.classList.add("dragging");
    });
  });
  ["dragleave", "drop"].forEach((name) => {
    dropZone.addEventListener(name, () => dropZone.classList.remove("dragging"));
  });
  dropZone.addEventListener("drop", async (event) => {
    event.preventDefault();
    try {
      await importDroppedItems(event.dataTransfer);
    } catch (error) {
      toast(error.message || "拖拽导入失败");
    }
  });

  $$(".nav-item").forEach((button) => button.addEventListener("click", () => switchView(button.dataset.view)));
  $$("[data-view-shortcut]").forEach((button) => button.addEventListener("click", () => switchView(button.dataset.viewShortcut)));
  $$(".segment").forEach((button) => {
    button.addEventListener("click", () => {
      state.filter = button.dataset.filter;
      renderQueueTabs();
      renderFiles();
    });
  });
  $$(".quick-task").forEach((button) => button.addEventListener("click", () => createTask(button.dataset.task)));

  document.addEventListener("change", (event) => {
    const target = event.target;
    if (target.matches(".file-check")) {
      target.checked ? state.selectedFiles.add(target.dataset.id) : state.selectedFiles.delete(target.dataset.id);
      render();
    }
    if (target.id === "selectAllFiles") {
      filteredFiles().forEach((file) => (target.checked ? state.selectedFiles.add(file.id) : state.selectedFiles.delete(file.id)));
      render();
    }
  });

  document.addEventListener("input", (event) => {
    const target = event.target;
    if (target.id === "previewSearchInput") {
      state.previewSearch = target.value || "";
      state.previewPageOffset = 0;
      renderFilePreview();
    }
  });

  if (window.PointerEvent) {
    document.addEventListener("pointerdown", startPlannerPointerDrag);
    document.addEventListener("pointermove", updatePlannerPointerDrag);
    document.addEventListener("pointerup", finishPlannerPointerDrag);
    document.addEventListener("pointercancel", cancelPlannerPointerDrag);
  } else {
    document.addEventListener("mousedown", startPlannerPointerDrag);
    document.addEventListener("mousemove", updatePlannerPointerDrag);
    document.addEventListener("mouseup", finishPlannerPointerDrag);
  }

  document.addEventListener("dragstart", (event) => {
    const macroItem = event.target.closest(".macro-order-item");
    if (macroItem) {
      event.dataTransfer?.setData("text/plain", macroItem.dataset.macro || "");
      event.dataTransfer?.setData("application/x-k12-macro-order", macroItem.dataset.macro || "");
      macroItem.classList.add("dragging");
      return;
    }
    const plannerHandle = event.target.closest(".planner-drag-handle");
    const item = plannerHandle?.closest(".planner-item");
    if (!item) return;
    event.dataTransfer?.setData("text/plain", item.dataset.task || "");
    event.dataTransfer?.setData("application/x-k12-planner-task", item.dataset.task || "");
    item.classList.add("dragging");
  });

  document.addEventListener("dragover", (event) => {
    if (event.target.closest(".macro-order-item")) event.preventDefault();
    const plannerTarget = plannerDropTargetAt(event.clientX, event.clientY);
    if (plannerTarget) {
      event.preventDefault();
      markPlannerDropTarget(plannerTarget.item, plannerTarget.placement);
    }
  });

  document.addEventListener("drop", (event) => {
    resetPlannerDropTargets();
    const macroItem = event.target.closest(".macro-order-item");
    if (macroItem) {
      event.preventDefault();
      const sourceMacro = event.dataTransfer?.getData("application/x-k12-macro-order") || event.dataTransfer?.getData("text/plain") || "";
      moveMacroBefore(sourceMacro, macroItem.dataset.macro || "");
      return;
    }
    const plannerTarget = plannerDropTargetAt(event.clientX, event.clientY);
    if (!plannerTarget) return;
    event.preventDefault();
    const sourceTask = event.dataTransfer?.getData("application/x-k12-planner-task") || event.dataTransfer?.getData("text/plain") || "";
    movePlannerTaskBefore(sourceTask, plannerTarget.item.dataset.task || "", plannerTarget.placement);
  });

  document.addEventListener("dragend", () => {
    resetPlannerDropTargets();
    $$(".planner-item.dragging").forEach((item) => item.classList.remove("dragging"));
    $$(".macro-order-item.dragging").forEach((item) => item.classList.remove("dragging"));
  });
  document.addEventListener("keydown", handlePlannerHandleKeydown);

  document.addEventListener("click", async (event) => {
    const button = event.target.closest("button");
    if (!button) return;
    if (button.matches(".menu-group-head, .menu-leaf")) {
      openMenuTarget(button);
      return;
    }
    if (button.matches(".launch-local-client")) {
      await launchLocalClient(button.dataset.id || "");
      return;
    }
    if (button.matches(".copy-local-payload")) {
      await copyLocalPayload(button.dataset.id || "");
      return;
    }
    if (button.matches(".upload-manifest")) {
      await loadLocalUploadManifest(button.dataset.endpoint || "");
      return;
    }
    if (button.matches(".risk-jump")) {
      switchView("settings");
      const anchor = button.dataset.riskAnchor || "";
      if (anchor) window.setTimeout(() => document.getElementById(anchor)?.scrollIntoView({ block: "start" }), 0);
      return;
    }
    if (button.matches(".exception-jump")) {
      switchView("settings");
      const anchor = button.dataset.exceptionAnchor || "";
      if (anchor) window.setTimeout(() => document.getElementById(anchor)?.scrollIntoView({ block: "start" }), 0);
      return;
    }
    if (button.matches(".acceptance-jump")) {
      const view = button.dataset.acceptanceView || "workspace";
      switchView(view);
      const anchor = button.dataset.acceptanceAnchor || "";
      if (anchor) window.setTimeout(() => document.getElementById(anchor)?.scrollIntoView({ block: "start" }), 0);
      return;
    }
    if (button.matches(".security-jump")) {
      switchView("settings");
      const anchor = button.dataset.securityAnchor || "";
      if (anchor) window.setTimeout(() => document.getElementById(anchor)?.scrollIntoView({ block: "start" }), 0);
      return;
    }
    if (button.matches(".compatibility-jump")) {
      switchView("settings");
      const anchor = button.dataset.compatibilityAnchor || "";
      if (anchor) window.setTimeout(() => document.getElementById(anchor)?.scrollIntoView({ block: "start" }), 0);
      return;
    }
    if (button.matches(".file-inspect")) {
      state.activeFileId = button.dataset.id;
      renderFiles();
      renderFileDetail();
    }
    if (button.matches(".file-preview")) {
      await loadFilePreview(button.dataset.id);
    }
    if (button.matches(".file-convert")) {
      const file = state.files.find((item) => item.id === button.dataset.id);
      await createFileTask(button.dataset.id, defaultConversionTask(file));
    }
    if (button.matches(".file-scan")) {
      const file = state.files.find((item) => item.id === button.dataset.id);
      await createFileTask(button.dataset.id, defaultSearchTask(file));
    }
    if (button.matches(".file-reports")) {
      await openFileReports(button.dataset.id);
    }
    if (button.matches(".word-select")) {
      state.selectedFiles = new Set([button.dataset.id]);
      state.activeFileId = button.dataset.id;
      if ($("#taskTypeSelect")) $("#taskTypeSelect").value = "word_to_ppt";
      render();
    }
    if (button.matches(".pdf-select")) {
      state.selectedFiles = new Set([button.dataset.id]);
      state.activeFileId = button.dataset.id;
      if ($("#taskTypeSelect")) $("#taskTypeSelect").value = "pdf_to_word";
      render();
    }
    if (button.matches(".ppt-select")) {
      state.selectedFiles = new Set([button.dataset.id]);
      state.activeFileId = button.dataset.id;
      if ($("#taskTypeSelect")) $("#taskTypeSelect").value = "ppt_to_word";
      render();
    }
    if (button.matches(".preview-page-prev")) {
      state.previewPageOffset = Math.max(0, state.previewPageOffset - 4);
      renderFilePreview();
    }
    if (button.matches(".preview-page-next")) {
      const total = filterPreviewPages(state.filePreview?.pages || []).length;
      state.previewPageOffset = Math.min(Math.max(0, total - 1), state.previewPageOffset + 4);
      renderFilePreview();
    }
    if (button.matches(".preview-zoom-out")) {
      state.previewZoom = Math.max(0.8, Number((state.previewZoom - 0.1).toFixed(1)));
      renderFilePreview();
    }
    if (button.matches(".preview-zoom-in")) {
      state.previewZoom = Math.min(1.4, Number((state.previewZoom + 0.1).toFixed(1)));
      renderFilePreview();
    }
    if (button.matches(".preview-jump-chip")) {
      jumpPreviewTo(button.dataset.term || "");
    }
    if (button.matches(".file-password")) {
      await setFilePassword(button.dataset.id);
    }
    if (button.matches(".file-download")) {
      window.open(withToken(`/api/files/${button.dataset.id}/download`), "_blank", "noopener");
    }
    if (button.matches(".file-replace")) {
      replaceFile(button.dataset.id);
    }
    if (button.matches(".file-delete")) {
      if (!window.confirm("确定删除这个文件记录吗？上传缓存会被移除，外部原始文件不会被删除。")) return;
      await api(`/api/files/${button.dataset.id}`, { method: "DELETE" });
      state.selectedFiles.delete(button.dataset.id);
      if (state.activeFileId === button.dataset.id) state.activeFileId = "";
      if (state.filePreview?.file_id === button.dataset.id) state.filePreview = null;
      await loadAll();
    }
    if (button.matches(".task-retry")) {
      await api(`/api/tasks/${button.dataset.id}/retry`, { method: "POST", body: "{}" });
      await loadAll();
    }
    if (button.matches(".batch-skip-file")) {
      await api(`/api/tasks/${button.dataset.task}/skip-file`, { method: "POST", body: JSON.stringify({ file_id: button.dataset.file }) });
      toast("已跳过该批量失败文件");
      await loadAll();
    }
    if (button.matches(".task-pause")) {
      await api(`/api/tasks/${button.dataset.id}/pause`, { method: "POST", body: "{}" });
      await loadAll();
    }
    if (button.matches(".task-resume")) {
      await api(`/api/tasks/${button.dataset.id}/resume`, { method: "POST", body: "{}" });
      await loadAll();
    }
    if (button.matches(".task-restore")) {
      if (window.confirm("确定用宏执行前备份覆盖当前文件吗？")) {
        const result = await api(`/api/tasks/${button.dataset.id}/restore-backups`, { method: "POST", body: "{}" });
        toast(`已恢复 ${result.restored_count || 0} 个备份`);
        await loadAll();
      }
    }
    if (button.matches(".task-download")) {
      window.open(withToken(`/api/tasks/${button.dataset.id}/download`), "_blank", "noopener");
    }
    if (button.matches(".task-detail")) {
      state.activeTaskDetailId = button.dataset.id;
      renderTaskDetail();
      $("#taskDetailPanel")?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
    if (button.matches(".task-notice-ack")) {
      await api(`/api/tasks/${button.dataset.id}/completion-notice/ack`, { method: "POST", body: "{}" });
      toast("任务完成提醒已确认");
      await loadAll();
    }
    if (button.matches(".task-log")) {
      await viewTaskLogs(button.dataset.id);
    }
    if (button.matches(".task-cancel")) {
      await api(`/api/tasks/${button.dataset.id}/cancel`, { method: "POST", body: "{}" });
      await loadAll();
    }
    if (button.matches(".report-download")) {
      window.open(withToken(`/api/reports/${button.dataset.id}/download?format=${button.dataset.format || "html"}`), "_blank", "noopener");
    }
    if (button.matches(".failure-download")) {
      window.open(withToken(`/api/reports/${button.dataset.id}/failures.csv`), "_blank", "noopener");
    }
    if (button.matches(".comparison-load")) {
      await loadComparison(button.dataset.id);
    }
    if (button.matches(".report-delete")) {
      if (window.confirm("确定删除这份报告及报告缓存吗？源文件和任务不会删除。")) {
        await api(`/api/reports/${button.dataset.id}`, { method: "DELETE" });
        await loadAll();
      }
    }
    if (button.matches(".image-open")) {
      if (button.dataset.url) window.open(withToken(button.dataset.url), "_blank", "noopener");
    }
    if (button.matches(".image-detail")) {
      state.activeImageId = button.dataset.id;
      renderImages();
    }
    if (button.matches(".image-download")) {
      if (button.dataset.url) window.open(withToken(button.dataset.url), "_blank", "noopener");
    }
    if (button.matches(".image-reexport")) {
      await reexportImage(button.dataset.report, button.dataset.id);
    }
    if (button.matches(".image-locate")) {
      await locateImageInSource(button.dataset.file, button.dataset.location, button.dataset.report, button.dataset.id);
    }
    if (button.matches(".image-confirm-type")) {
      await confirmImageType(button.dataset.report, button.dataset.id, button.dataset.label);
    }
    if (button.matches(".image-false-positive")) {
      await markImageFalsePositive(button.dataset.report, button.dataset.id);
    }
    if (button.matches(".image-delete-mark")) {
      await annotateImage(button.dataset.report, button.dataset.id, "删除待处理", "用户要求删除该图片，等待本地客户端改写源文档");
    }
    if (button.matches(".image-replace-mark")) {
      await markImageReplacement(button.dataset.report, button.dataset.id);
    }
    if (button.matches(".formula-confirm")) {
      await annotateFormula(button.dataset.report, button.dataset.id, "已确认");
    }
    if (button.matches(".formula-select")) {
      state.activeFormulaId = button.dataset.id;
      renderFormulaList();
    }
    if (button.matches(".formula-skip")) {
      await annotateFormula(button.dataset.report, button.dataset.id, "跳过");
    }
    if (button.matches(".formula-rerecognize")) {
      await annotateFormula(button.dataset.report, button.dataset.id, "重新识别");
    }
    if (button.matches(".formula-save")) {
      await annotateFormula(button.dataset.report, button.dataset.id, "已修正");
    }
    if (button.matches(".omml-keep")) {
      await annotateOmml(button.dataset.report, button.dataset.id, "保留OMML", "用户选择保留 Word 自带公式");
    }
    if (button.matches(".omml-retry")) {
      await annotateOmml(button.dataset.report, button.dataset.id, "重新转换", "用户要求重新执行 OMML 转 MathType");
    }
    if (button.matches(".omml-fixed")) {
      await annotateOmml(button.dataset.report, button.dataset.id, "已修复", "用户确认 OMML 依赖问题已修复");
    }
    if (button.matches(".omml-manual")) {
      const manualPath = window.prompt("请输入 OMML2MML.XSL 或等效依赖文件的本地路径", "");
      if (manualPath === null) return;
      const trimmedPath = manualPath.trim();
      if (!trimmedPath) {
        toast("请填写 OMML 依赖文件路径");
        return;
      }
      await annotateOmml(button.dataset.report, button.dataset.id, "手动指定依赖", "用户手动指定 OMML 依赖文件，等待本地客户端复核并重新转换", {
        manual_omml_path: trimmedPath,
        retry_conversion: true,
      });
    }
    if (button.matches(".planner-task")) {
      $("#taskTypeSelect").value = button.dataset.task;
      await createTask(button.dataset.task);
    }
    if (button.matches(".planner-up")) {
      movePlannerTask(button.dataset.task, -1);
    }
    if (button.matches(".planner-down")) {
      movePlannerTask(button.dataset.task, 1);
    }
    if (button.matches(".macro-detail")) {
      state.activeMacroId = button.dataset.id;
      renderMacros();
      renderMacroDetail();
    }
    if (button.matches(".macro-add")) {
      state.activeMacroId = button.dataset.id;
      if (!state.macroSequence.includes(button.dataset.id)) {
        state.macroSequence.push(button.dataset.id);
        renderMacros();
        renderMacroDetail();
      }
    }
    if (button.matches(".macro-remove")) {
      state.macroSequence = state.macroSequence.filter((id) => id !== button.dataset.id);
      renderMacros();
      renderMacroDetail();
    }
    if (button.matches(".macro-rerun")) {
      await rerunMacroFromReport(button.dataset.report, button.dataset.id, button.dataset.index);
    }
    if (button.matches(".macro-up")) {
      moveMacro(button.dataset.id, -1);
      renderMacros();
    }
    if (button.matches(".macro-down")) {
      moveMacro(button.dataset.id, 1);
      renderMacros();
    }
    if (button.matches(".user-activate")) {
      await activateUser(button.dataset.id);
    }
    if (button.matches(".user-delete")) {
      if (window.confirm("确定删除这个用户吗？")) {
        await deleteUser(button.dataset.id);
      }
    }
    if (button.matches(".template-delete")) {
      await deleteTemplate(button.dataset.id);
    }
    if (button.matches(".authorization-toggle")) {
      await toggleAuthorization(button.dataset.key);
    }
  });
}

function recommendedTasks(files) {
  if (!files.length) return ["word_to_ppt", "formula_precheck", "omml_to_mathtype", "mathtype_format", "pdf_to_word", "excel_to_pdf"];
  const tasks = [];
  const hasType = (type) => files.some((file) => file.file_type === type);
  const any = (predicate) => files.some(predicate);
  if (hasType("Word")) tasks.push("word_to_ppt", "formula_precheck");
  if (any((file) => file.has_omml)) tasks.push("omml_to_mathtype");
  if (any((file) => file.has_mathtype || file.has_formula)) tasks.push("mathtype_format");
  if (any((file) => file.has_macro)) tasks.push("macro_sequence");
  if (hasType("PPT")) tasks.push("ppt_to_word");
  if (hasType("PDF")) tasks.push("pdf_to_word");
  if (hasType("Excel")) tasks.push("excel_to_pdf", "excel_to_word", "excel_to_ppt");
  if (any((file) => file.has_small_image)) tasks.push("small_image_scan");
  if (files.length > 1) tasks.push("batch_process");
  return [...new Set(tasks)].slice(0, 6);
}

function orderedPlannerTasks(files = contextualFiles()) {
  const recommended = recommendedTasks(files);
  const ordered = state.plannerOrder.filter((taskType) => recommended.includes(taskType));
  const merged = [...ordered, ...recommended.filter((taskType) => !ordered.includes(taskType))];
  state.plannerOrder = merged;
  return merged;
}

function defaultPlannerTask(files = contextualFiles()) {
  return orderedPlannerTasks(files)[0] || "word_to_ppt";
}

function defaultConversionTask(file) {
  if (!file) return "word_to_ppt";
  if (file.file_type === "Word") return "word_to_ppt";
  if (file.file_type === "PPT") return "ppt_to_word";
  if (file.file_type === "PDF") return "pdf_to_word";
  if (file.file_type === "Excel") return "excel_to_pdf";
  if (file.file_type === "图片") return "small_image_scan";
  return "batch_process";
}

function defaultSearchTask(file) {
  if (!file) return "formula_precheck";
  if (file.has_omml || file.missing_omml_dependency) return "omml_to_mathtype";
  if (file.has_image || file.has_small_image || ["Word", "PPT", "Excel", "PDF", "图片"].includes(file.file_type)) return "small_image_scan";
  return "formula_precheck";
}

function movePlannerTask(taskType, direction) {
  const tasks = orderedPlannerTasks(contextualFiles());
  const index = tasks.indexOf(taskType);
  const target = index + direction;
  if (index < 0 || target < 0 || target >= tasks.length) return;
  [tasks[index], tasks[target]] = [tasks[target], tasks[index]];
  state.plannerOrder = tasks;
  renderTaskPlanner();
  renderRouteAdvisor();
}

function movePlannerTaskToIndex(taskType, targetIndex) {
  const tasks = orderedPlannerTasks(contextualFiles());
  const index = tasks.indexOf(taskType);
  if (index < 0 || targetIndex < 0 || targetIndex >= tasks.length || index === targetIndex) return;
  const [task] = tasks.splice(index, 1);
  tasks.splice(targetIndex, 0, task);
  state.plannerOrder = tasks;
  renderTaskPlanner();
  renderRouteAdvisor();
}

function movePlannerTaskBefore(sourceTask, targetTask, placement = "before") {
  if (!sourceTask || !targetTask || sourceTask === targetTask) return;
  const current = orderedPlannerTasks(contextualFiles());
  if (!current.includes(sourceTask) || !current.includes(targetTask)) return;
  const tasks = current.filter((taskType) => taskType !== sourceTask);
  const targetIndex = tasks.indexOf(targetTask);
  if (targetIndex < 0) return;
  const insertIndex = placement === "after" ? targetIndex + 1 : targetIndex;
  tasks.splice(insertIndex, 0, sourceTask);
  state.plannerOrder = tasks;
  renderTaskPlanner();
  renderRouteAdvisor();
}

function handlePlannerHandleKeydown(event) {
  if (!(event.target instanceof Element)) return;
  const handle = event.target.closest(".planner-drag-handle");
  const taskType = handle?.dataset.task || "";
  if (!taskType) return;
  const tasks = orderedPlannerTasks(contextualFiles());
  if (!tasks.includes(taskType)) return;
  if (event.key === "ArrowUp") movePlannerTask(taskType, -1);
  else if (event.key === "ArrowDown") movePlannerTask(taskType, 1);
  else if (event.key === "Home") movePlannerTaskToIndex(taskType, 0);
  else if (event.key === "End") movePlannerTaskToIndex(taskType, tasks.length - 1);
  else return;
  event.preventDefault();
  focusPlannerHandle(taskType);
}

function focusPlannerHandle(taskType) {
  window.requestAnimationFrame(() => {
    $$(".planner-drag-handle").find((button) => button.dataset.task === taskType)?.focus();
  });
}

function startPlannerPointerDrag(event) {
  const handle = event.target.closest(".planner-drag-handle");
  const item = handle?.closest(".planner-item");
  if (!item) return;
  if ("button" in event && event.button !== 0) return;
  event.preventDefault();
  handle.setPointerCapture?.(event.pointerId);
  plannerPointerDrag = {
    active: false,
    handle,
    item,
    pointerId: event.pointerId,
    startY: event.clientY,
    task: item.dataset.task || "",
  };
}

function updatePlannerPointerDrag(event) {
  if (!plannerPointerDrag || !samePlannerPointer(event)) return;
  const distance = Math.abs(event.clientY - plannerPointerDrag.startY);
  if (!plannerPointerDrag.active && distance < 4) return;
  plannerPointerDrag.active = true;
  plannerPointerDrag.item.classList.add("dragging");
  event.preventDefault();
  const target = plannerDropTargetAt(event.clientX, event.clientY);
  if (target) markPlannerDropTarget(target.item, target.placement);
}

function finishPlannerPointerDrag(event) {
  if (!plannerPointerDrag || !samePlannerPointer(event)) return;
  const drag = plannerPointerDrag;
  if (drag.active) {
    event.preventDefault();
    const target = plannerDropTargetAt(event.clientX, event.clientY);
    if (target) movePlannerTaskBefore(drag.task, target.item.dataset.task || "", target.placement);
  }
  drag.handle.releasePointerCapture?.(drag.pointerId);
  cancelPlannerPointerDrag();
}

function cancelPlannerPointerDrag() {
  resetPlannerDropTargets();
  $$(".planner-item.dragging").forEach((item) => item.classList.remove("dragging"));
  plannerPointerDrag = null;
}

function samePlannerPointer(event) {
  return plannerPointerDrag.pointerId === undefined || event.pointerId === undefined || event.pointerId === plannerPointerDrag.pointerId;
}

function plannerDropItemAt(clientX, clientY) {
  return document.elementFromPoint(clientX, clientY)?.closest(".planner-item");
}

function plannerDropTargetAt(clientX, clientY) {
  const directItem = plannerDropItemAt(clientX, clientY);
  if (directItem) return { item: directItem, placement: plannerDropPlacement(directItem, clientY) };
  const list = $("#taskPlanner");
  if (!list) return null;
  const items = Array.from(list.querySelectorAll(".planner-item"));
  if (!items.length) return null;
  const rect = list.getBoundingClientRect();
  if (clientX < rect.left || clientX > rect.right || clientY < rect.top - 24 || clientY > rect.bottom + 24) return null;
  for (const item of items) {
    const itemRect = item.getBoundingClientRect();
    if (clientY < itemRect.top + itemRect.height / 2) return { item, placement: "before" };
  }
  return { item: items[items.length - 1], placement: "after" };
}

function plannerDropPlacement(item, clientY) {
  const rect = item.getBoundingClientRect();
  return clientY > rect.top + rect.height / 2 ? "after" : "before";
}

function resetPlannerDropTargets() {
  $$(".planner-item.drop-before, .planner-item.drop-after").forEach((item) => {
    item.classList.remove("drop-before", "drop-after");
  });
}

function markPlannerDropTarget(item, placementOrClientY) {
  resetPlannerDropTargets();
  const placement = typeof placementOrClientY === "string" ? placementOrClientY : plannerDropPlacement(item, placementOrClientY);
  item.classList.add(placement === "after" ? "drop-after" : "drop-before");
}

function resolveExecuteMode(taskType, files = contextualFiles()) {
  const forced = $("#executeModeSelect")?.value;
  if (forced && forced !== "auto") return forced;
  if (["formula_precheck", "omml_to_mathtype", "mathtype_format", "macro_sequence"].includes(taskType)) return "local";
  if (files.some((file) => file.has_macro || file.has_omml || file.has_mathtype)) return "hybrid";
  if (["word_to_ppt", "ppt_to_word", "small_image_scan", "batch_process"].includes(taskType)) return "hybrid";
  return taskSpecs[taskType]?.mode || "web";
}

function compatibilityIssues(taskType, files) {
  if (!files.length) return ["请先选择或上传文件"];
  const spec = taskSpecs[taskType];
  const invalid = files.filter((file) => !spec.inputs.includes(file.file_type));
  const errors = files.flatMap((file) => file.validation_errors || []);
  const issues = [];
  if (errors.length) issues.push(`存在 ${errors.length} 条文件校验问题`);
  if (taskType === "macro_sequence" && !state.macroSequence.length) issues.push("尚未选择宏，系统会使用默认宏库预览");
  if (taskType === "macro_sequence" && !state.settings.enableMacroExecution) issues.push("宏执行队列已禁用，只生成安全报告");
  if (taskType === "macro_sequence" && state.settings.macroWhitelistOnly) issues.push("仅白名单宏可进入执行队列");
  if (invalid.length && taskType !== "batch_process") issues.push(`有 ${invalid.length} 个文件类型可能不适配`);
  if (taskType === "pdf_to_word" && !state.settings.allowExternalMathpixUpload) issues.push("Mathpix 外部上传未授权时会生成授权提示");
  return issues;
}

function moveMacro(id, delta) {
  const index = state.macroSequence.indexOf(id);
  const target = index + delta;
  if (index < 0 || target < 0 || target >= state.macroSequence.length) return;
  [state.macroSequence[index], state.macroSequence[target]] = [state.macroSequence[target], state.macroSequence[index]];
}

function moveMacroBefore(sourceId, targetId) {
  if (!sourceId || !targetId || sourceId === targetId) return;
  const ordered = state.macroSequence.filter((id) => id !== sourceId);
  const targetIndex = ordered.indexOf(targetId);
  if (targetIndex < 0 || !state.macroSequence.includes(sourceId)) return;
  ordered.splice(targetIndex, 0, sourceId);
  state.macroSequence = ordered;
  renderMacroOrder();
  renderRouteAdvisor();
}

function fileStats(files) {
  return {
    Word: files.filter((file) => file.file_type === "Word").length,
    PPT: files.filter((file) => file.file_type === "PPT").length,
    PDF: files.filter((file) => file.file_type === "PDF").length,
    Excel: files.filter((file) => file.file_type === "Excel").length,
    formula: files.filter((file) => file.has_formula).length,
    mathtype: files.filter((file) => file.has_mathtype).length,
    omml: files.filter((file) => file.has_omml).length,
    macro: files.filter((file) => file.has_macro).length,
    image: files.filter((file) => file.has_image).length,
    smallImage: files.filter((file) => file.has_small_image).length,
  };
}

function taskStats() {
  return {
    pending: state.tasks.filter((task) => task.status === "待处理").length,
    running: state.tasks.filter((task) => task.status === "处理中").length,
    success: state.tasks.filter((task) => task.status === "成功").length,
    failed: state.tasks.filter((task) => task.status === "失败" || task.status === "已取消").length,
  };
}

function metricCard([label, value, detail]) {
  return `<article class="metric-card"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span><small>${escapeHtml(detail)}</small></article>`;
}

function profileList(items = []) {
  return `<ul class="profile-list">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
}

function compatibilityLabel(mode) {
  return {
    "platform-specific": "区分平台",
    "mathml-latex": "MathML/LaTeX 优先",
    "image-fallback": "图片兜底",
  }[mode] || mode || "区分平台";
}

function installContractBadgeClass(contract = {}) {
  const blockers = contract.native_handoff_blocking_reasons || [];
  if (blockers.includes("platform_mismatch") || contract.status === "平台不符") return "warn";
  if (blockers.length || contract.status === "需同平台") return "warn";
  if (contract.native_handoff_allowed || contract.status === "同平台可交接" || contract.status === "跨平台兜底") return "good";
  return "blue";
}

function syncSelectAll(files) {
  const box = $("#selectAllFiles");
  if (!box) return;
  const ids = files.map((file) => file.id);
  const checked = ids.filter((id) => state.selectedFiles.has(id)).length;
  box.checked = ids.length > 0 && checked === ids.length;
  box.indeterminate = checked > 0 && checked < ids.length;
}

function nodeValue(node) {
  if (node.type === "checkbox") return node.checked;
  if (node.type === "number") return Number(node.value);
  return node.value;
}

function setNodeValue(node, value) {
  if (node.type === "checkbox") node.checked = Boolean(value);
  else node.value = value;
}

function setSelectIfOptionExists(node, value) {
  if (!node) return;
  const option = [...node.options].find((item) => item.value === value || item.textContent === value);
  if (option) node.value = option.value;
}

function fileMetric(file) {
  return file.page_count || file.slide_count || file.sheet_count || file.archive_entry_count || "-";
}

function fileCapabilityBadges(file) {
  const badges = [];
  if (file.has_mathtype) badges.push(["MathType", "blue"]);
  if (file.has_omml) badges.push([file.missing_omml_dependency ? "OMML缺失" : "OMML", file.missing_omml_dependency ? "bad" : "warn"]);
  if (file.has_image) badges.push(["图片", "blue"]);
  if (file.has_small_image) badges.push(["微小图", "warn"]);
  if (file.encrypted) badges.push(["加密", "bad"]);
  if (!badges.length) return '<span class="badge">常规</span>';
  return `<div class="capability-badges">${badges.map(([label, level]) => `<span class="badge ${level}">${escapeHtml(label)}</span>`).join("")}</div>`;
}

function primaryValidationError(file) {
  return (file?.validation_errors || []).find(Boolean) || "";
}

function uploadValidationSummary(files) {
  const uploaded = Array.from(files || []);
  const invalid = uploaded.filter((file) => primaryValidationError(file) || file.status === "校验失败");
  if (!invalid.length) return `已上传 ${uploaded.length} 个文件`;
  const firstReason = primaryValidationError(invalid[0]) || "校验失败";
  return `已上传 ${uploaded.length} 个文件，${invalid.length} 个需处理：${firstReason}`;
}

function formulaCount(file) {
  const summary = file.content_summary || {};
  return summary.formulas || summary.formulaHints || summary.ommlFormulas || (file.has_formula ? 1 : 0) || 0;
}

function ommlCount(file) {
  const summary = file.content_summary || {};
  return summary.ommlFormulas || (file.has_omml ? 1 : 0) || 0;
}

function macroCount(file) {
  return file.has_macro ? (file.content_summary?.macros || 1) : "-";
}

function fileTypeShort(type) {
  return {
    Word: "W",
    PDF: "PDF",
    PPT: "PPT",
    Excel: "X",
    图片: "I",
    ZIP: "Z",
  }[type] || "F";
}

function fileTypeClass(type) {
  return {
    Word: "word",
    PDF: "pdf",
    PPT: "ppt",
    Excel: "excel",
    图片: "image",
    ZIP: "zip",
  }[type] || "";
}

function sourceLabel(source) {
  return {
    upload: "上传",
    archive_entry: "压缩包",
    metadata: "元数据",
  }[source] || source || "未知";
}

function formatBytes(bytes) {
  if (!bytes) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit += 1;
  }
  return `${value.toFixed(value >= 10 || unit === 0 ? 0 : 1)} ${units[unit]}`;
}

function formatTime(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString("zh-CN", { hour12: false });
}

function formatDate(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString("zh-CN", { year: "numeric", month: "2-digit", day: "2-digit" }).replace(/\//g, "-");
}

function truncateText(text, maxLength) {
  const value = String(text || "");
  if (value.length <= maxLength) return value;
  return `${value.slice(0, Math.max(0, maxLength - 1))}…`;
}

function clampNumber(value, fallback, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.max(min, Math.min(max, Math.floor(number)));
}

function toCount(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) && number > 0 ? Math.floor(number) : 0;
}

function formatDuration(task) {
  if (task?.duration_label) return task.duration_label;
  if (!task?.start_time) return "-";
  const seconds = Number(task.duration_seconds || 0);
  const safeSeconds = Number.isFinite(seconds) && seconds > 0 ? Math.floor(seconds) : 0;
  const hours = Math.floor(safeSeconds / 3600);
  const minutes = Math.floor((safeSeconds % 3600) / 60);
  const secs = safeSeconds % 60;
  if (hours) return `${hours}小时${minutes}分${secs}秒`;
  if (minutes) return `${minutes}分${secs}秒`;
  return `${secs}秒`;
}

function yesNo(value) {
  return value ? '<span class="badge good">是</span>' : '<span class="badge">否</span>';
}

function ommlMissingBadge(file) {
  if (!file.has_omml) return '<span class="badge">不涉及</span>';
  return file.missing_omml_dependency ? '<span class="badge bad">缺失</span>' : '<span class="badge good">正常</span>';
}

function imageLabel(item) {
  if (item.is_formula_like) return "疑似公式";
  if (item.is_qrcode_like) return "疑似二维码";
  if (item.is_stamp_like) return "疑似印章";
  if (item.is_signature_like) return "疑似签名";
  return "疑似图标";
}

function imageKindLabel(item) {
  return imageLabel(item).replace("疑似", "") || "图标";
}

function imageAnnotation(imageId, reportId) {
  return state.imageAnnotations.find((item) => item.image_id === imageId && (!reportId || item.report_id === reportId));
}

function imageAnnotationMuted(annotation) {
  return ["误判", "删除待处理"].includes(annotation?.status);
}

function imageAnnotationBadge(status) {
  if (status === "删除待处理") return "bad";
  if (["替换待处理", "类型已确认"].includes(status)) return "blue";
  if (status === "位置未知") return "warn";
  return "warn";
}

function layoutStatusClass(status) {
  if (status === "已校正") return "good";
  if (status === "已忽略") return "blue";
  if (status === "需本地客户端处理") return "warn";
  return "";
}

function layoutChangeSummary(item) {
  const before = item.before ? `修复前：${item.before}` : "";
  const after = item.after ? `修复后：${item.after}` : "";
  return [before, after].filter(Boolean).join("；") || "等待补充修复前后描述";
}

function formulaAnnotation(formulaId, reportId) {
  return state.formulaAnnotations.find((item) => item.formula_id === formulaId && (!reportId || item.report_id === reportId));
}

function ommlAnnotation(dependencyId, reportId) {
  return state.ommlAnnotations.find((item) => item.dependency_id === dependencyId && (!reportId || item.report_id === reportId));
}

function formulaBadgeClass(status, confidence) {
  if (["已确认", "已修正"].includes(status)) return "good";
  if (status === "跳过") return "bad";
  return confidence >= 80 ? "good" : "warn";
}

function qualityBadgeClass(status) {
  if (status === "失败") return "bad";
  if (status === "需确认") return "warn";
  return "good";
}

function imagePreview(item) {
  if (item.asset_url) {
    return `<img src="${escapeHtml(withToken(item.asset_url))}" alt="${escapeHtml(imageLabel(item))}">`;
  }
  return `${item.width}×${item.height}`;
}

function latestReport() {
  return state.reports[0];
}

function reportForTask(taskId) {
  return state.reports.find((item) => item.task_id === taskId);
}

function hasMacroBackup(taskId) {
  const report = reportForTask(taskId);
  return Boolean(report?.analysis?.macros?.some((item) => item.backup_available || item.backup_path));
}

function reportButtons(taskId) {
  const report = reportForTask(taskId);
  if (!report) return "-";
  return `<div class="download-stack">${downloadButtons(report.id)}${artifactButtons(report)}</div>`;
}

function downloadButtons(reportId) {
  return `<div class="row-actions">
    <button class="mini-button report-download" data-id="${reportId}" data-format="html" title="下载 HTML 报告">HTML</button>
    <button class="mini-button report-download" data-id="${reportId}" data-format="json" title="下载 JSON 报告">JSON</button>
    <button class="mini-button report-download" data-id="${reportId}" data-format="pdf" title="下载 PDF 报告">PDF</button>
    <button class="mini-button report-download" data-id="${reportId}" data-format="xlsx" title="下载 XLSX 报告">XLSX</button>
    <button class="mini-button report-download" data-id="${reportId}" data-format="txt" title="下载 TXT 日志">TXT</button>
    <button class="mini-button failure-download" data-id="${reportId}" title="下载失败清单 CSV">失败</button>
    <button class="mini-button comparison-load" data-id="${reportId}" title="转换前后对比">对比</button>
  </div>`;
}

function artifactButtons(report) {
  const artifacts = report.analysis?.artifacts?.filter((item) => item.status === "成功" && item.url) || [];
  if (!artifacts.length) return "";
  return `<div class="row-actions">${artifacts
    .map((item) => `<a class="mini-button" href="${escapeHtml(withToken(item.url))}" target="_blank" rel="noopener" title="下载转换输出">${escapeHtml(item.output_type.toUpperCase())}</a>`)
    .join("")}</div>`;
}

function statusClass(status) {
  if (["成功", "可用", "已格式化", "已修复"].includes(status)) return "good";
  if (["失败", "校验失败"].includes(status)) return "bad";
  if (["处理中", "待处理", "待确认", "待本地客户端执行", "已取消", "已暂停", "已中断", "已禁用", "未授权", "重新转换"].includes(status)) return "warn";
  if (["保留OMML", "手动指定依赖"].includes(status)) return "blue";
  return "";
}

function modeClass(mode) {
  return {
    local: "warn",
    web: "blue",
    hybrid: "good",
  }[mode] || "";
}

function escapeHtml(text) {
  return String(text ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  })[char]);
}

function escapeRegExp(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function toast(message) {
  const note = document.createElement("div");
  note.textContent = message;
  note.className = "toast";
  document.body.appendChild(note);
  setTimeout(() => note.remove(), 1800);
}

bindEvents();
loadAll().catch((error) => {
  $("#serviceStatus").textContent = "连接失败";
  $("#localStatus").textContent = error.message;
  $("#localDot").classList.add("warn");
  console.error(error);
});
