# Modular Theme Dashboard - Obsidian Dashboard Plugin

> 一款功能丰富的 Obsidian 仪表盘插件，提供自由拖拽布局、10 个功能模块、8 款精美主题，以及无限实例化能力。

**V15** — 融合 V14 自由拖拽布局架构与 V11 全量功能模块，所有模块均支持无限实例化。

## Preview

```
+----------------------------------------------------------+
|  🏠 Modular Theme Dashboard  [➕ 添加板块] [🎨 主题] [🔄 刷新] [⚙️]  |
+----------------------------------------------------------+
|                                                          |
|  +----------+  +----------+  +----------+  +----------+   |
|  | 🌤️ 天气  |  | 📅 日历  |  | 📈 统计  |  | ✅ 待办  |   |
|  |          |  |          |  |          |  |          |   |
|  | 实时天气  |  | 农历日历  |  | 笔记统计  |  | 增删改查  |   |
|  | 3天预报   |  | 节日节气  |  | 文件排行  |  | 进度追踪  |   |
|  +----------+  +----------+  +----------+  +----------+   |
|                                                          |
|  +----------+  +----------+  +----------+  +----------+   |
|  | 🕐 最近  |  | 🔥 资讯  |  | 📂 目录  |  | 💡 AI   |   |
|  |          |  |          |  |          |  |          |   |
|  | 最近修改  |  | 热点新闻  |  | 树形目录  |  | AI 洞察  |   |
|  | 点击打开  |  | 分页浏览  |  | 快速导航  |  | 智能分析  |   |
|  +----------+  +----------+  +----------+  +----------+   |
|                                                          |
|  +-------------------+  +-------------------+            |
|  | 🌐 网页预览       |  | 📺 网页视频       |            |
|  |                   |  |                   |            |
|  | 内嵌网页浏览器    |  | 内嵌视频播放器    |            |
|  | 缩放 / 平移 / 刷新|  | 缩放 / 平移 / 刷新|            |
|  +-------------------+  +-------------------+            |
+----------------------------------------------------------+
```

## Features

### 🎨 自由拖拽布局

- **绝对定位画布** — 模块卡片以 `position: absolute` 自由布局，可放置在画布任意位置
- **拖拽移动** — 按住卡片标题栏拖拽，实时移动到任意位置
- **自由缩放** — 每张卡片支持 `resize: both`，拖动右下角调整宽高
- **布局持久化** — 所有卡片位置、尺寸自动保存到 `data.json`，重启不丢失

### 🧩 完全模块化架构

- **热加载模块系统** — `ModuleManager` 动态扫描 `modules/` 目录，自动加载所有 `.js` 模块
- **标准模块接口** — 每个模块导出 `id / title / icon / defaultSettings / styles / render / renderSettings`
- **运行时注入** — 通过 `with(_runtimeCtx)` 自动注入 `app / plugin / moment / requestUrl` 等上下文
- **零耦合** — 新增模块只需在 `modules/` 目录下新建 `.js` 文件，无需修改主代码

### ♾️ 无限实例化

- **任意模块可克隆** — 点击顶部 `➕` 按钮，可为任何模块创建新实例
- **独立配置** — 每个实例拥有独立的设置、缓存、会话（如网页预览的 partition 隔离）
- **实例 ID 规范** — 格式 `module#N`（如 `weather#1`、`web-preview#2`）
- **自动继承默认值** — 新实例自动从模块的 `defaultSettings` 深拷贝初始配置

### 🌤️ 天气模块

- **高德地图 API** — 基于高德天气 API 的实时天气数据
- **实时 + 预报** — 当前天气实况 + 未来 3 天预报
- **详细信息** — 温度、湿度、风力、风向、体感温度等
- **可配置城市** — 支持自定义城市，API Key 配置

### 📅 日历模块

- **农历完整支持** — 天干地支、农历月份、节气、传统节日
- **月历翻页** — 支持前后月份切换
- **节日标注** — 公历节日 + 农历节日 + 24 节气自动标注
- **今日高亮** — 当前日期主题色高亮，周末特殊颜色

### 📈 笔记统计模块

- **四维统计** — 笔记总数 / 总字数 / 文件夹数 / 平均字数
- **文件夹排行** — Top 5 文件夹按笔记数量排序，带可视化进度条
- **万级格式化** — 超过 10000 自动显示为 `X.X万`

### ✅ 待办事项模块

- **完整 CRUD** — 新增、编辑（双击）、删除、完成标记
- **三种筛选** — 全部 / 未完成 / 已完成
- **进度统计** — 实时显示完成进度（如 `3/7`）
- **Markdown 持久化** — 待办数据读写指定文件夹的 `.md` 文件

### 🕐 最近文件模块

- **最近修改** — 按文件修改时间排序显示最近打开的笔记
- **相对时间** — 智能显示 `刚刚 / X分钟前 / X小时前 / X天前`
- **一键打开** — 点击直接在编辑器中打开文件
- **可配置数量** — 支持设置显示文件数上限

### 🔥 资讯模块

- **热点新闻** — 基于 AI HOT RSS 的实时热点新闻
- **分类标签** — 支持多分类切换浏览
- **分页浏览** — 支持上/下翻页，显示当前页码
- **一键跳转** — 点击新闻标题直接打开原文链接

### 📂 目录模块

- **树形目录** — 递归显示文件夹结构，可展开/折叠
- **自定义根目录** — 支持配置多个根目录显示
- **展开状态持久化** — 折叠/展开状态保存到设置，重启后恢复
- **文件计数** — 每个文件夹旁显示包含的笔记数量

### 💡 AI 洞察模块

- **AI 笔记分析** — 自动分析最近 5 篇笔记，调用 AI API 生成洞察摘要
- **OpenAI 兼容接口** — 支持任意 OpenAI 兼容 API（自定义 URL、模型、温度参数）
- **当天缓存** — 同一天内不重复调用，节省 API 费用
- **全局节流** — 2 秒最小请求间隔 + 实例级可配置延迟（0~10 秒），避免多实例同时请求被限频

### 🌐 网页预览模块

- **内嵌浏览器** — 在仪表盘内嵌入网页，无需切换窗口
- **三层缩放架构** — viewport → wrapper（transform scale/translate）→ iframe，精确控制缩放和平移
- **工具栏** — URL 输入栏 + 缩放按钮（➖/➕）+ X/Y 偏移 + 🔄 刷新
- **完整沙箱** — 支持页面内登录、OAuth 弹窗等完整浏览器交互

### 📺 网页视频模块

- **内嵌视频播放** — 基于 Electron `<webview>` 在仪表盘内播放视频网站
- **独立会话** — 每个实例使用独立 partition，支持同时登录不同账号
- **三层缩放架构** — 同网页预览，支持缩放、平移、刷新
- **弹幕屏蔽** — 自动注入 CSS 屏蔽 B 站等网站的广告弹幕

### 🎨 8 款精美主题

| 主题 | 名称 | 风格 |
|------|------|------|
| 晨曦 | Dawn | 暖橙色调，温暖柔和 |
| 侘寂 | Sabi | 柔和绿灰，日式简约 |
| 暮光 | Dusk | 深蓝紫调，暗色优雅 |
| 海岸 | Coastal | 青绿清新，自然舒适 |
| 丰收 | Harvest | 金棕暖调，秋日丰收 |
| 墨迹 | Ink | 深灰冷调，墨色沉稳 |
| 亚麻 | Linen | 米色布纹，低调温暖 |
| 碳灰 | Carbon | 纯黑背景，科技感 |

支持通过顶部工具栏一键切换主题，也可在设置中自定义卡片背景色和透明度。

## Installation

### From GitHub (Manual)

1. Go to [Releases](https://github.com/liamzy2021/Obsidian--Modular-Theme-Dashboard-Free-Drag-and-Drop/releases) and download the latest `main.js`, `manifest.json`, `styles.css` and `assets/` folder
2. Create a folder named `ai-smart-dashboard-v15` in your Obsidian vault's `.obsidian/plugins/` directory
3. Copy all downloaded files into this folder, maintaining the structure:
   ```
   .obsidian/plugins/ai-smart-dashboard-v15/
   ├── main.js
   ├── manifest.json
   ├── styles.css
   ├── assets/
   │   └── donate-qrcode.png
   └── modules/
       ├── ai-insight.js
       ├── calendar.js
       ├── directory.js
       ├── news.js
       ├── recent.js
       ├── stats.js
       ├── todo.js
       ├── weather.js
       ├── web-preview.js
       └── web-video.js
   ```
4. Restart Obsidian or reload plugins
5. Go to Settings → Community Plugins → Enable "Modular Theme Dashboard"

### From Obsidian Community Plugins

1. Open Settings → Community Plugins → Browse
2. Search for "Modular Theme Dashboard"
3. Click Install, then Enable

## Usage

- **打开仪表盘** — 左侧边栏点击 🏠 图标，或通过命令面板搜索 "Modular Theme Dashboard"
- **拖拽卡片** — 按住卡片标题栏拖拽到目标位置
- **调整大小** — 拖动卡片右下角调整宽高
- **添加新实例** — 点击顶部 `➕` 按钮，选择要添加的模块
- **切换主题** — 点击顶部 `🎨` 按钮，从 8 款主题中选择
- **模块设置** — 通过顶部 `⚙️` 按钮 → 滚动到对应模块区域，或通过 Obsidian 设置面板

## Configuration

Each module has its own settings section. Key configurations:

| Module | Key Settings |
|--------|-------------|
| 天气 | 城市、高德 API Key |
| AI 洞察 | API URL、API Key、模型名称、温度参数、请求延迟 |
| 网页预览 | 默认 URL、缩放比例、X/Y 偏移 |
| 网页视频 | 默认 URL、缩放比例、X/Y 偏移 |
| 待办事项 | 存储文件夹路径 |
| 目录 | 显示的根目录列表 |

Global settings:
- **主题选择** — 8 款预设主题
- **卡片背景色** — 自定义卡片背景色 + 透明度
- **显示/隐藏顶部栏** — 控制顶部工具栏显示
- **模块开关** — 启用/禁用每个模块
- **模块排序** — 拖拽调整模块显示顺序

## Architecture

```
ai-dashboard-v15/
├── main.js              # 插件入口（Plugin + DashboardView + ModuleManager）
├── manifest.json        # Obsidian 插件清单
├── styles.css           # 全局样式 + 各模块基础样式
├── assets/
│   └── donate-qrcode.png
└── modules/
    ├── weather.js       # 天气模块（高德 API）
    ├── calendar.js      # 日历模块（农历/节气/节日）
    ├── stats.js         # 统计模块（笔记统计/文件夹排行）
    ├── todo.js          # 待办模块（CRUD + Markdown 持久化）
    ├── recent.js        # 最近文件模块
    ├── news.js          # 资讯模块（AI HOT RSS）
    ├── directory.js     # 目录模块（树形目录）
    ├── ai-insight.js    # AI 洞察模块（OpenAI 兼容 API）
    ├── web-preview.js   # 网页预览模块（iframe 三层缩放）
    └── web-video.js     # 网页视频模块（webview 三层缩放）
```

### Core Design

- **`ModuleManager`** — 动态扫描 `modules/` 目录，通过 `new Function()` + `with(_runtimeCtx)` 注入 Obsidian API 上下文并执行模块代码
- **`DashboardView`** — 基于 `ItemView` 的自由布局视图，管理卡片渲染、拖拽、缩放、主题切换
- **模块接口** — 每个模块导出标准接口 `{ id, title, icon, defaultSettings, styles, render, renderSettings }`
- **实例系统** — `settings.instances[]` 数组存储所有实例信息，`settings.modules[instanceId]` 存储实例独立配置

## Requirements

- Obsidian **0.15.0** or later
- Desktop app recommended (web-view module requires Electron)
- Weather module requires [Amap (高德) API Key](https://lbs.amap.com/)
- AI Insight module requires an OpenAI-compatible API endpoint

## License

MIT

## Author

**栗子仁儿**
