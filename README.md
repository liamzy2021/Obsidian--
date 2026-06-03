# 中文说明

> 一款功能丰富的自由拖拽仪表盘插件，提供 10 个功能模块、8 款精美主题，以及无限实例化能力。

**V16** — 融合 V14 自由拖拽布局架构与 V11 全量功能模块，所有模块均支持无限实例化。

---

## 预览

<img width="1905" height="1358" alt="Modular Theme Dashboard 预览" src="https://github.com/user-attachments/assets/09b4854a-7281-4c7c-bfc0-fdd7336e0230" />

---

## 功能特性

### 自由拖拽布局

- **绝对定位画布** — 模块卡片以 `position: absolute` 自由布局，可放置在画布任意位置
- **拖拽移动** — 按住卡片标题栏拖拽，实时移动到任意位置
- **自由缩放** — 每张卡片支持 `resize: both`，拖动右下角调整宽高
- **布局持久化** — 所有卡片位置、尺寸自动保存到 `data.json`，重启不丢失

### 完全模块化架构

- **内置模块系统** — 10 个功能模块全部内置于 `main.js` 中，安装即用，无需额外文件
- **标准模块接口** — 每个模块导出 `id / title / icon / defaultSettings / styles / render / renderSettings`
- **运行时注入** — 通过 `with(_runtimeCtx)` 自动注入 `app / plugin / moment / requestUrl` 等上下文
- **零耦合** — 模块间完全独立，互不影响

### 无限实例化

- **任意模块可克隆** — 点击顶部 `➕` 按钮，可为任何模块创建新实例
- **独立配置** — 每个实例拥有独立的设置、缓存、会话（如网页预览的 partition 隔离）
- **实例 ID 规范** — 格式 `module#N`（如 `weather#1`、`web-preview#2`）
- **自动继承默认值** — 新实例自动从模块的 `defaultSettings` 深拷贝初始配置

### 天气模块

- **高德地图 API** — 基于高德天气 API 的实时天气数据
- **实时 + 预报** — 当前天气实况 + 未来 3 天预报
- **详细信息** — 温度、湿度、风力、风向、体感温度等
- **可配置城市** — 支持自定义城市，API Key 配置

### 日历模块

- **农历完整支持** — 天干地支、农历月份、节气、传统节日
- **月历翻页** — 支持前后月份切换
- **节日标注** — 公历节日 + 农历节日 + 24 节气自动标注
- **今日高亮** — 当前日期主题色高亮，周末特殊颜色

### 笔记统计模块

- **四维统计** — 笔记总数 / 总字数 / 文件夹数 / 平均字数
- **文件夹排行** — Top 5 文件夹按笔记数量排序，带可视化进度条
- **万级格式化** — 超过 10000 自动显示为 `X.X万`

### 待办事项模块

- **完整 CRUD** — 新增、编辑（双击）、删除、完成标记
- **三种筛选** — 全部 / 未完成 / 已完成
- **进度统计** — 实时显示完成进度（如 `3/7`）
- **Markdown 持久化** — 待办数据读写指定文件夹的 `.md` 文件

### 最近文件模块

- **最近修改** — 按文件修改时间排序显示最近打开的笔记
- **相对时间** — 智能显示 `刚刚 / X分钟前 / X小时前 / X天前`
- **一键打开** — 点击直接在编辑器中打开文件
- **可配置数量** — 支持设置显示文件数上限

### 资讯模块

- **热点新闻** — 基于 AI HOT RSS 的实时热点新闻
- **分类标签** — 支持多分类切换浏览
- **分页浏览** — 支持上/下翻页，显示当前页码
- **一键跳转** — 点击新闻标题直接打开原文链接

### 目录模块

- **树形目录** — 递归显示文件夹结构，可展开/折叠
- **自定义根目录** — 支持配置多个根目录显示
- **展开状态持久化** — 折叠/展开状态保存到设置，重启后恢复
- **文件计数** — 每个文件夹旁显示包含的笔记数量

### AI 洞察模块

- **AI 笔记分析** — 自动分析最近 5 篇笔记，调用 AI API 生成洞察摘要
- **OpenAI 兼容接口** — 支持任意 OpenAI 兼容 API（自定义 URL、模型、温度参数）
- **当天缓存** — 同一天内不重复调用，节省 API 费用
- **全局节流** — 2 秒最小请求间隔 + 实例级可配置延迟（0~10 秒），避免多实例同时请求被限频

### 网页预览模块

- **内嵌浏览器** — 在仪表盘内嵌入网页，无需切换窗口
- **三层缩放架构** — viewport → wrapper（transform scale/translate）→ iframe，精确控制缩放和平移
- **工具栏** — URL 输入栏 + 缩放按钮（➖/➕）+ X/Y 偏移 + 🔄 刷新
- **完整沙箱** — 支持页面内登录、OAuth 弹窗等完整浏览器交互

### 网页视频模块

- **内嵌视频播放** — 基于 Electron `<webview>` 在仪表盘内播放视频网站
- **独立会话** — 每个实例使用独立 partition，支持同时登录不同账号
- **三层缩放架构** — 同网页预览，支持缩放、平移、刷新
- **弹幕屏蔽** — 自动注入 CSS 屏蔽 B 站等网站的广告弹幕

### 8 款精美主题

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

---

## 安装

### 从 GitHub 手动安装

1. 前往 [Releases](https://github.com/liamzy2021/Obsidian--Modular-Theme-Dashboard-Free-Drag-and-Drop/releases) 下载最新的 `main.js`、`manifest.json` 和 `styles.css`
2. 在 Obsidian 库的 `.obsidian/plugins/` 目录下创建名为 `modular-theme-dashboard` 的文件夹
3. 将下载的 3 个文件复制到该文件夹中：
   ```
   .obsidian/plugins/modular-theme-dashboard/
   ├── main.js
   ├── manifest.json
   └── styles.css
   ```
4. 重启 Obsidian 或重新加载插件
5. 前往 **设置 → 社区插件 → 启用 "Modular Theme Dashboard"**

### 从社区插件安装

1. 打开 **设置 → 社区插件 → 浏览**
2. 搜索 "Modular Theme Dashboard"
3. 点击 **安装**，然后 **启用**

---

## 使用说明

- **打开仪表盘** — 左侧边栏点击 🏠 图标，或通过命令面板搜索 "Modular Theme Dashboard"
- **拖拽卡片** — 按住卡片标题栏拖拽到目标位置
- **调整大小** — 拖动卡片右下角调整宽高
- **添加新实例** — 点击顶部 `➕` 按钮，选择要添加的模块
- **切换主题** — 点击顶部 `🎨` 按钮，从 8 款主题中选择
- **模块设置** — 通过顶部 `⚙️` 按钮 → 滚动到对应模块区域，或通过 Obsidian 设置面板

---

## 配置说明

每个模块都有独立的设置区域。主要配置项：

| 模块 | 关键设置 |
|------|---------|
| 天气 | 城市、高德 API Key |
| AI 洞察 | API URL、API Key、模型名称、温度参数、请求延迟 |
| 网页预览 | 默认 URL、缩放比例、X/Y 偏移 |
| 网页视频 | 默认 URL、缩放比例、X/Y 偏移 |
| 待办事项 | 存储文件夹路径 |
| 目录 | 显示的根目录列表 |

全局设置：
- **主题选择** — 8 款预设主题
- **卡片背景色** — 自定义卡片背景色 + 透明度
- **显示/隐藏顶部栏** — 控制顶部工具栏显示
- **模块开关** — 启用/禁用每个模块
- **模块排序** — 拖拽调整模块显示顺序

---

## 系统要求

- Obsidian **0.15.0** 或更高版本
- 推荐桌面端（网页视频模块需要 Electron）
- 天气模块需要[高德地图 API Key](https://lbs.amap.com/)
- AI 洞察模块需要 OpenAI 兼容的 API 接口

---

## 许可证

MIT

---

## 作者

**栗子仁儿** · [GitHub](https://github.com/liamzy2021)

---

## 打赏

![打赏](https://img-reg-ab.imagency.cn/e/19467f4b916c082ee6ef3b9d81aa9ecb.png)

如果您觉得这个插件有帮助，欢迎打赏支持！

---

# Modular Theme Dashboard

> A feature-rich free-layout dashboard plugin with 10 widgets, 8 beautiful themes, and unlimited module instancing support.

**V16** — Combines V14's free drag-and-drop layout architecture with V11's full feature set, with all modules supporting unlimited instancing.

---

## Preview

<img width="1905" height="1358" alt="Modular Theme Dashboard Preview" src="https://github.com/user-attachments/assets/09b4854a-7281-4c7c-bfc0-fdd7336e0230" />

---

## Features

### Free Drag-and-Drop Layout

- **Absolute Position Canvas** — Module cards use `position: absolute` and can be placed anywhere on the canvas
- **Drag to Move** — Hold the card title bar to drag and move cards in real time
- **Resize Freely** — Each card supports `resize: both`; drag the bottom-right corner to adjust dimensions
- **Persistent Layout** — All card positions and sizes are auto-saved to `data.json` and persist across restarts

### Fully Modular Architecture

- **Built-in Module System** — All 10 modules are built into `main.js`; install and use immediately, no extra files needed
- **Standard Module Interface** — Each module exports `id / title / icon / defaultSettings / styles / render / renderSettings`
- **Runtime Context Injection** — Automatically injects `app / plugin / moment / requestUrl` context via `with(_runtimeCtx)`
- **Zero Coupling** — Modules are completely independent from each other

### Unlimited Instancing

- **Clone Any Module** — Click the top `+` button to create new instances of any module
- **Independent Config** — Each instance has its own settings, cache, and session (e.g., web preview partition isolation)
- **Instance ID Format** — `module#N` (e.g., `weather#1`, `web-preview#2`)
- **Auto-Inherit Defaults** — New instances automatically deep-copy initial config from the module's `defaultSettings`

### Weather Module

- **Amap API** — Real-time weather data via the Amap Weather API
- **Current + Forecast** — Live weather conditions + 3-day forecast
- **Detailed Info** — Temperature, humidity, wind force, wind direction, perceived temperature
- **Configurable City** — Supports custom city selection and API Key configuration

### Calendar Module

- **Full Lunar Support** — Heavenly stems and earthly branches, lunar months, solar terms, traditional holidays
- **Monthly Navigation** — Support for switching between previous/next months
- **Holiday Annotations** — Public holidays + lunar holidays + 24 solar terms auto-annotated
- **Today Highlight** — Current date highlighted in theme color; weekends shown in special colors

### Note Statistics Module

- **Four-Dimensional Stats** — Total notes / total words / folder count / average words
- **Folder Rankings** — Top 5 folders sorted by note count, with visual progress bars
- **Smart Formatting** — Auto-formatting for large numbers

### To-Do Module

- **Full CRUD** — Add, edit (double-click), delete, and mark as complete
- **Three Filters** — All / Incomplete / Completed
- **Progress Stats** — Real-time completion progress display (e.g., `3/7`)
- **Markdown Persistence** — To-do data reads/writes `.md` files in a specified folder

### Recent Files Module

- **Recently Modified** — Sorts and displays recently opened notes by modification time
- **Relative Time** — Smart display: `Just now / X min ago / X hr ago / X day ago`
- **One-Click Open** — Click to open the file directly in the editor
- **Configurable Count** — Supports setting a maximum number of files to display

### News Module

- **Hot News** — Real-time trending news based on AI HOT RSS
- **Category Tags** — Supports switching between multiple categories
- **Pagination** — Supports previous/next page browsing with current page number display
- **One-Click Jump** — Click news title to open the original article link

### Directory Module

- **Tree Directory** — Recursively displays folder structure; expandable/collapsible
- **Custom Root Directories** — Supports configuring multiple root directories to display
- **Expanded State Persistence** — Collapse/expand state saved to settings and restored after restart
- **File Count** — Displays the number of notes contained in each folder

### AI Insight Module

- **AI Note Analysis** — Automatically analyzes the last 5 notes, calls AI API to generate insight summaries
- **OpenAI-Compatible API** — Supports any OpenAI-compatible API (custom URL, model, temperature)
- **Daily Cache** — Does not re-call within the same day, saving API costs
- **Global Throttling** — 2-second minimum request interval + per-instance configurable delay (0-10s) to prevent rate-limiting

### Web Preview Module

- **Embedded Browser** — Embeds web pages inside the dashboard without switching windows
- **Three-Layer Zoom Architecture** — viewport -> wrapper (transform scale/translate) -> iframe for precise zoom and pan control
- **Toolbar** — URL input bar + zoom buttons (-/+) + X/Y offset + refresh button
- **Full Sandbox** — Supports in-page login, OAuth popups, and full browser interaction

### Web Video Module

- **Embedded Video Player** — Plays video websites inside the dashboard via Electron webview
- **Independent Session** — Each instance uses an independent partition, supporting simultaneous login to different accounts
- **Three-Layer Zoom Architecture** — Same as Web Preview, supporting zoom, pan, and refresh
- **Danmaku Blocking** — Automatically injects CSS to block ad danmaku on video sites

### 8 Beautiful Themes

| Theme | Style |
|-------|-------|
| Dawn | Warm orange tones, gentle and soft |
| Sabi | Soft green-grey, Japanese minimalism |
| Dusk | Deep blue-purple, elegant dark |
| Coastal | Cyan-green, fresh and natural |
| Harvest | Golden-brown, autumn harvest |
| Ink | Deep grey-cool, ink-like composure |
| Linen | Beige linen, understated warmth |
| Carbon | Pure black background, tech feel |

Switch themes with one click via the top toolbar, or customize card background colors and transparency in settings.

---

## Installation

### From GitHub (Manual)

1. Go to [Releases](https://github.com/liamzy2021/Obsidian--Modular-Theme-Dashboard-Free-Drag-and-Drop/releases) and download the latest `main.js`, `manifest.json`, and `styles.css`
2. Create a folder named `modular-theme-dashboard` in your Obsidian vault's `.obsidian/plugins/` directory
3. Copy all 3 files into this folder:
   ```
   .obsidian/plugins/modular-theme-dashboard/
   ├── main.js
   ├── manifest.json
   └── styles.css
   ```
4. Restart Obsidian or reload plugins
5. Go to **Settings -> Community Plugins -> Enable "Modular Theme Dashboard"**

### From Community Plugins

1. Open **Settings -> Community Plugins -> Browse**
2. Search for "Modular Theme Dashboard"
3. Click **Install**, then **Enable**

---

## Usage

- **Open Dashboard** — Click the home icon in the left sidebar, or search for "Modular Theme Dashboard" in the command palette
- **Drag Cards** — Hold the card title bar to drag to the desired position
- **Resize** — Drag the bottom-right corner of a card to adjust width and height
- **Add New Instance** — Click the top `+` button and select the module to add
- **Switch Theme** — Click the theme button in the top toolbar and choose from 8 themes
- **Module Settings** — Via the top settings button, scroll to the module section, or through the Obsidian Settings panel

---

## Configuration

Each module has its own settings section. Key configurations:

| Module | Key Settings |
|--------|-------------|
| Weather | City, Amap API Key |
| AI Insight | API URL, API Key, Model Name, Temperature, Request Delay |
| Web Preview | Default URL, Zoom Scale, X/Y Offset |
| Web Video | Default URL, Zoom Scale, X/Y Offset |
| To-Do | Storage folder path |
| Directory | Root directory list to display |

Global settings:
- **Theme Selection** — 8 preset themes
- **Card Background** — Custom card background color + transparency
- **Show/Hide Top Bar** — Toggle top toolbar visibility
- **Module Toggles** — Enable/disable each module
- **Module Order** — Drag to adjust module display order

---

## Architecture

```
modular-theme-dashboard/
├── main.js              # Plugin entry + all 10 built-in modules (Plugin + DashboardView + ModuleManager)
├── manifest.json        # Plugin manifest
├── styles.css           # Global styles + base module styles
└── LICENSE              # MIT License
```

### Core Design

- **`ModuleManager`** — Loads built-in modules from memory, injects plugin API context via `new Function()` + `with(_runtimeCtx)`, and executes module code
- **`DashboardView`** — A free-layout view based on `ItemView`, managing card rendering, drag-and-drop, resizing, and theme switching
- **Module Interface** — Each module exports a standard interface `{ id, title, icon, defaultSettings, styles, render, renderSettings }`
- **Instance System** — `settings.instances[]` stores all instance info; `settings.modules[instanceId]` stores per-instance independent configuration

---

## Requirements

- Obsidian **0.15.0** or later
- Desktop app recommended (web-view module requires Electron)
- Weather module requires an [Amap API Key](https://lbs.amap.com/)
- AI Insight module requires an OpenAI-compatible API endpoint

---

## License

MIT

---

## Author

**liamzy2021** · [GitHub](https://github.com/liamzy2021)
