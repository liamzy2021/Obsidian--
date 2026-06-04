<p align="center">
  <a href="#modular-theme-dashboard">English</a> &nbsp;|&nbsp;
  <a href="#中文说明">中文</a> &nbsp;|&nbsp;
  <a href="#-support"><img src="https://img.shields.io/badge/☕-Donate-orange?style=flat-square" /></a>
  <a href="#打赏"><img src="https://img.shields.io/badge/💰-打赏-red?style=flat-square" /></a>
</p>

---

# Modular Theme Dashboard

> A feature-rich free-layout dashboard plugin with 25 widgets, 8 beautiful themes, and unlimited module instancing support.

**V17.0.7** — Based on V14's free drag-and-drop layout architecture, now featuring 25 modules with full FileViewer support for 9+ file formats.

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

- **25 Built-in Modules** — All modules built into the plugin, install and use immediately
- **Standard Module Interface** — Each module exports `id / title / icon / defaultSettings / styles / render / renderSettings`
- **Runtime Context Injection** — Automatically injects `app / plugin / moment / requestUrl` context
- **Zero Coupling** — Modules are completely independent; enable/disable individually

### Unlimited Instancing

- **Clone Any Module** — Create new instances of any module via the `+` button
- **Independent Config** — Each instance has its own settings, cache, and session
- **Instance ID Format** — `module#N` (e.g., `weather#1`, `web-preview#2`)

### Built-in FileViewer (9+ Formats)

- **Spreadsheets** — `.xlsx`, `.xls`, `.csv` — full table rendering with SheetJS
- **Word Documents** — `.docx` (mammoth.js), `.doc` (docstream / CFB)
- **HTML / Text / Code** — `.html`, `.txt`, `.json`, `.js`, etc.
- **Images & Video** — `.png`, `.jpg`, `.gif`, `.webp`, `.mp4`, `.webm`
- **Toggle per Format** — Enable/disable each format in Settings
- **System Fallback** — Disabled formats auto-open with system default app

### Core Modules

| Module | Description |
|--------|------------|
| Weather | Real-time weather + 3-day forecast via Amap API |
| Calendar | Full lunar calendar with solar terms, holidays, and monthly navigation |
| Stats | Note count, word count, folder count, average words, Top 5 folders |
| To-Do | Full CRUD with filters, progress tracking, Markdown persistence |
| Recent Files | Recently modified notes with relative timestamps |
| News | Real-time trending news via AI HOT RSS with categories & pagination |
| Directory | Tree-style folder structure with expand/collapse and file counts |
| AI Insight | AI-powered note analysis via OpenAI-compatible API with daily cache |
| Web Preview | Embedded browser with URL bar, zoom/pan controls, full sandbox |
| Web Video | Embedded video player via Electron webview with independent sessions |
| Image Gallery | Browse images in grid or masonry layout with lightbox preview |
| Media Gallery | Unified media browser for images, video, and audio files |
| Vault Stats | Comprehensive vault statistics with visual charts |
| Code Editor | Built-in code editor with syntax highlighting |
| Data Editor | Edit and manage structured data (JSON, YAML, CSV) directly |
| Spreadsheet | Full spreadsheet editing with cell formatting |
| Doc Viewer | Rich document preview for various formats |
| HTML Viewer | Render HTML files directly within the dashboard |
| URL Opener | Quick-launch URLs and bookmarks from the dashboard |
| XHS Importer | Import content from Xiaohongshu (RED) |
| Excel to Markdown | Convert clipboard Excel tables to Markdown format |

### Utility Modules (Global Features)

| Module | Description |
|--------|------------|
| Folder Counter | Show file counts next to folder names in the file explorer |
| Image Tools | 18 right-click operations: format conversion, resize, rotate, flip, align, compress |
| Table Resize | Drag to resize table column widths in reading view |
| Auto-Play Loop | Automatic media playback loop engine |

### 8 Beautiful Themes

| Theme | Name | Style |
|-------|------|-------|
| Dawn | 晨曦 | Warm orange tones, gentle and soft |
| Sabi | 侘寂 | Soft green-grey, Japanese minimalism |
| Dusk | 暮光 | Deep blue-purple, elegant dark |
| Coastal | 海岸 | Cyan-green, fresh and natural |
| Harvest | 丰收 | Golden-brown, autumn harvest |
| Ink | 墨迹 | Deep grey-cool, ink-like composure |
| Linen | 亚麻 | Beige linen, understated warmth |
| Carbon | 碳灰 | Pure black background, tech feel |

### Settings Backup

- **Export** — Download all settings as a JSON file (`obsidian-dashboard-settings-YYYY-MM-DD.json`)
- **Import** — Restore settings from a previously exported file
- Perfect for migrating settings between vaults or sharing configurations

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
5. Go to **Settings → Community Plugins → Enable "Modular Theme Dashboard"**

### From Community Plugins

1. Open **Settings → Community Plugins → Browse**
2. Search for "Modular Theme Dashboard"
3. Click **Install**, then **Enable**

---

## Usage

- **Open Dashboard** — Click the 🏠 icon in the left sidebar, or search in the command palette
- **Drag Cards** — Hold the card title bar to drag to any position
- **Resize** — Drag the bottom-right corner of a card
- **Add Instance** — Click the top `+` button and select a module
- **Switch Theme** — Click the `🎨` button in the top toolbar
- **Settings** — Via the `⚙️` button, or through Obsidian Settings panel

---

## Configuration

| Module | Key Settings |
|--------|-------------|
| Weather | City, Amap API Key |
| AI Insight | API URL, API Key, Model, Temperature, Delay |
| Web Preview | Default URL, Zoom, X/Y Offset |
| Web Video | Default URL, Zoom, X/Y Offset |
| To-Do | Storage folder path |
| Directory | Root directories to display |
| XHS Importer | Import configuration |

**Global Settings:** Theme selection, card background color + opacity, top bar visibility, module toggles, module ordering, FileViewer format toggles, utility module toggles.

---

## Architecture

```
modular-theme-dashboard/
├── main.js              # Plugin entry + all 25 built-in modules
├── manifest.json        # Plugin manifest
├── styles.css           # Global styles
├── src/core/            # Core source (00-header, 01-lib-loaders, 02-file-viewer, 03-themes, 04-defaults, 05-module-manager, 06-dashboard-view, 07-settings-tab, 08-plugin)
├── modules/             # 25 feature modules (dynamic loading in dev mode)
│   └── file-viewers/    # 9 FileViewer extensions (xlsx, doc, docx, html, image, video, text, legacy-office, utils)
└── libs/                # Third-party libraries (xlsx.full.min.js, mammoth.browser.min.js)
```

### Core Design

- **`ModuleManager`** — Loads built-in modules, injects plugin API context, executes module code
- **`DashboardView`** — Free-layout view based on `ItemView`, managing cards, drag-and-drop, resizing, themes
- **`DashboardFileViewer`** — Custom `FileView` for rendering 9+ file formats inline
- **Module Interface** — Standard `{ id, title, icon, defaultSettings, styles, render, renderSettings }`
- **Instance System** — `settings.instances[]` for layout; `settings.modules[instanceId]` for per-instance config

---

## Requirements

- Obsidian **0.15.0** or later
- Desktop app recommended (web preview/video modules require Electron)
- Weather module requires an [Amap API Key](https://lbs.amap.com/)
- AI Insight module requires an OpenAI-compatible API endpoint

---

## License

MIT

---

## Author

**栗子仁儿 (liamzy2021)** · [GitHub](https://github.com/liamzy2021)

---

## ☕ Support

If you find this plugin helpful, consider supporting its development:

<p align="center">
  <a href="https://opencollective.com/obsidian--modular-theme-dashboard-free-drag-and-drop" target="_blank" rel="noopener">
    <img src="https://bootflare.com/wp-content/uploads/2025/12/Opencollective-Logo.png" alt="Open Collective" width="120" /><br/>
    <img src="https://img.shields.io/badge/☕-Buy_Me_a_Coffee-orange?style=for-the-badge&logo=buy-me-a-coffee" />
  </a>
</p>

---

<a name="中文说明"></a>
<p align="center">
  <a href="#modular-theme-dashboard">English</a> &nbsp;|&nbsp;
  <a href="#中文说明">中文</a> &nbsp;|&nbsp;
  <a href="#-support"><img src="https://img.shields.io/badge/☕-Donate-orange?style=flat-square" /></a>
  <a href="#打赏"><img src="https://img.shields.io/badge/💰-打赏-red?style=flat-square" /></a>
</p>

---

# 中文说明

> 一款功能丰富的自由拖拽仪表盘插件，提供 25 个功能模块、8 款精美主题，以及无限实例化能力。

**V17.0.7** — 基于 V14 自由拖拽布局架构，融合 25 个全量功能模块，并支持 9+ 种文件格式的内置文件查看器。

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

- **25 个内置模块** — 全部内置于插件中，安装即用
- **标准模块接口** — 每个模块导出 `id / title / icon / defaultSettings / styles / render / renderSettings`
- **运行时注入** — 自动注入 `app / plugin / moment / requestUrl` 上下文
- **零耦合** — 模块间完全独立，可单独启用/禁用

### 无限实例化

- **任意模块可克隆** — 点击顶部 `+` 按钮，创建任意模块的新实例
- **独立配置** — 每个实例拥有独立的设置、缓存、会话
- **实例 ID 规范** — 格式 `module#N`（如 `weather#1`、`web-preview#2`）

### 内置文件查看器（9+ 种格式）

- **表格文件** — `.xlsx`、`.xls`、`.csv` — 基于 SheetJS 的完整表格渲染
- **Word 文档** — `.docx`（mammoth.js）、`.doc`（docstream / CFB 解析）
- **HTML / 文本 / 代码** — `.html`、`.txt`、`.json`、`.js` 等
- **图片与视频** — `.png`、`.jpg`、`.gif`、`.webp`、`.mp4`、`.webm`
- **按格式开关** — 在设置中单独启用/禁用每种格式
- **系统回退** — 禁用的格式自动用系统默认程序打开

### 核心模块

| 模块 | 说明 |
|------|------|
| 天气 | 基于高德 API 的实时天气 + 未来 3 天预报 |
| 日历 | 农历完整支持，含节气、节日标注、月历翻页 |
| 笔记统计 | 笔记总数 / 总字数 / 文件夹数 / 平均字数 / Top 5 文件夹 |
| 待办事项 | 完整增删改查，三种筛选，进度统计，Markdown 持久化 |
| 最近文件 | 最近修改笔记，智能相对时间显示 |
| 资讯 | 基于 AI HOT RSS 的实时热点新闻，分类 + 分页 |
| 目录 | 树形文件夹结构，可展开/折叠，自定义根目录 |
| AI 洞察 | AI 笔记分析，OpenAI 兼容接口，当天缓存 |
| 网页预览 | 内嵌浏览器，URL 栏 + 缩放/平移 + 完整沙箱 |
| 网页视频 | 基于 Electron webview 的内嵌视频播放，独立会话 |
| 图片画廊 | 网格/瀑布流图片浏览，点击灯箱预览 |
| 媒体画廊 | 图片、视频、音频统一媒体浏览器 |
| 笔记统计 | 全面的笔记数据统计与可视化图表 |
| 代码编辑器 | 内置代码编辑器，支持语法高亮 |
| 数据编辑器 | 直接在仪表盘中编辑 JSON、YAML、CSV 等结构化数据 |
| 电子表格 | 完整电子表格编辑，支持单元格格式化 |
| 文档查看器 | 多种格式的富文档预览 |
| HTML 查看器 | 在仪表盘中直接渲染 HTML 文件 |
| URL 打开器 | 从仪表盘快速启动 URL 和书签 |
| 小红书导入 | 从小红书导入内容 |
| Excel 转表格 | 将剪贴板中的 Excel 表格转换为 Markdown 格式 |

### 实用模块（全局功能）

| 模块 | 说明 |
|------|------|
| 文件夹计数器 | 在文件浏览器中文件夹旁显示文件数量 |
| 图片处理 | 18 项右键菜单：格式转换、调整大小、旋转、翻转、对齐、压缩 |
| 表格列宽调整 | 阅读模式下拖拽调整表格列宽 |
| 自动播放循环 | 媒体自动播放循环引擎 |

### 8 款精美主题

| 主题 | 名称 | 风格 |
|------|------|------|
| Dawn | 晨曦 | 暖橙色调，温暖柔和 |
| Sabi | 侘寂 | 柔和绿灰，日式简约 |
| Dusk | 暮光 | 深蓝紫调，暗色优雅 |
| Coastal | 海岸 | 青绿清新，自然舒适 |
| Harvest | 丰收 | 金棕暖调，秋日丰收 |
| Ink | 墨迹 | 深灰冷调，墨色沉稳 |
| Linen | 亚麻 | 米色布纹，低调温暖 |
| Carbon | 碳灰 | 纯黑背景，科技感 |

支持通过顶部工具栏一键切换主题，也可在设置中自定义卡片背景色和透明度。

### 设置存档

- **导出** — 将所有设置保存为 JSON 文件（`obsidian-dashboard-settings-YYYY-MM-DD.json`）
- **导入** — 从之前导出的文件恢复设置
- 适合在多个仓库间迁移设置或分享配置

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

- **打开仪表盘** — 左侧边栏点击 🏠 图标，或通过命令面板搜索
- **拖拽卡片** — 按住卡片标题栏拖拽到目标位置
- **调整大小** — 拖动卡片右下角调整宽高
- **添加新实例** — 点击顶部 `+` 按钮，选择要添加的模块
- **切换主题** — 点击顶部 `🎨` 按钮，从 8 款主题中选择
- **模块设置** — 通过顶部 `⚙️` 按钮，或通过 Obsidian 设置面板

---

## 配置说明

| 模块 | 关键设置 |
|------|---------|
| 天气 | 城市、高德 API Key |
| AI 洞察 | API URL、API Key、模型名称、温度参数、请求延迟 |
| 网页预览 | 默认 URL、缩放比例、X/Y 偏移 |
| 网页视频 | 默认 URL、缩放比例、X/Y 偏移 |
| 待办事项 | 存储文件夹路径 |
| 目录 | 显示的根目录列表 |
| 小红书导入 | 导入配置 |

**全局设置：** 主题选择、卡片背景色 + 透明度、顶部栏显示/隐藏、模块开关、模块排序、文件查看器格式开关、实用模块开关。

---

## 架构

```
modular-theme-dashboard/
├── main.js              # 插件入口 + 全部 25 个内置模块
├── manifest.json        # 插件清单
├── styles.css           # 全局样式
├── src/core/            # 核心源码（00-header, 01-lib-loaders, 02-file-viewer, 03-themes, 04-defaults, 05-module-manager, 06-dashboard-view, 07-settings-tab, 08-plugin）
├── modules/             # 25 个功能模块（开发模式下动态加载）
│   └── file-viewers/    # 9 个文件查看器扩展（xlsx, doc, docx, html, image, video, text, legacy-office, utils）
└── libs/                # 第三方库（xlsx.full.min.js, mammoth.browser.min.js）
```

### 核心设计

- **`ModuleManager`** — 加载内置模块，注入插件 API 上下文，执行模块代码
- **`DashboardView`** — 基于 `ItemView` 的自由布局视图，管理卡片、拖拽、缩放、主题
- **`DashboardFileViewer`** — 自定义 `FileView`，支持 9+ 种文件格式的内联渲染
- **模块接口** — 标准 `{ id, title, icon, defaultSettings, styles, render, renderSettings }`
- **实例系统** — `settings.instances[]` 存储布局；`settings.modules[instanceId]` 存储每个实例的独立配置

---

## 系统要求

- Obsidian **0.15.0** 或更高版本
- 推荐桌面端（网页预览/视频模块需要 Electron）
- 天气模块需要[高德地图 API Key](https://lbs.amap.com/)
- AI 洞察模块需要 OpenAI 兼容的 API 接口

---

## 许可证

MIT

---

## 作者

**栗子仁儿 (liamzy2021)** · [GitHub](https://github.com/liamzy2021)

---

<a name="打赏"></a>

## 💰 打赏

如果您觉得这个插件有帮助，欢迎打赏支持！

<p align="center">
  <img src="https://img-reg-ab.imagency.cn/e/19467f4b916c082ee6ef3b9d81aa9ecb.png" alt="微信打赏二维码" width="200" />
</p>
