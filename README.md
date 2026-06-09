<p align="center">
  <a href="#modular-theme-dashboard">English</a> &nbsp;|&nbsp;
  <a href="#中文说明">中文</a> &nbsp;|&nbsp;
  <a href="#-support"><img src="https://img.shields.io/badge/☕-Support-orange?style=flat-square" /></a>
  <a href="#打赏"><img src="https://img.shields.io/badge/💰-打赏-red?style=flat-square" /></a>
</p>

---

# Modular Theme Dashboard

> A free-layout Obsidian dashboard with 25 widgets, 8 themes, 10+ file viewers including a lightning-fast table viewer with decimal precision control and multi-sheet browsing, and AI-powered multi-language translation — drag, drop, and make it yours.

**V17** — Three major updates: **Lightning-Fast Table Viewer** (instant open for any size spreadsheet, adjustable decimal precision, multi-sheet tab switching), **PPT / PPT File Viewer** (full support for `.ppt` and `.pptx` with smart text extraction) and **Multi-Platform Weather Module** (5 providers: Amap, Open-Meteo, wttr.in, OpenWeatherMap + custom URL template). Global weather coverage, free and unlimited for most providers.

---

## Preview

<p align="center">
  <img width="1905" alt="Modular Theme Dashboard Preview" src="https://github.com/user-attachments/assets/09b4854a-7281-4c7c-bfc0-fdd7336e0230" />
  <br/><em>🏠 Modular Theme Dashboard — full overview</em>
</p>



<p align="center">
  <img width="926" alt="Language Settings" src="https://github.com/user-attachments/assets/63b0583e-8eb0-4eba-8eb1-bebf52e01780" />
  <br/><em>🌍 AI-powered language translation — translate the UI into any language</em>
</p>

<p align="center">
  <img width="838" alt="Theme Gallery" src="https://github.com/user-attachments/assets/a5e14c51-569c-4d0d-8d40-a0dcaae00b92" />
  <br/><em>🎨 8 beautiful themes, one click to switch</em>
</p>

<p align="center">
  <img width="1677" alt="Full Feature Showcase" src="https://github.com/user-attachments/assets/91cd02b7-5500-4069-8353-8e1e976ee24c" />
  <br/><em>📊 25 modules covering weather, calendar, news, AI analysis, and more</em>
</p>

---

## Features

### 🌍 AI Custom Language (New in V17.0.8)

> **Your AI. Your language. Your dashboard.**

Tired of English-only plugins? Want to read your dashboard in Japanese, Russian, or even a local dialect like Sichuanese? This is for you.

**How it works:** Enter your OpenAI-compatible API key, type the target language (e.g., `日本語`, `Русский`, `四川话`, `Klingon`), hit translate — the AI converts all 500+ UI strings in under a minute. Preview the result, apply it, and your dashboard instantly speaks your language.

| Feature | Detail |
|---------|--------|
| 🗣️ **Any language** | Japanese, Korean, Russian, French, German, Spanish, Arabic, Hindi, Thai, Vietnamese... |
| 🏠 **Local dialects** | Cantonese (广东话), Sichuanese (四川话), Hokkien (闽南语), Shanghainese (上海话), Hakka (客家话), Teochew (潮汕话)... |
| 🚀 **Any AI** | OpenAI, DeepSeek, Moonshot, or any OpenAI-compatible API |
| 💾 **Persistent** | Translations are saved and survive restarts — switch languages anytime from the dropdown |
| 🔁 **Reusable** | Once translated, the language appears in your language selector with a ⚡ badge |

**Quick Start:** Settings → Language → AI Custom → enter target language → Translate → Preview → Apply

### 🧩 Free Drag-and-Drop Layout

- **Absolute Positioning Canvas** — Every module card lives exactly where you place it
- **Drag to Move** — Grab the title bar and drag anywhere in real time
- **Resize Freely** — Pull the bottom-right corner to set the perfect size
- **Persistent** — All positions and sizes are auto-saved; nothing shifts after a restart

### 📦 25 Built-in Modules

All modules are built in — no extra downloads, no dependencies to chase. Install and use immediately.

| Module | What It Does |
|--------|-------------|
| Weather | Real-time weather + 3-day forecast (5 providers: Amap, Open-Meteo, wttr.in, OpenWeatherMap, Custom) |
| Calendar | Lunar calendar with solar terms, holidays, monthly navigation |
| Stats | Note count, word count, folder stats, Top 5 folders |
| To-Do | Full CRUD with filters, progress tracking, Markdown persistence |
| Recent Files | Recently modified notes with relative timestamps |
| News | Trending headlines via AI HOT RSS, with categories & pagination |
| Directory | Tree-style folder browser with expand/collapse |
| AI Insight | AI-powered note analysis with daily cache |
| Web Preview | Embedded browser with URL bar, zoom, and sandbox |
| Web Video | Embedded video player via Electron webview |
| Image Gallery | Grid / masonry image browser with lightbox |
| Media Gallery | Unified browser for images, video, and audio |
| Vault Stats | Comprehensive statistics with visual charts |
| Code Editor | Built-in code editor with syntax highlighting |
| Data Editor | Edit JSON, YAML, CSV directly in the dashboard |
| Spreadsheet | Full spreadsheet with cell formatting |
| Doc Viewer | Rich preview for Word, HTML, and more |
| HTML Viewer | Render HTML files inline |
| URL Opener | Quick-launch bookmarks and URLs |
| XHS Importer | Import content from Xiaohongshu (RED) |
| Excel to Markdown | Paste clipboard tables as Markdown |

### 🔧 Utility Modules

| Module | Description |
|--------|------------|
| Folder Counter | Show file counts next to folder names in the explorer |
| Image Tools | 18 right-click operations: crop, resize, rotate, compress, and more |
| Table Resize | Drag to resize table columns in reading view |
| Auto-Play Loop | Automatic media playback loop engine |
| Audio Transcribe | Right-click audio/video files to transcribe via Cloud API or Local AI (New in V17.1.7) |

### 🔄 Unlimited Instancing

- **Clone Anything** — Create multiple instances of any module via the `+` button
- **Independent Config** — Each instance has its own settings, cache, and session
- **Smart IDs** — Instances use `module#N` naming (e.g., `weather#1`, `web-preview#2`)

### 📁 Built-in File Viewer (10+ Formats)

Open files directly in the dashboard — no need to leave Obsidian:

| Format | Support |
|--------|---------|
| Spreadsheets | `.xlsx`, `.xls`, `.csv` (SheetJS rendering) |
| Word Documents | `.docx` (mammoth.js), `.doc` (CFB parser) |
| PowerPoint | `.pptx` (JSZip + XML text extraction), `.ppt` (OLE2 binary parser with UTF-16LE/ASCII dual-scan) |
| Code & Text | `.html`, `.txt`, `.json`, `.js`, and more |
| Images | `.png`, `.jpg`, `.gif`, `.webp` |
| Video | `.mp4`, `.webm` |

Toggle each format on/off in Settings. Disabled formats fall back to your system default app.

### 📊 Lightning-Fast Table Viewer (New in V17.1.6)

> Open any `.xlsx` / `.xls` file instantly — from small tables to 100,000-row monster spreadsheets, it renders faster than any desktop spreadsheet editor.

| Feature | Detail |
|---------|--------|
| ⚡ **Instant Open** | No loading bar, no wait — renders in milliseconds regardless of file size |
| 🔢 **Adjustable Decimals** | One-click `− / +` stepper to set decimal precision (0-10), or toggle "Auto" for raw values. No more eye-straining `0.093408691346` |
| 📑 **Multi-Sheet Tabs** | All worksheet tabs listed at the top — click any tab to switch sheets instantly |
| 📏 **Smart Columns** | Column widths auto-fit content, with manual drag-to-resize handles on every column |
| 🌍 **Multi-Language** | Toolbar labels follow your plugin language setting (Chinese / English / AI-custom) |

<p align="center">
  <img width="2560" height="1379" alt="Table Viewer Overview" src="https://github.com/user-attachments/assets/f27906f8-3bcc-456a-bcc6-bc274ba91742" />
  <br/><em>📊 Table Viewer — decimal precision control and multi-sheet tabs</em>
</p>

<p align="center">
  <img width="2555" height="1352" alt="Table Viewer Multi-Sheet" src="https://github.com/user-attachments/assets/15db28ab-eebb-4663-8af5-4299739777a7" />
  <br/><em>📑 Switch between sheets with one click — all sheets listed as tabs</em>
</p>

### 🎙️ Audio / Video Transcription (New in V17.1.7)

> Right-click any audio or video file → Transcribe → text inserted into your note. No dashboard panel needed — just configure and go.

| Feature | Detail |
|---------|--------|
| 🎯 **Right-click to Transcribe** | Right-click any `.mp3` / `.m4a` / `.wav` / `.mp4` / `.webm` / `.flac` / `.ogg` file in the file explorer |
| ☁️ **Cloud API** | Any OpenAI-compatible `/v1/audio/transcriptions` endpoint — fill in your Base URL, API Key, and model name |
| 🖥️ **Local AI** | faster-whisper-server, LocalAI, docker-whisper — any OpenAI-compatible /v1/audio/transcriptions service |
| 📝 **5 Output Formats** | Plain text, bullet list, action items, timestamp chapters, speaker diarization |
| 🌍 **Multi-Language** | All UI text follows the plugin's language setting |
| 💰 **Free Options** | SiliconFlow (TeleSpeech ASR, free tier), Groq (Whisper, free tier), Cloudflare Workers AI |

**Recommended platforms (all support OpenAI-compatible API):**

*Cloud API:*

| Platform | Base URL | Notes |
|----------|----------|-------|
| SiliconFlow | `https://api.siliconflow.cn/v1` | 🇨🇳 China access, free TeleSpeech ASR |
| Groq | `https://api.groq.com/openai/v1` | Fast, free Whisper tier |
| Cloudflare Workers AI | `https://api.cloudflare.com/client/v4/accounts/{id}/ai/v1` | Free tier |
| OpenAI | `https://api.openai.com/v1` | Standard paid API |

*Local AI:*

| Platform | Base URL | Notes |
|----------|----------|-------|
| faster-whisper-server | `http://localhost:8000/v1` | Python, fully OpenAI-compatible, recommended |
| LocalAI | `http://localhost:8080/v1` | Go, multi-backend (whisper/faster-whisper/llama-cpp) |
| docker-whisper | `http://localhost:9000` | Docker one-click deploy, speaker diarization |

<p align="center">
  <img width="1105" height="1003" alt="Audio Transcribe Settings" src="https://github.com/user-attachments/assets/10424883-fa9d-4e98-93f6-857066158bbb" />
  <br/><em>🎙️ Audio Transcribe — cloud API configuration with platform recommendations</em>
</p>

<p align="center">
  <img width="442" height="462" alt="Audio Transcribe Context Menu" src="https://github.com/user-attachments/assets/6775045c-31a0-47ed-9d55-7d52fa5c175d" />
  <br/><em>🎯 Right-click an audio file → Transcribe Audio/Video</em>
</p>

### 🖼️ In-Note Code Block Galleries

> Place ````t```` (image) or ````s```` (media) code blocks in any note to render a gallery in reading view. Right-click the block for layout, size, columns, and spacing controls.

<p align="center">
  <img width="1207" alt="Code Block Gallery" src="https://github.com/user-attachments/assets/5efa44a2-ccdc-41d4-9461-eed7d1b43525" />
  <br/><em>📸 In-note gallery with spacing controls and smart center</em>
</p>

#### Image Gallery (```t)

````markdown
```t
/path/to/images|horizontal|4|200|0|0|10|12|true
```
````

**Quick format** (pipe-separated):

| Position | Field | Default | Description |
|----------|-------|---------|-------------|
| 1 | `path` | — | Folder path (required) |
| 2 | `layout` | `horizontal` | `horizontal` / `vertical` (waterfall) / `grid` |
| 3 | `columns` | `4` | Images per row |
| 4 | `height` | `200` | Image height in px |
| 5 | `width` | `0` | Image width in px (0 = auto) |
| 6 | `spacingLeft` | `0` | Left padding in px |
| 7 | `spacingRight` | `0` | Right padding in px |
| 8 | `itemGap` | `12` | Gap between images in px |
| 9 | `smartCenter` | `false` | `true` for auto-centered flex layout |

**Key-value format** (more readable):

```
path: /path/to/images
type: horizontal
columns: 4
height: 200
spacingleft: 10
spacingright: 10
itemgap: 12
smartcenter: true
```

**Supported image formats:** jpeg, jpg, gif, png, webp, tiff, tif, avif, bmp

**Right-click menu:** Edit / Delete / Custom Layout / Custom Size / Custom Columns / Spacing Settings (with Smart Center toggle)

#### Media Gallery (``s)

````markdown
```s
/path/to/media|grid|220|10|10|12|true
```
````

**Quick format** (pipe-separated):

| Position | Field | Default | Description |
|----------|-------|---------|-------------|
| 1 | `path` | — | Folder path (required) |
| 2 | `type` | `grid` | `grid` / `list` / `full` (full-width) |
| 3 | `size` | `220` | Media cell size in px |
| 4 | `spacingLeft` | `0` | Left padding in px |
| 5 | `spacingRight` | `0` | Right padding in px |
| 6 | `itemGap` | `10` | Gap between items in px |
| 7 | `smartCenter` | `false` | `true` for auto-centered flex layout |

**Key-value format:**

```
path: /path/to/media1, /path/to/media2
type: grid
size: 220
sort: name
limit: 50
spacingleft: 10
spacingright: 10
itemgap: 12
smartcenter: true
```

**Supported media formats:** Images (jpeg, jpg, gif, png, webp, tiff, svg, ico, heic, avif, bmp) · Video (mp4, webm, ogg, mov, mkv, avi) · Audio (mp3, wav, flac, aac, m4a, ogg)

**Right-click menu:** Edit / Delete / Custom Layout / Custom Size / Spacing Settings (with Smart Center toggle)

### 🎨 8 Beautiful Themes

| Theme | Style |
|-------|-------|
| Dawn (晨曦) | Warm orange, gentle and soft |
| Sabi (侘寂) | Soft green-grey, Japanese minimalism |
| Dusk (暮光) | Deep blue-purple, elegant dark |
| Coastal (海岸) | Cyan-green, fresh and natural |
| Harvest (丰收) | Golden-brown, autumn warmth |
| Ink (墨迹) | Deep grey, ink-wash composure |
| Linen (亚麻) | Beige linen, understated warmth |
| Carbon (碳灰) | Pure black, modern tech |

Switch themes from the top toolbar, or customize card background color and opacity in Settings.

### 💾 Settings Backup

- **Export** — Save all settings as `obsidian-dashboard-settings-YYYY-MM-DD.json`
- **Import** — Restore from a previously exported file
- Great for migrating between vaults or sharing configurations

---

## Installation

### From GitHub (Manual)

1. Go to [Releases](https://github.com/liamzy2021/Obsidian--Modular-Theme-Dashboard-Free-Drag-and-Drop/releases) and download the latest `main.js`, `manifest.json`, and `styles.css`
2. Create a folder named `modular-theme-dashboard` in your vault's `.obsidian/plugins/` directory
3. Copy all 3 files in:
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

- **Open Dashboard** — Click the 🏠 icon in the left sidebar, or use the command palette
- **Drag Cards** — Hold the title bar to drag cards anywhere
- **Resize** — Drag the bottom-right corner of any card
- **Add Module** — Click the `+` button at the top and pick a module
- **Switch Theme** — Click the `🎨` button in the top toolbar
- **Settings** — Via the `⚙️` button or Obsidian's plugin settings panel

---

## Configuration

| Module | Key Settings |
|--------|-------------|
| Weather | Provider, City, API Key (Amap/OWM), Custom URL template |
| AI Insight | API URL, API Key, Model, Temperature, Delay |
| AI Language | API URL, API Key, Model, Target Language |
| Web Preview | Default URL, Zoom, X/Y Offset |
| Web Video | Default URL, Zoom, X/Y Offset |
| To-Do | Storage folder path |
| Directory | Root directories to display |

**Global Settings:** Theme, card background color & opacity, top bar visibility, module toggles, module ordering, FileViewer format toggles, utility module toggles.

---

## Requirements

- Obsidian **0.15.0** or later
- Desktop app recommended (web preview/video modules use Electron)
- Weather module: [Amap API Key](https://lbs.amap.com/) (optional — Open-Meteo and wttr.in work without any key)
- AI modules: OpenAI-compatible API endpoint

---

## License

MIT

---

## Author

**栗子仁儿 (liamzy2021)** · [GitHub](https://github.com/liamzy2021)

---

## ☕ Support

If you enjoy this plugin, feel free to support via:

<div align="center">
  <table>
    <tr>
      <td align="center" width="240">
        <a href="https://paypal.me/lilirenzy" target="_blank" rel="noopener">
          <img src="https://img.shields.io/badge/💳-PayPal-003087?style=for-the-badge&logo=paypal&logoColor=white" />
        </a>
        <br/><b>PayPal</b>
      </td>
      <td align="center" width="240">
        <a href="https://ko-fi.com/liamzy" target="_blank" rel="noopener">
          <img src="https://cdn.ko-fi.com/cdn/kofi3.png?v=3" height="36" alt="Buy Me a Coffee at ko-fi.com" />
        </a>
        <br/><b>Ko-fi</b>
      </td>
    </tr>
  </table>
</div>

---

<a name="中文说明"></a>

# 中文说明

> 一款自由拖拽的 Obsidian 仪表盘插件，内置 25 个功能模块、8 款精美主题、10+ 种文件查看器（含秒开级表格查看器，小数精度可调、多工作簿切换），以及 AI 驱动的多语言翻译——随心摆放，你的桌面你做主。

**V17** — 三大重要更新：**秒开级表格查看器**（任意大小表格即点即看，速度碾压桌面级编辑器；可调小数位数，不再眼花缭乱；多工作簿标签一键切换）、**PPT / PPT 文件查看器**（完整支持 `.ppt` 和 `.pptx` 格式，智能文本提取）和 **天气模块多平台国际化升级**（5 大平台：高德、Open-Meteo、wttr.in、OpenWeatherMap + 自定义 URL 模板）。覆盖全球天气数据，大部分平台免费无限使用。

---

## 预览

<p align="center">
  <img width="1905" alt="Modular Theme Dashboard 全貌" src="https://github.com/user-attachments/assets/09b4854a-7281-4c7c-bfc0-fdd7336e0230" />
  <br/><em>🏠 Modular Theme Dashboard — 全貌一览</em>
</p>

<p align="center">
  <img width="840" alt="仪表盘总览" src="https://github.com/user-attachments/assets/4b54c878-0e85-4c05-81e0-e26187d7cef7" />
  <br/><em>🌍 AI 驱动多语言翻译，任意语言一键切换</em>
</p>


<p align="center">
  <img width="838" alt="主题展示" src="https://github.com/user-attachments/assets/a5e14c51-569c-4d0d-8d40-a0dcaae00b92" />
  <br/><em>🎨 8 款精美主题，一键切换</em>
</p>

<p align="center">
  <img width="1677" alt="全功能展示" src="https://github.com/user-attachments/assets/91cd02b7-5500-4069-8353-8e1e976ee24c" />
  <br/><em>📊 25 个模块覆盖天气、日历、资讯、AI 分析等方方面面</em>
</p>

---

## 功能特性

### 🌍 AI 自定义语言（V17.0.8 全新功能）

> **你的 AI，你的语言，你的仪表盘。**

厌倦了只有英文的插件界面？想把仪表盘变成日文、俄文，甚至是四川话、广东话？这个功能就是为你准备的。

**怎么用：** 填入你的 OpenAI 兼容 API 密钥，输入目标语言（比如 `日本語`、`Русский`、`四川话`、`Klingon`），点击翻译——AI 会在几十秒内翻译完 500+ 条界面文字。预览结果，一键应用，你的仪表盘立刻说你的语言。

| 特点 | 详情 |
|------|------|
| 🗣️ **任意语言** | 日语、韩语、俄语、法语、德语、西班牙语、阿拉伯语、印地语、泰语、越南语…… |
| 🏠 **当地方言** | 广东话、四川话、闽南语、上海话、客家话、潮汕话……只有你想不到 |
| 🚀 **任意 AI** | OpenAI、DeepSeek、Moonshot，或任何 OpenAI 兼容接口 |
| 💾 **持久保存** | 翻译结果自动保存，重启不丢失——下次直接从下拉菜单切换 |
| 🔁 **一次翻译，永久使用** | 翻译过的语言会带 ⚡ 标记出现在语言选择器里，无需重复翻译 |

**快速上手：** 设置 → 语言 → AI 自定义 → 输入目标语言 → 点击翻译 → 预览 → 应用

### 🧩 自由拖拽布局

- **绝对定位画布** — 每个模块卡片精确放置在你想要的位置
- **拖拽移动** — 按住标题栏，拖到任意位置，实时响应
- **自由缩放** — 拖动右下角，调到你满意的尺寸
- **自动保存** — 所有位置和尺寸自动持久化，重启不跑位

### 📦 25 个内置模块

所有模块内置于插件中，无需额外下载，安装即用。

| 模块 | 功能说明 |
|------|---------|
| 天气 | 实时天气 + 未来 3 天预报（5 大平台：高德、Open-Meteo、wttr.in、OpenWeatherMap、自定义） |
| 日历 | 农历万年历，节气、节日标注，月历翻页 |
| 笔记统计 | 笔记总数、总字数、文件夹统计、Top 5 文件夹 |
| 待办事项 | 完整增删改查，三种筛选模式，进度追踪，Markdown 持久化 |
| 最近文件 | 最近修改笔记列表，智能相对时间显示 |
| 资讯 | AI HOT RSS 实时热点新闻，分类 + 分页浏览 |
| 目录 | 树形文件夹结构，可展开折叠，自定义根目录 |
| AI 洞察 | AI 驱动的笔记智能分析，当天缓存不重复请求 |
| 网页预览 | 内嵌浏览器，支持 URL 输入、缩放平移、完整沙箱 |
| 网页视频 | Electron webview 内嵌视频播放，独立会话 |
| 图片画廊 | 网格 / 瀑布流图片浏览，点击灯箱预览 |
| 媒体画廊 | 图片、视频、音频统一媒体浏览器 |
| 仓库统计 | 全面的笔记数据统计与可视化图表 |
| 代码编辑器 | 内置代码编辑器，支持语法高亮 |
| 数据编辑器 | 在仪表盘中直接编辑 JSON、YAML、CSV |
| 电子表格 | 完整电子表格编辑，支持单元格格式化 |
| 文档查看器 | Word、HTML 等格式的富文档预览 |
| HTML 查看器 | 在仪表盘中直接渲染 HTML 文件 |
| URL 打开器 | 快速启动书签和 URL |
| 小红书导入 | 从小红书导入内容 |
| Excel 转表格 | 将剪贴板 Excel 表格一键转为 Markdown |

### 🔧 实用模块

| 模块 | 说明 |
|------|------|
| 文件夹计数器 | 文件浏览器中文件夹旁显示文件数量 |
| 图片处理 | 18 项右键操作：裁剪、缩放、旋转、翻转、压缩等 |
| 表格列宽调整 | 阅读模式下拖拽调整表格列宽 |
| 自动播放循环 | 媒体自动播放循环引擎 |
| 音频转录 | 右键音视频文件一键转录（云端 API / 本地 AI）（V17.1.7 新增） |

### 🔄 无限实例化

- **任意模块可克隆** — 点击顶部 `+` 按钮，创建任意模块的多个实例
- **独立配置** — 每个实例拥有独立的设置、缓存和会话
- **智能命名** — 实例使用 `模块名#编号` 格式（如 `天气#1`、`网页预览#2`）

### 📁 内置文件查看器（10+ 种格式）

直接在仪表盘中打开文件，无需切换窗口：

| 格式 | 支持 |
|------|------|
| 表格 | `.xlsx`、`.xls`、`.csv`（SheetJS 完整渲染） |
| Word 文档 | `.docx`（mammoth.js）、`.doc`（CFB 解析） |
| PowerPoint | `.pptx`（JSZip + XML 文本提取）、`.ppt`（OLE2 二进制解析器，UTF-16LE/ASCII 双重扫描） |
| 代码与文本 | `.html`、`.txt`、`.json`、`.js` 等 |
| 图片 | `.png`、`.jpg`、`.gif`、`.webp` |
| 视频 | `.mp4`、`.webm` |

在设置中可按格式开关，禁用的格式自动回退到系统默认程序打开。

### 📊 秒开级表格查看器（V17.1.6 新功能）

> 打开任意 `.xlsx` / `.xls` 文件——小到几行表格，大到十万行巨型数据表，渲染速度碾压市面上任何表格编辑器。

| 特色 | 说明 |
|------|------|
| ⚡ **秒开不等待** | 无加载条、不卡顿——无论文件多大，毫秒级渲染 |
| 🔢 **自定义小数位** | 一键 `− / +` 调节小数精度（0-10 位），或切到"自动"保留原始值。告别 `0.093408691346` 这种眼花缭乱的长数字 |
| 📑 **多工作簿标签** | 所有 Sheet 标签在顶部列出，点哪个看哪个，即时切换 |
| 📏 **智能列宽** | 列宽按内容自适应，同时支持手动拖拽调整每列宽度 |
| 🌍 **多语言跟随** | 工具栏文字跟随插件语言设置（中文/英文/AI 自定义语言） |

<p align="center">
  <img width="2560" height="1379" alt="表格查看器概览" src="https://github.com/user-attachments/assets/f27906f8-3bcc-456a-bcc6-bc274ba91742" />
  <br/><em>📊 表格查看器——小数精度控制与多工作簿标签</em>
</p>

<p align="center">
  <img width="2555" height="1352" alt="表格查看器多工作簿" src="https://github.com/user-attachments/assets/15db28ab-eebb-4663-8af5-4299739777a7" />
  <br/><em>📑 一键切换工作簿——所有 Sheet 以标签形式展示</em>
</p>

### 🎙️ 音频/视频转录（V17.1.7 新增）

> 右键音视频文件 → 一键转录 → 文字直接插入笔记。无需仪表盘面板——配好参数即可使用。

| 特色 | 说明 |
|------|------|
| 🎯 **右键即转录** | 在文件浏览器中右键任意 `.mp3` / `.m4a` / `.wav` / `.mp4` / `.webm` / `.flac` / `.ogg` 文件 |
| ☁️ **云端 API** | 支持任何 OpenAI 兼容的 `/v1/audio/transcriptions` 接口——填入 API 地址、密钥和模型名即可 |
| 🖥️ **本地 AI** | faster-whisper-server、LocalAI、docker-whisper——任何 OpenAI 兼容 /v1/audio/transcriptions 服务 |
| 📝 **5 种输出格式** | 纯文本、项目符号列表、待办行动项、时间戳章节、说话人分离 |
| 🌍 **多语言跟随** | 所有界面文字跟随插件语言设置（中文/英文/AI 自定义语言） |
| 💰 **免费选项** | 硅基流动（TeleSpeech ASR 免费额度）、Groq（Whisper 免费额度）、Cloudflare Workers AI |

**推荐平台（均支持 OpenAI 兼容接口）：**

*云端 API：*

| 平台 | API 地址 | 说明 |
|------|----------|------|
| 硅基流动 | `https://api.siliconflow.cn/v1` | 🇨🇳 国内访问，TeleSpeech ASR 免费额度 |
| Groq | `https://api.groq.com/openai/v1` | 速度快，Whisper 免费额度 |
| Cloudflare Workers AI | `https://api.cloudflare.com/client/v4/accounts/{id}/ai/v1` | 免费额度 |
| OpenAI | `https://api.openai.com/v1` | 标准付费 API |

*本地 AI：*

| 平台 | API 地址 | 说明 |
|------|----------|------|
| faster-whisper-server | `http://localhost:8000/v1` | Python，完全 OpenAI 兼容，推荐 |
| LocalAI | `http://localhost:8080/v1` | Go，多后端（whisper/faster-whisper/llama-cpp） |
| docker-whisper | `http://localhost:9000` | Docker 一键部署，含说话人分离 |

<p align="center">
  <img width="1105" height="1003" alt="音频转录设置界面" src="https://github.com/user-attachments/assets/10424883-fa9d-4e98-93f6-857066158bbb" />
  <br/><em>🎙️ 音频转录——云端 API 配置与平台推荐</em>
</p>

<p align="center">
  <img width="442" height="462" alt="右键转录菜单" src="https://github.com/user-attachments/assets/6775045c-31a0-47ed-9d55-7d52fa5c175d" />
  <br/><em>🎯 右键音视频文件 → 转录音频/视频</em>
</p>

### 🖼️ 笔记内代码块画廊

> 在任意笔记中插入 ````t````（图片画廊）或 ````s````（媒体画廊）代码块，即可在阅读模式下渲染画廊。右键代码块可控制排版、大小、列数、间距等。

<p align="center">
  <img width="1207" alt="代码块画廊" src="https://github.com/user-attachments/assets/5efa44a2-ccdc-41d4-9461-eed7d1b43525" />
  <br/><em>📸 笔记内画廊 — 间距控制 + 智能居中</em>
</p>

#### 图片画廊（```t）

`````markdown
```t
/path/to/images|horizontal|4|200|0|0|10|12|true
```
`````

**快速格式**（管道分隔）：

| 位置 | 字段 | 默认值 | 说明 |
|------|------|--------|------|
| 1 | `path` | — | 图片文件夹路径（必填） |
| 2 | `layout` | `horizontal` | `horizontal`（水平）/ `vertical`（瀑布流）/ `grid`（网格） |
| 3 | `columns` | `4` | 每行图片数量 |
| 4 | `height` | `200` | 图片高度（px） |
| 5 | `width` | `0` | 图片宽度（px，0=自适应） |
| 6 | `spacingLeft` | `0` | 左侧距离（px） |
| 7 | `spacingRight` | `0` | 右侧距离（px） |
| 8 | `itemGap` | `12` | 图片间距（px） |
| 9 | `smartCenter` | `false` | `true` 启用智能居中 |

**键值格式**（更易读）：

```
path: /path/to/images
type: horizontal
columns: 4
height: 200
spacingleft: 10
spacingright: 10
itemgap: 12
smartcenter: true
```

**支持格式：** jpeg、jpg、gif、png、webp、tiff、tif、avif、bmp

**右键菜单：** 编辑 / 删除 / 自定义排版 / 自定义大小 / 自定义每行数量 / 间距设置（含智能居中开关）

#### 媒体画廊（```s）

`````markdown
```s
/path/to/media|grid|220|10|10|12|true
```
`````

**快速格式**（管道分隔）：

| 位置 | 字段 | 默认值 | 说明 |
|------|------|--------|------|
| 1 | `path` | — | 媒体文件夹路径（必填） |
| 2 | `type` | `grid` | `grid`（网格）/ `list`（列表）/ `full`（全宽） |
| 3 | `size` | `220` | 媒体格子大小（px） |
| 4 | `spacingLeft` | `0` | 左侧距离（px） |
| 5 | `spacingRight` | `0` | 右侧距离（px） |
| 6 | `itemGap` | `10` | 媒体间距（px） |
| 7 | `smartCenter` | `false` | `true` 启用智能居中 |

**键值格式：**

```
path: /path/to/media1, /path/to/media2
type: grid
size: 220
sort: name
limit: 50
spacingleft: 10
spacingright: 10
itemgap: 12
smartcenter: true
```

**支持格式：** 图片（jpeg、jpg、gif、png、webp、tiff、svg、ico、heic、avif、bmp）· 视频（mp4、webm、ogg、mov、mkv、avi）· 音频（mp3、wav、flac、aac、m4a、ogg）

**右键菜单：** 编辑 / 删除 / 自定义排版 / 自定义大小 / 间距设置（含智能居中开关）

### 🎨 8 款精美主题

| 主题 | 风格 |
|------|------|
| Dawn（晨曦） | 暖橙色调，温暖柔和 |
| Sabi（侘寂） | 柔和绿灰，日式简约 |
| Dusk（暮光） | 深蓝紫调，暗色优雅 |
| Coastal（海岸） | 青绿清新，自然舒适 |
| Harvest（丰收） | 金棕暖调，秋日丰收 |
| Ink（墨迹） | 深灰冷调，墨色沉稳 |
| Linen（亚麻） | 米色布纹，低调温暖 |
| Carbon（碳灰） | 纯黑背景，现代科技 |

点击顶部工具栏 `🎨` 按钮一键切换，也可在设置中自定义卡片背景色和透明度。

### 💾 设置存档

- **导出** — 将所有设置保存为 `obsidian-dashboard-settings-YYYY-MM-DD.json`
- **导入** — 从之前导出的文件一键恢复
- 适合在多仓库间迁移配置或与朋友分享设置

---

## 安装

### 从 GitHub 手动安装

1. 前往 [Releases](https://github.com/liamzy2021/Obsidian--Modular-Theme-Dashboard-Free-Drag-and-Drop/releases) 下载最新的 `main.js`、`manifest.json` 和 `styles.css`
2. 在 Obsidian 仓库的 `.obsidian/plugins/` 目录下新建 `modular-theme-dashboard` 文件夹
3. 将 3 个文件复制进去：
   ```
   .obsidian/plugins/modular-theme-dashboard/
   ├── main.js
   ├── manifest.json
   └── styles.css
   ```
4. 重启 Obsidian 或重新加载插件
5. **设置 → 社区插件 → 启用 "Modular Theme Dashboard"**

### 从社区插件商店安装

1. 打开 **设置 → 社区插件 → 浏览**
2. 搜索 "Modular Theme Dashboard"
3. 点击 **安装**，然后 **启用**

---

## 使用说明

- **打开仪表盘** — 左侧边栏点击 🏠 图标，或通过命令面板搜索
- **拖拽卡片** — 按住标题栏拖拽到任意位置
- **调整大小** — 拖动卡片右下角
- **添加模块** — 点击顶部 `+` 按钮，选择要添加的模块
- **切换主题** — 点击顶部 `🎨` 按钮
- **模块设置** — 点击顶部 `⚙️` 按钮，或通过 Obsidian 设置面板

---

## 配置说明

| 模块 | 关键设置 |
|------|---------|
| 天气 | 平台选择、城市、API Key（高德/OWM）、自定义 URL 模板 |
| AI 洞察 | API URL、API Key、模型、温度参数、请求延迟 |
| AI 语言 | API URL、API Key、模型、目标语言 |
| 网页预览 | 默认 URL、缩放比例、XY 偏移 |
| 网页视频 | 默认 URL、缩放比例、XY 偏移 |
| 待办事项 | 存储文件夹路径 |
| 目录 | 显示的根目录列表 |

**全局设置：** 主题、卡片背景色和透明度、顶栏显示/隐藏、模块开关、模块排序、文件查看器格式开关、实用模块开关。

---

## 系统要求

- Obsidian **0.15.0** 或更高版本
- 推荐桌面端使用（网页预览/视频模块依赖 Electron）
- 天气模块可选：[高德地图 API Key](https://lbs.amap.com/)（也可使用 Open-Meteo 或 wttr.in，无需任何 API Key）
- AI 模块需要 OpenAI 兼容的 API 接口

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

<div align="center">
  <table>
    <tr>
      <td align="center" width="280">
        <a href="https://ifdian.net/a/liamzy2021" target="_blank" rel="noopener">
          <img src="https://static.afdiancdn.com/static/img/logo/logo.png" width="80" alt="爱发电 Afdian" />
          <br/>
          <b>❤️ 前往爱发电支持</b>
        </a>
      </td>
      <td align="center" width="280">
        <img src="https://img-reg-ab.imagency.cn/e/19467f4b916c082ee6ef3b9d81aa9ecb.png" width="200" alt="微信赞赏" />
        <br/>
        <b>微信赞赏</b>
      </td>
    </tr>
  </table>
</div>
