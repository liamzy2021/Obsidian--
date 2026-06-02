/**
 * AI Dashboard V15
 * 底层：V14 自由拖拽/resize 布局 + 完全模块化架构
 * 功能：V11 完整功能迁移 + 无限实例化系统（所有模块默认可克隆）
 * 主题：V11 8个精美主题
 */

const { Plugin, ItemView, Setting, PluginSettingTab, Modal, Notice, setIcon, requestUrl, moment } = require('obsidian');

const VIEW_TYPE = 'ai-dashboard-v15';

// ===================== 主题配置 =====================
const THEMES = {
    dawn: {
        name: '晨曦',
        primary: '#e8956d',
        secondary: '#f0b27a',
        accent: '#d4785a',
        bg: '#fdf6f0',
        card: '#ffffff',
        text: '#3d2b1f',
        muted: '#9c7b6e'
    },
    sabi: {
        name: '侘寂',
        primary: '#8b9e87',
        secondary: '#a8b9a4',
        accent: '#6b7e67',
        bg: '#f2f0eb',
        card: '#faf8f5',
        text: '#2c2c2c',
        muted: '#8a8279'
    },
    dusk: {
        name: '暮光',
        primary: '#7986cb',
        secondary: '#9fa8da',
        accent: '#5c6bc0',
        bg: '#1a1b2e',
        card: '#22243d',
        text: '#e8eaf6',
        muted: '#7986cb'
    },
    coastal: {
        name: '海岸',
        primary: '#4db6ac',
        secondary: '#80cbc4',
        accent: '#26a69a',
        bg: '#f0f8f7',
        card: '#ffffff',
        text: '#1a3c40',
        muted: '#6db3ae'
    },
    harvest: {
        name: '丰收',
        primary: '#c49a3c',
        secondary: '#d4b06a',
        accent: '#a07c28',
        bg: '#faf5e8',
        card: '#fff9ed',
        text: '#2d2010',
        muted: '#9a7d42'
    },
    ink: {
        name: '墨迹',
        primary: '#546e7a',
        secondary: '#78909c',
        accent: '#37474f',
        bg: '#1c1f24',
        card: '#252930',
        text: '#eceff1',
        muted: '#78909c'
    },
    linen: {
        name: '亚麻',
        primary: '#a0856c',
        secondary: '#c4a882',
        accent: '#7d6455',
        bg: '#f8f2ea',
        card: '#fdfaf5',
        text: '#2e2219',
        muted: '#a08060'
    },
    carbon: {
        name: '碳灰',
        primary: '#64b5f6',
        secondary: '#90caf9',
        accent: '#42a5f5',
        bg: '#121212',
        card: '#1e1e1e',
        text: '#eeeeee',
        muted: '#757575'
    }
};

// ===================== 默认设置 =====================
const DEFAULT_SETTINGS = {
    theme: 'ink',
    layout: {},
    modules: {
        weather: {
            enabled: true,
            city: '北京',
            apiKey: ''
        },
        calendar: {
            enabled: true,
            showLunar: true,
            showHoliday: true
        },
        stats: {
            enabled: true,
            showFileCount: true,
            showWordCount: true
        },
        todo: {
            enabled: true,
            folder: '待办'
        },
        recent: {
            enabled: true,
            maxFiles: 10
        },
        news: {
            enabled: true,
            source: 'aihot',
            pageSize: 10
        },
        directory: {
            enabled: true,
            folders: [],
            expandedNodes: []
        },
        'ai-insight': {
            enabled: true,
            apiKey: '',
            apiUrl: 'https://api.openai.com/v1/chat/completions',
            model: 'gpt-3.5-turbo',
            temperature: 0.7,
            requestDelay: 0
        },
        'web-preview': {
            enabled: true,
            url: 'https://www.baidu.com',
            zoom: 1,
            posX: 0,
            posY: 0
        },
        'web-video': {
            enabled: true,
            url: 'https://www.bilibili.com',
            zoom: 1,
            posX: 0,
            posY: 0
        },
    },
    // 实例列表：每个实例 { id: 'weather#1', baseModule: 'weather', label: '天气 1' }
    instances: [],
    instanceCounter: 0,
    moduleOrder: ['weather', 'calendar', 'stats', 'todo', 'recent', 'news', 'directory', 'ai-insight', 'web-preview', 'web-video'],
    headerBg: '',
    showHeader: true,
    cardBgColor: '',
    cardBgOpacity: 0.95
};

// ===================== 模块管理器 =====================
class ModuleManager {
    constructor(plugin) {
        this.plugin = plugin;
        this.modules = new Map();
        this._runtimeCtx = {};
        this._moduleDefaults = new Map(); // moduleId -> defaultSettings
    }

    async loadModules() {
        this.modules.clear();
        this._moduleDefaults.clear();
        const modulesDir = this.plugin.manifest.dir + '/modules/';

        let files = [];
        try {
            files = await this.plugin.app.vault.adapter.list(modulesDir);
        } catch (e) {
            console.warn('[V15] 扫描模块目录失败:', e);
            return;
        }

        const jsFiles = (files.files || []).filter(f => f.endsWith('.js'));

        for (const filePath of jsFiles) {
            const moduleId = filePath.replace(modulesDir, '').replace('.js', '');
            try {
                const content = await this.plugin.app.vault.adapter.read(filePath);
                const mod = this._evalModule(content, moduleId);
                if (mod && mod.render) {
                    const id = mod.id || moduleId;
                    this.modules.set(id, mod);
                    // 捕获模块导出的默认设置
                    if (mod.defaultSettings && typeof mod.defaultSettings === 'object') {
                        this._moduleDefaults.set(id, mod.defaultSettings);
                    }
                }
            } catch (e) {
                console.warn('[V15] 模块 ' + moduleId + ' 加载失败:', e);
            }
        }

        console.log('[V15] 已加载模块: ' + [...this.modules.keys()].join(', '));
    }

    getLoadedModuleIds() {
        return [...this.modules.keys()];
    }

    // 获取模块的默认设置（从模块导出或全局默认）
    getModuleDefaultSettings(baseId) {
        if (this._moduleDefaults.has(baseId)) {
            return JSON.parse(JSON.stringify(this._moduleDefaults.get(baseId)));
        }
        if (DEFAULT_SETTINGS.modules[baseId]) {
            const defs = { ...DEFAULT_SETTINGS.modules[baseId] };
            delete defs.enabled; // 实例单独控制
            return defs;
        }
        return {};
    }

    // 解析模块 ID：实例 ID（如 web-preview#1）→ 基础模块代码
    resolveModule(moduleId) {
        if (this.modules.has(moduleId)) return this.modules.get(moduleId);
        const hashIdx = moduleId.indexOf('#');
        if (hashIdx > 0) {
            const base = moduleId.substring(0, hashIdx);
            return this.modules.get(base) || null;
        }
        return null;
    }

    // 获取基础模块 ID（weather#1 → weather）
    resolveBaseModuleId(moduleId) {
        const hashIdx = moduleId.indexOf('#');
        return hashIdx > 0 ? moduleId.substring(0, hashIdx) : moduleId;
    }

    // 获取实例信息
    getInstanceInfo(moduleId) {
        const instances = this.plugin.settings.instances || [];
        return instances.find(i => i.id === moduleId) || null;
    }

    _evalModule(code, fallbackId) {
        try {
            const moduleExports = {};
            const module = { exports: moduleExports };
            const exports = moduleExports;
            const _require = (pkg) => {
                if (pkg === 'obsidian') return require('obsidian');
                throw new Error('Unknown module: ' + pkg);
            };

            const fn = new Function(
                'module', 'exports', 'require',
                'app', 'plugin', 'moment', 'Notice', 'requestUrl', 'setIcon',
                '_runtimeCtx',
                'with (_runtimeCtx) {\n' + code + '\n}'
            );

            fn(
                module, exports, _require,
                this.plugin.app, this.plugin, moment, Notice, requestUrl, setIcon,
                this._runtimeCtx
            );

            if (module.exports && typeof module.exports === 'object' &&
                Object.keys(module.exports).length > 0) {
                return module.exports;
            }
            return exports;
        } catch (e) {
            console.error('[V15] 模块解析错误 (' + fallbackId + '):', e);
            return null;
        }
    }

    createContext(moduleId) {
        const plugin = this.plugin;
        const app = plugin.app;
        const baseId = this.resolveBaseModuleId(moduleId);

        // 确保设置条目存在，且包含完整默认值
        if (!plugin.settings.modules[moduleId]) {
            const defaults = this.getModuleDefaultSettings(baseId);
            plugin.settings.modules[moduleId] = {
                enabled: true,
                ...defaults
            };
        }

        const getAllFiles = () => app.vault.getMarkdownFiles();
        const getRecentFiles = (n = 10) => app.vault.getMarkdownFiles()
            .sort((a, b) => b.stat.mtime - a.stat.mtime).slice(0, n);
        const getFilesInFolder = (path) => app.vault.getMarkdownFiles()
            .filter(f => f.path.startsWith(path));

        const saveCallback = async () => {
            await plugin.saveSettings();
        };

        return {
            plugin,
            app,
            moment,
            Notice,
            Setting,
            requestUrl,
            setIcon,
            saveCallback,
            settings: plugin.settings.modules[moduleId],
            theme: THEMES[plugin.settings.theme] || THEMES.dawn,
            data: { getAllFiles, getRecentFiles, getFilesInFolder },
            _moduleId: moduleId
        };
    }

    async renderModule(moduleId, container) {
        const mod = this.resolveModule(moduleId);
        if (!mod) {
            container.createEl('div', {
                text: '模块 "' + moduleId + '" 未加载',
                attr: { style: 'color: var(--text-muted); text-align: center; padding: 20px;' }
            });
            return;
        }

        // 注入内联样式（实例共享基础模块样式）
        const styleModuleId = moduleId.indexOf('#') > 0 ? moduleId.substring(0, moduleId.indexOf('#')) : moduleId;
        if (mod.styles) {
            const styleId = 'v15-module-style-' + styleModuleId;
            let styleEl = document.getElementById(styleId);
            if (!styleEl) {
                styleEl = document.createElement('style');
                styleEl.id = styleId;
                document.head.appendChild(styleEl);
            }
            styleEl.textContent = mod.styles;
        }

        // 更新运行时上下文
        const ctx = this.createContext(moduleId);
        Object.assign(this._runtimeCtx, ctx);

        try {
            await mod.render(container);
        } catch (e) {
            console.error('[V15] 模块 ' + moduleId + ' 渲染错误:', e);
            container.createEl('div', {
                text: '渲染失败: ' + e.message,
                attr: { style: 'color: var(--text-muted); font-size: 12px; padding: 10px;' }
            });
        }
    }

    getModule(moduleId) {
        return this.resolveModule(moduleId);
    }

    getAllModules() {
        return [...this.modules.values()];
    }
}

// ===================== 仪表盘视图 =====================
class DashboardView extends ItemView {
    constructor(leaf, plugin) {
        super(leaf);
        this.plugin = plugin;
        this.draggedCard = null;
        this.dragOffset = { x: 0, y: 0 };
    }

    getViewType() { return VIEW_TYPE; }
    getDisplayText() { return 'Modular Theme Dashboard'; }
    getIcon() { return 'layout-dashboard'; }

    async onOpen() {
        this.containerEl.empty();
        this.containerEl.addClass('v15-view');

        this.registerDomEvent(document, 'mousemove', (e) => this._onMouseMove(e));
        this.registerDomEvent(document, 'mouseup', (e) => this._onMouseUp(e));

        await this.render();
    }

    async onClose() {}

    applyTheme() {
        const theme = THEMES[this.plugin.settings.theme] || THEMES.dawn;
        const root = this.containerEl;
        root.style.setProperty('--v6-primary', theme.primary);
        root.style.setProperty('--v6-secondary', theme.secondary);
        root.style.setProperty('--v6-accent', theme.accent);
        root.style.setProperty('--v6-bg', theme.bg);
        root.style.setProperty('--v6-text', theme.text);
        root.style.setProperty('--v6-muted', theme.muted);

        const customBg = this.plugin.settings.cardBgColor;
        const opacity = this.plugin.settings.cardBgOpacity != null
            ? this.plugin.settings.cardBgOpacity
            : 0.95;
        if (customBg) {
            root.style.setProperty('--v6-card', this._hexToRgba(customBg, opacity));
        } else {
            root.style.setProperty('--v6-card', this._hexToRgba(theme.card, opacity));
        }

        root.style.setProperty('--background-primary', theme.bg);
        root.style.setProperty('--text-normal', theme.text);
        root.style.setProperty('--text-muted', theme.muted);
        root.style.setProperty('--interactive-accent', theme.primary);
    }

    _hexToRgba(hex, alpha) {
        if (!hex || typeof hex !== 'string') return `rgba(255,255,255,${alpha})`;
        if (hex.startsWith('rgba') || hex.startsWith('rgb')) return hex;
        let h = hex.replace('#', '');
        if (h.length === 3) h = h.split('').map(c => c + c).join('');
        if (h.length !== 6) return hex;
        const r = parseInt(h.substring(0, 2), 16);
        const g = parseInt(h.substring(2, 4), 16);
        const b = parseInt(h.substring(4, 6), 16);
        if (isNaN(r) || isNaN(g) || isNaN(b)) return hex;
        return `rgba(${r},${g},${b},${alpha})`;
    }

    async render() {
        const { containerEl } = this;
        containerEl.empty();
        containerEl.addClass('v15-view');
        this.applyTheme();

        if (this.plugin.settings.showHeader !== false) {
            this._renderHeader(containerEl);
        }

        const canvas = containerEl.createDiv({ cls: 'v6-canvas' });

        if (this.plugin.moduleManager.modules.size === 0) {
            await this.plugin.moduleManager.loadModules();
        }

        // 收集所有要渲染的模块 ID（基础模块 + 实例）
        const loadedIds = this.plugin.moduleManager.getLoadedModuleIds();
        const instances = this.plugin.settings.instances || [];
        const instanceIds = instances.map(i => i.id);

        const moduleOrder = this.plugin.settings.moduleOrder || [];

        // 按照 moduleOrder 顺序渲染，然后是未在 order 中的
        const rendered = new Set();

        // 先按 moduleOrder 渲染
        for (const moduleId of moduleOrder) {
            if (rendered.has(moduleId)) continue;
            if (!loadedIds.includes(moduleId) && !instanceIds.includes(moduleId)) continue;

            const modSettings = this.plugin.settings.modules[moduleId];
            if (!modSettings || modSettings.enabled === false) continue;

            const mod = this.plugin.moduleManager.resolveModule(moduleId);
            if (!mod) continue;

            this.renderModuleCard(canvas, moduleId, mod);
            rendered.add(moduleId);
        }

        // 渲染剩余的基础模块（不在 moduleOrder 中的）
        for (const moduleId of loadedIds) {
            if (rendered.has(moduleId)) continue;
            const modSettings = this.plugin.settings.modules[moduleId];
            if (!modSettings || modSettings.enabled === false) continue;
            const mod = this.plugin.moduleManager.getModule(moduleId);
            if (!mod) continue;
            this.renderModuleCard(canvas, moduleId, mod);
            rendered.add(moduleId);
        }

        // 渲染剩余的实例（不在 moduleOrder 中的）
        for (const inst of instances) {
            if (rendered.has(inst.id)) continue;
            const modSettings = this.plugin.settings.modules[inst.id];
            if (!modSettings || modSettings.enabled === false) continue;
            const mod = this.plugin.moduleManager.resolveModule(inst.id);
            if (!mod) continue;
            this.renderModuleCard(canvas, inst.id, mod);
            rendered.add(inst.id);
        }
    }

    _renderHeader(parent) {
        const header = parent.createDiv({ cls: 'v15-header' });

        const left = header.createDiv({ cls: 'v15-header-left' });
        left.createEl('span', { text: '🏠', cls: 'v15-header-icon' });
        left.createEl('span', { text: 'Modular Theme Dashboard', cls: 'v15-header-title' });

        const right = header.createDiv({ cls: 'v15-header-right' });

        // ★ 新增：添加板块按钮
        const addBtn = right.createEl('button', {
            cls: 'v15-header-btn',
            attr: { title: '添加板块' }
        });
        addBtn.innerHTML = '➕';
        addBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this._showAddMenu(addBtn);
        });

        // 主题切换
        const themeBtn = right.createEl('button', {
            cls: 'v15-header-btn',
            attr: { title: '切换主题' }
        });
        themeBtn.innerHTML = '🎨';
        themeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this._showThemeMenu(themeBtn);
        });

        // 刷新
        const refreshBtn = right.createEl('button', {
            cls: 'v15-header-btn',
            attr: { title: '刷新' }
        });
        refreshBtn.innerHTML = '🔄';
        refreshBtn.addEventListener('click', () => this.render());

        // 设置
        const settingsBtn = right.createEl('button', {
            cls: 'v15-header-btn',
            attr: { title: '设置' }
        });
        settingsBtn.innerHTML = '⚙️';
        settingsBtn.addEventListener('click', () => {
            try {
                this.plugin.app.setting.open();
            } catch (e) {
                console.warn('[V15] 打开设置失败:', e);
            }
        });
    }

    _showAddMenu(anchor) {
        const existing = document.querySelector('.v15-add-menu');
        if (existing) { existing.remove(); return; }

        const menu = document.createElement('div');
        menu.className = 'v15-add-menu';
        document.body.appendChild(menu);

        const rect = anchor.getBoundingClientRect();
        menu.style.cssText =
            'position:fixed;' +
            'top:' + (rect.bottom + 4) + 'px;' +
            'right:' + (window.innerWidth - rect.right) + 'px;' +
            'z-index:9999;' +
            'background:var(--background-primary);' +
            'border:1px solid var(--background-modifier-border);' +
            'border-radius:8px;' +
            'padding:6px;' +
            'min-width:180px;' +
            'max-height:400px;' +
            'overflow-y:auto;' +
            'box-shadow:0 4px 16px rgba(0,0,0,0.15);';

        menu.createEl('div', {
            text: '添加板块（所有模块均可添加多个）',
            attr: { style: 'padding:4px 8px;font-size:11px;color:var(--text-muted);font-weight:bold;' }
        });

        const allModules = this.plugin.moduleManager.getAllModules();
        allModules.forEach(mod => {
            const baseId = mod.id;
            if (!baseId) return;
            const item = menu.createEl('div', {
                text: (mod.icon || '📦') + ' ' + (mod.title || baseId),
                attr: {
                    style: 'padding:8px 10px;border-radius:6px;cursor:pointer;font-size:13px;display:flex;align-items:center;gap:6px;transition:background 0.15s;'
                }
            });
            item.addEventListener('mouseenter', () => item.style.background = 'var(--background-modifier-hover)');
            item.addEventListener('mouseleave', () => item.style.background = '');
            item.addEventListener('click', async () => {
                menu.remove();
                await this._addInstance(baseId);
            });
        });

        const dismiss = (e) => {
            if (!menu.contains(e.target)) {
                menu.remove();
                document.removeEventListener('click', dismiss);
            }
        };
        setTimeout(() => document.addEventListener('click', dismiss), 0);
    }

    async _addInstance(baseModule) {
        const mod = this.plugin.moduleManager.getModule(baseModule);
        if (!mod) {
            new Notice('模块 ' + baseModule + ' 未加载');
            return;
        }

        const instances = this.plugin.settings.instances || [];
        const counter = (this.plugin.settings.instanceCounter || 0) + 1;
        this.plugin.settings.instanceCounter = counter;

        const instanceId = baseModule + '#' + counter;
        const label = (mod.title || baseModule) + ' ' + counter;

        instances.push({ id: instanceId, baseModule: baseModule, label: label });
        this.plugin.settings.instances = instances;

        // 创建实例设置（深拷贝默认值）
        const defaults = this.plugin.moduleManager.getModuleDefaultSettings(baseModule);
        this.plugin.settings.modules[instanceId] = {
            enabled: true,
            ...JSON.parse(JSON.stringify(defaults))
        };

        // 添加到 moduleOrder
        if (!this.plugin.settings.moduleOrder) {
            this.plugin.settings.moduleOrder = [];
        }
        this.plugin.settings.moduleOrder.push(instanceId);

        // 分配默认布局：从顶部开始排列，避免新卡片出现在屏幕外
        if (!this.plugin.settings.layout[instanceId]) {
            const idx = instances.length - 1;
            const cols = 4;
            this.plugin.settings.layout[instanceId] = {
                x: 20 + (idx % cols) * 320,
                y: 80 + Math.floor(idx / cols) * 270,
                width: 300,
                height: 250
            };
        }

        await this.plugin.saveSettings();
        new Notice('已添加: ' + label);
        this.render();
    }

    async _removeInstance(instanceId) {
        const instances = this.plugin.settings.instances || [];
        const idx = instances.findIndex(i => i.id === instanceId);
        if (idx === -1) return;

        const label = instances[idx].label;
        instances.splice(idx, 1);
        this.plugin.settings.instances = instances;

        // 删除实例设置和布局
        delete this.plugin.settings.modules[instanceId];
        delete this.plugin.settings.layout[instanceId];

        // 从 moduleOrder 移除
        if (this.plugin.settings.moduleOrder) {
            this.plugin.settings.moduleOrder = this.plugin.settings.moduleOrder.filter(id => id !== instanceId);
        }

        await this.plugin.saveSettings();
        new Notice('已移除: ' + label);
        this.render();
    }

    _showThemeMenu(anchor) {
        const existing = document.querySelector('.v15-theme-menu');
        if (existing) { existing.remove(); return; }

        const menu = document.createElement('div');
        menu.className = 'v15-theme-menu';
        document.body.appendChild(menu);

        const rect = anchor.getBoundingClientRect();
        menu.style.cssText =
            'position:fixed;' +
            'top:' + (rect.bottom + 4) + 'px;' +
            'right:' + (window.innerWidth - rect.right) + 'px;' +
            'z-index:9999;' +
            'background:var(--background-primary);' +
            'border:1px solid var(--background-modifier-border);' +
            'border-radius:8px;' +
            'padding:8px;' +
            'display:grid;' +
            'grid-template-columns:repeat(4,1fr);' +
            'gap:6px;' +
            'box-shadow:0 4px 16px rgba(0,0,0,0.15);';

        Object.entries(THEMES).forEach(([id, t]) => {
            const btn = document.createElement('button');
            btn.title = t.name;
            btn.style.cssText =
                'width:36px;height:36px;border-radius:50%;' +
                'border:2px solid ' + (this.plugin.settings.theme === id ? t.primary : 'transparent') + ';' +
                'background:' + t.primary + ';cursor:pointer;outline:none;' +
                'transition:transform 0.15s;';
            btn.addEventListener('click', async () => {
                this.plugin.settings.theme = id;
                await this.plugin.saveSettings();
                menu.remove();
                this.render();
            });
            menu.appendChild(btn);
        });

        const dismiss = (e) => {
            if (!menu.contains(e.target)) {
                menu.remove();
                document.removeEventListener('click', dismiss);
            }
        };
        setTimeout(() => document.addEventListener('click', dismiss), 0);
    }

    renderModuleCard(canvas, moduleId, mod) {
        let layout = this.plugin.settings.layout[moduleId];
        const defaults = this._defaultLayout(moduleId);
        const x = layout && layout.x != null ? layout.x : defaults.x;
        const y = layout && layout.y != null ? layout.y : defaults.y;
        const width = layout && layout.width >= 200 ? layout.width : defaults.width;
        const height = layout && layout.height >= 150 ? layout.height : defaults.height;

        const card = canvas.createDiv({ cls: 'v6-card' });
        card.dataset.moduleId = moduleId;
        card.style.left = x + 'px';
        card.style.top = y + 'px';
        card.style.width = width + 'px';
        card.style.height = height + 'px';
        card.style.resize = 'both';
        card.style.overflow = 'hidden';
        card.style.minWidth = '200px';
        card.style.minHeight = '150px';

        // 判断是否为实例
        const isInstance = moduleId.indexOf('#') > 0;
        const instanceInfo = isInstance ? this.plugin.moduleManager.getInstanceInfo(moduleId) : null;

        // 卡片头部
        const cardHeader = card.createDiv({ cls: 'v6-card-header' });
        const titleArea = cardHeader.createDiv({ cls: 'v6-card-title' });

        if (isInstance && instanceInfo) {
            const baseMod = this.plugin.moduleManager.getModule(instanceInfo.baseModule);
            titleArea.createEl('span', { text: (baseMod ? baseMod.icon : '📦') + ' ' + instanceInfo.label, cls: 'v6-card-label' });
        } else {
            titleArea.createEl('span', { text: mod.icon || '📦', cls: 'v6-card-icon' });
            titleArea.createEl('span', { text: mod.title || moduleId, cls: 'v6-card-label' });
        }

        // 刷新按钮
        const refreshBtn = cardHeader.createEl('button', {
            cls: 'v6-card-btn',
            attr: { title: '刷新' }
        });
        refreshBtn.innerHTML = '↺';
        refreshBtn.addEventListener('click', async (e) => {
            e.stopPropagation();
            content.empty();
            await this.plugin.moduleManager.renderModule(moduleId, content);
        });

        // 实例：添加移除按钮
        if (isInstance) {
            const removeBtn = cardHeader.createEl('button', {
                cls: 'v6-card-btn',
                attr: { title: '移除此板块' }
            });
            removeBtn.innerHTML = '✕';
            removeBtn.style.color = 'var(--text-error)';
            removeBtn.addEventListener('click', async (e) => {
                e.stopPropagation();
                await this._removeInstance(moduleId);
            });
        }

        // 内容区域
        const content = card.createDiv({ cls: 'v6-card-content' });
        content.style.overflow = 'auto';
        content.style.height = 'calc(100% - 50px)';

        this.plugin.moduleManager.renderModule(moduleId, content);

        // 拖拽
        cardHeader.addEventListener('mousedown', (e) => this._onDragStart(e, card));

        // resize 保存
        let lastWidth = width;
        let lastHeight = height;
        const saveSize = () => {
            const newWidth = parseInt(card.style.width) || card.offsetWidth;
            const newHeight = parseInt(card.style.height) || card.offsetHeight;
            if (newWidth !== lastWidth || newHeight !== lastHeight) {
                lastWidth = newWidth;
                lastHeight = newHeight;
                this._saveLayout(moduleId, card);
            }
        };
        card.addEventListener('mouseup', saveSize);
        const globalMouseUp = (e) => {
            if (card.contains(e.target)) {
                setTimeout(saveSize, 50);
            }
        };
        document.addEventListener('mouseup', globalMouseUp);
        this.register(() => document.removeEventListener('mouseup', globalMouseUp));
    }

    _defaultLayout(moduleId) {
        const defaults = {
            weather:      { x: 20,  y: 20,  width: 300, height: 280 },
            calendar:     { x: 340, y: 20,  width: 340, height: 380 },
            stats:        { x: 700, y: 20,  width: 300, height: 280 },
            todo:         { x: 20,  y: 320, width: 300, height: 360 },
            recent:       { x: 340, y: 420, width: 340, height: 280 },
            news:         { x: 700, y: 320, width: 340, height: 360 },
            directory:    { x: 20,  y: 700, width: 300, height: 360 },
            'ai-insight': { x: 340, y: 720, width: 700, height: 300 },
            'web-preview':{ x: 20,  y: 1080,width: 500, height: 400 },
            'web-video':  { x: 540, y: 1080,width: 500, height: 400 }
        };
        return defaults[moduleId] || { x: 20, y: 20, width: 300, height: 280 };
    }

    _onDragStart(e, card) {
        if (e.target.closest('.v6-card-btn')) return;
        this.draggedCard = card;
        this.dragOffset = {
            x: e.clientX - card.offsetLeft,
            y: e.clientY - card.offsetTop
        };
        card.style.zIndex = '100';
        card.style.transition = 'none';
        e.preventDefault();
    }

    _onMouseMove(e) {
        if (!this.draggedCard) return;
        const x = e.clientX - this.dragOffset.x;
        const y = e.clientY - this.dragOffset.y;
        this.draggedCard.style.left = Math.max(0, x) + 'px';
        this.draggedCard.style.top = Math.max(0, y) + 'px';
    }

    async _onMouseUp(e) {
        if (this.draggedCard) {
            const moduleId = this.draggedCard.dataset.moduleId;
            this._saveLayout(moduleId, this.draggedCard);
            this.draggedCard.style.zIndex = '';
            this.draggedCard.style.transition = '';
            this.draggedCard = null;
        }
    }

    _saveLayout(moduleId, card) {
        const computedStyle = window.getComputedStyle(card);
        const width = parseInt(computedStyle.width) || 300;
        const height = parseInt(computedStyle.height) || 250;
        this.plugin.settings.layout[moduleId] = {
            x: parseInt(card.style.left) || 0,
            y: parseInt(card.style.top) || 0,
            width: Math.max(width, 200),
            height: Math.max(height, 150)
        };
        this.plugin.saveSettings();
    }
}

// ===================== 设置面板 =====================
class DashboardSettingTab extends PluginSettingTab {
    constructor(app, plugin) {
        super(app, plugin);
        this.plugin = plugin;
        this._currentModuleId = null;
    }

    display() {
        const { containerEl } = this;
        containerEl.empty();

        containerEl.createEl('h2', { text: 'Modular Theme Dashboard V15 设置' });

        this._renderAppearanceSettings(containerEl);
        this._renderModuleToggles(containerEl);
        this._renderInstanceManager(containerEl);

        if (this._currentModuleId) {
            this._renderModuleSettings(containerEl, this._currentModuleId);
        }

        // 设置面板底部固定展示打赏（可通过模块设置关闭）
        this._renderDonateSection(containerEl);
    }

    _renderAppearanceSettings(containerEl) {
        containerEl.createEl('h3', { text: '外观' });

        new Setting(containerEl)
            .setName('主题')
            .setDesc('选择仪表盘主题风格')
            .addDropdown(d => {
                Object.entries(THEMES).forEach(([id, t]) => d.addOption(id, t.name));
                d.setValue(this.plugin.settings.theme)
                    .onChange(async (v) => {
                        this.plugin.settings.theme = v;
                        await this.plugin.saveSettings();
                        this.plugin.refreshView();
                    });
            });

        new Setting(containerEl)
            .setName('显示顶栏')
            .setDesc('显示或隐藏仪表盘顶部工具栏')
            .addToggle(t => {
                t.setValue(this.plugin.settings.showHeader !== false)
                    .onChange(async (v) => {
                        this.plugin.settings.showHeader = v;
                        await this.plugin.saveSettings();
                        this.plugin.refreshView();
                    });
            });

        new Setting(containerEl)
            .setName('卡片背景颜色')
            .setDesc('自定义卡片背景色，留空则使用主题默认')
            .addText(t => {
                t.setPlaceholder('#1a1a1a 或 #ffffff')
                    .setValue(this.plugin.settings.cardBgColor || '')
                    .onChange(async (v) => {
                        this.plugin.settings.cardBgColor = v;
                        await this.plugin.saveSettings();
                        this.plugin.refreshView();
                    });
            });

        new Setting(containerEl)
            .setName('卡片背景透明度')
            .setDesc('0 = 完全透明，1 = 完全不透明')
            .addSlider(s => {
                s.setLimits(0, 1, 0.05)
                    .setValue(this.plugin.settings.cardBgOpacity != null ? this.plugin.settings.cardBgOpacity : 0.95)
                    .setDynamicTooltip()
                    .onChange(async (v) => {
                        this.plugin.settings.cardBgOpacity = v;
                        await this.plugin.saveSettings();
                        this.plugin.refreshView();
                    });
            });

        new Setting(containerEl)
            .setName('重置布局')
            .setDesc('清除所有模块的位置和尺寸设置，恢复默认布局')
            .addButton(b => {
                b.setButtonText('重置').setWarning()
                    .onClick(async () => {
                        this.plugin.settings.layout = {};
                        await this.plugin.saveSettings();
                        this.plugin.refreshView();
                        new Notice('布局已重置');
                    });
            });
    }

    _renderModuleToggles(containerEl) {
        containerEl.createEl('h3', { text: '模块管理' });

        const loadedModules = this.plugin.moduleManager.getAllModules();
        if (loadedModules.length === 0) {
            containerEl.createEl('p', {
                text: '未找到任何模块文件，请检查 modules/ 目录',
                attr: { style: 'color: var(--text-muted); font-size: 13px;' }
            });
            return;
        }

        loadedModules.forEach(mod => {
            const moduleId = mod.id;
            if (!moduleId) return;
            const modSettings = this.plugin.settings.modules[moduleId] || {};
            new Setting(containerEl)
                .setName((mod.icon || '📦') + ' ' + (mod.title || moduleId))
                .setDesc(mod.id)
                .addToggle(t => {
                    t.setValue(modSettings.enabled !== false)
                        .onChange(async (v) => {
                            if (!this.plugin.settings.modules[moduleId]) {
                                this.plugin.settings.modules[moduleId] = {};
                            }
                            this.plugin.settings.modules[moduleId].enabled = v;
                            await this.plugin.saveSettings();
                            this.plugin.refreshView();
                        });
                })
                .addButton(b => {
                    b.setButtonText('配置')
                        .onClick(() => {
                            this._currentModuleId = this._currentModuleId === moduleId ? null : moduleId;
                            this.display();
                        });
                });
        });
    }

    _renderInstanceManager(containerEl) {
        const instances = this.plugin.settings.instances || [];
        if (instances.length === 0) return;

        containerEl.createEl('h3', { text: '实例管理' });

        containerEl.createEl('p', {
            text: '以下是通过 ➕ 按钮添加的额外板块实例，可在此管理',
            attr: { style: 'color: var(--text-muted); font-size: 12px; margin-bottom: 12px;' }
        });

        instances.forEach(inst => {
            const baseMod = this.plugin.moduleManager.getModule(inst.baseModule);
            const modSettings = this.plugin.settings.modules[inst.id] || {};

            new Setting(containerEl)
                .setName((baseMod ? baseMod.icon : '📦') + ' ' + inst.label)
                .setDesc('类型: ' + (baseMod ? baseMod.title : inst.baseModule) + ' (ID: ' + inst.id + ')')
                .addToggle(t => {
                    t.setValue(modSettings.enabled !== false)
                        .onChange(async (v) => {
                            if (!this.plugin.settings.modules[inst.id]) {
                                this.plugin.settings.modules[inst.id] = { enabled: true };
                            }
                            this.plugin.settings.modules[inst.id].enabled = v;
                            await this.plugin.saveSettings();
                            this.plugin.refreshView();
                        });
                })
                .addButton(b => {
                    b.setButtonText('配置')
                        .onClick(() => {
                            this._currentModuleId = this._currentModuleId === inst.id ? null : inst.id;
                            this.display();
                        });
                })
                .addButton(b => {
                    b.setButtonText('删除').setWarning()
                        .onClick(async () => {
                            instances.splice(instances.indexOf(inst), 1);
                            this.plugin.settings.instances = instances;
                            delete this.plugin.settings.modules[inst.id];
                            delete this.plugin.settings.layout[inst.id];
                            if (this.plugin.settings.moduleOrder) {
                                this.plugin.settings.moduleOrder = this.plugin.settings.moduleOrder.filter(id => id !== inst.id);
                            }
                            if (this._currentModuleId === inst.id) this._currentModuleId = null;
                            await this.plugin.saveSettings();
                            this.plugin.refreshView();
                            new Notice('已删除: ' + inst.label);
                            this.display();
                        });
                });
        });
    }

    _renderModuleSettings(containerEl, moduleId) {
        const mod = this.plugin.moduleManager.resolveModule(moduleId);
        if (!mod || !mod.renderSettings) return;

        const ctx = this.plugin.moduleManager.createContext(moduleId);
        Object.assign(this.plugin.moduleManager._runtimeCtx, ctx);

        const instanceInfo = this.plugin.moduleManager.getInstanceInfo(moduleId);
        const displayTitle = instanceInfo ? instanceInfo.label : (mod.title || moduleId);

        const wrapper = containerEl.createDiv({ cls: 'v15-module-settings-wrapper' });
        wrapper.createEl('h3', { text: '⚙️ ' + displayTitle + ' 设置' });

        const saveCallback = async () => {
            await this.plugin.saveSettings();
        };

        try {
            mod.renderSettings(wrapper, this.plugin, saveCallback);
        } catch (e) {
            console.error('[V15] 模块 ' + moduleId + ' 设置渲染失败:', e);
            wrapper.createEl('p', { text: '设置加载失败: ' + e.message, attr: { style: 'color: var(--text-muted);' } });
        }
    }

    _renderDonateSection(containerEl) {
        const section = containerEl.createDiv({
            attr: {
                style: 'margin-top:32px;padding-top:20px;border-top:2px dashed var(--background-modifier-border);'
            }
        });

        section.createEl('h3', {
            text: '☕ 支持开发者',
            attr: { style: 'text-align:center;margin-bottom:12px;' }
        });

        // 加载二维码图片
        let qrSrc = '';
        try {
            const nodePath = require('path');
            const fs = require('fs');
            const adapter = this.app.vault.adapter;
            const vaultBase = adapter.basePath || adapter.getBasePath?.() || '';
            const relDir = (this.plugin.manifest.dir || '').replace(/\\/g, '/').replace(/^\//, '');

            const candidates = [
                nodePath.join(vaultBase, relDir, 'assets', 'donate-qrcode.png'),
                nodePath.join(vaultBase, '.obsidian', 'plugins', 'ai-smart-dashboard-v15', 'assets', 'donate-qrcode.png'),
            ];
            if (nodePath.isAbsolute(relDir || '')) {
                candidates.unshift(nodePath.join(relDir, 'assets', 'donate-qrcode.png'));
            }

            let foundPath = '';
            for (const p of candidates) {
                if (p && fs.existsSync(p)) { foundPath = p; break; }
            }
            console.log('[V15 Donate] 找到路径:', foundPath || '未找到', '候选:', candidates);

            if (foundPath) {
                const buf = fs.readFileSync(foundPath);
                qrSrc = 'data:image/png;base64,' + buf.toString('base64');
            }
        } catch (e) {
            console.error('[V15] 加载打赏二维码失败:', e);
        }

        const qrWrap = section.createDiv({
            attr: { style: 'text-align:center;' }
        });

        if (qrSrc) {
            qrWrap.createEl('img', {
                attr: {
                    src: qrSrc,
                    style: 'width:280px;height:280px;object-fit:contain;border-radius:10px;border:2px solid var(--background-modifier-border);background:#fff;display:block;margin:0 auto;'
                }
            });
        } else {
            qrWrap.createEl('div', {
                text: '二维码加载失败，请检查 assets/donate-qrcode.png 是否存在',
                attr: { style: 'color:var(--text-muted);font-size:12px;text-align:center;' }
            });
        }
    }
}

// ===================== 主插件类 =====================
class DashboardPlugin extends Plugin {
    async onload() {
        await this.loadSettings();
        this.moduleManager = new ModuleManager(this);

        await this.initModuleLayouts();

        this.registerView(VIEW_TYPE, (leaf) => new DashboardView(leaf, this));

        this.addRibbonIcon('layout-dashboard', 'Modular Theme Dashboard', () => this.activateView());

        this.addCommand({
            id: 'open-dashboard',
            name: 'Open Modular Theme Dashboard',
            callback: () => this.activateView()
        });

        this.addSettingTab(new DashboardSettingTab(this.app, this));

        this.app.workspace.onLayoutReady(() => {
            this.activateView();
        });
    }

    async initModuleLayouts() {
        if (this.moduleManager.modules.size === 0) {
            await this.moduleManager.loadModules();
        }
        const loadedIds = this.moduleManager.getLoadedModuleIds();
        const instances = this.settings.instances || [];
        const instanceIds = instances.map(i => i.id);
        const allIds = new Set([...loadedIds, ...instanceIds]);

        let changed = false;

        // 确保每个基础模块有 settings 条目（含完整默认值）
        loadedIds.forEach(moduleId => {
            if (!this.settings.modules[moduleId]) {
                const defaults = this.moduleManager.getModuleDefaultSettings(moduleId);
                this.settings.modules[moduleId] = {
                    enabled: true,
                    ...defaults
                };
                changed = true;
            }
        });

        // 确保每个实例有 settings 条目（含完整默认值）
        instances.forEach(inst => {
            if (!this.settings.modules[inst.id]) {
                const defaults = this.moduleManager.getModuleDefaultSettings(inst.baseModule);
                this.settings.modules[inst.id] = {
                    enabled: true,
                    ...JSON.parse(JSON.stringify(defaults))
                };
                changed = true;
            }
        });

        // 分配默认布局（基础模块按 index 排列）
        loadedIds.forEach((moduleId, index) => {
            if (!this.settings.layout[moduleId]) {
                const col = index % 3;
                const row = Math.floor(index / 3);
                this.settings.layout[moduleId] = {
                    x: 20 + col * 320,
                    y: 20 + row * 270,
                    width: 300,
                    height: 250
                };
                changed = true;
            }
        });

        // 实例布局：放在基础模块下方
        instances.forEach((inst, index) => {
            if (!this.settings.layout[inst.id]) {
                this.settings.layout[inst.id] = {
                    x: 20 + (index % 3) * 320,
                    y: 20 + (loadedIds.length + index) * 270,
                    width: 300,
                    height: 250
                };
                changed = true;
            }
        });

        // 清理已删除的布局和设置（保留实例的）
        Object.keys(this.settings.layout).forEach(moduleId => {
            if (!allIds.has(moduleId)) {
                delete this.settings.layout[moduleId];
                changed = true;
            }
        });
        Object.keys(this.settings.modules).forEach(moduleId => {
            if (!allIds.has(moduleId)) {
                delete this.settings.modules[moduleId];
                changed = true;
            }
        });

        // 确保 instances 数组存在
        if (!this.settings.instances) {
            this.settings.instances = [];
            changed = true;
        }
        if (this.settings.instanceCounter == null) {
            this.settings.instanceCounter = 0;
            changed = true;
        }

        if (changed) await this.saveSettings();
    }

    onunload() {
        this.app.workspace.detachLeavesOfType(VIEW_TYPE);
    }

    async activateView() {
        const existing = this.app.workspace.getLeavesOfType(VIEW_TYPE);
        if (existing.length > 0) {
            this.app.workspace.revealLeaf(existing[0]);
            return;
        }
        const leaf = this.app.workspace.getLeaf(true);
        await leaf.setViewState({ type: VIEW_TYPE, active: true });
        this.app.workspace.revealLeaf(leaf);
    }

    refreshView() {
        const leaves = this.app.workspace.getLeavesOfType(VIEW_TYPE);
        leaves.forEach(leaf => {
            if (leaf.view instanceof DashboardView) {
                leaf.view.render();
            }
        });
    }

    async loadSettings() {
        const saved = await this.loadData();
        this.settings = Object.assign({}, DEFAULT_SETTINGS, saved);

        // 深拷贝 modules，确保默认值存在
        this.settings.modules = Object.assign({}, DEFAULT_SETTINGS.modules);
        if (saved && saved.modules) {
            // 合并基础模块设置
            Object.keys(DEFAULT_SETTINGS.modules).forEach(mid => {
                this.settings.modules[mid] = Object.assign(
                    {},
                    DEFAULT_SETTINGS.modules[mid],
                    saved.modules[mid] || {}
                );
            });

            // 合并实例设置（确保包含基础模块的默认值）
            if (saved.instances) {
                saved.instances.forEach(inst => {
                    const baseDefaults = DEFAULT_SETTINGS.modules[inst.baseModule] || {};
                    const savedInstSettings = saved.modules[inst.id] || {};
                    this.settings.modules[inst.id] = Object.assign(
                        {},
                        baseDefaults,
                        savedInstSettings
                    );
                });
            }

            // 合并任何其他已保存的模块设置（兼容旧数据）
            Object.keys(saved.modules).forEach(mid => {
                if (!this.settings.modules[mid]) {
                    this.settings.modules[mid] = saved.modules[mid];
                }
            });
        }

        if (!this.settings.layout) this.settings.layout = {};
        if (!this.settings.moduleOrder) {
            this.settings.moduleOrder = DEFAULT_SETTINGS.moduleOrder;
        }
        if (!this.settings.instances) this.settings.instances = [];
        if (this.settings.instanceCounter == null) this.settings.instanceCounter = 0;
    }

    async saveSettings() {
        await this.saveData(this.settings);
    }
}

module.exports = DashboardPlugin;
