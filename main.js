/**
 * AI Dashboard V17
 * 底层：V14 自由拖拽/resize 布局 + 完全模块化架构
 * 功能：V11 完整功能迁移 + 无限实例化系统（所有模块默认可克隆）
 * 主题：V11 8个精美主题
 * 构建版本：17.1.6 (release)
 */

const { Plugin, ItemView, FileView, Setting, PluginSettingTab, Modal, Notice, setIcon, requestUrl, moment } = require('obsidian');

const VIEW_TYPE = 'ai-dashboard-v15';
// ★ 文件查看器：接管所有非 md 文件扩展名，在 Obsidian 页签内直接预览
// 参考原独立插件（ViewItAll、Univer、html-plugin），registerExtensions 注册扩展名，
// 当 Obsidian 尝试打开非 md 文件时路由到此视图，在页签内渲染内容，
// 避免 shell.openPath() 调用外部程序（WPS 等）。
const FILE_VIEWER_TYPE = 'dashboard-file-viewer';

// ============ 文件查看器 库加载（SheetJS + mammoth.js + docstream）============
// ★ SheetJS：仍用隐藏 iframe 沙箱（XLSX.read 同步，无跨上下文问题）
// ★ mammoth.js：改用 eval+IIFE 包装 → 屏蔽 module/exports/define，强制走 browser 挂载
//   → 避免 iframe 沙箱中 JSZip 异步处理 ArrayBuffer 跨上下文失败问题
// ★ docstream（@jose.espana/docstream）：用于解析 .doc 旧版 Word 二进制格式
//   → eval+IIFE 加载（同 mammoth），约 1.2MB，支持 DOC/XLS/PPT 遗留格式
// ★ 所有库：本地 libs/ 优先，CDN 回退（离线环境完全可用）
var _xlsxLib = null, _xlsxLoaded = false, _xlsxLoading = false, _xlsxWaiters = [];
var _mammothLib = null, _mammothLoaded = false, _mammothLoading = false, _mammothWaiters = [];
var _docstreamLib = null, _docstreamLoaded = false, _docstreamLoading = false, _docstreamWaiters = [];
var _cfbLib = null, _cfbLoaded = false, _cfbLoading = false, _cfbWaiters = [];
var _jszipLib = null, _jszipLoaded = false, _jszipLoading = false, _jszipWaiters = [];
var _sandboxFrame = null;  // 共享的隐藏 iframe（仅 SheetJS 使用）

// ★ 本地库读取器：从插件 libs/ 目录读取库文件（100%离线可用）
// 返回 Promise<string>（库源码），失败时 reject
function _readLocalLib(filename) {
    var pluginId = window.__PLUGIN_ID__ || '';
    if (!pluginId) {
        return Promise.reject(new Error('[DFV] __PLUGIN_ID__ 未设置，无法读取本地库'));
    }
    if (!app || !app.vault || !app.vault.adapter) {
        return Promise.reject(new Error('[DFV] app.vault.adapter 不可用'));
    }
    var vaultPath = '.obsidian/plugins/' + pluginId + '/libs/' + filename;
    return app.vault.adapter.read(vaultPath).then(function(code) {
        if (!code || code.length < 100) {
            throw new Error('[DFV] 本地库文件为空或过短: ' + filename);
        }
        console.log('[DFV] ✓ 本地库加载成功: libs/' + filename + ' (' + (code.length / 1024).toFixed(0) + 'KB)');
        return code;
    });
}

function getSandboxFrame() {
    if (_sandboxFrame && _sandboxFrame.contentWindow) return _sandboxFrame;
    _sandboxFrame = document.createElement('iframe');
    _sandboxFrame.style.cssText = 'display:none;position:absolute;width:0;height:0;border:0;';
    document.body.appendChild(_sandboxFrame);
    return _sandboxFrame;
}

function loadLibInSandbox(libCode) {
    // 在隐藏 iframe 的纯浏览器沙箱中执行库代码
    var iframe = getSandboxFrame();
    var win = iframe.contentWindow;
    // 在 iframe 上下文中执行库代码（纯 eval，不创建 DOM 节点）
    try { win.eval(libCode); } catch(e) {
        console.warn('[Dashboard] Sandbox eval 失败:', e.message);
    }
    return win;
}

function loadXLSXOnce() {
    if (_xlsxLoaded || _xlsxLoading) return;
    _xlsxLoading = true;
    // ★ 策略0：直接检查 window.XLSX（可能是其他插件或浏览器缓存已加载）
    if (window.XLSX && typeof window.XLSX.read === 'function') {
        _xlsxLib = window.XLSX;
        _xlsxLoaded = true; _xlsxLoading = false;
        console.log('[DFV] SheetJS 使用主窗口已有实例');
        _xlsxWaiters.forEach(function(w) { w(_xlsxLib); }); _xlsxWaiters = [];
        return;
    }
    // eval 执行器（本地和 CDN 共用）
    function evalXLSX(code) {
        var fn = new Function('var module=undefined,exports=undefined,define=undefined;' + code + ';return typeof XLSX!=="undefined"?XLSX:null;');
        var lib = fn();
        if (lib && typeof lib === 'object' && typeof lib.read === 'function') {
            return lib;
        }
        return null;
    }
    function onSuccess(lib) {
        _xlsxLib = lib;
        _xlsxLoaded = true; _xlsxLoading = false;
        _xlsxWaiters.forEach(function(w) { w(_xlsxLib); }); _xlsxWaiters = [];
    }
    function onFail() {
        _xlsxLib = null; _xlsxLoaded = true; _xlsxLoading = false;
        console.error('[DFV] SheetJS 所有加载方式均失败（本地+CDN）');
        _xlsxWaiters.forEach(function(w) { w(null); }); _xlsxWaiters = [];
    }
    // CDN 容灾（仅作为回退）
    function tryFetchFromCDN(urls, idx) {
        if (idx >= urls.length) { onFail(); return; }
        requestUrl({ url: urls[idx] })
            .then(function(resp) {
                var lib = evalXLSX(resp.text);
                if (lib) {
                    console.log('[DFV] SheetJS 加载成功 (CDN #' + (idx+1) + '):', urls[idx]);
                    onSuccess(lib);
                } else {
                    tryFetchFromCDN(urls, idx + 1);
                }
            })
            .catch(function(err) {
                console.warn('[DFV] SheetJS CDN #' + (idx+1) + ' 失败:', String(err));
                tryFetchFromCDN(urls, idx + 1);
            });
    }
    // ★ 策略1：本地 libs/ 优先（离线可用）
    _readLocalLib('xlsx.full.min.js').then(function(code) {
        var lib = evalXLSX(code);
        if (lib) {
            console.log('[DFV] SheetJS ✓ 从本地加载成功');
            onSuccess(lib);
        } else {
            console.warn('[DFV] SheetJS 本地文件 eval 失败，回退CDN');
            tryFetchFromCDN([
                'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
                'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
                'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
            ], 0);
        }
    }).catch(function(err) {
        console.warn('[DFV] SheetJS 本地加载失败:', String(err), '→ 回退CDN');
        tryFetchFromCDN([
            'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
            'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
            'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
        ], 0);
    });
}

function getXLSXAsync() {
    if (_xlsxLoaded) return Promise.resolve(_xlsxLib);
    return new Promise(function(resolve) {
        _xlsxWaiters.push(resolve);
        loadXLSXOnce();
    });
}

function loadMammothOnce() {
    if (_mammothLoaded || _mammothLoading) return;
    _mammothLoading = true;
    function evalMammoth(code) {
        var wrapped = '(function(){var module=undefined,exports=undefined,define=undefined;' + code + '\nreturn mammoth;})()';
        return eval(wrapped);
    }
    function onSuccess(lib) {
        _mammothLib = lib;
        _mammothLoaded = true; _mammothLoading = false;
        console.log('[DFV] mammoth 加载完成');
        _mammothWaiters.forEach(function(w) { w(_mammothLib); }); _mammothWaiters = [];
    }
    function onFail(label) {
        _mammothLib = null; _mammothLoaded = true; _mammothLoading = false;
        console.error('[DFV] mammoth 加载失败 (' + label + ')');
        _mammothWaiters.forEach(function(w) { w(null); }); _mammothWaiters = [];
    }
    // ★ 本地优先
    _readLocalLib('mammoth.browser.min.js').then(function(code) {
        try {
            var lib = evalMammoth(code);
            if (lib && typeof lib.convertToHtml === 'function') {
                console.log('[DFV] mammoth ✓ 从本地加载成功');
                onSuccess(lib);
            } else {
                console.warn('[DFV] mammoth 本地 eval 返回无效对象，回退CDN');
                cdnFallback();
            }
        } catch(e) {
            console.warn('[DFV] mammoth 本地 eval 异常:', String(e), '→ 回退CDN');
            cdnFallback();
        }
    }).catch(function(err) {
        console.warn('[DFV] mammoth 本地加载失败:', String(err), '→ 回退CDN');
        cdnFallback();
    });
    function cdnFallback() {
        requestUrl({ url: 'https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js' })
            .then(function(resp) {
                try {
                    var lib = evalMammoth(resp.text);
                    if (lib && typeof lib.convertToHtml === 'function') {
                        console.log('[DFV] mammoth 从 CDN 加载成功');
                        onSuccess(lib);
                    } else {
                        onFail('CDN eval 返回无效对象');
                    }
                } catch(e) { onFail('CDN eval 异常: ' + String(e)); }
            })
            .catch(function(err) { onFail('CDN 网络错误: ' + String(err)); });
    }
}

function getMammothAsync() {
    if (_mammothLoaded) return Promise.resolve(_mammothLib);
    return new Promise(function(resolve) {
        _mammothWaiters.push(resolve);
        loadMammothOnce();
    });
}

// ============ docstream 加载（.doc 解析） ============
function loadDocstreamOnce() {
    if (_docstreamLoaded || _docstreamLoading) return;
    _docstreamLoading = true;
    function evalDocstream(code) {
        var clean = code.replace(/^#![^\n]*\n/, '');
        var wrapped = '(function(){var module=undefined,exports=undefined,define=undefined;' + clean + '\nreturn officeParser;})()';
        return eval(wrapped);
    }
    function onSuccess(lib) {
        _docstreamLib = lib;
        _docstreamLoaded = true; _docstreamLoading = false;
        console.log('[DFV] docstream 加载完成');
        _docstreamWaiters.forEach(function(w) { w(_docstreamLib); }); _docstreamWaiters = [];
    }
    function onFail(label) {
        _docstreamLib = null; _docstreamLoaded = true; _docstreamLoading = false;
        console.error('[DFV] docstream 加载失败 (' + label + ')');
        _docstreamWaiters.forEach(function(w) { w(null); }); _docstreamWaiters = [];
    }
    // ★ 本地优先
    _readLocalLib('docstream.browser@0.1.3.js').then(function(code) {
        try {
            var lib = evalDocstream(code);
            if (lib && typeof lib.parseOffice === 'function') {
                console.log('[DFV] docstream ✓ 从本地加载成功');
                onSuccess(lib);
            } else {
                console.warn('[DFV] docstream 本地 eval 返回无效对象，回退CDN');
                cdnFallback();
            }
        } catch(e) {
            console.warn('[DFV] docstream 本地 eval 异常:', String(e), '→ 回退CDN');
            cdnFallback();
        }
    }).catch(function(err) {
        console.warn('[DFV] docstream 本地加载失败:', String(err), '→ 回退CDN');
        cdnFallback();
    });
    function cdnFallback() {
        requestUrl({ url: 'https://cdn.jsdelivr.net/npm/@jose.espana/docstream@0.1.3/dist/docstream.browser@0.1.3.js' })
            .then(function(resp) {
                try {
                    var lib = evalDocstream(resp.text);
                    if (lib && typeof lib.parseOffice === 'function') {
                        console.log('[DFV] docstream 从 CDN 加载成功');
                        onSuccess(lib);
                    } else {
                        onFail('CDN eval 返回无效对象');
                    }
                } catch(e) { onFail('CDN eval 异常: ' + String(e)); }
            })
            .catch(function(err) { onFail('CDN 网络错误: ' + String(err)); });
    }
}

function getDocstreamAsync() {
    if (_docstreamLoaded) return Promise.resolve(_docstreamLib);
    return new Promise(function(resolve) {
        _docstreamWaiters.push(resolve);
        loadDocstreamOnce();
    });
}

// ============ CFB 加载（OLE2 复合文档解析 — .doc 核心依赖）============
// CFB（Compound File Binary）由 SheetJS 社区维护，稳定性远超 docstream
// 用于解析 Office 97-2003 遗留格式的 OLE2 容器结构
function loadCfbOnce() {
    if (_cfbLoaded || _cfbLoading) return;
    _cfbLoading = true;
    function evalCFB(code) {
        var wrapped = '(function(){var module=undefined,exports=undefined,define=undefined;' + code + '\nreturn CFB;})()';
        return eval(wrapped);
    }
    function onSuccess(lib) {
        _cfbLib = lib;
        _cfbLoaded = true; _cfbLoading = false;
        console.log('[DFV] CFB 加载完成');
        _cfbWaiters.forEach(function(w) { w(_cfbLib); }); _cfbWaiters = [];
    }
    function onFail(label) {
        _cfbLib = null; _cfbLoaded = true; _cfbLoading = false;
        console.error('[DFV] CFB 加载失败 (' + label + ')');
        _cfbWaiters.forEach(function(w) { w(null); }); _cfbWaiters = [];
    }
    // ★ 本地优先
    _readLocalLib('cfb.min.js').then(function(code) {
        try {
            var lib = evalCFB(code);
            if (lib && typeof lib.read === 'function') {
                console.log('[DFV] CFB ✓ 从本地加载成功');
                onSuccess(lib);
            } else { console.warn('[DFV] CFB 本地 eval 返回无效对象，回退CDN'); cdnFallback(); }
        } catch(e) { console.warn('[DFV] CFB 本地 eval 异常:', String(e), '→ 回退CDN'); cdnFallback(); }
    }).catch(function(err) { console.warn('[DFV] CFB 本地加载失败:', String(err), '→ 回退CDN'); cdnFallback(); });
    function cdnFallback() {
        requestUrl({ url: 'https://cdn.jsdelivr.net/npm/cfb@1.2.2/dist/cfb.min.js' })
            .then(function(resp) {
                try {
                    var lib = evalCFB(resp.text);
                    if (lib && typeof lib.read === 'function') {
                        console.log('[DFV] CFB 从 CDN 加载成功');
                        onSuccess(lib);
                    } else { onFail('CDN eval 返回无效对象'); }
                } catch(e) { onFail('CDN eval 异常: ' + String(e)); }
            })
            .catch(function(err) { onFail('CDN 网络错误: ' + String(err)); });
    }
}

function getCfbAsync() {
    if (_cfbLoaded) return Promise.resolve(_cfbLib);
    return new Promise(function(resolve) {
        _cfbWaiters.push(resolve);
        loadCfbOnce();
    });
}

// ============ JSZip 加载（DOCX 兜底解压器 — mammoth 失败后手动提取文本）============
function loadJSZipOnce() {
    if (_jszipLoaded || _jszipLoading) return;
    _jszipLoading = true;
    function evalJSZip(code) {
        var wrapped = '(function(){var module=undefined,exports=undefined,define=undefined;' + code + '\nreturn JSZip;})()';
        return eval(wrapped);
    }
    function onSuccess(lib) {
        _jszipLib = lib;
        _jszipLoaded = true; _jszipLoading = false;
        console.log('[DFV] JSZip 加载完成');
        _jszipWaiters.forEach(function(w) { w(_jszipLib); }); _jszipWaiters = [];
    }
    function onFail(label) {
        _jszipLib = null; _jszipLoaded = true; _jszipLoading = false;
        console.error('[DFV] JSZip 加载失败 (' + label + ')');
        _jszipWaiters.forEach(function(w) { w(null); }); _jszipWaiters = [];
    }
    // ★ 本地优先
    _readLocalLib('jszip.min.js').then(function(code) {
        try {
            var lib = evalJSZip(code);
            if (lib) {
                console.log('[DFV] JSZip ✓ 从本地加载成功');
                onSuccess(lib);
            } else { console.warn('[DFV] JSZip 本地 eval 返回无效对象，回退CDN'); cdnFallback(); }
        } catch(e) { console.warn('[DFV] JSZip 本地 eval 异常:', String(e), '→ 回退CDN'); cdnFallback(); }
    }).catch(function(err) { console.warn('[DFV] JSZip 本地加载失败:', String(err), '→ 回退CDN'); cdnFallback(); });
    function cdnFallback() {
        requestUrl({ url: 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js' })
            .then(function(resp) {
                try {
                    var lib = evalJSZip(resp.text);
                    if (lib) { console.log('[DFV] JSZip 从 CDN 加载成功'); onSuccess(lib); }
                    else { onFail('CDN eval 返回无效对象'); }
                } catch(e) { onFail('CDN eval 异常: ' + String(e)); }
            })
            .catch(function(err) { onFail('CDN 网络错误: ' + String(err)); });
    }
}

function getJSZipAsync() {
    if (_jszipLoaded) return Promise.resolve(_jszipLib);
    return new Promise(function(resolve) {
        _jszipWaiters.push(resolve);
        loadJSZipOnce();
    });
}

function startViewerLibPreload() {
    loadXLSXOnce();
    loadMammothOnce();
    loadDocstreamOnce();
    loadCfbOnce();
    loadJSZipOnce();
}

// ============ 文件查看器类（对齐原插件：FileView 继承）============
// ★ 根因修复：改为继承 FileView（非 ItemView），
// Obsidian 框架在通过 registerExtensions 打开文件时会自动设置 this.file
// 并调用 onLoadFile(file) 回调。对齐 ViewItAll、Univer、html-plugin。
// ★ 文件类型处理已拆分为 modules/file-viewers/*.js 独立扩展

// ============ 全局插件引用（供 canAcceptExtension 访问设置） ============
var __DBFV_PLUGIN__ = null;

// ============ FileViewer 扩展注册表 ============
var FILE_VIEWER_HANDLERS = {};

// ============ 扩展名 → 扩展组名 映射表 ============
// 用于设置开关：按组（xlsx/doc/docx/html/image/video/office/text）控制
var FILE_VIEWER_EXT_GROUPS = {
    "xlsx": "xlsx", "xls": "xlsx", "xlsm": "xlsx", "csv": "xlsx", "ods": "xlsx",
    "docx": "docx",
    "doc": "doc",
    "html": "html", "htm": "html",
    "png": "image", "jpg": "image", "jpeg": "image", "gif": "image",
    "svg": "image", "webp": "image", "bmp": "image", "ico": "image",
    "mp4": "video", "webm": "video", "mov": "video", "avi": "video",
    "mkv": "video", "m4v": "video",
    "mp3": "video", "wav": "video", "ogg": "video", "flac": "video", "m4a": "video", "aac": "video",
    "pdf": "video",
    "ppt": "office", "pptx": "office",
    "txt": "text", "json": "text", "yaml": "text", "yml": "text", "xml": "text",
    "toml": "text", "ini": "text", "cfg": "text", "env": "text",
    "js": "text", "ts": "text", "jsx": "text", "tsx": "text",
    "css": "text", "scss": "text", "less": "text",
    "py": "text", "rb": "text", "java": "text", "go": "text", "rs": "text",
    "c": "text", "cpp": "text", "cs": "text", "sh": "text", "bat": "text", "ps1": "text"
};

class DashboardFileViewer extends FileView {
    constructor(leaf, plugin) {
        super(leaf);
        this.plugin = plugin;
    }
    getViewType() { return FILE_VIEWER_TYPE; }
    getDisplayText() { return this.file ? '📄 ' + this.file.name : '📄 文件查看器'; }
    getIcon() { return 'file'; }

    canAcceptExtension(ext) {
        // ★ 只接受非 Obsidian 原生扩展（让 md/canvas 等走原生视图）
        //    否则从表格/Word 切回 MD 时会被拦截为原始文本
        var e = ext.toLowerCase();
        var nativeExts = ['md', 'canvas'];
        if (nativeExts.indexOf(e) !== -1) return false;

        // ★ 检查 FileViewer 设置开关：禁用时返回 false，让 Obsidian 用系统默认程序打开
        var fvSettings = null;
        if (this.plugin && this.plugin.settings) {
            fvSettings = this.plugin.settings.fileViewerExtensions;
        } else if (__DBFV_PLUGIN__) {
            fvSettings = __DBFV_PLUGIN__.settings.fileViewerExtensions;
        }
        if (fvSettings) {
            var extGroup = FILE_VIEWER_EXT_GROUPS[e];
            if (extGroup && fvSettings[extGroup] === false) {
                return false;  // 释放扩展名，让 Obsidian 回退到系统默认应用
            }
        }
        return true;
    }

    // ★ 核心：FileView 的 onLoadFile 回调，Obsidian 框架传入 TFile 对象
    async onLoadFile(file) {
        var container = this.contentEl;
        container.empty();
        container.addClass('dbfv-container');
        container.style.cssText = 'display:flex;flex-direction:column;height:100%;overflow:hidden;user-select:text;-webkit-user-select:text;';

        var fileName = file.name;
        var ext = file.extension.toLowerCase();
        var filePath = file.path;

        console.log('[DashboardFileViewer] 打开文件:', fileName, '扩展名:', ext);

        // 内容区（直接使用全部空间，不添加自定义 header）
        var contentArea = container.createEl('div', { cls: 'dbfv-content',
            attr: { style: 'flex:1;overflow:auto;min-height:0;position:relative;user-select:text;-webkit-user-select:text;' }
        });

        // ★ 轻量浮动工具栏（右上角，避让内容）
        var floatBar = contentArea.createEl('div', { cls: 'dbfv-floatbar',
            attr: { style: 'position:absolute;top:6px;right:6px;z-index:10;display:flex;gap:6px;align-items:center;' }
        });
        var extBadge = floatBar.createEl('span', { cls: 'dbfv-ext-badge', text: '.' + ext.toUpperCase(),
            attr: { style: 'font-size:10px;color:var(--text-faint);background:var(--background-modifier-form-field);padding:1px 6px;border-radius:8px;opacity:0.7;' }
        });
        var openBtn = floatBar.createEl('button', { text: '🔗', attr: {
            style: 'padding:2px 6px;font-size:11px;border:1px solid var(--background-modifier-border);border-radius:4px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;opacity:0.6;',
            title: '用外部程序打开 ' + fileName
        }});
        var vault = this.app.vault;
        var adapter = vault.adapter;

        openBtn.addEventListener('click', function() {
            try {
                var fullPath = adapter.getFullPath(filePath);
                require('electron').shell.openPath(fullPath);
            } catch(e) { new Notice('无法打开: ' + e.message); }
        });

        // hover 时提高可见度
        floatBar.addEventListener('mouseenter', function() {
            extBadge.style.opacity = '1';
            openBtn.style.opacity = '1';
            if (zoomGroup) zoomGroup.style.opacity = '1';
        });
        floatBar.addEventListener('mouseleave', function() {
            extBadge.style.opacity = '0.7';
            openBtn.style.opacity = '0.6';
            if (zoomGroup) zoomGroup.style.opacity = '0.6';
        });

        // ======== 缩放控件（HTML + 表格） ========
        var zoomLevel = 1.0;
        var zoomLabel = null;
        var zoomTargetEl = null;  // 被缩放的 DOM 元素
        var zoomGroup = null;

        function updateZoom(newLevel) {
            zoomLevel = Math.max(0.3, Math.min(3.0, parseFloat(newLevel.toFixed(2))));
            if (zoomLabel) zoomLabel.textContent = Math.round(zoomLevel * 100) + '%';
            if (zoomTargetEl) {
                // 使用 zoom 属性（Chromium/Electron 原生支持，自动处理滚动区域）
                // ★ 只有 zoom !== 1 时才设置，避免 zoom:1 干扰文本选择
                if (zoomLevel === 1.0) {
                    zoomTargetEl.style.zoom = '';
                } else {
                    zoomTargetEl.style.zoom = zoomLevel;
                }
            }
        }

    var htmlExts = ['html', 'htm'];  // 前向声明，下面分类处理中也会定义
    var spreadsheetExts2 = ['xlsx', 'xls', 'xlsm', 'csv', 'ods'];
    var docxExts2 = ['docx'];
    var officeExts2 = ['ppt', 'pptx'];
    var isZoomable = (htmlExts.indexOf(ext) !== -1 || spreadsheetExts2.indexOf(ext) !== -1 || docxExts2.indexOf(ext) !== -1 || officeExts2.indexOf(ext) !== -1);
        if (isZoomable) {
            zoomGroup = floatBar.createEl('div', { cls: 'dbfv-zoom-group',
                attr: { style: 'display:flex;align-items:center;gap:1px;opacity:0.6;' }
            });
            var zoomOutBtn = zoomGroup.createEl('button', { text: '−', attr: {
                style: 'padding:2px 6px;font-size:12px;border:1px solid var(--background-modifier-border);border-radius:4px 0 0 4px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;line-height:1;',
                title: '缩小'
            }});
            zoomLabel = zoomGroup.createEl('span', { text: '100%', attr: {
                style: 'font-size:10px;padding:1px 5px;background:var(--background-modifier-form-field);border-top:1px solid var(--background-modifier-border);border-bottom:1px solid var(--background-modifier-border);color:var(--text-muted);user-select:none;line-height:1.6;'
            }});
            var zoomInBtn = zoomGroup.createEl('button', { text: '+', attr: {
                style: 'padding:2px 6px;font-size:12px;border:1px solid var(--background-modifier-border);border-radius:0 4px 4px 0;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;line-height:1;',
                title: '放大'
            }});
            // 重置
            var zoomResetBtn = zoomGroup.createEl('button', { text: '↺', attr: {
                style: 'padding:2px 5px;font-size:11px;border:1px solid var(--background-modifier-border);border-radius:4px;background:var(--background-secondary);color:var(--text-muted);cursor:pointer;line-height:1;margin-left:3px;',
                title: '重置缩放'
            }});

            zoomOutBtn.addEventListener('click', function(e) { e.stopPropagation(); updateZoom(zoomLevel - 0.1); });
            zoomInBtn.addEventListener('click', function(e) { e.stopPropagation(); updateZoom(zoomLevel + 0.1); });
            zoomResetBtn.addEventListener('click', function(e) { e.stopPropagation(); updateZoom(1.0); });
            // 滚轮缩放（Ctrl+滚轮 或 pinch）
            contentArea.addEventListener('wheel', function(e) {
                if (e.ctrlKey || e.metaKey) {
                    e.preventDefault();
                    updateZoom(zoomLevel + (e.deltaY < 0 ? 0.05 : -0.05));
                }
            }, { passive: false });
        }

        try {
            var zoomCtx = isZoomable ? {
                getLevel: function() { return zoomLevel; },
                setTarget: function(el) { zoomTargetEl = el; updateZoom(zoomLevel); },
                getTarget: function() { return zoomTargetEl; }
            } : null;

            // ======== 通过注册表查找处理器（替代硬编码 if-else）========
            // ★ 所有处理函数由 modules/file-viewers/ 下的扩展通过 IIFE 注册
            var handler = FILE_VIEWER_HANDLERS[ext];
            if (handler) {
                // ★ 检查设置中的开关
                var fvSettings = (this.plugin && this.plugin.settings && this.plugin.settings.fileViewerExtensions) || {};
                var extGroup = FILE_VIEWER_EXT_GROUPS[ext];
                if (extGroup && fvSettings[extGroup] === false) {
                    // ★ 用户关闭了此扩展，自动用系统默认程序打开
                    try {
                        var fullPath = adapter.getFullPath(filePath);
                        require('electron').shell.openPath(fullPath);
                    } catch(e) {
                        new Notice('无法打开文件: ' + e.message);
                    }
                    // 关闭当前视图页签
                    this.leaf.detach();
                    return;
                }
                // 调用处理器（this 保持为 DashboardFileViewer 实例）
                await handler.call(this, contentArea, file, ext, zoomCtx, vault);
            } else {
                // 未知类型 → 尝试当文本读
                await this._renderUnknown(contentArea, file, vault);
            }
        } catch(e) {
            showFileViewerErr(contentArea, '⚠ 读取失败: ' + e.message);
            console.error('[DashboardFileViewer] 渲染错误:', e);
        }
    }

    // ★ 兜底：未知文件类型当纯文本处理
    async _renderUnknown(area, file, vault) {
        try {
            var raw = await vault.read(file);
            var pre = area.createEl('pre', { cls: 'dbfv-code',
                attr: { style: 'margin:0;padding:16px;font-family:var(--font-monospace);font-size:12px;line-height:1.6;white-space:pre-wrap;word-break:break-all;color:var(--text-muted);min-height:100%;' }
            });
            pre.textContent = raw;
        } catch(e) {
            showFileViewerErr(area, '⚠ 无法读取文件: ' + e.message);
        }
    }


    async onClose() {
        this.contentEl.empty();
    }
}

// ============ 开发模式：FileViewer 扩展动态加载器 ============
// ★ 架构说明：
//   - release 模式：build.js 将 modules/file-viewers/*.js 内联到 main.js，__FV_BUILTIN__=true
//     → 此函数直接返回，无需任何操作
//   - dev 模式：build.js 不内联 file-viewers，__FV_BUILTIN__=false
//     → 此函数在插件启动时从磁盘读取并 eval 每个 file-viewer JS 文件
// ★ 效果：dev 模式下修改 file-viewer 代码后，只需 Ctrl+R 重载插件即可生效，无需 build！
//
// 使用方式（08-plugin.js onload 中调用）：
//   await initFileViewers(this);
//
async function initFileViewers(plugin) {
    // release模式：代码已由 build.js 内联到 main.js 中
    if (typeof __FV_BUILTIN__ !== 'undefined' && __FV_BUILTIN__) {
        console.log('[FV] 内置模式：FileViewer 扩展已内联到 main.js');
        return;
    }

    // dev模式：从文件系统动态加载
    var fvDir = plugin.manifest.dir + '/modules/file-viewers/';
    console.log('[FV] Dev 模式：开始动态加载 FileViewer 扩展...');

    try {
        var result = await plugin.app.vault.adapter.list(fvDir);
        var files = (result.files || []).filter(function(f) { return f.endsWith('.js'); }).sort();

        if (files.length === 0) {
            console.warn('[FV] modules/file-viewers/ 目录为空或不存在，无扩展加载');
            return;
        }

        var loaded = 0;
        for (var i = 0; i < files.length; i++) {
            var filePath = files[i];
            try {
                var content = await plugin.app.vault.adapter.read(filePath);
                // eval 执行：file-viewer 文件通过 IIFE 注册到 FILE_VIEWER_HANDLERS
                eval(content);
                loaded++;
                console.log('[FV] ✓', filePath.replace(fvDir, ''));
            } catch(e) {
                console.warn('[FV] ✗ 加载失败:', filePath.replace(fvDir, ''), e.message || e);
            }
        }

        console.log('[FV] 动态加载完成:', loaded + '/' + files.length, '个扩展可用');
    } catch(e) {
        console.warn('[FV] 扫描 file-viewers 目录失败:', e.message || e);
    }
}

// ============ 开发模式：Gallery 处理器动态加载器 ============
// ★ 架构说明（与 FileViewer 动态加载一致）：
//   - release 模式：build.js 将 src/utils/*.js 内联到 main.js，__GALLERY_BUILTIN__=true
//     → 此函数直接返回，无需任何操作
//   - dev 模式：build.js 不内联处理器代码，__GALLERY_BUILTIN__=false
//     → 此函数从 src/utils/ 目录读取并 eval 每个处理器 JS 文件
// ★ 效果：dev 模式下修改处理器代码后，只需 Ctrl+R 重载插件即可生效，无需 build！
//
async function initGalleryProcessors(plugin) {
    // release模式：代码已由 build.js 内联到 main.js 中
    if (typeof __GALLERY_BUILTIN__ !== 'undefined' && __GALLERY_BUILTIN__) {
        console.log('[Gallery] 内置模式：处理器已内联到 main.js');
        return;
    }

    var utilsDir = plugin.manifest.dir;
    if (utilsDir) {
        if (!utilsDir.endsWith('/') && !utilsDir.endsWith('\\')) {
            utilsDir += '/';
        }
        utilsDir += 'src/utils/';
    } else {
        console.warn('[Gallery] manifest.dir is undefined');
        return;
    }
    console.log('[Gallery] Dev 模式：开始动态加载 Gallery 处理器...');
    console.log('[Gallery] utilsDir:', utilsDir);

    var files = ['img-gallery-processor.js', 'memories-processor.js'];
    var loaded = 0;

    for (var i = 0; i < files.length; i++) {
        try {
            var filePath = utilsDir + files[i];
            console.log('[Gallery] Loading:', filePath);
            var content = await plugin.app.vault.adapter.read(filePath);
            console.log('[Gallery] Content length:', content.length);
            eval(content);
            loaded++;
            console.log('[Gallery] ✓', files[i]);
        } catch(e) {
            console.warn('[Gallery] ✗ 加载失败:', files[i], e.message || e);
        }
    }
    console.log('[Gallery] 动态加载完成:', loaded + '/' + files.length, '个处理器可用');
}

var __FV_BUILTIN__ = true;

// ===================== FileViewer 文件查看器扩展（modules/file-viewers/）=====================
// ============ FileViewer 扩展：旧版 DOC（docstream + CFB + 全流扫描）============

    // ============ .doc 文本提取辅助函数 ============

    // 构建 .doc 的渲染 HTML
    function buildDocHtml(text, metadata, method) {
        var badgeText = '旧版 Word 97-2003 (.doc)';
        if (method) badgeText += ' — ' + method + ' 解析';
        if (metadata && metadata.author) badgeText += ' | 作者: ' + metadata.author;
        if (metadata && metadata.title) badgeText += ' | 标题: ' + metadata.title;

        // 智能分段：按双换行拆段落，单换行保留在段内
        var rawParagraphs = text.split(/\n\s*\n/);
        var bodyHtml = '';
        for (var i = 0; i < rawParagraphs.length; i++) {
            var p = rawParagraphs[i].trim();
            if (!p) continue;

            // 检测可能为标题的行（长度 <= 80 且不以标点结尾）
            var lines = p.split('\n');
            var isHeading = (lines.length === 1 && p.length <= 80 &&
                !/[，。；：、！？,.;:!?]$/.test(p.trimEnd()) &&
                !/^[a-z]/.test(p.trimStart()));

            // 常见中文标题模式
            if (!isHeading && lines.length === 1 && p.length <= 100) {
                if (/^(第[一二三四五六七八九十百千\d]+[章节条篇]|[\d一二三四五六七八九十]+[\.\、\)）]|一[\.\、]|二[\.\、]|三[\.\、]|四[\.\、]|五[\.\、]|六[\.\、]|七[\.\、]|八[\.\、]|九[\.\、]|十[\.\、])/.test(p)) {
                    isHeading = true;
                }
            }

            if (isHeading) {
                bodyHtml += '<h3 class="dbfv-doc-heading">' + escapeHtml(p) + '</h3>';
            } else {
                // 段内换行转 <br>，转义 HTML
                var escapedLines = escapeHtml(p).replace(/\n/g, '<br>');
                bodyHtml += '<p class="dbfv-doc-para">' + escapedLines + '</p>';
            }
        }

        // ★ DOC 专属样式表（类似 DOCX 的排版风格）
        var docStyles = '<style>' +
            '.dbfv-doc-body { width:100% !important; min-width:100% !important; box-sizing:border-box !important; }' +
            '.dbfv-doc-body > * { max-width:none !important; }' +
            '.dbfv-doc-badge { margin-bottom:20px; padding:8px 12px; background:var(--background-modifier-form-field); border-radius:6px; font-size:11px; color:var(--text-muted); }' +
            '.dbfv-doc-heading { font-size:18px; font-weight:700; margin:24px 0 12px; padding-bottom:8px; border-bottom:1px solid var(--background-modifier-border); color:var(--text-normal); line-height:1.4; }' +
            '.dbfv-doc-para { margin:0 0 12px; text-indent:2em; line-height:1.9; color:var(--text-normal); }' +
            '.dbfv-doc-para:first-of-type { margin-top:0; }' +
            '</style>';

        return docStyles +
            '<div class="dbfv-doc-body">' +
            '<div class="dbfv-doc-badge">' + escapeHtml(badgeText) + '</div>' +
            (bodyHtml || ('<p class="dbfv-doc-para">' + escapeHtml(text).replace(/\n/g, '<br>') + '</p>')) +
            '</div>';
    }

    // UTF-16LE + ANSI 混合文本扫描：从二进制中提取可读文本序列
    //   ★ 关键改进：自动跳过图像二进制块（JPEG/PNG/GIF/BMP magic bytes）
    //   降低文本运行阈值到2个连续字符，更宽容地提取分散文本
    function scanForDocText(data) {
        if (!data || data.length < 10) return '';

        // --- 辅助：检测并跳过图像二进制块 ---
        function isImageBlock(buf, idx) {
            // JPEG: FF D8 FF
            if (buf[idx] === 0xFF && buf[idx+1] === 0xD8 && buf[idx+2] === 0xFF) return true;
            // PNG: 89 50 4E 47
            if (buf[idx] === 0x89 && buf[idx+1] === 0x50 && buf[idx+2] === 0x4E && buf[idx+3] === 0x47) return true;
            // GIF: 47 49 46 38
            if (buf[idx] === 0x47 && buf[idx+1] === 0x49 && buf[idx+2] === 0x46 && buf[idx+3] === 0x38) return true;
            // BMP: 42 4D
            if (buf[idx] === 0x42 && buf[idx+1] === 0x4D) return true;
            // EMF/WMF: 01 00 00 00 或 D7 CD C6 9A
            if (buf[idx] === 0xD7 && buf[idx+1] === 0xCD && buf[idx+2] === 0xC6 && buf[idx+3] === 0x9A) return true;
            return false;
        }

        function skipImageBlock(buf, start) {
            var len = buf.length;
            // 查找下一个明显的文本边界（两个连续的\0或\x20+）
            for (var i = start + 4; i < len - 2; i++) {
                if (buf[i] >= 0x20 && buf[i+1] >= 0x20) return i;
            }
            return len;
        }

        // --- 先尝试 UTF-16LE ---
        var utf16Text = '';
        var len = data.length - 1;
        var consecutive = 0, runStart = -1;
        var i = 0;

        while (i < len) {
            // 跳过图像块
            if (i < len - 4 && isImageBlock(data, i)) {
                if (runStart >= 0 && consecutive >= 2) {
                    if (utf16Text) utf16Text += '\n';
                    for (var j = runStart; j < i; j += 2) {
                        var c = data[j] | (data[j + 1] << 8);
                        if (c === 0x0D) utf16Text += '\n';
                        else if (c === 0x07) utf16Text += '\t';
                        else if (c >= 0x20 && c < 0xD800) utf16Text += String.fromCharCode(c);
                    }
                }
                i = skipImageBlock(data, i);
                runStart = -1; consecutive = 0;
                continue;
            }

            var code = data[i] | (data[i + 1] << 8);
            var literal = (code === 0x0D || code === 0x07) ||
                          (code >= 0x20 && code < 0xD800) ||
                          (code >= 0xE000 && code < 0xFFFE);
            if (literal) {
                if (code >= 0x20) consecutive++;
                if (runStart < 0) runStart = i;
            } else {
                if (runStart >= 0 && consecutive >= 2) {
                    if (utf16Text) utf16Text += '\n';
                    for (var j = runStart; j < i; j += 2) {
                        var c = data[j] | (data[j + 1] << 8);
                        if (c === 0x0D) utf16Text += '\n';
                        else if (c === 0x07) utf16Text += '\t';
                        else if (c >= 0x20 && c < 0xD800) utf16Text += String.fromCharCode(c);
                    }
                }
                runStart = -1; consecutive = 0;
            }
            i += 2;
        }

        if (utf16Text.trim().length > 30) {
            console.log('[DFV] UTF-16LE 扫描成功, 长度:', utf16Text.length);
            return utf16Text.trim();
        }

        // --- UTF-16LE 效果差 → 尝试 ANSI/CP1252 ---
        console.log('[DFV] UTF-16LE 仅扫描到', utf16Text.length, '字符，尝试 ANSI');
        var ansiText = '';
        var ansiConsecutive = 0, ansiRunStart = -1;
        var i = 0;

        while (i < data.length) {
            // 跳过图像块
            if (i < data.length - 4 && isImageBlock(data, i)) {
                if (ansiRunStart >= 0 && ansiConsecutive >= 2) {
                    if (ansiText) ansiText += '\n';
                    for (var k = ansiRunStart; k < i; k++) {
                        var c = data[k];
                        if (c === 0x0D) ansiText += '\n';
                        else if (c === 0x0A) { /* skip LF, already handled by CR */ }
                        else if (c === 0x09) ansiText += '\t';
                        else if (c >= 0x20 && c < 0x7F) ansiText += String.fromCharCode(c);
                    }
                }
                i = skipImageBlock(data, i);
                ansiRunStart = -1; ansiConsecutive = 0;
                continue;
            }

            var b = data[i];
            var lit = (b >= 0x20 && b < 0x7F) || b === 0x0D || b === 0x0A || b === 0x09;
            if (lit) {
                if (b >= 0x20) ansiConsecutive++;
                if (ansiRunStart < 0) ansiRunStart = i;
            } else {
                if (ansiRunStart >= 0 && ansiConsecutive >= 2) {
                    if (ansiText) ansiText += '\n';
                    for (var k = ansiRunStart; k < i; k++) {
                        var c = data[k];
                        if (c === 0x0D) ansiText += '\n';
                        else if (c === 0x0A) { /* skip LF, already handled by CR */ }
                        else if (c === 0x09) ansiText += '\t';
                        else if (c >= 0x20 && c < 0x7F) ansiText += String.fromCharCode(c);
                    }
                }
                ansiRunStart = -1; ansiConsecutive = 0;
            }
            i++;
        }
        console.log('[DFV] ANSI 扫描结果长度:', ansiText.length);
        return (ansiText.trim().length > 0) ? ansiText.trim() : (utf16Text.trim());
    }

    // 旧版 Word 97-2003 .doc：三重策略（docstream → CFB+扫描 → 全流兜底）
    async function _renderDoc(area, file, ext, zoomCtx, vault) {
        area.style.cssText += 'background:var(--background-primary);overflow:hidden;user-select:text;-webkit-user-select:text;';
        var wrapEl = area.createEl('div', { cls: 'dbfv-doc-content',
            attr: { style: 'width:100%;min-width:100%;box-sizing:border-box;height:100%;overflow:auto;padding:16px 12px;font-size:13px;line-height:1.8;color:var(--text-normal);user-select:text;-webkit-user-select:text;' }
        });
        wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">⏳ 正在加载文档引擎...</div>';

        try {
            var data = await this.app.vault.readBinary(file);
            var u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
            console.log('[DFV] DOC 文件大小:', u8.byteLength, 'bytes');

            // ============================================================
            // 策略一：docstream → AST 解析（质量最高，含元数据）
            // ============================================================
            var officeParser = await getDocstreamAsync();
            if (officeParser) {
                wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">⏳ docstream 正在解析...</div>';
                try {
                    var ab = u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
                    var ast = await officeParser.parseOffice(ab);
                    var docText = ast.toText ? ast.toText() : (ast.text || '');
                    if (docText && docText.trim().length > 0) {
                        wrapEl.innerHTML = buildDocHtml(docText, ast.metadata || null, 'docstream');
                        console.log('[DFV] .doc docstream 解析成功, 文本长度:', docText.length);
                        return;
                    }
                    console.log('[DFV] docstream 返回空文本，降级到 CFB');
                } catch(dsErr) {
                    console.warn('[DFV] docstream 解析失败:', String(dsErr).substring(0,120));
                }
            }

            // ============================================================
            // 策略二：CFB OLE2 解析 → WordDocument 流文本扫描
            // ============================================================
            var CFB = await getCfbAsync();
            if (!CFB) {
                wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-error);padding:40px 0;">⚠ 文档解析库加载失败<br><small>请检查网络连接后刷新</small></div>';
                return;
            }

            wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">⏳ CFB 正在解析 OLE2...</div>';

            var cfb = null, dirs = [];
            try {
                cfb = CFB.read(u8, { type: 'array' });
                dirs = cfb.Directory || [];
            } catch(cfbErr) {
                console.warn('[DFV] CFB.read 失败:', String(cfbErr).substring(0,120));
            }

            if (dirs.length > 0) {
                // 查找 WordDocument 流（.doc 主文本存储）
                var wdStream = null;
                for (var i = 0; i < dirs.length; i++) {
                    if (dirs[i].name === 'WordDocument' && dirs[i].content) {
                        wdStream = dirs[i].content;
                        break;
                    }
                }

                if (wdStream && wdStream.length >= 10) {
                    console.log('[DFV] 找到 WordDocument 流, 大小:', wdStream.length, 'bytes');
                    var magic = wdStream[0] | (wdStream[1] << 8);
                    console.log('[DFV] WordDocument magic:', '0x' + magic.toString(16).toUpperCase());

                    wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">⏳ 扫描文本...</div>';
                    var scanned = scanForDocText(wdStream);
                    if (scanned && scanned.trim().length > 0) {
                        wrapEl.innerHTML = buildDocHtml(scanned, null, 'CFB 文本扫描');
                        console.log('[DFV] .doc CFB 扫描成功, 文本长度:', scanned.length);
                        return;
                    }
                    console.log('[DFV] WordDocument 流扫描无结果，降级到全流扫描');
                } else {
                    console.log('[DFV] 未找到 WordDocument 流（CFB dirs:', dirs.length, '个条目）');
                }

                // ============================================================
                // 策略三：全流扫描（遍历所有 OLE2 流找文本）
                // ============================================================
                wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">⏳ 全流扫描...</div>';
                var allText = '';
                for (var k = 0; k < dirs.length; k++) {
                    var d = dirs[k];
                    if (d.type !== 2 || !d.content || d.content.length < 10) continue;
                    var t = scanForDocText(d.content);
                    if (t && t.trim()) allText += (allText ? '\n--- ' + (d.name || '流#' + k) + ' ---\n' : '') + t;
                }

                if (allText.trim()) {
                    wrapEl.innerHTML = buildDocHtml(allText, null, '全流扫描');
                    console.log('[DFV] .doc 全流扫描完成, 文本长度:', allText.length);
                    return;
                }
            }

            // ============================================================
            // 策略四：原始二进制全文扫描（CFB 无法解析或所有流无文本）
            // ============================================================
            console.log('[DFV] OLE2 扫描无结果，尝试原始二进制扫描');
            wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">⏳ 原始二进制扫描...</div>';
            var rawText = scanForDocText(u8);
            if (rawText && rawText.trim().length > 0) {
                wrapEl.innerHTML = buildDocHtml(rawText, null, '原始二进制扫描');
                console.log('[DFV] .doc 原始二进制扫描完成, 文本长度:', rawText.length);
                return;
            }

            wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">⚠ 未能提取到文本内容<br><small>该文档可能损坏、加密或仅包含图片</small></div>';
        } catch(e) {
            wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-error);padding:40px 0;">⚠ 解析 .doc 失败: ' + escapeHtml(String(e).substring(0,200)) + '</div>';
            console.error('[DashboardFileViewer] .doc 解析失败:', e);
        }
    }

(function() {
    FILE_VIEWER_HANDLERS["doc"] = _renderDoc;
})();

// ============ FileViewer 扩展：DOCX（mammoth.js + JSZip 兜底）============

    // Word 文档：mammoth.js 渲染 DOCX（支持 zoom 缩放）
    async function _renderDocx(area, file, ext, zoomCtx, vault) {
        area.style.cssText += 'background:var(--background-primary);overflow:hidden;';
        var wrapEl = area.createEl('div', { cls: 'dbfv-docx-content',
            attr: { style: 'width:100%;min-width:100%;box-sizing:border-box;height:100%;overflow:auto;padding:16px 12px;font-size:13px;line-height:1.8;color:var(--text-normal);user-select:text;-webkit-user-select:text;' }
        });
        // ★ 连接 zoom 控件
        if (zoomCtx) zoomCtx.setTarget(wrapEl);
        wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">⏳ 正在加载文档引擎...</div>';

        var mammoth = await getMammothAsync();
        if (!mammoth) {
            wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-error);padding:40px 0;">⚠ Word解析库(mammoth.js)加载失败<br><small>请检查网络连接后刷新</small></div>';
            return;
        }

        try {
            var data = await this.app.vault.readBinary(file);
            // ★ mammoth 现在在主上下文执行，ArrayBuffer 不再有跨上下文问题
            // readBinary 返回 ArrayBuffer，确保以正确类型传递
            var arrayBuffer;
            if (data instanceof ArrayBuffer) {
                arrayBuffer = data;
            } else if (data instanceof Uint8Array) {
                arrayBuffer = data.buffer;
            } else if (data && data.buffer) {
                arrayBuffer = data.buffer;
            } else {
                arrayBuffer = new Uint8Array(data).buffer;
            }
            console.log('[DFV] DOCX ArrayBuffer 大小:', arrayBuffer.byteLength, 'bytes');
            var result;
            try {
                result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer }, {
                    styleMap: [
                        "p[style-name='Heading 1'] => h1:fresh",
                        "p[style-name='Heading 2'] => h2:fresh",
                        "p[style-name='Heading 3'] => h3:fresh",
                        "r[style-name='Strong'] => strong",
                        "r[style-name='Emphasis'] => em"
                    ]
                });
            } catch(mamErr) {
                // ★ 容错：内部文件引用缺失时（Could not find file in options）
                // 策略2：降级为纯文本提取
                console.warn('[DFV] mammoth HTML 转换失败，降级为纯文本:', String(mamErr).substring(0,120));
                try {
                    var rawResult = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
                    result = { value: '<p>' + escapeHtml(rawResult.value).replace(/\n/g, '</p><p>') + '</p>', messages: rawResult.messages };
                    console.log('[DFV] mammoth 纯文本提取成功');
                } catch(rawErr) {
                    console.warn('[DFV] mammoth extractRawText 也失败:', String(rawErr).substring(0,120));
                    // 策略3：JSZip 手动解包 → 读 word/document.xml → 剥离 XML 标签
                    try {
                        var fallbackText = await extractDocxTextViaJSZip(arrayBuffer);
                        result = { value: buildDocxFallbackHtml(fallbackText), messages: [] };
                        console.log('[DFV] JSZip 手动解包成功, 文本长度:', fallbackText.length);
                    } catch(jszipErr) {
                        console.warn('[DFV] JSZip 解包也失败:', String(jszipErr).substring(0,120));
                        throw jszipErr;  // 抛到外层统一处理
                    }
                }
            }

            // 注入基础样式（★ 强制内容填满容器，覆盖原文档可能的居中/缩进）
            var docStyles = '<style>' +
                '.dbfv-docx-content { width:100% !important; min-width:100% !important; box-sizing:border-box !important; }' +
                '.dbfv-docx-content * { max-width:none !important; box-sizing:border-box !important; }' +
                '.dbfv-docx-content h1 { font-size:24px; margin:16px 0 12px; border-bottom:1px solid var(--background-modifier-border); padding-bottom:8px; }' +
                '.dbfv-docx-content h2 { font-size:20px; margin:14px 0 10px; }' +
                '.dbfv-docx-content h3 { font-size:16px; margin:12px 0 8px; }' +
                '.dbfv-docx-content p { margin:0 0 8px; }' +
                '.dbfv-docx-content table { border-collapse:collapse; width:100% !important; margin:12px 0; }' +
                '.dbfv-docx-content th,.dbfv-docx-content td { border:1px solid var(--background-modifier-border); padding:6px 10px; text-align:left; font-size:12px; }' +
                '.dbfv-docx-content th { background:var(--background-modifier-form-field); font-weight:600; }' +
                '.dbfv-docx-content img { max-width:100% !important; height:auto !important; }' +
                '.dbfv-docx-content ul,.dbfv-docx-content ol { padding-left:24px; margin-bottom:8px; }' +
                '.dbfv-docx-content blockquote { border-left:3px solid var(--v6-primary); padding-left:12px; margin:8px 0; color:var(--text-muted); }' +
                '</style>';
            wrapEl.innerHTML = docStyles + result.value;

            if (result.messages && result.messages.length > 0) {
                console.log('[DashboardFileViewer] mammoth 警告:', result.messages);
            }
        } catch(e) {
            var errMsg = String(e.message || e).substring(0, 200);
            wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">⚠ 解析文档失败<br><small>' + escapeHtml(errMsg) + '</small><br><small style="margin-top:8px;display:block;">该 DOCX 可能包含损坏的内部引用，建议用 WPS 另存为新格式</small></div>';
            console.error('[DashboardFileViewer] DOCX渲染失败:', e);
        }
    }

    // 当 mammoth 的两重策略（convertToHtml + extractRawText）都失败时使用
    async function extractDocxTextViaJSZip(arrayBuffer) {
        var JSZip = await getJSZipAsync();
        if (!JSZip) throw new Error('JSZip 加载失败');

        var zip = await JSZip.loadAsync(arrayBuffer);

        // 优先读取 word/document.xml（正文）
        var docXml = null;
        var docFile = zip.file('word/document.xml');
        if (docFile) {
            docXml = await docFile.async('string');
        }

        // 兜底：尝试 header/footer
        var headerXml = '';
        try {
            var hFiles = zip.filter(function(relativePath) { return /word\/header\d*\.xml/.test(relativePath); });
            if (hFiles && hFiles.length > 0) {
                headerXml = await hFiles[0].async('string');
            }
        } catch(e) {}

        if (!docXml && !headerXml) throw new Error('DOCX 内未找到 word/document.xml');

        // 合并所有 XML 文本，按段落顺序提取
        var xml = (docXml || '') + '\n' + (headerXml || '');

        // 正则剥离：匹配 <w:t ...> 标签内容，按 <w:p> 边界分段
        var paragraphs = xml.split(/<\/w:p\s*>/i);
        var lines = [];

        for (var i = 0; i < paragraphs.length; i++) {
            var p = paragraphs[i];
            // 提取 <w:t>content</w:t> 或 <w:t ...>content</w:t>
            var tMatch = p.match(/<w:t[^>]*>([\s\S]*?)<\/w:t\s*>/gi);
            if (tMatch) {
                var lineText = '';
                for (var j = 0; j < tMatch.length; j++) {
                    var inner = tMatch[j].replace(/<w:t[^>]*>/gi, '').replace(/<\/w:t\s*>/gi, '');
                    lineText += inner;
                }
                var trimmed = lineText.trim();
                if (trimmed) lines.push(trimmed);
            }
        }

        if (lines.length === 0) throw new Error('DOCX XML 中未提取到文本');
        return lines.join('\n\n');
    }

    // DOCX 兜底渲染 HTML 生成（类似 buildDocHtml 但用 DOCX 样式类名）
    function buildDocxFallbackHtml(text) {
        var rawParagraphs = text.split(/\n\s*\n/);
        var bodyHtml = '';
        for (var i = 0; i < rawParagraphs.length; i++) {
            var p = rawParagraphs[i].trim();
            if (!p) continue;
            bodyHtml += '<p class="dbfv-docx-fallback-p">' + escapeHtml(p).replace(/\n/g, '<br>') + '</p>';
        }

        var docStyles = '<style>' +
            '.dbfv-docx-fallback-badge { margin-bottom:20px; padding:8px 12px; background:var(--background-modifier-form-field); border-radius:6px; font-size:11px; color:var(--text-muted); }' +
            '.dbfv-docx-fallback-p { margin:0 0 8px; text-indent:2em; line-height:1.9; color:var(--text-normal); }' +
            '</style>';

        return docStyles +
            '<div class="dbfv-docx-content" style="width:100% !important;min-width:100% !important;box-sizing:border-box !important;">' +
            '<div class="dbfv-docx-fallback-badge">DOCX — 手动解包提取文本（原始解析失败，已降级）</div>' +
            (bodyHtml || '<p class="dbfv-docx-fallback-p">' + escapeHtml(text).replace(/\n/g, '<br>') + '</p>') +
            '</div>';
    }

(function() {
    FILE_VIEWER_HANDLERS["docx"] = _renderDocx;
})();

// ============ FileViewer 扩展：HTML（iframe + srcdoc + 基础安全清洗）============

    // HTML：iframe + srcdoc（基础安全清洗） + zoom 支持
    async function _renderHtml(area, file, ext, zoomCtx, vault) {
        var content = await vault.read(file);
        // 基础清洗：移除 script 标签和内联事件
        var safe = content
            .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
            .replace(/\son\w+\s*=\s*"[^"]*"/gi, '')
            .replace(/\son\w+\s*=\s*'[^']*'/gi, '');

        // 外层 wrapper（zoom 作用于 wrapper，iframe 内容由 browser 自行缩放）
        var wrapper = area.createEl('div', { cls: 'dbfv-html-wrapper',
            attr: { style: 'width:100%;height:100%;overflow:auto;' }
        });
        if (zoomCtx) zoomCtx.setTarget(wrapper);

        var iframe = wrapper.createEl('iframe', { cls: 'dbfv-iframe',
            attr: { style: 'width:100%;height:100%;border:none;background:#fff;', sandbox: 'allow-same-origin' }
        });
        iframe.srcdoc = safe;
    }

(function() {
    var exts = ["html", "htm"];
    exts.forEach(function(ext) {
        FILE_VIEWER_HANDLERS[ext] = _renderHtml;
    });
})();

// ============ FileViewer 扩展：图片（PNG/JPG/GIF/SVG/WEBP/BMP/ICO）============

    // 图片
    function _renderImage(area, file, ext, zoomCtx, vault) {
        area.style.cssText += 'display:flex;align-items:center;justify-content:center;background:var(--background-secondary);';
        area.createEl('img', { cls: 'dbfv-image',
            attr: { src: vault.getResourcePath(file), style: 'max-width:100%;max-height:100%;object-fit:contain;' }
        });
    }

(function() {
    var exts = ["png", "jpg", "jpeg", "gif", "svg", "webp", "bmp", "ico"];
    exts.forEach(function(ext) {
        FILE_VIEWER_HANDLERS[ext] = _renderImage;
    });
})();

// ============ FileViewer 扩展：PPT / PPTX 查看器 ============
// PPTX（新版）：JSZip 解压 → 解析 ppt/slides/*.xml 中的 <a:t> 文本
// PPT （旧版）：OLE2流提取 + 暴力签名扫描 → 元数据过滤
// ★ 只提取和渲染文字，不解析/渲染图片、形状等媒体内容
// ★ 支持 zoom 缩放（Ctrl+滚轮 / 按钮）、文字选择复制

// ============ PPTX 解析：JSZip + XML 文本提取 ============
async function _extractPptxText(arrayBuffer) {
    var JSZip = await getJSZipAsync();
    if (!JSZip) throw new Error('JSZip 库加载失败，无法解压 PPTX');

    var zip = await JSZip.loadAsync(arrayBuffer);

    // 收集所有幻灯片文件（按数字排序）
    var slideFiles = [];
    zip.forEach(function(relativePath) {
        var match = relativePath.match(/^ppt\/slides\/slide(\d+)\.xml$/i);
        if (match) {
            slideFiles.push({ path: relativePath, index: parseInt(match[1], 10) });
        }
    });
    slideFiles.sort(function(a, b) { return a.index - b.index; });

    if (slideFiles.length === 0) throw new Error('PPTX 内未找到任何幻灯片文件 (ppt/slides/*.xml)');

    var slides = [];

    for (var i = 0; i < slideFiles.length; i++) {
        var slideFile = zip.file(slideFiles[i].path);
        if (!slideFile) continue;

        var xml = await slideFile.async('string');
        var paragraphs = extractTextFromPptxXml(xml);
        slides.push({
            index: slideFiles[i].index,
            paragraphs: paragraphs,
            hasContent: paragraphs.length > 0
        });
    }

    return slides;
}

// 从单张幻灯片 XML 中提取所有文本段落
function extractTextFromPptxXml(xml) {
    var paragraphs = [];

    // 按 <a:p> 段落标签分割
    var pBlocks = xml.split(/<\/a:p\s*>/i);
    for (var i = 0; i < pBlocks.length; i++) {
        var block = pBlocks[i];
        // 提取该段落内所有 <a:t> 标签的文本
        var tMatches = block.match(/<a:t[^>]*>([\s\S]*?)<\/a:t\s*>/gi);
        if (!tMatches || tMatches.length === 0) continue;

        var lineText = '';
        for (var j = 0; j < tMatches.length; j++) {
            var inner = tMatches[j].replace(/<a:t[^>]*>/gi, '').replace(/<\/a:t\s*>/gi, '');
            lineText += inner;
        }
        var trimmed = lineText.trim();
        if (trimmed) paragraphs.push(trimmed);
    }

    return paragraphs;
}

// 将提取到的幻灯片数据渲染为 HTML
function renderPptxSlidesHtml(slides) {
    var html = '';
    html += '<style type="text/css">' +
        '.dbfv-pptx-container { width:100%; padding:16px 20px; box-sizing:border-box; }' +
        '.dbfv-pptx-slide { margin-bottom:24px; padding:16px 20px; border-left:3px solid var(--v6-primary); background:var(--background-secondary); border-radius:0 8px 8px 0; page-break-inside:avoid; }' +
        '.dbfv-pptx-slide-header { display:flex; align-items:center; gap:10px; margin-bottom:12px; padding-bottom:8px; border-bottom:1px solid var(--background-modifier-border); }' +
        '.dbfv-pptx-slide-num { display:inline-flex; align-items:center; justify-content:center; min-width:28px; height:28px; background:var(--v6-primary); color:#fff; border-radius:50%; font-size:12px; font-weight:600; flex-shrink:0; }' +
        '.dbfv-pptx-slide-label { font-size:12px; color:var(--text-muted); font-weight:500; }' +
        '.dbfv-pptx-para { margin:6px 0; padding:4px 0; font-size:13.5px; line-height:1.8; color:var(--text-normal); text-indent:2em; }' +
        '.dbfv-pptx-para:first-of-type { text-indent:0; }' +
        '.dbfv-pptx-empty { font-style:italic; color:var(--text-faint); font-size:12px; padding:8px 0; }' +
        '</style>';

    html += '<div class="dbfv-pptx-container">';

    for (var i = 0; i < slides.length; i++) {
        var slide = slides[i];
        html += '<div class="dbfv-pptx-slide">';
        html += '<div class="dbfv-pptx-slide-header">';
        html += '<span class="dbfv-pptx-slide-num">' + slide.index + '</span>';
        html += '<span class="dbfv-pptx-slide-label">Slide ' + slide.index + (slide.hasContent ? '' : ' \u2014 \u7A7A\u767D\u9875') + '</span>';
        html += '</div>';

        if (slide.paragraphs.length > 0) {
            for (var j = 0; j < slide.paragraphs.length; j++) {
                html += '<p class="dbfv-pptx-para">' + escapeHtml(slide.paragraphs[j]) + '</p>';
            }
        } else {
            html += '<p class="dbfv-pptx-empty">（\u6B64\u5E7B\u706F\u7247\u65E0\u53EF\u63D0\u53D6\u7684\u6587\u5B57\u5185\u5BB9）</p>';
        }

        html += '</div>'; // .dbfv-pptx-slide
    }

    html += '</div>'; // .dbfv-pptx-container
    return html;
}

// ============ PPT 旧版格式解析（v8 架构）============
//
// ★ Step 1: OLE2 容器解析 → 提取 "PowerPoint Document" 流（消除85%垃圾源）
// ★ Step 2: 暴力签名扫描 → 遍历PP流每个字节，匹配文本记录头(0x0FA8/0x0FA0/0x0FAD)
// ★ Step 3: 轻量级元数据过滤 → 黑名单正则过滤模板文字
//
// 设计哲学：
//   PPT 的容器嵌套结构极其复杂，手写递归解析器不可靠（v7已验证失败）
//   但文本记录的 type 签名是固定的，暴力扫描 100% 可靠！
//   OLE2 流提取确保我们只扫描 38KB PP流而非全文件，效率足够

async function _extractPptText(arrayBuffer) {
    var rawBytes = new Uint8Array(arrayBuffer);
    console.log('[DFV-PPT] 开始解析 PPT, 文件大小:', rawBytes.length, 'bytes');

    // ══════════════════════════════════════
    //  Step 1: 验证 OLE2 签名 + 提取PP流
    // ══════════════════════════════════════
    var OLE_SIG = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    for (var s = 0; s < 8; s++) {
        if (rawBytes[s] !== OLE_SIG[s]) throw new Error('非有效的 PPT (OLE2) 文件格式');
    }

    var ppStream = _extractOle2Stream(rawBytes, 'PowerPoint Document');
    if (!ppStream || ppStream.length < 16) {
        throw new Error('无法在 PPT 文件中找到 PowerPoint Document 数据流');
    }
    console.log('[DFV-PPT] PowerPoint Document 流大小:', ppStream.length, 'bytes',
        '(原始文件', rawBytes.length, 'bytes, 瘦身', Math.round(100 * (1 - ppStream.length / rawBytes.length)) + '%)');

    // ══════════════════════════════════════
    //  Step 2: 暴力签名扫描（v8核心 — 替代v7的容器递归）
    //  遍历PP流每个字节位置，查找 TextCharsAtom(0x0FA8)/TextBytesAtom(0x0FA0)/CString(0x0FAD)
    // ══════════════════════════════════════
    var recordTexts = _scanForTextRecordSignatures(ppStream);
    console.log('[DFV-PPT] 签名扫描找到:', recordTexts.length, '段文本');

    // ══════════════════════════════════════
    //  Step 3: 元数据过滤
    // ══════════════════════════════════════
    var filtered = _pptMetadataFilter(recordTexts);
    console.log('[DFV-PPT] 过滤后有效文本:', filtered.length, '段');

    if (filtered.length === 0) {
        throw new Error('未能从 PPT 中提取到有效文本内容（可能该文件只包含图片/形状而无文字）');
    }

    return splitPptTextIntoSlides(filtered.join('\n'));
}

// ================================================================
//  OLE2 流提取器（轻量级 — 只需读取目录和目标流）
// ================================================================
function _extractOle2Stream(data, streamName) {
    // 读取 OLE2 头部参数
    var sectorSize = 1 << (_readU16(data, 30));   // 通常512
    var dirStartSec = _readU32(data, 48);          // 目录起始扇区

    // 辅助函数：读取扇区数据
    function readSector(s) {
        return data.subarray(512 + s * sectorSize, 512 + (s + 1) * sectorSize);
    }

    // 解析 FAT 链
    function followChain(startSector) {
        var chain = [], visited = {}, s = startSector;
        while (s !== 0xFFFFFFFE && s !== 0xFFFFFFFC && s >= 0 && !visited[s]) {
            visited[s] = true;
            chain.push(s);
            // 从 FAT 中获取下一扇区号
            var fatIdx = Math.floor(s / (sectorSize / 4));
            var entryIdx = s % (sectorSize / 4);
            if (fatIdx * sectorSize + entryIdx * 4 + 3 >= data.length) break;
            s = _readU32(readSector(fatIdx), entryIdx * 4);
        }
        return chain;
    }

    // 读取目录流
    var dirChain = followChain(dirStartSec);
    var dirData = [];
    for (var di = 0; di < dirChain.length; di++) {
        var secData = readSector(dirChain[di]);
        for (var si = 0; si < secData.length; si++) dirData.push(secData[si]);
    }
    dirData = new Uint8Array(dirData);

    // 遍历目录项（每项128字节），查找目标流
    for (var ei = 0; ei + 128 <= dirData.length; ei += 128) {
        var nameLen = dirData[ei + 0x40] | (dirData[ei + 0x41] << 8);
        if (nameLen === 0 || nameLen > 64) continue;
        
        // 读取名称（UTF-16LE）
        var entryName = '';
        for (var ni = 0; ni < nameLen; ni += 2) {
            entryName += String.fromCharCode(dirData[ei + ni] | (dirData[ei + ni + 1] << 8));
        }
        entryName = entryName.replace(/\0/g, '');

        if (entryName === streamName) {
            var entryType = dirData[ei + 0x42];  // 2=Stream
            if (entryType !== 2) return null;

            var startSector = _readU32(dirData, ei + 0x74);
            var size = Number(_readU64Raw(dirData, ei + 0x78));

            if (size === 0) return new Uint8Array(0);

            // 小于 mini stream cutoff 的使用 mini FAT
            var miniCutOff = _readU32(data, 56);

            var chain2;
            if (size < miniCutOff && sectorSize > 0x1000) {
                // Mini stream — 使用 Mini FAT（简化处理：直接返回空，极少情况触发）
                return new Uint8Array(0);
            } else {
                chain2 = followChain(startSector);
            }

            // 拼接所有扇区
            var result = [];
            for (var ci = 0; ci < chain2.length; ci++) {
                var sd = readSector(chain2[ci]);
                var remaining = size - result.length;
                for (var bi = 0; bi < sd.length && bi < remaining; bi++) {
                    result.push(sd[bi]);
                }
            }
            return new Uint8Array(result);
        }
    }
    
    return null;
}

// ================================================================
//  v8 核心：暴力签名扫描器（替代v7的容器递归解析）
// ================================================================
//
// ★ 为什么不用容器递归？
//   PPT 的容器嵌套结构极其复杂（Document→Slide→Shape→...）
//   手写递归解析器在第2层就会因为 rlen 异常而迷失（v7已验证）
//
// ★ 暴力扫描为什么可靠？
//   文本记录的 type 标识符是固定的16位值：
//     0x0FA8 = TextCharsAtom (UTF-16LE 文本)
//     0x0FA0 = TextBytesAtom (ANSI/单字节文本)
//     0x0FAD = CString (null-terminated UTF-16LE)
//   这些值在二进制中几乎不可能随机出现 + 配合合理rlen = 100%精确
//
// ★ 性能：PP流通常 10~80KB，遍历每个字节 < 1ms，完全可接受
//
function _scanForTextRecordSignatures(ppStream) {
    var results = [];
    var limit = ppStream.length;

    // 只需检查到 limit-8，确保能读取完整的8字节记录头
    for (var pos = 0; pos + 8 <= limit; pos++) {
        var recType = _readU16(ppStream, pos + 2);
        var recLen = _readU32(ppStream, pos + 4);

        // 快速过滤：只处理3种文本记录类型
        if (recType !== 0x0FA8 && recType !== 0x0FA0 && recType !== 0x0FAD) continue;

        // 合理性检查：长度必须在有效范围内
        if (recLen <= 0 || recLen > 10000) continue;
        if (pos + 8 + recLen > limit) continue;

        var bodyStart = pos + 8;
        var recData = ppStream.slice(bodyStart, bodyStart + recLen);

        var text = '';
        switch (recType) {
            case 0x0FA8:  // TextCharsAtom — UTF-16LE
                text = _decodeUtf16Atom(recData);
                break;
            case 0x0FA0:  // TextBytesAtom — ANSI/单字节
                text = _decodeTextBytesAtom(recData);
                break;
            case 0x0FAD:  // CString — null-terminated UTF-16LE
                text = _decodeCString(recData);
                break;
        }

        if (text && text.length >= 1) {
            results.push(text.trim());
        }
    }

    return results;
}

// ================================================================
//  二进制读取辅助函数
// ================================================================
function _readU16(buf, off) {
    return buf[off] | (buf[off + 1] << 8);
}

function _readU32(buf, off) {
    return ((buf[off]|(buf[off+1]<<8)|(buf[off+2]<<16)|(buf[off+3]<<24))>>>0);
}

function _readU64Raw(buf, off) {
    // 返回近似值（JavaScript 不支持 64 位整数，只用于比较）
    var lo = _readU32(buf, off);
    var hi = _readU32(buf, off + 4);
    return hi > 0 ? hi * 4294967296 + lo : lo;
}

// ================================================================
//  文本解码器
// ================================================================

// 解码 TextCharsAtom (type 0x0FA8)
function _decodeUtf16Atom(data) {
    if (data.length < 2) return '';

    var start = 0;
    if (data.length >= 4) {
        var possibleLen = data[0] | (data[1] << 8);
        if (possibleLen >= 2 && possibleLen <= 10000 && possibleLen * 2 <= data.length - 0) {
            start = 2;
        }
    }

    var chars = [];
    for (var i = start; i + 1 < data.length; i += 2) {
        var code = data[i] | (data[i + 1] << 8);
        if (code === 0) break;
        if (code < 32 && code !== 9 && code !== 10 && code !== 13) continue;
        chars.push(String.fromCharCode(code));
    }

    var r = chars.join('').trim();
    return (r.length >= 1) ? r : '';
}

// 解码 TextBytesAtom (type 0x0FA0)
// ★ v8.1 关键修复：自动检测 UTF-16LE vs 单字节编码
//   PPT 的 TextBytesAtom 实际存储格式取决于文件的字符集设置：
//   - 某些文件用 ANSI/单字节（纯 ASCII 值，无 0x00 间隔）
//   - **更多文件用 UTF-16LE！（即使 type=0x0FA0 也可能存 UTF-16LE）**
//   检测方法：采样前20字节，如果 >=25% 的奇数位是 0x00 → 判定为 UTF-16LE
function _decodeTextBytesAtom(data) {
    if (!data || data.length < 1) return '';

    // ═══ UTF-16LE 编码检测 ═══
    var sampleLen = Math.min(data.length, 20);
    var nullCount = 0;
    for (var si = 1; si < sampleLen; si += 2) {
        if (data[si] === 0) nullCount++;
    }
    // 如果采样中有足够多的 0x00 间隔 → UTF-16LE 编码
    if (sampleLen >= 4 && nullCount >= Math.floor(sampleLen / 4)) {
        return _decodeUtf16Atom(data);
    }

    // ═══ 回退：单字节/ANSI 解码 ═══
    var chars = [], hasPrintable = false;
    for (var i = 0; i < data.length; i++) {
        var b = data[i];
        if (b === 0) break;
        if (b >= 32 && b < 127) { chars.push(String.fromCharCode(b)); hasPrintable = true; }
        else if (b >= 0x80) hasPrintable = true;
    }
    var r = chars.join('').trim();
    return (hasPrintable && r.length >= 1) ? r : '';
}

// 解码 CString (type 0x0FAD)
function _decodeCString(data) {
    if (data.length >= 4) {
        var u = _decodeUtf16Atom(data);
        if (u.length >= 2) return u;
    }
    return _decodeTextBytesAtom(data);
}

// ================================================================
// ================================================================
//  PPT 元数据过滤器（v7 — 轻量级，无打分机制！）
// ================================================================
// ★ 设计哲学：结构化解析已经消除了95%的垃圾
//   剩余的少量"泄漏"都是 PowerPoint 自带的模板文字
//   不需要复杂的打分系统，只需要一个简单的黑名单匹配
//   → 准确、快速、可维护
function _pptMetadataFilter(texts) {
    var result = [], seen = {};

    // 模板文字黑名单（PowerPoint 自动生成的占位符/提示文字）
    var TEMPLATE_PATTERNS = [
        // 模板编辑提示
        /^(单击|双击)\s*(此处|这里)\s*(编辑|添加|插入|键入)/,
        /(编辑|添加|插入)\s*(母版|幻灯片|文本框|标题|副标题|正文)/,
        /^Click (here|to )/i,
        // 层级名称
        /^(第一级|第二级|第三级|第四级|第五级|第六级)$/,
        /^(Level \d+|Heading \d+|Title|Subtitle|Body Text)$/i,
        // 占位符标识（带数字ID）
        /(占位符|文本|标题|日期|页脚|对象|内容|形状|图片|表格|备注|批注)\s+\d{3,}/,
        /[一-鿿]+\d{3,}$/,
        // 字体名（常见字体家族前缀）
        /^(Arial|Times New Roman|Calibri|Verdana|Tahoma|Segoe|Georgia|Consolas|SimSun|SimHei|Microsoft YaHei|KaiTi|FangSong|Wingdings|Songti)(\s+\w+)*$/,
        // ★ PPT 内部 ID 标识符（混合大小写字母+数字+下划线/标点，无空格，长度≥6，排除纯数字）
        // 匹配如 USQdkYkHrh7h_、USQdkYkHre,g7h_
        // 排除纯数字（用户可能输入纯数字作为内容）
        /^(?!\d+$)[A-Za-z][A-Za-z0-9_\-,.:;~!@#$%^&*()+=\[\]{}<>?\\/]{4,}$/,
        // ★ 单符号/单字符标记（项目符号等，非有意义内容）
        /^[^a-zA-Z\u4e00-\u9fff0-9]$/,
        // 软件/产品标识
        /^(WPS ?Office|Microsoft ?Office|KSOP?ProductBuildVer|ICV|PID_GUID|MSPowerPoint|MS ?PPT|KSO)$/i,
        // GUID
        /^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$/i,
        /^[0-9A-F]{16,}$/i,
        // OLE2 内部名称
        /^(Root Entry|SummaryInformation|DocumentSummaryInformation|Current User|Pictures|CompObj)$/i,
    ];

    for (var ti = 0; ti < texts.length; ti++) {
        var txt = texts[ti].trim();
        if (!txt || txt.length < 1 || seen[txt]) continue;
        seen[txt] = true;

        var isMeta = false;
        for (var pi = 0; pi < TEMPLATE_PATTERNS.length && !isMeta; pi++) {
            if (TEMPLATE_PATTERNS[pi].test(txt)) isMeta = true;
        }

        if (!isMeta) {
            result.push(txt);
        } else {
            console.log('[DFV-PPT] 过滤模板文字: ' + txt.substring(0, 40));
        }
    }

    return result;
}

// ============ 幻灯片分段与渲染 ============

function splitPptTextIntoSlides(text) {
    var parts = text.split(/\n\s*\n|\n{3,}/);
    var slides = [];

    for (var i = 0; i < parts.length; i++) {
        var lines = parts[i].split('\n').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
        if (lines.length > 0) {
            slides.push({ index: i + 1, paragraphs: lines, hasContent: true });
        }
    }

    if (slides.length === 0 && text.trim()) {
        slides.push({ index: 1, paragraphs: [text.trim()], hasContent: true });
    }

    return slides;
}

function renderPptSlidesHtml(slides) {
    var html = '';
    html += '<style type="text/css">' +
        '.dbfv-ppt-container { width:100%; padding:16px 20px; box-sizing:border-box; }' +
        '.dbfv-ppt-slide { margin-bottom:24px; padding:16px 20px; border-left:3px solid var(--v6-primary); background:var(--background-secondary); border-radius:0 8px 8px 0; page-break-inside:avoid; }' +
        '.dbfv-ppt-slide-header { display:flex; align-items:center; gap:10px; margin-bottom:12px; padding-bottom:8px; border-bottom:1px solid var(--background-modifier-border); }' +
        '.dbfv-ppt-slide-num { display:inline-flex; align-items:center; justify-content:center; min-width:28px; height:28px; background:var(--v6-primary); color:#fff; border-radius:50%; font-size:12px; font-weight:600; flex-shrink:0; }' +
        '.dbfv-ppt-slide-label { font-size:12px; color:var(--text-muted); font-weight:500; }' +
        '.dbfv-ppt-para { margin:6px 0; padding:4px 0; font-size:13.5px; line-height:1.8; color:var(--text-normal); text-indent:2em; }' +
        '.dbfv-ppt-para:first-of-type { text-indent:0; }' +
        '</style>';

    html += '<div class="dbfv-ppt-container">';

    for (var i = 0; i < slides.length; i++) {
        var slide = slides[i];
        html += '<div class="dbfv-ppt-slide">';
        html += '<div class="dbfv-ppt-slide-header">';
        html += '<span class="dbfv-ppt-slide-num">' + slide.index + '</span>';
        html += '<span class="dbfv-ppt-slide-label">Slide ' + slide.index + (slide.hasContent ? '' : ' \u2014 \u7A7A\u767D\u9875') + '</span>';
        html += '</div>';
        if (slide.paragraphs.length > 0) {
            for (var j = 0; j < slide.paragraphs.length; j++) {
                html += '<p class="dbfv-ppt-para">' + escapeHtml(slide.paragraphs[j]) + '</p>';
            }
        } else {
            html += '<p style="font-style:italic;color:var(--text-faint);font-size:12px;padding:8px 0;">（\u6B64\u5E7B\u706F\u7247\u65E0\u53EF\u63D0\u53D6\u7684\u6587\u5B57\u5185\u5BB9）</p>';
        }
        html += '</div>';
    }

    html += '</div>';
    return html;
}

// ============ 统一入口：_renderLegacyOffice ============
async function _renderLegacyOffice(area, file, ext, zoomCtx, vault) {
    area.style.cssText += 'background:var(--background-primary);overflow:hidden;';
    var wrapEl = area.createEl('div', { cls: 'dbfv-ppt-content',
        attr: { style: 'width:100%;min-width:100%;box-sizing:border-box;height:100%;overflow:auto;padding:16px 12px;font-size:13px;line-height:1.8;color:var(--text-normal);user-select:text;-webkit-user-select:text;' }
    });

    if (zoomCtx) zoomCtx.setTarget(wrapEl);

    var isPptx = (ext === 'pptx');
    var formatLabel = isPptx ? 'PowerPoint (.pptx)' : 'PowerPoint 97-2003 (.ppt)';
    wrapEl.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:40px 0;">\u23F3 \u6B63\u5728\u89E3\u6790 ' + formatLabel + ' ...</div>';

    try {
        var data = await vault.readBinary(file);
        var arrayBuffer;
        if (data instanceof ArrayBuffer) arrayBuffer = data;
        else if (data instanceof Uint8Array) arrayBuffer = data.buffer;
        else if (data && data.buffer) arrayBuffer = data.buffer;
        else arrayBuffer = new Uint8Array(data).buffer;

        console.log('[DFV-PPT] 开始解析', formatLabel, ', 大小:', arrayBuffer.byteLength, 'bytes');

        var slides;

        if (isPptx) {
            slides = await _extractPptxText(arrayBuffer);
            wrapEl.innerHTML = renderPptxSlidesHtml(slides);
        } else {
            slides = await _extractPptText(arrayBuffer);
            wrapEl.innerHTML = renderPptSlidesHtml(slides);
        }

    } catch(e) {
        var errMsg = String(e.message || e).substring(0, 200);
        console.error('[DFV-PPT] 错误:', e);

        wrapEl.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;flex-direction:column;gap:12px;background:var(--background-secondary);height:100%;padding:40px;">' +
            '<div style="text-align:center;color:var(--text-muted);max-width:420px;">' +
            '<div style="font-size:48px;margin-bottom:12px;">\uD83D\uDCCB</div>' +
            '<p style="font-size:15px;margin-bottom:8px;font-weight:600;color:var(--text-error);">\u26A0 \u89E3\u6790\u5931\u8D25</p>' +
            '<p style="font-size:12px;margin-bottom:4px;">' + formatLabel + '</p>' +
            '<p style="font-size:11px;color:var(--text-accent);margin-bottom:16px;">' + escapeHtml(errMsg) + '</p>' +
            '</div></div>';

        var btnArea = document.createElement('div');
        btnArea.style.cssText = 'text-align:center;';
        wrapEl.appendChild(btnArea);
        var openBtn = document.createElement('button');
        openBtn.textContent = '\uD83D\uDCC2 用默认程序打开';
        openBtn.setAttribute('style', 'padding:8px 20px;font-size:13px;border:none;border-radius:6px;background:var(--v6-primary);color:white;cursor:pointer;');
        btnArea.appendChild(openBtn);

        var self = this;
        openBtn.addEventListener('click', function() {
            try {
                var fp = self.app.vault.adapter.getFullPath(file.path);
                require('electron').shell.openPath(fp);
            } catch(ex) { new Notice('打开失败: ' + ex.message); }
        });
    }
}

// ============ 注册到 FileViewer 全局处理器表 ============
(function() {
    var exts = ["ppt", "pptx"];
    exts.forEach(function(ext) {
        FILE_VIEWER_HANDLERS[ext] = _renderLegacyOffice;
    });
})();

// ============ 导出共享函数（供 ppt-viewer.js Dashboard 模块复用）============
window.__pptExtractor = {
    extractPptText: _extractPptText,
    extractPptxText: _extractPptxText,
    splitIntoSlides: splitPptTextIntoSlides,
    renderPptHtml: renderPptSlidesHtml,
    renderPptxHtml: renderPptxSlidesHtml,
    extractPptxXml: extractTextFromPptxXml,
    metadataFilter: _pptMetadataFilter
};

// ============ FileViewer 扩展：纯文本 / 代码 / 未知类型兜底 ============

    // 文本/代码
    async function _renderCode(area, file, ext, zoomCtx, vault) {
        var code = await vault.read(file);
        var pre = area.createEl('pre', { cls: 'dbfv-code',
            attr: { style: 'margin:0;padding:16px;font-family:var(--font-monospace);font-size:12px;line-height:1.6;white-space:pre-wrap;word-break:break-all;overflow-wrap:break-word;color:var(--text-normal);min-height:100%;' }
        });
        pre.textContent = code;
    }

    // 未知类型 → 尝试当文本读
    async function _renderUnknown(area, file, vault) {
        try {
            var raw = await vault.read(file);
            var pre = area.createEl('pre', { cls: 'dbfv-code',
                attr: { style: 'margin:0;padding:16px;font-family:var(--font-monospace);font-size:12px;line-height:1.6;white-space:pre-wrap;word-break:break-all;color:var(--text-muted);min-height:100%;' }
            });
            pre.textContent = raw;
        } catch(e) {
            showFileViewerErr(area, '⚠ 无法读取文件: ' + e.message);
        }
    }

(function() {
    // 代码/文本
    var codeExts = ["txt", "json", "yaml", "yml", "xml", "toml", "ini", "cfg", "env",
        "js", "ts", "jsx", "tsx", "css", "scss", "less",
        "py", "rb", "java", "go", "rs", "c", "cpp", "cs", "sh", "bat", "ps1"];
    codeExts.forEach(function(ext) {
        FILE_VIEWER_HANDLERS[ext] = _renderCode;
    });
    // ★ 未知类型作为兜底（最后注册，不覆盖已有处理器）
    DashboardFileViewer.prototype._fallbackHandler = _renderUnknown;
})();

// ============ FileViewer 辅助函数 ============

// ============ 辅助函数 ============

// 在内容区显示错误
function showFileViewerErr(area, text) {
    area.style.cssText += 'display:flex;align-items:center;justify-content:center;';
    area.createEl('div', { cls: 'dbfv-error', text: text,
        attr: { style: 'color:var(--text-error);text-align:center;padding:20px;' }
    });
}

// HTML 转义
function escapeHtml(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ============ FileViewer 扩展：视频 / 音频 + PDF ============

    // 视频 / 音频：HTML5 <video> / <audio> 标签
    function _renderMedia(area, file, ext, zoomCtx, vault) {
        var videoExts = ['mp4', 'webm', 'mov', 'avi', 'mkv', 'm4v'];
        var isVideo = videoExts.indexOf(ext) !== -1;
        var resourcePath = vault.getResourcePath(file);

        area.style.cssText += 'display:flex;align-items:center;justify-content:center;background:#000;';
        if (isVideo) {
            var vid = area.createEl('video', { cls: 'dbfv-video',
                attr: { src: resourcePath, controls: '', style: 'max-width:100%;max-height:100%;' }
            });
            // 防止 MP4 加载失败后无提示
            vid.addEventListener('error', function() {
                area.style.cssText += 'background:var(--background-secondary);';
                area.empty();
                showFileViewerErr(area, '⚠ 无法播放视频: ' + file.name + '\n\n编码格式可能不受浏览器支持');
            });
        } else {
            area.style.cssText += 'background:var(--background-primary);';
            area.createEl('audio', { cls: 'dbfv-audio',
                attr: { src: resourcePath, controls: '', style: 'width:80%;max-width:500px;' }
            });
            area.createEl('div', { text: '🎵 ' + file.name,
                attr: { style: 'position:absolute;bottom:20px;color:var(--text-muted);font-size:12px;' }
            });
        }
    }

    // PDF
    function _renderPdf(area, file, ext, zoomCtx, vault) {
        area.createEl('iframe', { cls: 'dbfv-pdf',
            attr: { src: vault.getResourcePath(file), style: 'width:100%;height:100%;border:none;' }
        });
    }

(function() {
    // 视频/音频
    var mediaExts = ["mp4", "webm", "mov", "avi", "mkv", "m4v", "mp3", "wav", "ogg", "flac", "m4a", "aac"];
    mediaExts.forEach(function(ext) {
        FILE_VIEWER_HANDLERS[ext] = _renderMedia;
    });
    // PDF
    FILE_VIEWER_HANDLERS["pdf"] = _renderPdf;
})();

// ============ FileViewer 扩展：表格文件（XLSX/XLS/CSV/ODS）============
// ★ 通过 FILE_VIEWER_HANDLERS 注册表注入
// ★ 依赖：libs/xlsx.full.min.js（SheetJS 运行时）

    // 表格文件：SheetJS 渲染 xlsx/xls/csv
    async function _renderSpreadsheet(area, file, ext, zoomCtx) {
        // ★ 强制文本可选择：Obsidian 默认 user-select:none，必须所有层级覆盖
        area.style.cssText += 'background:var(--background-primary);position:relative;overflow:hidden;user-select:text;-webkit-user-select:text;display:flex;flex-direction:column;';
        var wrapEl = area.createEl('div', { cls: 'dbfv-table-wrap',
            attr: { style: 'width:100%;min-width:100%;box-sizing:border-box;flex:1;min-height:0;overflow:auto;user-select:text;-webkit-user-select:text;' }
        });
        // ★ 连接 zoom 控件
        if (zoomCtx) zoomCtx.setTarget(wrapEl);

        // CSV 直接解析（不需要 SheetJS）
        if (ext === 'csv') {
            var data = await this.app.vault.readBinary(file);
            var text = new TextDecoder('utf-8').decode(data);
            var rows = parseCSVInline(text);
            renderTableToEl(wrapEl, rows);
            return;
        }

        // 加载状态
        wrapEl.innerHTML = '<div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);color:var(--text-muted);font-size:13px;">⏳ 正在加载表格引擎...</div>';

        var XLSX = await getXLSXAsync();
        if (!XLSX) {
            wrapEl.innerHTML = '<div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);color:var(--text-error);text-align:center;">⚠ SheetJS 加载失败<br><small>请检查网络连接后刷新</small></div>';
            return;
        }

        try {
            var data = await this.app.vault.readBinary(file);
            var wb = XLSX.read(new Uint8Array(data), { type: 'array' });

            // === 工具栏：小数位数 + 共享渲染状态 ===
            var _currentSheetName = wb.SheetNames[0];
            var _reRenderCurrent = function() {
                var ws = wb.Sheets[_currentSheetName];
                var json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                wrapEl.innerHTML = '';
                renderTableToEl(wrapEl, json, _decimalPlaces);
            };

            var tb = area.createEl('div', { cls: 'dbfv-toolbar',
                attr: { style: 'display:flex;align-items:center;gap:6px;padding:3px 10px;flex-shrink:0;font-size:11px;user-select:none;border-bottom:1px solid var(--background-modifier-border);' }
            });
            tb.createEl('span', { text: t('fv.table.decimals') + ':', attr: { style: 'color:var(--text-muted);' } });
            var decDisp = tb.createEl('span', { text: _decimalPlaces === null ? t('fv.table.auto') : _decimalPlaces,
                attr: { style: 'min-width:20px;text-align:center;font-weight:600;color:var(--text-normal);' }
            });
            var btnMinus = tb.createEl('button', {
                text: '\u2212',
                attr: { style: 'width:20px;height:20px;padding:0;border:1px solid var(--background-modifier-border);border-radius:3px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:12px;line-height:18px;text-align:center;' }
            });
            var btnPlus = tb.createEl('button', {
                text: '+',
                attr: { style: 'width:20px;height:20px;padding:0;border:1px solid var(--background-modifier-border);border-radius:3px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:12px;line-height:18px;text-align:center;' }
            });
            var btnAuto = tb.createEl('button', {
                text: t('fv.table.auto'),
                attr: { style: 'height:20px;padding:0 7px;border:1px solid var(--background-modifier-border);border-radius:3px;background:' + (_decimalPlaces === null ? 'var(--background-modifier-hover)' : 'var(--background-secondary)') + ';color:' + (_decimalPlaces === null ? 'var(--text-accent)' : 'var(--text-normal)') + ';cursor:pointer;font-size:11px;line-height:18px;' }
            });

            btnMinus.addEventListener('click', function() {
                var cur = _decimalPlaces === null ? 2 : _decimalPlaces;
                if (cur > 0) { _setDecimalPlaces(cur - 1); decDisp.textContent = cur - 1; btnAuto.style.background = 'var(--background-secondary)'; btnAuto.style.color = 'var(--text-normal)'; _reRenderCurrent(); }
            });
            btnPlus.addEventListener('click', function() {
                var cur = _decimalPlaces === null ? 2 : _decimalPlaces;
                if (cur < 10) { _setDecimalPlaces(cur + 1); decDisp.textContent = cur + 1; btnAuto.style.background = 'var(--background-secondary)'; btnAuto.style.color = 'var(--text-normal)'; _reRenderCurrent(); }
            });
            btnAuto.addEventListener('click', function() {
                _setDecimalPlaces(null); decDisp.textContent = t('fv.table.auto'); btnAuto.style.background = 'var(--background-modifier-hover)'; btnAuto.style.color = 'var(--text-accent)'; _reRenderCurrent();
            });

            // 多 Sheet 标签栏
            if (wb.SheetNames && wb.SheetNames.length > 1) {
                var sheetBar = area.createEl('div', { cls: 'dbfv-sheets',
                    attr: { style: 'display:flex;gap:3px;padding:4px 8px;flex-shrink:0;flex-wrap:wrap;' }
                });
                var self = this;
                wb.SheetNames.forEach(function(name, idx) {
                    var tab = sheetBar.createEl('span', { cls: 'dbfv-sheet-tab',
                        text: name,
                        attr: { style: 'padding:2px 10px;border:1px solid var(--background-modifier-border);border-radius:4px 4px 0 0;font-size:11px;cursor:pointer;background:' + (idx === 0 ? 'var(--background-modifier-form-field)' : 'var(--background-secondary)') + ';color:' + (idx === 0 ? 'var(--text-normal)' : 'var(--text-muted)') + ';user-select:none;' }
                    });
                    tab.addEventListener('click', function() {
                        var tabs = sheetBar.querySelectorAll('.dbfv-sheet-tab');
                        tabs.forEach(function(t) { t.style.background = 'var(--background-secondary)'; t.style.color = 'var(--text-muted)'; });
                        tab.style.background = 'var(--background-modifier-form-field)';
                        tab.style.color = 'var(--text-normal)';
                        _currentSheetName = name;
                        _reRenderCurrent();
                    });
                });
            }

            // 默认显示第一个 Sheet
            _reRenderCurrent();

        } catch(e) {
            wrapEl.innerHTML = '<div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);color:var(--text-error);text-align:center;">⚠ 解析表格失败: ' + escapeHtml(e.message) + '</div>';
            console.error('[DashboardFileViewer] 表格渲染失败:', e);
        }
    }

// 小数位数设置（持久化到 localStorage，null=自动不格式化）
var _decimalPlaces = (function() {
    try { var v = localStorage.getItem('dbfv-decimal-places'); if (v === null || v === 'auto') return null; return Math.max(0, Math.min(10, parseInt(v) || 0)); }
    catch(e) { return null; }
})();
function _setDecimalPlaces(v) {
    _decimalPlaces = v;
    try { localStorage.setItem('dbfv-decimal-places', v === null ? 'auto' : String(v)); } catch(e) {}
}

// CSV 简易解析
function parseCSVInline(text) {
    var rows = [];
    var lines = text.split(/\r?\n/).filter(function(l) { return l.trim(); });
    for (var i = 0; i < lines.length; i++) {
        var cols = [];
        var inQuote = false, col = '';
        for (var j = 0; j < lines[i].length; j++) {
            var ch = lines[i][j];
            if (ch === '"' && !inQuote) { inQuote = true; }
            else if (ch === '"' && inQuote) { inQuote = false; }
            else if (ch === ',' && !inQuote) { cols.push(col.trim()); col = ''; }
            else { col += ch; }
        }
        cols.push(col.trim());
        rows.push(cols);
    }
    return rows;
}

var _tableSelectState = null;

function startColResize(e, colIndex, table) {
    e.preventDefault(); e.stopPropagation();
    var startX = e.clientX;
    var colgroup = table.querySelector('colgroup');
    var col = colgroup ? colgroup.children[colIndex] : null;
    if (!col) return;
    var startWidth = parseInt(col.style.width || '120');

    function onMove(ev) {
        var delta = ev.clientX - startX;
        var newWidth = Math.max(40, startWidth + delta);
        col.style.width = newWidth + 'px';
    }
    function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
    }
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
}

function clearTableSelection() {
    if (_tableSelectState && _tableSelectState.table) {
        var cells = _tableSelectState.table.querySelectorAll('td.cell-selected, th.cell-selected');
        for (var i = 0; i < cells.length; i++) { cells[i].classList.remove('cell-selected'); }
    }
    _tableSelectState = null;
}

function startCellSelect(e, row, col, table) {
    // ★ 只有按住 Alt 键时才启用单元格选区，否则让原生文本选择正常工作
    if (!e.altKey) {
        clearTableSelection();
        return;
    }
    e.preventDefault();  // Alt+拖拽时阻止原生选区，启用单元格选区

    // Shift+Alt+Click: 扩展选区
    if (e.shiftKey && _tableSelectState && _tableSelectState.table === table) {
        _tableSelectState.endRow = row;
        _tableSelectState.endCol = col;
        highlightSelection(table);
        return;
    }
    // Alt+点击: 新选区
    clearTableSelection();
    _tableSelectState = { table: table, startRow: row, startCol: col, endRow: row, endCol: col };
    highlightSelection(table);

    function onMove(ev) {
        ev.preventDefault();
        // 找鼠标所在单元格
        var target = document.elementFromPoint(ev.clientX, ev.clientY);
        while (target && target !== table) {
            if (target.tagName === 'TD' || target.tagName === 'TH') {
                var r = parseInt(target.getAttribute('data-row') || target.getAttribute('data-r'));
                var c = parseInt(target.getAttribute('data-col') || target.getAttribute('data-c'));
                if (!isNaN(r) && !isNaN(c) && _tableSelectState) {
                    _tableSelectState.endRow = r;
                    _tableSelectState.endCol = c;
                    highlightSelection(table);
                }
                break;
            }
            target = target.parentElement;
        }
    }
    function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
    }
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
}

function highlightSelection(table) {
    if (!_tableSelectState || _tableSelectState.table !== table) return;
    var s = _tableSelectState;
    var minRow = Math.min(s.startRow, s.endRow);
    var maxRow = Math.max(s.startRow, s.endRow);
    var minCol = Math.min(s.startCol, s.endCol);
    var maxCol = Math.max(s.startCol, s.endCol);

    // 清除旧选区
    var cells = table.querySelectorAll('td.cell-selected, th.cell-selected');
    for (var i = 0; i < cells.length; i++) { cells[i].classList.remove('cell-selected'); }

    // 高亮新选区
    for (var r = minRow; r <= maxRow; r++) {
        for (var c = minCol; c <= maxCol; c++) {
            var cell = table.querySelector('td[data-row="' + r + '"][data-col="' + c + '"], th[data-col="' + c + '"]');
            if (r === 0 && cell) {
                // header row: use th
                var th = table.querySelector('th[data-col="' + c + '"]');
                if (th) th.classList.add('cell-selected');
            } else if (cell) {
                cell.classList.add('cell-selected');
            }
        }
    }
}

function copyCellSelection() {
    if (!_tableSelectState || !_tableSelectState.table) return;
    var s = _tableSelectState;
    var minRow = Math.min(s.startRow, s.endRow);
    var maxRow = Math.max(s.startRow, s.endRow);
    var minCol = Math.min(s.startCol, s.endCol);
    var maxCol = Math.max(s.startCol, s.endCol);
    var table = s.table;

    var lines = [];
    for (var r = minRow; r <= maxRow; r++) {
        var line = [];
        for (var c = minCol; c <= maxCol; c++) {
            var cell;
            if (r === 0) {
                cell = table.querySelector('th[data-col="' + c + '"]');
            } else {
                cell = table.querySelector('td[data-row="' + r + '"][data-col="' + c + '"]');
            }
            var txt = cell ? (cell.textContent || '').trim() : '';
            // 内含 tab/换行的值用引号包裹
            if (txt.indexOf('\t') !== -1 || txt.indexOf('\n') !== -1) {
                txt = '"' + txt.replace(/"/g, '""') + '"';
            }
            line.push(txt);
        }
        lines.push(line.join('\t'));
    }

    var tsv = lines.join('\n');
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(tsv).catch(function() { /* 静默 */ });
    }
    // 反馈：短暂闪烁状态栏
    if (window.Notice) {
        new Notice('已复制 ' + lines.length + ' 行 × ' + (maxCol - minCol + 1) + ' 列');
    }
}

// ★ 全局 Ctrl+C / Esc 钩子（文档级）
document.addEventListener('keydown', function(e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 'c' && _tableSelectState && _tableSelectState.table) {
        // 确保不在 input/textarea 中
        var tag = document.activeElement ? document.activeElement.tagName : '';
        if (tag !== 'INPUT' && tag !== 'TEXTAREA') {
            // ★ 关键：如果用户有原生文本选区，优先走原生复制，不拦截
            var nativeSel = window.getSelection ? window.getSelection().toString() : '';
            if (nativeSel && nativeSel.length > 0) {
                return;  // 让原生 Ctrl+C 正常工作
            }
            // 只有没有原生文本选区时，才用单元格选区复制
            copyCellSelection();
            e.preventDefault();
        }
    }
    // Esc 清除选区
    if (e.key === 'Escape' && _tableSelectState) {
        clearTableSelection();
    }
});

// ★ 注入选择器样式 + 全局内容区文本选择保障
(function injectTableSelectStyle() {
    if (document.getElementById('dbfv-cell-select-style')) return;
    var styleEl = document.createElement('style');
    styleEl.id = 'dbfv-cell-select-style';
    styleEl.textContent =
        // 单元格选区高亮（仅 Alt+拖拽 时触发）
        '.cell-selected { background: rgba(66,133,244,0.15) !important; outline: 2px solid #4285f4 !important; outline-offset: -2px; z-index:1; position:relative; }' +
        // ★ 关键：确保 FileViewer 所有内容区文本可选择（覆盖 Obsidian 默认的 user-select:none）
        //   多层覆盖：container → content → table-wrap → table/cell
        '.dbfv-container, .dbfv-container *, .dbfv-content, .dbfv-content *,' +
        '.dbfv-docx-content, .dbfv-docx-content *,' +
        '.dbfv-doc-content, .dbfv-doc-content *,' +
        '.dbfv-table-wrap, .dbfv-table-wrap * { user-select:text !important; -webkit-user-select:text !important; }' +
        // 选中文本高亮（蓝底白字，在所有主题下可见）
        '.dbfv-container ::selection, .dbfv-content ::selection,' +
        '.dbfv-docx-content ::selection, .dbfv-doc-content ::selection,' +
        '.dbfv-table-wrap ::selection { background: rgba(66,133,244,0.45) !important; color:inherit; }' +
        '.dbfv-container ::-moz-selection, .dbfv-content ::-moz-selection,' +
        '.dbfv-docx-content ::-moz-selection, .dbfv-doc-content ::-moz-selection,' +
        '.dbfv-table-wrap ::-moz-selection { background: rgba(66,133,244,0.45) !important; color:inherit; }';
    document.head.appendChild(styleEl);
})();

// 渲染二维数组为增强型 HTML 表格
function renderTableToEl(el, rows, decimalPlaces) {
    if (!rows || rows.length === 0) {
        el.innerHTML = '<div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);color:var(--text-muted);font-size:13px;">📊 表格为空</div>';
        return;
    }

    var maxCols = 0;
    for (var i = 0; i < rows.length; i++) {
        if (rows[i] && rows[i].length > maxCols) maxCols = rows[i].length;
    }

    // 清除旧选区
    clearTableSelection();

    // --- 构建 colgroup ---
    var table = document.createElement('table');
    table.style.cssText = 'border-collapse:collapse;font-size:11px;table-layout:auto;box-sizing:border-box;user-select:text;-webkit-user-select:text;';

    var colgroup = document.createElement('colgroup');
    for (var c = 0; c < maxCols; c++) {
        var colEl = document.createElement('col');
        colEl.style.cssText = 'min-width:80px;max-width:360px;';
        colgroup.appendChild(colEl);
    }
    table.appendChild(colgroup);

    // --- 构建 thead / tbody ---
    var thead = document.createElement('thead');
    var tbody = document.createElement('tbody');
    table.appendChild(thead);
    table.appendChild(tbody);

    for (var r = 0; r < rows.length; r++) {
        var tr = document.createElement('tr');
        var row = rows[r] || [];
        var isHeader = (r === 0);

        if (isHeader) { thead.appendChild(tr); }
        else { tbody.appendChild(tr); }

        for (var cIdx = 0; cIdx < maxCols; cIdx++) {
            var val = row[cIdx];
            if (val === undefined || val === null) val = '';
            var cellText;
            if (typeof val === 'number' && decimalPlaces !== null && decimalPlaces !== undefined && decimalPlaces >= 0) {
                cellText = val.toFixed(decimalPlaces);
            } else {
                cellText = String(val);
            }

            var cell;
            if (isHeader) {
                cell = document.createElement('th');
                cell.setAttribute('data-col', cIdx);
                cell.style.cssText = 'border:1px solid var(--background-modifier-border);padding:4px 8px;position:relative;background:var(--background-modifier-form-field);font-weight:600;overflow:hidden;white-space:pre-wrap;word-break:break-word;box-sizing:border-box;';
                cell.textContent = cellText;

                // 列宽拖拽手柄
                var handle = document.createElement('div');
                handle.style.cssText = 'position:absolute;right:0;top:0;width:6px;height:100%;cursor:col-resize;z-index:5;background:transparent;transition:background 0.15s;';
                handle.addEventListener('mouseenter', function() { this.style.background = 'var(--v6-primary)'; });
                handle.addEventListener('mouseleave', function() { this.style.background = 'transparent'; });
                (function(colIdx) {
                    handle.addEventListener('mousedown', function(ev) { startColResize(ev, colIdx, table); });
                })(cIdx);
                cell.appendChild(handle);
            } else {
                cell = document.createElement('td');
                cell.setAttribute('data-row', r);
                cell.setAttribute('data-col', cIdx);
                cell.style.cssText = 'border:1px solid var(--background-modifier-border);padding:4px 8px;overflow:hidden;white-space:pre-wrap;word-break:break-word;box-sizing:border-box;';
                cell.textContent = cellText;
            }

            // 选区事件（所有数据格 + 表头）
            cell.setAttribute('data-r', r);
            cell.setAttribute('data-c', cIdx);
            cell.addEventListener('mousedown', (function(rowIdx, colIdx) {
                return function(ev) { startCellSelect(ev, rowIdx, colIdx, table); };
            })(r, cIdx));

            tr.appendChild(cell);
        }
    }

    el.innerHTML = '';
    el.appendChild(table);
}


// ============ 注册到 FileViewer ============
(function() {
    var exts = ["xlsx", "xls", "xlsm", "csv", "ods"];
    exts.forEach(function(ext) {
        FILE_VIEWER_HANDLERS[ext] = _renderSpreadsheet;
    });
})();

// ===================== 内置模块（构建脚本自动注入）=====================
// 发布模式：模块代码内联，无需文件系统读取
const BUILTIN_MODULES = {
  "ai-insight": "/**\n * AI洞察模块 V15\n * 格式：V14（含 id/styles/renderSettings/defaultSettings）\n * 功能：V11 完整版（分析最近5篇笔记 + 调用 AI API + 格式化显示 + 当天缓存）\n * 新增：全局请求节流 + 实例独立缓存 + 可配置请求延迟\n */\nconst id = 'ai-insight';\nconst title = t('mod.aiInsight');\nconst icon = '💡';\n\nconst defaultSettings = {\n    apiKey: '',\n    apiUrl: 'https://api.openai.com/v1/chat/completions',\n    model: 'gpt-3.5-turbo',\n    temperature: 0.7,\n    requestDelay: 0\n};\n\nconst styles = `/* AI洞察模块样式已在 styles.css 中定义 */`;\n\n// 全局 AI 请求节流器（跨实例共享，避免同时触发多个 AI 请求）\nif (!window._v15AIThrottle) {\n    window._v15AIThrottle = {\n        lastRequestTime: 0,\n        minInterval: 2000, // 默认最小间隔 2 秒\n        async waitForTurn(extraDelayMs = 0) {\n            const now = Date.now();\n            const nextAvailable = this.lastRequestTime + this.minInterval;\n            const waitTime = Math.max(0, nextAvailable - now) + extraDelayMs;\n            if (waitTime > 0) {\n                await new Promise(r => setTimeout(r, waitTime));\n            }\n            this.lastRequestTime = Date.now();\n        }\n    };\n}\n\n// 实例级缓存（以 settings 对象为 key，确保每个实例独立缓存）\nif (!window._v15AICaches) {\n    window._v15AICaches = new Map();\n}\n\nfunction getInstanceCache() {\n    let state = window._v15AICaches.get(settings);\n    if (!state) {\n        state = { lastDate: null, analysisResult: null };\n        window._v15AICaches.set(settings, state);\n    }\n    return state;\n}\n\nasync function getRecentNotes(limit = 5) {\n    const files = app.vault.getMarkdownFiles()\n        .sort((a, b) => b.stat.mtime - a.stat.mtime)\n        .slice(0, limit);\n\n    const notes = [];\n    for (const file of files) {\n        try {\n            const content = await app.vault.read(file);\n            const cleanContent = content\n                .replace(/^---[\\s\\S]*?---\\n?/, '')\n                .replace(/```[\\s\\S]*?```/g, '')\n                .trim();\n            notes.push({\n                title: file.basename,\n                content: cleanContent.substring(0, 600),\n                path: file.path\n            });\n        } catch (e) { /* ignore */ }\n    }\n    return notes;\n}\n\nasync function analyzeWithAI(notes) {\n    const apiKey = settings.apiKey || '';\n    const apiModel = settings.model || 'gpt-3.5-turbo';\n    const temperature = settings.temperature || 0.7;\n\n    let apiUrl = settings.apiUrl || 'https://api.openai.com/v1/chat/completions';\n    if (apiUrl && !apiUrl.includes('/v1/') && !apiUrl.includes('/chat')) {\n        apiUrl = apiUrl.replace(/\\/$/, '') + '/v1/chat/completions';\n    }\n\n    if (!apiKey) throw new Error(t('mod.ai.error.noKey'));\n\n    const prompt = `请分析以下笔记内容，提供：\n1. 主题总结（2-3句话）\n2. 关键知识点提取（3-5个）\n3. 建议的关联方向或行动\n\n笔记内容：\n${notes.map((n, i) => `${i + 1}. 《${n.title}》\\n${n.content}`).join('\\n\\n')}`;\n\n    try {\n        const response = await requestUrl({\n            url: apiUrl,\n            method: 'POST',\n            headers: {\n                'Content-Type': 'application/json',\n                'Authorization': 'Bearer ' + apiKey\n            },\n            body: JSON.stringify({\n                model: apiModel,\n                messages: [{ role: 'user', content: prompt }],\n                temperature: parseFloat(temperature)\n            })\n        });\n\n        let data = response;\n        if (response.text) {\n            try { data = JSON.parse(response.text); } catch (e) { return response.text; }\n        }\n        if (typeof data === 'object' && data.json) data = data.json;\n\n        if (data.choices?.[0]?.message?.content) return data.choices[0].message.content;\n        if (data.content) return data.content;\n        if (data.response) return data.response;\n        if (data.text) return data.text;\n        if (data.result) return data.result;\n        if (data.error) throw new Error(data.error.message || t('mod.ai.error.apiError'));\n\n        throw new Error(t('mod.ai.error.parseFailed'));\n    } catch (e) {\n        if (e.message.includes('401')) throw new Error(t('mod.ai.error.invalidKey'));\n        if (e.message.includes('404')) throw new Error(t('mod.ai.error.invalidUrl'));\n        if (e.message.includes('429')) throw new Error(t('mod.ai.error.rateLimited'));\n        throw new Error(t('mod.ai.error.callFailed') + e.message);\n    }\n}\n\nfunction displayContent(resultArea, text) {\n    resultArea.empty();\n    const lines = text.split('\\n').filter(l => l.trim());\n    lines.forEach(line => {\n        if (line.startsWith('###')) {\n            resultArea.createEl('h4', {\n                text: line.replace(/^###\\s*/, ''),\n                attr: { style: 'margin: 10px 0 5px; font-size: 13px; color: var(--v6-primary);' }\n            });\n        } else if (line.startsWith('##')) {\n            resultArea.createEl('h3', {\n                text: line.replace(/^##\\s*/, ''),\n                attr: { style: 'margin: 12px 0 6px; font-size: 14px; color: var(--v6-primary);' }\n            });\n        } else if (line.startsWith('- ') || line.startsWith('* ')) {\n            resultArea.createEl('div', {\n                text: '• ' + line.substring(2),\n                attr: { style: 'margin: 5px 0; padding-left: 10px; font-size: 13px;' }\n            });\n        } else if (/^\\d+\\./.test(line)) {\n            resultArea.createEl('div', {\n                text: line,\n                attr: { style: 'margin: 5px 0; padding-left: 6px; font-size: 13px;' }\n            });\n        } else {\n            resultArea.createEl('p', {\n                text: line,\n                attr: { style: 'margin: 6px 0; font-size: 13px; line-height: 1.7;' }\n            });\n        }\n    });\n}\n\nasync function render(content) {\n    const state = getInstanceCache();\n    const today = moment().format('YYYY-MM-DD');\n\n    content.empty();\n    const container = content.createDiv({ cls: 'ai-insight-container' });\n\n    // 工具栏\n    const toolbar = container.createDiv({ cls: 'ai-insight-toolbar' });\n    const analyzeBtn = toolbar.createEl('button', { text: t('mod.ai.btn.analyze'), cls: 'ai-insight-btn' });\n    const clearBtn = toolbar.createEl('button', { text: t('mod.ai.btn.clearCache'), cls: 'ai-insight-btn secondary' });\n\n    // 结果区域\n    const resultArea = container.createDiv({ cls: 'ai-insight-response' });\n\n    // 时间戳\n    const dateEl = container.createDiv({ cls: 'ai-insight-date' });\n    if (state.lastDate) dateEl.textContent = `上次分析：${state.lastDate}`;\n\n    const doAnalyze = async () => {\n        resultArea.empty();\n        resultArea.createEl('div', {\n            cls: 'ai-insight-loading',\n            text: t('mod.ai.analyzing')\n        });\n        analyzeBtn.disabled = true;\n\n        try {\n            // 请求节流：等待轮到自己的回合\n            const extraDelay = (Number(settings.requestDelay) || 0) * 1000;\n            await window._v15AIThrottle.waitForTurn(extraDelay);\n\n            const notes = await getRecentNotes(5);\n            if (notes.length === 0) {\n                resultArea.empty();\n                resultArea.createEl('div', { cls: 'ai-insight-empty', text: t('mod.ai.empty') });\n                analyzeBtn.disabled = false;\n                return;\n            }\n\n            const result = await analyzeWithAI(notes);\n            state.analysisResult = result;\n            state.lastDate = today;\n            dateEl.textContent = `分析于：${today}`;\n            displayContent(resultArea, result);\n        } catch (e) {\n            resultArea.empty();\n            resultArea.createEl('div', {\n                cls: 'ai-insight-error',\n                text: e.message\n            });\n        } finally {\n            analyzeBtn.disabled = false;\n        }\n    };\n\n    analyzeBtn.addEventListener('click', doAnalyze);\n    clearBtn.addEventListener('click', () => {\n        state.analysisResult = null;\n        state.lastDate = null;\n        resultArea.empty();\n        resultArea.createEl('div', { cls: 'ai-insight-empty', text: t('mod.ai.cacheCleared') });\n        dateEl.textContent = '';\n    });\n\n    // 有缓存直接显示，无缓存自动触发分析\n    if (state.lastDate === today && state.analysisResult) {\n        displayContent(resultArea, state.analysisResult);\n        dateEl.textContent = `分析于：${today}`;\n    } else if (settings.apiKey) {\n        doAnalyze();\n    } else {\n        resultArea.createEl('div', {\n            cls: 'ai-insight-empty',\n            text: t('mod.ai.needKey')\n        });\n    }\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    const { Setting } = require('obsidian');\n\n    containerEl.createEl('h3', { text: t('mod.ai.settings.title') });\n\n    new Setting(containerEl)\n        .setName(t('mod.ai.settings.apiKey'))\n        .setDesc(t('mod.ai.settings.apiKeyDesc'))\n        .addText(t => {\n            t.setPlaceholder('sk-...')\n                .setValue(settings.apiKey || '')\n                .onChange(async (v) => {\n                    settings.apiKey = v.trim();\n                    await saveCallback();\n                });\n            t.inputEl.style.width = '100%';\n        });\n\n    new Setting(containerEl)\n        .setName(t('mod.ai.settings.apiUrl'))\n        .setDesc(t('mod.ai.settings.apiUrlDesc'))\n        .addText(t => {\n            t.setPlaceholder('https://api.openai.com/v1/chat/completions')\n                .setValue(settings.apiUrl || '')\n                .onChange(async (v) => {\n                    settings.apiUrl = v.trim();\n                    await saveCallback();\n                });\n            t.inputEl.style.width = '100%';\n        });\n\n    new Setting(containerEl)\n        .setName(t('mod.ai.settings.model'))\n        .setDesc(t('mod.ai.settings.modelDesc'))\n        .addDropdown(d => {\n            d.addOption('gpt-3.5-turbo', 'GPT-3.5 Turbo')\n                .addOption('gpt-4o-mini', 'GPT-4o Mini')\n                .addOption('gpt-4o', 'GPT-4o')\n                .addOption('deepseek-chat', 'DeepSeek Chat')\n                .addOption('moonshot-v1-8k', 'Moonshot v1-8k')\n                .addOption('custom', t('mod.ai.settings.customOption'));\n\n            const knownModels = ['gpt-3.5-turbo', 'gpt-4o-mini', 'gpt-4o', 'deepseek-chat', 'moonshot-v1-8k'];\n            const currentModel = settings.model || 'gpt-3.5-turbo';\n            d.setValue(knownModels.includes(currentModel) ? currentModel : 'custom')\n                .onChange(async (v) => {\n                    if (v !== 'custom') {\n                        settings.model = v;\n                        await saveCallback();\n                    }\n                });\n        })\n        .addText(t => {\n            t.setPlaceholder(t('mod.ai.settings.customModelPlaceholder'))\n                .setValue(['gpt-3.5-turbo', 'gpt-4o-mini', 'gpt-4o', 'deepseek-chat', 'moonshot-v1-8k'].includes(settings.model || 'gpt-3.5-turbo') ? '' : (settings.model || ''))\n                .onChange(async (v) => {\n                    if (v.trim()) {\n                        settings.model = v.trim();\n                        await saveCallback();\n                    }\n                });\n        });\n\n    new Setting(containerEl)\n        .setName(t('mod.ai.settings.temperature'))\n        .setDesc(t('mod.ai.settings.temperatureDesc'))\n        .addSlider(s => {\n            s.setLimits(0, 1, 0.1)\n                .setValue(settings.temperature || 0.7)\n                .setDynamicTooltip()\n                .onChange(async (v) => {\n                    settings.temperature = v;\n                    await saveCallback();\n                });\n        });\n\n    new Setting(containerEl)\n        .setName(t('mod.ai.settings.requestDelay'))\n        .setDesc(t('mod.ai.settings.requestDelayDesc'))\n        .addSlider(s => {\n            s.setLimits(0, 10, 0.5)\n                .setValue(Number(settings.requestDelay) || 0)\n                .setDynamicTooltip()\n                .onChange(async (v) => {\n                    settings.requestDelay = v;\n                    await saveCallback();\n                });\n        });\n\n    new Setting(containerEl)\n        .setName(t('mod.ai.settings.globalInterval'))\n        .setDesc(t('mod.ai.settings.globalIntervalDesc'))\n        .addText(t => {\n            t.setPlaceholder('2000')\n                .setValue(String(window._v15AIThrottle ? window._v15AIThrottle.minInterval : 2000))\n                .onChange(async (v) => {\n                    const val = parseInt(v);\n                    if (window._v15AIThrottle && isFinite(val) && val >= 0) {\n                        window._v15AIThrottle.minInterval = val;\n                    }\n                    await saveCallback();\n                });\n        });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "aquarium": "/**\n * 水族箱模块 V1\n * 基于 aquariumtank (MIT License) 改编适配\n * 功能：添加鱼/气泡/水草/珊瑚/投食/主题切换/夜间模式/LocalStorage持久化\n */\nconst id = 'aquarium';\nconst title = t('mod.aquarium');\nconst icon = '🐠';\n\nconst defaultSettings = {\n    theme: 'tropical',\n    sandColor: 'golden',\n    nightMode: false,\n    autoSave: true\n};\n\nconst styles = `\n.aquarium-wrap { height:100%; display:flex; flex-direction:column; overflow:hidden; border-radius:12px; }\n.aquarium-toolbar { display:flex; gap:4px; padding:6px 10px; flex-shrink:0; flex-wrap:wrap; background:rgba(0,0,0,0.2); }\n.aq-btn { padding:4px 10px; border:none; border-radius:14px; cursor:pointer; font-size:11px; font-weight:600; transition:all .2s; text-transform:uppercase; letter-spacing:.5px; color:#fff; }\n.aq-btn-fish { background:linear-gradient(135deg,#ff6b6b,#ff8e53); box-shadow:0 2px 8px rgba(255,107,107,.3); }\n.aq-btn-bubble { background:linear-gradient(135deg,#4facfe,#00f2fe); box-shadow:0 2px 8px rgba(79,172,254,.3); }\n.aq-btn-plant { background:linear-gradient(135deg,#38ef7d,#11998e); box-shadow:0 2px 8px rgba(56,239,125,.3); }\n.aq-btn-food { background:linear-gradient(135deg,#f093fb,#f5576c); box-shadow:0 2px 8px rgba(240,147,251,.3); }\n.aq-btn-decor { background:linear-gradient(135deg,#a18cd1,#fbc2eb); box-shadow:0 2px 8px rgba(161,140,209,.3); color:#333; }\n.aq-btn-clear { background:linear-gradient(135deg,#868f96,#596164); box-shadow:0 2px 8px rgba(134,143,150,.3); }\n.aq-btn-reset { background:linear-gradient(135deg,#fc4a1a,#f7b733); box-shadow:0 2px 8px rgba(252,74,26,.3); }\n.aq-btn:hover { transform:translateY(-1px); filter:brightness(1.1); }\n.aq-body { flex:1; position:relative; overflow:hidden; border-radius:0 0 12px 12px;\n    box-shadow:inset 0 0 60px rgba(255,255,255,.08), 0 0 0 3px #2d4a5e, 0 0 0 5px #1a3a4a; transition:background .8s ease; }\n/* Themes */\n.aq-theme-tropical { background:linear-gradient(180deg,#0077b6,#0096c7 20%,#00b4d8 40%,#48cae4 60%,#90e0ef 80%,#ade8f4 100%); }\n.aq-theme-ocean { background:linear-gradient(180deg,#003366,#004080 20%,#0059b3 40%,#0073e6 60%,#1a8cff 80%,#4da6ff 100%); }\n.aq-theme-pond { background:linear-gradient(180deg,#2d5a27,#3d7a37 20%,#4a9a47 40%,#5cb85c 60%,#7dcea0 80%,#a8e6cf 100%); }\n.aq-theme-sunset { background:linear-gradient(180deg,#2c1810,#4a2c1a 20%,#6b3d22 40%,#8b5230 60%,#b87a45 80%,#d4a56a 100%); }\n.aq-theme-midnight { background:linear-gradient(180deg,#0a0a1a,#0d1028 20%,#101838 40%,#152050 60%,#1a2868 80%,#203080 100%); }\n.aq-night { filter:brightness(.4); }\n.aq-night .aq-bio { filter:brightness(3) drop-shadow(0 0 8px currentColor); }\n/* Sand */\n.aq-sand { position:absolute; bottom:0; left:0; width:100%; height:60px; border-radius:0 0 10px 10px; transition:background .5s; }\n.aq-sand-golden { background:linear-gradient(180deg,#f4d03f,#e6b800 30%,#c9a227 60%,#b8860b 100%); }\n.aq-sand-white { background:linear-gradient(180deg,#f5f5f5,#e0e0e0 30%,#bdbdbd 60%,#9e9e9e 100%); }\n.aq-sand-dark { background:linear-gradient(180deg,#5d4e37,#4a3f2f 30%,#3d3427 60%,#2e2820 100%); }\n/* Light rays */\n.aq-rays { position:absolute; top:0; left:0; width:100%; height:100%; pointer-events:none; overflow:hidden; }\n.aq-ray { position:absolute; top:-30px; width:120px; height:350px; background:linear-gradient(180deg,rgba(255,255,255,.25) 0%,rgba(255,255,255,.08) 50%,transparent 100%); transform:skewX(-20deg); animation:aqRayShimmer 7s ease-in-out infinite; opacity:.3; }\n@keyframes aqRayShimmer { 0%,100%{opacity:.25} 50%{opacity:.5} }\n/* Water surface */\n.aq-surface { position:absolute; top:0; left:0; width:100%; height:24px; background:linear-gradient(180deg,rgba(255,255,255,.25) 0%,rgba(255,255,255,.08) 50%,transparent 100%); pointer-events:none; animation:aqSurface 3s ease-in-out infinite; }\n@keyframes aqSurface { 0%,100%{transform:scaleY(1)} 50%{transform:scaleY(1.08)} }\n/* Fish */\n.aq-fish { position:absolute; cursor:pointer; z-index:10; transition:filter .3s; }\n.aq-fish:hover { filter:brightness(1.3) drop-shadow(0 0 8px rgba(255,255,255,.6)); }\n.aq-fish-body { position:relative; width:48px; height:28px; border-radius:50% 40% 40% 50%;\n    box-shadow:inset -3px -3px 10px rgba(0,0,0,.15), inset 3px 3px 10px rgba(255,255,255,.25); }\n.aq-fish-eye { position:absolute; top:6px; width:9px; height:9px; background:radial-gradient(circle at 35% 35%,#fff,#333); border-radius:50%; }\n.aq-fish-eye::after { content:''; position:absolute; top:2px; left:2px; width:4px; height:4px; background:#000; border-radius:50%; }\n.aq-fish-eye.right { right:7px; } .aq-fish-eye.left { left:7px; }\n.aq-fish-tail { position:absolute; top:50%; transform:translateY(-50%); width:0; height:0; border-style:solid; animation:aqTailWag .3s ease-in-out infinite; }\n.aq-fish.right .aq-fish-tail { left:-14px; border-width:12px 16px 12px 0; border-color:transparent currentColor transparent transparent; }\n.aq-fish.left .aq-fish-tail { right:-14px; border-width:12px 0 12px 16px; border-color:transparent transparent transparent currentColor; }\n@keyframes aqTailWag { 0%,100%{transform:translateY(-50%) scaleY(1)} 50%{transform:translateY(-50%) scaleY(.75)} }\n.aq-fin-top { position:absolute; top:-9px; left:50%; transform:translateX(-50%); width:0; height:0; border-style:solid; border-width:0 8px 12px 8px; border-color:transparent transparent currentColor transparent; animation:aqFinWave 1.2s ease-in-out infinite; }\n@keyframes aqFinWave { 0%,100%{transform:translateX(-50%) rotate(-4deg)} 50%{transform:translateX(-50%) rotate(4deg)} }\n.aq-fin-bottom { position:absolute; bottom:-6px; left:50%; transform:translateX(-50%); width:0; height:0; border-style:solid; border-width:8px 6px 0 6px; border-color:currentColor transparent transparent transparent; }\n/* Bubbles */\n.aq-bubble { position:absolute; border-radius:50%; background:radial-gradient(circle at 30% 30%,rgba(255,255,255,.85),rgba(255,255,255,.3) 40%,rgba(255,255,255,.08) 60%,transparent);\n    box-shadow:inset -1px -1px 6px rgba(255,255,255,.4), inset 1px 1px 6px rgba(0,100,150,.15);\n    animation:aqBubbleRise linear forwards; pointer-events:none; }\n@keyframes aqBubbleRise {\n    0%{transform:translateY(0) translateX(0) scale(1); opacity:.75}\n    50%{transform:translateY(-50%) translateX(8px) scale(1.06)}\n    100%{transform:translateY(-100%) translateX(0) scale(1); opacity:0}\n}\n/* Food */\n.aq-food { position:absolute; width:6px; height:6px; background:radial-gradient(circle,#f4a460,#d2691e); border-radius:50%;\n    animation:aqFoodFall linear forwards; pointer-events:none; z-index:5; }\n@keyframes aqFoodFall { 0%{transform:translateY(0); opacity:1} 100%{transform:translateY(var(--aqFd)); opacity:0} }\n/* Plants */\n.aq-plant { position:absolute; bottom:0; pointer-events:none; }\n.aq-seaweed { width:18px; }\n.aq-seg { width:100%; height:24px; background:linear-gradient(90deg,#1e7b1e,#2e8b2e,#1e7b1e); border-radius:50%; margin-bottom:-12px; animation:aqSway 2.2s ease-in-out infinite; }\n@keyframes aqSway { 0%,100%{transform:translateX(-2px) rotate(-2deg)} 50%{transform:translateX(2px) rotate(2deg)} }\n.aq-broad-leaf { width:16px; height:60px; background:linear-gradient(90deg,#228b22,#32cd32,#228b22); border-radius:50% 50% 0 0; transform-origin:bottom center; animation:aqPlantSway 2.8s ease-in-out infinite; position:absolute; bottom:0; }\n@keyframes aqPlantSway { 0%,100%{transform:rotate(-4deg) scaleY(1)} 50%{transform:rotate(4deg) scaleY(1.02)} }\n/* Coral */\n.aq-coral { position:absolute; bottom:55px; pointer-events:none; }\n.aq-coral-br { width:12px; border-radius:10px 10px 0 0; position:absolute; bottom:0; }\n.aq-anemone { position:absolute; bottom:55px; pointer-events:none; }\n.aq-anemone-base { position:absolute; bottom:0; left:50%; transform:translateX(-50%); width:30px; height:15px; background:linear-gradient(to top,#ff5722,#ff8a65); border-radius:50%; }\n.aq-anemone-tentacle { position:absolute; bottom:11px; width:3px; height:30px; background:linear-gradient(to top,#ff5722,#ffccbc); border-radius:2px; transform-origin:bottom center; animation:aqTentacle 2s ease-in-out infinite; }\n@keyframes aqTentacle { 0%,100%{transform:rotate(var(--ar)) translateX(0)} 50%{transform:rotate(calc(var(--ar)+10deg)) translateX(2px)} 75%{transform:rotate(calc(var(--ar)-10deg)) translateX(-2px)} }\n/* Decorations */\n.aq-castle { position:absolute; bottom:55px; left:50%; transform:translateX(-50%); width:80px; height:100px; pointer-events:none; }\n.aq-castle-tower { position:absolute; bottom:50px; width:22px; height:52px; background:linear-gradient(135deg,#9b8365,#7b6354); border-radius:4px 4px 0 0; }\n.aq-castle-tower.l { left:8px; } .aq-castle-tower.r { right:8px; }\n.aq-castle-tower::before { content:''; position:absolute; top:-12px; left:50%; transform:translateX(-50%);\n    border-left:14px solid transparent; border-right:14px solid transparent; border-bottom:14px solid #7b6354; }\n.aq-castle-body { position:absolute; bottom:0; left:50%; transform:translateX(-50%); width:54px; height:54px; background:linear-gradient(135deg,#8b7355,#6b5344); border-radius:4px; }\n.aq-castle-door { position:absolute; bottom:0; left:50%; transform:translateX(-50%); width:18px; height:28px; background:#2c1810; border-radius:12px 12px 0 0; }\n.treasure { position:absolute; bottom:58px; right:8%; width:46px; height:34px; pointer-events:none; }\n.treasure-body { position:absolute; bottom:0; width:46px; height:26px; background:linear-gradient(135deg,#b8860b,#8b6914); border-radius:4px; border:1.5px solid #5a4008; }\n.treasure-lid { position:absolute; bottom:22px; width:46px; height:18px; background:linear-gradient(135deg,#d4a017,#b8860b); border-radius:8px 8px 0 0; border:1.5px solid #5a4008; transform-origin:bottom; animation:aqTreasureLid 5s ease-in-out infinite; }\n@keyframes aqTreasureLid { 0%,80%,100%{transform:rotateX(0)} 85%,95%{transform:rotateX(-25deg)} }\n.treasure-lock { position:absolute; bottom:14px; left:50%; transform:translateX(-50%); width:9px; height:9px; background:#ffd700; border-radius:50%; border:1.5px solid #b8860b; }\n/* Rocks & starfish */\n.aq-rocks { position:absolute; bottom:50px; left:0; width:100%; height:48px; pointer-events:none; }\n.aq-rock { position:absolute; bottom:0; border-radius:50% 50% 42% 42%; background:linear-gradient(135deg,#6b7280,#4b5563,#374151);\n    box-shadow:inset -3px -3px 10px rgba(0,0,0,.25), inset 2px 2px 8px rgba(255,255,255,.08); }\n.aq-starfish { position:absolute; bottom:62px; width:30px; height:30px; pointer-events:none; fill:#ff6347; filter:drop-shadow(1px 1px 2px rgba(0,0,0,.3)); }\n.aq-shell { position:absolute; bottom:58px; width:26px; height:18px; background:linear-gradient(135deg,#ffe4c4,#deb887,#d2b48c); border-radius:50% 50% 0 0;\n    box-shadow:inset -2px -2px 6px rgba(0,0,0,.15); }\n.aq-shell::after { content:''; position:absolute; bottom:-3px; left:50%; transform:translateX(-50%); width:22px; height:7px; background:linear-gradient(180deg,#deb887,#c4a882); border-radius:0 0 50% 50%; }\n/* Stats */\n.aq-stats { position:absolute; top:8px; left:8px; background:rgba(0,0,0,.5); backdrop-filter:blur(8px); padding:8px 12px; border-radius:10px; color:#fff; font-size:11px; z-index:100; border:1px solid rgba(255,255,255,.1); }\n.aq-stats h4 { margin:0 0 4px; color:#00d4ff; font-size:12px; }\n.aq-stats p { margin:2px 0; display:flex; align-items:center; gap:5px; }\n.aq-stats span { color:#ffd700; font-weight:600; }\n/* Bio plant */\n.aq-bio-stem { position:absolute; bottom:0; left:50%; transform:translateX(-50%); width:4px; height:38px; background:linear-gradient(to top,#1a237e,#3949ab); border-radius:2px; }\n.aq-bio-bulb { position:absolute; width:12px; height:12px; background:radial-gradient(circle,#00ffff,#0097a7); border-radius:50%;\n    animation:aqBioGlow 2s ease-in-out infinite; box-shadow:0 0 8px #00ffff, 0 0 14px #00ffff; }\n@keyframes aqBioGlow { 0%,100%{opacity:.6; transform:scale(1)} 50%{opacity:1; transform:scale(1.1)} }\n/* Ambient particles */\n.aq-particle { position:absolute; width:2px; height:2px; background:rgba(255,255,255,.35); border-radius:50%; pointer-events:none; animation:aqParticleFloat 10s linear infinite; }\n@keyframes aqParticleFloat { 0%{transform:translateY(100%) translateX(0); opacity:0} 10%{opacity:.5} 90%{opacity:.5} 100%{transform:translateY(-100%) translateX(16px); opacity:0} }\n`;\n\n// ===== State =====\nconst FISH_COLORS = [\n    { body:'#ff6b6b', accent:'#ee5a5a', name:'Coral' },\n    { body:'#ffd93d', accent:'#f5c000', name:'Gold' },\n    { body:'#6bcb77', accent:'#4caf50', name:'Mint' },\n    { body:'#4d96ff', accent:'#2196f3', name:'Ocean' },\n    { body:'#ff6eb4', accent:'#ff1493', name:'Pink' },\n    { body:'#9b59b6', accent:'#8e44ad', name:'Violet' },\n    { body:'#ff8c42', accent:'#ff7315', name:'Orange' },\n    { body:'#00d4aa', accent:'#00bcd4', name:'Teal' },\n    { body:'#ffa07a', accent:'#fa8072', name:'Salmon' },\n    { body:'#87ceeb', accent:'#6bb3d9', name:'Sky' }\n];\nconst FISH_NAMES = ['Bubbles','Nemo','Dory','Goldie','Finley','Splash','Coral','Neptune','Pearl','Sunny','Azure','Shimmer','Glitter','Wave','Aqua','Marina','Flash','Zippy','Dash','Spark'];\n\nfunction _loadState(key) {\n    try { const s = localStorage.getItem('aq_' + key); return s ? JSON.parse(s) : null; } catch(e) { return null; }\n}\nfunction _saveState(key, data) { try { localStorage.setItem('aq_' + key, JSON.stringify(data)); } catch(e) {} }\n\nfunction _getFishName() { return FISH_NAMES[Math.floor(Math.random() * FISH_NAMES.length)]; }\n\n// ===== Render =====\nasync function render(content) {\n    content.empty();\n    const wrap = content.createDiv({ cls: 'aquarium-wrap' });\n\n    // Toolbar\n    const toolbar = wrap.createDiv({ cls: 'aquarium-toolbar' });\n    const btn = (cls, label, fn) => {\n        const b = toolbar.createEl('button', { cls: `aq-btn ${cls}`, text: label });\n        b.onclick = fn; return b;\n    };\n    btn('aq-btn-fish', t('mod.aquarium.addFish') || '+ Fish', () => addFish(bodyEl));\n    btn('aq-btn-bubble', t('mod.aquarium.addBubble') || '+ Bubble', () => addBubbleSrc(bodyEl));\n    btn('aq-btn-plant', t('mod.aquarium.addPlant') || '+ Plant', () => addPlant(bodyEl));\n    btn('aq-btn-food', t('mod.aquarium.addFood') || '+ Food', () => dropFood(bodyEl));\n    btn('aq-btn-decor', (t('mod.aquarium.decor') || '+ Decor'), () => toggleDecorPanel());\n    btn('aq-btn-clear', (t('mod.aquarium.clear') || 'Clear'), () => clearFish(bodyEl));\n    btn('aq-btn-reset', (t('mod.aquarium.reset') || 'Reset'), () => resetAll(bodyEl));\n\n    // Body\n    const bodyEl = wrap.createDiv({ cls: 'aq-body' });\n    const theme = settings.theme || 'tropical';\n    bodyEl.addClass('aq-theme-' + theme);\n    if (settings.nightMode) bodyEl.addClass('aq-night');\n\n    // Rays\n    const rays = bodyEl.createDiv({ cls: 'aq-rays' });\n    for (let i = 0; i < 5; i++) {\n        const r = rays.createDiv({ cls: 'aq-ray' });\n        r.style.left = [8,26,50,72,88][i] + '%';\n        r.style.width = [100,70,90,60,80][i] + 'px';\n        r.style.animationDelay = (i * 1.5) + 's';\n    }\n\n    // Surface\n    bodyEl.createDiv({ cls: 'aq-surface' });\n\n    // Stats\n    const stats = bodyEl.createDiv({ cls: 'aq-stats' });\n    stats.createEl('h4', { text: '📊 ' + (t('mod.aquarium.stats') || 'Aquarium') });\n    const statFish = stats.createEl('p'); statFish.innerHTML = '🐟 ' + (t('mod.aquarium.fish') || 'Fish') + ': <span id=\"aqFc\">0</span>';\n    const statPlant = stats.createEl('p'); statPlant.innerHTML = '🌿 ' + (t('mod.aquarium.plants') || 'Plants') + ': <span id=\"aqPc\">0</span>';\n    const statBub = stats.createEl('p'); statBub.innerHTML = '💨 ' + (t('mod.aquarium.bubbles') || 'Bubbles') + ': <span id=\"aqBc\">0</span>';\n\n    // Sand\n    const sand = bodyEl.createDiv({ cls: 'aq-sand' });\n    sand.addClass('aq-sand-' + (settings.sandColor || 'golden'));\n\n    // Rocks\n    const rocks = bodyEl.createDiv({ cls: 'aq-rocks' });\n    const rockPos = [[4,60,36],[10,44,28],[88,68,40],[78,34,24],[42,50,32]];\n    rockPos.forEach(([l,w,h]) => {\n        const rk = rocks.createDiv({ cls: 'aq-rock' });\n        rk.style.left = l + '%'; rk.style.width = w + 'px'; rk.style.height = h + 'px';\n    });\n\n    // Castle\n    const castle = bodyEl.createDiv({ cls: 'aq-castle' });\n    castle.createDiv({ cls: 'aq-castle-tower l' });\n    castle.createDiv({ cls: 'aq-castle-tower r' });\n    castle.createDiv({ cls: 'aq-castle-body' }).createDiv({ cls: 'aq-castle-door' });\n\n    // Treasure\n    const treasure = bodyEl.createDiv({ cls: 'treasure' });\n    treasure.createDiv({ cls: 'treasure-lid' });\n    treasure.createDiv({ cls: 'treasure-body' }).createDiv({ cls: 'treasure-lock' });\n\n    // Starfish\n    const sf1 = bodyEl.createSvg('svg', { cls: 'aq-starfish' }); sf1.style.left='12%'; sf1.style.transform='rotate(12deg)';\n    sf1.innerHTML = '<polygon points=\"50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35\" fill=\"#ff6347\"/>';\n    const sf2 = bodyEl.createSvg('svg', { cls: 'aq-starfish' }); sf2.style.right='4%'; sf2.style.transform='rotate(-18deg)';\n    sf2.innerHTML = '<polygon points=\"50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35\" fill=\"#ff6347\"/>';\n\n    // Shells\n    const sh1 = bodyEl.createDiv({ cls: 'aq-shell' }); sh1.style.left='32%';\n    const sh2 = bodyEl.createDiv({ cls: 'aq-shell' }); sh2.style.right='22%'; sh2.style.transform='rotate(8deg)';\n\n    // Plants container\n    const plantsContainer = bodyEl.createDiv(); // for dynamic plants\n\n    // Ambient particles\n    for (let i = 0; i < 15; i++) {\n        const p = bodyEl.createDiv({ cls: 'aq-particle' });\n        p.style.left = Math.random() * 100 + '%';\n        p.style.top = Math.random() * 100 + '%';\n        p.style.animationDelay = Math.random() * 10 + 's';\n        p.style.animationDuration = (8 + Math.random() * 4) + 's';\n    }\n\n    // State\n    let state = _loadState('state') || { fish:[], plants:[], bubbleSources:[], foodCount:0 };\n    let fishIdCounter = 0;\n    const fishEls = new Map();\n    let afId = null;\n\n    function updateStats() {\n        const fc = document.getElementById('aqFc'); if(fc) fc.textContent = state.fish.length;\n        const pc = document.getElementById('aqPc'); if(pc) pc.textContent = state.plants.length;\n        const bc = document.getElementById('aqBc'); if(bc) bc.textContent = state.bubbleSources.length;\n    }\n\n    // Restore fish\n    state.fish.forEach(fd => {\n        createFishEl(bodyEl, fd);\n        const num = parseInt(fd.id, 10);\n        if (!isNaN(num) && num >= fishIdCounter) fishIdCounter = num + 1;\n    });\n    // Restore plants\n    state.plants.forEach(pd => createPlantEl(plantsContainer, pd));\n    updateStats();\n\n    // ===== Fish creation =====\n    function createFishEl(container, fd) {\n        const f = container.createDiv({ cls: 'aq-fish' });\n        f.style.left = fd.x + 'px'; f.style.top = fd.y + 'px';\n        f.style.color = fd.color.accent;\n\n        const fb = f.createDiv({ cls: 'aq-fish-body' });\n        fb.style.background = `linear-gradient(135deg, ${fd.color.body}, ${fd.color.accent})`;\n        const eyeCls = fd.dir === 'right' ? 'aq-fish-eye right' : 'aq-fish-eye left';\n        fb.createDiv({ cls: eyeCls });\n        f.createDiv({ cls: 'aq-fish-tail' });\n        f.createDiv({ cls: 'aq-fin-top' });\n        f.createDiv({ cls: 'aq-fin-bottom' });\n\n        f.onclick = (e) => { e.stopPropagation(); showFishTip(fd, e); };\n        fishEls.set(fd.id, f);\n        return f;\n    }\n\n    function showFishTip(fd, e) {\n        const age = Math.floor((Date.now() - new Date(fd.birth).getTime()) / 60000);\n        let tip = document.getElementById('aqTooltip');\n        if (!tip) {\n            tip = document.body.createDiv({ id: 'aqTooltip' });\n            tip.style.cssText = 'position:fixed;background:rgba(0,0,0,.82);color:#fff;padding:6px 10px;border-radius:8px;font-size:11px;pointer-events:none;z-index:9999;transition:opacity .2s;max-width:200px;';\n        }\n        tip.innerHTML `<strong>${fd.name}</strong><br>${t('mod.aquarium.color')||'Color'}: ${fd.color.name}<br>${t('mod.aquarium.age')||'Age'}: ${age}${t('mod.aquarium.min')||'min'}<br>${t('mod.aquarium.fed')||'Fed'}: ${fd.fed}`;\n        tip.style.left = (e.clientX + 8) + 'px'; tip.style.top = (e.clientY + 8) + 'px'; tip.style.opacity = '1';\n        setTimeout(() => { tip.style.opacity = '0'; }, 3000);\n    }\n\n    function addFish(container) {\n        const rect = container.getBoundingClientRect();\n        const fd = {\n            id: String(fishIdCounter++),\n            x: Math.random() * (rect.width - 80) + 16,\n            y: Math.random() * (rect.height - 160) + 40,\n            sx: (Math.random() * 1.6 + 0.8) * (Math.random() < 0.5 ? 1 : -1),\n            sy: (Math.random() - 0.5) * 0.8,\n            color: FISH_COLORS[Math.floor(Math.random() * FISH_COLORS.length)],\n            dir: Math.random() < 0.5 ? 'right' : 'left',\n            name: _getFishName(), birth: new Date().toISOString(), fed: 0\n        };\n        state.fish.push(fd);\n        createFishEl(container, fd);\n        updateStats(); _saveState('state', state);\n    }\n\n    // ===== Bubbles =====\n    function createBubble(container, x, y) {\n        const b = container.createDiv({ cls: 'aq-bubble' });\n        const sz = Math.random() * 12 + 6;\n        b.style.width = sz + 'px'; b.style.height = sz + 'px';\n        b.style.left = x + 'px'; b.style.bottom = (container.offsetHeight - y) + 'px';\n        b.style.animationDuration = (2.5 + Math.random() * 2.5) + 's';\n        setTimeout(() => b.remove(), 5500);\n    }\n\n    function addBubbleSrc(container) {\n        const src = { id: Date.now() + Math.random(), x: Math.random() * (container.offsetWidth - 30) + 15 };\n        state.bubbleSources.push(src);\n        src.interval = setInterval(() => {\n            createBubble(container, src.x + (Math.random() - 0.5) * 16, container.offsetHeight - 60);\n        }, 500 + Math.random() * 800);\n        updateStats(); _saveState('state', state);\n    }\n\n    // Start existing bubble sources\n    state.bubbleSources.forEach(src => {\n        src.interval = setInterval(() => {\n            createBubble(bodyEl, src.x, bodyEl.offsetHeight - 60);\n        }, 700);\n    });\n\n    // ===== Plants =====\n    function createPlantEl(cont, pd) {\n        let el;\n        if (pd.type === 'seaweed') {\n            el = cont.createDiv({ cls: 'aq-plant aq-seaweed' });\n            el.style.left = pd.x + 'px';\n            for (let i = 0; i < (pd.segs || 4); i++) {\n                const s = el.createDiv({ cls: 'aq-seg' });\n                s.style.background = `linear-gradient(90deg,${pd.c||'#1e7b1e'},${pd.cl||'#2e8b2e'},${pd.c||'#1e7b1e'})`;\n                s.style.animationDelay = (i * 0.12) + 's';\n            }\n        } else if (pd.type === 'broad') {\n            el = cont.createDiv({ cls: 'aq-plant' });\n            el.style.left = pd.x + 'px';\n            for (let i = 0; i < (pd.leaves || 3); i++) {\n                const l = el.createDiv({ cls: 'aq-broad-leaf' });\n                l.style.background = `linear-gradient(90deg,${pd.c||'#228b22'},${pd.cl||'#32cd32'},${pd.c||'#228b22'})`;\n                l.style.height = (45 + Math.random() * 30) + 'px';\n                l.style.transform = `rotate(${(i - 1) * 14}deg)`;\n                l.style.animationDelay = (i * 0.18) + 's';\n                l.style.zIndex = (pd.leaves || 3) - i;\n            }\n        } else if (pd.type === 'coral') {\n            el = cont.createDiv({ cls: 'aq-coral' });\n            el.style.left = pd.x + 'px';\n            for (let i = 0; i < 3; i++) {\n                const br = el.createDiv({ cls: 'aq-coral-br' });\n                br.style.background = `linear-gradient(180deg,${pd.c||'#ff6b6b'},${pd.cd||'#ee5a5a'})`;\n                br.style.left = (i * 12) + 'px';\n                br.style.height = (40 + Math.random() * 20) + 'px';\n            }\n        } else if (pd.type === 'anemone') {\n            el = cont.createDiv({ cls: 'aq-anemone' });\n            el.style.left = pd.x + 'px';\n            el.createDiv({ cls: 'aq-anemone-base' }).style.background = `linear-gradient(to top,${pd.c||'#ff5722'},${pd.cl||'#ff8a65'})`;\n            for (let i = 0; i < 10; i++) {\n                const t = el.createDiv({ cls: 'aq-anemone-tentacle' });\n                t.style.background = `linear-gradient(to top,${pd.c||'#ff5722'},${pd.ct||'#ffccbc'})`;\n                t.style.left = (4 + i * 2.6) + 'px';\n                t.style.setProperty('--ar', ((i - 5) * 7) + 'deg');\n                t.style.animationDelay = (i * 0.1) + 's';\n                t.style.height = (24 + Math.random() * 16) + 'px';\n            }\n        } else if (pd.type === 'bio') {\n            el = cont.createDiv({ cls: 'aq-plant aq-bio aq-bio' });\n            el.style.left = pd.x + 'px';\n            el.createDiv({ cls: 'aq-bio-stem' });\n            const colors = ['#00ffff','#00ff33','#ff00ff','#ffff33'];\n            const gc = pd.glow || colors[Math.floor(Math.random() * colors.length)];\n            for (let i = 0; i < 3; i++) {\n                const bulb = el.createDiv({ cls: 'aq-bio-bulb' });\n                bulb.style.background = `radial-gradient(circle,${gc},${gc}88)`;\n                bulb.style.boxShadow = `0 0 8px ${gc}, 0 0 14px ${gc}`;\n                bulb.style.top = (-8 - i * 13) + 'px';\n                bulb.style.left = (i % 2 === 0 ? -4 : 12) + 'px';\n                bulb.style.animationDelay = (i * 0.4) + 's';\n            }\n        }\n        return el;\n    }\n\n    function addPlant(container) {\n        const types = ['seaweed','broad','coral','anemone','bio'];\n        const type = types[Math.floor(Math.random() * types.length)];\n        const greens = [\n            { c:'#228b22', cl:'#32cd32' }, { c:'#006400', cl:'#228b22' },\n            { c:'#2e8b57', cl:'#3cb371' }, { c:'#20b2aa', cl:'#48d1cc' }\n        ];\n        const g = greens[Math.floor(Math.random() * greens.length)];\n        const coralColors = [\n            { c:'#ff6b6b', cd:'#ee5a5a' }, { c:'#9c27b0', cd:'#7b1fa2' },\n            { c:'#e91e63', cd:'#f48fb1' }, { c:'#ff5722', cd:'#ff8a65' }\n        ];\n        const cc = type === 'coral' || type === 'anemone' ? coralColors[Math.floor(Math.random() * coralColors.length)] : g;\n\n        const pd = { id: Date.now() + Math.random(), x: Math.random() * (container.offsetWidth - 40) + 20, type:type, ...cc };\n        if (type === 'seaweed') pd.segs = Math.floor(Math.random() * 3) + 3;\n        if (type === 'broad') pd.leaves = Math.floor(Math.random() * 2) + 2;\n\n        state.plants.push(pd);\n        createPlantEl(container, pd);\n        updateStats(); _saveState('state', state);\n    }\n\n    // ===== Food =====\n    function dropFood(container) {\n        for (let i = 0; i < 8; i++) {\n            setTimeout(() => {\n                const fd = container.createDiv({ cls: 'aq-food' });\n                fd.style.left = (Math.random() * (container.offsetWidth - 30) + 15) + 'px';\n                fd.style.top = '16px';\n                fd.style.setProperty('--aqFd', (container.offsetHeight - 100) + 'px');\n                fd.style.animationDuration = (3.5 + Math.random() * 1.5) + 's';\n                setTimeout(() => fd.remove(), 5000);\n\n                state.fish.forEach(f => { f.sx *= 1.4; setTimeout(()=>{ f.sx/=1.4; }, 2500); f.fed++; });\n            }, i * 80);\n        }\n        state.foodCount++;\n        updateStats(); _saveState('state', state);\n    }\n\n    // ===== Clear / Reset =====\n    function clearFish(container) {\n        state.fish.forEach(fd => { const el = fishEls.get(fd.id); if(el) el.remove(); fishEls.delete(fd.id); });\n        state.fish = []; updateStats(); _saveState('state', state);\n    }\n\n    function resetAll(container) {\n        if (!confirm(t('mod.aquarium.resetConfirm') || 'Reset entire aquarium?')) return;\n        clearFish(container);\n        plantsContainer.empty();\n        state.bubbleSources.forEach(s => { if(s.interval) clearInterval(s.interval); });\n        state = { fish:[], plants:[], bubbleSources:[], foodCount:0 };\n        fishIdCounter = 0;\n        updateStats(); _saveState('state', state);\n    }\n\n    // ===== Decor panel (simple) =====\n    function toggleDecorPanel() {\n        // Just add random decoration on click\n        const decorTypes = ['coral', 'anemone'];\n        const type = decorTypes[Math.floor(Math.random() * decorTypes.length)];\n        const pd = { id: Date.now()+Math.random(), x: Math.random()*((bodyEl.offsetWidth-80))+40, type, c:'#ff6b6b',cd:'#ee5a5a',cl:'#ff8a65',ct:'#ffccbc' };\n        state.plants.push(pd);\n        createPlantEl(plantsContainer, pd);\n        updateStats(); _saveState('state', state);\n    }\n\n    // Click to make bubbles\n    bodyEl.addEventListener('click', (e) => {\n        const rect = bodyEl.getBoundingClientRect();\n        const x = e.clientX - rect.left, y = e.clientY - rect.top;\n        for (let i = 0; i < 4; i++)\n            setTimeout(() => createBubble(bodyEl, x + (Math.random()-0.5)*24, rect.height-y), i*40);\n    });\n\n    // ===== Game loop =====\n    function gameLoop() {\n        const rect = bodyEl.getBoundingClientRect();\n        const w = rect.width, h = rect.height;\n        state.fish.forEach(fd => {\n            fd.x += fd.sx; fd.y += fd.sy;\n            if (fd.x <= 12 || fd.x >= w - 56) { fd.sx *= -1; fd.dir = fd.sx > 0 ? 'right' : 'left'; fd.x = Math.max(12, Math.min(fd.x, w-56)); }\n            if (fd.y <= 36 || fd.y >= h - 120) { fd.sy *= -1; fd.y = Math.max(36, Math.min(fd.y, h-120)); }\n            if (Math.random() < 0.012) fd.sy = (Math.random()-0.5)*1.6;\n            if (Math.random() < 0.005) { fd.sx *= -1; fd.dir = fd.sx>0?'right':'left'; }\n            const el = fishEls.get(fd.id);\n            if (el) { el.style.left = fd.x+'px'; el.style.top = fd.y+'px'; el.className = 'aq-fish'; }\n        });\n        afId = requestAnimationFrame(gameLoop);\n    }\n\n    gameLoop();\n\n    // Auto-save every 30s\n    setInterval(() => { _saveState('state', state); }, 30000);\n\n    // Cleanup on visibility change\n    document.addEventListener('visibilitychange', () => {\n        if (document.hidden) { if(afId) cancelAnimationFrame(afId); afId=null; }\n        else if (!afId) gameLoop();\n    });\n}\n\nfunction renderSettings(containerEl, settings, save) {\n    // Theme select\n    containerEl.createEl('label', { text: t('mod.aquarium.themeLabel') || 'Theme' });\n    const themeSel = containerEl.createEl('select');\n    ['tropical','ocean','pond','sunset','midnight'].forEach(t => {\n        const opt = themeSel.createEl('option', { value:t, text: t.charAt(0).toUpperCase()+t.slice(1) });\n        if (settings.theme === t) opt.selected = true;\n    });\n    themeSel.onchange = () => { settings.theme = themeSel.value; save(); };\n\n    // Sand color\n    containerEl.createEl('label', { text: t('mod.aquarium.sandLabel') || 'Sand' });\n    const sandSel = containerEl.createEl('select');\n    ['golden','white','dark'].forEach(s => {\n        const opt = sandSel.createEl('option', { value:s, text: s.charAt(0).toUpperCase()+s.slice(1) });\n        if (settings.sandColor === s) opt.selected = true;\n    });\n    sandSel.onchange = () => { settings.sandColor = sandSel.value; save(); };\n\n    // Night mode\n    const nightCb = containerEl.createEl('input', { type: 'checkbox' });\n    nightCb.checked = !!settings.nightMode;\n    const nightLbl = containerEl.createEl('label', { text: t('mod.aquarium.nightMode') || 'Night Mode', attr:{for:''} });\n    containerEl.createDiv().append(nightCb, nightLbl);\n    nightCb.onchange = () => { settings.nightMode = nightCb.checked; save(); };\n\n    // Reset button\n    const resetBtn = containerEl.createEl('button', { text: t('mod.aquarium.resetData') || 'Reset Saved Data' });\n    resetBtn.style.cssText = 'margin-top:10px;padding:6px 14px;border-radius:8px;border:1px solid var(--v6-error,#e53e3e);color:var(--v6-error,#e53e3e);background:transparent;cursor:pointer;';\n    resetBtn.onclick = () => { localStorage.removeItem('aq_state'); };\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "autoplay-loop": "// autoplay-loop 模块 - 全局媒体自动播放控制\n// 源插件: autoplay-and-loop (自动播放音频)\n// 核心功能保留: IntersectionObserver 实际控制 video/audio 元素\nconst id = 'autoplay-loop';\nconst title = t('mod.autoplay');\nconst icon = '▶️';\n\nconst defaultSettings = {\n    autoplayAudio: true,\n    autoplayVideo: true,\n    loopAudio: true,\n    loopVideo: true,\n    muteAutoplayedAudio: false,\n    muteAutoplayedVideo: true,\n    singlePlaybackAudio: true,\n    singlePlaybackVideo: false,\n    pauseOutOfViewAudio: true,\n    pauseOutOfViewVideo: true\n};\n\nconst styles = `\n.ap-wrap { padding: 10px 12px; display: flex; flex-direction: column; height: 100%; }\n.ap-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }\n.ap-title { font-size: 12px; font-weight: 600; color: var(--v6-primary); }\n.ap-toggle { font-size: 11px; padding: 3px 10px; border: 1px solid var(--background-modifier-border); border-radius: 10px; cursor: pointer; transition: all 0.15s; background: var(--background-modifier-form-field); color: var(--text-muted); }\n.ap-toggle.active { background: #4caf50; color: white; border-color: #4caf50; }\n.ap-section { margin-bottom: 10px; padding: 8px; background: var(--background-modifier-form-field); border-radius: 6px; }\n.ap-section h4 { font-size: 11px; color: var(--text-muted); margin: 0 0 4px; }\n.ap-row { display: flex; align-items: center; justify-content: space-between; padding: 3px 4px; font-size: 11px; }\n.ap-row:hover { background: var(--background-modifier-hover); border-radius: 3px; }\n.ap-row label { flex: 1; color: var(--text-normal); }\n.ap-row .ap-indicator { font-size: 10px; font-weight: 600; padding: 1px 6px; border-radius: 8px; }\n.ap-row .ap-indicator.on { color: #4caf50; background: rgba(76,175,80,0.1); }\n.ap-row .ap-indicator.off { color: var(--text-muted); }\n.ap-stats { margin-top: 8px; padding: 8px; background: var(--background-modifier-form-field); border-radius: 6px; font-size: 10px; color: var(--text-muted); }\n.ap-stats strong { color: var(--text-normal); }\n`;\n\n// ============ 媒体控制引擎 ============\nvar observer = null;\nvar intervalId = null;\nvar engineActive = false;\n\nfunction startEngine(s) {\n    if (engineActive) stopEngine(s);\n    engineActive = true;\n\n    function applyToElement(el) {\n        var tag = el.tagName.toLowerCase();\n        var isVideo = tag === 'video';\n        var isAudio = tag === 'audio';\n\n        if (isVideo && s.autoplayVideo) {\n            el.setAttribute('autoplay', '');\n            if (s.muteAutoplayedVideo) el.setAttribute('muted', '');\n            if (s.loopVideo) el.setAttribute('loop', '');\n            el.play().catch(function() {});\n        }\n        if (isAudio && s.autoplayAudio) {\n            el.setAttribute('autoplay', '');\n            if (s.muteAutoplayedAudio) el.setAttribute('muted', '');\n            if (s.loopAudio) el.setAttribute('loop', '');\n            el.play().catch(function() {});\n        }\n    }\n\n    // IntersectionObserver: 可见性控制\n    observer = new IntersectionObserver(function(entries) {\n        entries.forEach(function(entry) {\n            var el = entry.target;\n            var tag = el.tagName.toLowerCase();\n\n            if (entry.isIntersecting) {\n                // 进入视野\n                applyToElement(el);\n            } else {\n                // 离开视野\n                if (tag === 'video' && s.pauseOutOfViewVideo) el.pause();\n                if (tag === 'audio' && s.pauseOutOfViewAudio) el.pause();\n            }\n        });\n    }, { threshold: 0.1 });\n\n    // 扫描并附加observer\n    function scan() {\n        document.querySelectorAll('video, audio').forEach(function(el) {\n            if (!el.dataset.apObserved) {\n                el.dataset.apObserved = '1';\n                observer.observe(el);\n                applyToElement(el);\n            }\n        });\n    }\n\n    scan();\n\n    // 定时扫描新出现的媒体元素\n    intervalId = setInterval(scan, 3600000)/*TEMP_DISABLED*/;\n}\n\nfunction stopEngine(s) {\n    engineActive = false;\n    if (observer) { observer.disconnect(); observer = null; }\n    if (intervalId) { clearInterval(intervalId); intervalId = null; }\n    // 清理标记\n    document.querySelectorAll('[data-ap-observed]').forEach(function(el) {\n        delete el.dataset.apObserved;\n    });\n}\n\n// ============ 渲染 ============\n\nasync function render(container) {\n    container.addClass('ap-wrap');\n    var s = settings;\n\n    // 总控开关\n    var header = container.createDiv({ cls: 'ap-header' });\n    header.createDiv({ text: t('mod.autoplay.title'), cls: 'ap-title' });\n    var toggleBtn = header.createEl('button', { cls: 'ap-toggle' });\n    updateToggleBtn();\n\n    function updateToggleBtn() {\n        toggleBtn.textContent = engineActive ? t('mod.autoplay.running') : t('mod.autoplay.stopped');\n        toggleBtn.className = 'ap-toggle' + (engineActive ? ' active' : '');\n    }\n\n    toggleBtn.addEventListener('click', function() {\n        if (engineActive) { stopEngine(s); } else { startEngine(s); }\n        updateToggleBtn();\n        if (typeof saveCallback === 'function') saveCallback();\n    });\n\n    // 设置项\n    function makeRow(sectionEl, key, label) {\n        var row = sectionEl.createDiv({ cls: 'ap-row' });\n        row.createEl('label', { text: label });\n        var ind = row.createDiv({\n            text: s[key] ? t('mod.autoplay.on') : t('mod.autoplay.off'),\n            cls: 'ap-indicator ' + (s[key] ? 'on' : 'off')\n        });\n        row.addEventListener('click', function() {\n            s[key] = !s[key];\n            ind.textContent = s[key] ? t('mod.autoplay.on') : t('mod.autoplay.off');\n            ind.className = 'ap-indicator ' + (s[key] ? 'on' : 'off');\n            if (engineActive) { stopEngine(s); startEngine(s); }\n            if (typeof saveCallback === 'function') saveCallback();\n        });\n        row.style.cursor = 'pointer';\n    }\n\n    // 视频设置\n    var videoSection = container.createDiv({ cls: 'ap-section' });\n    videoSection.createEl('h4', { text: t('mod.autoplay.grp.video') });\n    makeRow(videoSection, 'autoplayVideo', t('mod.autoplay.opt.autoplay'));\n    makeRow(videoSection, 'muteAutoplayedVideo', t('mod.autoplay.opt.automute'));\n    makeRow(videoSection, 'loopVideo', t('mod.autoplay.opt.loop'));\n    makeRow(videoSection, 'pauseOutOfViewVideo', t('mod.autoplay.opt.pauseOut'));\n\n    // 音频设置\n    var audioSection = container.createDiv({ cls: 'ap-section' });\n    audioSection.createEl('h4', { text: t('mod.autoplay.grp.audio') });\n    makeRow(audioSection, 'autoplayAudio', t('mod.autoplay.opt.autoplay'));\n    makeRow(audioSection, 'muteAutoplayedAudio', t('mod.autoplay.opt.automute'));\n    makeRow(audioSection, 'loopAudio', t('mod.autoplay.opt.loop'));\n    makeRow(audioSection, 'pauseOutOfViewAudio', t('mod.autoplay.opt.pauseOut'));\n\n    // 高级设置\n    var advSection = container.createDiv({ cls: 'ap-section' });\n    advSection.createEl('h4', { text: t('mod.autoplay.grp.advanced') });\n    makeRow(advSection, 'singlePlaybackAudio', t('mod.autoplay.opt.singleAudio'));\n    makeRow(advSection, 'singlePlaybackVideo', t('mod.autoplay.opt.singleVideo'));\n\n    // 实时统计\n    var stats = container.createDiv({ cls: 'ap-stats' });\n    function updateStats() {\n        var videos = document.querySelectorAll('video').length;\n        var audios = document.querySelectorAll('audio').length;\n        stats.innerHTML = t('mod.autoplay.status') + '<strong>' + videos + '</strong>' + t('mod.autoplay.statusVideo') + '<strong>' + audios + '</strong>' + t('mod.autoplay.statusAudio') + '<strong>' + (engineActive ? t('mod.autoplay.engineRunning') : t('mod.autoplay.engineStopped')) + '</strong>';\n    }\n    updateStats();\n    var statsIntervalId = setInterval(updateStats, 3600000)/*TEMP_DISABLED*/;\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.autoplay.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.autoplay.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n}\n\n// 导出 onunload 供框架调用清理\nmodule.exports.onunload = function() {\n    stopEngine();\n    if (typeof statsIntervalId !== \"undefined\" && statsIntervalId) clearInterval(statsIntervalId);\n};\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "calendar": "/**\n * 日历模块 V15\n * 格式：V14（含 id/styles/renderSettings）\n * 功能：V11 完整版（月历翻页 + 农历 + 节日 + 节气 + 天干地支）\n */\nconst id = 'calendar';\nconst title = t('mod.calendar');\nconst icon = '📅';\n\nconst defaultSettings = {\n    showLunar: true,\n    showHoliday: true\n};\n\nconst styles = `/* 日历模块样式已在 styles.css 中定义 */`;\n\n// ===== 农历工具 =====\nconst LUNAR_INFO = [\n    0x04bd8,0x04ae0,0x0a570,0x054d5,0x0d260,0x0d950,0x16554,0x056a0,0x09ad0,0x055d2,\n    0x04ae0,0x0a5b6,0x0a4d0,0x0d250,0x1d255,0x0b540,0x0d6a0,0x0ada2,0x095b0,0x14977,\n    0x04970,0x0a4b0,0x0b4b5,0x06a50,0x06d40,0x1ab54,0x02b60,0x09570,0x052f2,0x04970,\n    0x06566,0x0d4a0,0x0ea50,0x06e95,0x05ad0,0x02b60,0x186e3,0x092e0,0x1c8d7,0x0c950,\n    0x0d4a0,0x1d8a6,0x0b550,0x056a0,0x1a5b4,0x025d0,0x092d0,0x0d2b2,0x0a950,0x0b557,\n    0x06ca0,0x0b550,0x15355,0x04da0,0x0a5b0,0x14573,0x052b0,0x0a9a8,0x0e950,0x06aa0,\n    0x0aea6,0x0ab50,0x04b60,0x0aae4,0x0a570,0x05260,0x0f263,0x0d950,0x05b57,0x056a0,\n    0x096d0,0x04dd5,0x04ad0,0x0a4d0,0x0d4d4,0x0d250,0x0d558,0x0b540,0x0b6a0,0x195a6,\n    0x095b0,0x049b0,0x0a974,0x0a4b0,0x0b27a,0x06a50,0x06d40,0x0af46,0x0ab60,0x09570,\n    0x04af5,0x04970,0x064b0,0x074a3,0x0ea50,0x06aa0,0x0a6b6,0x056a0,0x02b40,0x0acb6,\n    0x0a940,0x0a950,0x0b4a6,0x0b550,0x0d2a0,0x11d25,0x0d960,0x05954,0x056a0,0x0aba0,\n    0x1a3c5,0x09250,0x0a950,0x0b965,0x0aa40,0x0bccd,0x0b550,0x04b60,0x0a576,0x0a520,\n    0x0dd45,0x0d950,0x056a0,0x14ad5,0x055d0,0x0a9b0,0x14b75,0x04970,0x0a4b0,0x0e950,\n    0x06b60,0x0b4b5,0x05ab0,0x02b40,0x1ab60,0x096d5,0x095b0,0x049b0,0x0a4b0,0x0b8a6\n];\n\nconst TG = ['甲','乙','丙','丁','戊','己','庚','辛','壬','癸'];\nconst DZ = ['子','丑','寅','卯','辰','巳','午','未','申','酉','戌','亥'];\nconst ANIMALS = ['鼠','牛','虎','兔','龙','蛇','马','羊','猴','鸡','狗','猪'];\nconst LUNAR_MONTHS = ['正','二','三','四','五','六','七','八','九','十','十一','十二'];\nconst LUNAR_DAYS = ['初一','初二','初三','初四','初五','初六','初七','初八','初九','初十',\n    '十一','十二','十三','十四','十五','十六','十七','十八','十九','二十',\n    '廿一','廿二','廿三','廿四','廿五','廿六','廿七','廿八','廿九','三十'];\n\nfunction lYearDays(y) {\n    let i, sum = 348;\n    for (i = 0x8000; i > 0x8; i >>= 1) {\n        sum += (LUNAR_INFO[y - 1900] & i) ? 1 : 0;\n    }\n    return sum + leapDays(y);\n}\nfunction leapMonth(y) { return LUNAR_INFO[y - 1900] & 0xf; }\nfunction leapDays(y) {\n    if (leapMonth(y)) {\n        return (LUNAR_INFO[y - 1900] & 0x10000) ? 30 : 29;\n    }\n    return 0;\n}\nfunction monthDays(y, m) {\n    return (LUNAR_INFO[y - 1900] & (0x10000 >> m)) ? 30 : 29;\n}\n\nfunction solarToLunar(sYear, sMonth, sDay) {\n    let y, m, d, leapYear = false;\n    let dayCyclical, monthCyclical;\n    \n    const baseDate = new Date(1900, 0, 31);\n    const objDate = new Date(sYear, sMonth - 1, sDay);\n    let offset = Math.round((objDate - baseDate) / 86400000);\n    \n    let i;\n    for (i = 1900; i < 2100 && offset > 0; i++) {\n        let daysInYear = lYearDays(i);\n        offset -= daysInYear;\n    }\n    if (offset < 0) {\n        offset += lYearDays(--i);\n    }\n    \n    y = i;\n    const leap = leapMonth(y);\n    leapYear = false;\n    \n    for (i = 1; i < 13 && offset > 0; i++) {\n        if (leap > 0 && i === leap + 1 && !leapYear) {\n            --i;\n            leapYear = true;\n            d = leapDays(y);\n        } else {\n            d = monthDays(y, i);\n        }\n        if (leapYear && i === leap + 1) leapYear = false;\n        offset -= d;\n    }\n    \n    if (offset === 0 && leap > 0 && i === leap + 1) {\n        if (leapYear) {\n            leapYear = false;\n        } else {\n            leapYear = true;\n            --i;\n        }\n    }\n    if (offset < 0) {\n        offset += d;\n        --i;\n    }\n    \n    m = i;\n    d = offset + 1;\n    \n    const cyclicalYear = y - 1900 + 36;\n    const gan = TG[cyclicalYear % 10];\n    const zhi = DZ[cyclicalYear % 12];\n    const animal = ANIMALS[cyclicalYear % 12];\n    \n    return {\n        year: y,\n        month: m,\n        day: d,\n        isLeap: leapYear,\n        ganZhi: gan + zhi,\n        animal,\n        monthStr: (leapYear ? '闰' : '') + LUNAR_MONTHS[m - 1] + '月',\n        dayStr: LUNAR_DAYS[d - 1]\n    };\n}\n\n// 节气表（每年近似，精度够用）\nconst SOLAR_TERMS = {\n    '1-6': '小寒', '1-20': '大寒',\n    '2-4': '立春', '2-19': '雨水',\n    '3-6': '惊蛰', '3-21': '春分',\n    '4-5': '清明', '4-20': '谷雨',\n    '5-6': '立夏', '5-21': '小满',\n    '6-6': '芒种', '6-21': '夏至',\n    '7-7': '小暑', '7-23': '大暑',\n    '8-7': '立秋', '8-23': '处暑',\n    '9-8': '白露', '9-23': '秋分',\n    '10-8': '寒露', '10-23': '霜降',\n    '11-7': '立冬', '11-22': '小雪',\n    '12-7': '大雪', '12-22': '冬至'\n};\n\n// 法定节假日\nconst HOLIDAYS = {\n    '1-1': '元旦',\n    '2-14': '情人节',\n    '3-8': '妇女节',\n    '3-12': '植树节',\n    '4-4': '清明',\n    '4-5': '清明',\n    '5-1': '劳动节',\n    '5-4': '青年节',\n    '6-1': '儿童节',\n    '7-1': '建党节',\n    '8-1': '建军节',\n    '9-9': '重阳',\n    '10-1': '国庆节',\n    '10-2': '国庆节',\n    '10-3': '国庆节',\n    '11-11': '双十一',\n    '12-25': '圣诞节'\n};\n\n// 农历节日\nconst LUNAR_FESTIVALS = {\n    '1-1': '春节',\n    '1-15': '元宵',\n    '5-5': '端午',\n    '7-7': '七夕',\n    '7-15': '中元',\n    '8-15': '中秋',\n    '9-9': '重阳',\n    '12-30': '除夕',\n    '12-29': '除夕'\n};\n\nfunction getDayInfo(year, month, day) {\n    const solarKey = `${month}-${day}`;\n    if (HOLIDAYS[solarKey]) return { text: HOLIDAYS[solarKey], isHoliday: true };\n    \n    const termKey = solarKey;\n    if (SOLAR_TERMS[termKey]) return { text: SOLAR_TERMS[termKey], isHoliday: false };\n    \n    try {\n        const lunar = solarToLunar(year, month, day);\n        const lunarKey = `${lunar.month}-${lunar.day}`;\n        if (LUNAR_FESTIVALS[lunarKey]) return { text: LUNAR_FESTIVALS[lunarKey], isHoliday: true };\n        return { text: lunar.dayStr, isHoliday: false };\n    } catch (e) {\n        return { text: '', isHoliday: false };\n    }\n}\n\n// 全局状态\nif (!window._v15CalState) {\n    window._v15CalState = {\n        year: new Date().getFullYear(),\n        month: new Date().getMonth() + 1\n    };\n}\n\nasync function render(content) {\n    const state = window._v15CalState;\n    content.empty();\n\n    const container = content.createDiv({ cls: 'calendar-container' });\n\n    const today = new Date();\n    const todayY = today.getFullYear();\n    const todayM = today.getMonth() + 1;\n    const todayD = today.getDate();\n\n    let { year, month } = state;\n\n    // 天干地支年份信息\n    try {\n        const lunarYear = solarToLunar(year, month, 1);\n        const yearInfo = container.createDiv({ cls: 'calendar-year-info' });\n        yearInfo.textContent = `${lunarYear.ganZhi}年 · ${lunarYear.animal}年`;\n    } catch (e) {}\n\n    // 导航栏\n    const nav = container.createDiv({ cls: 'calendar-nav' });\n    const prevBtn = nav.createEl('button', { cls: 'calendar-nav-btn', text: '‹' });\n    const titleEl = nav.createEl('span', {\n        cls: 'calendar-title',\n        text: `${year}年${month}月`\n    });\n    const todayBtn = nav.createEl('button', { cls: 'calendar-nav-btn', text: '今', attr: { style: 'font-size: 11px; width: 28px;' } });\n    const nextBtn = nav.createEl('button', { cls: 'calendar-nav-btn', text: '›' });\n\n    prevBtn.addEventListener('click', () => {\n        state.month--;\n        if (state.month < 1) { state.month = 12; state.year--; }\n        render(content);\n    });\n    nextBtn.addEventListener('click', () => {\n        state.month++;\n        if (state.month > 12) { state.month = 1; state.year++; }\n        render(content);\n    });\n    todayBtn.addEventListener('click', () => {\n        state.year = todayY;\n        state.month = todayM;\n        render(content);\n    });\n\n    // 星期头\n    const weekdays = container.createDiv({ cls: 'calendar-weekdays' });\n    ['日','一','二','三','四','五','六'].forEach(d => {\n        weekdays.createEl('div', { cls: 'calendar-weekday', text: d });\n    });\n\n    // 构建日期格子\n    const grid = container.createDiv({ cls: 'calendar-grid' });\n    const firstDay = new Date(year, month - 1, 1).getDay();\n    const daysInMonth = new Date(year, month, 0).getDate();\n    const daysInPrevMonth = new Date(year, month - 1, 0).getDate();\n\n    // 补充上月\n    for (let i = firstDay - 1; i >= 0; i--) {\n        const d = daysInPrevMonth - i;\n        const cell = grid.createDiv({ cls: 'calendar-day other-month' });\n        cell.createEl('div', { cls: 'calendar-day-num', text: String(d) });\n        cell.createEl('div', { cls: 'calendar-lunar', text: '' });\n    }\n\n    // 当月日期\n    for (let d = 1; d <= daysInMonth; d++) {\n        const isToday = year === todayY && month === todayM && d === todayD;\n        const dow = new Date(year, month - 1, d).getDay();\n        const isWeekend = dow === 0 || dow === 6;\n\n        let cls = 'calendar-day';\n        if (isToday) cls += ' today';\n        if (isWeekend) cls += ' weekend';\n\n        const cell = grid.createDiv({ cls });\n        cell.createEl('div', { cls: 'calendar-day-num', text: String(d) });\n\n        // 农历 / 节日 / 节气\n        const showLunar = settings.showLunar !== false;\n        const showHoliday = settings.showHoliday !== false;\n\n        if (showLunar || showHoliday) {\n            const dayInfo = getDayInfo(year, month, d);\n            const lunarEl = cell.createEl('div', {\n                cls: dayInfo.isHoliday ? 'calendar-holiday' : 'calendar-lunar',\n                text: dayInfo.text\n            });\n        }\n    }\n\n    // 补充下月\n    const totalCells = firstDay + daysInMonth;\n    const remaining = totalCells % 7 === 0 ? 0 : 7 - (totalCells % 7);\n    for (let d = 1; d <= remaining; d++) {\n        const cell = grid.createDiv({ cls: 'calendar-day other-month' });\n        cell.createEl('div', { cls: 'calendar-day-num', text: String(d) });\n        cell.createEl('div', { cls: 'calendar-lunar', text: '' });\n    }\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    const { Setting } = require('obsidian');\n\n    containerEl.createEl('h3', { text: t('mod.calendar.settings.title') });\n\n    new Setting(containerEl)\n        .setName(t('mod.calendar.settings.lunar'))\n        .setDesc(t('mod.calendar.settings.lunarDesc'))\n        .addToggle(t => {\n            t.setValue(settings.showLunar !== false)\n                .onChange(async (v) => {\n                    settings.showLunar = v;\n                    await saveCallback();\n                });\n        });\n\n    new Setting(containerEl)\n        .setName(t('mod.calendar.settings.holiday'))\n        .setDesc(t('mod.calendar.settings.holidayDesc'))\n        .addToggle(t => {\n            t.setValue(settings.showHoliday !== false)\n                .onChange(async (v) => {\n                    settings.showHoliday = v;\n                    await saveCallback();\n                });\n        });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "code-editor": "// code-editor 模块 - 代码编辑器\n// 源插件: ace-code-editor\n// 核心功能保留: 代码文件浏览+编辑+保存\nconst id = 'code-editor';\nconst title = t('mod.codeEditor');\nconst icon = '💻';\n\nconst defaultSettings = {\n    fontSize: 14,\n    tabSize: 4,\n    theme: 'monokai',\n    showLineNumbers: true\n};\n\nconst styles = `\n.ce-wrap { padding: 8px 10px; display: flex; flex-direction: column; height: 100%; }\n.ce-toolbar { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; flex-shrink: 0; flex-wrap: wrap; }\n.ce-toolbar button { padding: 3px 8px; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 11px; cursor: pointer; white-space: nowrap; }\n.ce-toolbar button:hover { background: var(--background-modifier-hover); }\n.ce-toolbar button.primary { background: var(--v6-primary); color: white; border-color: var(--v6-primary); }\n.ce-toolbar button.primary:hover { opacity: 0.85; }\n.ce-toolbar .ce-spacer { flex: 1; }\n.ce-filelist { max-height: 120px; overflow-y: auto; margin-bottom: 4px; flex-shrink: 0; border: 1px solid var(--background-modifier-border); border-radius: 4px; }\n.ce-filelist.hidden { display: none; }\n.ce-file-item { display: flex; align-items: center; justify-content: space-between; padding: 3px 8px; cursor: pointer; font-size: 11px; transition: background 0.1s; }\n.ce-file-item:hover { background: var(--background-modifier-hover); }\n.ce-file-item.selected { background: rgba(var(--v6-primary-rgb, 232,149,109), 0.15); color: var(--v6-primary); }\n.ce-file-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.ce-file-type { color: var(--text-faint); font-size: 10px; flex-shrink: 0; margin-left: 8px; background: var(--background-modifier-form-field); padding: 1px 6px; border-radius: 8px; }\n.ce-editor-wrap { flex: 1; border: 1px solid var(--background-modifier-border); border-radius: 4px; overflow: hidden; position: relative; display: flex; }\n.ce-linenums { background: var(--background-secondary); color: var(--text-faint); padding: 8px 6px 8px 10px; font-family: var(--font-monospace); font-size: 13px; line-height: 1.5; text-align: right; user-select: none; overflow: hidden; border-right: 1px solid var(--background-modifier-border); min-width: 30px; }\n.ce-linenums div { min-height: 19.5px; }\n.ce-textarea { flex: 1; border: none; padding: 8px 10px; font-family: var(--font-monospace); font-size: 13px; line-height: 1.5; tab-size: 4; resize: none; outline: none; background: var(--background-primary); color: var(--text-normal); white-space: pre; overflow-wrap: normal; overflow-x: auto; }\n.ce-textarea:focus { background: var(--background-primary); }\n.ce-empty { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; text-align: center; pointer-events: none; }\n.ce-statusbar { display: flex; align-items: center; justify-content: space-between; padding: 3px 6px; font-size: 10px; color: var(--text-faint); background: var(--background-secondary); border-top: 1px solid var(--background-modifier-border); flex-shrink: 0; }\n.ce-statusbar span { margin-right: 10px; }\n.ce-modified { color: #ff9800; }\n.ce-saved { color: #4caf50; }\n`;\n\n// 代码文件扩展名映射（★ 不含 md，排除笔记文件）\nvar CODE_EXTENSIONS = {\n    js: 'JavaScript',    ts: 'TypeScript',   jsx: 'React JSX',\n    tsx: 'React TSX',    py: 'Python',        rb: 'Ruby',\n    java: 'Java',        go: 'Go',            rs: 'Rust',\n    c: 'C',              cpp: 'C++',          cs: 'C#',\n    php: 'PHP',          swift: 'Swift',      kt: 'Kotlin',\n    css: 'CSS',          scss: 'SCSS',        less: 'Less',\n    html: 'HTML',        htm: 'HTML',         xml: 'XML',\n    json: 'JSON',        yaml: 'YAML',        yml: 'YAML',\n    sql: 'SQL',          sh: 'Shell',\n    bat: 'Batch',        ps1: 'PowerShell',   toml: 'TOML',\n    lua: 'Lua',          r: 'R',              dart: 'Dart',\n    vue: 'Vue',          svelte: 'Svelte',    ini: 'INI',\n    cfg: 'Config',       env: 'Env',          txt: 'Text'\n};\n\nasync function render(container) {\n    container.addClass('ce-wrap');\n    var s = settings;\n\n    // 工具栏\n    var toolbar = container.createDiv({ cls: 'ce-toolbar' });\n    var toggleBtn = toolbar.createEl('button', { text: t('mod.codeEditor.btn.files') });\n    var newBtn = toolbar.createEl('button', { text: t('mod.codeEditor.btn.new') });\n    var spacer = toolbar.createEl('span', { cls: 'ce-spacer' });\n    var saveBtn = toolbar.createEl('button', { text: t('mod.codeEditor.btn.save'), cls: 'primary' });\n    var openObsidianBtn = toolbar.createEl('button', { text: t('mod.codeEditor.btn.openObs') });\n\n    // 文件列表\n    var fileList = container.createDiv({ cls: 'ce-filelist' });\n\n    // 编辑器\n    var editorWrap = container.createDiv({ cls: 'ce-editor-wrap' });\n    var lineNums = editorWrap.createDiv({ cls: 'ce-linenums' });\n    var textarea = editorWrap.createEl('textarea', { cls: 'ce-textarea', attr: { spellcheck: 'false' } });\n    var emptyHint = editorWrap.createDiv({ cls: 'ce-empty' });\n    emptyHint.innerHTML = t('mod.codeEditor.hint') + '<br><small>' + t('mod.codeEditor.hintFormats') + '</small>';\n\n    // 状态栏\n    var statusbar = container.createDiv({ cls: 'ce-statusbar' });\n    var statusFile = statusbar.createSpan();\n    var statusType = statusbar.createSpan();\n    var statusSave = statusbar.createSpan();\n    var statusLines = statusbar.createSpan();\n\n    var currentFile = null;\n    var originalContent = '';\n    var modified = false;\n\n    // 扫描代码文件\n    function scanFiles() {\n        files = [];\n        app.vault.getFiles().forEach(function(f) {\n            var ext = f.extension.toLowerCase();\n            if (CODE_EXTENSIONS[ext]) {\n                files.push({ path: f.path, name: f.name, ext: ext, size: f.stat ? f.stat.size : 0 });\n            }\n        });\n        files.sort(function(a, b) { return a.name.localeCompare(b.name); });\n        renderFileList();\n    }\n\n    var files = [];\n\n    function renderFileList() {\n        fileList.innerHTML = '';\n        if (files.length === 0) {\n            fileList.createDiv({ text: t('mod.codeEditor.empty'), cls: 'ce-file-item' }).style.cssText = 'cursor:default;color:var(--text-muted);justify-content:center;';\n            return;\n        }\n        files.forEach(function(f) {\n            var item = fileList.createDiv({ cls: 'ce-file-item' });\n            item.createSpan({ text: f.name, cls: 'ce-file-name' });\n            item.createSpan({ text: CODE_EXTENSIONS[f.ext] || f.ext, cls: 'ce-file-type' });\n            if (currentFile && currentFile.path === f.path) item.addClass('selected');\n            // 安全点击\n            item.addEventListener('mousedown', function(evt) { evt.preventDefault(); evt.stopPropagation(); }, true);\n            item.addEventListener('click', function(evt) {\n                evt.preventDefault(); evt.stopPropagation(); evt.stopImmediatePropagation();\n                openFile(f);\n            }, true);\n        });\n    }\n\n    // 更新行号\n    function updateLineNumbers() {\n        var lines = textarea.value.split('\\n').length;\n        var currentLines = lineNums.children.length;\n        while (lineNums.children.length < lines) {\n            lineNums.createDiv({ text: String(lineNums.children.length + 1) });\n        }\n        while (lineNums.children.length > lines) {\n            lineNums.lastChild.remove();\n        }\n    }\n\n    // 同步滚动\n    textarea.addEventListener('scroll', function() {\n        lineNums.scrollTop = textarea.scrollTop;\n    });\n\n    // 监听修改\n    textarea.addEventListener('input', function() {\n        updateLineNumbers();\n        modified = (textarea.value !== originalContent);\n        updateStatus();\n    });\n\n    function updateStatus() {\n        statusLines.textContent = t('mod.codeEditor.line') + textarea.value.split('\\n').length;\n        if (modified) {\n            statusSave.textContent = t('mod.codeEditor.modified');\n            statusSave.className = 'ce-modified';\n        } else {\n            statusSave.textContent = t('mod.codeEditor.saved');\n            statusSave.className = 'ce-saved';\n        }\n    }\n\n    async function openFile(file) {\n        // 如果有未保存修改，确认\n        if (modified && currentFile) {\n            var confirmed = confirm(t('mod.codeEditor.confirmDiscard'));\n            if (!confirmed) return;\n        }\n\n        currentFile = file;\n        renderFileList();\n\n        try {\n            var fileObj = app.vault.getAbstractFileByPath(file.path);\n            if (!fileObj) { showError(t('mod.dataEditor.notFound')); return; }\n            var content = await app.vault.read(fileObj);\n            textarea.value = content;\n            originalContent = content;\n            modified = false;\n\n            updateLineNumbers();\n            emptyHint.style.display = 'none';\n            textarea.style.display = '';\n            lineNums.style.display = '';\n\n            statusFile.textContent = file.name;\n            statusType.textContent = CODE_EXTENSIONS[file.ext] || file.ext;\n            updateStatus();\n        } catch (e) {\n            showError(t('mod.dataEditor.readFailed') + e.message);\n        }\n    }\n\n    async function saveFile() {\n        if (!currentFile) return;\n        if (!modified) {\n            new Notice(t('mod.codeEditor.nothingToSave'));\n            return;\n        }\n\n        try {\n            var fileObj = app.vault.getAbstractFileByPath(currentFile.path);\n            if (!fileObj) { new Notice('文件不存在: ' + currentFile.path); return; }\n            await app.vault.modify(fileObj, textarea.value);\n            originalContent = textarea.value;\n            modified = false;\n            updateStatus();\n            new Notice(t('mod.codeEditor.savedNotice') + currentFile.name);\n        } catch (e) {\n            new Notice(t('mod.codeEditor.saveFailed') + e.message);\n        }\n    }\n\n    function showError(msg) {\n        textarea.style.display = 'none';\n        lineNums.style.display = 'none';\n        emptyHint.style.display = '';\n        emptyHint.innerHTML = '<span style=\"color:var(--text-error)\">⚠ ' + msg + '</span>';\n        statusFile.textContent = '';\n        statusType.textContent = '';\n    }\n\n    // 事件\n    toggleBtn.addEventListener('click', function() { fileList.classList.toggle('hidden'); });\n    newBtn.addEventListener('click', function() { new Notice(t('mod.codeEditor.useObsidianNew')); });\n    saveBtn.addEventListener('click', saveFile);\n\n    openObsidianBtn.addEventListener('click', function() {\n        if (currentFile) {\n            app.workspace.openLinkText(currentFile.path, '', false);\n        } else {\n            new Notice(t('mod.codeEditor.selectFirst'));\n        }\n    });\n\n    // Ctrl+S 快捷键保存\n    textarea.addEventListener('keydown', function(e) {\n        if ((e.ctrlKey || e.metaKey) && e.key === 's') {\n            e.preventDefault();\n            saveFile();\n        }\n        // Tab缩进\n        if (e.key === 'Tab') {\n            e.preventDefault();\n            var start = textarea.selectionStart;\n            var end = textarea.selectionEnd;\n            var spaces = ' '.repeat(s.tabSize || 4);\n            textarea.value = textarea.value.substring(0, start) + spaces + textarea.value.substring(end);\n            textarea.selectionStart = textarea.selectionEnd = start + spaces.length;\n            updateLineNumbers();\n        }\n    });\n\n    // 延迟初始化\n    setTimeout(function() { scanFiles(); }, 700);\n    updateLineNumbers();\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.codeEditor.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.codeEditor.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "data-editor": "// data-editor 模块 - JSON/YAML/XML 数据文件查看+格式化+验证\n// 源插件: data-files-editor\n// 核心功能保留: 格式化预览 + JSON验证 + 文件浏览\nconst id = 'data-editor';\nconst title = t('mod.dataEditor');\nconst icon = '📋';\n\nconst defaultSettings = {\n    doLoadTxt: true,\n    doLoadXml: true,\n    doLoadJson: true,\n    doLoadYaml: true,\n    lineWrapping: true\n};\n\nconst styles = `\n.de-wrap { padding: 8px 10px; display: flex; flex-direction: column; height: 100%; }\n.de-toolbar { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; flex-shrink: 0; flex-wrap: wrap; }\n.de-toolbar button { padding: 3px 8px; border: 1px solid var(--background-modifier-border); border-radius: 3px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 11px; cursor: pointer; }\n.de-toolbar button:hover { background: var(--background-modifier-hover); }\n.de-toolbar button.primary { background: var(--v6-primary); color: white; border-color: var(--v6-primary); }\n.de-toolbar button.danger { color: var(--text-error); }\n.de-filelist { max-height: 100px; overflow-y: auto; margin-bottom: 4px; flex-shrink: 0; border: 1px solid var(--background-modifier-border); border-radius: 4px; }\n.de-filelist.hidden { display: none; }\n.de-file-item { display: flex; align-items: center; justify-content: space-between; padding: 3px 8px; cursor: pointer; font-size: 11px; transition: background 0.1s; }\n.de-file-item:hover { background: var(--background-modifier-hover); }\n.de-file-item.selected { background: rgba(var(--v6-primary-rgb, 232,149,109), 0.15); color: var(--v6-primary); }\n.de-file-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.de-file-type { color: var(--text-faint); font-size: 10px; flex-shrink: 0; margin-left: 8px; }\n.de-viewer { flex: 1; border: 1px solid var(--background-modifier-border); border-radius: 4px; overflow: auto; background: var(--background-primary); position: relative; }\n.de-viewer pre { margin: 0; padding: 10px 14px; font-family: var(--font-monospace); font-size: 12px; line-height: 1.6; white-space: pre-wrap; word-break: break-word; }\n.de-empty { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; text-align: center; }\n.de-statusbar { display: flex; align-items: center; justify-content: space-between; padding: 3px 6px; font-size: 10px; color: var(--text-faint); background: var(--background-secondary); border-top: 1px solid var(--background-modifier-border); flex-shrink: 0; }\n.de-statusbar .de-valid { color: #4caf50; }\n.de-statusbar .de-invalid { color: var(--text-error); }\n`;\n\n// JSON语法高亮（简单版）\nfunction highlightJSON(text) {\n    return text.replace(/(\"(?:[^\"\\\\]|\\\\.)*\")\\s*:/g, '<span style=\"color:#e06c75;\">$1</span>:')\n        .replace(/: (\".*?\"|true|false|null|\\d+(?:\\.\\d+)?)/g, ': <span style=\"color:#98c379;\">$1</span>')\n        .replace(/[{}[\\]]/g, '<span style=\"color:#61afef;\">$&</span>');\n}\n\nasync function render(container) {\n    container.addClass('de-wrap');\n\n    // 工具栏\n    var toolbar = container.createDiv({ cls: 'de-toolbar' });\n    var toggleBtn = toolbar.createEl('button', { text: t('mod.dataEditor.btn.files') });\n    var formatBtn = toolbar.createEl('button', { text: t('mod.dataEditor.btn.format'), cls: 'primary' });\n    var validateBtn = toolbar.createEl('button', { text: t('mod.dataEditor.btn.validate') });\n    var copyBtn = toolbar.createEl('button', { text: t('mod.dataEditor.btn.copy') });\n    var refreshBtn = toolbar.createEl('button', { text: t('mod.dataEditor.btn.refresh') });\n\n    // 文件列表\n    var fileList = container.createDiv({ cls: 'de-filelist' });\n\n    // 查看器\n    var viewer = container.createDiv({ cls: 'de-viewer' });\n    viewer.innerHTML = '<div class=\"de-empty\">' + t('mod.dataEditor.hint') + '<br><small>' + t('mod.dataEditor.hintFormats') + '</small></div>';\n\n    // 状态栏\n    var statusbar = container.createDiv({ cls: 'de-statusbar' });\n    var statusFile = statusbar.createSpan();\n    var statusValid = statusbar.createSpan();\n\n    var currentFile = null;\n    var currentContent = '';\n    var files = [];\n\n    function scanFiles() {\n        files = [];\n        app.vault.getFiles().forEach(function(f) {\n            var ext = f.extension.toLowerCase();\n            if (['json', 'yaml', 'yml', 'xml', 'txt', 'toml', 'ini'].indexOf(ext) >= 0) {\n                files.push({ path: f.path, name: f.name, ext: ext });\n            }\n        });\n        files.sort(function(a, b) { return a.name.localeCompare(b.name); });\n        renderFileList();\n    }\n\n    function renderFileList() {\n        fileList.innerHTML = '';\n        if (files.length === 0) {\n            fileList.createDiv({ text: t('mod.dataEditor.empty'), cls: 'de-file-item' }).style.cssText = 'cursor:default;color:var(--text-muted);justify-content:center;';\n            return;\n        }\n        files.forEach(function(f) {\n            var item = fileList.createDiv({ cls: 'de-file-item' });\n            item.createSpan({ text: f.name, cls: 'de-file-name' });\n            item.createSpan({ text: f.ext.toUpperCase(), cls: 'de-file-type' });\n            if (currentFile && currentFile.path === f.path) item.addClass('selected');\n            // 安全点击：阻止Obsidian事件系统\n            item.addEventListener('mousedown', function(evt) { evt.preventDefault(); evt.stopPropagation(); }, true);\n            item.addEventListener('click', function(evt) {\n                evt.preventDefault(); evt.stopPropagation(); evt.stopImmediatePropagation();\n                loadFile(f);\n            }, true);\n        });\n    }\n\n    async function loadFile(file) {\n        currentFile = file;\n        renderFileList();\n        try {\n            var fileObj = app.vault.getAbstractFileByPath(file.path);\n            if (!fileObj) { showError(t('mod.dataEditor.notFound')); return; }\n            currentContent = await app.vault.read(fileObj);\n            displayContent();\n            statusFile.textContent = file.name;\n            validateContent();\n            } catch (e) {\n            showError(t('mod.dataEditor.readFailed') + e.message);\n        }\n    }\n\n    function displayContent() {\n        viewer.innerHTML = '';\n        var pre = viewer.createEl('pre');\n\n        if (currentFile.ext === 'json') {\n            pre.innerHTML = highlightJSON(currentContent);\n        } else {\n            pre.textContent = currentContent;\n        }\n    }\n\n    function formatContent() {\n        if (!currentContent) return;\n\n        var formatted = '';\n        var ext = currentFile.ext;\n\n        if (ext === 'json') {\n            try { formatted = JSON.stringify(JSON.parse(currentContent), null, 2); }\n            catch(e) { new Notice(t('mod.dataEditor.jsonError') + e.message); return; }\n        } else if (ext === 'yaml' || ext === 'yml') {\n            // YAML简单格式化：无法完全解析，保持原样\n            new Notice(t('mod.dataEditor.yamlNotSupport'));\n            return;\n        } else if (ext === 'xml') {\n            // XML简单格式化\n            formatted = currentContent.replace(/></g, '>\\n<');\n            var indent = 0;\n            formatted = formatted.split('\\n').map(function(line) {\n                if (line.match(/<\\/\\w/)) indent = Math.max(0, indent - 1);\n                var result = '  '.repeat(indent) + line.trim();\n                if (line.match(/<\\w[^>]*[^/]>$/)) indent++;\n                return result;\n            }).join('\\n');\n        } else {\n            new Notice(t('mod.dataEditor.typeNotSupport'));\n            return;\n        }\n\n        currentContent = formatted;\n        displayContent();\n        validateContent();\n\n        // 自动保存格式化结果\n        if (currentFile && (ext === 'json' || ext === 'xml')) {\n            var fileObj = app.vault.getAbstractFileByPath(currentFile.path);\n            if (fileObj) {\n                app.vault.modify(fileObj, currentContent).then(function() {\n                    statusFile.textContent = currentFile.name + ' (已格式化并保存)';\n                }).catch(function() {\n                    new Notice(t('mod.dataEditor.autoSaveFailed'));\n                });\n            }\n        }\n    }\n\n    function validateContent() {\n        if (!currentContent || !currentFile) {\n            statusValid.textContent = '';\n            statusValid.className = '';\n            return;\n        }\n\n        var ext = currentFile.ext;\n        if (ext === 'json') {\n            try {\n                JSON.parse(currentContent);\n                statusValid.textContent = t('mod.dataEditor.jsonValid');\n                statusValid.className = 'de-valid';\n            } catch (e) {\n                statusValid.textContent = t('mod.dataEditor.jsonInvalid') + e.message;\n                statusValid.className = 'de-invalid';\n            }\n        } else if (ext === 'yaml' || ext === 'yml') {\n            statusValid.textContent = '(YAML)';\n            statusValid.className = '';\n        } else {\n            statusValid.textContent = '';\n        }\n    }\n\n    function showError(msg) {\n        viewer.innerHTML = '<div class=\"de-empty\" style=\"color:var(--text-error)\">⚠ ' + msg + '</div>';\n        statusFile.textContent = '';\n        statusValid.textContent = '';\n    }\n\n    // 事件\n    toggleBtn.addEventListener('click', function() { fileList.classList.toggle('hidden'); });\n    formatBtn.addEventListener('click', formatContent);\n    validateBtn.addEventListener('click', validateContent);\n    refreshBtn.addEventListener('click', function() { scanFiles(); if (currentFile) loadFile(currentFile); });\n    copyBtn.addEventListener('click', function() {\n        if (!currentContent) return;\n        navigator.clipboard.writeText(currentContent).then(function() {\n            new Notice(t('mod.dataEditor.copied'));\n        }).catch(function() {\n            new Notice(t('mod.dataEditor.copyFailed'));\n        });\n    });\n\n    // 延迟初始化\n    setTimeout(function() { scanFiles(); }, 900);\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.dataEditor.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.dataEditor.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "directory": "/**\n * 目录模块 V15\n * 格式：V14（含 id/styles/renderSettings）\n * 功能：树形目录，折叠/展开，文件图标，点击打开\n * 特性：展开状态持久化到 settings.expandedNodes（使用 child.path 作为 key）\n */\nconst id = 'directory';\nconst title = t('mod.directory');\nconst icon = '📂';\n\nconst defaultSettings = {\n    folders: [],\n    expandedNodes: []\n};\n\nconst styles = `\n.dir-tree { padding: 4px 0; }\n.dir-root { margin-bottom: 8px; }\n.dir-root-node {\n    display: flex;\n    align-items: center;\n    gap: 6px;\n    padding: 6px 8px;\n    border-radius: 6px;\n    cursor: default;\n    font-weight: 600;\n    font-size: 13px;\n    color: var(--text-normal);\n    background: var(--background-modifier-form-field);\n}\n.dir-root-label { flex: 1; }\n.dir-count {\n    font-size: 10px;\n    color: var(--text-muted);\n    background: var(--background-secondary);\n    padding: 1px 6px;\n    border-radius: 10px;\n}\n.dir-node { }\n.dir-node-header {\n    display: flex;\n    align-items: center;\n    gap: 4px;\n    padding: 3px 6px;\n    border-radius: 4px;\n    cursor: pointer;\n    font-size: 12px;\n    color: var(--text-normal);\n    transition: background 0.15s;\n}\n.dir-node-header:hover {\n    background: var(--background-modifier-hover);\n}\n.dir-toggle {\n    width: 14px;\n    text-align: center;\n    font-size: 9px;\n    color: var(--text-muted);\n    cursor: pointer;\n    flex-shrink: 0;\n}\n.dir-icon { flex-shrink: 0; }\n.dir-label { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.dir-children { padding-left: 14px; }\n.dir-children.collapsed { display: none; }\n.dir-empty {\n    text-align: center;\n    padding: 24px;\n    color: var(--text-muted);\n    font-size: 13px;\n}\n`;\n\nconst FILE_ICONS = {\n    'md': '📝', 'markdown': '📝',\n    'png': '🖼️', 'jpg': '🖼️', 'jpeg': '🖼️', 'gif': '🖼️', 'webp': '🖼️', 'svg': '🖼️',\n    'pdf': '📄',\n    'doc': '📘', 'docx': '📘',\n    'xls': '📗', 'xlsx': '📗',\n    'ppt': '📙', 'pptx': '📙',\n    'mp3': '🎵', 'wav': '🎵', 'flac': '🎵',\n    'mp4': '🎬', 'mov': '🎬', 'mkv': '🎬',\n    'zip': '📦', 'rar': '📦', '7z': '📦',\n    'txt': '📃', 'csv': '📊', 'json': '🔧', 'js': '🔧', 'ts': '🔧', 'py': '🐍'\n};\n\nfunction getFileIcon(file) {\n    const ext = (file.extension || '').toLowerCase();\n    return FILE_ICONS[ext] || '📄';\n}\n\nfunction countFiles(folder) {\n    if (!folder.children) return 0;\n    let count = 0;\n    folder.children.forEach(child => {\n        count += child.children ? countFiles(child) : 1;\n    });\n    return count;\n}\n\nfunction renderFolder(container, folder, saveCallback) {\n    if (!folder.children) return;\n\n    const sorted = [...folder.children].sort((a, b) => {\n        if (a.children && !b.children) return -1;\n        if (!a.children && b.children) return 1;\n        return a.name.localeCompare(b.name, 'zh-CN');\n    });\n\n    sorted.forEach(child => {\n        const node = container.createDiv({ cls: 'dir-node' });\n\n        if (child.children !== undefined) {\n            // === 子文件夹 ===\n            // 用 child.path 作为持久化 key（Obsidian 提供的完整路径，绝对可靠）\n            const nodePath = child.path;\n            const isExpanded = settings.expandedNodes && settings.expandedNodes.includes(nodePath);\n\n            const header = node.createDiv({ cls: 'dir-node-header' });\n            const toggle = header.createEl('span', { text: isExpanded ? '▼' : '▶', cls: 'dir-toggle' });\n            header.createEl('span', { text: '📁', cls: 'dir-icon' });\n            header.createEl('span', { text: child.name, cls: 'dir-label' });\n            const cnt = countFiles(child);\n            if (cnt > 0) header.createEl('span', { text: String(cnt), cls: 'dir-count' });\n\n            const childContainer = node.createDiv({ cls: 'dir-children' + (isExpanded ? '' : ' collapsed') });\n\n            // 若已展开，递归渲染子内容\n            if (isExpanded) {\n                renderFolder(childContainer, child, saveCallback);\n            }\n\n            header.addEventListener('click', async () => {\n                const nowCollapsed = !childContainer.hasClass('collapsed');\n                childContainer.toggleClass('collapsed', nowCollapsed);\n                toggle.textContent = nowCollapsed ? '▶' : '▼';\n\n                // 持久化展开状态\n                if (!settings.expandedNodes) settings.expandedNodes = [];\n                if (nowCollapsed) {\n                    settings.expandedNodes = settings.expandedNodes.filter(p => p !== nodePath);\n                } else {\n                    if (!settings.expandedNodes.includes(nodePath)) {\n                        settings.expandedNodes.push(nodePath);\n                    }\n                    // 展开时若子内容为空则渲染\n                    if (childContainer.childElementCount === 0) {\n                        renderFolder(childContainer, child, saveCallback);\n                    }\n                }\n\n                // 调试日志\n                console.log('[directory] 展开状态变更:', nodePath, nowCollapsed ? '折叠' : '展开', 'expandedNodes:', settings.expandedNodes);\n\n                try {\n                    await saveCallback();\n                    console.log('[directory] 保存成功');\n                } catch (e) {\n                    console.error('[directory] 保存失败:', e);\n                }\n            });\n        } else {\n            // === 文件 ===\n            const header = node.createDiv({ cls: 'dir-node-header' });\n            header.createEl('span', { cls: 'dir-toggle' }); // 占位\n            header.createEl('span', { text: getFileIcon(child), cls: 'dir-icon' });\n            header.createEl('span', { text: child.name, cls: 'dir-label' });\n\n            header.addEventListener('click', () => {\n                app.workspace.openLinkText(child.path, '', false);\n            });\n        }\n    });\n}\n\nasync function render(content) {\n    content.empty();\n\n    // 确保 expandedNodes 已初始化\n    if (!settings.expandedNodes) {\n        settings.expandedNodes = [];\n        console.log('[directory] 初始化 expandedNodes 为空数组');\n    }\n    console.log('[directory] 当前 expandedNodes:', settings.expandedNodes);\n\n    const container = content.createDiv({ cls: 'dir-tree' });\n    const folders = settings.folders || [];\n\n    if (folders.length === 0) {\n        container.createEl('div', {\n            cls: 'dir-empty',\n            text: t('mod.directory.empty')\n        });\n        return;\n    }\n\n    for (const folderPath of folders) {\n        const folder = app.vault.getAbstractFileByPath(folderPath);\n        if (!folder || folder.children === undefined) {\n            const errNode = container.createDiv({ cls: 'dir-root' });\n            const errHeader = errNode.createDiv({ cls: 'dir-root-node' });\n            errHeader.createEl('span', { text: '⚠️' });\n            errHeader.createEl('span', {\n                text: t('mod.directory.error.notFound') + folderPath,\n                cls: 'dir-root-label',\n                attr: { style: 'color: var(--text-muted);' }\n            });\n            continue;\n        }\n\n        const rootNode = container.createDiv({ cls: 'dir-root' });\n        const rootHeader = rootNode.createDiv({ cls: 'dir-root-node' });\n        rootHeader.createEl('span', { text: '📁' });\n        rootHeader.createEl('span', { text: folder.name || folderPath, cls: 'dir-root-label' });\n        const totalFiles = countFiles(folder);\n        rootHeader.createEl('span', { text: totalFiles + t('mod.directory.fileCount'), cls: 'dir-count' });\n\n        const childContainer = rootNode.createDiv({ cls: 'dir-children' });\n        renderFolder(childContainer, folder, async () => {\n            console.log('[directory] 调用 saveSettings...');\n            await plugin.saveSettings();\n        });\n    }\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    const { Setting } = require('obsidian');\n\n    containerEl.createEl('h3', { text: t('mod.directory.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.directory.settings.desc'),\n        attr: { style: 'font-size: 12px; color: var(--text-muted); margin: 0 0 8px;' }\n    });\n\n    // 初始化\n    if (!settings.folders) settings.folders = [];\n    if (!settings.expandedNodes) settings.expandedNodes = [];\n\n    // 添加文件夹\n    const addSetting = new Setting(containerEl)\n        .setName('添加文件夹')\n        .setDesc('输入文件夹路径后点击添加');\n\n    let tempPath = '';\n    addSetting.addText(t => {\n        t.setPlaceholder('例如：笔记/日记')\n            .onChange(v => { tempPath = v; });\n    });\n    addSetting.addButton(b => {\n        b.setButtonText('添加')\n            .setCta()\n            .onClick(async () => {\n                const path = tempPath.trim();\n                if (!path) return new Notice('路径不能为空');\n                if (settings.folders.includes(path)) return new Notice('已存在');\n                const folder = app.vault.getAbstractFileByPath(path);\n                if (!folder) return new Notice(t('mod.directory.error.notFound') + path);\n                settings.folders.push(path);\n                await saveCallback();\n                containerEl.querySelectorAll('.dir-path-setting').forEach(el => el.remove());\n                renderFolderList();\n            });\n    });\n\n    // 已有文件夹列表\n    const renderFolderList = () => {\n        if (!settings.folders || settings.folders.length === 0) return;\n        settings.folders.forEach((path, index) => {\n            const s = new Setting(containerEl)\n                .setName('📁 ' + path)\n                .addButton(b => {\n                    b.setButtonText('移除').setWarning()\n                        .onClick(async () => {\n                            settings.folders.splice(index, 1);\n                            // 清理该文件夹相关的展开记录\n                            if (settings.expandedNodes) {\n                                settings.expandedNodes = settings.expandedNodes.filter(p => !p.startsWith(path + '/'));\n                            }\n                            await saveCallback();\n                            containerEl.querySelectorAll('.dir-path-setting').forEach(el => el.remove());\n                            renderFolderList();\n                        });\n                });\n            s.settingEl.addClass('dir-path-setting');\n        });\n    };\n    renderFolderList();\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "doc-viewer": "// doc-viewer 模块 - Word/PDF文档查看器\n// 源插件: ViewItAll (word查看器)\n// 核心功能: mammoth.js渲染DOCX + iframe渲染PDF（面板内完成）\nconst id = 'doc-viewer';\nconst title = t('mod.docViewer');\nconst icon = '📄';\n\nconst defaultSettings = {\n    docxEnabled: true,\n    pdfEnabled: true,\n    defaultZoom: 100\n};\n\nconst styles = `\n.dv-wrap { padding: 8px 0; display: flex; flex-direction: column; height: 100%; }\n.dv-toolbar { display: flex; align-items: center; gap: 6px; padding: 0 10px; margin-bottom: 6px; flex-shrink: 0; flex-wrap: wrap; }\n.dv-toolbar button { padding: 3px 8px; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 11px; cursor: pointer; white-space: nowrap; }\n.dv-toolbar button:hover { background: var(--background-modifier-hover); }\n.dv-toolbar button.active { background: var(--v6-primary); color: white; border-color: var(--v6-primary); }\n.dv-filelist { max-height: 100px; overflow-y: auto; margin: 0 10px 6px; flex-shrink: 0; border: 1px solid var(--background-modifier-border); border-radius: 4px; }\n.dv-filelist.hidden { display: none; }\n.dv-file-item { display: flex; align-items: center; justify-content: space-between; padding: 4px 8px; cursor: pointer; font-size: 11px; transition: background 0.1s; border-radius: 2px; user-select: none; -webkit-user-select: none; }\n.dv-file-item:hover { background: var(--background-modifier-hover); }\n.dv-file-item.selected { background: rgba(var(--v6-primary-rgb, 232,149,109), 0.2); color: var(--v6-primary); }\n.dv-file-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.dv-file-type { color: var(--text-faint); font-size: 10px; flex-shrink: 0; margin-left: 8px; background: var(--background-modifier-form-field); padding: 1px 6px; border-radius: 8px; }\n.dv-viewer { flex: 1; min-height: 0; margin: 0 10px; border: 1px solid var(--background-modifier-border); border-radius: 4px; overflow: auto; background: var(--background-primary); position: relative; }\n.dv-viewer iframe { width: 100%; height: 100%; border: none; }\n.dv-content { padding: 16px 20px; font-size: 13px; line-height: 1.7; color: var(--text-normal); width: 100%; box-sizing: border-box; word-break: break-word; }\n.dv-content * { max-width: none !important; box-sizing: border-box !important; }\n.dv-content h1 { font-size: 24px; margin-top: 0; margin-bottom: 12px; }\n.dv-content h2 { font-size: 20px; margin-top: 16px; margin-bottom: 8px; }\n.dv-content h3 { font-size: 16px; margin-top: 12px; margin-bottom: 6px; }\n.dv-content p { margin: 0 0 8px; }\n.dv-content img { max-width: 100%; height: auto; }\n.dv-content table { border-collapse: collapse; width: 100%; margin: 12px 0; }\n.dv-content th, .dv-content td { border: 1px solid var(--background-modifier-border); padding: 6px 10px; text-align: left; font-size: 12px; }\n.dv-content th { background: var(--background-modifier-form-field); font-weight: 600; }\n.dv-content ul, .dv-content ol { padding-left: 24px; margin-bottom: 8px; }\n.dv-content blockquote { border-left: 3px solid var(--v6-primary); padding-left: 12px; margin: 8px 0; color: var(--text-muted); }\n.dv-empty { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; text-align: center; pointer-events: none; }\n.dv-loading { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; }\n.dv-statusbar { display: flex; align-items: center; justify-content: space-between; padding: 3px 10px; font-size: 10px; color: var(--text-faint); background: var(--background-secondary); border-top: 1px solid var(--background-modifier-border); flex-shrink: 0; margin-top: 4px; }\n`;\n\n// ============ mammoth.js 异步加载 ============\nvar _mammothLoaded = false;\nvar _mammothLib = null;\nvar _mammothLoading = false;\nvar _mammothWaiters = [];\n\nfunction getMammoth() {\n    if (_mammothLoaded) return Promise.resolve(_mammothLib);\n    return new Promise(function(resolve) {\n        _mammothWaiters.push(resolve);\n        if (!_mammothLoading) loadMammoth();\n    });\n}\n\nfunction loadMammoth() {\n    _mammothLoading = true;\n    try {\n        requestUrl({ url: 'https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js' })\n            .then(function(resp) {\n                try {\n                    var code = resp.text;\n                    // 用 eval+IIFE 确保隔离执行，同时返回 mammoth 对象\n                    var wrapped = '(function(){var module=undefined,exports=undefined,define=undefined;' + code + '\\\\nreturn mammoth;})()';\n                    _mammothLib = eval(wrapped);\n                    if (typeof _mammothLib !== 'object' || typeof _mammothLib.convertToHtml !== 'function') {\n                        _mammothLib = window.mammoth || null;\n                    }\n                    _mammothLoaded = true;\n                    _mammothLoading = false;\n                    console.log('[doc-viewer] mammoth 加载完成, convertToHtml:', typeof (_mammothLib && _mammothLib.convertToHtml));\n                    _mammothWaiters.forEach(function(w) { w(_mammothLib); });\n                    _mammothWaiters = [];\n                } catch(e) {\n                    console.error('mammoth 执行失败:', e);\n                    _mammothLoading = false;\n                    _mammothLib = null;\n                    _mammothLoaded = true;\n                    _mammothWaiters.forEach(function(w) { w(null); });\n                    _mammothWaiters = [];\n                }\n            })\n            .catch(function() {\n                _mammothLoading = false;\n                _mammothLib = null;\n                _mammothLoaded = true;\n                _mammothWaiters.forEach(function(w) { w(null); });\n                _mammothWaiters = [];\n            });\n    } catch(e) {\n        _mammothLoading = false;\n        _mammothLib = null;\n        _mammothLoaded = true;\n        _mammothWaiters.forEach(function(w) { w(null); });\n        _mammothWaiters = [];\n    }\n}\n\n// ============ 安全点击 ============\nfunction safeClick(el, handler) {\n    el.addEventListener('mousedown', function(evt) {\n        evt.preventDefault();\n        evt.stopPropagation();\n    }, true);\n    el.addEventListener('click', function(evt) {\n        evt.preventDefault();\n        evt.stopPropagation();\n        evt.stopImmediatePropagation();\n        handler(evt);\n    }, true);\n}\n\n// ============ 主渲染（懒加载）============\nasync function render(container) {\n    container.addClass('dv-wrap');\n    var s = settings;\n\n    var toolbar = container.createDiv({ cls: 'dv-toolbar' });\n    var toggleBtn = toolbar.createEl('button', { text: t('mod.docViewer.btn.files') });\n    var refreshBtn = toolbar.createEl('button', { text: t('mod.docViewer.btn.refresh') });\n    var loadBtn = toolbar.createEl('button', { text: t('mod.docViewer.btn.view'), attr: { style: 'background:var(--v6-primary);color:white;border-color:var(--v6-primary);' } });\n\n    var fileList = container.createDiv({ cls: 'dv-filelist' });\n    var viewer = container.createDiv({ cls: 'dv-viewer' });\n    viewer.innerHTML = '<div class=\"dv-empty\">' + t('mod.docViewer.hint') + '<br><small>' + t('mod.docViewer.hintFormats') + '</small></div>';\n\n    var statusbar = container.createDiv({ cls: 'dv-statusbar' });\n    var statusFile = statusbar.createSpan();\n    var statusPages = statusbar.createSpan();\n\n    var currentFile = null;\n    var files = [];\n\n    function scanFiles() {\n        files = [];\n        var allFiles = app.vault.getFiles();\n        for (var i = 0; i < allFiles.length; i++) {\n            var f = allFiles[i];\n            var ext = f.extension.toLowerCase();\n            if (ext === 'docx' || ext === 'pdf') {\n                files.push({ path: f.path, name: f.name, ext: ext, size: f.stat ? f.stat.size : 0 });\n            }\n        }\n        files.sort(function(a, b) { return a.name.localeCompare(b.name); });\n        renderFileList();\n    }\n\n    function renderFileList() {\n        fileList.innerHTML = '';\n        if (files.length === 0) {\n            var empty = document.createElement('div');\n            empty.className = 'dv-file-item';\n            empty.textContent = t('mod.docViewer.empty');\n            empty.style.cssText = 'cursor:default;color:var(--text-muted);justify-content:center;';\n            fileList.appendChild(empty);\n            return;\n        }\n        for (var i = 0; i < files.length; i++) {\n            (function(f) {\n                var item = document.createElement('div');\n                item.className = 'dv-file-item';\n                if (currentFile && currentFile.path === f.path) item.classList.add('selected');\n\n                var icon = f.ext === 'pdf' ? '📕' : '📘';\n                var nameSpan = document.createElement('span');\n                nameSpan.className = 'dv-file-name';\n                nameSpan.textContent = icon + ' ' + f.name;\n                item.appendChild(nameSpan);\n\n                var typeSpan = document.createElement('span');\n                typeSpan.className = 'dv-file-type';\n                typeSpan.textContent = f.ext.toUpperCase();\n                item.appendChild(typeSpan);\n\n                // 安全单击：选中\n                safeClick(item, function() {\n                    currentFile = f;\n                    renderFileList();\n                    statusFile.textContent = t('mod.docViewer.selected') + f.name;\n                });\n\n                // 双击加载\n                item.addEventListener('dblclick', function(evt) {\n                    evt.preventDefault();\n                    evt.stopPropagation();\n                    evt.stopImmediatePropagation();\n                    currentFile = f;\n                    renderFileList();\n                    loadFile(f);\n                }, true);\n\n                fileList.appendChild(item);\n            })(files[i]);\n        }\n    }\n\n    async function loadFile(file) {\n        currentFile = file;\n        renderFileList();\n        viewer.innerHTML = '<div class=\"dv-loading\">' + t('mod.docViewer.loading') + '</div>';\n        statusFile.textContent = file.name;\n        statusPages.textContent = '';\n\n        try {\n            var fileObj = app.vault.getAbstractFileByPath(file.path);\n            if (!fileObj) { showError(t('mod.docViewer.notFound') + file.path); return; }\n\n            if (file.ext === 'pdf') {\n                var data = await app.vault.readBinary(fileObj);\n                var blob = new Blob([data], { type: 'application/pdf' });\n                var url = URL.createObjectURL(blob);\n                viewer.innerHTML = '';\n                var iframe = document.createElement('iframe');\n                iframe.src = url;\n                viewer.appendChild(iframe);\n                statusPages.textContent = t('mod.docViewer.pdfReader');\n                setTimeout(function() { URL.revokeObjectURL(url); }, 120000);\n            } else if (file.ext === 'docx') {\n                var data = await app.vault.readBinary(fileObj);\n                var mammoth = await getMammoth();\n                if (!mammoth) {\n                    showError(t('mod.docViewer.error.mammoth') + '\\n' + t('mod.docViewer.error.network'));\n                    return;\n                }\n                var arrayBuffer;\n                if (data instanceof ArrayBuffer) arrayBuffer = data;\n                else if (data instanceof Uint8Array) arrayBuffer = data.buffer;\n                else if (data && data.buffer) arrayBuffer = data.buffer;\n                else arrayBuffer = new Uint8Array(data).buffer;\n\n                var result;\n                try {\n                    // 策略1：完整 HTML 转换\n                    result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer }, {\n                        styleMap: [\n                            \"p[style-name='Heading 1'] => h1:fresh\",\n                            \"p[style-name='Heading 2'] => h2:fresh\",\n                            \"p[style-name='Heading 3'] => h3:fresh\",\n                            \"p[style-name='Heading 4'] => h4:fresh\",\n                            \"r[style-name='Strong'] => strong\",\n                            \"r[style-name='Emphasis'] => em\"\n                        ]\n                    });\n                } catch(mamErr) {\n                    console.warn('[doc-viewer] mammoth HTML 转换失败，降级为纯文本:', String(mamErr).substring(0,120));\n                    try {\n                        // 策略2：纯文本提取\n                        var rawResult = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });\n                        result = { value: '<p>' + rawResult.value.replace(/</g, '&lt;').replace(/\\n/g, '</p><p>') + '</p>', messages: rawResult.messages };\n                    } catch(rawErr) {\n                        console.warn('[doc-viewer] mammoth extractRawText 也失败:', String(rawErr).substring(0,120));\n                        throw new Error(t('mod.docViewer.error.corrupt') + '\\n（' + String(rawErr).substring(0,80) + '）\\n' + t('mod.docViewer.error.corruptHint'));\n                    }\n                }\n                viewer.innerHTML = '<div class=\"dv-content\">' + result.value + '</div>';\n                statusPages.textContent = t('mod.docViewer.docxRendered');\n                if (result.messages && result.messages.length > 0) {\n                    console.log('mammoth 警告:', result.messages);\n                }\n            }\n        } catch (e) {\n            showError(t('mod.docViewer.error.load') + (e.message || e));\n            console.error('doc-viewer loadFile error:', e);\n        }\n    }\n\n    function showError(msg) {\n        viewer.innerHTML = '<div class=\"dv-empty\" style=\"color:var(--text-error);white-space:pre-line;\">⚠ ' + msg.replace(/</g, '&lt;') + '</div>';\n        statusPages.textContent = '';\n    }\n\n    toggleBtn.addEventListener('click', function() { fileList.classList.toggle('hidden'); });\n    refreshBtn.addEventListener('click', function() { scanFiles(); });\n    loadBtn.addEventListener('click', function() {\n        if (currentFile) loadFile(currentFile);\n        else if (files.length > 0) { currentFile = files[0]; renderFileList(); loadFile(files[0]); }\n    });\n\n    // 懒初始化\n    setTimeout(function() {\n        scanFiles();\n        getMammoth().catch(function(){});\n    }, 500);\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.docViewer.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.docViewer.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n    containerEl.createEl('p', {\n        text: t('mod.docViewer.settings.hint'),\n        attr: { style: 'color:#4caf50;font-size:11px;' }\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "excel-to-markdown": "// excel-to-markdown 模块 - Excel粘贴自动转Markdown表格\n// 源插件: obsidian-excel-to-markdown-table\n// 核心功能保留: 全局拦截笔记编辑器粘贴事件，自动检测Tab分隔数据转表格\nconst id = 'excel-to-markdown';\nconst title = t('mod.excelToMd');\nconst icon = '📊';\n\nconst defaultSettings = {\n    enabledAutoConvert: true\n};\n\nconst styles = `\n.excel-md-panel { padding: 10px 12px; display: flex; flex-direction: column; height: 100%; }\n.excel-md-hint { font-size: 11px; color: var(--text-muted); margin-bottom: 8px; line-height: 1.5; }\n.excel-md-hint .badge { display: inline-block; background: #4caf50; color: white; padding: 1px 8px; border-radius: 10px; font-size: 10px; font-weight: 600; margin-right: 4px; }\n.excel-md-hint .badge.off { background: var(--text-faint); }\n.excel-md-toggle-row { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }\n.excel-md-switch { position: relative; width: 36px; height: 20px; background: var(--background-modifier-border); border-radius: 10px; cursor: pointer; transition: background 0.2s; flex-shrink: 0; }\n.excel-md-switch.on { background: #4caf50; }\n.excel-md-switch-knob { position: absolute; top: 2px; left: 2px; width: 16px; height: 16px; background: white; border-radius: 50%; transition: left 0.2s; }\n.excel-md-switch.on .excel-md-switch-knob { left: 18px; }\n.excel-md-switch-label { font-size: 12px; color: var(--text-normal); }\n.excel-md-textarea { width: 100%; min-height: 60px; max-height: 100px; border: 1px solid var(--background-modifier-border); border-radius: 6px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 12px; font-family: var(--font-monospace); resize: vertical; padding: 8px; outline: none; box-sizing: border-box; }\n.excel-md-textarea:focus { border-color: var(--v6-primary); }\n.excel-md-output { flex: 1; min-height: 40px; max-height: 160px; overflow: auto; padding: 10px; border: 1px solid var(--background-modifier-border); border-radius: 6px; background: var(--background-modifier-form-field); font-family: var(--font-monospace); font-size: 11px; white-space: pre-wrap; color: var(--text-normal); margin-top: 8px; word-break: break-all; tab-size: 4; }\n.excel-md-output.empty { color: var(--text-faint); font-family: inherit; }\n.excel-md-btn-row { display: flex; gap: 6px; margin-top: 8px; flex-wrap: wrap; }\n.excel-md-btn { padding: 4px 10px; border: none; border-radius: 4px; background: var(--v6-primary); color: white; cursor: pointer; font-size: 11px; transition: opacity 0.15s; }\n.excel-md-btn:hover { opacity: 0.85; }\n.excel-md-btn.secondary { background: var(--background-modifier-form-field); color: var(--text-normal); border: 1px solid var(--background-modifier-border); }\n.excel-md-btn.success { background: #4caf50; color: white; }\n.excel-md-status { font-size: 10px; color: var(--text-muted); margin-top: 6px; min-height: 14px; }\n.excel-md-status.success { color: #4caf50; }\n.excel-md-status.error { color: var(--text-error); }\n.excel-md-table-preview { margin-top: 6px; overflow-x: auto; max-height: 100px; }\n.excel-md-table-preview table { border-collapse: collapse; font-size: 11px; width: 100%; }\n.excel-md-table-preview th, .excel-md-table-preview td { border: 1px solid var(--background-modifier-border); padding: 4px 8px; text-align: left; }\n.excel-md-table-preview th { background: var(--background-modifier-form-field); font-weight: 600; }\n`;\n\n// ============ 全局粘贴事件引用（用于清理） ============\nvar _pasteEventRef = null;\nvar _autoConvertEnabled = true;\n\n// ============ 核心转换引擎 ============\n\n// 单元格内换行处理（原插件 replaceIntraCellNewline）\nfunction replaceIntraCellNewline(data) {\n    return data.replace(/\"([^\\t]*(?<=[^\\r])\\n[^\\t]*)\"/g, function(match) {\n        return match.slice(1, -1).replace(/\"\"/g, '\"').replace(/\\n/g, '<br/>');\n    });\n}\n\n// 列对齐\nvar COL_ALIGN_REGEX = /^(\\^[lcr])/i;\n\nfunction getColumnWidthsAndAlignments(rows) {\n    var colAlignments = [];\n    var columnWidths = rows[0].map(function(col, i) {\n        var align = 'l';\n        var m = col.match(COL_ALIGN_REGEX);\n        if (m) {\n            align = m[1][1].toLowerCase();\n            rows[0][i] = col.replace(COL_ALIGN_REGEX, '');\n        }\n        colAlignments.push(align);\n        return Math.max.apply(null, rows.map(function(r) { return String(r[i] || '').length; }));\n    });\n    return { columnWidths: columnWidths, colAlignments: colAlignments };\n}\n\n// 主转换函数\nfunction excelToMarkdown(rawData) {\n    var data = rawData.trim();\n    if (!data) return null;\n\n    data = replaceIntraCellNewline(data);\n\n    var rows = data.split(/[\\n\\u0085\\u2028\\u2029]|\\r\\n?/g).map(function(r) {\n        return r.split('\\t');\n    });\n\n    if (!rows[0] || rows[0].length < 2) return null;\n\n    rows = rows.filter(function(r) { return r.some(function(c) { return c.trim(); }); });\n    if (rows.length === 0) return null;\n\n    var sizes = getColumnWidthsAndAlignments(rows);\n    var colWidths = sizes.columnWidths;\n    var colAlignments = sizes.colAlignments;\n\n    var mdRows = rows.map(function(row) {\n        return '| ' + row.map(function(col, i) {\n            return String(col).replace(/\\|/g, '\\\\|') + ' '.repeat(Math.max(0, colWidths[i] - String(col).length + 1));\n        }).join(' | ') + ' |';\n    });\n\n    var ALIGN_MAP = { l: ' ', r: ':', c: ':' };\n    var ALIGN_POST = { l: ' ', r: '', c: ':' };\n    var alignRow = '|' + colWidths.map(function(w, i) {\n        var a = colAlignments[i] || 'l';\n        return ALIGN_MAP[a] + '-'.repeat(w + 2) + ALIGN_POST[a];\n    }).join('|') + '|';\n    mdRows.splice(1, 0, alignRow);\n\n    return mdRows.join('\\n');\n}\n\n// 预览：解析Markdown表格为HTML\nfunction markdownToPreviewTable(md) {\n    var lines = md.trim().split('\\n').filter(function(l) { return l.trim(); });\n    if (lines.length < 2) return null;\n\n    var parseRow = function(line) {\n        return line.split('|').slice(1, -1).map(function(c) { return c.trim(); });\n    };\n\n    var header = parseRow(lines[0]);\n    var alignLine = lines[1];\n    var bodyLines;\n    if (/^[\\s\\|\\:\\-]+$/.test(alignLine)) {\n        bodyLines = lines.slice(2);\n    } else {\n        bodyLines = lines.slice(1);\n    }\n\n    var html = '<table>';\n    html += '<thead><tr>' + header.map(function(h) { return '<th>' + h + '</th>'; }).join('') + '</tr></thead>';\n    html += '<tbody>';\n    bodyLines.forEach(function(line) {\n        var cells = parseRow(line);\n        html += '<tr>' + cells.map(function(c) { return '<td>' + c + '</td>'; }).join('') + '</tr>';\n    });\n    html += '</tbody></table>';\n    return html;\n}\n\n// 判断剪贴板数据是否为Excel表格（包含Tab分隔符）\nfunction isExcelData(text) {\n    if (!text) return false;\n    var trimmed = text.trim();\n    if (!trimmed) return false;\n    // 必须包含Tab\n    if (trimmed.indexOf('\\t') === -1) return false;\n    // 第一行至少2列\n    var firstLine = trimmed.split(/[\\n\\r]/)[0] || '';\n    return firstLine.split('\\t').length >= 2;\n}\n\n// ============ 全局粘贴拦截器 ============\n// 这是核心功能：在 Obsidian 笔记编辑器中 Ctrl+V 粘贴 Excel 数据时，\n// 自动拦截并转换为 Markdown 表格格式\n\nfunction registerGlobalPasteHandler() {\n    // 先清理旧的\n    unregisterGlobalPasteHandler();\n\n    try {\n        _pasteEventRef = app.workspace.on('editor-paste', function(evt, editor, view) {\n            if (!_autoConvertEnabled) return;\n\n            // 获取剪贴板文本\n            var clipboardData = evt.clipboardData;\n            if (!clipboardData) return;\n\n            var text = clipboardData.getData('text/plain');\n            if (!isExcelData(text)) return;\n\n            // 转换为Markdown表格\n            var mdTable = excelToMarkdown(text);\n            if (!mdTable) return;\n\n            // 阻止默认粘贴行为\n            evt.preventDefault();\n\n            // 用转换后的Markdown表格替换选区\n            try {\n                editor.replaceSelection(mdTable + '\\n');\n                // 通知用户\n                new Notice('✓ Excel数据已自动转为Markdown表格');\n            } catch(e) {\n                console.error('Excel->Markdown插入失败:', e);\n            }\n        });\n    } catch(e) {\n        console.error('注册全局粘贴处理器失败:', e);\n    }\n}\n\nfunction unregisterGlobalPasteHandler() {\n    if (_pasteEventRef && app && app.workspace) {\n        try {\n            app.workspace.offref(_pasteEventRef);\n        } catch(e) {}\n        _pasteEventRef = null;\n    }\n}\n\n// ============ 渲染 ============\n\nasync function render(container) {\n    container.addClass('excel-md-panel');\n\n    // 同步状态\n    _autoConvertEnabled = settings.enabledAutoConvert !== false;\n\n    // 注册全局粘贴拦截器（每次渲染都确保已注册）\n    registerGlobalPasteHandler();\n\n    // 状态指示\n    var hint = container.createDiv({ cls: 'excel-md-hint' });\n    updateHint(hint);\n\n    // 开关行\n    var toggleRow = container.createDiv({ cls: 'excel-md-toggle-row' });\n    var switchEl = toggleRow.createDiv({ cls: 'excel-md-switch' + (_autoConvertEnabled ? ' on' : '') });\n    var knob = switchEl.createDiv({ cls: 'excel-md-switch-knob' });\n    var switchLabel = toggleRow.createSpan({ cls: 'excel-md-switch-label', text: _autoConvertEnabled ? t('mod.excelToMd.autoOn') : t('mod.excelToMd.autoOff') });\n\n    switchEl.addEventListener('click', function() {\n        _autoConvertEnabled = !_autoConvertEnabled;\n        settings.enabledAutoConvert = _autoConvertEnabled;\n        if (_autoConvertEnabled) {\n            switchEl.addClass('on');\n            switchLabel.textContent = t('mod.excelToMd.autoOn');\n            registerGlobalPasteHandler();\n        } else {\n            switchEl.removeClass('on');\n            switchLabel.textContent = t('mod.excelToMd.autoOff');\n            unregisterGlobalPasteHandler();\n        }\n        updateHint(hint);\n        saveCallback();\n    });\n\n    // 输入文本区（手动粘贴到模块内也可以）\n    var textarea = container.createEl('textarea', {\n        cls: 'excel-md-textarea',\n        attr: { placeholder: t('mod.excelToMd.placeholder') }\n    });\n\n    // 输出预览区\n    var output = container.createDiv({ cls: 'excel-md-output empty' });\n    output.textContent = t('mod.excelToMd.autoConvertHint');\n\n    // 表格预览区\n    var tablePreview = container.createDiv({ cls: 'excel-md-table-preview' });\n\n    // 按钮行\n    var btnRow = container.createDiv({ cls: 'excel-md-btn-row' });\n    var insertBtn = btnRow.createEl('button', { text: t('mod.excelToMd.btn.insert'), cls: 'excel-md-btn' });\n    var copyBtn = btnRow.createEl('button', { text: t('mod.excelToMd.btn.copy'), cls: 'excel-md-btn secondary' });\n    var clearBtn = btnRow.createEl('button', { text: t('mod.excelToMd.btn.clear'), cls: 'excel-md-btn secondary' });\n\n    // 状态行\n    var status = container.createDiv({ cls: 'excel-md-status' });\n\n    var currentMarkdown = '';\n\n    function updateHint(hintEl) {\n        if (_autoConvertEnabled) {\n            hintEl.innerHTML = '<span class=\"badge\">● 已激活</span> ' + t('mod.excelToMd.autoConvertHint');\n        } else {\n            hintEl.innerHTML = '<span class=\"badge off\">○ 已关闭</span> 仅手动模式：在此面板内粘贴来转换';\n        }\n    }\n\n    function convertAndShow(raw) {\n        if (!raw.trim()) {\n            output.textContent = t('mod.excelToMd.waiting');\n            output.addClass('empty');\n            tablePreview.innerHTML = '';\n            currentMarkdown = '';\n            status.textContent = '';\n            return;\n        }\n\n        var result = excelToMarkdown(raw);\n        if (result) {\n            output.textContent = result;\n            output.removeClass('empty');\n            currentMarkdown = result;\n            var html = markdownToPreviewTable(result);\n            tablePreview.innerHTML = html || '';\n            status.textContent = t('mod.excelToMd.detected') + result.split('\\n').length + t('mod.excelToMd.detectedRows');\n            status.className = 'excel-md-status success';\n        } else {\n            output.textContent = t('mod.excelToMd.notDetected');\n            output.addClass('empty');\n            tablePreview.innerHTML = '';\n            currentMarkdown = '';\n            status.textContent = t('mod.excelToMd.notDetected');\n            status.className = 'excel-md-status error';\n        }\n    }\n\n    // 面板内粘贴事件\n    textarea.addEventListener('paste', function() {\n        setTimeout(function() { convertAndShow(textarea.value); }, 50);\n    });\n\n    textarea.addEventListener('input', function() {\n        convertAndShow(textarea.value);\n    });\n\n    // 插入当前笔记\n    insertBtn.addEventListener('click', function() {\n        if (!currentMarkdown) {\n            status.textContent = t('mod.excelToMd.nothingToCopy');\n            status.className = 'excel-md-status error';\n            return;\n        }\n        try {\n            var editor = app.workspace.activeEditor;\n            if (editor && editor.editor) {\n                editor.editor.replaceSelection(currentMarkdown + '\\n');\n                status.textContent = t('mod.excelToMd.inserted');\n                status.className = 'excel-md-status success';\n            } else {\n                var leaf = app.workspace.activeLeaf;\n                if (leaf && leaf.view && leaf.view.editor) {\n                    leaf.view.editor.replaceSelection(currentMarkdown + '\\n');\n                    status.textContent = t('mod.excelToMd.inserted');\n                    status.className = 'excel-md-status success';\n                } else {\n                    throw new Error('未找到活动编辑器');\n                }\n            }\n        } catch (e) {\n            status.textContent = '插入失败: ' + e.message + '。请打开一篇笔记后再试。';\n            status.className = 'excel-md-status error';\n        }\n    });\n\n    // 复制\n    copyBtn.addEventListener('click', function() {\n        if (!currentMarkdown) {\n            status.textContent = t('mod.excelToMd.nothingToCopy2');\n            status.className = 'excel-md-status error';\n            return;\n        }\n        try {\n            navigator.clipboard.writeText(currentMarkdown).then(function() {\n                status.textContent = t('mod.excelToMd.copied');\n                status.className = 'excel-md-status success';\n                setTimeout(function() { status.textContent = ''; }, 2000);\n            }).catch(function() {\n                status.textContent = '复制失败，请手动选中上方文本复制';\n                status.className = 'excel-md-status error';\n            });\n        } catch (e) {\n            status.textContent = t('mod.excelToMd.copyFailed');\n            status.className = 'excel-md-status error';\n        }\n    });\n\n    // 清空\n    clearBtn.addEventListener('click', function() {\n        textarea.value = '';\n        output.textContent = t('mod.excelToMd.waiting');\n        output.addClass('empty');\n        tablePreview.innerHTML = '';\n        currentMarkdown = '';\n        status.textContent = '';\n    });\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.excelToMd.settings.title') });\n\n    containerEl.createEl('p', {\n        text: t('mod.excelToMd.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n\n    new Setting(containerEl)\n        .setName(t('mod.excelToMd.settings.auto'))\n        .setDesc(t('mod.excelToMd.settings.autoDesc'))\n        .addToggle(function(t) {\n            t.setValue(settings.enabledAutoConvert !== false);\n            t.onChange(async function(v) {\n                settings.enabledAutoConvert = v;\n                await saveCallback();\n            });\n        });\n\n    containerEl.createEl('p', {\n        text: t('mod.excelToMd.settings.hint1'),\n        attr: { style: 'color:var(--text-muted);font-size:11px;margin-top:10px;' }\n    });\n    containerEl.createEl('p', {\n        text: t('mod.excelToMd.settings.hint2'),\n        attr: { style: 'color:var(--text-muted);font-size:11px;' }\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "farm-clicker": "/**\n * 农场点击模块 V1\n * 简约放置农场 — 点击收获/自动产出/升级系统\n */\nconst id = 'farm-clicker';\nconst title = t('mod.farmClicker');\nconst icon = '🌾';\n\nconst defaultSettings = { autoSave: true };\n\nconst styles = `\n.fc-wrap { height:100%; display:flex; flex-direction:column; overflow:hidden; font-family:'Courier New',monospace; image-rendering:pixelated; }\n.fc-bar { display:flex; gap:4px; padding:6px 10px; flex-shrink:0; background:rgba(139,90,43,.2); border-radius:8px 8px 0 0;\n    align-items:center; justify-content:center; }\n.fc-btn { padding:5px 14px; border:none; border-radius:12px; cursor:pointer; font-size:11px; font-weight:700;\n    color:#fff; text-shadow:1px 1px 0 rgba(0,0,0,.3); transition:all .15s; }\n.fc-btn-harvest { background:#e67e22; box-shadow:0 3px 0 #a55c10,inset 0 -2px 6px rgba(0,0,0,.2); }\n.fc-btn-harvest:hover { background:#f39c12; transform:translateY(-1px); }\n.fc-btn-upgrade { background:#27ae60; box-shadow:0 3px 0 #1a7d42,inset 0 -2px 6px rgba(0,0,0,.2); }\n.fc-btn-upgrade:hover { background:#2ecc71; transform:translateY(-1px); }\n.fc-body { flex:1; display:flex; flex-direction:column; overflow-y:auto;\n    background:linear-gradient(180deg,#87CEEB 0%,#b8e6b8 30%,#7cb342 65%,#5D4037 100%);\n    border:3px solid #5D4037; border-top:none; border-radius:0 0 10px 10px; position:relative; }\n/* HUD */\n.fc-hud { display:flex; gap:8px; padding:8px 12px; z-index:5; flex-wrap:wrap; justify-content:center; }\n.fc-hud-item { background:rgba(255,255,255,.75); padding:4px 12px; border-radius:12px; font-size:12px;\n    box-shadow:0 2px 6px rgba(0,0,0,.08); min-width:80px; text-align:center; }\n.fc-hud-label { font-size:9px; color:#888; text-transform:uppercase; letter-spacing:.5px; }\n.fc-hud-val { font-weight:700; color:#333; }\n.coins-gold { color:#f39c12 !important; }\n/* Field area */\n.fc-field { flex:1; display:flex; flex-direction:column; padding:10px; gap:6px; overflow-y:auto; }\n.fc-crop-row { display:flex; gap:6px; align-items:center; background:rgba(255,255,255,.35); border-radius:10px; padding:6px 10px;\n    border:2px solid rgba(93,64,55,.15); transition:border-color .2s; cursor:pointer; }\n.fc-crop-row:hover { border-color:rgba(93,64,55,.4); background:rgba(255,255,255,.45); }\n.fc-crop-row.ready { border-color:rgba(200,180,50,.5); background:rgba(255,248,190,.45);\n    animation:fcReadyPulse 1.5s ease-in-out infinite alternate; }\n@keyframes fcReadyPulse { 0%{box-shadow:0 0 0 0 rgba(240,200,60,0)} 100%{box-shadow:0 0 16px 4px rgba(240,200,60,.15)} }\n.fc-crop-icon { font-size:26px; width:36px; height:36px; display:flex; align-items:center; justify-content:center; }\n.fc-crop-info { flex:1; min-width:0; }\n.fc-crop-name { font-size:11px; font-weight:600; color:#333; }\n.fc-crop-progress { height:6px; background:rgba(0,0,0,.08); border-radius:3px; margin-top:3px; overflow:hidden; }\n.fc-crop-fill { height:100%; border-radius:3px; transition:width .25s ease; }\n.fill-grow { background:linear-gradient(90deg,#81c784,#66bb6a); }\n.fill-ready { background:linear-gradient(90deg,#ffd54f,#ffb300); animation:fcFillShine 1.2s linear infinite; }\n@keyframes fcFillShine { 0%{background-position:0} 100%{background-position:20px} }\n.fc-crop-stats { font-size:9px; color:#888; display:flex; gap:6px; margin-top:2px; }\n.fc-crop-earn { font-size:13px; font-weight:700; color:#27ae60; white-space:nowrap; }\n/* Upgrade panel */\n.fc-upgrade-panel { padding:8px 12px; border-top:1px solid rgba(93,64,55,.15); background:rgba(255,255,255,.18); }\n.fc-upg-title { font-size:11px; font-weight:700; color:#5D4037; margin-bottom:6px; text-transform:uppercase; letter-spacing:.5px; }\n.fc-upg-list { display:flex; gap:6px; overflow-x:auto; padding-bottom:4px; scrollbar-width:none; }\n.fc-upg-list::-webkit-scrollbar { display:none; }\n.fc-upg-card { flex-shrink:0; background:rgba(255,255,255,.6); border:2px solid rgba(93,64,55,.15); border-radius:10px; padding:8px 10px; text-align:center;\n    cursor:pointer; min-width:72px; transition:all .15s; }\n.fc-upg-card:hover { border-color:rgba(46,204,113,.4); transform:translateY(-2px); }\n.fc-upg-card.maxed { opacity:.45; cursor:not-allowed; filter:grayscale(.5); border-color:transparent; }\n.fc-upg-card .upg-icon { font-size:22px; margin-bottom:2px; }\n.fc-upg-card .upg-name { font-size:9px; color:#5D4037; font-weight:600; }\n.fc-upg-card .upg-lv { font-size:10px; font-weight:700; color:#27ae60; }\n.fc-upg-card .upg-cost { font-size:9px; color:#888; }\n.fc-upg-card .upg-cost.can-buy { color:#e67e22; font-weight:700; }\n\n/* Floating +N text */\n.fc-float { position:absolute; pointer-events:none; z-index:30; font-size:14px; font-weight:700;\n    color:#27ae60; text-shadow:0 1px 3px rgba(0,0,0,.3); animation:fcFloatUp .8s ease-out forwards; white-space:nowrap; }\n@keyframes fcFloatUp { 0%{opacity:1;transform:translateY(0) scale(1);} 100%{opacity:0;transform:translateY(-32px) scale(1.2);} }\n\n/* Ground decoration */\n.fc-ground-dec { position:absolute; bottom:0; left:0; right:0; height:28px; pointer-events:none; z-index:0;\n    background:repeating-linear-gradient(90deg,transparent 0, transparent 20px, rgba(101,67,33,.15) 20px, rgba(101,67,33,.15) 21px,\n        transparent 21px, transparent 48px, rgba(160,130,80,.1) 48px, rgba(160,130,80,.1) 49px);\n}\n`;\n\nconst CROPS = [\n    { id:'wheat', icon:'🌾', name:t('mod.farm.wheat')||'Wheat', basePrice:1, growSec:3, baseYield:1, unlockCost:0 },\n    { id:'carrot', icon:'🥕', name:t('mod.farm.carrot')||'Carrot', basePrice:3, growSec:5, baseYield:2, unlockCost:15 },\n    { id:'tomato', icon:'🍅', name:t('mod.farm.tomato')||'Tomato', basePrice:8, growSec:8, baseYield:5, unlockCost:50 },\n    { id:'corn', icon:'🌽', name:t('mod.farm.corn')||'Corn', basePrice:20, growSec:12, baseYield:12, unlockCost:150 },\n    { id:'strawberry', icon:'🍓', name:t('mod.farm.strawberry')||'Strawberry', basePrice:50, growSec:18, baseYield:30, unlockCost:400 },\n    { id:'pumpkin', icon:'🎃', name:t('mod.farm.pumpkin')||'Pumpkin', basePrice:120, growSec:30, baseYield:70, unlockCost:1200 },\n];\n\nconst UPGRADES = [\n    { id:'clickPower', icon:'👆', name:(t('mod.farm.upgClick')||'Click Power'), desc:'+1 per click', baseCost:10, mult:1.5 },\n    { id:'autoHarvest', icon:'⏱️', name:(t('mod.farm.upgAuto')||'Auto Speed'), desc:'+20% speed', baseCost:50, mult:1.8 },\n    { id:'yieldBonus', icon:'✨', name:(t('mod.farm.upgYield')||'Yield Bonus'), desc:'+15% yield', baseCost:200, mult:2 },\n    { id:'goldenTouch', icon:'🏆', name:(t('mod.farm.upgGolden')||'Golden Touch'), desc:'+5% gold crops', baseCost:800, mult:2.5 },\n];\n\nfunction _fk() { return 'farm_v1'; }\nfunction _fs() { try{ const d=localStorage.getItem(_fk()); return d?JSON.parse(d):null; }catch(e){return null;} }\nfunction _fsave(d){ try{ localStorage.setItem(_fk(),JSON.stringify(d)); }catch(e){} }\n\nasync function render(content) {\n    content.empty();\n    const wrap = content.createDiv({ cls: 'fc-wrap' });\n\n    // State\n    let state = _fs() || {\n        coins:0, totalEarned:0, totalClicks:0,\n        crops:CROPS.map(c=>({id:c.id, count:Math.floor(c.id==='wheat'?3:0), plantedAt:null, ready:false})),\n        upgrades:{ clickPower:1, autoHarvest:1, yieldBonus:1, goldenTouch:0 },\n        lastSave:Date.now()\n    };\n\n    // Toolbar\n    const bar = wrap.createDiv({ cls: 'fc-bar' });\n    bar.createEl('button',{cls:'fc-btn fc-btn-harvest',text:'🌾 '+((t('mod.farm.harvestAll')||'Harvest All'))}).onclick=()=>harvestAll(wrap);\n    bar.createEl('button',{cls:'fc-btn fc-btn-upgrade',text:'⬆ '+((t('mod.farm.upgrades')||'Upgrades'))}).onclick=()=>scrollToUpgrades();\n\n    // Body\n    const body = wrap.createDiv({ cls: 'fc-body' });\n\n    // Ground deco\n    body.createDiv({ cls: 'fc-ground-dec' });\n\n    // HUD\n    const hud = body.createDiv({ cls: 'fc-hud' });\n    const hudCoins = hud.createDiv({ cls: 'fc-hud-item coins-gold' });\n    hudCoins.innerHTML = `<div class=\"fc-hud-label\">${t('mod.farm.coins')||'Coins'}</div><div class=\"fc-hud-val\" id=\"fcCoins\">0</div>`;\n    const hudRate = hud.createDiv({ cls: 'fc-hud-item' });\n    hudRate.innerHTML = `<div class=\"fc-hud-label\">${t('mod.farm.perSec')||'/sec'}</div><div class=\"fc-hud-val\" id=\"fcRate\">--</div>`;\n\n    // Field\n    const field = body.createDiv({ cls: 'fc-field' });\n    field.onclick=(e)=>{ if(e.target===field || e.target.classList.contains('fc-field')) doClick(body,e); };\n\n    // Upgrade panel\n    const upgPanel = body.createDiv({ cls: 'fc-upgrade-panel' });\n    upgPanel.createEl('h4',{cls:'fc-upg-title',text:'⬆ '+(t('mod.farm.upgradeShop')||'Upgrade Shop')});\n    const upgList = upgPanel.createDiv({ cls: 'fc-upg-list' });\n\n    function scrollToUpgrades(){ upgPanel.scrollIntoView({behavior:'smooth'}); }\n\n    // Render upgrades\n    function renderUpgrades(){\n        upgList.empty();\n        UPGRADES.forEach(u => {\n            const lvl = state.upgrades[u.id] || 0;\n            const cost = Math.round(u.baseCost * Math.pow(u.mult, lvl));\n            const canBuy = state.coins >= cost && lvl < 10;\n\n            const card = upgList.createDiv({ cls:`fc-upg-card ${lvl>=10?'maxed':''}` });\n            card.createDiv({cls:'upg-icon',text:u.icon});\n            card.createDiv({cls:'upg-name',text:u.name});\n            card.createDiv({cls:'upg-lv',text:`Lv.${lvl}/10`});\n            card.createDiv({cls:`upg-cost ${canBuy?'can-buy':''}`,text:canBuy?`💰${cost}`:'MAX'});\n            if (canBuy) card.onclick=()=>buyUpgrade(u.id,cost);\n        });\n    }\n\n    function buyUpgrade(uid,cost){\n        state.coins -= cost;\n        state.upgrades[uid] = (state.upgrades[uid]||0)+1;\n        _fsave(state); updateHUD(); renderUpgrades();\n        showFloat(body, `${UPGRADES.find(u=>u.id===uid).icon} UP!`, '#27ae60');\n    }\n\n    // Render crop rows\n    function renderCrops(){\n        field.empty();\n        state.crops.forEach((cp,i)=>{\n            if (cp.count <= 0) return;\n            const cropDef = CROPS.find(c=>c.id===cp.id);\n            if (!cropDef) return;\n\n            const row = field.createDiv({ cls:`fc-crop-row ${cp.ready?'ready':''}` });\n            row.onclick = (e) => { e.stopPropagation(); harvestCrop(i,body,row); };\n\n            row.createDiv({ cls: 'fc-crop-icon', text: cropDef.icon });\n\n            const info = row.createDiv({ cls: 'fc-crop-info' });\n            info.createDiv({ cls: 'fc-crop-name', text: `${cropDef.name} x${cp.count}` });\n\n            const progBar = info.createDiv({ cls: 'fc-crop-progress' });\n            const fill = progBar.createDiv({ cls: `fc-crop-fill ${cp.ready?'fill-ready':'fill-grow'}` });\n            fill.style.width = cp.ready ? '100%' : getProgress(cp,cropDef)+'%';\n\n            const stats = info.createDiv({ cls: 'fc-crop-stats' });\n            stats.innerHTML `<span>⏱${cropDef.growSec}s</span>\n                <span>💰${getEarning(cropDef)}</span><span>${cp.ready?((t('mod.farm.ready')||'READY!')):((t('mod.farm.growing')||'Growing...'))}</span>`;\n\n            if (cp.ready){\n                row.createDiv({ cls:'fc-crop-earn', text: `+${getTotalEarning(cropDef,cp.count)}` });\n            }\n        });\n    }\n\n    function getProgress(cp,cd){\n        if(!cp.plantedAt||cp.ready)return 100;\n        const elapsed=(Date.now()-cp.plantedAt)/1000;\n        return Math.min(100,(elapsed/cd.growSec)*100*(state.upgrades.autoHarvest||1));\n    }\n\n    function getEarning(cd){\n        let e=cd.baseYield;\n        e*=state.upgrades.yieldBonus||(1);\n        if(state.upgrades.goldenTouch>=1&&Math.random()<.05*state.upgrades.goldenTouch) e*=3;\n        return Math.floor(e);\n    }\n\n    function getTotalEarning(cd,count){ return getEarning(cd)*count; }\n\n    function doClick(containerBody,e){\n        const rect=containerBody.getBoundingClientRect();\n        const x=e.clientX-rect.left,y=e.clientY-rect.top;\n        const earn=state.upgrades.clickPower||1;\n        state.coins+=earn; state.totalEarned+=earn; state.totalClicks++;\n        updateHUD(); _fsave(state);\n        showFloatAt(containerBody,x,y,`+${earn}`);\n\n        // Click particles (simple)\n        for(let i=0;i<3;i++){\n            setTimeout(()=>{\n                showFloatAt(containerBody,x+(Math.random()-0.5)*30,y-(i*8),'+1');\n            },i*50);\n        }\n    }\n\n    function harvestCrop(idx,containerBody,rowEl){\n        const cp=state.crops[idx]; if(!cp||!cp.ready)return;\n        const cd=CROPS.find(c=>c.id===cp.id); if(!cd)return;\n        const earn=getTotalEarning(cd,cp.count);\n\n        state.coins+=earn; state.totalEarned+=earn;\n        cp.plantedAt=null; cp.ready=false;\n\n        _fsave(state); updateHUD(); renderCrops();\n\n        // Float from the row's position\n        if(rowEl){\n            const rRect=rowEl.getBoundingClientRect();\n            const cRect=containerBody.getBoundingClientRect();\n            showFloatAt(containerBody,rRect.left-cRect.left+rRect.width/2,rRect.top-cRect.top,`+${earn}`);\n        }\n    }\n\n    function harvestAll(containerBody){\n        let total=0;\n        state.crops.forEach((cp,i)=>{\n            if(!cp.ready)return;\n            const cd=CROPS.find(c=>c.id===cp.id); if(!cd)return;\n            total+=getTotalEarning(cd,cp.count);\n            cp.plantedAt=null; cp.ready=false;\n        });\n        if(total>0){\n            state.coins+=total; state.totalEarned+=total;\n            updateHUD(); _fsave(state); renderCrops();\n            showFloat(containerBody,`+${total} 🎉`);\n        }\n    }\n\n    function showFloat(el,text,color){\n        const f=document.createElement('div'); f.className='fc-float'; f.textContent=text; f.style.color=color||'#27ae60';\n        el.appendChild(f); setTimeout(()=>f.remove(),800);\n    }\n    function showFloatAt(el,x,y,text,color){\n        const f=document.createElement('div'); f.className='fc-float'; f.textContent=text; f.style.color=color||'#27ae60';\n        f.style.left=x+'px'; f.style.top=y+'px';\n        el.appendChild(f); setTimeout(()=>f.remove(),800);\n    }\n\n    function updateHUD(){\n        document.getElementById('fcCoins').textContent=formatNum(state.coins);\n        // Calculate approximate /sec\n        let rate=0;\n        state.crops.forEach(cp=>{\n            if(!cp.ready)return;\n            const cd=CROPS.find(c=>c.id===cp.id); if(!cd)return;\n            rate+=getTotalEarning(cd,cp.count)/(cd.growSec/(state.upgrades.autoHarvest||1));\n        });\n        document.getElementById('fcRate').textContent=rate>0?formatNum(rate.toFixed(1)): '--';\n    }\n\n    function formatNum(n){\n        if(n>=1e6)return (n/1e6).toFixed(1)+'M';\n        if(n>=1e3)return (n/1e3).toFixed(1)+'K';\n        return String(Math.floor(n));\n    }\n\n    // Auto-plant on start\n    state.crops.forEach((cp,i)=>{\n        if(cp.count>0&&!cp.plantedAt&&!cp.ready){\n            cp.plantedAt=Date.now()-Math.random()*3000; // stagger\n        }\n    });\n\n    // Growth loop — every 500ms\n    setInterval(()=>{\n        state.crops.forEach((cp,i)=>{\n            if(!cp||cp.count<=0||cp.plantedAt||cp.ready)return;\n            const delay=Math.random()*800;\n            setTimeout(()=>{ cp.plantedAt=Date.now(); },delay);\n        });\n\n        // Check growth\n        let changed=false;\n        state.crops.forEach(cp=>{\n            if(!cp||!cp.plantedAt||cp.ready)return;\n            const cd=CROPS.find(c=>c.id===cp.id); if(!cd)return;\n            const elapsed=(Date.now()-cp.plantedAt)/1000;\n            const speedMult=state.upgrades.autoHarvest||1;\n            if(elapsed>=cd.growSec/speedMult){\n                cp.ready=true; changed=true;\n            }\n        });\n        if(changed){ renderCrops(); updateHUD(); _fsave(state);}\n    },500);\n\n    // Initial renders\n    renderCrops(); renderUpgrades(); updateHUD();\n\n    // Auto-save every 30s\n    setInterval(()=>{_fsave(state);},30000);\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render };\n",
  "folder-counter": "// folder-counter 模块 - 文件夹笔记计数器（递归统计）\n// 源插件: file-explorer-note-count（文件夹文件计数器）\nconst id = 'folder-counter';\nconst title = t('mod.folderCounter');\nconst icon = '📁';\n\nconst defaultSettings = {\n    showAllNumbers: true,\n    addRootFolder: false\n};\n\nconst styles = `\n.fc-panel { padding: 10px 12px; display: flex; flex-direction: column; height: 100%; }\n.fc-header { font-size: 12px; font-weight: 600; color: var(--v6-primary); margin-bottom: 6px; padding-bottom: 4px; border-bottom: 1px solid var(--background-modifier-border); display: flex; align-items: center; justify-content: space-between; }\n.fc-refresh { font-size: 10px; color: var(--text-faint); cursor: pointer; padding: 2px 6px; border-radius: 3px; }\n.fc-refresh:hover { background: var(--background-modifier-hover); color: var(--text-normal); }\n.fc-list { flex: 1; overflow-y: auto; }\n.fc-folder { margin-bottom: 2px; }\n.fc-folder-header { display: flex; align-items: center; padding: 5px 8px; border-radius: 4px; cursor: pointer; transition: background 0.1s; font-size: 12px; }\n.fc-folder-header:hover { background: var(--background-modifier-hover); }\n.fc-folder-arrow { width: 14px; font-size: 10px; color: var(--text-faint); flex-shrink: 0; transition: transform 0.15s; }\n.fc-folder-arrow.open { transform: rotate(90deg); }\n.fc-folder-icon { margin-right: 4px; flex-shrink: 0; }\n.fc-folder-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--text-normal); }\n.fc-folder-count { font-weight: 600; color: var(--text-muted); background: var(--background-modifier-form-field); padding: 1px 8px; border-radius: 10px; font-size: 10px; flex-shrink: 0; margin-left: 6px; }\n.fc-subfolders { padding-left: 16px; display: none; }\n.fc-subfolders.open { display: block; }\n.fc-subfolder .fc-folder-header { font-size: 11px; }\n.fc-total { padding: 6px 8px; font-size: 10px; color: var(--text-faint); border-top: 1px solid var(--background-modifier-border); margin-top: 4px; flex-shrink: 0; }\n.fc-total strong { color: var(--text-normal); }\n.fc-empty { text-align: center; color: var(--text-muted); padding: 20px; font-size: 12px; }\n`;\n\nasync function render(container) {\n    container.addClass('fc-panel');\n\n    var header = container.createDiv({ cls: 'fc-header' });\n    header.createSpan({ text: t('mod.folderCounter.title') });\n    var refreshBtn = header.createEl('span', { text: t('mod.folderCounter.refresh'), cls: 'fc-refresh' });\n\n    var listEl = container.createDiv({ cls: 'fc-list' });\n    var totalEl = container.createDiv({ cls: 'fc-total' });\n\n    // 递归构建文件夹树\n    function buildFolderTree() {\n        var files = app.vault.getMarkdownFiles();\n        var root = { name: '/', subfolders: {}, files: 0, depth: 0 };\n\n        files.forEach(function(f) {\n            var parts = f.path.split('/');\n            var node = root;\n            for (var i = 0; i < parts.length - 1; i++) {\n                if (!node.subfolders[parts[i]]) {\n                    node.subfolders[parts[i]] = { name: parts[i], subfolders: {}, files: 0, depth: node.depth + 1 };\n                }\n                node = node.subfolders[parts[i]];\n            }\n            node.files++;\n        });\n\n        return root;\n    }\n\n    function scanFolders() {\n        listEl.innerHTML = '';\n        var tree = buildFolderTree();\n\n        var totalFiles = 0;\n        var totalFolders = 0;\n\n        function countAll(node) {\n            totalFiles += node.files;\n            var folderCount = 0;\n            Object.keys(node.subfolders).forEach(function(k) {\n                folderCount += countAll(node.subfolders[k]);\n            });\n            if (node.depth > 0) totalFolders++;\n            return folderCount + 1;\n        }\n        countAll(tree);\n\n        if (Object.keys(tree.subfolders).length === 0) {\n            listEl.createDiv({ text: t('mod.folderCounter.empty'), cls: 'fc-empty' });\n            totalEl.innerHTML = '';\n            return;\n        }\n\n        // 只渲染前两层（顶级文件夹 + 子文件夹可展开）\n        Object.keys(tree.subfolders).sort(function(a, b) { return a.localeCompare(b.name); }).forEach(function(key) {\n            renderFolder(listEl, tree.subfolders[key], key);\n        });\n\n        totalEl.innerHTML = t('mod.folderCounter.total') + '<strong>' + totalFolders + '</strong>' + t('mod.folderCounter.totalFolders') + '<strong>' + totalFiles + '</strong>' + t('mod.folderCounter.totalNotes');\n    }\n\n    function countFilesRecursive(node) {\n        var count = node.files;\n        Object.keys(node.subfolders).forEach(function(k) {\n            count += countFilesRecursive(node.subfolders[k]);\n        });\n        return count;\n    }\n\n    function renderFolder(parentEl, node, fullPath) {\n        var folderDiv = parentEl.createDiv({ cls: 'fc-folder' });\n        var hasSub = Object.keys(node.subfolders).length > 0;\n        var totalInFolder = countFilesRecursive(node);\n\n        var header = folderDiv.createDiv({ cls: 'fc-folder-header' });\n\n        var arrow = header.createSpan({ cls: 'fc-folder-arrow', text: hasSub ? '▶' : '  ' });\n        header.createSpan({ text: '📂', cls: 'fc-folder-icon' });\n        header.createSpan({ text: node.name, cls: 'fc-folder-name' });\n        header.createSpan({ text: totalInFolder + t('mod.folderCounter.noteCount'), cls: 'fc-folder-count' });\n\n        header.addEventListener('click', function(evt) {\n            evt.preventDefault(); evt.stopPropagation(); evt.stopImmediatePropagation();\n            // 在文件浏览器中定位\n            try {\n                var explorer = app.workspace.getLeavesOfType('file-explorer');\n                if (explorer.length > 0) {\n                    app.workspace.revealLeaf(explorer[0]);\n                }\n            } catch(e) {}\n        });\n\n        if (hasSub) {\n            var subEl = folderDiv.createDiv({ cls: 'fc-subfolders' });\n\n            arrow.addEventListener('click', function(e) {\n                e.stopPropagation();\n                var isOpen = subEl.classList.contains('open');\n                if (isOpen) {\n                    subEl.classList.remove('open');\n                    arrow.classList.remove('open');\n                    arrow.textContent = '▶';\n                } else {\n                    subEl.classList.add('open');\n                    arrow.classList.add('open');\n                    arrow.textContent = '▼';\n                }\n            });\n\n            var subFolders = Object.keys(node.subfolders).sort();\n            // t('mod.folderCounter.limit20')\n            subFolders.slice(0, 20).forEach(function(k) {\n                var subNode = node.subfolders[k];\n                var subDiv = subEl.createDiv({ cls: 'fc-folder fc-subfolder' });\n                var subHeader = subDiv.createDiv({ cls: 'fc-folder-header' });\n                subHeader.createSpan({ text: '  ', cls: 'fc-folder-arrow' });\n                subHeader.createSpan({ text: '📁', cls: 'fc-folder-icon' });\n                subHeader.createSpan({ text: k, cls: 'fc-folder-name' });\n                subHeader.createSpan({ text: countFilesRecursive(subNode) + t('mod.folderCounter.noteCount'), cls: 'fc-folder-count' });\n            });\n\n            if (subFolders.length > 20) {\n                subEl.createDiv({ text: t('mod.folderCounter.more') + (subFolders.length - 20) + t('mod.folderCounter.moreSuffix'), cls: 'fc-empty', attr: { style: 'font-size:10px;padding:4px;' } });\n            }\n        }\n    }\n\n    setTimeout(function() { scanFolders(); }, 1100);\n    refreshBtn.addEventListener('click', scanFolders);\n\n    // 定时刷新\n    var interval = setInterval(scanFolders, 3600000)/*TEMP_DISABLED*/;\n    // 保存引用供 onunload 清理\n    // 容器从DOM断开时自动清理\n    var observer = new MutationObserver(function() {\n        if (!container.isConnected) {\n            clearInterval(interval);\n            interval = null;\n            observer.disconnect();\n            observer = null;\n        }\n    });\n    if (container.parentElement) observer.observe(container.parentElement, { childList: true });\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.folderCounter.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.folderCounter.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n}\n\n// 导出 onunload 供框架调用清理\nmodule.exports.onunload = function() {\n    if (typeof interval !== \"undefined\" && interval) { clearInterval(interval); interval = null; }\n    if (typeof observer !== \"undefined\" && observer) { observer.disconnect(); observer = null; }\n};\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "html-viewer": "// html-viewer 模块 - 浏览器内渲染HTML文件\n// 源插件: obsidian-html-plugin\n// 核心功能保留: iframe渲染 + 安全模式切换 + 文件浏览\nconst id = 'html-viewer';\nconst title = t('mod.htmlViewer');\nconst icon = '🌐';\n\nconst defaultSettings = {\n    opMode: 'BalanceMode',\n    zoomValue: 1\n};\n\nconst styles = `\n.htmlv-wrap { padding: 8px 10px; display: flex; flex-direction: column; height: 100%; }\n.htmlv-toolbar { display: flex; align-items: center; gap: 6px; margin-bottom: 8px; flex-shrink: 0; flex-wrap: wrap; }\n.htmlv-toolbar select { padding: 3px 6px; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: var(--background-secondary); color: var(--text-normal); font-size: 11px; }\n.htmlv-toolbar button { padding: 3px 8px; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 11px; cursor: pointer; }\n.htmlv-toolbar button:hover { background: var(--background-modifier-hover); }\n.htmlv-toolbar button.active { background: var(--v6-primary); color: white; border-color: var(--v6-primary); }\n.htmlv-filelist { max-height: 100px; overflow-y: auto; margin-bottom: 6px; flex-shrink: 0; border: 1px solid var(--background-modifier-border); border-radius: 4px; }\n.htmlv-filelist.hidden { display: none; }\n.htmlv-file-item { display: flex; align-items: center; justify-content: space-between; padding: 4px 8px; cursor: pointer; font-size: 11px; transition: background 0.1s; }\n.htmlv-file-item:hover { background: var(--background-modifier-hover); }\n.htmlv-file-item.selected { background: rgba(var(--v6-primary-rgb, 232,149,109), 0.15); color: var(--v6-primary); }\n.htmlv-file-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.htmlv-file-size { color: var(--text-faint); font-size: 10px; flex-shrink: 0; margin-left: 8px; }\n.htmlv-viewer { flex: 1; border: 1px solid var(--background-modifier-border); border-radius: 4px; overflow: hidden; background: #fff; position: relative; min-height: 100px; }\n.htmlv-viewer iframe { width: 100%; height: 100%; border: none; }\n.htmlv-viewer.text-mode { background: var(--background-secondary); padding: 10px; overflow: auto; }\n.htmlv-viewer.text-mode pre { margin: 0; font-size: 12px; white-space: pre-wrap; word-break: break-all; font-family: var(--font-monospace); }\n.htmlv-viewer .htmlv-empty { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; text-align: center; }\n.htmlv-nav { display: flex; gap: 4px; align-items: center; margin-left: auto; }\n.htmlv-nav button { padding: 2px 6px; font-size: 11px; }\n.htmlv-statusbar { display: flex; align-items: center; justify-content: space-between; padding: 3px 6px; font-size: 10px; color: var(--text-faint); background: var(--background-secondary); border-top: 1px solid var(--background-modifier-border); flex-shrink: 0; }\n`;\n\n// ============ 安全模式定义 ============\nvar MODE_CONFIG = {\n    TextMode:      { sandbox: '', csp: '', sanitize: true,  allowScripts: false, allowSameOrigin: false, renderAs: 'text' },\n    BalanceMode:   { sandbox: 'allow-same-origin', csp: \"default-src 'none'; style-src 'unsafe-inline'; img-src data:;\", sanitize: true, allowScripts: false, allowSameOrigin: true, renderAs: 'iframe' },\n    UnrestrictedMode: { sandbox: 'allow-same-origin allow-scripts', csp: '', sanitize: false, allowScripts: true, allowSameOrigin: true, renderAs: 'iframe' }\n};\n\n// 简易HTML清洗（移除script标签）\nfunction sanitizeHtml(html) {\n    return html.replace(/<script\\b[^<]*(?:(?!<\\/script>)<[^<]*)*<\\/script>/gi, '')\n               .replace(/\\son\\w+\\s*=\\s*\"[^\"]*\"/gi, '')\n               .replace(/\\son\\w+\\s*=\\s*'[^']*'/gi, '');\n}\n\nasync function render(container) {\n    container.addClass('htmlv-wrap');\n\n    var s = settings; // 来自 with(_runtimeCtx)\n\n    // 顶部工具栏\n    var toolbar = container.createDiv({ cls: 'htmlv-toolbar' });\n\n    var toggleFilesBtn = toolbar.createEl('button', { text: t('mod.htmlViewer.btn.files') });\n    var modeLabel = toolbar.createEl('span', { text: t('mod.htmlViewer.btn.mode'), attr: { style: 'font-size:11px;color:var(--text-muted);' } });\n    var modeSelect = toolbar.createEl('select');\n    ['TextMode', 'BalanceMode', 'UnrestrictedMode'].forEach(function(m) {\n        var opt = modeSelect.createEl('option', { text: m === 'TextMode' ? t('mod.htmlViewer.mode.text') : m === 'BalanceMode' ? t('mod.htmlViewer.mode.safe') : t('mod.htmlViewer.mode.trust'), attr: { value: m } });\n        if (s.opMode === m) opt.selected = true;\n    });\n\n    var refreshBtn = toolbar.createEl('button', { text: t('mod.htmlViewer.btn.refresh') });\n\n    // 文件列表\n    var fileList = container.createDiv({ cls: 'htmlv-filelist' });\n    var viewer = container.createDiv({ cls: 'htmlv-viewer' });\n    viewer.innerHTML = '<div class=\"htmlv-empty\">' + t('mod.htmlViewer.hint') + '</div>';\n\n    var statusbar = container.createDiv({ cls: 'htmlv-statusbar' });\n    var statusFile = statusbar.createSpan();\n    var statusMode = statusbar.createSpan();\n\n    var currentFile = null;\n    var currentContent = '';\n    var files = [];\n\n    // 扫描HTML文件\n    function scanFiles() {\n        files = [];\n        app.vault.getFiles().forEach(function(f) {\n            if (f.extension === 'html' || f.extension === 'htm') {\n                files.push({ path: f.path, name: f.name, size: f.stat ? f.stat.size : 0 });\n            }\n        });\n        files.sort(function(a, b) { return a.name.localeCompare(b.name); });\n        renderFileList();\n    }\n\n    // 渲染文件列表\n    function renderFileList() {\n        fileList.innerHTML = '';\n        if (files.length === 0) {\n            fileList.createDiv({ text: t('mod.htmlViewer.empty'), cls: 'htmlv-file-item' }).style.cssText = 'cursor:default;color:var(--text-muted);';\n        } else {\n            files.forEach(function(f) {\n                var item = fileList.createDiv({ cls: 'htmlv-file-item' });\n                item.createSpan({ text: f.name, cls: 'htmlv-file-name' });\n                var sizeKB = Math.round(f.size / 1024);\n                if (sizeKB > 0) item.createSpan({ text: sizeKB + 'KB', cls: 'htmlv-file-size' });\n\n                if (currentFile && currentFile.path === f.path) {\n                    item.addClass('selected');\n                }\n\n                // 安全点击：阻止Obsidian事件系统\n                item.addEventListener('mousedown', function(evt) { evt.preventDefault(); evt.stopPropagation(); }, true);\n                item.addEventListener('click', function(evt) {\n                    evt.preventDefault(); evt.stopPropagation(); evt.stopImmediatePropagation();\n                    loadFile(f);\n                }, true);\n            });\n        }\n    }\n\n    // 加载并渲染文件\n    async function loadFile(file) {\n        currentFile = file;\n        renderFileList();\n\n        try {\n            var fileObj = app.vault.getAbstractFileByPath(file.path);\n            if (!fileObj) { showError(t('mod.htmlViewer.error.load') + file.path); return; }\n            currentContent = await app.vault.read(fileObj);\n            renderContent();\n        } catch (e) {\n            showError(t('mod.dataEditor.readFailed') + e.message);\n        }\n    }\n\n    // 按当前模式渲染\n    function renderContent() {\n        var mode = s.opMode || 'BalanceMode';\n        var config = MODE_CONFIG[mode] || MODE_CONFIG['BalanceMode'];\n        viewer.innerHTML = '';\n\n        if (!currentContent) {\n            viewer.innerHTML = '<div class=\"htmlv-empty\">' + t('mod.htmlViewer.emptyContent') + '</div>';\n            return;\n        }\n\n        if (config.renderAs === 'text') {\n            // 文本模式：显示原始HTML\n            viewer.addClass('text-mode');\n            var pre = viewer.createEl('pre');\n            pre.textContent = currentContent;\n        } else {\n            // iframe模式\n            viewer.removeClass('text-mode');\n            var iframe = viewer.createEl('iframe');\n\n            var content = currentContent;\n            if (config.sanitize) {\n                content = sanitizeHtml(content);\n            }\n\n            var cspMeta = config.csp ? '<meta http-equiv=\"Content-Security-Policy\" content=\"' + config.csp + '\">' : '';\n            var blobContent = '<!DOCTYPE html><html><head><meta charset=\"utf-8\">' + cspMeta + '<base target=\"_blank\"></head><body>' + content + '</body></html>';\n\n            var blob = new Blob([blobContent], { type: 'text/html' });\n            var url = URL.createObjectURL(blob);\n            iframe.src = url;\n\n            // 缩放\n            if (s.zoomValue && s.zoomValue !== 1) {\n                iframe.style.transform = 'scale(' + s.zoomValue + ')';\n                iframe.style.transformOrigin = '0 0';\n            }\n\n            // 清理blob URL（延迟释放）\n            setTimeout(function() { URL.revokeObjectURL(url); }, 30000);\n        }\n\n        statusFile.textContent = currentFile ? currentFile.name : '';\n    }\n\n    function showError(msg) {\n        viewer.innerHTML = '';\n        viewer.innerHTML = '<div class=\"htmlv-empty\" style=\"color:var(--text-error) !important;\">⚠ ' + msg + '</div>';\n        statusFile.textContent = '';\n    }\n\n    // 事件\n    modeSelect.addEventListener('change', function() {\n        s.opMode = modeSelect.value;\n        if (typeof saveCallback === 'function') saveCallback();\n        if (currentContent) renderContent();\n        updateStatusMode();\n    });\n\n    toggleFilesBtn.addEventListener('click', function() { fileList.classList.toggle('hidden'); });\n    refreshBtn.addEventListener('click', function() {\n        scanFiles();\n        if (currentFile) loadFile(currentFile);\n    });\n\n    function updateStatusMode() {\n        var labels = { TextMode: t('mod.htmlViewer.mode.text'), BalanceMode: t('mod.htmlViewer.mode.safe'), UnrestrictedMode: t('mod.htmlViewer.mode.trust') };\n        statusMode.textContent = t('mod.htmlViewer.modeLabel') + (labels[s.opMode] || s.opMode);\n    }\n\n    // 延迟初始化：避免所有模块同时扫描导致卡顿\n    setTimeout(function() {\n        scanFiles();\n    }, 150);\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.htmlViewer.settings.title') });\n\n    // 安全模式说明\n    var modeTable = containerEl.createEl('table', { attr: { style: 'width:100%;font-size:11px;border-collapse:collapse;margin-bottom:12px;' } });\n    var header = modeTable.createEl('tr');\n    header.createEl('th', { text: t('mod.htmlViewer.settings.mode'), attr: { style: 'border:1px solid var(--background-modifier-border);padding:4px;' } });\n    header.createEl('th', { text: t('mod.htmlViewer.settings.desc'), attr: { style: 'border:1px solid var(--background-modifier-border);padding:4px;' } });\n\n    [\n        [t('mod.htmlViewer.mode.text'), t('mod.htmlViewer.settings.text')],\n        [t('mod.htmlViewer.mode.safe'), t('mod.htmlViewer.settings.safe')],\n        [t('mod.htmlViewer.mode.trust'), t('mod.htmlViewer.settings.trust')]\n    ].forEach(function(row) {\n        var tr = modeTable.createEl('tr');\n        tr.createEl('td', { text: row[0], attr: { style: 'border:1px solid var(--background-modifier-border);padding:4px;font-weight:600;' } });\n        tr.createEl('td', { text: row[1], attr: { style: 'border:1px solid var(--background-modifier-border);padding:4px;color:var(--text-muted);' } });\n    });\n\n    containerEl.createEl('p', {\n        text: t('mod.htmlViewer.settings.hint'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;' }\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "image-gallery": "// image-gallery 模块 - 图片画廊+灯箱\n// 源插件: obsidian-image-gallery-diy (图片画廊)\n// 核心功能保留: 图片网格 + 内联灯箱（缩放/滑动/键盘导航）\n// 展示模式: 正方形(grid) / 瀑布流(masonry) / 全智能(auto)\nconst id = 'image-gallery';\nconst title = t('mod.imageGallery');\nconst icon = '🏞️';\n\nconst defaultSettings = {\n    imgFolder: '',\n    sortby: 'mtime',\n    sort: 'desc',\n    gridCols: 3,\n    displayMode: 'square',\n    showCount: true,\n    spacingLeft: 10,\n    spacingRight: 10,\n    itemGap: 4\n};\n\nconst styles = `\n.ig-wrap { padding: 8px 10px; display: block !important; width: 100% !important; box-sizing: border-box !important; }\n.ig-toolbar { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; flex-shrink: 0; flex-wrap: wrap; }\n.ig-toolbar select, .ig-toolbar input { padding: 3px 6px; border: 1px solid var(--background-modifier-border); border-radius: 3px; background: var(--background-secondary); color: var(--text-normal); font-size: 11px; }\n.ig-toolbar button { padding: 3px 8px; border: 1px solid var(--background-modifier-border); border-radius: 3px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 11px; cursor: pointer; }\n.ig-toolbar button:hover { background: var(--background-modifier-hover); }\n.ig-toolbar button.active { background: var(--interactive-accent); color: var(--text-on-accent); border-color: var(--interactive-accent); }\n.ig-toolbar label { font-size: 10px; color: var(--text-muted); }\n.ig-toolbar .ig-folder-input { flex: 1; min-width: 120px; max-width: 220px; }\n.ig-toolbar .ig-sep { width: 1px; height: 20px; background: var(--background-modifier-border); margin: 0 2px; }\n\n/* 间距调节面板 */\n.ig-spacing-panel { display: flex; align-items: center; gap: 6px; padding: 4px 0; margin-bottom: 4px; flex-wrap: wrap; }\n.ig-spacing-panel label { font-size: 10px; color: var(--text-muted); }\n.ig-spacing-panel input { width: 42px; padding: 2px 4px; border: 1px solid var(--background-modifier-border); border-radius: 3px; background: var(--background-secondary); color: var(--text-normal); font-size: 11px; text-align: center; }\n\n/* 正方形网格 — inline-block + text-align:center，不依赖容器flex宽度计算 */\n.ig-grid-square { display: block !important; width: 100% !important; text-align: center !important; font-size: 0 !important; }\n.ig-square-cell { display: inline-block !important; width: 100px; height: 100px; margin: 4px; vertical-align: top !important; overflow: hidden !important; }\n.ig-square-cell img { width: 100% !important; height: 100% !important; object-fit: cover !important; display: block !important; position: static !important; }\n\n/* 瀑布流 — 保持原始比例 */\n.ig-grid-masonry { column-gap: 4px !important; overflow-y: auto; flex: 1; width: 100%; }\n.ig-grid-masonry .ig-thumb { break-inside: avoid; margin-bottom: 4px; display: block; }\n.ig-grid-masonry .ig-thumb img { width: 100%; height: auto; display: block; object-fit: contain; }\n\n/* 全智能 */\n.ig-grid-auto { display: flex !important; flex-wrap: wrap !important; gap: 4px !important; overflow-y: auto; align-content: start; flex: 1; width: 100%; }\n.ig-grid-auto .ig-thumb { flex: 1 1 auto; min-width: 80px; max-width: 300px; }\n.ig-grid-auto .ig-thumb img { width: 100%; height: auto; display: block; object-fit: contain; }\n\n.ig-thumb { border-radius: 4px; overflow: hidden; cursor: pointer; border: 2px solid transparent; transition: border-color 0.15s, transform 0.15s; background: var(--background-modifier-form-field); position: relative; }\n.ig-thumb:hover { border-color: var(--v6-primary); transform: scale(1.03); z-index: 1; }\n.ig-empty { text-align: center; color: var(--text-muted); padding: 30px 20px; font-size: 13px; grid-column: 1 / -1; }\n\n/* 灯箱 */\n.ig-lightbox { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(0,0,0,0.92); z-index: 99999; display: flex; align-items: center; justify-content: center; flex-direction: column; }\n.ig-lightbox img { max-width: 90vw; max-height: 80vh; object-fit: contain; border-radius: 4px; box-shadow: 0 4px 30px rgba(0,0,0,0.5); transition: opacity 0.2s; }\n.ig-lb-close { position: absolute; top: 16px; right: 20px; font-size: 28px; color: #fff; cursor: pointer; opacity: 0.7; z-index: 2; transition: opacity 0.15s; line-height: 1; }\n.ig-lb-close:hover { opacity: 1; }\n.ig-lb-nav { position: absolute; top: 50%; transform: translateY(-50%); font-size: 36px; color: #fff; cursor: pointer; opacity: 0.6; padding: 20px; transition: opacity 0.15s; user-select: none; z-index: 2; }\n.ig-lb-nav:hover { opacity: 1; }\n.ig-lb-prev { left: 10px; }\n.ig-lb-next { right: 10px; }\n.ig-lb-counter { color: #fff; margin-top: 12px; font-size: 12px; opacity: 0.7; }\n.ig-lb-caption { color: #fff; margin-top: 6px; font-size: 13px; max-width: 80vw; text-align: center; }\n`;\n\nasync function render(container) {\n    container.addClass('ig-wrap');\n    // 强制容器宽度=100%（防御 Obsidian 主题覆盖）\n    container.style.setProperty('display', 'block', 'important');\n    container.style.setProperty('width', '100%', 'important');\n    container.style.setProperty('box-sizing', 'border-box', 'important');\n    var s = settings;\n\n    // 工具栏\n    var toolbar = container.createDiv({ cls: 'ig-toolbar' });\n\n    toolbar.createEl('label', { text: t('mod.imageGallery.display') });\n    var modes = [\n        { v: 'square', t: t('mod.imageGallery.mode.square'), tip: '微信朋友圈风格裁剪' },\n        { v: 'masonry', t: t('mod.imageGallery.mode.masonry'), tip: '保持原始比例' },\n        { v: 'auto', t: t('mod.imageGallery.mode.smart'), tip: '自适应宽度' }\n    ];\n    var modeBtns = {};\n    modes.forEach(function(m) {\n        var btn = toolbar.createEl('button', { text: m.t, attr: { title: m.tip } });\n        if (s.displayMode === m.v) btn.classList.add('active');\n        modeBtns[m.v] = btn;\n    });\n\n    toolbar.createEl('span', { cls: 'ig-sep' });\n\n    toolbar.createEl('label', { text: t('mod.imageGallery.folder') });\n    var folderInput = toolbar.createEl('input', {\n        cls: 'ig-folder-input',\n        attr: { type: 'text', placeholder: t('mod.imageGallery.folderHint'), value: s.imgFolder || '' }\n    });\n\n    toolbar.createEl('span', { cls: 'ig-sep' });\n\n    toolbar.createEl('label', { text: t('mod.imageGallery.cols'), cls: 'ig-cols-label' });\n    var colsSelect = toolbar.createEl('select', { cls: 'ig-cols-select' });\n    [1, 2, 3, 4, 5, 6].forEach(function(n) {\n        var opt = colsSelect.createEl('option', { text: String(n), attr: { value: n } });\n        if (s.gridCols === n) opt.selected = true;\n    });\n\n    toolbar.createEl('label', { text: t('mod.imageGallery.sort'), cls: 'ig-sort-label' });\n    var sortSelect = toolbar.createEl('select');\n    [\n        { v: 'mtime', t: t('mod.imageGallery.sort.mtime') },\n        { v: 'name', t: t('mod.imageGallery.sort.name') },\n        { v: 'size', t: t('mod.imageGallery.sort.size') }\n    ].forEach(function(o) {\n        var opt = sortSelect.createEl('option', { text: o.t, attr: { value: o.v } });\n        if (s.sortby === o.v) opt.selected = true;\n    });\n\n    var orderBtn = toolbar.createEl('button', { text: s.sort === 'desc' ? t('mod.imageGallery.sort.desc') : t('mod.imageGallery.sort.asc') });\n    var refreshBtn = toolbar.createEl('button', { text: t('mod.imageGallery.refresh') });\n\n    var countLabel = toolbar.createEl('span', { attr: { style: 'font-size:10px;color:var(--text-muted);margin-left:auto;' } });\n    countLabel.style.display = (s.showCount !== false) ? '' : 'none';\n\n    // 间距调节面板\n    var spacingBtn = toolbar.createEl('button', { text: '间距', attr: { style: 'margin-left:6px;' } });\n    var spacingPanel = container.createDiv({ cls: 'ig-spacing-panel' });\n    spacingPanel.style.display = 'none';\n\n    spacingPanel.createEl('label', { text: '左距:' });\n    var leftInput = spacingPanel.createEl('input', { attr: { type: 'number', min: '0', max: '200', value: String(s.spacingLeft || 10) } });\n    spacingPanel.createEl('label', { text: '右距:' });\n    var rightInput = spacingPanel.createEl('input', { attr: { type: 'number', min: '0', max: '200', value: String(s.spacingRight || 10) } });\n    spacingPanel.createEl('label', { text: '间距:' });\n    var gapInput = spacingPanel.createEl('input', { attr: { type: 'number', min: '0', max: '50', value: String(s.itemGap || 4) } });\n\n    function applySpacing() {\n        grid.style.paddingLeft = (s.spacingLeft || 10) + 'px';\n        grid.style.paddingRight = (s.spacingRight || 10) + 'px';\n        var gap = s.itemGap || 4;\n        grid.querySelectorAll('.ig-square-cell').forEach(function(cell) {\n            cell.style.margin = gap + 'px';\n        });\n        // 覆盖 CSS 中的 margin:4px（用 inline !important）\n        grid.style.setProperty('--ig-gap', gap + 'px', 'important');\n    }\n\n    spacingBtn.addEventListener('click', function() {\n        var visible = spacingPanel.style.display !== 'none';\n        spacingPanel.style.display = visible ? 'none' : 'flex';\n        spacingBtn.classList.toggle('active', !visible);\n    });\n\n    [leftInput, rightInput, gapInput].forEach(function(inp) { inp.addEventListener('input', function() {\n        s.spacingLeft = parseInt(leftInput.value) || 0;\n        s.spacingRight = parseInt(rightInput.value) || 0;\n        s.itemGap = parseInt(gapInput.value) || 0;\n        applySpacing();\n        if (typeof saveCallback === 'function') saveCallback();\n    }); });\n\n    // 网格容器\n    var grid = container.createDiv({ cls: 'ig-grid-square' });\n\n    var images = [];\n\n    function updateColsVisibility() {\n        var show = s.displayMode === 'square';\n        container.querySelectorAll('.ig-cols-label, .ig-cols-select').forEach(function(el) {\n            el.style.display = show ? '' : 'none';\n        });\n    }\n\n    // 强制正方形：block + text-align:center + inline-block，不依赖flex容器宽度\n    function enforceSquareGrid() {\n        grid.style.cssText = '';\n        var st = grid.style;\n        st.setProperty('display', 'block', 'important');\n        st.setProperty('width', '100%', 'important');\n        st.setProperty('text-align', 'center', 'important');\n        st.setProperty('font-size', '0', 'important');\n        applySpacing();\n        updateColsVisibility();\n    }\n\n    function applyDisplayMode() {\n        // 先清除所有 inline 样式\n        grid.style.cssText = '';\n\n        grid.className = 'ig-grid-' + s.displayMode;\n\n        if (s.displayMode === 'square') {\n            enforceSquareGrid();\n        } else if (s.displayMode === 'masonry') {\n            grid.style.columnCount = s.gridCols || 3;\n            updateColsVisibility();\n        } else {\n            grid.style.flexWrap = 'wrap';\n            updateColsVisibility();\n        }\n    }\n\n    function scanImages() {\n        images = [];\n        var folderFilter = (s.imgFolder || '').replace(/\\\\/g, '/').replace(/^\\//, '').replace(/\\/$/, '');\n\n        // ★ 定点路径扫描：未设置路径时，不扫描任何文件，提示输入路径\n        if (!folderFilter) {\n            countLabel.textContent = '0' + t('mod.imageGallery.count');\n            grid.innerHTML = '<div class=\"ig-empty\">\\uD83D\\uDCC1 ' + t('mod.imageGallery.folderHint') + '</div>';\n            return;\n        }\n\n        app.vault.getFiles().forEach(function(f) {\n            var ext = f.extension.toLowerCase();\n            if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'ico'].indexOf(ext) >= 0) {\n                if (folderFilter && !f.path.startsWith(folderFilter + '/') && f.path !== folderFilter) return;\n                images.push({ path: f.path, name: f.name, mtime: f.stat ? f.stat.mtime : 0, size: f.stat ? f.stat.size : 0 });\n            }\n        });\n\n        var sortBy = s.sortby || 'mtime';\n        images.sort(function(a, b) {\n            var cmp = 0;\n            if (sortBy === 'name') cmp = a.name.localeCompare(b.name);\n            else if (sortBy === 'size') cmp = a.size - b.size;\n            else cmp = a.mtime - b.mtime;\n            return s.sort === 'asc' ? cmp : -cmp;\n        });\n\n        countLabel.textContent = images.length + t('mod.imageGallery.count');\n        renderGrid();\n    }\n\n    // 正方形模式：inline-block + 固定宽高，由容器 text-align:center 整体居中，间距由 s.itemGap 控制\n    function makeSquare(thumb) {\n        var gap = s.itemGap || 4;\n        thumb.style.cssText = [\n            'display:inline-block',\n            'width:100px',\n            'height:100px',\n            'margin:' + gap + 'px',\n            'vertical-align:top',\n            'overflow:hidden',\n            'border-radius:4px',\n            'cursor:pointer',\n            'border:2px solid transparent',\n            'transition:border-color 0.15s,transform 0.15s',\n            'background:var(--background-modifier-form-field)',\n        ].join(';') + ';';\n        return thumb;\n    }\n\n    function renderGrid() {\n        grid.innerHTML = '';\n        if (images.length === 0) {\n            var emptyMsg = (s.imgFolder || '') ?\n                t('mod.imageGallery.empty.folder') + s.imgFolder + t('mod.imageGallery.empty.folder2') :\n                t('mod.imageGallery.empty.vault');\n            grid.innerHTML = '<div class=\"ig-empty\">' + emptyMsg + '<br><small>' + t('mod.imageGallery.formats') + '</small></div>';\n            return;\n        }\n\n        var isSquare = s.displayMode === 'square';\n\n        images.forEach(function(img, idx) {\n            var thumb = grid.createDiv({ cls: 'ig-thumb' });\n\n            if (isSquare) {\n                // 正方形模式：清除 class 的默认样式，完全由 JS 控制\n                thumb.className = 'ig-thumb ig-square-cell';\n                makeSquare(thumb);\n            }\n\n            // 安全点击\n            thumb.addEventListener('mousedown', function(evt) { evt.preventDefault(); evt.stopPropagation(); }, true);\n            thumb.addEventListener('click', function(evt) {\n                evt.preventDefault(); evt.stopPropagation(); evt.stopImmediatePropagation();\n                openLightbox(idx);\n            }, true);\n\n            // 加载缩略图\n            var fileObj = app.vault.getAbstractFileByPath(img.path);\n            if (fileObj) {\n                app.vault.readBinary(fileObj).then(function(data) {\n                    var blob = new Blob([data]);\n                    var url = URL.createObjectURL(blob);\n                    var el = thumb.createEl('img', { attr: { src: url, loading: 'lazy', alt: img.name } });\n\n                    if (isSquare) {\n                        el.style.cssText = 'width:100%;height:100%;object-fit:cover;display:block;flex-shrink:0;';\n                    }\n\n                    el.addEventListener('load', function() { URL.revokeObjectURL(url); });\n                }).catch(function() {\n                    var fallback = thumb.createDiv({\n                        cls: 'ig-icon-fallback',\n                        text: '\\uD83D\\uDCF7'\n                    });\n                    fallback.style.cssText = 'display:flex;align-items:center;justify-content:center;font-size:30px;color:var(--text-muted);width:100%;height:100%;';\n                });\n            }\n        });\n    }\n\n    // ===== 灯箱 =====\n    function openLightbox(index) {\n        var lb = document.createElement('div');\n        lb.className = 'ig-lightbox';\n\n        var imgEl = lb.appendChild(document.createElement('img'));\n        var closeBtn = lb.appendChild(document.createElement('span'));\n        closeBtn.className = 'ig-lb-close';\n        closeBtn.textContent = '\\u2715';\n\n        var prevBtn = lb.appendChild(document.createElement('span'));\n        prevBtn.className = 'ig-lb-nav ig-lb-prev';\n        prevBtn.textContent = '\\u2039';\n\n        var nextBtn = lb.appendChild(document.createElement('span'));\n        nextBtn.className = 'ig-lb-nav ig-lb-next';\n        nextBtn.textContent = '\\u203A';\n\n        var counter = lb.appendChild(document.createElement('div'));\n        counter.className = 'ig-lb-counter';\n\n        var caption = lb.appendChild(document.createElement('div'));\n        caption.className = 'ig-lb-caption';\n\n        var currentIdx = index;\n\n        function showImage(idx) {\n            currentIdx = idx;\n            imgEl.style.opacity = '0';\n            setTimeout(function() {\n                var imgData = images[idx];\n                var fileObj = app.vault.getAbstractFileByPath(imgData.path);\n                if (fileObj) {\n                    app.vault.readBinary(fileObj).then(function(data) {\n                        var blob = new Blob([data]);\n                        var url = URL.createObjectURL(blob);\n                        imgEl.src = url;\n                        imgEl.style.opacity = '1';\n                    }).catch(function() { imgEl.style.opacity = '1'; });\n                }\n                caption.textContent = imgData.name;\n                counter.textContent = (idx + 1) + ' / ' + images.length;\n                prevBtn.style.visibility = idx > 0 ? 'visible' : 'hidden';\n                nextBtn.style.visibility = idx < images.length - 1 ? 'visible' : 'hidden';\n            }, 150);\n        }\n\n        showImage(index);\n\n        closeBtn.addEventListener('click', function() { document.body.removeChild(lb); });\n        prevBtn.addEventListener('click', function() { if (currentIdx > 0) showImage(currentIdx - 1); });\n        nextBtn.addEventListener('click', function() { if (currentIdx < images.length - 1) showImage(currentIdx + 1); });\n        lb.addEventListener('click', function(e) { if (e.target === lb) document.body.removeChild(lb); });\n\n        document.addEventListener('keydown', function handler(e) {\n            if (e.key === 'Escape') { document.body.removeChild(lb); document.removeEventListener('keydown', handler); }\n            if (e.key === 'ArrowLeft' && currentIdx > 0) showImage(currentIdx - 1);\n            if (e.key === 'ArrowRight' && currentIdx < images.length - 1) showImage(currentIdx + 1);\n        });\n\n        lb.addEventListener('wheel', function(e) {\n            e.preventDefault();\n            if (e.deltaY > 0 && currentIdx < images.length - 1) showImage(currentIdx + 1);\n            if (e.deltaY < 0 && currentIdx > 0) showImage(currentIdx - 1);\n        });\n\n        document.body.appendChild(lb);\n    }\n\n    // ===== 事件绑定 =====\n\n    modes.forEach(function(m) {\n        modeBtns[m.v].addEventListener('click', function() {\n            s.displayMode = m.v;\n            Object.keys(modeBtns).forEach(function(k) {\n                modeBtns[k].classList.toggle('active', k === m.v);\n            });\n            applyDisplayMode();\n            renderGrid();\n            if (typeof saveCallback === 'function') saveCallback();\n        });\n    });\n\n    folderInput.addEventListener('change', function() {\n        s.imgFolder = folderInput.value.trim();\n        if (typeof saveCallback === 'function') saveCallback();\n        scanImages();\n    });\n    folderInput.addEventListener('keydown', function(e) {\n        if (e.key === 'Enter') {\n            s.imgFolder = folderInput.value.trim();\n            if (typeof saveCallback === 'function') saveCallback();\n            scanImages();\n        }\n    });\n\n    colsSelect.addEventListener('change', function() {\n        s.gridCols = parseInt(colsSelect.value);\n        applyDisplayMode();\n        if (typeof saveCallback === 'function') saveCallback();\n    });\n\n    sortSelect.addEventListener('change', function() {\n        s.sortby = sortSelect.value;\n        if (typeof saveCallback === 'function') saveCallback();\n        scanImages();\n    });\n\n    orderBtn.addEventListener('click', function() {\n        s.sort = s.sort === 'asc' ? 'desc' : 'asc';\n        orderBtn.textContent = s.sort === 'desc' ? t('mod.imageGallery.sort.desc') : t('mod.imageGallery.sort.asc');\n        if (typeof saveCallback === 'function') saveCallback();\n        scanImages();\n    });\n\n    refreshBtn.addEventListener('click', scanImages);\n\n    // 初始化\n    applyDisplayMode();\n    setTimeout(function() { scanImages(); }, 1300);\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.imageGallery.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.imageGallery.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n\n    // 计数条显示开关\n    var s = plugin.settings && plugin.settings.modules && plugin.settings.modules['image-gallery'];\n    if (!s) return;\n\n    var row = containerEl.createDiv({ attr: { style: 'display:flex;align-items:center;justify-content:space-between;padding:10px 0;border-top:1px solid var(--background-modifier-border);' } });\n    row.createEl('div', {\n        attr: { style: 'flex:1;' }\n    }).createEl('span', { text: t('mod.imageGallery.showCount'), attr: { style: 'font-size:13px;color:var(--text-normal);' } });\n    row.querySelector('div').createEl('small', { text: t('mod.imageGallery.showCount.desc'), attr: { style: 'display:block;font-size:11px;color:var(--text-muted);margin-top:2px;' } });\n\n    var toggle = row.createEl('input', { attr: { type: 'checkbox' } });\n    toggle.style.cssText = 'width:16px;height:16px;cursor:pointer;';\n    toggle.checked = (s.showCount !== false);\n    toggle.addEventListener('change', function() {\n        s.showCount = toggle.checked;\n        if (typeof saveCallback === 'function') saveCallback();\n    });\n}\n\n\n// === 自动生成的 onunload 清理函数 ===\nvar _cleanupFns = [];\nmodule.exports.onunload = function() {\n    _cleanupFns.forEach(function(fn){ try{fn();}catch(e){} });\n    _cleanupFns = [];\n};\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "image-tools": "// image-tools 模块 - 图片拖放处理（格式转换/缩放/压缩）\n// 源插件: image-converter\n// 核心功能保留: 拖放图片 → Canvas处理 → 保存到库\nconst id = 'image-tools';\nconst title = t('mod.imageTools');\nconst icon = '🖼️';\n\nconst defaultSettings = {\n    autoRename: false,\n    resizeWidth: 800,\n    quality: 80,\n    format: 'webp'\n};\n\nconst styles = `\n.it-wrap { padding: 10px 12px; display: flex; flex-direction: column; height: 100%; }\n.it-hint { font-size: 11px; color: var(--text-muted); margin-bottom: 8px; }\n.it-dropzone { flex: 1; min-height: 80px; border: 2px dashed var(--background-modifier-border); border-radius: 8px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 6px; color: var(--text-muted); font-size: 12px; transition: border-color 0.2s, background 0.2s; cursor: pointer; }\n.it-dropzone:hover, .it-dropzone.drag-over { border-color: var(--v6-primary); background: rgba(var(--v6-primary-rgb, 232,149,109), 0.05); }\n.it-dropzone.drag-over { border-style: solid; }\n.it-settings { margin-bottom: 8px; flex-shrink: 0; display: flex; flex-wrap: wrap; gap: 6px; align-items: center; font-size: 11px; }\n.it-settings label { color: var(--text-muted); font-size: 10px; margin-right: 2px; }\n.it-settings select, .it-settings input { padding: 2px 6px; border: 1px solid var(--background-modifier-border); border-radius: 3px; background: var(--background-secondary); color: var(--text-normal); font-size: 11px; }\n.it-settings input[type=number] { width: 60px; }\n.it-queue { max-height: 120px; overflow-y: auto; margin-top: 8px; flex-shrink: 0; }\n.it-queue-item { display: flex; align-items: center; gap: 6px; padding: 4px 8px; font-size: 11px; border-bottom: 1px solid var(--background-modifier-border); }\n.it-queue-item:last-child { border-bottom: none; }\n.it-queue-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.it-queue-status { font-size: 10px; color: var(--text-muted); white-space: nowrap; }\n.it-queue-status.done { color: #4caf50; }\n.it-queue-status.error { color: var(--text-error); }\n.it-queue-status.processing { color: var(--v6-primary); }\n.it-status { font-size: 10px; color: var(--text-muted); margin-top: 6px; text-align: center; }\n.it-status.success { color: #4caf50; }\n`;\n\nasync function render(container) {\n    container.addClass('it-wrap');\n    var s = settings;\n\n    // 设置区\n    var settingsBar = container.createDiv({ cls: 'it-settings' });\n\n    var fmtLabel = settingsBar.createEl('label', { text: t('mod.imageTools.format') });\n    var fmtSelect = settingsBar.createEl('select');\n    ['webp', 'jpeg', 'png'].forEach(function(f) {\n        var opt = fmtSelect.createEl('option', { text: f.toUpperCase(), attr: { value: f } });\n        if (s.format === f) opt.selected = true;\n    });\n\n    var wLabel = settingsBar.createEl('label', { text: t('mod.imageTools.width') });\n    var wInput = settingsBar.createEl('input', { attr: { type: 'number', value: s.resizeWidth || 800, min: 50, max: 4000 } });\n\n    var qLabel = settingsBar.createEl('label', { text: t('mod.imageTools.quality') });\n    var qInput = settingsBar.createEl('input', { attr: { type: 'number', value: s.quality || 80, min: 10, max: 100 } });\n\n    // 提示\n    var hint = container.createDiv({ cls: 'it-hint', text: t('mod.imageTools.dropHint') });\n\n    // 拖放区\n    var dropzone = container.createDiv({ cls: 'it-dropzone' });\n    dropzone.innerHTML = t('mod.imageTools.dropZone') + '<br><small>或点击选择文件</small>';\n\n    // 处理队列\n    var queue = container.createDiv({ cls: 'it-queue' });\n    var statusEl = container.createDiv({ cls: 'it-status' });\n\n    // 文件选择\n    var fileInput = document.createElement('input');\n    fileInput.type = 'file';\n    fileInput.accept = 'image/*';\n    fileInput.multiple = true;\n    fileInput.style.display = 'none';\n    document.body.appendChild(fileInput);\n\n    dropzone.addEventListener('click', function() { fileInput.click(); });\n\n    fileInput.addEventListener('change', function() {\n        if (fileInput.files.length > 0) processFiles(Array.from(fileInput.files));\n        fileInput.value = '';\n    });\n\n    // 拖放事件\n    dropzone.addEventListener('dragover', function(e) { e.preventDefault(); dropzone.addClass('drag-over'); });\n    dropzone.addEventListener('dragleave', function() { dropzone.removeClass('drag-over'); });\n    dropzone.addEventListener('drop', function(e) {\n        e.preventDefault();\n        dropzone.removeClass('drag-over');\n        var files = Array.from(e.dataTransfer.files).filter(function(f) { return f.type.startsWith('image/'); });\n        if (files.length > 0) processFiles(files);\n    });\n\n    // 设置变更保存\n    fmtSelect.addEventListener('change', function() { s.format = fmtSelect.value; if (typeof saveCallback === 'function') saveCallback(); });\n    wInput.addEventListener('change', function() { s.resizeWidth = parseInt(wInput.value) || 800; if (typeof saveCallback === 'function') saveCallback(); });\n    qInput.addEventListener('change', function() { s.quality = Math.min(100, Math.max(10, parseInt(qInput.value) || 80)); if (typeof saveCallback === 'function') saveCallback(); });\n\n    // ============ 图片处理引擎 ============\n    async function processFiles(files) {\n        if (files.length === 0) return;\n        var format = s.format || 'webp';\n        var maxWidth = s.resizeWidth || 800;\n        var quality = (s.quality || 80) / 100;\n        var success = 0, fail = 0;\n\n        statusEl.textContent = t('mod.imageTools.processing');\n        statusEl.className = 'it-status';\n        queue.innerHTML = '';\n\n        for (var i = 0; i < files.length; i++) {\n            var file = files[i];\n            var item = queue.createDiv({ cls: 'it-queue-item' });\n            item.createSpan({ text: file.name, cls: 'it-queue-name' });\n            var statusSpan = item.createSpan({ text: t('mod.imageTools.processing'), cls: 'it-queue-status processing' });\n\n            try {\n                var resultBlob = await processImage(file, format, maxWidth, quality);\n                var ext = format === 'jpeg' ? 'jpg' : format;\n                var baseName = file.name.replace(/\\.[^.]+$/, '');\n                var newName = (s.autoRename ? baseName + '_' + Date.now() : baseName) + '.' + ext;\n\n                // 保存到vault\n                var arrayBuf = await resultBlob.arrayBuffer();\n                var targetPath = newName;\n\n                // 检查同名文件\n                var existing = app.vault.getAbstractFileByPath(targetPath);\n                if (existing) {\n                    targetPath = baseName + '_' + Date.now() + '.' + ext;\n                }\n\n                await app.vault.createBinary(targetPath, arrayBuf);\n                statusSpan.textContent = '✓ ' + targetPath;\n                statusSpan.className = 'it-queue-status done';\n                success++;\n            } catch (e) {\n                statusSpan.textContent = '✗ ' + e.message;\n                statusSpan.className = 'it-queue-status error';\n                fail++;\n            }\n        }\n\n        statusEl.textContent = t('mod.imageTools.done') + success + t('mod.imageTools.success') + (fail > 0 ? fail + t('mod.imageTools.failed') : '');\n        statusEl.className = fail > 0 ? 'it-status' : 'it-status success';\n        if (success > 0) new Notice(t('mod.imageTools.saved') + success + t('mod.imageTools.savedSuffix'));\n    }\n\n    function processImage(file, format, maxWidth, quality) {\n        return new Promise(function(resolve, reject) {\n            var img = new Image();\n            var url = URL.createObjectURL(file);\n\n            img.onload = function() {\n                URL.revokeObjectURL(url);\n                var w = img.width, h = img.height;\n\n                // 等比缩放\n                if (w > maxWidth) { h = Math.round(h * maxWidth / w); w = maxWidth; }\n\n                var canvas = document.createElement('canvas');\n                canvas.width = w;\n                canvas.height = h;\n                var ctx = canvas.getContext('2d');\n                ctx.drawImage(img, 0, 0, w, h);\n\n                var mimeType = 'image/' + (format === 'jpeg' ? 'jpeg' : format);\n                canvas.toBlob(function(blob) {\n                    if (blob) resolve(blob);\n                    else reject(new Error(t('mod.imageTools.error.convert')));\n                }, mimeType, quality);\n            };\n\n            img.onerror = function() { URL.revokeObjectURL(url); reject(new Error(t('mod.imageTools.error.load'))); };\n            img.src = url;\n        });\n    }\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.imageTools.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.imageTools.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n    containerEl.createEl('p', {\n        text: t('mod.imageTools.settings.hint'),\n        attr: { style: 'color:var(--text-muted);font-size:11px;' }\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "media-gallery": "// media-gallery 模块 - 媒体画廊（图片/视频/音频缩略图 + 灯箱播放）\n// 源插件: memories(视频画廊)\n// 核心功能保留: 视频帧缩略图提取 + 内联播放 + 拖放上传\n// 展示模式: 正方形(grid) / 瀑布流(masonry) / 全智能(auto)\nconst id = 'media-gallery';\nconst title = t('mod.mediaGallery');\nconst icon = '\\uD83C\\uDFAC';\n\nconst defaultSettings = {\n    scanFolder: '',\n    sortOrder: 'date-desc',\n    gridSize: 200,\n    limit: 50,\n    displayMode: 'square',\n    mediaType: 'all',\n    showCount: true,\n    showUploadZone: true,\n    spacingLeft: 10,\n    spacingRight: 10,\n    itemGap: 4\n};\n\nconst styles = `\n.mg-wrap { padding: 8px 10px; display: block !important; width: 100% !important; box-sizing: border-box !important; }\n.mg-toolbar { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; flex-shrink: 0; flex-wrap: wrap; }\n.mg-toolbar button { padding: 3px 8px; border: 1px solid var(--background-modifier-border); border-radius: 3px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 11px; cursor: pointer; }\n.mg-toolbar button:hover { background: var(--background-modifier-hover); }\n.mg-toolbar button.active { background: var(--interactive-accent); color: var(--text-on-accent); border-color: var(--interactive-accent); }\n.mg-toolbar label { font-size: 10px; color: var(--text-muted); }\n.mg-toolbar input { padding: 3px 6px; border: 1px solid var(--background-modifier-border); border-radius: 3px; background: var(--background-secondary); color: var(--text-normal); font-size: 11px; }\n.mg-toolbar .mg-count { font-size: 10px; color: var(--text-muted); margin-left: auto; }\n.mg-toolbar .mg-folder-input { flex: 1; min-width: 120px; max-width: 200px; }\n.mg-toolbar .mg-sep { width: 1px; height: 20px; background: var(--background-modifier-border); margin: 0 2px; }\n\n/* 间距调节面板 */\n.mg-spacing-panel { display: flex; align-items: center; gap: 6px; padding: 4px 0; margin-bottom: 4px; flex-wrap: wrap; }\n.mg-spacing-panel label { font-size: 10px; color: var(--text-muted); }\n.mg-spacing-panel input { width: 42px; padding: 2px 4px; border: 1px solid var(--background-modifier-border); border-radius: 3px; background: var(--background-secondary); color: var(--text-normal); font-size: 11px; text-align: center; }\n\n/* 正方形网格 — inline-block + text-align:center，不依赖容器flex宽度计算 */\n.mg-grid-square { display: block !important; width: 100% !important; text-align: center !important; font-size: 0 !important; }\n.mg-square-cell { display: inline-block !important; width: 100px; height: 100px; margin: 4px; vertical-align: top !important; overflow: hidden !important; position: relative !important; }\n.mg-square-cell > img { width: 100% !important; height: 100% !important; object-fit: cover !important; display: block !important; position: static !important; }\n.mg-square-cell > .mg-item-icon { display: flex !important; align-items: center !important; justify-content: center !important; width: 100% !important; height: 100% !important; font-size: 36px; }\n/* 瀑布流 */\n.mg-grid-masonry { column-gap: 4px !important; width: 100%; }\n.mg-grid-masonry .mg-item { break-inside: avoid; margin-bottom: 4px; display: block; }\n.mg-grid-masonry .mg-item img { width: 100%; height: auto; display: block; object-fit: contain; }\n\n/* 全智能 */\n.mg-grid-auto { display: flex !important; flex-wrap: wrap !important; gap: 4px !important; align-content: start; width: 100%; }\n.mg-grid-auto .mg-item { flex: 1 1 auto; min-width: 80px; max-width: 300px; }\n.mg-grid-auto .mg-item img { width: 100%; height: auto; display: block; object-fit: contain; }\n\n.mg-item { border-radius: 4px; overflow: hidden; cursor: pointer; border: 2px solid transparent; transition: border-color 0.15s, transform 0.15s; background: var(--background-modifier-form-field); position: relative; }\n.mg-item:hover { border-color: var(--v6-primary); transform: scale(1.03); z-index: 1; }\n.mg-item-type { position: absolute; bottom: 4px; right: 4px; font-size: 9px; background: rgba(0,0,0,0.7); color: #fff; padding: 1px 5px; border-radius: 3px; z-index: 2; }\n.mg-empty { text-align: center; color: var(--text-muted); padding: 30px 20px; font-size: 13px; }\n\n.mg-upload-zone { min-height: 50px; border: 2px dashed var(--background-modifier-border); border-radius: 6px; display: flex; align-items: center; justify-content: center; color: var(--text-muted); font-size: 12px; margin-bottom: 6px; cursor: pointer; transition: border-color 0.2s; flex-shrink: 0; }\n.mg-upload-zone:hover { border-color: var(--v6-primary); }\n\n/* 播放器灯箱 */\n.mg-player { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(0,0,0,0.95); z-index: 99999; display: flex; align-items: center; justify-content: center; flex-direction: column; }\n.mg-player video, .mg-player img { max-width: 90vw; max-height: 80vh; object-fit: contain; border-radius: 4px; }\n.mg-player .mg-close { position: absolute; top: 16px; right: 20px; font-size: 28px; color: #fff; cursor: pointer; opacity: 0.7; z-index: 2; }\n.mg-player .mg-close:hover { opacity: 1; }\n.mg-player .mg-nav { position: absolute; top: 50%; transform: translateY(-50%); font-size: 36px; color: #fff; cursor: pointer; opacity: 0.6; padding: 20px; }\n.mg-player .mg-nav:hover { opacity: 1; }\n.mg-player .mg-prev { left: 10px; }\n.mg-player .mg-next { right: 10px; }\n.mg-player .mg-info { color: #fff; margin-top: 10px; font-size: 12px; }\n`;\n\nasync function render(container) {\n    container.addClass('mg-wrap');\n    // 强制容器宽度=100%（防御 Obsidian 主题覆盖）\n    container.style.setProperty('display', 'block', 'important');\n    container.style.setProperty('width', '100%', 'important');\n    container.style.setProperty('box-sizing', 'border-box', 'important');\n    var s = settings;\n\n    // 上传区\n    var uploadZone = container.createDiv({ cls: 'mg-upload-zone', text: t('mod.mediaGallery.dropZone') });\n    if (s.showUploadZone === false) uploadZone.style.display = 'none';\n\n    // 工具栏\n    var toolbar = container.createDiv({ cls: 'mg-toolbar' });\n\n    toolbar.createEl('label', { text: t('mod.imageGallery.display') });\n    var modes = [\n        { v: 'square', t: t('mod.imageGallery.mode.square') },\n        { v: 'masonry', t: t('mod.imageGallery.mode.masonry') },\n        { v: 'auto', t: t('mod.imageGallery.mode.smart') }\n    ];\n    var modeBtns = {};\n    modes.forEach(function(m) {\n        var btn = toolbar.createEl('button', { text: m.t });\n        if (s.displayMode === m.v) btn.classList.add('active');\n        modeBtns[m.v] = btn;\n    });\n\n    toolbar.createEl('span', { cls: 'mg-sep' });\n\n    toolbar.createEl('label', { text: '\\u7C7B\\u578B:' });\n    var typeBtns = {};\n    var types = [\n        { v: 'all', t: '\\u5168\\u90E8' },\n        { v: 'image', t: '\\u56FE\\u7247' },\n        { v: 'video', t: '\\u89C6\\u9891' },\n        { v: 'audio', t: '\\u97F3\\u9891' }\n    ];\n    // Note: type button labels are not in i18n dictionary, keeping original\n    types.forEach(function(tp) {\n        var btn = toolbar.createEl('button', { text: tp.t });\n        if (s.mediaType === tp.v) btn.classList.add('active');\n        typeBtns[tp.v] = btn;\n    });\n\n    toolbar.createEl('span', { cls: 'mg-sep' });\n\n    toolbar.createEl('label', { text: t('mod.imageGallery.folder') });\n    var folderInput = toolbar.createEl('input', {\n        cls: 'mg-folder-input',\n        attr: { type: 'text', placeholder: t('mod.imageGallery.folderHint'), value: s.scanFolder || '' }\n    });\n\n    var refreshBtn = toolbar.createEl('button', { text: t('mod.imageGallery.refresh') });\n    var countEl = toolbar.createEl('span', { cls: 'mg-count' });\n    countEl.style.display = (s.showCount !== false) ? '' : 'none';\n\n    // 间距调节面板\n    var spacingBtn = toolbar.createEl('button', { text: '间距', attr: { style: 'margin-left:6px;' } });\n    var spacingPanel = container.createDiv({ cls: 'mg-spacing-panel' });\n    spacingPanel.style.display = 'none';\n\n    spacingPanel.createEl('label', { text: '左距:' });\n    var leftInput = spacingPanel.createEl('input', { attr: { type: 'number', min: '0', max: '200', value: String(s.spacingLeft || 10) } });\n    spacingPanel.createEl('label', { text: '右距:' });\n    var rightInput = spacingPanel.createEl('input', { attr: { type: 'number', min: '0', max: '200', value: String(s.spacingRight || 10) } });\n    spacingPanel.createEl('label', { text: '间距:' });\n    var gapInput = spacingPanel.createEl('input', { attr: { type: 'number', min: '0', max: '50', value: String(s.itemGap || 4) } });\n\n    function applySpacing() {\n        grid.style.paddingLeft = (s.spacingLeft || 10) + 'px';\n        grid.style.paddingRight = (s.spacingRight || 10) + 'px';\n        var gap = s.itemGap || 4;\n        grid.querySelectorAll('.mg-square-cell').forEach(function(cell) {\n            cell.style.margin = gap + 'px';\n        });\n    }\n\n    spacingBtn.addEventListener('click', function() {\n        var visible = spacingPanel.style.display !== 'none';\n        spacingPanel.style.display = visible ? 'none' : 'flex';\n        spacingBtn.classList.toggle('active', !visible);\n    });\n\n    [leftInput, rightInput, gapInput].forEach(function(inp) { inp.addEventListener('input', function() {\n        s.spacingLeft = parseInt(leftInput.value) || 0;\n        s.spacingRight = parseInt(rightInput.value) || 0;\n        s.itemGap = parseInt(gapInput.value) || 0;\n        applySpacing();\n        if (typeof saveCallback === 'function') saveCallback();\n    }); });\n\n    // 网格\n    var grid = container.createDiv({ cls: 'mg-grid-square' });\n\n    var mediaFiles = [];\n    var MediaTypes = {\n        image: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'ico'],\n        video: ['mp4', 'webm', 'mov', 'avi', 'mkv', 'flv'],\n        audio: ['mp3', 'wav', 'ogg', 'flac', 'aac', 'm4a']\n    };\n\n    function getMediaType(ext) {\n        ext = ext.toLowerCase();\n        if (MediaTypes.image.indexOf(ext) >= 0) return 'image';\n        if (MediaTypes.video.indexOf(ext) >= 0) return 'video';\n        if (MediaTypes.audio.indexOf(ext) >= 0) return 'audio';\n        return null;\n    }\n\n    // 正方形模式：inline-block + 固定宽高，间距由 s.itemGap 控制\n    function makeSquareCell(item) {\n        var gap = s.itemGap || 4;\n        item.style.cssText = [\n            'display:inline-block',\n            'width:100px',\n            'height:100px',\n            'margin:' + gap + 'px',\n            'vertical-align:top',\n            'overflow:hidden',\n            'border-radius:4px',\n            'cursor:pointer',\n            'border:2px solid transparent',\n            'transition:border-color 0.15s,transform 0.15s',\n            'background:var(--background-modifier-form-field)',\n            'position:relative',\n        ].join(';') + ';';\n        return item;\n    }\n\n    function applyDisplayMode() {\n        grid.style.cssText = '';\n        grid.className = 'mg-grid-' + s.displayMode;\n\n        if (s.displayMode === 'square') {\n            var st = grid.style;\n            st.setProperty('display', 'block', 'important');\n            st.setProperty('width', '100%', 'important');\n            st.setProperty('text-align', 'center', 'important');\n            st.setProperty('font-size', '0', 'important');\n            applySpacing();\n        } else if (s.displayMode === 'masonry') {\n            grid.style.columnCount = s.gridCols || 3;\n        } else {\n            grid.style.flexWrap = 'wrap';\n        }\n    }\n\n    // 扫描\n    function scanMedia() {\n        mediaFiles = [];\n        var folderFilter = (s.scanFolder || '').replace(/\\\\/g, '/').replace(/^\\//, '').replace(/\\/$/, '');\n        var typeFilter = s.mediaType || 'all';\n\n        // ★ 定点路径扫描：未设置路径时，不扫描任何文件，直接显示空状态\n        if (!folderFilter) {\n            renderGrid();\n            return;\n        }\n\n        app.vault.getFiles().forEach(function(f) {\n            var ext = f.extension.toLowerCase();\n            var type = getMediaType(ext);\n            if (!type) return;\n            if (typeFilter !== 'all' && type !== typeFilter) return;\n            if (folderFilter && !f.path.startsWith(folderFilter + '/') && f.path !== folderFilter) return;\n\n            mediaFiles.push({\n                path: f.path,\n                name: f.name,\n                ext: ext,\n                type: type,\n                mtime: f.stat ? f.stat.mtime : 0\n            });\n        });\n\n        mediaFiles.sort(function(a, b) { return b.mtime - a.mtime; });\n        renderGrid();\n    }\n\n    function renderGrid() {\n        grid.innerHTML = '';\n        var limit = s.limit || 50;\n        var total = mediaFiles.length;\n        countEl.textContent = Math.min(total, limit) + ' / ' + total + t('mod.imageGallery.count');\n\n        if (total === 0) {\n            var msg;\n            if (!(s.scanFolder || '')) {\n                msg = '\\uD83D\\uDCC1 \\u8BF7\\u8F93\\u5165\\u7167\\u7247\\u8DEF\\u5F84';\n            } else {\n                msg = t('mod.imageGallery.empty.folder') + s.scanFolder + t('mod.imageGallery.empty.folder2');\n            }\n            grid.innerHTML = '<div class=\"mg-empty\">' + msg + '<br><small>' + t('mod.imageGallery.formats') + '</small></div>';\n            return;\n        }\n\n        var count = Math.min(total, limit);\n        var isSquare = s.displayMode === 'square';\n\n        for (var i = 0; i < count; i++) {\n            (function(idx) {\n                var f = mediaFiles[idx];\n                var item = grid.createDiv({ cls: 'mg-item' });\n\n                if (isSquare) {\n                    item.className = 'mg-item mg-square-cell';\n                    makeSquareCell(item);\n                }\n\n                // 类型 badge\n                var typeLabel = item.createDiv({ cls: 'mg-item-type', text: f.ext.toUpperCase() });\n\n                if (f.type === 'image') {\n                    loadThumb(f.path, function(url) {\n                        item.createEl('img', { attr: { src: url } });\n                    });\n                } else if (f.type === 'video') {\n                    extractVideoFrame(f.path, function(imgDataUrl) {\n                        if (imgDataUrl) {\n                            item.createEl('img', { attr: { src: imgDataUrl } });\n                        } else {\n                            item.createDiv({ cls: 'mg-item-icon', text: '\\uD83C\\uDFAC' });\n                        }\n                    });\n                } else {\n                    item.createDiv({ cls: 'mg-item-icon', text: '\\uD83C\\uDFB5' });\n                }\n\n                // 安全点击\n                item.addEventListener('mousedown', function(evt) { evt.preventDefault(); evt.stopPropagation(); }, true);\n                item.addEventListener('click', function(evt) {\n                    evt.preventDefault(); evt.stopPropagation(); evt.stopImmediatePropagation();\n                    openPlayer(idx);\n                }, true);\n            })(i);\n        }\n    }\n\n    function loadThumb(filePath, callback) {\n        var fileObj = app.vault.getAbstractFileByPath(filePath);\n        if (!fileObj) return;\n        app.vault.readBinary(fileObj).then(function(data) {\n            var blob = new Blob([data]);\n            var url = URL.createObjectURL(blob);\n            callback(url);\n            setTimeout(function() { URL.revokeObjectURL(url); }, 30000);\n        });\n    }\n\n    function extractVideoFrame(filePath, callback) {\n        var fileObj = app.vault.getAbstractFileByPath(filePath);\n        if (!fileObj) { callback(null); return; }\n        app.vault.readBinary(fileObj).then(function(data) {\n            var blob = new Blob([data], { type: 'video/mp4' });\n            var url = URL.createObjectURL(blob);\n            var video = document.createElement('video');\n            video.crossOrigin = 'anonymous';\n            video.preload = 'metadata';\n            video.muted = true;\n\n            var timeout = setTimeout(function() {\n                URL.revokeObjectURL(url);\n                callback(null);\n            }, 5000);\n\n            video.addEventListener('loadeddata', function() {\n                clearTimeout(timeout);\n                video.currentTime = 1;\n            });\n\n            video.addEventListener('seeked', function() {\n                clearTimeout(timeout);\n                try {\n                    var canvas = document.createElement('canvas');\n                    canvas.width = video.videoWidth || 320;\n                    canvas.height = video.videoHeight || 180;\n                    var ctx = canvas.getContext('2d');\n                    ctx.drawImage(video, 0, 0, canvas.width, canvas.height);\n                    callback(canvas.toDataURL('image/jpeg', 0.7));\n                } catch (e) {\n                    callback(null);\n                }\n                URL.revokeObjectURL(url);\n            });\n\n            video.addEventListener('error', function() {\n                clearTimeout(timeout);\n                URL.revokeObjectURL(url);\n                callback(null);\n            });\n\n            video.src = url;\n        }).catch(function() { callback(null); });\n    }\n\n    // 播放器\n    function openPlayer(index) {\n        var files = mediaFiles;\n        var currentIdx = index;\n\n        var player = document.createElement('div');\n        player.className = 'mg-player';\n\n        var closeBtn = player.appendChild(document.createElement('span'));\n        closeBtn.className = 'mg-close';\n        closeBtn.textContent = '\\u2715';\n\n        var prevBtn = player.appendChild(document.createElement('span'));\n        prevBtn.className = 'mg-nav mg-prev';\n        prevBtn.textContent = '\\u2039';\n\n        var nextBtn = player.appendChild(document.createElement('span'));\n        nextBtn.className = 'mg-nav mg-next';\n        nextBtn.textContent = '\\u203A';\n\n        var info = player.appendChild(document.createElement('div'));\n        info.className = 'mg-info';\n\n        var mediaContainer = player.appendChild(document.createElement('div'));\n        mediaContainer.style.cssText = 'width:100%;height:100%;display:flex;align-items:center;justify-content:center;';\n\n        function showMedia(idx) {\n            currentIdx = idx;\n            mediaContainer.innerHTML = '';\n            var f = files[idx];\n            loadThumb(f.path, function(url) {\n                if (f.type === 'image') {\n                    var img = mediaContainer.appendChild(document.createElement('img'));\n                    img.src = url;\n                } else if (f.type === 'video') {\n                    var video = mediaContainer.appendChild(document.createElement('video'));\n                    video.src = url;\n                    video.controls = true;\n                    video.autoplay = true;\n                    video.style.maxWidth = '90vw';\n                    video.style.maxHeight = '80vh';\n                } else if (f.type === 'audio') {\n                    var audio = mediaContainer.appendChild(document.createElement('audio'));\n                    audio.src = url;\n                    audio.controls = true;\n                    audio.autoplay = true;\n                    audio.style.width = '400px';\n                    var label = mediaContainer.appendChild(document.createElement('div'));\n                    label.style.cssText = 'text-align:center;color:#fff;margin-top:20px;';\n                    label.textContent = '\\uD83C\\uDFB5 ' + f.name;\n                }\n            });\n            info.textContent = (idx + 1) + ' / ' + files.length + ' - ' + f.name;\n            prevBtn.style.visibility = idx > 0 ? 'visible' : 'hidden';\n            nextBtn.style.visibility = idx < files.length - 1 ? 'visible' : 'hidden';\n        }\n\n        showMedia(index);\n\n        closeBtn.addEventListener('click', function() { document.body.removeChild(player); });\n        prevBtn.addEventListener('click', function() { if (currentIdx > 0) showMedia(currentIdx - 1); });\n        nextBtn.addEventListener('click', function() { if (currentIdx < files.length - 1) showMedia(currentIdx + 1); });\n        player.addEventListener('click', function(e) { if (e.target === player) document.body.removeChild(player); });\n\n        document.addEventListener('keydown', function handler(e) {\n            if (e.key === 'Escape') { document.body.removeChild(player); document.removeEventListener('keydown', handler); }\n            if (e.key === 'ArrowLeft' && currentIdx > 0) showMedia(currentIdx - 1);\n            if (e.key === 'ArrowRight' && currentIdx < files.length - 1) showMedia(currentIdx + 1);\n        });\n\n        document.body.appendChild(player);\n    }\n\n    // ===== 上传处理 =====\n    var fileInput = document.createElement('input');\n    fileInput.type = 'file';\n    fileInput.accept = 'image/*,video/*,audio/*';\n    fileInput.multiple = true;\n    fileInput.style.display = 'none';\n    document.body.appendChild(fileInput);\n\n    uploadZone.addEventListener('click', function() { fileInput.click(); });\n    fileInput.addEventListener('change', function() {\n        if (fileInput.files.length > 0) uploadFiles(Array.from(fileInput.files));\n        fileInput.value = '';\n    });\n\n    uploadZone.addEventListener('dragover', function(e) { e.preventDefault(); });\n    uploadZone.addEventListener('drop', function(e) {\n        e.preventDefault();\n        var files = Array.from(e.dataTransfer.files).filter(function(f) {\n            var ext = (f.name.split('.').pop() || '').toLowerCase();\n            return getMediaType(ext) !== null;\n        });\n        if (files.length > 0) uploadFiles(files);\n    });\n\n    async function uploadFiles(files) {\n        var count = 0;\n        for (var i = 0; i < files.length; i++) {\n            try {\n                var buf = await files[i].arrayBuffer();\n                var existing = app.vault.getAbstractFileByPath(files[i].name);\n                var targetPath = existing ? files[i].name.replace(/(\\.[^.]+)$/, '_' + Date.now() + '$1') : files[i].name;\n                await app.vault.createBinary(targetPath, buf);\n                count++;\n            } catch (e) {\n                console.error('[upload failed]', files[i].name, e);\n            }\n        }\n        if (count > 0) { new Notice('\\u5DF2\\u4E0A\\u4F20 ' + count + ' \\u4E2A\\u6587\\u4EF6'); scanMedia(); }\n    }\n\n    // 展示模式切换\n    modes.forEach(function(m) {\n        modeBtns[m.v].addEventListener('click', function() {\n            s.displayMode = m.v;\n            Object.keys(modeBtns).forEach(function(k) {\n                modeBtns[k].classList.toggle('active', k === m.v);\n            });\n            applyDisplayMode();\n            renderGrid();\n            if (typeof saveCallback === 'function') saveCallback();\n        });\n    });\n\n    types.forEach(function(tp) {\n        typeBtns[tp.v].addEventListener('click', function() {\n            s.mediaType = tp.v;\n            Object.keys(typeBtns).forEach(function(k) {\n                typeBtns[k].classList.toggle('active', k === tp.v);\n            });\n            if (typeof saveCallback === 'function') saveCallback();\n            scanMedia();\n        });\n    });\n\n    folderInput.addEventListener('change', function() {\n        s.scanFolder = folderInput.value.trim();\n        if (typeof saveCallback === 'function') saveCallback();\n        scanMedia();\n    });\n    folderInput.addEventListener('keydown', function(e) {\n        if (e.key === 'Enter') {\n            s.scanFolder = folderInput.value.trim();\n            if (typeof saveCallback === 'function') saveCallback();\n            scanMedia();\n        }\n    });\n\n    refreshBtn.addEventListener('click', scanMedia);\n\n    applyDisplayMode();\n    setTimeout(function() { scanMedia(); }, 1500);\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.mediaGallery.settings.title') });\n    containerEl.createEl('p', {\n        text: '\\u5C55\\u793A\\u5E93\\u4E2D\\u56FE\\u7247\\u3001\\u89C6\\u9891\\u3001\\u97F3\\u9891\\u6587\\u4EF6\\u3002\\u89C6\\u9891\\u81EA\\u52A8\\u63D0\\u53D6\\u9996\\u5E27\\u4F5C\\u4E3A\\u7F29\\u7565\\u56FE\\u3002\\u652F\\u6301\\u4E09\\u79CD\\u5C55\\u793A\\u6A21\\u5F0F\\u3001\\u6587\\u4EF6\\u5939\\u7B5B\\u9009\\u548C\\u5A92\\u4F53\\u7C7B\\u578B\\u8FC7\\u6EE4\\u3002\\u70B9\\u51FB\\u4EFB\\u610F\\u5A92\\u4F53\\u6253\\u5F00\\u64AD\\u653E\\u5668\\uFF0C\\u652F\\u6301\\u952E\\u76D8\\u5BFC\\u822A\\u3002\\u62D6\\u653E\\u65B0\\u6587\\u4EF6\\u5230\\u9762\\u677F\\u4E0A\\u4F20\\u3002',\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n\n    var s = plugin.settings && plugin.settings.modules && plugin.settings.modules['media-gallery'];\n    if (!s) return;\n\n    // 通用 toggle 行构造函数\n    function addToggle(labelText, descText, getValue, setValue) {\n        var row = containerEl.createDiv({ attr: { style: 'display:flex;align-items:center;justify-content:space-between;padding:10px 0;border-top:1px solid var(--background-modifier-border);' } });\n        var info = row.createEl('div', { attr: { style: 'flex:1;' } });\n        info.createEl('span', { text: labelText, attr: { style: 'font-size:13px;color:var(--text-normal);' } });\n        info.createEl('small', { text: descText, attr: { style: 'display:block;font-size:11px;color:var(--text-muted);margin-top:2px;' } });\n        var cb = row.createEl('input', { attr: { type: 'checkbox' } });\n        cb.style.cssText = 'width:16px;height:16px;cursor:pointer;';\n        cb.checked = getValue();\n        cb.addEventListener('change', function() { setValue(cb.checked); if (typeof saveCallback === 'function') saveCallback(); });\n    }\n\n    addToggle(\n        t('mod.imageGallery.showCount'),\n        t('mod.imageGallery.showCount.desc'),\n        function() { return s.showCount !== false; },\n        function(v) { s.showCount = v; }\n    );\n    addToggle(\n        t('mod.mediaGallery.showUploadZone'),\n        t('mod.mediaGallery.showUploadZone.desc'),\n        function() { return s.showUploadZone !== false; },\n        function(v) { s.showUploadZone = v; }\n    );\n}\n\n\n// === 自动生成的 onunload 清理函数 ===\nvar _cleanupFns = [];\nmodule.exports.onunload = function() {\n    _cleanupFns.forEach(function(fn){ try{fn();}catch(e){} });\n    _cleanupFns = [];\n};\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "mobile-adapter": "/**\n * mobile-adapter 模块 — 移动端自适应适配器 v12\n *\n * v12 大改：\n * 1. 竖屏/横屏设置完全分开：各自独立设置列数、宽度、高度\n * 2. 竖屏默认 2列 180px；横屏默认 4列 180px\n * 3. 长按卡片弹出\"隐藏\"按钮，可隐藏不合适的模块\n * 4. 内容缩放：zoom 整体缩小卡片内容（日历/画廊等自适应）\n * 5. 设置面板按竖屏/横屏分区，含已隐藏模块管理\n * 6. 旧版 v11 扁平设置自动迁移到 v12 嵌套结构\n * 只需 Ctrl+R，无需 build。\n */\n\nconst id = 'mobile-adapter';\nconst title = 'Mobile Adapter';\nconst icon = '📱';\n\n// ─── i18n 辅助（安全调用 t()） ─────────────────────────\nfunction _t(key, fallback) {\n    try { return t(key); } catch(e) { return fallback || key; }\n}\n\n// ─── 移动端检测 ─────────────────────────────────────────\nfunction _detectMobile() {\n    try {\n        if (typeof app !== 'undefined' && app.isMobile) return true;\n        var obs = require('obsidian');\n        if (obs && obs.Platform && obs.Platform.isMobile) return true;\n    } catch (e) {}\n    if (/Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent)) return true;\n    if (window.innerWidth <= 768) return true;\n    return false;\n}\n\nfunction _isPortrait() {\n    if (screen && screen.orientation) return screen.orientation.type.startsWith('portrait');\n    return window.innerWidth <= window.innerHeight;\n}\n\n// ─── 默认设置（v12 嵌套结构）───────────────────────────\nconst defaultSettings = {\n    adaptLayout: false,\n    portrait: {\n        cols: 2,\n        cardWidth: 180,    // px, 0=自动按列等分；~360px手机/2列=180/列\n        cardHeight: 0,     // px, 0=用 cardHeightVh\n        cardHeightVh: 28   // vh, cardHeight=0 时生效\n    },\n    landscape: {\n        cols: 4,\n        cardWidth: 180,    // ~720px横屏/4列=180/列\n        cardHeight: 0,\n        cardHeightVh: 28\n    },\n    hideHeader: true,\n    hideNavbar: true,           // ★ 隐藏 Obsidian 移动端顶部/底部导航栏\n    contentScale: 0.75,         // 内容缩放比例 (0.4-1.0), zoom 属性\n    hiddenCardTitles: []        // 已隐藏的模块标题列表\n};\n\n// ─── 全局引擎状态 ─────────────────────────────────────\nvar _engineActive = false;\nvar _styleEl = null;\nvar _resizeHandler = null;\nvar _cardObserver = null;\nvar _canvasObserver = null;\nvar _globalCanvasObserver = null;\n\n// 长按相关\nvar _longPressTimer = null;\nvar _longPressTarget = null;\nvar _hideOverlay = null;\nvar LONG_PRESS_MS = 600;\nvar LONG_PRESS_MOVE = 10;\n\n// ─── 设置迁移：v11 扁平 → v12 嵌套 ─────────────────────\nfunction _migrateSettings(mods) {\n    var m = mods['mobile-adapter'];\n    if (!m) return;\n\n    // 已经是 v12 嵌套结构 — 修复旧版错误的默认值\n    if (m.portrait && typeof m.portrait === 'object') {\n        var fixed = false;\n        // 竖屏默认 2 列，横屏默认 4 列，宽度默认 180\n        if (!m.portrait.cols || m.portrait.cols < 2)     { m.portrait.cols = 2; fixed = true; }\n        if (!m.landscape.cols || m.landscape.cols < 4)   { m.landscape.cols = 4; fixed = true; }\n        if (!m.portrait.cardWidth || m.portrait.cardWidth <= 0)   { m.portrait.cardWidth = 180; fixed = true; }\n        if (!m.landscape.cardWidth || m.landscape.cardWidth <= 0) { m.landscape.cardWidth = 180; fixed = true; }\n        if (m.hideNavbar === undefined) { m.hideNavbar = true; fixed = true; }\n        if (fixed && plugin && plugin.saveSettings) {\n            try { plugin.saveSettings(); } catch(e) {}\n        }\n        return;\n    }\n\n    // v11 扁平结构 → 迁移\n    console.log('[mobile-adapter] 检测到旧版设置，自动迁移到 v12 嵌套结构');\n    var oldCols = m.portraitCols || 2;\n    var oldLandCols = m.landscapeCols || 4;\n    var oldW = m.cardWidthPx || 180;   // ★ v12 默认 180，旧版 0 无意义\n    var oldH = m.cardHeightPx || 0;\n    var oldVh = m.cardHeightVh || 28;\n\n    m.portrait = {\n        cols: oldCols,\n        cardWidth: oldW,\n        cardHeight: oldH,\n        cardHeightVh: oldVh\n    };\n    m.landscape = {\n        cols: oldLandCols,\n        cardWidth: oldW,\n        cardHeight: oldH,\n        cardHeightVh: oldVh\n    };\n\n    // 删除旧扁平字段\n    delete m.portraitCols;\n    delete m.landscapeCols;\n    delete m.cardWidthPx;\n    delete m.cardHeightPx;\n    delete m.cardHeightVh;\n\n    // 补 v12 新字段\n    if (m.contentScale === undefined) m.contentScale = 0.75;\n    if (m.hideNavbar === undefined) m.hideNavbar = true;\n    if (!m.hiddenCardTitles) m.hiddenCardTitles = [];\n\n    // ★ 持久化迁移结果\n    if (plugin && plugin.saveSettings) {\n        try { plugin.saveSettings(); } catch(e) {}\n    }\n}\n\n// ─── 实时读取设置（每次调用都读最新值）──────────────────\nfunction _getCurrentSettings() {\n    try {\n        var mods = plugin && plugin.settings && plugin.settings.modules;\n        if (mods && mods['mobile-adapter']) {\n            _migrateSettings(mods);\n            return Object.assign({}, JSON.parse(JSON.stringify(defaultSettings)), mods['mobile-adapter']);\n        }\n    } catch(e) {}\n    return JSON.parse(JSON.stringify(Object.assign({}, defaultSettings, { adaptLayout: true })));\n}\n\n// ─── 获取当前朝向对应的设置 ────────────────────────────\nfunction _getOrientSettings(s) {\n    return _isPortrait() ? s.portrait : s.landscape;\n}\n\n// ─── CSS 构建 ───────────────────────────────────────────\nfunction _buildCSS(s) {\n    var orient = _getOrientSettings(s);\n    var cols = parseInt(orient.cols) || 2;\n    var cardW = parseInt(orient.cardWidth) || 0;\n    var cardH = parseInt(orient.cardHeight) || 0;\n    var cardVh = parseInt(orient.cardHeightVh) || 28;\n    var hideHeader = s.hideHeader !== false;\n    var scale = parseFloat(s.contentScale);\n    if (isNaN(scale) || scale < 0.4) scale = 0.75;\n    if (scale > 1.0) scale = 1.0;\n\n    // 列宽规则\n    var colRule = cardW > 0 ? (cardW + 'px') : '1fr';\n    var cardWidthRule = cardW > 0 ? (cardW + 'px') : '100%';\n    // 高度规则\n    var cardHeightRule = cardH > 0 ? (cardH + 'px') : (cardVh + 'vh');\n\n    var css = [\n        '/* mobile-adapter-v12 */',\n\n        // 1. 视口容器防横向溢出\n        'body, .app-container, .workspace, .workspace-split,',\n        '.workspace-tabs, .workspace-leaf, .workspace-leaf-content,',\n        '.view-content, .v15-view {',\n        '  max-width: 100vw !important;',\n        '  overflow-x: hidden !important;',\n        '  box-sizing: border-box !important;',\n        '}',\n\n        // 2. 视图容器允许纵向滚动\n        '.view-content, .workspace-leaf-content, .v15-view {',\n        '  overflow-y: auto !important;',\n        '  -webkit-overflow-scrolling: touch !important;',\n        '}',\n        // ★ 补偿隐藏的顶部栏（44px），防止内容紧贴状态栏\n        '.v15-view, .view-content {',\n        '  margin-top: 44px !important;',\n        '}',\n\n        // 3. canvas → grid 布局\n        '.v6-canvas {',\n        '  position: relative !important;',\n        '  display: grid !important;',\n        '  grid-template-columns: repeat(' + cols + ', ' + colRule + ') !important;',\n        '  grid-auto-rows: ' + cardHeightRule + ' !important;',\n        '  gap: 0 !important;',\n        '  padding: 0 !important;',\n        '  width: 100% !important;',\n        '  max-width: 100vw !important;',\n        '  height: auto !important;',\n        '  min-height: 0 !important;',\n        '  overflow-x: hidden !important;',\n        '  overflow-y: visible !important;',\n        '  box-sizing: border-box !important;',\n        '}',\n\n        // 4. 卡片\n        '.v6-canvas > .v6-card {',\n        '  position: relative !important;',\n        '  left: auto !important;',\n        '  top: auto !important;',\n        '  width: ' + cardWidthRule + ' !important;',\n        '  min-width: 0 !important;',\n        '  max-width: 100% !important;',\n        '  height: ' + cardHeightRule + ' !important;',\n        '  min-height: 0 !important;',\n        '  max-height: none !important;',\n        '  resize: none !important;',\n        '  overflow: hidden !important;',\n        '  box-sizing: border-box !important;',\n        '  display: flex !important;',\n        '  flex-direction: column !important;',\n        '}',\n\n        // 5. 内容缩放 ★\n        '.v6-canvas > .v6-card .v6-card-content {',\n        '  flex: 1 1 0 !important;',\n        '  height: 0 !important;',\n        '  min-height: 0 !important;',\n        '  overflow: hidden !important;',\n        '  box-sizing: border-box !important;',\n        '}',\n        '.v6-canvas > .v6-card .v6-card-content > * {',\n        '  zoom: ' + scale + ' !important;',\n        '  -moz-transform: scale(' + scale + ');',\n        '  -moz-transform-origin: top left;',\n        '}',\n\n        // 6. 已隐藏的卡片\n        '.v6-canvas > .v6-card[data-mobile-hidden=\"true\"] {',\n        '  display: none !important;',\n        '}',\n    ];\n\n        // NOTE: hideNavbar CSS 由 _injectNavbarCSS / _removeNavbarCSS 独立管理，不在此处\n\n    if (hideHeader) {\n        css = css.concat([\n            '.v6-canvas > .v6-card .v6-card-header { display: none !important; }'\n        ]);\n    } else {\n        css = css.concat([\n            '.v6-canvas > .v6-card .v6-card-header {',\n            '  flex: 0 0 38px !important; height: 38px !important; min-height: 0 !important;',\n            '  overflow: hidden !important; font-size: 11px !important;',\n            '}'\n        ]);\n    }\n\n    return css.join('\\n');\n}\n\n// ★★★ 独立的 navbar 隐藏 —— JS遍历 + CSS兜底 ★★★\nvar _navbarStyleEl = null;\nvar _navbarWatchTimer = null;\nvar HIDE_CLASS = 'mobile-adapter-navbar-hidden';\n\nfunction _injectNavbarCSS() {\n    // 1. CSS：用类名控制隐藏（可撤销）\n    if (!_navbarStyleEl) {\n        _navbarStyleEl = document.createElement('style');\n        _navbarStyleEl.id = 'mobile-adapter-navbar-css';\n        document.head.appendChild(_navbarStyleEl);\n    }\n    _navbarStyleEl.textContent = [\n        '/* mobile-adapter: hide obsidian top header (类驱动，可撤销) */',\n        '.' + HIDE_CLASS + ' { display: none !important; }',\n        // CSS 兜底：直接匹配 desktop 的 view-header\n        '.view-header { display: none !important; }',\n        '.workspace-tab-header-container { display: none !important; }',\n    ].join('\\n');\n\n    // 2. JS 遍历：加类名（不用内联 style）\n    if (_navbarWatchTimer) clearInterval(_navbarWatchTimer);\n    _navbarWatchTimer = setInterval(_hideTopBarByJS, 500);\n    setTimeout(function() {\n        if (_navbarWatchTimer) clearInterval(_navbarWatchTimer);\n        _navbarWatchTimer = setInterval(_hideTopBarByJS, 2000);\n    }, 3000);\n}\n\nfunction _hideTopBarByJS() {\n    try {\n        var appEl = document.querySelector('.app-container');\n        if (appEl) {\n            var children = appEl.children;\n            for (var i = 0; i < children.length; i++) {\n                var el = children[i];\n                var cls = el.className || '';\n                var tag = el.tagName || '';\n                if (cls.indexOf('header') !== -1 || cls.indexOf('toolbar') !== -1 ||\n                    cls.indexOf('navbar') !== -1 || cls.indexOf('titlebar') !== -1 ||\n                    tag === 'HEADER') {\n                    el.classList.add(HIDE_CLASS);\n                }\n            }\n        }\n        // 扫描顶部矮 div\n        var allDivs = document.querySelectorAll('div[class]');\n        for (var j = 0; j < allDivs.length; j++) {\n            var dv = allDivs[j];\n            var dvCls = dv.className || '';\n            if (dvCls.indexOf('header') >= 0 && dv.offsetHeight > 0 && dv.offsetHeight < 80) {\n                var rect = dv.getBoundingClientRect();\n                if (rect.top < 10) {\n                    dv.classList.add(HIDE_CLASS);\n                }\n            }\n        }\n    } catch(e) {}\n}\n\nfunction _restoreNavbarByJS() {\n    // ★ 批量移除所有被加上的隐藏类\n    try {\n        var hidden = document.querySelectorAll('.' + HIDE_CLASS);\n        for (var i = 0; i < hidden.length; i++) {\n            hidden[i].classList.remove(HIDE_CLASS);\n        }\n    } catch(e) {}\n}\n\nfunction _removeNavbarCSS() {\n    if (_navbarStyleEl && _navbarStyleEl.parentNode) {\n        _navbarStyleEl.parentNode.removeChild(_navbarStyleEl);\n    }\n    _navbarStyleEl = null;\n    if (_navbarWatchTimer) {\n        clearInterval(_navbarWatchTimer);\n        _navbarWatchTimer = null;\n    }\n    // ★ 恢复所有被 JS 隐藏的元素\n    _restoreNavbarByJS();\n}\n\nfunction _injectCSS(s) {\n    if (!_styleEl) {\n        _styleEl = document.createElement('style');\n        _styleEl.id = 'mobile-adapter-v12-css';\n        document.head.appendChild(_styleEl);\n    }\n    _styleEl.textContent = _buildCSS(s);\n}\n\nfunction _removeCSS() {\n    var ids = [];\n    for (var i = 5; i <= 12; i++) ids.push('mobile-adapter-v' + i + '-css');\n    ids.forEach(function(sid) {\n        var el = document.getElementById(sid);\n        if (el && el.parentNode) el.parentNode.removeChild(el);\n    });\n    _styleEl = null;\n}\n\n// ─── 获取卡片标题（用于隐藏标识）────────────────────────\nfunction _getCardTitle(card) {\n    var header = card.querySelector('.v6-card-header');\n    if (header) {\n        var t = header.textContent || header.innerText || '';\n        return t.trim();\n    }\n    // fallback: 找 data 属性\n    var dataMod = card.getAttribute('data-module-id') || card.getAttribute('data-module') || '';\n    return dataMod.trim();\n}\n\n// ─── 清除单张卡片内联样式 ──────────────────────────────\nfunction _clearCard(card) {\n    var s = _getCurrentSettings();\n    var orient = _getOrientSettings(s);\n    var cardW = parseInt(orient.cardWidth) || 0;\n    var cardH = parseInt(orient.cardHeight) || 0;\n    var cardVh = parseInt(orient.cardHeightVh) || 28;\n    var isHidden = false;\n\n    // 检查是否该隐藏\n    if (s.hiddenCardTitles && s.hiddenCardTitles.length > 0) {\n        var cardTitle = _getCardTitle(card);\n        if (cardTitle && s.hiddenCardTitles.indexOf(cardTitle) >= 0) {\n            isHidden = true;\n        }\n    }\n    card.setAttribute('data-mobile-hidden', isHidden ? 'true' : 'false');\n\n    ['left','top','width','height','min-width','min-height',\n     'max-width','max-height','resize','overflow','position'].forEach(function(p) {\n        card.style.removeProperty(p);\n    });\n\n    card.style.setProperty('width',      cardW > 0 ? (cardW + 'px') : '100%', 'important');\n    card.style.setProperty('height',     cardH > 0 ? (cardH + 'px') : (cardVh + 'vh'), 'important');\n    card.style.setProperty('min-width',  '0',       'important');\n    card.style.setProperty('max-width',  '100%',     'important');\n    card.style.setProperty('position',   'relative', 'important');\n    card.style.setProperty('left',       'auto',     'important');\n    card.style.setProperty('top',        'auto',     'important');\n    card.style.setProperty('resize',     'none',     'important');\n    card.style.setProperty('box-sizing', 'border-box','important');\n\n    var content = card.querySelector('.v6-card-content');\n    if (content) {\n        ['height','min-height','max-height','overflow','width','max-width'].forEach(function(p) {\n            content.style.removeProperty(p);\n        });\n        content.style.setProperty('width',      '100%',      'important');\n        content.style.setProperty('max-width',  '100%',      'important');\n        content.style.setProperty('overflow-x', 'hidden',    'important');\n        content.style.setProperty('overflow-y', 'hidden',    'important');\n        content.style.setProperty('box-sizing', 'border-box','important');\n    }\n\n    card.querySelectorAll('.v6-card-content *').forEach(function(el) {\n        el.style.setProperty('max-width',  '100%',       'important');\n        el.style.setProperty('box-sizing', 'border-box', 'important');\n    });\n\n    // ★ 绑定长按事件\n    _bindLongPress(card);\n}\n\nfunction _clearAllCards() {\n    document.querySelectorAll('.v6-canvas > .v6-card').forEach(_clearCard);\n}\n\n// ─── 长按隐藏 ───────────────────────────────────────────\nfunction _removeHideOverlay() {\n    if (_hideOverlay && _hideOverlay.parentNode) {\n        _hideOverlay.parentNode.removeChild(_hideOverlay);\n    }\n    _hideOverlay = null;\n}\n\nfunction _showHideButton(card) {\n    _removeHideOverlay();\n\n    var btn = document.createElement('div');\n    btn.textContent = _t('mod.mobileAdapter.hideModuleBtn', '隐藏此模块');\n    btn.style.cssText = [\n        'position:absolute;top:4px;right:4px;z-index:9999;',\n        'padding:4px 10px;border-radius:6px;',\n        'background:rgba(0,0,0,0.75);color:#fff;',\n        'font-size:12px;cursor:pointer;white-space:nowrap;',\n        'pointer-events:auto;'\n    ].join('');\n\n    btn.addEventListener('click', function(e) {\n        e.stopPropagation(); e.preventDefault();\n        var s = _getCurrentSettings();\n        if (!s.hiddenCardTitles) s.hiddenCardTitles = [];\n        var title = _getCardTitle(card);\n        if (title && s.hiddenCardTitles.indexOf(title) < 0) {\n            s.hiddenCardTitles.push(title);\n            // 写回存储\n            try {\n                if (plugin && plugin.settings && plugin.settings.modules) {\n                    if (!plugin.settings.modules['mobile-adapter']) {\n                        plugin.settings.modules['mobile-adapter'] = {};\n                    }\n                    plugin.settings.modules['mobile-adapter'].hiddenCardTitles = s.hiddenCardTitles;\n                    if (plugin.saveSettings) plugin.saveSettings();\n                }\n            } catch(e) {}\n        }\n        _removeHideOverlay();\n        // 立即隐藏\n        card.setAttribute('data-mobile-hidden', 'true');\n    });\n\n    // 点其他地方关闭\n    var closeHandler = function(e) {\n        if (e.target !== btn) {\n            _removeHideOverlay();\n            document.removeEventListener('click', closeHandler);\n        }\n    };\n    setTimeout(function() {\n        document.addEventListener('click', closeHandler);\n    }, 50);\n\n    card.appendChild(btn);\n    _hideOverlay = btn;\n}\n\nfunction _cancelLongPress() {\n    if (_longPressTimer) { clearTimeout(_longPressTimer); _longPressTimer = null; }\n    _longPressTarget = null;\n}\n\nfunction _initLongPress(card) {\n    var startX = 0, startY = 0;\n    var started = false;\n\n    card.addEventListener('touchstart', function(e) {\n        _cancelLongPress();\n        var t = e.touches[0];\n        startX = t.clientX; startY = t.clientY;\n        started = true;\n        _longPressTarget = card;\n        _longPressTimer = setTimeout(function() {\n            if (started) {\n                _showHideButton(card);\n                _longPressTimer = null;\n            }\n        }, LONG_PRESS_MS);\n    }, { passive: true });\n\n    card.addEventListener('touchmove', function(e) {\n        if (!started) return;\n        var t = e.touches[0];\n        var dx = Math.abs(t.clientX - startX);\n        var dy = Math.abs(t.clientY - startY);\n        if (dx > LONG_PRESS_MOVE || dy > LONG_PRESS_MOVE) {\n            _cancelLongPress();\n            started = false;\n        }\n    }, { passive: true });\n\n    card.addEventListener('touchend', function() {\n        _cancelLongPress();\n        started = false;\n    });\n    card.addEventListener('touchcancel', function() {\n        _cancelLongPress();\n        started = false;\n    });\n}\n\nfunction _bindLongPress(card) {\n    if (card._mobileLongPressBound) return;\n    card._mobileLongPressBound = true;\n    _initLongPress(card);\n}\n\n// ─── 布局应用 ─────────────────────────────────────────\nfunction _applyLayout() {\n    var s = _getCurrentSettings();\n    if (!s.adaptLayout) { _removeCSS(); return; }\n    _injectCSS(s);\n    setTimeout(_clearAllCards, 0);\n}\n\n// ─── 监听卡片动态插入 ──────────────────────────────────\nfunction _watchCanvas(canvas) {\n    if (_cardObserver) { _cardObserver.disconnect(); _cardObserver = null; }\n    _cardObserver = new MutationObserver(function(mutations) {\n        var hasNew = false;\n        mutations.forEach(function(m) {\n            m.addedNodes.forEach(function(node) {\n                if (node.nodeType === 1 && node.classList && node.classList.contains('v6-card')) {\n                    hasNew = true;\n                }\n            });\n        });\n        if (hasNew) {\n            clearTimeout(_watchCanvas._timer);\n            _watchCanvas._timer = setTimeout(_clearAllCards, 80);\n        }\n    });\n    _cardObserver.observe(canvas, { childList: true });\n}\n_watchCanvas._timer = null;\n\n// ─── 等待 .v6-canvas ──────────────────────────────────\nfunction _waitForCanvas() {\n    var canvas = document.querySelector('.v6-canvas');\n    if (canvas) {\n        setTimeout(function() { _applyLayout(); _watchCanvas(canvas); }, 300);\n        return;\n    }\n    if (_canvasObserver) { _canvasObserver.disconnect(); _canvasObserver = null; }\n    _canvasObserver = new MutationObserver(function() {\n        var c = document.querySelector('.v6-canvas');\n        if (c) {\n            _canvasObserver.disconnect(); _canvasObserver = null;\n            setTimeout(function() { _applyLayout(); _watchCanvas(c); }, 300);\n        }\n    });\n    _canvasObserver.observe(document.body, { childList: true, subtree: true });\n}\n\n// ─── 全局 canvas 重建监听 ─────────────────────────────\nfunction _startGlobalCanvasWatch() {\n    if (_globalCanvasObserver) { _globalCanvasObserver.disconnect(); _globalCanvasObserver = null; }\n    var _lastCanvas = null;\n    _globalCanvasObserver = new MutationObserver(function() {\n        var c = document.querySelector('.v6-canvas');\n        if (c && c !== _lastCanvas) {\n            _lastCanvas = c;\n            setTimeout(function() { _applyLayout(); _watchCanvas(c); }, 300);\n        }\n    });\n    _globalCanvasObserver.observe(document.body, { childList: true, subtree: true });\n}\n\n// ─── 引擎启停 ──────────────────────────────────────────\nfunction _startEngine() {\n    if (_engineActive) _stopEngine();\n    _engineActive = true;\n    _waitForCanvas();\n    _startGlobalCanvasWatch();\n    _resizeHandler = function() {\n        clearTimeout(_resizeHandler._t);\n        _resizeHandler._t = setTimeout(_applyLayout, 200);\n    };\n    _resizeHandler._t = null;\n    window.addEventListener('resize', _resizeHandler);\n    window.addEventListener('orientationchange', _resizeHandler);\n}\n\nfunction _stopEngine() {\n    _engineActive = false;\n    _removeCSS();\n    _removeNavbarCSS();\n    _removeHideOverlay();\n    if (_resizeHandler) {\n        window.removeEventListener('resize', _resizeHandler);\n        window.removeEventListener('orientationchange', _resizeHandler);\n        _resizeHandler = null;\n    }\n    if (_cardObserver)       { _cardObserver.disconnect();       _cardObserver = null; }\n    if (_canvasObserver)     { _canvasObserver.disconnect();     _canvasObserver = null; }\n    if (_globalCanvasObserver) { _globalCanvasObserver.disconnect(); _globalCanvasObserver = null; }\n}\n\n// ─── ★ 自执行：模块加载时立即启动 ──────────────────────\n(function autoStart() {\n    if (!_detectMobile()) {\n        console.log('[mobile-adapter] 桌面端，跳过');\n        return;\n    }\n    try {\n        if (plugin && plugin.settings) {\n            if (!plugin.settings.modules) plugin.settings.modules = {};\n\n            if (!plugin.settings.modules['mobile-adapter']) {\n                // 首次初始化：写入 v12 默认值\n                plugin.settings.modules['mobile-adapter'] = JSON.parse(JSON.stringify(\n                    Object.assign({}, defaultSettings, { adaptLayout: true })\n                ));\n            } else {\n                // 已有设置：迁移 + 确保启动\n                _migrateSettings(plugin.settings.modules);\n                var ma = plugin.settings.modules['mobile-adapter'];\n                ma.adaptLayout = true;\n                if (ma.hideNavbar === undefined || ma.hideNavbar === null) {\n                    ma.hideNavbar = true;\n                    try { plugin.saveSettings(); } catch(e) {}\n                }\n                console.log('[mobile-adapter] hideNavbar=' + ma.hideNavbar + ' adaptLayout=' + ma.adaptLayout);\n            }\n        }\n    } catch(e) {}\n\n    // ★★★ 立即注入 navbar CSS（不依赖 canvas） ★★★\n    var s = _getCurrentSettings();\n    if (s.hideNavbar !== false) {\n        _injectNavbarCSS();\n        console.log('[mobile-adapter] navbar CSS 已立即注入');\n    }\n\n    _startEngine();\n    console.log('[mobile-adapter] v12 启动 竖' + s.portrait.cols + '列/' + s.portrait.cardWidth + 'px '\n        + '横' + s.landscape.cols + '列/' + s.landscape.cardWidth + 'px 缩放' + s.contentScale);\n})();\n\n// ─── 设置面板渲染（顶部提示行）────────────────────────\nasync function render(container) {\n    var isMobile = _detectMobile();\n    container.createEl('span', {\n        text: isMobile\n            ? _t('mod.mobileAdapter.status.mobileEnabled', '📱 Mobile (enabled)')\n            : _t('mod.mobileAdapter.status.desktopDisabled', '🖥️ Desktop (auto-enable on mobile)'),\n        attr: { style: 'font-size:11px;color:var(--text-muted);' }\n    });\n}\n\n// ─── 设置面板：完整配置界面 ────────────────────────────\nfunction renderSettings(containerEl, pluginRef, saveCallback) {\n    containerEl.empty();\n\n    var plug = pluginRef || plugin;\n    if (!plug || !plug.settings) return;\n    if (!plug.settings.modules) plug.settings.modules = {};\n\n    var stored = plug.settings.modules['mobile-adapter'];\n    if (!stored) {\n        stored = JSON.parse(JSON.stringify(defaultSettings));\n        plug.settings.modules['mobile-adapter'] = stored;\n    }\n    // 迁移旧版\n    _migrateSettings(plug.settings.modules);\n\n    var s = plug.settings.modules['mobile-adapter'];\n    // 确保嵌套对象存在\n    if (!s.portrait) s.portrait = JSON.parse(JSON.stringify(defaultSettings.portrait));\n    if (!s.landscape) s.landscape = JSON.parse(JSON.stringify(defaultSettings.landscape));\n    if (!s.hiddenCardTitles) s.hiddenCardTitles = [];\n    if (s.contentScale === undefined) s.contentScale = 0.75;\n    if (s.hideNavbar === undefined) s.hideNavbar = true;\n\n    var isMobile = _detectMobile();\n\n    containerEl.createEl('p', {\n        text: isMobile\n            ? _t('mod.mobileAdapter.setup.mobileActive', '📱 Mobile detected, adapter auto-enabled.')\n            : _t('mod.mobileAdapter.setup.desktopPreview', '🖥️ Desktop. Changes apply on mobile.'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;margin:0 0 12px 0;' }\n    });\n\n    // ── 工具函数 ──\n    function addInput(labelKey, descKey, descExtra, curVal, onChangeFn) {\n        var row = containerEl.createDiv({ attr: { style: 'margin-bottom:8px;' } });\n        var labelRow = row.createDiv({ attr: { style: 'display:flex;align-items:baseline;gap:6px;margin-bottom:2px;' } });\n        labelRow.createEl('span', { text: _t(labelKey), attr: { style: 'font-size:12px;font-weight:600;color:var(--text-normal);' } });\n        if (descKey) {\n            var descText = _t(descKey) + (descExtra || '');\n            labelRow.createEl('span', { text: descText, attr: { style: 'font-size:10px;color:var(--text-faint);' } });\n        }\n        var input = row.createEl('input', {\n            attr: {\n                type: 'number', value: String(curVal),\n                style: 'width:70px;padding:3px 6px;border-radius:5px;border:1px solid var(--background-modifier-border);background:var(--background-secondary);color:var(--text-normal);font-size:12px;'\n            }\n        });\n        var handler = function() { onChangeFn(input.value); };\n        input.addEventListener('change', handler);\n        input.addEventListener('blur',   handler);\n        return input;\n    }\n\n    function addInputText(labelKey, descKey, descExtra, curVal, onChangeFn) {\n        var row = containerEl.createDiv({ attr: { style: 'margin-bottom:8px;' } });\n        var labelRow = row.createDiv({ attr: { style: 'display:flex;align-items:baseline;gap:6px;margin-bottom:2px;' } });\n        labelRow.createEl('span', { text: _t(labelKey), attr: { style: 'font-size:12px;font-weight:600;color:var(--text-normal);' } });\n        if (descKey) {\n            var descText = _t(descKey) + (descExtra || '');\n            labelRow.createEl('span', { text: descText, attr: { style: 'font-size:10px;color:var(--text-faint);' } });\n        }\n        var input = row.createEl('input', {\n            attr: {\n                type: 'text', value: String(curVal || ''),\n                style: 'width:100%;padding:3px 6px;border-radius:5px;border:1px solid var(--background-modifier-border);background:var(--background-secondary);color:var(--text-normal);font-size:12px;'\n            }\n        });\n        var handler = function() { onChangeFn(input.value); };\n        input.addEventListener('change', handler);\n        input.addEventListener('blur',   handler);\n        return input;\n    }\n\n    function addToggleRow(labelKey, descKey, curVal, onChangeFn) {\n        var row = containerEl.createDiv({ attr: { style: 'display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;' } });\n        var left = row.createDiv();\n        left.createEl('span', { text: _t(labelKey), attr: { style: 'font-size:12px;font-weight:600;color:var(--text-normal);' } });\n        if (descKey) left.createEl('div', { text: _t(descKey), attr: { style: 'font-size:10px;color:var(--text-faint);margin-top:1px;' } });\n        var tog = row.createEl('div', { attr: { style:\n            'width:34px;height:18px;border-radius:9px;cursor:pointer;transition:background .2s;flex-shrink:0;position:relative;' +\n            'background:' + (curVal ? 'var(--interactive-accent,#4c8bf5)' : 'var(--background-modifier-border)') + ';'\n        }});\n        var dot = tog.createEl('div', { attr: { style:\n            'width:14px;height:14px;border-radius:50%;background:#fff;position:absolute;top:2px;transition:left .2s;' +\n            'left:' + (curVal ? '18px' : '2px') + ';'\n        }});\n        var state = { val: curVal };\n        tog.addEventListener('click', function() {\n            state.val = !state.val;\n            tog.style.background = state.val ? 'var(--interactive-accent,#4c8bf5)' : 'var(--background-modifier-border)';\n            dot.style.left = state.val ? '18px' : '2px';\n            onChangeFn(state.val);\n        });\n    }\n\n    function addDivider(titleKey) {\n        containerEl.createEl('div', {\n            text: _t(titleKey),\n            attr: { style: 'font-size:11px;font-weight:700;color:var(--text-accent);margin:14px 0 4px 0;border-bottom:1px solid var(--background-modifier-border);padding-bottom:2px;' }\n        });\n    }\n\n    async function save() {\n        try {\n            if (typeof saveCallback === 'function') await saveCallback();\n            else if (plug.saveSettings) await plug.saveSettings();\n        } catch(e) {}\n        if (_engineActive) _applyLayout();\n    }\n\n    // ═══════════════════════════════════════════════════\n    // 竖屏设置\n    // ═══════════════════════════════════════════════════\n    addDivider('mod.mobileAdapter.setup.portraitTitle');\n    addInput('mod.mobileAdapter.setup.cols', 'mod.mobileAdapter.setup.colsDesc', '2', s.portrait.cols || 2, function(v) {\n        s.portrait.cols = Math.max(1, Math.min(6, parseInt(v) || 2)); save();\n    });\n    addInput('mod.mobileAdapter.setup.cardWidth', 'mod.mobileAdapter.setup.cardWidthDesc', '', s.portrait.cardWidth || 180, function(v) {\n        s.portrait.cardWidth = Math.max(0, parseInt(v) || 0); save();\n    });\n    addInput('mod.mobileAdapter.setup.cardHeight', 'mod.mobileAdapter.setup.cardHeightDesc', '', s.portrait.cardHeight || 0, function(v) {\n        s.portrait.cardHeight = Math.max(0, parseInt(v) || 0); save();\n    });\n    addInput('mod.mobileAdapter.setup.cardHeightVh', 'mod.mobileAdapter.setup.cardHeightVhDesc', '', s.portrait.cardHeightVh || 28, function(v) {\n        s.portrait.cardHeightVh = Math.max(10, Math.min(100, parseInt(v) || 28)); save();\n    });\n\n    // ═══════════════════════════════════════════════════\n    // 横屏设置\n    // ═══════════════════════════════════════════════════\n    addDivider('mod.mobileAdapter.setup.landscapeTitle');\n    addInput('mod.mobileAdapter.setup.cols', 'mod.mobileAdapter.setup.colsDesc', '4', s.landscape.cols || 4, function(v) {\n        s.landscape.cols = Math.max(1, Math.min(8, parseInt(v) || 4)); save();\n    });\n    addInput('mod.mobileAdapter.setup.cardWidth', 'mod.mobileAdapter.setup.cardWidthDesc', '', s.landscape.cardWidth || 180, function(v) {\n        s.landscape.cardWidth = Math.max(0, parseInt(v) || 0); save();\n    });\n    addInput('mod.mobileAdapter.setup.cardHeight', 'mod.mobileAdapter.setup.cardHeightDesc', '', s.landscape.cardHeight || 0, function(v) {\n        s.landscape.cardHeight = Math.max(0, parseInt(v) || 0); save();\n    });\n    addInput('mod.mobileAdapter.setup.cardHeightVh', 'mod.mobileAdapter.setup.cardHeightVhDesc', '', s.landscape.cardHeightVh || 28, function(v) {\n        s.landscape.cardHeightVh = Math.max(10, Math.min(100, parseInt(v) || 28)); save();\n    });\n\n    // ═══════════════════════════════════════════════════\n    // 通用设置\n    // ═══════════════════════════════════════════════════\n    addDivider('mod.mobileAdapter.setup.generalTitle');\n    addToggleRow('mod.mobileAdapter.hideHeader', 'mod.mobileAdapter.hideHeaderDesc',\n        s.hideHeader !== false,\n        function(v) { s.hideHeader = v; save(); }\n    );\n\n    addToggleRow('mod.mobileAdapter.hideNavbar', 'mod.mobileAdapter.hideNavbarDesc',\n        s.hideNavbar !== false,\n        function(v) {\n            s.hideNavbar = v;\n            save();\n            // ★ 即时生效：注入或移除 navbar CSS\n            if (v) {\n                _injectNavbarCSS();\n            } else {\n                _removeNavbarCSS();\n            }\n        }\n    );\n\n    addInput('mod.mobileAdapter.setup.contentScale', 'mod.mobileAdapter.setup.contentScaleDesc', '', s.contentScale || 0.75, function(v) {\n        var n = parseFloat(v);\n        if (isNaN(n)) n = 0.75;\n        s.contentScale = Math.max(0.4, Math.min(1.0, n)); save();\n    });\n\n    // ── 立即应用（仅移动端） ──\n    if (isMobile) {\n        var applyBtn = containerEl.createEl('button', {\n            text: _t('mod.mobileAdapter.setup.applyBtn', 'Apply Now'),\n            attr: { style: 'margin-top:10px;padding:6px 18px;border-radius:6px;border:none;background:var(--interactive-accent,#4c8bf5);color:#fff;font-size:13px;cursor:pointer;' }\n        });\n        applyBtn.addEventListener('click', function() {\n            _stopEngine();\n            _startEngine();\n            applyBtn.textContent = _t('mod.mobileAdapter.setup.applied', '✅ Applied');\n            setTimeout(function() { applyBtn.textContent = _t('mod.mobileAdapter.setup.applyBtn', 'Apply Now'); }, 1500);\n        });\n    }\n\n    // ═══════════════════════════════════════════════════\n    // 已隐藏模块管理\n    // ═══════════════════════════════════════════════════\n    addDivider('mod.mobileAdapter.setup.hiddenTitle');\n    var hidden = s.hiddenCardTitles || [];\n    if (hidden.length === 0) {\n        containerEl.createEl('p', {\n            text: _t('mod.mobileAdapter.setup.hiddenEmpty', 'No hidden modules. Long-press a card on mobile.'),\n            attr: { style: 'font-size:11px;color:var(--text-faint);margin:4px 0;' }\n        });\n    } else {\n        hidden.forEach(function(title, idx) {\n            var row = containerEl.createDiv({ attr: { style: 'display:flex;align-items:center;justify-content:space-between;padding:4px 0;border-bottom:1px solid var(--background-modifier-border);' } });\n            row.createEl('span', { text: title, attr: { style: 'font-size:11px;color:var(--text-muted);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;' } });\n            var restoreBtn = row.createEl('button', {\n                text: _t('mod.mobileAdapter.setup.restoreBtn', 'Restore'),\n                attr: { style: 'padding:2px 8px;border-radius:4px;border:none;background:var(--interactive-accent,#4c8bf5);color:#fff;font-size:10px;cursor:pointer;flex-shrink:0;margin-left:6px;' }\n            });\n            restoreBtn.addEventListener('click', function() {\n                s.hiddenCardTitles.splice(idx, 1);\n                save();\n                renderSettings(containerEl, pluginRef, saveCallback);\n            });\n        });\n    }\n\n    // 底部提示\n    containerEl.createEl('p', {\n        text: _t('mod.mobileAdapter.setup.tip', '💡 Tip: Width 0 = auto; enter px for fixed.'),\n        attr: { style: 'font-size:10px;color:var(--text-faint);margin-top:12px;line-height:1.5;' }\n    });\n}\n\n// ─── 卸载 ───────────────────────────────────────────────\nfunction onunload() {\n    _stopEngine();\n}\n\nmodule.exports = { id, title, icon, defaultSettings, render, renderSettings, onunload };\n",
  "news": "/**\n * 新闻模块 V15 - AI HOT RSS (全新UI)\n * 格式：V14（含 id/styles/renderSettings）\n */\nconst id = 'news';\nconst title = (function() {\n    try { if (typeof t === 'function') { var v = t('mod.news.title'); if (v && v !== 'mod.news.title') return v; } } catch(e) {}\n    return _isZh() ? '资讯' : 'News';\n})();\nconst icon = '🔥';\n\n/** ★ i18n 辅助 — 多路径回退确保可靠性 */\nfunction _isZh() {\n    // ★ 最优先：用 t() 检测，与 UI 标签的 t() 共享同一个 _currentLang\n    try {\n        if (typeof t === 'function') {\n            var v = t('mod.news.title');\n            if (v === '资讯') return true;\n            if (v === 'News') return false;\n        }\n    } catch(e) {}\n    try {\n        var fn = typeof getCurrentLang === 'function' ? getCurrentLang : null;\n        if (!fn && typeof window !== 'undefined' && window.getCurrentLang) fn = window.getCurrentLang;\n        if (fn) { var l = fn(); if (l) return l.startsWith('zh'); }\n    } catch(e) {}\n    try {\n        var htmlLang = document.documentElement.getAttribute('lang') || '';\n        if (htmlLang) return htmlLang.startsWith('zh');\n    } catch(e) {}\n    return true;\n}\n\n/** 双语字符串获取 */\nfunction _(zh, en) {\n    try {\n        if (typeof t === 'function') {\n            // 尝试通过 i18n key 查找\n            var k = 'mod.news.' + en.toLowerCase().replace(/\\s+/g, '_').replace(/[^a-z_]/g, '');\n            var v = t(k);\n            if (v && v !== k && v !== en) return v;\n        }\n    } catch(e) {}\n    return _isZh() ? zh : en;\n}\n\nconst defaultSettings = {\n    source: 'aihot',\n    pageSize: 10,\n    customFeeds: []  // [{ label_zh: '科技', label_en: 'Tech', url: 'https://...' }]\n};\n\nconst styles = `\n/* Tab 栏 */\n.aihot-tabs {\n    display: flex;\n    gap: 4px;\n    padding: 10px 12px 6px;\n    border-bottom: 1px solid var(--background-modifier-border);\n}\n.aihot-tab {\n    flex: 1;\n    padding: 5px 4px;\n    border: none;\n    background: transparent;\n    border-radius: 6px;\n    cursor: pointer;\n    font-size: 12px;\n    color: var(--text-muted);\n    font-weight: 500;\n    transition: all 0.2s ease;\n    text-align: center;\n}\n.aihot-tab:hover {\n    background: var(--background-modifier-hover);\n    color: var(--text-normal);\n}\n.aihot-tab.active {\n    background: var(--v6-primary);\n    color: white;\n}\n\n/* 文章卡片 */\n.aihot-card {\n    padding: 12px;\n    display: flex;\n    flex-direction: column;\n    height: calc(100% - 80px);\n}\n.aihot-source-badge {\n    display: inline-flex;\n    align-items: center;\n    gap: 4px;\n    font-size: 10px;\n    font-weight: 600;\n    color: var(--v6-primary);\n    background: var(--v6-primary);\n    opacity: 0.15;\n    padding: 2px 8px;\n    border-radius: 10px;\n    margin-bottom: 8px;\n    width: fit-content;\n}\n.aihot-source-badge span {\n    opacity: 6;\n    color: var(--v6-primary);\n}\n.aihot-article-title {\n    font-size: 15px;\n    font-weight: 600;\n    color: var(--text-normal);\n    line-height: 1.45;\n    margin-bottom: 8px;\n    display: -webkit-box;\n    -webkit-line-clamp: 3;\n    -webkit-box-orient: vertical;\n    overflow: hidden;\n}\n.aihot-article-meta {\n    display: flex;\n    align-items: center;\n    gap: 10px;\n    font-size: 11px;\n    color: var(--text-muted);\n    margin-bottom: 10px;\n}\n.aihot-article-meta .dot {\n    width: 3px;\n    height: 3px;\n    border-radius: 50%;\n    background: var(--text-muted);\n    opacity: 0.5;\n}\n.aihot-article-body {\n    flex: 1;\n    overflow: auto;\n    background: var(--background-modifier-form-field);\n    border-radius: 8px;\n    padding: 10px 12px;\n    margin-bottom: 10px;\n}\n.aihot-article-body p {\n    font-size: 13px;\n    color: var(--text-normal);\n    line-height: 1.65;\n    margin: 0;\n    display: -webkit-box;\n    -webkit-line-clamp: 8;\n    -webkit-box-orient: vertical;\n    overflow: hidden;\n}\n\n/* 操作区 */\n.aihot-actions {\n    display: flex;\n    gap: 8px;\n    margin-bottom: 10px;\n}\n.aihot-btn {\n    flex: 1;\n    padding: 8px;\n    border: 1px solid var(--background-modifier-border);\n    background: var(--background-secondary);\n    border-radius: 6px;\n    cursor: pointer;\n    font-size: 12px;\n    color: var(--text-normal);\n    text-align: center;\n    transition: all 0.15s;\n}\n.aihot-btn:hover {\n    background: var(--background-modifier-hover);\n}\n.aihot-btn.primary {\n    background: var(--v6-primary);\n    border-color: var(--v6-primary);\n    color: white;\n}\n.aihot-btn.primary:hover {\n    opacity: 0.9;\n}\n\n/* 导航栏 */\n.aihot-footer {\n    display: flex;\n    align-items: center;\n    justify-content: space-between;\n    padding-top: 8px;\n    border-top: 1px solid var(--background-modifier-border);\n}\n.aihot-footer-btn {\n    padding: 5px 10px;\n    border: none;\n    background: transparent;\n    border-radius: 6px;\n    cursor: pointer;\n    font-size: 11px;\n    color: var(--text-muted);\n    transition: all 0.15s;\n}\n.aihot-footer-btn:hover:not(:disabled) {\n    background: var(--background-modifier-hover);\n    color: var(--text-normal);\n}\n.aihot-footer-btn:disabled {\n    opacity: 0.3;\n    cursor: not-allowed;\n}\n.aihot-footer-counter {\n    font-size: 11px;\n    color: var(--text-muted);\n    font-weight: 500;\n    font-variant-numeric: tabular-nums;\n}\n\n/* 状态 */\n.v5-loading {\n    display: flex;\n    flex-direction: column;\n    align-items: center;\n    justify-content: center;\n    height: 100%;\n    gap: 8px;\n    color: var(--text-muted);\n    font-size: 13px;\n}\n.v5-error {\n    display: flex;\n    flex-direction: column;\n    align-items: center;\n    justify-content: center;\n    height: 100%;\n    text-align: center;\n    padding: 24px 16px;\n    gap: 8px;\n    color: var(--text-error);\n}\n.v5-error .err-title {\n    font-size: 13px;\n    font-weight: 600;\n}\n.v5-error .err-detail {\n    font-size: 11px;\n    color: var(--text-muted);\n    line-height: 1.5;\n    max-width: 100%;\n    word-break: break-all;\n}\n.v5-error .err-retry {\n    margin-top: 4px;\n    padding: 6px 16px;\n    border: none;\n    background: var(--v6-primary);\n    border-radius: 6px;\n    cursor: pointer;\n    font-size: 12px;\n    color: white;\n}\n.v5-warning {\n    padding: 10px 12px;\n    font-size: 11px;\n    color: var(--v6-primary);\n    background: var(--v6-primary);\n    opacity: 0.1;\n    border-radius: 6px;\n    margin: 8px 12px;\n}\n.v5-warning span {\n    opacity: 10;\n    color: var(--v6-primary);\n}\n`;\n\nvar AIHOT_LABELS = { 'featured': '精选', 'all': '全部', 'daily': '日报' };\nvar AIHOT_LABELS_EN = { 'featured': 'Featured', 'all': 'All', 'daily': 'Daily' };\n\nfunction _getAihotFeeds() {\n    var isZh = _isZh();\n    var labels = isZh ? AIHOT_LABELS : AIHOT_LABELS_EN;\n    var feeds = {};\n    feeds[labels.featured] = 'https://aihot.virxact.com/feed.xml';\n    feeds[labels.all] = 'https://aihot.virxact.com/feed/all.xml';\n    feeds[labels.daily] = 'https://aihot.virxact.com/feed/daily.xml';\n    return feeds;\n}\n\nfunction _getAihotLabelKey(zhLabel) {\n    for (var k in AIHOT_LABELS) {\n        if (AIHOT_LABELS[k] === zhLabel) return k;\n    }\n    return 'featured';\n}\n\nfunction _getAihotDefaultFeed(isZh) {\n    return isZh ? AIHOT_LABELS.featured : AIHOT_LABELS_EN.featured;\n}\n\nif (!window._v15NewsState) {\n    window._v15NewsState = {\n        currentFeed: '',      // 运行时动态设置\n        currentFeedUrl: '',   // ★ 当前源的 URL\n        currentIndex: 0,\n        cachedData: null,\n        currentItems: null,\n        customFeedsUsed: false\n    };\n}\n\nfunction parseRSS_DOM(text) {\n    if (typeof DOMParser === 'undefined') throw new Error('DOMParser 不可用');\n    const parser = new DOMParser();\n    const xml = parser.parseFromString(text, 'application/xml');\n    const parseError = xml.querySelector('parsererror');\n    if (parseError) throw new Error('DOMParser 解析 XML 出错');\n\n    const items = [];\n    xml.querySelectorAll('item').forEach(item => {\n        const getText = (sel) => {\n            const el = item.querySelector(sel);\n            return el ? el.textContent.trim() : '';\n        };\n        const description = getText('content\\\\:encoded') || getText('content:encoded') || getText('description');\n        const author = getText('dc\\\\:creator') || getText('dc:creator') || getText('author');\n        items.push({\n            title: getText('title'),\n            link: getText('link'),\n            description: description,\n            pubDate: getText('pubDate'),\n            author: author\n        });\n    });\n    if (items.length === 0) throw new Error('未找到 item 节点');\n    return items;\n}\n\nfunction parseRSS_Regex(text) {\n    const items = [];\n    const itemMatches = text.match(/<item[\\s\\S]*?<\\/item>/gi);\n    if (!itemMatches || itemMatches.length === 0) throw new Error('正则未匹配到 item');\n\n    itemMatches.forEach(itemBlock => {\n        const getTag = (tag) => {\n            const re = new RegExp('<' + tag + '(?:\\\\s[^>]*)?>([\\\\s\\\\S]*?)<\\\\/' + tag + '>', 'i');\n            const m = itemBlock.match(re);\n            return m ? m[1].replace(/<!\\[CDATA\\[|\\]\\]>/g, '').trim() : '';\n        };\n        items.push({\n            title: getTag('title'),\n            link: getTag('link'),\n            description: getTag('content:encoded') || getTag('description'),\n            pubDate: getTag('pubDate'),\n            author: getTag('dc:creator') || getTag('author')\n        });\n    });\n    return items;\n}\n\nfunction parseRSS(text) {\n    try { return parseRSS_DOM(text); }\n    catch (e) { return parseRSS_Regex(text); }\n}\n\nfunction isValidXML(text) {\n    const t = text.trim();\n    return t.startsWith('<?xml') || t.startsWith('<rss') || t.startsWith('<feed');\n}\n\nfunction formatTime(pubDate) {\n    if (!pubDate) return '';\n    try {\n        const m = moment(pubDate);\n        if (m.isValid()) return m.fromNow();\n    } catch (e) {}\n    return pubDate;\n}\n\nfunction stripHtml(html) {\n    if (!html) return '';\n    return html\n        .replace(/<script[^>]*>.*?<\\/script>/gi, '')\n        .replace(/<style[^>]*>.*?<\\/style>/gi, '')\n        .replace(/<[^>]+>/g, ' ')\n        .replace(/\\s+/g, ' ')\n        .trim();\n}\n\nfunction _getAllFeeds() {\n    var feeds = _getAihotFeeds();\n    // 合并用户自定义源\n    var custom = settings.customFeeds || [];\n    for (var i = 0; i < custom.length; i++) {\n        var cf = custom[i];\n        if (cf && cf.url) {\n            var label = _isZh() ? (cf.label_zh || cf.url) : (cf.label_en || cf.url);\n            feeds[label] = cf.url;\n        }\n    }\n    return feeds;\n}\n\nfunction _getFeedUrl(label) {\n    var all = _getAllFeeds();\n    if (all[label]) return all[label];\n    // 回退：查找第一个匹配\n    var keys = Object.keys(all);\n    if (keys.length > 0) return all[keys[0]];\n    var aihot = _getAihotFeeds();\n    return aihot[_getAihotDefaultFeed(true)];\n}\n\nasync function render(content) {\n    var state = window._v15NewsState;\n    var isZh = _isZh();\n\n    // ★ 动态获取当前源的 URL\n    if (!state.currentFeed || state.currentFeedUrl !== _getFeedUrl(state.currentFeed)) {\n        var defaultKey = _getAihotDefaultFeed(isZh);\n        state.currentFeed = defaultKey;\n        state.currentFeedUrl = _getFeedUrl(defaultKey);\n    }\n    var feedUrl = _getFeedUrl(state.currentFeed);\n\n    content.empty();\n    var loading = content.createDiv({ cls: 'v5-loading' });\n    loading.createEl('div', { text: '🔥', attr: { style: 'font-size: 28px;' } });\n    loading.createEl('div', { text: _('加载中...', 'Loading...') });\n\n    try {\n        const res = await requestUrl({\n            url: feedUrl,\n            method: 'GET',\n            headers: {\n                'Accept': 'application/rss+xml, application/xml, text/xml, */*',\n                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'\n            }\n        });\n\n        if (res.status !== 200) {\n            throw new Error('HTTP ' + res.status + (res.text ? ': ' + res.text.substring(0, 80) : ''));\n        }\n\n        const rssText = res.text;\n        if (!rssText) throw new Error(_('响应内容为空', 'Empty response'));\n\n        if (!isValidXML(rssText)) {\n            var preview = rssText.substring(0, 120).replace(/\\s+/g, ' ');\n            throw new Error(_('返回的不是 RSS/XML', 'Response is not RSS/XML') + '\\n' + preview);\n        }\n\n        const items = parseRSS(rssText);\n        if (!items || items.length === 0) throw new Error(_('解析成功但无内容', 'Parsed OK but no items'));\n\n        state.cachedData = items;\n        state.currentItems = items;\n        state.currentIndex = 0;\n\n        content.empty();\n        renderUI(content, state);\n        updateArticle(content, state);\n\n    } catch (e) {\n        content.empty();\n\n        if (state.cachedData && state.cachedData.length > 0) {\n            const warning = content.createDiv({ cls: 'v5-warning' });\n            warning.createEl('span', { text: _('⚠️ 网络异常，显示缓存内容', '⚠️ Network error, showing cached') });\n            state.currentItems = state.cachedData;\n            state.currentIndex = 0;\n            renderUI(content, state);\n            updateArticle(content, state);\n            return;\n        }\n\n        const err = content.createDiv({ cls: 'v5-error' });\n        err.createEl('div', { text: '❌', attr: { style: 'font-size: 28px;' } });\n        err.createEl('div', { text: _('加载失败', 'Load failed'), cls: 'err-title' });\n        err.createEl('div', { text: e.message || _('未知错误', 'Unknown error'), cls: 'err-detail' });\n        const retry = err.createEl('button', { text: _('重新加载', 'Retry'), cls: 'err-retry' });\n        retry.addEventListener('click', () => render(content));\n    }\n}\n\nfunction renderUI(content, state) {\n    // Tab 栏\n    var tabs = content.createDiv({ cls: 'aihot-tabs' });\n    var allFeeds = _getAllFeeds();\n    Object.keys(allFeeds).forEach(function(feedName) {\n        var btn = tabs.createEl('button', {\n            text: feedName,\n            cls: 'aihot-tab' + (state.currentFeed === feedName ? ' active' : '')\n        });\n        btn.addEventListener('click', function() {\n            state.currentFeed = feedName;\n            state.currentFeedUrl = allFeeds[feedName];\n            state.currentIndex = 0;\n            state.cachedData = null;\n            state.currentItems = null;\n            render(content);\n        });\n    });\n\n    // 文章卡片\n    var card = content.createDiv({ cls: 'aihot-card' });\n\n    var badge = card.createDiv({ cls: 'aihot-source-badge' });\n    badge.createEl('span', { text: state.currentFeed });\n\n    card.createEl('h3', { cls: 'aihot-article-title', attr: { 'data-role': 'title' } });\n\n    var meta = card.createDiv({ cls: 'aihot-article-meta' });\n    meta.createEl('span', { attr: { 'data-role': 'author' } });\n    meta.createEl('span', { cls: 'dot' });\n    meta.createEl('span', { attr: { 'data-role': 'time' } });\n\n    var body = card.createDiv({ cls: 'aihot-article-body' });\n    body.createEl('p', { attr: { 'data-role': 'desc' } });\n\n    // 操作按钮\n    var actions = card.createDiv({ cls: 'aihot-actions' });\n    var readBtn = actions.createEl('button', { text: _('查看原文 →', 'Read →'), cls: 'aihot-btn primary' });\n    readBtn.addEventListener('click', function() {\n        var item = state.currentItems[state.currentIndex];\n        if (item && item.link) window.open(item.link, '_blank');\n    });\n\n    // 底部导航\n    var footer = card.createDiv({ cls: 'aihot-footer' });\n    var prevBtn = footer.createEl('button', { text: _('← 上一条', '← Prev'), cls: 'aihot-footer-btn', attr: { 'data-role': 'prev' } });\n    prevBtn.addEventListener('click', function() {\n        if (state.currentIndex > 0) {\n            state.currentIndex--;\n            updateArticle(content, state);\n        }\n    });\n\n    footer.createEl('span', { cls: 'aihot-footer-counter', attr: { 'data-role': 'counter' } });\n\n    var nextBtn = footer.createEl('button', { text: _('下一条 →', 'Next →'), cls: 'aihot-footer-btn', attr: { 'data-role': 'next' } });\n    nextBtn.addEventListener('click', function() {\n        if (state.currentIndex < state.currentItems.length - 1) {\n            state.currentIndex++;\n            updateArticle(content, state);\n        }\n    });\n}\n\nfunction updateArticle(content, state) {\n    const items = state.currentItems;\n    if (!items || items.length === 0) return;\n\n    const item = items[state.currentIndex] || items[0];\n\n    const titleEl = content.querySelector('[data-role=\"title\"]');\n    if (titleEl) titleEl.textContent = item.title || _('无标题', 'Untitled');\n\n    const authorEl = content.querySelector('[data-role=\"author\"]');\n    if (authorEl) authorEl.textContent = item.author || 'AI HOT';\n\n    const timeEl = content.querySelector('[data-role=\"time\"]');\n    if (timeEl) timeEl.textContent = formatTime(item.pubDate);\n\n    const descEl = content.querySelector('[data-role=\"desc\"]');\n    if (descEl) {\n        const text = stripHtml(item.description);\n        descEl.textContent = text.substring(0, 400) + (text.length >= 400 ? '...' : '');\n    }\n\n    const prevBtn = content.querySelector('[data-role=\"prev\"]');\n    const nextBtn = content.querySelector('[data-role=\"next\"]');\n    const counterEl = content.querySelector('[data-role=\"counter\"]');\n\n    if (prevBtn) prevBtn.disabled = state.currentIndex === 0;\n    if (nextBtn) nextBtn.disabled = state.currentIndex >= items.length - 1;\n    if (counterEl) counterEl.textContent = (state.currentIndex + 1) + ' / ' + items.length;\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    const { Setting } = require('obsidian');\n\n    containerEl.createEl('h3', { text: t('mod.news.settings.title') });\n\n    // ★ 默认源（AI HOT 子源选择）\n    new Setting(containerEl)\n        .setName(t('mod.news.settings.default_feed'))\n        .setDesc(t('mod.news.settings.default_feed_desc'))\n        .addDropdown(function(d) {\n            var aihotFeeds = _getAihotFeeds();\n            Object.keys(aihotFeeds).forEach(function(name) { d.addOption(name, name); });\n            // 也加入自定义源\n            var custom = settings.customFeeds || [];\n            for (var i = 0; i < custom.length; i++) {\n                var cf = custom[i];\n                if (cf && cf.url) {\n                    var label = _isZh() ? (cf.label_zh || cf.url) : (cf.label_en || cf.url);\n                    d.addOption(label, label);\n                }\n            }\n            var defaultFeed = settings.defaultFeed || _getAihotDefaultFeed(_isZh());\n            d.setValue(defaultFeed)\n                .onChange(async function(v) {\n                    settings.defaultFeed = v;\n                    window._v15NewsState.currentFeed = v;\n                    window._v15NewsState.currentFeedUrl = _getFeedUrl(v);\n                    await saveCallback();\n                });\n        });\n\n    // ★ 自定义 RSS 源\n    containerEl.createEl('h4', {\n        text: t('mod.news.settings.custom_feeds'),\n        style: 'margin-top:16px;margin-bottom:4px;'\n    });\n    containerEl.createEl('div', {\n        text: t('mod.news.settings.custom_feeds_desc'),\n        style: 'font-size:10px;color:var(--text-muted);margin-bottom:10px;'\n    });\n\n    var custom = settings.customFeeds || [];\n\n    function renderCustomFeeds() {\n        // 清除旧的列表区域（在描述 div 之后的所有兄弟元素）\n        var descEl = containerEl.querySelector('[data-role=\"custom-feeds-desc\"]');\n        if (descEl) {\n            var next = descEl.nextElementSibling;\n            while (next) {\n                var toRemove = next;\n                next = next.nextElementSibling;\n                toRemove.remove();\n            }\n        }\n\n        for (var i = 0; i < custom.length; i++) {\n            (function(idx) {\n                var cf = custom[idx];\n                var row = containerEl.createDiv({ style: 'display:flex;gap:4px;align-items:center;margin-bottom:4px;' });\n                row.createEl('input', {\n                    attr: { placeholder: t('mod.news.settings.zh_label'), value: cf.label_zh || '', style: 'flex:1;min-width:0;' },\n                    type: 'text'\n                }).addEventListener('input', function(e) {\n                    custom[idx].label_zh = e.target.value;\n                    saveCallback();\n                });\n                row.createEl('input', {\n                    attr: { placeholder: t('mod.news.settings.en_label'), value: cf.label_en || '', style: 'flex:1;min-width:0;' },\n                    type: 'text'\n                }).addEventListener('input', function(e) {\n                    custom[idx].label_en = e.target.value;\n                    saveCallback();\n                });\n                row.createEl('input', {\n                    attr: { placeholder: 'RSS URL', value: cf.url || '', style: 'flex:2;min-width:0;' },\n                    type: 'text'\n                }).addEventListener('input', function(e) {\n                    custom[idx].url = e.target.value;\n                    saveCallback();\n                });\n                var delBtn = row.createEl('button', {\n                    text: '✕',\n                    style: 'padding:2px 6px;border:1px solid var(--text-error);background:transparent;color:var(--text-error);border-radius:4px;cursor:pointer;'\n                });\n                delBtn.addEventListener('click', function() {\n                    custom.splice(idx, 1);\n                    settings.customFeeds = custom;\n                    saveCallback();\n                    renderCustomFeeds();\n                });\n            })(i);\n        }\n    }\n\n    // 标记容器\n    var descMarker = containerEl.createDiv({ attr: { 'data-role': 'custom-feeds-desc' } });\n    renderCustomFeeds();\n\n    var addBtn = containerEl.createEl('button', {\n        text: t('mod.news.settings.add_feed'),\n        style: 'margin-top:8px;padding:6px 12px;border:1px solid var(--v6-primary);background:transparent;color:var(--v6-primary);border-radius:6px;cursor:pointer;width:100%;'\n    });\n    addBtn.addEventListener('click', function() {\n        custom.push({ label_zh: '', label_en: '', url: '' });\n        settings.customFeeds = custom;\n        saveCallback();\n        renderCustomFeeds();\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "particle-toy": "/**\n * 粒子玩具模块 V1\n * 鼠标互动粒子效果 — 多模式禅意解压\n */\nconst id = 'particle-toy';\nconst title = t('mod.particleToy');\nconst icon = '✨';\n\nconst defaultSettings = {\n    mode: 'attract',\n    count: 150,\n    color: '#4facfe'\n};\n\nconst styles = `\n.pt-wrap { height:100%; display:flex; flex-direction:column; overflow:hidden; border-radius:12px; }\n.pt-canvas { flex:1; display:block; cursor:crosshair; background:#0f0f1a; border-radius:0 0 12px 12px; touch-action:none; }\n.pt-bar { display:flex; gap:4px; padding:6px 10px; flex-shrink:0; align-items:center;\n    background:rgba(15,15,26,.9); border-radius:8px 8px 0 0; flex-wrap:wrap; }\n.pt-mode-btn { padding:5px 14px; border:1.5px solid rgba(255,255,255,.1); border-radius:20px; cursor:pointer;\n    color:#aaa; font-size:11px; font-weight:600; transition:all .2s ease; background:transparent; }\n.pt-mode-btn.active { color:#fff; border-color:var(--pt-ac,#4facfe); background:rgba(79,172,254,.12);\n    box-shadow:0 0 10px rgba(79,172,254,.15), inset 0 0 6px rgba(79,172,254,.08); }\n.pt-mode-btn:hover:not(.active) { border-color:rgba(255,255,255,.25); color:#ddd; }\n.pt-info { margin-left:auto; color:#666; font-size:10px; }\n`;\n\nasync function render(content) {\n    content.empty();\n    const wrap = content.createDiv({ cls: 'pt-wrap' });\n\n    // Toolbar\n    const bar = wrap.createDiv({ cls: 'pt-bar' });\n\n    const modes = [\n        { id:'attract', icon:'🧲', label:t('mod.pt.attract')||'Attract' },\n        { id:'repel', icon:'💨', label:t('mod.pt.repel')||'Repel' },\n        { id:'flow', icon:'🌊', label:t('mod.pt.flow')||'Flow' },\n        { id:'firework', icon:'🎆', label:t('mod.pt.firework')||'Firework' },\n        { id:'snow', icon:'❄️', label:t('mod.pt.snow')||'Snow' },\n        { id:'galaxy', icon:'🌌', label:t('mod.pt.galaxy')||'Galaxy' },\n        { id:'rainbow', icon:'🌈', label:t('mod.pt.rainbow')||'Rainbow' },\n    ];\n\n    const modeBtns = [];\n    modes.forEach(m => {\n        const btn = bar.createEl('button', { cls:'pt-mode-btn', text:m.icon + ' ' + m.label });\n        btn.dataset.mode = m.id;\n        if (m.id === settings.mode) btn.addClass('active');\n        btn.onclick = () => { modeBtns.forEach(b=>b.removeClass('active')); btn.addClass('active'); settings.mode=m.id; };\n        modeBtns.push(btn);\n    });\n\n    const info = bar.createSpan({ cls: 'pt-info', text:`FPS: -- | ${t('pt.count')||'Particles'}: --` });\n\n    // Canvas\n    const canvas = wrap.createEl('canvas', { cls: 'pt-canvas' });\n\n    let ctx, w, h, particles = [], mouse = { x:0, y:0, down:false }, rafId;\n\n    function resize() {\n        const rect = canvas.parentElement.getBoundingClientRect();\n        canvas.width = rect.width * (window.devicePixelRatio || 1);\n        canvas.height = rect.height * (window.devicePixelRatio || 1);\n        canvas.style.width = rect.width + 'px';\n        canvas.style.height = rect.height + 'px';\n        ctx = canvas.getContext('2d');\n        if (ctx) ctx.scale(window.devicePixelRatio || 1, window.devicePixelRatio || 1);\n        w = rect.width; h = rect.height;\n    }\n    resize();\n\n    window.addEventListener('resize', resize);\n\n    // Particle class\n    class P {\n        constructor(x, y, vx, vy, c, s, life, type) {\n            this.x=x; this.y=y; this.vx=vx; this.vy=vy; this.c=c; this.s=s||Math.random()*3+1; this.life=life||1; this.maxLife=this.life; this.type=type||'circle'; this.angle=Math.random()*Math.PI*2; this.spin=(Math.random()-.5)*0.08; this.alpha=1;\n        }\n        update(mode) {\n            switch(mode) {\n                case 'attract': {\n                    const dx=mouse.x-this.x, dy=mouse.y-this.y, d=Math.sqrt(dx*dx+dy*dy)||1;\n                    this.vx+=dx/d*.8; this.vy+=dy/d*.8; this.vx*=.96; this.vy*=.96;\n                } break;\n                case 'repel': {\n                    const dx=mouse.x-this.x, dy=mouse.y-this.y, d=Math.sqrt(dx*dx+dy*dy)||1;\n                    this.vx-=dx/d*1.5; this.vy-=dy/d*1.5; this.vx*=.95; this.vy*=.95;\n                } break;\n                case 'flow':\n                    this.vx+=Math.sin(this.y*.01+this.x*.008)*.15; this.vy+=Math.cos(this.x*.01+this.y*.008)*.12;\n                    this.vx*=.99; this.vy*=.99;\n                break;\n                case 'firework':\n                    this.vy += .03; this.vx *= .98; this.vy *= .98;\n                    break;\n                case 'snow':\n                    this.vy += .02; this.vx += Math.sin(Date.now()*.001+this.x*.005)*.04;\n                    this.vx *= .999; this.vy *= .999;\n                    break;\n                case 'galaxy':\n                    const cx=w/2, cy=h/2, dx=cx-this.x, dy=cy-this.y, d=Math.sqrt(dx*dx+dy*dy)||1;\n                    this.vx+=dx/d*.05; this.vy+=dy/d*.05; this.vx+=Math.cos(this.y*.02)*.06; this.vy+=Math.sin(this.x*.02)*.06;\n                    this.vx*=.997; this.vy*=.997;\n                    break;\n                case 'rainbow':\n                    this.vx+=(Math.random()-.5)*.3; this.vy+=(Math.random()-.5)*.3;\n                    this.vx*=.95; this.vy*=.95;\n                    this.hue=(this.hue||0)+1;\n                    break;\n            }\n\n            this.x+=this.vx; this.y+=this.vy;\n            this.angle+=this.spin; this.life-=this.maxLife/(settings.count>200?600:400);\n\n            if (this.life<=0 || this.x<-20||this.x>w+20||this.y<-20||this.y+h<20) return false;\n            return true;\n        }\n        draw(ctx) {\n            ctx.save(); ctx.globalAlpha=Math.max(0,this.life/this.maxLife)*this.alpha;\n            ctx.translate(this.x,this.y); ctx.rotate(this.angle);\n\n            let col = this.c;\n            if (settings.mode === 'rainbow') col = `hsl(${(this.hue||0)%360},80%,60%)`;\n            else if (this.c.startsWith('#')) col = this.c;\n\n            ctx.fillStyle = col;\n            ctx.shadowColor = col; ctx.shadowBlur = this.s * 2;\n\n            const size = this.s * (0.5 + this.life / this.maxLife);\n            if (this.type === 'star') {\n                drawStar(ctx, 0, 0, 5, size, size/2, col);\n            } else if (this.type === 'heart') {\n                drawHeart(ctx, 0, 0, size, col);\n            } else {\n                ctx.beginPath(); ctx.arc(0,0,size,0,Math.PI*2); ctx.fill();\n            }\n\n            ctx.restore();\n        }\n    }\n\n    function spawn(x, y, burst) {\n        const n = burst ? Math.floor(settings.count / 8) : Math.ceil(settings.count / 60);\n        for (let i = 0; i < n; i++) {\n            const a = Math.random() * Math.PI * 2;\n            const sp = settings.mode === 'firework' ? (2 + Math.random() * 4) : (0.5 + Math.random() * 2.5);\n            let vx = Math.cos(a) * sp * (burst ? (1 + Math.random() * 2) : 1);\n            let vy = Math.sin(a) * sp * (burst ? (1 + Math.random() * 2) : 1);\n            if (!burst && settings.mode === 'firework') vy -= 1 + Math.random() * 2;\n\n            const types = ['circle','star','heart'];\n            const type = types[Math.floor(Math.random() * types.length)];\n            particles.push(new P(x||w*Math.random(), y||h*Math.random(), vx, vy, settings.color, undefined, undefined, type));\n        }\n    }\n\n    function drawStar(ctx, x, y, spikes, outerR, innerR, col) {\n        ctx.beginPath();\n        for (let i = 0; i < spikes * 2; i++) {\n            const r = i % 2 === 0 ? outerR : innerR;\n            const a = (i * Math.PI / spikes) - Math.PI / 2;\n            ctx.lineTo(x + Math.cos(a) * r, y + Math.sin(a) * r);\n        }\n        ctx.closePath(); ctx.fillStyle = col; ctx.fill();\n    }\n\n    function drawHeart(ctx, x, y, size, col) {\n        ctx.beginPath();\n        ctx.moveTo(x, y + size * 0.35);\n        ctx.bezierCurveTo(x, y - size * 0.5, x - size * 0.5, y - size * 0.3, x, y + size * 0.25);\n        ctx.bezierCurveTo(x + size * 0.5, y - size * 0.3, x, y + size * 0.5, x, y + size * 0.35);\n        ctx.fillStyle = col; ctx.fill();\n    }\n\n    // Init particles\n    for (let i = 0; i < settings.count; i++) particles.push(new P(w*Math.random(), h*Math.random(), (Math.random()-.5)*1, (Math.random()-.5)*1, settings.color));\n\n    // Events\n    canvas.onmousemove = e => {\n        const r = canvas.getBoundingClientRect();\n        mouse.x = e.clientX - r.left; mouse.y = e.clientY - r.top;\n        if (mouse.down && settings.mode !== 'firework') spawn(mouse.x, mouse.y, false);\n    };\n    canvas.onmousedown = e => { mouse.down = true; const r = canvas.getBoundingClientRect(); mouse.x=e.clientX-r.left; mouse.y=e.clientY-r.top; spawn(mouse.x, mouse.y, true); };\n    canvas.onmouseup = () => { mouse.down = false; };\n    canvas.onmouseleave = () => { mouse.down = false; };\n\n    // Touch\n    canvas.ontouchstart = e => { e.preventDefault(); mouse.down=true; const t=e.touches[0]; const r=canvas.getBoundingClientRect();\n        mouse.x=t.clientX-r.left; mouse.y=t.clientY-r.top; spawn(mouse.x,mouse.y,true); };\n    canvas.ontouchmove = e => { e.preventDefault(); const t=e.touches[0]; const r=canvas.getBoundingClientRect();\n        mouse.x=t.clientX-r.left; mouse.y=t.clientY-r.top; if(mouse.down) spawn(mouse.x,mouse.y,false); };\n    canvas.ontouchend = () => { mouse.down=false; };\n\n    // Loop\n    let lastTime = performance.now(), fpsCounter = 0, fpsTime = 0;\n    function loop(now) {\n        ctx.fillStyle = '#0f0f1a';\n        ctx.fillRect(0, 0, w, h);\n\n        // Subtle grid\n        ctx.strokeStyle = 'rgba(255,255,255,.03)';\n        ctx.lineWidth = .5;\n        const gridSize = 40;\n        for (let gx = 0; gx < w; gx += gridSize) { ctx.beginPath(); ctx.moveTo(gx, 0); ctx.lineTo(gx, h); ctx.stroke(); }\n        for (let gy = 0; gy < h; gy += gridSize) { ctx.beginPath(); ctx.moveTo(0, gy); ctx.lineTo(w, gy); ctx.stroke(); }\n\n        // Update & draw\n        particles = particles.filter(p => p.update(settings.mode));\n        while (particles.length < settings.count) {\n            particles.push(new P(w * Math.random(), h * Math.random(), (Math.random()-.5)*.5, (Math.random()-.5)*.5, settings.color));\n        }\n\n        particles.forEach(p => p.draw(ctx));\n\n        // FPS\n        fpsCounter++;\n        if (now - fpsTime >= 1000) { info.textContent = `FPS: ${fpsCounter} | ${t('pt.count')||'Particles'}: ${particles.length}`; fpsCounter = 0; fpsTime = now; }\n\n        rafId = requestAnimationFrame(loop);\n    }\n    loop(performance.now());\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render };\n",
  "pixel-garden": "/**\n * 像素花园模块 V1\n * 8-bit 像素风虚拟花园 — 种植、浇水、开花\n */\nconst id = 'pixel-garden';\nconst title = t('mod.pixelGarden');\nconst icon = '🌻';\n\nconst defaultSettings = {\n    autoGrow: true\n};\n\nconst styles = `\n.pg-wrap { height:100%; display:flex; flex-direction:column; overflow:hidden; font-family:'Courier New',monospace; image-rendering:pixelated; }\n.pg-bar { display:flex; gap:4px; padding:6px 10px; flex-shrink:0; background:rgba(0,0,0,.2); border-radius:8px 8px 0 0; }\n.pg-btn { padding:5px 12px; border:none; border-radius:10px; cursor:pointer; font-size:11px; font-weight:700; transition:all .15s;\n    color:#fff; text-shadow:1px 1px 0 rgba(0,0,0,.4); }\n.pg-btn-plant { background:#2d7d46; box-shadow:0 3px 0 #1a5029,inset 0 -2px 6px rgba(0,0,0,.2); }\n.pg-btn-plant:hover { background:#35924f; transform:translateY(-1px); }\n.pg-btn-water { background:#2980b9; box-shadow:0 3px 0 #1a5276,inset 0 -2px 6px rgba(0,0,0,.2); }\n.pg-btn-water:hover { background:#3498db; transform:translateY(-1px); }\n.pg-btn-harvest { background:#e67e22; box-shadow:0 3px 0 #a55c10,inset 0 -2px 6px rgba(0,0,0,.2); }\n.pg-btn-harvest:hover { background:#f39c12; transform:translateY(-1px); }\n.pg-btn-shop { background:#8e44ad; box-shadow:0 3px 0 #5b2c87,inset 0 -2px 6px rgba(0,0,0,.2); }\n.pg-btn-shop:hover { background:#9b59b6; transform:translateY(-1px); }\n.pg-body { flex:1; position:relative; background:linear-gradient(180deg,#87CEEB 0%,#98D8C8 40%,#7CB342 70%,#558B2F 100%); overflow-y:auto; overflow-x:hidden; border:3px solid #5D4037; border-top:none; border-radius:0 0 10px 10px;\n    image-rendering:pixelated; cursor:crosshair; }\n/* Grid */\n.pg-grid { display:grid; grid-template-columns:repeat(6,1fr); gap:3px; padding:10px; position:relative; z-index:2; }\n.pg-slot { aspect-ratio:1; background:rgba(139,119,101,.35); border:2px solid rgba(93,64,55,.4); border-radius:3px; position:relative; transition:background .2s;\n    display:flex; align-items:center; justify-content:center; font-size:20px; cursor:pointer; min-height:44px; }\n.pg-slot:hover { background:rgba(139,119,101,.55); }\n.pg-slot.tilled { background:rgba(101,67,33,.45); border-color:rgba(160,130,80,.5); }\n.pg-slot.planted { background:rgba(60,120,40,.3); border-color:rgba(100,180,80,.4); animation:pgPlanted .3s ease; }\n@keyframes pgPlanted { 0%{transform:scale(1.1)} 100%{transform:scale(1)} }\n.pg-slot.growing { background:rgba(76,175,80,.25); border-color:rgba(129,199,132,.5);\n    animation:pgGrowing 2s ease-in-out infinite alternate; }\n@keyframes pgGrowing { 0%{box-shadow:inset 0 0 6px rgba(76,175,80,.2)} 100%{box-shadow:inset 0 0 14px rgba(76,175,80,.4)} }\n.pg-slot.blooming { background:rgba(200,230,150,.35); border-color:rgba(156,204,101,.6); animation:pgBloom 1.5s ease-in-out infinite alternate; }\n@keyframes pgBloom { 0%{box-shadow:inset 0 0 10px rgba(255,224,102,.15)} 100%{box-shadow:inset 0 0 18px rgba(255,224,102,.35)} }\n.pg-slot.ready { background:rgba(220,240,180,.4); border-color:rgba(139,195,74,.7); cursor:pointer; animation:pgReady 1s ease-in-out infinite; }\n@keyframes pgReady { 0%,100%{transform:scale(1)} 50%{transform:scale(1.04)} }\n/* Plant sprites (CSS art) */\n.pg-plant-icon { font-size:22px; filter:drop-shadow(1px 1px 0 rgba(0,0,0,.25)); transition:transform .2s; line-height:1; }\n.pg-plant-icon.stage0 { font-size:10px; opacity:.5; } /* seed */\n.pg-plant-icon.stage1 { font-size:13px; animation:pgSprout .8s ease-in-out infinite alternate; } /* sprout */\n@keyframes pgSprout { 0%{transform:scaleY(.85) translateY(2px)} 100%{transform:scaleY(1.1) translateY(-1px)} }\n.pg-plant-icon.stage2 { font-size:17px; } /* growing */\n.pg-plant-icon.stage3 { font-size:22px; animation:pgSway 2s ease-in-out infinite alternate; } /* blooming */\n@keyframes pgSway { 0%{transform:rotate(-4deg)} 100%{transform:rotate(4deg)} }\n/* HUD */\n.pg-hud { position:absolute; top:6px; right:10px; z-index:20; display:flex; flex-direction:column; gap:4px; align-items:flex-end; pointer-events:none; }\n.pg-hud-item { background:rgba(0,0,0,.55); color:#fff; padding:4px 10px; border-radius:8px; font-size:11px;\n    backdrop-filter:blur(4px); border:1px solid rgba(255,255,255,.12); pointer-events:auto; }\n.pg-hud-coins { color:#ffd700; font-weight:700; }\n/* Shop overlay */\n.pg-shop-overlay { position:absolute; inset:0; background:rgba(0,0,0,.65); z-index:50; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(3px);\n    border-radius:0 0 10px 10px; }\n.pg-shop-panel { background:linear-gradient(135deg,#2c3e50,#34495e); border-radius:16px; padding:20px; max-width:320px; width:90%;\n    box-shadow:0 20px 60px rgba(0,0,0,.5), 0 0 0 1px rgba(255,255,255,.1); border:1px solid rgba(255,255,255,.1); }\n.pg-shop-title { color:#ffd700; text-align:center; margin-bottom:14px; font-size:16px; text-shadow:2px 2px 0 rgba(0,0,0,.4); }\n.pg-seed-list { display:grid; grid-template-columns:repeat(3,1fr); gap:8px; }\n.pg-seed-card { background:rgba(255,255,255,.08); border:2px solid rgba(255,255,255,.12); border-radius:10px; padding:10px 6px; text-align:center;\n    cursor:pointer; transition:all .15s; color:#fff; }\n.pg-seed-card:hover { background:rgba(255,255,255,.16); transform:translateY(-2px); border-color:rgba(255,215,0,.4); }\n.pg-seed-icon { font-size:26px; margin-bottom:4px; }\n.pg-seed-name { font-size:10px; margin-bottom:2px; opacity:.85; }\n.pg-seed-price { font-size:11px; color:#ffd700; font-weight:700; }\n/* Water indicator */\n.pg-water-ind { position:absolute; bottom:6px; left:10px; z-index:20; display:flex; align-items:center; gap:4px;\n    background:rgba(0,0,0,.5); padding:4px 10px; border-radius:8px; color:#87CEEB; font-size:11px; }\n`;\n\nconst SEED_TYPES = [\n    { id:'flower', icon:'🌸', name:t('mod.garden.flower')||'Flower', price:5, stages:['🌱','🌿','🌼','🌺'], growTime:[3000,6000,9000,12000] },\n    { id:'tomato', icon:'🍅', name:t('mod.garden.tomato')||'Tomato', price:8, stages:['🌱','🍃','🌿','🍅'], growTime:[4000,8000,14000,20000] },\n    { id:'carrot', icon:'🥕', name:t('mod.garden.carrot')||'Carrot', price:6, stages:['🌱','🌿','🥬','🥕'], growTime:[3500,7000,11000,16000] },\n    { id:'sunflower', icon:'🌻', name:t('mod.garden.sunflower')||'Sunflower', price:10, stages:['🌱','🌿','🌼','🌻'], growTime:[5000,10000,16000,24000] },\n    { id:'cactus', icon:'🌵', name:t('mod.garden.cactus')||'Cactus', price:12, stages:['🌱','🌵','🌵','🌵'], growTime:[8000,16000,28000,99999] },\n    { id:'mushroom', icon:'🍄', name:t('mod.garden.mushroom')||'Mushroom', price:7, stages:['🟤','🟤','🟤','🍄'], growTime:[2500,5000,7500,10000] },\n];\n\nfunction _gk() { return 'pg_garden_v1'; }\nfunction _gs() { try { const d = localStorage.getItem(_gk()); return d ? JSON.parse(d) : null; } catch(e){return null;} }\nfunction _gsave(d) { try { localStorage.setItem(_gk(), JSON.stringify(d)); }catch(e){} }\n\nasync function render(content) {\n    content.empty();\n    const wrap = content.createDiv({ cls: 'pg-wrap' });\n\n    // Toolbar\n    const bar = wrap.createDiv({ cls: 'pg-bar' });\n    const btn = (cls, label, fn) => {\n        const b = bar.createEl('button', { cls:`pg-btn ${cls}`, text:label }); b.onclick=fn; return b;\n    };\n    btn('pg-btn-plant', t('mod.garden.plant')||'Plant', () => doPlant());\n    btn('pg-btn-water', t('mod.garden.waterAll')||'Water All', () => waterAll());\n    btn('pg-btn-harvest', t('mod.garden.harvest')||'Harvest', () => harvestAll());\n    const shopBtn = btn('pg-btn-shop', t('mod.garden.shop')||'Shop', () => toggleShop());\n\n    // Body + Grid\n    const body = wrap.createDiv({ cls: 'pg-body' });\n    const grid = body.createDiv({ cls: 'pg-grid' });\n\n    // HUD\n    const hud = body.createDiv({ cls: 'pg-hud' });\n    const coinEl = hud.createDiv({ cls: 'pg-hud-item pg-hud-coins' });\n\n    // Water ind\n    const wInd = body.createDiv({ cls: 'pg-water-ind' });\n    wInd.innerHTML = '💧 <span id=\"pgWl\">--</span>';\n\n    // Shop panel\n    let shopOverlay = null;\n\n    function toggleShop() {\n        if (shopOverlay) { shopOverlay.remove(); shopOverlay=null; return; }\n        shopOverlay = body.createDiv({ cls: 'pg-shop-overlay' });\n        shopOverlay.onclick=(e)=>{ if(e.target===shopOverlay){shopOverlay.remove();shopOverlay=null;} };\n        const pnl = shopOverlay.createDiv({ cls: 'pg-shop-panel' });\n        pnl.createEl('h3',{cls:'pg-shop-title',text:'🛒 '+(t('mod.garden.seedShop')||'Seed Shop')});\n        const list = pnl.createDiv({ cls: 'pg-seed-list' });\n        SEED_TYPES.forEach(s => {\n            const card = list.createDiv({ cls: 'pg-seed-card' });\n            card.createDiv({ cls: 'pg-seed-icon', text:s.icon });\n            card.createDiv({ cls: 'pg-seed-name', text:s.name });\n            card.createDiv({ cls: 'pg-seed-price', text:`💰${s.price}` });\n            card.onclick = () => buySeed(s.id);\n        });\n        pnl.createEl('button',{text:(t('mod.garden.close')||'Close'),style:'margin-top:12px;width:100%;padding:8px;border-radius:10px;border:1px solid rgba(255,255,255,.2);background:transparent;color:#fff;cursor:pointer;',onclick:()=>{shopOverlay.remove();shopOverlay=null;}});\n    }\n\n    // State\n    let state = _gs() || { coins:50, slots:Array(24).fill(null).map(()=>({})), selected:null, waterLevel:100 };\n\n    function updateHUD() { coinEl.textContent = `💰 ${state.coins}`; document.getElementById('pgWl').textContent = Math.floor(state.waterLevel)+'%'; }\n    updateHUD();\n\n    function buySeed(seedId) {\n        const s = SEED_TYPES.find(x=>x.id===seedId);\n        if (!s || state.coins < s.price) return;\n        state.coins -= s.price;\n        state.selected = s;\n        _gsave(state); updateHUD();\n        // Visual feedback\n        if (shopOverlay) { shopOverlay.remove(); shopOverlay=null; }\n    }\n\n    // Create grid slots\n    for (let i = 0; i < 24; i++) {\n        const slot = grid.createDiv({ cls: 'pg-slot' });\n        slot.dataset.idx = i;\n\n        slot.onclick = () => handleSlotClick(i);\n\n        renderSlot(slot, i);\n    }\n\n    function renderSlot(slot, idx) {\n        const s = state.slots[idx];\n        slot.empty();\n        slot.className = 'pg-slot';\n\n        if (!s || !s.seedId) {\n            if (state.selected) slot.innerHTML = `<span class=\"pg-plant-icon stage0\">+</span>`;\n            return;\n        }\n\n        const st = SEED_TYPES.find(x=>x.id===s.seedId);\n        if (!st) return;\n\n        if (s.stage === 0) { slot.addClass('planted'); slot.innerHTML=`<span class=\"pg-plant-icon stage0\">${st.stages[0]}</span>`; }\n        else if (s.stage === 1) { slot.addClass('growing'); slot.innerHTML=`<span class=\"pg-plant-icon stage1\">${st.stages[1]}</span>`; }\n        else if (s.stage === 2) { slot.addClass('growing'); slot.innerHTML=`<span class=\"pg-plant-icon stage2\">${st.stages[2]}</span>`; }\n        else if (s.stage >= 3) {\n            if (Date.now() - s.plantTime >= st.growTime[3]) { slot.addClass('ready'); slot.innerHTML=`<span class=\"pg-plant-icon stage3\">${st.stages[3]}</span>`; }\n            else { slot.addClass('blooming'); slot.innerHTML=`<span class=\"pg-plant-icon stage3\">${st.stages[2]}</span>`; }\n        }\n    }\n\n    function handleSlotClick(idx) {\n        const s = state.slots[idx];\n        if (!s || !s.seedId) {\n            // Plant mode\n            if (state.selected) {\n                const st = SEED_TYPES.find(x=>x.id===state.selected.id);\n                if (state.coins >= st.price && !s.seedId) {\n                    state.slots[idx] = { seedId:st.id, plantTime:Date.now(), stage:0, watered:false };\n                    // Deduct was done on buy, but re-check\n                    _gsave(state); renderSlot(grid.children[idx], idx);\n                }\n            }\n        } else if (s.stage >= 3) {\n            // Harvest single\n            harvestOne(idx);\n        }\n    }\n\n    function doPlant() {\n        if (state.selected) return; // already have seed\n        toggleShop();\n    }\n\n    function waterAll() {\n        if (state.waterLevel <= 5) return;\n        state.waterLevel = Math.max(0, state.waterLevel - 20);\n        state.slots.forEach((s, i) => { if (s && s.seedId !== undefined) s.watered = true; });\n        _gsave(state); updateHUD();\n\n        // Re-render all planted slots with a visual flash\n        for (let i = 0; i < 24; i++) renderSlot(grid.children[i], i);\n    }\n\n    function harvestOne(idx) {\n        const s = state.slots[idx];\n        if (!s || !s.seedId || s.stage < 3) return;\n        const st = SEED_TYPES.find(x=>x.id===s.seedId);\n        if (!st) return;\n        if (Date.now() - s.plantTime < st.growTime[3]) return;\n\n        // Earn coins!\n        const earn = Math.floor(st.price * (1.5 + Math.random()));\n        state.coins += earn;\n        state.slots[idx] = {};\n        _gsave(state); updateHUD(); renderSlot(grid.children[idx], idx);\n\n        // Show floating coin\n        showFloatingEarn(grid.children[idx], `+${earn}`);\n    }\n\n    function showFloatingEarn(slotEl, txt) {\n        const fl = document.createElement('div');\n        fl.textContent = txt;\n        fl.style.cssText = 'position:absolute;top:-8px;left:50%;transform:translateX(-50%);color:#ffd700;font-weight:700;font-size:14px;text-shadow:0 1px 3px rgba(0,0,0,.5);pointer-events:none;z-index:30;animation:pgFloatUp 1.2s ease-out forwards;';\n        slotEl.appendChild(fl);\n        setTimeout(() => fl.remove(), 1200);\n    }\n\n    function harvestAll() {\n        let earned = 0;\n        for (let i = 0; i < 24; i++) {\n            const s = state.slots[i];\n            if (!s || !s.seedId || s.stage < 3) continue;\n            const st = SEED_TYPES.find(x=>x.id===s.seedId);\n            if (!st) continue;\n            if (Date.now() - s.plantTime < st.growTime[3]) continue;\n            earned += Math.floor(st.price * (1.5 + Math.random()));\n            state.slots[i] = {};\n        }\n        if (earned > 0) {\n            state.coins += earned;\n            _gsave(state); updateHUD();\n            // Re-render all\n            for (let i = 0; i < 24; i++) renderSlot(grid.children[i], i);\n        }\n    }\n\n    // Growth loop — check every 3 seconds\n    setInterval(() => {\n        let changed = false;\n        state.slots.forEach((s, i) => {\n            if (!s || !s.seedId || s.stage >= 3) return;\n            const st = SEED_TYPES.find(x=>x.id===s.seedId);\n            if (!st) return;\n            const elapsed = Date.now() - s.plantTime;\n\n            let newStage = 0;\n            for (let j = 3; j >= 0; j--) {\n                if (elapsed >= st.growTime[j]) { newStage = j+1; break; }\n            }\n            // Water speeds up growth by 1.5x\n            if (s.watered && elapsed >= st.growTime[s.stage] * 0.66) newStage = Math.max(newStage, s.stage + 1);\n\n            if (newStage > s.stage) {\n                s.stage = newStage;\n                s.watered = false;\n                changed = true;\n                renderSlot(grid.children[i], i);\n            }\n        });\n        // Slowly recover water\n        state.waterLevel = Math.min(100, state.waterLevel + 1);\n        if (changed) { _gsave(state); }\n        updateHUD();\n    }, 3000);\n}\n\n// CSS keyframe for floating text\nconst styleEl = document.createElement('style');\nstyleEl.textContent = '@keyframes pgFloatUp{0%{opacity:1;transform:translateY(0) translateX(-50%);}100%{opacity:0;transform:translateY(-28px) translateX(-50%);}}';\ndocument.head.appendChild(styleEl);\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render };\n",
  "ppt-viewer": "// ppt-viewer 模块 - PowerPoint (PPT/PPTX) 演示文稿查看器\n// 核心功能: JSZip 解压 PPTX + 复用 FileViewer 的 PPT OLE2 解析 → 纯文字模式预览\n// ★ 只提取和渲染文字，不解析/渲染图片、形状、动画等媒体内容\n// ★ 模块化设计：删除此文件 = 功能消失；放回即恢复\n// ★ PPT 解析逻辑复用 fv-legacy-office.js 导出的 window.__pptExtractor（避免代码重复）\nconst id = 'ppt-viewer';\nconst title = t('mod.pptViewer');\nconst icon = '📊';\n\nconst defaultSettings = {\n    pptxEnabled: true,\n    pptEnabled: true\n};\n\nconst styles = `\n.ptv-wrap { padding: 8px 0; display: flex; flex-direction: column; height: 100%; }\n.ptv-toolbar { display: flex; align-items: center; gap: 6px; padding: 0 10px; margin-bottom: 6px; flex-shrink: 0; flex-wrap: wrap; }\n.ptv-toolbar button { padding: 3px 8px; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 11px; cursor: pointer; white-space: nowrap; }\n.ptv-toolbar button:hover { background: var(--background-modifier-hover); }\n.ptv-toolbar button.active { background: var(--v6-primary); color: white; border-color: var(--v6-primary); }\n.ptv-filelist { max-height: 100px; overflow-y: auto; margin: 0 10px 6px; flex-shrink: 0; border: 1px solid var(--background-modifier-border); border-radius: 4px; }\n.ptv-filelist.hidden { display: none; }\n.ptv-file-item { display: flex; align-items: center; justify-content: space-between; padding: 4px 8px; cursor: pointer; font-size: 11px; transition: background 0.1s; border-radius: 2px; user-select: none; -webkit-user-select: none; }\n.ptv-file-item:hover { background: var(--background-modifier-hover); }\n.ptv-file-item.selected { background: rgba(var(--v6-primary-rgb, 232,149,109), 0.2); color: var(--v6-primary); }\n.ptv-file-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.ptv-file-type { color: var(--text-faint); font-size: 10px; flex-shrink: 0; margin-left: 8px; background: var(--background-modifier-form-field); padding: 1px 6px; border-radius: 8px; }\n.ptv-viewer { flex: 1; min-height: 0; margin: 0 10px; border: 1px solid var(--background-modifier-border); border-radius: 4px; overflow: auto; background: var(--background-primary); position: relative; }\n.ptv-empty { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; text-align: center; pointer-events: none; }\n.ptv-loading { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; }\n.ptv-statusbar { display: flex; align-items: center; justify-content: space-between; padding: 3px 10px; font-size: 10px; color: var(--text-faint); background: var(--background-secondary); border-top: 1px solid var(--background-modifier-border); flex-shrink: 0; margin-top: 4px; }\n\n/* 幻灯片卡片样式 */\n.ptv-slide-card { margin-bottom:20px; padding:16px 20px; border-left:3px solid var(--v6-primary); background:var(--background-secondary); border-radius:0 8px 8px 0; page-break-inside:avoid; }\n.ptv-slide-header { display:flex; align-items:center; gap:10px; margin-bottom:12px; padding-bottom:8px; border-bottom:1px solid var(--background-modifier-border); }\n.ptv-slide-num { display:inline-flex; align-items:center; justify-content:center; min-width:28px; height:28px; background:var(--v6-primary); color:#fff; border-radius:50%; font-size:12px; font-weight:600; flex-shrink:0; }\n.ptv-slide-label { font-size:12px; color:var(--text-muted); font-weight:500; }\n.ptv-para { margin:6px 0; padding:4px 0; font-size:13.5px; line-height:1.8; color:var(--text-normal); text-indent:2em; user-select:text;-webkit-user-select:text;}\n.ptv-para:first-of-type { text-indent:0; }\n.ptv-empty-slide { font-style:italic; color:var(--text-faint); font-size:12px; padding:8px 0; }\n`;\n\n// ============ JSZip 异步加载（PPTX 解压用，Dashboard 模块独立需要）============\nvar _ptvJSZipLoaded = false;\nvar _ptvJSZipLib = null;\nvar _ptvJSZipLoading = false;\nvar _ptvJSZipWaiters = [];\n\nfunction ptvGetJSZip() {\n    if (_ptvJSZipLoaded) return Promise.resolve(_ptvJSZipLib);\n    return new Promise(function(resolve) {\n        _ptvJSZipWaiters.push(resolve);\n        if (!_ptvJSZipLoading) ptvLoadJSZip();\n    });\n}\n\nfunction ptvLoadJSZip() {\n    _ptvJSZipLoading = true;\n    requestUrl({ url: 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js' })\n        .then(function(resp) {\n            try {\n                var wrapped = '(function(){var module=undefined,exports=undefined,define=undefined;' + resp.text + '\\nreturn JSZip;})()';\n                _ptvJSZipLib = eval(wrapped);\n            } catch(e) { _ptvJSZipLib = null; }\n            _ptvJSZipLoaded = true;\n            _ptvJSZipLoading = false;\n            _ptvJSZipWaiters.forEach(function(w) { w(_ptvJSZipLib); });\n            _ptvJSZipWaiters = [];\n        })\n        .catch(function() {\n            _ptvJSZipLib = null;\n            _ptvJSZipLoaded = true;\n            _ptvJSZipLoading = false;\n            _ptvJSZipWaiters.forEach(function(w) { w(null); });\n            _ptvJSZipWaiters = [];\n        });\n}\n\n// ============ PPTX 文本提取（Dashboard 模块独立使用）============\nasync function ptvExtractPptxText(arrayBuffer) {\n    var JSZip = await ptvGetJSZip();\n    if (!JSZip) throw new Error(t('mod.pptViewer.error.jszip'));\n\n    var zip = await JSZip.loadAsync(arrayBuffer);\n\n    var slideFiles = [];\n    zip.forEach(function(relativePath) {\n        var match = relativePath.match(/^ppt\\/slides\\/slide(\\d+)\\.xml$/i);\n        if (match) slideFiles.push({ path: relativePath, index: parseInt(match[1], 10) });\n    });\n    slideFiles.sort(function(a, b) { return a.index - b.index; });\n\n    if (slideFiles.length === 0) throw new Error('PPTX 内未找到任何幻灯片');\n\n    var slides = [];\n    for (var i = 0; i < slideFiles.length; i++) {\n        var sf = zip.file(slideFiles[i].path);\n        if (!sf) continue;\n        var xml = await sf.async('string');\n        // 复用 fv-legacy-office 的 XML 文本提取函数\n        var extractFn = (window.__pptExtractor && window.__pptExtractor.extractPptxXml) || function(x){return [];};\n        var paragraphs = extractFn(xml);\n        slides.push({ index: slideFiles[i].index, paragraphs: paragraphs, hasContent: paragraphs.length > 0 });\n    }\n    return slides;\n}\n\n// ============ PPT 文本提取（委托给共享的 OLE2 解析器）============\nasync function ptvExtractPptText(arrayBuffer) {\n    // 委托给 fv-legacy-office.js 导出的全局函数\n    if (window.__pptExtractor && window.__pptExtractor.extractPptText) {\n        return window.__pptExtractor.extractPptText(arrayBuffer);\n    }\n    throw new Error('PPT 解析模块未加载（需要 FileViewer 的 PPT 扩展支持）');\n}\n\n// ============ 渲染函数 ============\nfunction ptvRenderSlides(container, slides, ext) {\n    container.innerHTML = '';\n    var isPptx = (ext === 'pptx');\n\n    // 尝试复用共享的渲染函数，否则使用内嵌的简化版\n    var renderPptxFn = (window.__pptExtractor && window.__pptExtractor.renderPptxHtml) || null;\n\n    if (isPptx) {\n        for (var i = 0; i < slides.length; i++) {\n            var slide = slides[i];\n            var card = document.createElement('div');\n            card.className = 'ptv-slide-card';\n\n            var header = document.createElement('div');\n            header.className = 'ptv-slide-header';\n            header.innerHTML = '<span class=\"ptv-slide-num\">' + slide.index + '</span>' +\n                '<span class=\"ptv-slide-label\">Slide ' + slide.index + (slide.hasContent ? '' : ' — \\u7A7A\\u767D\\u9875') + '</span>';\n            card.appendChild(header);\n\n            if (slide.paragraphs.length > 0) {\n                for (var j = 0; j < slide.paragraphs.length; j++) {\n                    var para = document.createElement('p');\n                    para.className = 'ptv-para';\n                    para.textContent = slide.paragraphs[j];\n                    card.appendChild(para);\n                }\n            } else {\n                var empty = document.createElement('p');\n                empty.className = 'ptv-empty-slide';\n                empty.textContent = '（\\u6B64\\u5E7B\\u706F\\u7247\\u65E0\\u53EF\\u63D0\\u53D6\\u7684\\u6587\\u5B57\\u5185\\u5BB9）';\n                card.appendChild(empty);\n            }\n\n            container.appendChild(card);\n        }\n    } else {\n        // PPT 旧版格式 — 样式与 PPTX 完全统一\n        for (var i2 = 0; i2 < slides.length; i2++) {\n            var s2 = slides[i2];\n            var c2 = document.createElement('div');\n            c2.className = 'ptv-slide-card';\n\n            var h2 = document.createElement('div');\n            h2.className = 'ptv-slide-header';\n            h2.innerHTML = '<span class=\"ptv-slide-num\">' + s2.index + '</span>' +\n                '<span class=\"ptv-slide-label\">Slide ' + s2.index + (s2.hasContent ? '' : ' — \\u7A7A\\u767D\\u9875') + '</span>';\n            c2.appendChild(h2);\n\n            if (s2.paragraphs.length > 0) {\n                for (var j2 = 0; j2 < s2.paragraphs.length; j2++) {\n                    var p2 = document.createElement('p');\n                    p2.className = 'ptv-para';\n                    p2.textContent = s2.paragraphs[j2];\n                    c2.appendChild(p2);\n                }\n            } else {\n                var empty2 = document.createElement('p');\n                empty2.className = 'ptv-empty-slide';\n                empty2.textContent = '（\\u6B64\\u5E7B\\u706F\\u7247\\u65E0\\u53EF\\u63D0\\u53D6\\u7684\\u6587\\u5B57\\u5185\\u5BB9）';\n                c2.appendChild(empty2);\n            }\n\n            container.appendChild(c2);\n        }\n    }\n}\n\n// ============ 安全点击 ============\nfunction ptvSafeClick(el, handler) {\n    el.addEventListener('mousedown', function(evt) { evt.preventDefault(); evt.stopPropagation(); }, true);\n    el.addEventListener('click', function(evt) {\n        evt.preventDefault(); evt.stopPropagation(); evt.stopImmediatePropagation();\n        handler(evt);\n    }, true);\n}\n\n// ============ 主渲染（懒加载）============\nasync function render(container) {\n    container.addClass('ptv-wrap');\n\n    var toolbar = container.createDiv({ cls: 'ptv-toolbar' });\n    var toggleBtn = toolbar.createEl('button', { text: t('mod.pptViewer.btn.files') });\n    var refreshBtn = toolbar.createEl('button', { text: t('mod.pptViewer.btn.refresh') });\n    var loadBtn = toolbar.createEl('button', { text: t('mod.pptViewer.btn.view'), attr: { style: 'background:var(--v6-primary);color:white;border-color:var(--v6-primary);' } });\n\n    var fileList = container.createDiv({ cls: 'ptv-filelist' });\n    var viewer = container.createDiv({ cls: 'ptv-viewer' });\n    viewer.innerHTML = '<div class=\"ptv-empty\">' + t('mod.pptViewer.hint') + '<br><small>' + t('mod.pptViewer.hintFormats') + '</small></div>';\n\n    var statusbar = container.createDiv({ cls: 'ptv-statusbar' });\n    var statusFile = statusbar.createSpan();\n    var statusInfo = statusbar.createSpan();\n\n    var currentFile = null;\n    var files = [];\n\n    function scanFiles() {\n        files = [];\n        var allFiles = app.vault.getFiles();\n        for (var i = 0; i < allFiles.length; i++) {\n            var f = allFiles[i];\n            var e = f.extension.toLowerCase();\n            if (e === 'pptx' || e === 'ppt') {\n                files.push({ path: f.path, name: f.name, ext: e, size: f.stat ? f.stat.size : 0 });\n            }\n        }\n        files.sort(function(a, b) { return a.name.localeCompare(b.name); });\n        renderFileList();\n    }\n\n    function renderFileList() {\n        fileList.innerHTML = '';\n        if (files.length === 0) {\n            var empty = document.createElement('div');\n            empty.className = 'ptv-file-item';\n            empty.textContent = t('mod.pptViewer.empty');\n            empty.style.cssText = 'cursor:default;color:var(--text-muted);justify-content:center;';\n            fileList.appendChild(empty);\n            return;\n        }\n\n        for (var i = 0; i < files.length; i++) {\n            (function(f) {\n                var item = document.createElement('div');\n                item.className = 'ptv-file-item';\n                if (currentFile && currentFile.path === f.path) item.classList.add('selected');\n\n                var nameSpan = document.createElement('span');\n                nameSpan.className = 'ptv-file-name';\n                nameSpan.textContent = (f.ext === 'pptx' ? '\\uD83D\\uDCCA ' : '\\uD83D\\uDCCB ') + f.name;\n                item.appendChild(nameSpan);\n\n                var typeSpan = document.createElement('span');\n                typeSpan.className = 'ptv-file-type';\n                typeSpan.textContent = f.ext.toUpperCase();\n                item.appendChild(typeSpan);\n\n                ptvSafeClick(item, function() {\n                    currentFile = f;\n                    renderFileList();\n                    statusFile.textContent = t('mod.pptViewer.selected') + f.name;\n                });\n\n                item.addEventListener('dblclick', function(evt) {\n                    evt.preventDefault(); evt.stopPropagation(); evt.stopImmediatePropagation();\n                    currentFile = f;\n                    renderFileList();\n                    loadFile(f);\n                }, true);\n\n                fileList.appendChild(item);\n            })(files[i]);\n        }\n    }\n\n    async function loadFile(file) {\n        currentFile = file;\n        renderFileList();\n        viewer.innerHTML = '<div class=\"ptv-loading\">' + t('mod.pptViewer.loading') + '</div>';\n        statusFile.textContent = file.name;\n        statusInfo.textContent = '';\n\n        try {\n            var fileObj = app.vault.getAbstractFileByPath(file.path);\n            if (!fileObj) { showError(t('mod.pptViewer.notFound') + file.path); return; }\n\n            var data = await app.vault.readBinary(fileObj);\n\n            var arrayBuffer;\n            if (data instanceof ArrayBuffer) arrayBuffer = data;\n            else if (data instanceof Uint8Array) arrayBuffer = data.buffer;\n            else if (data && data.buffer) arrayBuffer = data.buffer;\n            else arrayBuffer = new Uint8Array(data).buffer;\n\n            var slides;\n\n            if (file.ext === 'pptx') {\n                slides = await ptvExtractPptxText(arrayBuffer);\n                statusInfo.textContent = t('mod.pptViewer.pptxRendered') + ' (' + slides.length + ' slides)';\n            } else {\n                slides = await ptvExtractPptText(arrayBuffer);\n                statusInfo.textContent = t('mod.pptViewer.pptRendered') + ' (' + slides.length + ' sections)';\n            }\n\n            viewer.innerHTML = '';\n            ptvRenderSlides(viewer, slides, file.ext);\n\n        } catch (e) {\n            var errStr = String(e && (e.message || e)).substring(0, 300);\n            showError(t('mod.pptViewer.error.load') + errStr);\n            console.error('[ppt-viewer] loadFile error:', e);\n        }\n    }\n\n    function showError(msg) {\n        var safeMsg = String(msg || '').replace(/</g, '&lt;');\n        viewer.innerHTML = '<div class=\"ptv-empty\" style=\"color:var(--text-error);white-space:pre-line;\">\\u26A0 ' + safeMsg + '</div>';\n        statusInfo.textContent = '';\n    }\n\n    toggleBtn.addEventListener('click', function() { fileList.classList.toggle('hidden'); });\n    refreshBtn.addEventListener('click', function() { scanFiles(); });\n    loadBtn.addEventListener('click', function() {\n        if (currentFile) loadFile(currentFile);\n        else if (files.length > 0) { currentFile = files[0]; renderFileList(); loadFile(files[0]); }\n    });\n\n    // 懒初始化：提前加载 JSZip\n    setTimeout(function() {\n        scanFiles();\n        ptvGetJSZip().catch(function(){});\n    }, 500);\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.pptViewer.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.pptViewer.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n    containerEl.createEl('p', {\n        text: t('mod.pptViewer.settings.hint'),\n        attr: { style: 'color:#4caf50;font-size:11px;' }\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "recent": "/**\n * 最近文件模块 V15\n * 格式：V14（含 id/styles/renderSettings）\n * 功能：V11/V14 一致（最近修改文件列表，相对时间，点击打开）\n */\nconst id = 'recent';\nconst title = t('mod.recent');\nconst icon = '🕐';\n\nconst defaultSettings = {\n    maxFiles: 10\n};\n\nconst styles = `/* 最近文件模块样式已在 styles.css 中定义 */`;\n\nfunction formatTime(timestamp) {\n    const diff = Date.now() - timestamp;\n    const minutes = Math.floor(diff / 60000);\n    const hours = Math.floor(diff / 3600000);\n    const days = Math.floor(diff / 86400000);\n    if (minutes < 1) return t('mod.recent.justNow');\n    if (minutes < 60) return minutes + t('mod.recent.minutesAgo');\n    if (hours < 24) return hours + t('mod.recent.hoursAgo');\n    if (days === 1) return t('mod.recent.yesterday');\n    if (days < 7) return days + t('mod.recent.daysAgo');\n    return moment(timestamp).format('MM-DD');\n}\n\nasync function render(content) {\n    content.empty();\n\n    const container = content.createDiv({ cls: 'recent-container' });\n    const maxFiles = settings.maxFiles || 10;\n\n    try {\n        const files = app.vault.getMarkdownFiles()\n            .sort((a, b) => b.stat.mtime - a.stat.mtime)\n            .slice(0, maxFiles);\n\n        if (files.length === 0) {\n            container.createEl('div', { text: t('mod.recent.empty'), cls: 'recent-empty' });\n            return;\n        }\n\n        files.forEach(file => {\n            const item = container.createDiv({ cls: 'recent-item' });\n            item.createEl('div', { text: '📝', cls: 'recent-icon' });\n\n            const info = item.createEl('div', { cls: 'recent-info' });\n            info.createEl('div', { text: file.basename, cls: 'recent-title' });\n\n            const pathParts = file.path.split('/');\n            pathParts.pop();\n            const folderPath = pathParts.join('/') || t('mod.recent.rootDir');\n            info.createEl('div', { text: folderPath, cls: 'recent-path' });\n\n            item.createEl('div', { text: formatTime(file.stat.mtime), cls: 'recent-time' });\n\n            item.addEventListener('click', () => {\n                app.workspace.openLinkText(file.path, '', false);\n            });\n        });\n\n    } catch (e) {\n        container.createEl('div', {\n            text: t('mod.recent.error.loadFailed') + e.message,\n            attr: { style: 'padding: 20px; text-align: center; color: var(--text-muted); font-size: 12px;' }\n        });\n    }\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    const { Setting } = require('obsidian');\n\n    containerEl.createEl('h3', { text: t('mod.recent.settings.title') });\n\n    new Setting(containerEl)\n        .setName(t('mod.recent.settings.maxFiles'))\n        .setDesc(t('mod.recent.settings.maxFilesDesc'))\n        .addSlider(s => {\n            s.setLimits(5, 30, 5)\n                .setValue(settings.maxFiles || 10)\n                .setDynamicTooltip()\n                .onChange(async (v) => {\n                    settings.maxFiles = v;\n                    await saveCallback();\n                });\n        });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "spreadsheet": "// spreadsheet 模块 - 表格文件查看器（xlsx/xls/csv）\n// 源插件: univer（表格查看器）\n// 核心功能: 面板内SheetJS渲染xlsx/csv为HTML表格\nconst id = 'spreadsheet';\nconst title = t('mod.spreadsheet');\nconst icon = '📈';\n\nconst defaultSettings = {\n    language: 'ZH',\n    isSupportXlsx: true\n};\n\nconst styles = `\n.ss-viewer { padding: 8px 10px; display: flex; flex-direction: column; height: 100%; min-height: 0; }\n.ss-toolbar { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; flex-shrink: 0; flex-wrap: wrap; }\n.ss-toolbar button { padding: 3px 8px; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 11px; cursor: pointer; white-space: nowrap; }\n.ss-toolbar button:hover { background: var(--background-modifier-hover); }\n.ss-toolbar button.active { background: var(--v6-primary); color: white; border-color: var(--v6-primary); }\n.ss-filelist { max-height: 90px; overflow-y: auto; margin-bottom: 6px; flex-shrink: 0; border: 1px solid var(--background-modifier-border); border-radius: 4px; }\n.ss-filelist.hidden { display: none; }\n.ss-file-item { display: flex; align-items: center; justify-content: space-between; padding: 4px 8px; cursor: pointer; font-size: 11px; transition: background 0.1s; border-radius: 2px; user-select: none; -webkit-user-select: none; }\n.ss-file-item:hover { background: var(--background-modifier-hover); }\n.ss-file-item.selected { background: rgba(var(--v6-primary-rgb, 232,149,109), 0.2); color: var(--v6-primary); }\n.ss-file-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.ss-file-type { color: var(--text-faint); font-size: 10px; flex-shrink: 0; margin-left: 8px; background: var(--background-modifier-form-field); padding: 1px 6px; border-radius: 8px; }\n.ss-sheets { display: flex; gap: 2px; margin-bottom: 4px; flex-shrink: 0; flex-wrap: wrap; }\n.ss-sheet-tab { padding: 2px 10px; border: 1px solid var(--background-modifier-border); border-radius: 4px 4px 0 0; font-size: 10px; cursor: pointer; background: var(--background-secondary); color: var(--text-muted); border-bottom: none; user-select: none; }\n.ss-sheet-tab.active { background: var(--background-modifier-form-field); color: var(--text-normal); font-weight: 600; }\n.ss-sheet-tab:hover { color: var(--text-normal); }\n.ss-table-wrap { flex: 1; overflow: auto; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: var(--background-primary); min-height: 60px; position: relative; }\n.ss-table-wrap table { border-collapse: collapse; font-size: 11px; min-width: 100%; }\n.ss-table-wrap th, .ss-table-wrap td { border: 1px solid var(--background-modifier-border); padding: 4px 8px; white-space: nowrap; min-width: 40px; max-width: 300px; overflow: hidden; text-overflow: ellipsis; }\n.ss-table-wrap th { background: var(--background-modifier-form-field); font-weight: 600; position: sticky; top: 0; z-index: 1; }\n.ss-table-wrap tr:hover td { background: var(--background-modifier-hover); }\n.ss-table-wrap tr:nth-child(even) td { background: rgba(128,128,128,0.05); }\n.ss-empty { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; text-align: center; pointer-events: none; }\n.ss-loading { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; }\n.ss-statusbar { display: flex; align-items: center; justify-content: space-between; padding: 2px 6px; font-size: 10px; color: var(--text-faint); background: var(--background-secondary); border-top: 1px solid var(--background-modifier-border); flex-shrink: 0; margin-top: 3px; }\n.ss-row-header { background: var(--background-modifier-form-field) !important; font-weight: 600; text-align: center !important; color: var(--text-muted) !important; font-size: 10px !important; }\n`;\n\n// ============ SheetJS 异步加载 ============\nvar _xlsxLoaded = false;\nvar _xlsxLib = null;\nvar _xlsxLoading = false;\nvar _xlsxWaiters = [];\n\nfunction getXLSX() {\n    if (_xlsxLoaded) return Promise.resolve(_xlsxLib);\n    return new Promise(function(resolve) {\n        _xlsxWaiters.push(resolve);\n        if (!_xlsxLoading) loadXLSX();\n    });\n}\n\nfunction loadXLSX() {\n    _xlsxLoading = true;\n    try {\n        requestUrl({ url: 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js' })\n            .then(function(resp) {\n                try {\n                    var code = resp.text;\n                    // eval+IIFE 隔离执行（屏蔽 module/exports/define），直接返回 XLSX 对象\n                    var wrapped = '(function(){var module=undefined,exports=undefined,define=undefined;' + code + '\\nreturn XLSX;})()';\n                    _xlsxLib = eval(wrapped);\n                    if (typeof _xlsxLib !== 'object' || typeof _xlsxLib.read !== 'function') {\n                        console.warn('[spreadsheet] eval返回无效对象，尝试 window.XLSX');\n                        _xlsxLib = window.XLSX || null;\n                    }\n                    _xlsxLoaded = true;\n                    _xlsxLoading = false;\n                    console.log('[spreadsheet] SheetJS 加载完成, read:', typeof (_xlsxLib && _xlsxLib.read));\n                    _xlsxWaiters.forEach(function(w) { w(_xlsxLib); });\n                    _xlsxWaiters = [];\n                } catch(e) {\n                    console.error('SheetJS eval 失败:', e);\n                    _xlsxLoading = false;\n                    _xlsxLib = null;\n                    _xlsxLoaded = true;\n                    _xlsxWaiters.forEach(function(w) { w(null); });\n                    _xlsxWaiters = [];\n                }\n            })\n            .catch(function() {\n                _xlsxLoading = false;\n                _xlsxLib = null;\n                _xlsxLoaded = true;\n                _xlsxWaiters.forEach(function(w) { w(null); });\n                _xlsxWaiters = [];\n            });\n    } catch(e) {\n        _xlsxLoading = false;\n        _xlsxLib = null;\n        _xlsxLoaded = true;\n        _xlsxWaiters.forEach(function(w) { w(null); });\n        _xlsxWaiters = [];\n    }\n}\n\n// ============ CSV解析 ============\nfunction parseCSV(text) {\n    var rows = [];\n    var current = [];\n    var cell = '';\n    var inQuotes = false;\n    for (var i = 0; i < text.length; i++) {\n        var ch = text[i];\n        if (inQuotes) {\n            if (ch === '\"') {\n                if (i + 1 < text.length && text[i + 1] === '\"') { cell += '\"'; i++; }\n                else { inQuotes = false; }\n            } else { cell += ch; }\n        } else {\n            if (ch === '\"') { inQuotes = true; }\n            else if (ch === ',') { current.push(cell); cell = ''; }\n            else if (ch === '\\n' || ch === '\\r') {\n                if (cell || current.length > 0) { current.push(cell); cell = ''; rows.push(current); current = []; }\n                if (ch === '\\r' && i + 1 < text.length && text[i + 1] === '\\n') i++;\n            } else { cell += ch; }\n        }\n    }\n    if (cell) current.push(cell);\n    if (current.length > 0) rows.push(current);\n    return rows;\n}\n\n// ============ 渲染表格 ============\nfunction renderTable(container, rows) {\n    container.innerHTML = '';\n    if (!rows || rows.length === 0) {\n        container.innerHTML = '<div class=\"ss-empty\">' + t('mod.spreadsheet.emptyTable') + '</div>';\n        return { rows: 0, cols: 0 };\n    }\n    var maxCols = 0;\n    for (var i = 0; i < rows.length; i++) {\n        if (rows[i].length > maxCols) maxCols = rows[i].length;\n    }\n    if (maxCols === 0) {\n        container.innerHTML = '<div class=\"ss-empty\">' + t('mod.spreadsheet.emptyTable') + '</div>';\n        return { rows: 0, cols: 0 };\n    }\n    var table = document.createElement('table');\n    var thead = document.createElement('thead');\n    var trH = document.createElement('tr');\n    var cornerTh = document.createElement('th');\n    cornerTh.textContent = '#';\n    cornerTh.style.cssText = 'width:35px;text-align:center;';\n    trH.appendChild(cornerTh);\n\n    var headerRow = rows[0] || [];\n    for (var ci = 0; ci < maxCols; ci++) {\n        var th = document.createElement('th');\n        th.textContent = headerRow[ci] !== undefined ? String(headerRow[ci]) : '';\n        trH.appendChild(th);\n    }\n    thead.appendChild(trH);\n    table.appendChild(thead);\n\n    var tbody = document.createElement('tbody');\n    for (var ri = 1; ri < rows.length; ri++) {\n        var tr = document.createElement('tr');\n        var rowNumTd = document.createElement('td');\n        rowNumTd.textContent = ri;\n        rowNumTd.className = 'ss-row-header';\n        tr.appendChild(rowNumTd);\n        var rowData = rows[ri] || [];\n        for (var cj = 0; cj < maxCols; cj++) {\n            var td = document.createElement('td');\n            td.textContent = rowData[cj] !== undefined ? String(rowData[cj]) : '';\n            tr.appendChild(td);\n        }\n        tbody.appendChild(tr);\n    }\n    table.appendChild(tbody);\n    container.appendChild(table);\n    return { rows: rows.length, cols: maxCols };\n}\n\n// ============ 安全点击：阻止事件冒泡到Obsidian ============\nfunction safeClick(el, handler) {\n    // 多重防护：mousedown + click + capture\n    el.addEventListener('mousedown', function(evt) {\n        evt.preventDefault();\n        evt.stopPropagation();\n    }, true);\n    el.addEventListener('click', function(evt) {\n        evt.preventDefault();\n        evt.stopPropagation();\n        evt.stopImmediatePropagation();\n        handler(evt);\n    }, true);\n}\n\n// ============ 主渲染（懒加载：不自动打开文件）============\nasync function render(container) {\n    container.addClass('ss-viewer');\n    var s = settings;\n\n    var toolbar = container.createDiv({ cls: 'ss-toolbar' });\n    var toggleBtn = toolbar.createEl('button', { text: t('mod.spreadsheet.btn.files') });\n    var refreshBtn = toolbar.createEl('button', { text: t('mod.spreadsheet.btn.refresh') });\n    var loadBtn = toolbar.createEl('button', { text: t('mod.spreadsheet.btn.view'), attr: { style: 'background:var(--v6-primary);color:white;border-color:var(--v6-primary);' } });\n\n    var fileList = container.createDiv({ cls: 'ss-filelist' });\n    var sheetsBar = container.createDiv({ cls: 'ss-sheets' });\n    var tableWrap = container.createDiv({ cls: 'ss-table-wrap' });\n    tableWrap.innerHTML = '<div class=\"ss-empty\">' + t('mod.spreadsheet.hint') + '<br><small>' + t('mod.spreadsheet.hintFormats') + '</small></div>';\n    var statusbar = container.createDiv({ cls: 'ss-statusbar' });\n    var statusInfo = statusbar.createSpan();\n    var statusStats = statusbar.createSpan();\n\n    var currentFile = null;\n    var workbookData = null;\n    var activeSheet = '';\n    var files = [];\n    var _scanned = false;\n\n    function scanFiles() {\n        files = [];\n        var allFiles = app.vault.getFiles();\n        for (var i = 0; i < allFiles.length; i++) {\n            var f = allFiles[i];\n            var ext = f.extension.toLowerCase();\n            if (ext === 'xlsx' || ext === 'xls' || ext === 'csv' || ext === 'ods') {\n                files.push({ path: f.path, name: f.name, ext: ext });\n            }\n        }\n        files.sort(function(a, b) { return a.name.localeCompare(b.name); });\n        renderFileList();\n    }\n\n    function renderFileList() {\n        fileList.innerHTML = '';\n        if (files.length === 0) {\n            var emptyItem = document.createElement('div');\n            emptyItem.className = 'ss-file-item';\n            emptyItem.textContent = t('mod.spreadsheet.empty');\n            emptyItem.style.cssText = 'cursor:default;color:var(--text-muted);justify-content:center;';\n            fileList.appendChild(emptyItem);\n            return;\n        }\n        for (var i = 0; i < files.length; i++) {\n            (function(f) {\n                var item = document.createElement('div');\n                item.className = 'ss-file-item';\n                if (currentFile && currentFile.path === f.path) item.classList.add('selected');\n\n                var iconMap = { csv: '📄', ods: '📗', xlsx: '📊', xls: '📊' };\n                var nameSpan = document.createElement('span');\n                nameSpan.className = 'ss-file-name';\n                nameSpan.textContent = (iconMap[f.ext] || '📊') + ' ' + f.name;\n                item.appendChild(nameSpan);\n\n                var typeSpan = document.createElement('span');\n                typeSpan.className = 'ss-file-type';\n                typeSpan.textContent = f.ext.toUpperCase();\n                item.appendChild(typeSpan);\n\n                // 安全点击：多重防护\n                safeClick(item, function() {\n                    // 选中文件（不自动加载）\n                    currentFile = f;\n                    renderFileList();\n                    statusInfo.textContent = t('mod.spreadsheet.selected') + f.name;\n                });\n\n                // 双击加载\n                item.addEventListener('dblclick', function(evt) {\n                    evt.preventDefault();\n                    evt.stopPropagation();\n                    evt.stopImmediatePropagation();\n                    currentFile = f;\n                    renderFileList();\n                    loadFile(f);\n                }, true);\n\n                fileList.appendChild(item);\n            })(files[i]);\n        }\n    }\n\n    async function loadFile(file) {\n        currentFile = file;\n        renderFileList();\n        tableWrap.innerHTML = '<div class=\"ss-loading\">' + t('mod.spreadsheet.loading') + '</div>';\n        sheetsBar.innerHTML = '';\n        statusInfo.textContent = file.name;\n        statusStats.textContent = '';\n\n        try {\n            var fileObj = app.vault.getAbstractFileByPath(file.path);\n            if (!fileObj) { showError(t('mod.spreadsheet.notFound') + file.path); return; }\n\n            var data = await app.vault.readBinary(fileObj);\n\n            if (file.ext === 'csv') {\n                var text = new TextDecoder('utf-8').decode(data);\n                var rows = parseCSV(text);\n                workbookData = { sheets: { 'Sheet1': rows }, sheetNames: ['Sheet1'] };\n                showSheet('Sheet1');\n                statusStats.textContent = 'CSV | ' + rows.length + ' 行';\n            } else {\n                var XLSX = await getXLSX();\n                if (!XLSX) {\n                    showError(t('mod.spreadsheet.error.xlsx') + '\\n\\n' + t('mod.spreadsheet.error.csvOk'));\n                    return;\n                }\n                // 确保 XLSX.read 可用（可能是挂载到window的）\n                if (typeof XLSX.read !== 'function' && window.XLSX && typeof window.XLSX.read === 'function') {\n                    XLSX = window.XLSX;\n                }\n\n                var wb = XLSX.read(new Uint8Array(data), { type: 'array' });\n                workbookData = { sheets: {}, sheetNames: wb.SheetNames };\n\n                wb.SheetNames.forEach(function(name) {\n                    var ws = wb.Sheets[name];\n                    workbookData.sheets[name] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });\n                });\n\n                renderSheetTabs();\n                if (workbookData.sheetNames.length > 0) {\n                    showSheet(workbookData.sheetNames[0]);\n                }\n                statusStats.textContent = 'XLSX | ' + workbookData.sheetNames.length + ' sheet(s)';\n            }\n        } catch (e) {\n            showError(t('mod.docViewer.error.load') + (e.message || e));\n            console.error('spreadsheet loadFile error:', e);\n        }\n    }\n\n    function renderSheetTabs() {\n        sheetsBar.innerHTML = '';\n        if (!workbookData || workbookData.sheetNames.length <= 1) return;\n\n        workbookData.sheetNames.forEach(function(name) {\n            var tab = document.createElement('div');\n            tab.className = 'ss-sheet-tab';\n            if (name === activeSheet) tab.classList.add('active');\n            tab.textContent = name;\n\n            safeClick(tab, function() {\n                showSheet(name);\n            });\n            sheetsBar.appendChild(tab);\n        });\n    }\n\n    function showSheet(name) {\n        activeSheet = name;\n        renderSheetTabs();\n        var rows = workbookData.sheets[name] || [];\n        var stats = renderTable(tableWrap, rows);\n        statusStats.textContent = (stats.rows || 0) + ' 行 × ' + (stats.cols || 0) + ' 列';\n    }\n\n    function showError(msg) {\n        tableWrap.innerHTML = '<div class=\"ss-empty\" style=\"color:var(--text-error);white-space:pre-line;\">⚠ ' + msg.replace(/</g, '&lt;') + '</div>';\n        sheetsBar.innerHTML = '';\n        statusStats.textContent = '';\n    }\n\n    // 按钮事件\n    toggleBtn.addEventListener('click', function() { fileList.classList.toggle('hidden'); });\n    refreshBtn.addEventListener('click', function() { scanFiles(); });\n    loadBtn.addEventListener('click', function() {\n        if (currentFile) loadFile(currentFile);\n        else if (files.length > 0) { currentFile = files[0]; renderFileList(); loadFile(files[0]); }\n    });\n\n    // 懒初始化：延迟扫描 + 预加载SheetJS\n    setTimeout(function() {\n        scanFiles();\n        getXLSX().catch(function(){});\n    }, 300);\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.spreadsheet.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.spreadsheet.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n    containerEl.createEl('p', {\n        text: t('mod.spreadsheet.settings.hint1'),\n        attr: { style: 'color:#4caf50;font-size:11px;' }\n    });\n    containerEl.createEl('p', {\n        text: t('mod.spreadsheet.settings.hint2'),\n        attr: { style: 'color:var(--text-muted);font-size:11px;' }\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "stats": "/**\n * 统计模块 V15\n * 格式：V14（含 id/styles/renderSettings）\n * 功能：V11 完整版（笔记数/总字数/文件夹数/平均字数 + 文件夹排行Top5 带进度条）\n */\nconst id = 'stats';\nconst title = t('mod.stats');\nconst icon = '📈';\n\nconst defaultSettings = {\n    showFileCount: true,\n    showWordCount: true\n};\n\nconst styles = `/* 统计模块样式已在 styles.css 中定义 */`;\n\nfunction formatNumber(num) {\n    if (num >= 10000) return (num / 10000).toFixed(1) + '万';\n    return num.toLocaleString();\n}\n\nasync function render(content) {\n    content.empty();\n\n    const container = content.createDiv({ cls: 'stats-container' });\n\n    // 加载提示\n    const loading = container.createEl('div', {\n        text: t('mod.stats.loading'),\n        attr: { style: 'grid-column: 1/-1; text-align: center; padding: 20px; color: var(--text-muted); font-size: 13px;' }\n    });\n\n    try {\n        const files = app.vault.getMarkdownFiles();\n\n        let totalWords = 0;\n        const folderCount = new Set();\n        const folderFiles = {};\n\n        for (const file of files) {\n            try {\n                const fileContent = await app.vault.read(file);\n                // 移除 YAML frontmatter 和 Markdown 符号再统计字符数\n                const clean = fileContent\n                    .replace(/^---[\\s\\S]*?---\\n?/, '')\n                    .replace(/```[\\s\\S]*?```/g, '')\n                    .replace(/`[^`]*`/g, '')\n                    .replace(/[#*\\[\\]>!\\-_~|]/g, '');\n                totalWords += clean.replace(/\\s+/g, '').length;\n            } catch (e) { /* 忽略单文件读取失败 */ }\n\n            const parts = file.path.split('/');\n            if (parts.length > 1) {\n                folderCount.add(parts[0]);\n                folderFiles[parts[0]] = (folderFiles[parts[0]] || 0) + 1;\n            }\n        }\n\n        const avgWords = files.length > 0 ? Math.round(totalWords / files.length) : 0;\n        const topFolders = Object.entries(folderFiles)\n            .sort((a, b) => b[1] - a[1])\n            .slice(0, 5);\n        const maxCount = topFolders.length > 0 ? topFolders[0][1] : 1;\n\n        // 清空加载提示\n        container.empty();\n\n        // 四个统计卡片\n        const showFileCount = settings.showFileCount !== false;\n        const showWordCount = settings.showWordCount !== false;\n\n        const items = [];\n        if (showFileCount) {\n            items.push({ icon: '', value: files.length, label: t('mod.stats.totalNotes') });\n        }\n        if (showWordCount) {\n            items.push({ icon: '', value: totalWords, label: t('mod.stats.totalWords') });\n        }\n        items.push({ icon: '', value: folderCount.size, label: t('mod.stats.folders') });\n        if (showWordCount) {\n            items.push({ icon: '', value: avgWords, label: t('mod.stats.avgWords') });\n        }\n\n        items.forEach(item => {\n            const itemEl = container.createDiv({ cls: 'stats-item' });\n            itemEl.createEl('div', { text: item.icon, cls: 'stats-icon' });\n            itemEl.createEl('div', { text: formatNumber(item.value), cls: 'stats-value' });\n            itemEl.createEl('div', { text: item.label, cls: 'stats-label' });\n        });\n\n        // 文件夹排行（带进度条）\n        if (topFolders.length > 0) {\n            const rankDiv = container.createDiv({ cls: 'stats-rank' });\n            rankDiv.createEl('div', { text: t('mod.stats.folderRank'), cls: 'stats-rank-title' });\n\n            topFolders.forEach((folder, index) => {\n                const rankItem = rankDiv.createDiv({ cls: 'stats-rank-item' });\n                rankItem.createEl('span', {\n                    text: ['🥇','🥈','🥉','4️⃣','5️⃣'][index] || String(index + 1)\n                });\n\n                const info = rankItem.createDiv({ cls: 'stats-rank-info' });\n                info.createEl('div', { text: folder[0], cls: 'stats-rank-name' });\n\n                const barWrap = info.createDiv({ cls: 'stats-rank-bar-wrap' });\n                const bar = barWrap.createDiv({ cls: 'stats-rank-bar' });\n                const pct = Math.round((folder[1] / maxCount) * 100);\n                bar.style.width = pct + '%';\n\n                rankItem.createEl('span', { text: folder[1] + ' 篇', cls: 'stats-rank-count' });\n            });\n        }\n\n    } catch (e) {\n        container.empty();\n        container.createEl('div', {\n            text: t('mod.stats.error.loadFailed') + e.message,\n            attr: { style: 'grid-column: 1/-1; text-align: center; padding: 20px; color: var(--text-muted); font-size: 12px;' }\n        });\n    }\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    const { Setting } = require('obsidian');\n\n    containerEl.createEl('h3', { text: t('mod.stats.settings.title') });\n\n    new Setting(containerEl)\n        .setName(t('mod.stats.settings.showCount'))\n        .addToggle(t => {\n            t.setValue(settings.showFileCount !== false)\n                .onChange(async (v) => {\n                    settings.showFileCount = v;\n                    await saveCallback();\n                });\n        });\n\n    new Setting(containerEl)\n        .setName(t('mod.stats.settings.showWords'))\n        .addToggle(t => {\n            t.setValue(settings.showWordCount !== false)\n                .onChange(async (v) => {\n                    settings.showWordCount = v;\n                    await saveCallback();\n                });\n        });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "table-resize": "// table-resize 模块 - Markdown查看模式下表格列宽拖拽调整\n// 源插件: obsidian-table-column-resize\n// 核心功能保留: 向页面中表格注入拖拽手柄，支持列宽调整\nconst id = 'table-resize';\nconst title = t('mod.tableResize');\nconst icon = '📐';\n\nconst defaultSettings = {\n    minColumnWidth: 50\n};\n\nconst styles = `\n.trs-wrap { padding: 10px 12px; display: flex; flex-direction: column; height: 100%; }\n.trs-title { font-size: 12px; font-weight: 600; color: var(--v6-primary); margin-bottom: 8px; }\n.trs-row { display: flex; align-items: center; justify-content: space-between; padding: 6px 0; font-size: 12px; }\n.trs-row label { color: var(--text-normal); }\n.trs-row input { width: 70px; padding: 4px 8px; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: var(--background-modifier-form-field); color: var(--text-normal); font-size: 12px; text-align: center; outline: none; }\n.trs-row input:focus { border-color: var(--v6-primary); }\n.trs-btn { padding: 6px 14px; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; margin-top: 8px; transition: opacity 0.15s; }\n.trs-btn.primary { background: var(--v6-primary); color: white; }\n.trs-btn.primary:hover { opacity: 0.85; }\n.trs-btn.secondary { background: var(--background-modifier-form-field); color: var(--text-normal); border: 1px solid var(--background-modifier-border); }\n.trs-status { margin-top: 10px; padding: 8px; background: var(--background-modifier-form-field); border-radius: 6px; font-size: 11px; color: var(--text-muted); }\n.trs-status strong { color: var(--text-normal); }\n/* 拖拽手柄样式 */\n.trs-resize-handle { position: absolute; top: 0; right: -3px; width: 6px; height: 100%; cursor: col-resize; z-index: 10; background: transparent; transition: background 0.15s; }\n.trs-resize-handle:hover, .trs-resize-handle.active { background: var(--v6-primary); opacity: 0.5; }\n.trs-resizing { user-select: none !important; cursor: col-resize !important; }\n`;\n\n// ============ 拖拽引擎 ============\nvar resizeActive = false;\nvar resizeHandles = [];\n\nfunction injectHandles(minWidth) {\n    removeHandles();\n\n    // 只在阅读模式的markdown渲染结果中注入\n    document.querySelectorAll('.markdown-reading-view table, .markdown-preview-view table, .markdown-rendered table').forEach(function(table) {\n        if (table.dataset.trsInjected) return;\n        table.dataset.trsInjected = '1';\n        table.style.position = 'relative';\n\n        var rows = table.querySelectorAll('tr');\n        if (rows.length === 0) return;\n\n        // 取第一行所有单元格\n        var firstRow = rows[0];\n        var cells = firstRow.querySelectorAll('th, td');\n        if (cells.length === 0) return;\n\n        cells.forEach(function(cell, idx) {\n            cell.style.position = 'relative';\n            var handle = document.createElement('div');\n            handle.className = 'trs-resize-handle';\n            handle.title = t('mod.tableResize.handleTitle');\n            cell.appendChild(handle);\n\n            var startX, startWidth;\n            handle.addEventListener('mousedown', function(e) {\n                e.preventDefault();\n                e.stopPropagation();\n                resizeActive = true;\n                handle.classList.add('active');\n                document.body.classList.add('trs-resizing');\n\n                startX = e.clientX;\n                startWidth = cell.getBoundingClientRect().width;\n\n                function onMove(ev) {\n                    if (!resizeActive) return;\n                    var diff = ev.clientX - startX;\n                    var newWidth = Math.max(minWidth || 50, startWidth + diff);\n                    cell.style.width = newWidth + 'px';\n                    cell.style.minWidth = newWidth + 'px';\n                }\n\n                function onUp() {\n                    resizeActive = false;\n                    handle.classList.remove('active');\n                    document.body.classList.remove('trs-resizing');\n                    document.removeEventListener('mousemove', onMove);\n                    document.removeEventListener('mouseup', onUp);\n                }\n\n                document.addEventListener('mousemove', onMove);\n                document.addEventListener('mouseup', onUp);\n            });\n\n            resizeHandles.push(handle);\n        });\n    });\n}\n\nfunction removeHandles() {\n    resizeHandles.forEach(function(h) { if (h.parentNode) h.parentNode.removeChild(h); });\n    resizeHandles = [];\n    document.querySelectorAll('[data-trs-injected]').forEach(function(el) { delete el.dataset.trsInjected; });\n}\n\nasync function render(container) {\n    container.addClass('trs-wrap');\n    var s = settings;\n\n    container.createDiv({ text: t('mod.tableResize.title'), cls: 'trs-title' });\n\n    // 最小列宽设置\n    var row = container.createDiv({ cls: 'trs-row' });\n    row.createEl('label', { text: t('mod.tableResize.minWidth') });\n    var widthInput = row.createEl('input', { attr: { type: 'number', value: s.minColumnWidth || 50, min: 20, max: 500 } });\n    widthInput.addEventListener('change', function() {\n        s.minColumnWidth = Math.max(20, parseInt(widthInput.value) || 50);\n        if (typeof saveCallback === 'function') saveCallback();\n        if (resizeActive) { removeHandles(); injectHandles(s.minColumnWidth); }\n    });\n\n    // 操作按钮\n    var injectBtn = container.createEl('button', { text: t('mod.tableResize.btn.inject'), cls: 'trs-btn primary', attr: { style: 'margin-right:6px;' } });\n    var removeBtn = container.createEl('button', { text: t('mod.tableResize.btn.remove'), cls: 'trs-btn secondary' });\n\n    // 状态显示\n    var status = container.createDiv({ cls: 'trs-status' });\n    updateStatus();\n\n    function updateStatus() {\n        var tableCount = document.querySelectorAll('.markdown-reading-view table, .markdown-preview-view table, .markdown-rendered table').length;\n        var injectedCount = document.querySelectorAll('[data-trs-injected]').length;\n        status.innerHTML = t('mod.tableResize.status') + '<strong>' + tableCount + '</strong>' + t('mod.tableResize.statusInjected') + '<strong>' + injectedCount + '</strong>' + t('mod.tableResize.statusHandles') + '<strong>' + resizeHandles.length + '</strong>';\n    }\n\n    injectBtn.addEventListener('click', function() {\n        injectHandles(s.minColumnWidth || 50);\n        updateStatus();\n        new Notice(t('mod.tableResize.injected') + document.querySelectorAll('[data-trs-injected]').length + t('mod.tableResize.injectedSuffix'));\n    });\n\n    removeBtn.addEventListener('click', function() {\n        removeHandles();\n        updateStatus();\n        new Notice(t('mod.tableResize.removed'));\n    });\n\n    // 定时刷新状态\n    var statusInterval = (function(){ var id = setInterval(updateStatus, 3600000)/*TEMP_DISABLED*/; _cleanupFns.push(function(){ clearInterval(id); }); return id; })();\n\n    // 模块销毁时清理（由框架调用？这里用简单方案：组件卸载时清理）\n    // 注: dashboard框架的render会在每次切换时重新创建container，旧DOM会被销毁\n    // 我们用一个MutationObserver确保新出现的表格也被注入\n    var mutationObserver = new MutationObserver(function() {\n        if (resizeHandles.length > 0) {\n            injectHandles(s.minColumnWidth || 50);\n            updateStatus();\n        }\n    });\n\n    // ★ 修复：只监听仪表盘容器内的变化，不再监听 document.body\n    // 原代码监听 document.body + subtree:true 会导致任何DOM变化触发回调\n    // 回调操作DOM → 又触发observer → 无限循环 → CPU占满\n    mutationObserver.observe(container, { childList: true, subtree: true });\n\n    // 自动注入\n    setTimeout(function() { injectHandles(s.minColumnWidth || 50); updateStatus(); }, 500);\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.tableResize.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.tableResize.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n    containerEl.createEl('p', {\n        text: t('mod.tableResize.settings.hint'),\n        attr: { style: 'color:var(--text-muted);font-size:11px;' }\n    });\n}\n\n\n// === 自动生成的 onunload 清理函数 ===\nvar _cleanupFns = [];\nmodule.exports.onunload = function() {\n    _cleanupFns.forEach(function(fn){ try{fn();}catch(e){} });\n    _cleanupFns = [];\n};\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "todo": "/**\n * 待办模块 V15\n * 格式：V14（含 id/styles/renderSettings）\n * 功能：V11 完整版（增删改查 + 双击编辑 + 筛选 + 进度统计 + 读写 Markdown 文件）\n */\nconst id = 'todo';\nconst title = t('mod.todo');\nconst icon = '✅';\n\nconst defaultSettings = {\n    folder: '待办'\n};\n\nconst styles = `/* 待办模块样式已在 styles.css 中定义 */`;\n\n// 全局筛选状态\nif (!window._v15TodoState) {\n    window._v15TodoState = { filter: 'all' };\n}\n\nfunction parseTodos(content) {\n    const todos = [];\n    content.split('\\n').forEach(line => {\n        const matchActive = line.match(/^\\s*- \\[ \\] (.*)$/);\n        const matchDone = line.match(/^\\s*- \\[x\\] (.*)$/i);\n        if (matchActive) todos.push({ text: matchActive[1].trim(), completed: false, rawLine: line });\n        else if (matchDone) todos.push({ text: matchDone[1].trim(), completed: true, rawLine: line });\n    });\n    return todos;\n}\n\nasync function ensureTodoFile(folder, filename) {\n    const today = moment().format('YYYY-MM-DD');\n    let file = app.vault.getAbstractFileByPath(filename);\n    if (!file) {\n        const folderExists = app.vault.getAbstractFileByPath(folder);\n        if (!folderExists) {\n            await app.vault.createFolder(folder);\n        }\n        await app.vault.create(filename, `# ${today} 待办事项\\n\\n`);\n        file = app.vault.getAbstractFileByPath(filename);\n    }\n    return file;\n}\n\nasync function addTodo(filename, text) {\n    const file = app.vault.getAbstractFileByPath(filename);\n    const c = await app.vault.read(file);\n    await app.vault.modify(file, c + `- [ ] ${text}\\n`);\n}\n\nasync function toggleTodo(filename, todo) {\n    const file = app.vault.getAbstractFileByPath(filename);\n    const c = await app.vault.read(file);\n    const lines = c.split('\\n');\n    const idx = lines.findIndex(l => l === todo.rawLine);\n    if (idx >= 0) {\n        lines[idx] = todo.completed\n            ? lines[idx].replace(/- \\[x\\]/i, '- [ ]')\n            : lines[idx].replace('- [ ]', '- [x]');\n        await app.vault.modify(file, lines.join('\\n'));\n    }\n}\n\nasync function deleteTodo(filename, todo) {\n    const file = app.vault.getAbstractFileByPath(filename);\n    const c = await app.vault.read(file);\n    const lines = c.split('\\n');\n    const idx = lines.findIndex(l => l === todo.rawLine);\n    if (idx >= 0) {\n        lines.splice(idx, 1);\n        await app.vault.modify(file, lines.join('\\n'));\n    }\n}\n\nasync function editTodo(filename, todo, newText) {\n    const file = app.vault.getAbstractFileByPath(filename);\n    const c = await app.vault.read(file);\n    const lines = c.split('\\n');\n    const idx = lines.findIndex(l => l === todo.rawLine);\n    if (idx >= 0) {\n        const prefix = todo.completed ? '- [x] ' : '- [ ] ';\n        lines[idx] = prefix + newText;\n        await app.vault.modify(file, lines.join('\\n'));\n    }\n}\n\nasync function render(content) {\n    const state = window._v15TodoState;\n    content.empty();\n\n    const folder = settings.folder || '待办';\n    const today = moment().format('YYYY-MM-DD');\n    const filename = `${folder}/${today}.md`;\n\n    const container = content.createDiv({ cls: 'todo-container' });\n\n    // 输入区域\n    const inputArea = container.createDiv({ cls: 'todo-input-area' });\n    const inputWrapper = inputArea.createDiv({ cls: 'todo-input-wrapper' });\n    inputWrapper.createDiv({ cls: 'todo-input-icon', text: '⭕' });\n    const input = inputWrapper.createEl('input', {\n        cls: 'todo-input',\n        attr: { placeholder: t('mod.todo.placeholder') }\n    });\n    input.addEventListener('keypress', async (e) => {\n        if (e.key === 'Enter' && input.value.trim()) {\n            await ensureTodoFile(folder, filename);\n            await addTodo(filename, input.value.trim());\n            input.value = '';\n            render(content);\n        }\n    });\n\n    let todos = [];\n    try {\n        await ensureTodoFile(folder, filename);\n        const file = app.vault.getAbstractFileByPath(filename);\n        const fileContent = await app.vault.read(file);\n        todos = parseTodos(fileContent);\n    } catch (e) {\n        container.createEl('div', { text: t('mod.todo.error.readFailed') + e.message, attr: { style: 'padding: 10px; color: var(--text-muted); font-size: 12px;' } });\n        return;\n    }\n\n    const completed = todos.filter(t => t.completed).length;\n    const total = todos.length;\n\n    // 筛选栏\n    const filterArea = container.createDiv({ cls: 'todo-filter-area' });\n    [\n        { key: 'all', label: `${t('mod.todo.filter.all')} ${total}` },\n        { key: 'active', label: `${t('mod.todo.filter.todo')} ${total - completed}` },\n        { key: 'done', label: `${t('mod.todo.filter.done')} ${completed}` }\n    ].forEach(f => {\n        const btn = filterArea.createEl('button', {\n            cls: 'todo-filter-btn' + (state.filter === f.key ? ' active' : ''),\n            text: f.label\n        });\n        btn.addEventListener('click', () => {\n            state.filter = f.key;\n            render(content);\n        });\n    });\n\n    // 进度提示\n    if (total > 0) {\n        const progress = container.createDiv({ cls: 'todo-progress' });\n        progress.textContent = `${t('mod.todo.progress')} ${completed} / ${total}${t('mod.todo.progressSuffix')}${total - completed}${t('mod.todo.progressSuffix2')}`;\n    }\n\n    // 列表区域\n    const listArea = container.createDiv({ cls: 'todo-list-area' });\n\n    const filtered = todos.filter(t => {\n        if (state.filter === 'active') return !t.completed;\n        if (state.filter === 'done') return t.completed;\n        return true;\n    });\n\n    if (filtered.length === 0) {\n        const empty = listArea.createDiv({ cls: 'todo-empty' });\n        empty.createEl('div', { text: '📝', cls: 'todo-empty-icon' });\n        empty.createEl('div', {\n            text: state.filter === 'done' ? t('mod.todo.empty.done') : t('mod.todo.empty.today'),\n            cls: 'todo-empty-text'\n        });\n        return;\n    }\n\n    filtered.forEach((todo) => {\n        const item = listArea.createDiv({ cls: 'todo-item' + (todo.completed ? ' completed' : '') });\n\n        const checkbox = item.createDiv({ cls: 'todo-checkbox' + (todo.completed ? ' checked' : '') });\n        if (todo.completed) checkbox.textContent = '✓';\n\n        const textEl = item.createEl('div', { text: todo.text, cls: 'todo-text' });\n        const deleteBtn = item.createEl('div', { text: '✕', cls: 'todo-delete' });\n\n        // 点击勾选/取消\n        checkbox.addEventListener('click', async (e) => {\n            e.stopPropagation();\n            await toggleTodo(filename, todo);\n            render(content);\n        });\n\n        // 双击编辑\n        textEl.addEventListener('dblclick', (e) => {\n            e.stopPropagation();\n            const editInput = item.createEl('input', {\n                cls: 'todo-text-edit',\n                attr: { value: todo.text }\n            });\n            textEl.remove();\n            editInput.select();\n            editInput.addEventListener('blur', async () => {\n                const newText = editInput.value.trim();\n                if (newText && newText !== todo.text) {\n                    await editTodo(filename, todo, newText);\n                }\n                render(content);\n            });\n            editInput.addEventListener('keypress', async (e) => {\n                if (e.key === 'Enter') {\n                    editInput.blur();\n                }\n            });\n            editInput.addEventListener('keydown', (e) => {\n                if (e.key === 'Escape') render(content);\n            });\n        });\n\n        // 删除\n        deleteBtn.addEventListener('click', async (e) => {\n            e.stopPropagation();\n            await deleteTodo(filename, todo);\n            render(content);\n        });\n    });\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    const { Setting } = require('obsidian');\n\n    containerEl.createEl('h3', { text: t('mod.todo.settings.title') });\n\n    new Setting(containerEl)\n        .setName(t('mod.todo.settings.folder'))\n        .setDesc(t('mod.todo.settings.folderDesc'))\n        .addText(t => {\n            t.setPlaceholder('待办')\n                .setValue(settings.folder || '待办')\n                .onChange(async (v) => {\n                    settings.folder = v.trim() || '待办';\n                    await saveCallback();\n                });\n        });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "url-opener": "// url-opener 模块 - 面板内浏览器\n// 源插件: url-webview-opener (笔记浏览器打开网址)\n// 核心功能保留: iframe内嵌浏览 + 书签管理\nconst id = 'url-opener';\nconst title = t('mod.urlOpener');\nconst icon = '🔗';\n\nconst defaultSettings = {\n    bookmarks: []\n};\n\nconst styles = `\n.uo-wrap { padding: 8px 10px; display: flex; flex-direction: column; height: 100%; }\n.uo-inputbar { display: flex; gap: 6px; margin-bottom: 6px; flex-shrink: 0; }\n.uo-inputbar input { flex: 1; padding: 5px 8px; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: var(--background-secondary); color: var(--text-normal); font-size: 12px; outline: none; min-width: 0; }\n.uo-inputbar input:focus { border-color: var(--v6-primary); }\n.uo-btn { padding: 5px 10px; border: none; border-radius: 4px; background: var(--v6-primary); color: white; cursor: pointer; font-size: 11px; transition: opacity 0.15s; white-space: nowrap; }\n.uo-btn:hover { opacity: 0.85; }\n.uo-btn.secondary { background: var(--background-modifier-form-field); color: var(--text-normal); border: 1px solid var(--background-modifier-border); }\n.uo-btn.danger { background: transparent; color: var(--text-muted); border: 1px solid var(--background-modifier-border); }\n.uo-btn.danger:hover { color: var(--text-error); border-color: var(--text-error); }\n.uo-navbar { display: flex; gap: 4px; margin-bottom: 4px; flex-shrink: 0; align-items: center; }\n.uo-navbar button { padding: 2px 7px; border: 1px solid var(--background-modifier-border); border-radius: 3px; background: var(--background-secondary); color: var(--text-muted); font-size: 11px; cursor: pointer; }\n.uo-navbar button:hover { color: var(--text-normal); background: var(--background-modifier-hover); }\n.uo-navbar button:disabled { opacity: 0.4; cursor: default; }\n.uo-urlbar { font-size: 10px; color: var(--text-faint); padding: 2px 6px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; min-width: 0; }\n.uo-bookmarks { max-height: 80px; overflow-y: auto; margin-bottom: 4px; flex-shrink: 0; border: 1px solid var(--background-modifier-border); border-radius: 4px; }\n.uo-bookmarks.hidden { display: none; }\n.uo-bm-item { display: flex; align-items: center; justify-content: space-between; padding: 3px 8px; cursor: pointer; font-size: 11px; transition: background 0.1s; }\n.uo-bm-item:hover { background: var(--background-modifier-hover); }\n.uo-bm-title { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.uo-bm-url { color: var(--text-faint); font-size: 10px; margin-left: 8px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 150px; }\n.uo-bm-del { color: var(--text-faint); cursor: pointer; padding: 1px 4px; font-size: 12px; opacity: 0; transition: opacity 0.15s; }\n.uo-bm-item:hover .uo-bm-del { opacity: 1; }\n.uo-bm-del:hover { color: var(--text-error); }\n.uo-viewer { flex: 1; border: 1px solid var(--background-modifier-border); border-radius: 4px; background: #fff; position: relative; min-height: 60px; }\n.uo-viewer iframe { width: 100%; height: 100%; border: none; }\n.uo-empty { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: var(--text-muted); font-size: 13px; text-align: center; }\n`;\n\nasync function render(container) {\n    container.addClass('uo-wrap');\n    var s = settings;\n\n    // 地址栏\n    var inputbar = container.createDiv({ cls: 'uo-inputbar' });\n    var urlInput = inputbar.createEl('input', { attr: { type: 'text', placeholder: t('mod.urlOpener.placeholder') } });\n    var goBtn = inputbar.createEl('button', { text: t('mod.urlOpener.btn.go'), cls: 'uo-btn' });\n    var bmBtn = inputbar.createEl('button', { text: t('mod.urlOpener.btn.bookmark'), cls: 'uo-btn secondary' });\n\n    // 导航栏\n    var navbar = container.createDiv({ cls: 'uo-navbar' });\n    var backBtn = navbar.createEl('button', { text: '◀', attr: { disabled: true, title: t('mod.urlOpener.btn.back') } });\n    var fwdBtn = navbar.createEl('button', { text: '▶', attr: { disabled: true, title: t('mod.urlOpener.btn.forward') } });\n    var refreshBtn = navbar.createEl('button', { text: '⟳', attr: { title: t('mod.urlOpener.btn.refresh') } });\n    var extBtn = navbar.createEl('button', { text: t('mod.urlOpener.btn.external'), attr: { title: '外部浏览器打开' } });\n    var urlBar = navbar.createDiv({ cls: 'uo-urlbar' });\n\n    // 书签栏\n    var bmToggle = container.createDiv({ cls: 'uo-navbar', attr: { style: 'margin-top:0;' } });\n    bmToggle.createEl('button', { text: t('mod.urlOpener.btn.bookmarks'), cls: 'uo-btn secondary' });\n    var bookmarks = container.createDiv({ cls: 'uo-bookmarks' });\n\n    // 查看器\n    var viewer = container.createDiv({ cls: 'uo-viewer' });\n    viewer.innerHTML = '<div class=\"uo-empty\">' + t('mod.urlOpener.hint') + '<br><small>' + t('mod.urlOpener.hintXFrame') + '</small></div>';\n\n    var iframe = null;\n    var history = [];\n    var historyIndex = -1;\n\n    // 加载书签\n    if (!s.bookmarks) s.bookmarks = [];\n    if (!Array.isArray(s.bookmarks)) s.bookmarks = [];\n\n    function renderBookmarks() {\n        bookmarks.innerHTML = '';\n        if (s.bookmarks.length === 0) {\n            bookmarks.createDiv({ text: t('mod.urlOpener.noBookmarks'), cls: 'uo-bm-item' }).style.cssText = 'cursor:default;color:var(--text-muted);justify-content:center;';\n        }\n        s.bookmarks.forEach(function(bm, idx) {\n            var item = bookmarks.createDiv({ cls: 'uo-bm-item' });\n            item.createSpan({ text: bm.title || bm.url, cls: 'uo-bm-title' });\n            item.createSpan({ text: bm.url, cls: 'uo-bm-url' });\n            var del = item.createSpan({ text: '✕', cls: 'uo-bm-del' });\n            item.addEventListener('mousedown', function(evt) { evt.preventDefault(); evt.stopPropagation(); }, true);\n            item.addEventListener('click', function(e) {\n                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();\n                if (e.target === del || e.target.closest('.uo-bm-del')) return;\n                navigate(bm.url);\n                urlInput.value = bm.url;\n            }, true);\n            del.addEventListener('click', function(e) {\n                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();\n                s.bookmarks.splice(idx, 1);\n                if (typeof saveCallback === 'function') saveCallback();\n                renderBookmarks();\n            });\n        });\n    }\n    renderBookmarks();\n\n    // 导航到URL\n    function navigate(url) {\n        if (!url) return;\n        // 自动补全\n        if (!/^https?:\\/\\//i.test(url)) url = 'https://' + url;\n\n        viewer.innerHTML = '';\n        iframe = viewer.createEl('iframe');\n        iframe.src = url;\n        iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin allow-forms allow-popups');\n\n        urlBar.textContent = url;\n        urlInput.value = url;\n\n        // 添加到历史\n        if (historyIndex >= 0 && history[historyIndex] === url) return;\n        // 删除当前位置之后的历史\n        history = history.slice(0, historyIndex + 1);\n        history.push(url);\n        historyIndex = history.length - 1;\n\n        updateNavButtons();\n    }\n\n    function updateNavButtons() {\n        backBtn.disabled = historyIndex <= 0;\n        fwdBtn.disabled = historyIndex >= history.length - 1;\n    }\n\n    // 事件\n    goBtn.addEventListener('click', function() { navigate(urlInput.value); });\n    urlInput.addEventListener('keydown', function(e) { if (e.key === 'Enter') navigate(urlInput.value); });\n\n    backBtn.addEventListener('click', function() {\n        if (historyIndex > 0) { historyIndex--; loadHistoryUrl(); }\n    });\n    fwdBtn.addEventListener('click', function() {\n        if (historyIndex < history.length - 1) { historyIndex++; loadHistoryUrl(); }\n    });\n\n    function loadHistoryUrl() {\n        var url = history[historyIndex];\n        if (iframe) iframe.src = url;\n        urlInput.value = url;\n        urlBar.textContent = url;\n        updateNavButtons();\n    }\n\n    refreshBtn.addEventListener('click', function() {\n        if (iframe) {\n            var src = iframe.src;\n            iframe.src = '';\n            setTimeout(function() { iframe.src = src; }, 50);\n        }\n    });\n\n    extBtn.addEventListener('click', function() {\n        var url = urlInput.value;\n        if (url) window.open(url, '_blank');\n    });\n\n    bmBtn.addEventListener('click', function() {\n        var url = urlInput.value;\n        if (!url) return;\n        if (!/^https?:\\/\\//i.test(url)) url = 'https://' + url;\n\n        // 检查是否已存在\n        var exists = s.bookmarks.some(function(b) { return b.url === url; });\n        if (exists) {\n            new Notice(t('mod.urlOpener.bookmarkExists'));\n            return;\n        }\n        s.bookmarks.push({ title: url.replace(/^https?:\\/\\//, '').split('/')[0], url: url });\n        if (typeof saveCallback === 'function') saveCallback();\n        renderBookmarks();\n        new Notice(t('mod.urlOpener.bookmarkAdded'));\n    });\n\n    bmToggle.addEventListener('click', function() { bookmarks.classList.toggle('hidden'); });\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.urlOpener.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.urlOpener.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n    containerEl.createEl('p', {\n        text: t('mod.urlOpener.settings.hint'),\n        attr: { style: 'color:var(--text-muted);font-size:11px;' }\n    });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "vault-stats": "/**\n * 笔记统计模块 (替代 better-word-count)\n * 功能：文件数、字数、文件夹数、最近修改数、最大文件列表\n */\nconst id = 'vault-stats';\nconst title = t('mod.vaultStats');\nconst icon = '📈';\n\nconst defaultSettings = {\n    countComments: true,\n    pageWords: 300\n};\n\nconst styles = `\n.vault-stats-wrap { padding: 12px; }\n.vs-header {\n    display: flex;\n    align-items: center;\n    justify-content: space-between;\n    margin-bottom: 10px;\n}\n.vs-header h4 { font-size: 13px; margin: 0; color: var(--text-normal); }\n.vs-refresh {\n    padding: 4px 10px;\n    border-radius: 4px;\n    border: 1px solid var(--background-modifier-border);\n    background: var(--background-modifier-form-field);\n    color: var(--text-muted);\n    cursor: pointer;\n    font-size: 11px;\n}\n.vs-refresh:hover { color: var(--text-normal); background: var(--background-modifier-hover); }\n.vs-grid {\n    display: grid;\n    grid-template-columns: repeat(2, 1fr);\n    gap: 8px;\n    margin-bottom: 12px;\n}\n.vs-card {\n    background: var(--background-modifier-form-field);\n    border-radius: 8px;\n    padding: 10px;\n    text-align: center;\n}\n.vs-card-icon { font-size: 18px; margin-bottom: 4px; }\n.vs-card-value { font-size: 20px; font-weight: 700; color: var(--text-normal); }\n.vs-card-label { font-size: 10px; color: var(--text-muted); margin-top: 2px; }\n.vs-section-title {\n    font-size: 12px;\n    font-weight: 600;\n    color: var(--text-normal);\n    margin: 12px 0 6px;\n    padding-bottom: 4px;\n    border-bottom: 1px solid var(--background-modifier-border);\n}\n.vs-large-list { list-style: none; padding: 0; margin: 0; }\n.vs-large-item {\n    display: flex;\n    align-items: center;\n    gap: 8px;\n    padding: 4px 0;\n    font-size: 11px;\n    color: var(--text-normal);\n    cursor: pointer;\n    border-radius: 4px;\n    transition: background 0.15s;\n}\n.vs-large-item:hover { background: var(--background-modifier-hover); }\n.vs-large-rank {\n    width: 18px;\n    height: 18px;\n    border-radius: 50%;\n    background: var(--background-secondary);\n    color: var(--text-muted);\n    display: flex;\n    align-items: center;\n    justify-content: center;\n    font-size: 10px;\n    flex-shrink: 0;\n}\n.vs-large-rank.top { background: var(--interactive-accent); color: var(--text-on-accent); }\n.vs-large-name {\n    flex: 1;\n    min-width: 0;\n    overflow: hidden;\n    text-overflow: ellipsis;\n    white-space: nowrap;\n}\n.vs-large-size {\n    color: var(--text-muted);\n    font-size: 10px;\n    flex-shrink: 0;\n}\n.vs-loading {\n    text-align: center;\n    padding: 24px;\n    color: var(--text-muted);\n    font-size: 12px;\n}\n`;\n\nfunction formatSize(bytes) {\n    if (bytes < 1024) return bytes + ' B';\n    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';\n    return (bytes / 1048576).toFixed(1) + ' MB';\n}\n\nfunction formatNumber(num) {\n    if (num >= 10000) return (num / 10000).toFixed(1) + '万';\n    return num.toLocaleString();\n}\n\nfunction countWords(text, countComments) {\n    let cleaned = text\n        .replace(/^---[\\s\\S]*?---\\n?/, '')\n        .replace(/```[\\s\\S]*?```/g, '')\n        .replace(/`[^`]*`/g, '')\n        .replace(/[#*\\[\\]>!\\-_~|]/g, '');\n    if (!countComments) {\n        cleaned = cleaned.replace(/%%[\\s\\S]*?%%/g, '');\n    }\n    return Math.ceil(cleaned.replace(/\\s+/g, '').length / 2);\n}\n\nasync function render(content) {\n    content.empty();\n    const wrap = content.createDiv({ cls: 'vault-stats-wrap' });\n\n    const header = wrap.createDiv({ cls: 'vs-header' });\n    header.createEl('h4', { text: t('mod.vaultStats.title') });\n    const refreshBtn = header.createEl('button', { cls: 'vs-refresh', text: t('mod.vaultStats.refresh') });\n\n    const loading = wrap.createEl('div', { cls: 'vs-loading', text: t('mod.vaultStats.loading') });\n\n    async function doStats(full) {\n        wrap.querySelectorAll('.vs-grid, .vs-section, .vs-large-list').forEach(el => el.remove());\n        loading.style.display = 'block';\n\n        const files = app.vault.getFiles();\n        const mdFiles = files.filter(f => f.extension === 'md');\n\n        let totalWords = 0;\n        let totalSize = 0;\n        const folderSet = new Set();\n        const recentCount = { count: 0, threshold: Date.now() - 7 * 24 * 60 * 60 * 1000 };\n        const largeFiles = [];\n\n        for (const file of files) {\n            totalSize += file.stat.size;\n            const parts = file.path.split('/');\n            if (parts.length > 1) folderSet.add(parts.slice(0, -1).join('/'));\n            if (file.stat.mtime > recentCount.threshold) recentCount.count++;\n        }\n\n        // ★ 修复：初始加载只做文件级统计（不读内容），避免卡死\n        // 完整的字数统计仅在用户点击\"刷新\"按钮时执行\n        if (full) {\n            for (const file of mdFiles) {\n                try {\n                    const text = await app.vault.read(file);\n                    totalWords += countWords(text, settings.countComments !== false);\n                } catch (e) { /* skip */ }\n                largeFiles.push({ name: file.name, path: file.path, size: file.stat.size });\n            }\n        } else {\n            // 轻量模式：只统计文件大小，不读内容\n            for (const file of mdFiles) {\n                largeFiles.push({ name: file.name, path: file.path, size: file.stat.size });\n            }\n            totalWords = -1; // 标记为\"未统计\"\n        }\n\n        largeFiles.sort((a, b) => b.size - a.size);\n        const topLarge = largeFiles.slice(0, 5);\n        const pageEstimate = settings.pageWords ? Math.round(totalWords / settings.pageWords) : 0;\n\n        loading.style.display = 'none';\n\n        // 统计卡片\n        const grid = wrap.createDiv({ cls: 'vs-grid' });\n        const cards = [\n            { icon: '📄', value: mdFiles.length, label: t('mod.vaultStats.totalNotes') },\n            { icon: '✏️', value: totalWords, label: t('mod.vaultStats.totalWords') },\n            { icon: '📁', value: folderSet.size, label: t('mod.vaultStats.folderCount') },\n            { icon: '🕐', value: recentCount.count, label: t('mod.vaultStats.recent7d') },\n            { icon: '📐', value: pageEstimate, label: t('mod.vaultStats.estPages') },\n            { icon: '💾', value: formatSize(totalSize), label: t('mod.vaultStats.vaultSize') }\n        ];\n\n        cards.forEach(c => {\n            const card = grid.createDiv({ cls: 'vs-card' });\n            card.createEl('div', { cls: 'vs-card-icon', text: c.icon });\n            const val = c.value;\n            if (val === -1) {\n                card.createEl('div', { cls: 'vs-card-value', text: '—', attr: { style: 'font-size:14px;color:var(--text-muted);' } });\n                card.createEl('div', { cls: 'vs-card-label', text: c.label + ' (点击刷新)' });\n            } else {\n                card.createEl('div', { cls: 'vs-card-value', text: typeof val === 'number' ? formatNumber(val) : val });\n                card.createEl('div', { cls: 'vs-card-label', text: c.label });\n            }\n        });\n\n        // 最大文件列表\n        if (topLarge.length > 0) {\n            wrap.createEl('div', { cls: 'vs-section-title', text: t('mod.vaultStats.largestNote') });\n            const list = wrap.createEl('ul', { cls: 'vs-large-list' });\n            topLarge.forEach((f, i) => {\n                const li = list.createEl('li', { cls: 'vs-large-item' });\n                const rank = li.createEl('span', { cls: 'vs-large-rank' + (i < 3 ? ' top' : ''), text: String(i + 1) });\n                li.createEl('span', { cls: 'vs-large-name', text: f.name });\n                li.createEl('span', { cls: 'vs-large-size', text: formatSize(f.size) });\n                li.addEventListener('click', () => {\n                    app.workspace.openLinkText(f.path, '', false);\n                });\n            });\n        }\n    }\n\n        // 当前文件统计\n        try {\n            var activeFile = app.workspace.getActiveFile();\n            if (activeFile && activeFile.extension === 'md') {\n                wrap.createEl('div', { cls: 'vs-section-title', text: t('mod.vaultStats.currentNote') });\n                var currentStats = wrap.createDiv({ cls: 'vs-grid', attr: { style: 'margin-top:4px;' } });\n                var curText = await app.vault.read(activeFile);\n                var curWords = countWords(curText, settings.countComments !== false);\n                var curChars = curText.replace(/\\s/g, '').length;\n                var curLines = curText.split('\\n').length;\n                var curCards = [\n                    { icon: '📄', value: activeFile.name, label: t('mod.vaultStats.fileName'), isSmall: true },\n                    { icon: '✏️', value: formatNumber(curWords), label: t('mod.vaultStats.wordCount') },\n                    { icon: '🔤', value: formatNumber(curChars), label: t('mod.vaultStats.charCount') },\n                    { icon: '📏', value: formatNumber(curLines), label: t('mod.vaultStats.lineCount') }\n                ];\n                curCards.forEach(function(c) {\n                    var card = currentStats.createDiv({ cls: 'vs-card' });\n                    if (c.isSmall) {\n                        card.style.cssText = 'grid-column: span 2;';\n                        card.createEl('div', { cls: 'vs-card-icon', text: c.icon });\n                        card.createEl('div', { cls: 'vs-card-value', text: c.value, attr: { style: 'font-size:13px;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;' } });\n                        card.createEl('div', { cls: 'vs-card-label', text: c.label });\n                    } else {\n                        card.createEl('div', { cls: 'vs-card-icon', text: c.icon });\n                        card.createEl('div', { cls: 'vs-card-value', text: c.value });\n                        card.createEl('div', { cls: 'vs-card-label', text: c.label });\n                    }\n                });\n            }\n        } catch(e) {}\n\n    // ★ 修复：初始加载使用轻量模式（不读文件内容），避免卡死\n    setTimeout(function() { doStats(false); }, 500);\n    refreshBtn.addEventListener('click', () => doStats(true));\n}\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    const { Setting } = require('obsidian');\n\n    containerEl.createEl('h3', { text: t('mod.vaultStats.settings.title') });\n\n    new Setting(containerEl)\n        .setName(t('mod.vaultStats.settings.desc'))\n        .setDesc('字数统计时包含 %%注释%% 内容')\n        .addToggle(t => t.setValue(settings.countComments !== false).onChange(async v => {\n            settings.countComments = v;\n            await saveCallback();\n        }));\n\n    new Setting(containerEl)\n        .setName(t('mod.vaultStats.settings.pageWords'))\n        .setDesc(t('mod.vaultStats.settings.pageWordsDesc'))\n        .addText(t => t.setValue(String(settings.pageWords || 300)).onChange(async v => {\n            const n = parseInt(v);\n            if (!isNaN(n) && n > 0) {\n                settings.pageWords = n;\n                await saveCallback();\n            }\n        }));\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "weather": "/**\n * 天气模块 V16 - 多平台适配器架构\n *\n * 支持的天气平台：\n *   1. 高德地图 Amap        — 需API Key，中国覆盖极佳\n *   2. Open-Meteo          — 无需Key，免费无限，19国气象模型，全球+中国\n *   3. wttr.in             — 无需Key，极简易用，城市名直接查\n *   4. OpenWeatherMap      — 需API Key，行业标准，1000次/天免费\n *   5. 自定义              — 用户自定义URL模板（高级）\n *\n * 所有平台统一输出为相同的内部数据格式 → render() 只关心统一格式\n */\n\n/** 获取当前语言代码 (zh/en) — 优先用 t() 确保与 i18n 同步 */\nfunction _isZh() {\n    // ★ 最优先：用 t() 检测，与 UI 标签的 t() 共享同一个 _currentLang\n    try {\n        if (typeof t === 'function') {\n            var v = t('mod.weather');\n            if (v === '天气') return true;\n            if (v === 'Weather') return false;\n        }\n    } catch(e) {}\n    // 回退1：getCurrentLang 全局函数\n    try {\n        var fn = typeof getCurrentLang === 'function' ? getCurrentLang : null;\n        if (!fn && typeof window !== 'undefined' && window.getCurrentLang) fn = window.getCurrentLang;\n        if (fn) { var l = fn(); if (l) return l.startsWith('zh'); }\n    } catch(e) {}\n    // 回退2：检查 html lang 属性\n    try {\n        var htmlLang = document.documentElement.getAttribute('lang') || '';\n        if (htmlLang) return htmlLang.startsWith('zh');\n    } catch(e) {}\n    return true; // 默认中文\n}\n\nconst id = 'weather';\nconst title = (function() {\n    try { if (typeof t === 'function') { var v = t('mod.weather'); if (v && v !== 'mod.weather') return v; } } catch(e) {}\n    return _isZh() ? '天气' : 'Weather';\n})();\nconst icon = '🌤️';\n\n// ══════════════════════════════════════\n//  默认设置\n// ══════════════════════════════════════\nconst defaultSettings = {\n    provider: 'amap',       // amap | open-meteo | wttr | openweathermap | custom\n    city: '北京',\n    apiKey: '',\n    customApiUrl: ''        // 自定义平台的 URL 模板（{city} 占位符）\n};\n\n// ══════════════════════════════════════\n//  样式（与V15完全一致）\n// ══════════════════════════════════════\nconst styles = `\n.weather-wrap {\n    padding: 0;\n    height: 100%;\n    display: flex;\n    flex-direction: column;\n}\n/* 顶部实况区 - 居中 */\n.weather-live {\n    padding: 16px 14px 12px;\n    display: flex;\n    flex-direction: column;\n    align-items: center;\n    text-align: center;\n    gap: 4px;\n}\n.weather-emoji {\n    font-size: 52px;\n    line-height: 1;\n}\n.weather-city-line {\n    display: flex;\n    align-items: center;\n    gap: 8px;\n    margin-top: 4px;\n}\n.weather-city {\n    font-size: 15px;\n    font-weight: 600;\n    color: var(--v6-text);\n}\n.weather-update-time {\n    font-size: 10px;\n    color: var(--v6-muted);\n    background: var(--background-modifier-form-field);\n    padding: 1px 6px;\n    border-radius: 10px;\n}\n.weather-temp-main {\n    font-size: 36px;\n    font-weight: 700;\n    color: var(--v6-primary);\n    line-height: 1.1;\n}\n.weather-temp-main .unit {\n    font-size: 20px;\n    font-weight: 400;\n    margin-left: 1px;\n}\n.weather-weather-text {\n    font-size: 13px;\n    color: var(--text-muted);\n}\n/* 实况详情网格 */\n.weather-detail-grid {\n    display: grid;\n    grid-template-columns: repeat(3, 1fr);\n    gap: 6px;\n    padding: 0 14px 10px;\n}\n.weather-detail-cell {\n    background: var(--background-modifier-form-field);\n    border-radius: 8px;\n    padding: 8px 6px;\n    text-align: center;\n}\n.weather-detail-cell .label {\n    font-size: 10px;\n    color: var(--text-muted);\n    margin-bottom: 2px;\n}\n.weather-detail-cell .value {\n    font-size: 13px;\n    font-weight: 600;\n    color: var(--text-normal);\n}\n/* 预报区 */\n.weather-forecast-wrap {\n    flex: 1;\n    overflow: auto;\n    padding: 0 14px 10px;\n}\n.weather-forecast-title {\n    font-size: 11px;\n    color: var(--text-muted);\n    font-weight: 600;\n    text-transform: uppercase;\n    letter-spacing: 0.5px;\n    margin-bottom: 6px;\n    padding-left: 2px;\n}\n.weather-forecast-list {\n    display: flex;\n    flex-direction: column;\n    gap: 6px;\n}\n.weather-forecast-card {\n    display: flex;\n    align-items: center;\n    gap: 10px;\n    padding: 8px 10px;\n    background: var(--background-modifier-form-field);\n    border-radius: 8px;\n}\n.weather-forecast-card .day-label {\n    width: 32px;\n    font-size: 11px;\n    font-weight: 600;\n    color: var(--text-muted);\n    text-align: center;\n}\n.weather-forecast-card .f-emoji {\n    font-size: 22px;\n    flex-shrink: 0;\n}\n.weather-forecast-card .f-desc {\n    flex: 1;\n    font-size: 12px;\n    color: var(--text-normal);\n}\n.weather-forecast-card .f-temp {\n    font-size: 12px;\n    font-weight: 600;\n    color: var(--v6-primary);\n    text-align: right;\n    white-space: nowrap;\n}\n.weather-forecast-card .f-temp .night {\n    font-size: 10px;\n    color: var(--text-muted);\n    font-weight: 400;\n}\n/* 错误/空状态 */\n.weather-empty {\n    display: flex;\n    flex-direction: column;\n    align-items: center;\n    justify-content: center;\n    height: 100%;\n    text-align: center;\n    padding: 24px;\n    color: var(--text-muted);\n    gap: 8px;\n}\n.weather-empty .big-icon {\n    font-size: 40px;\n    opacity: 0.6;\n}\n.weather-empty .tip {\n    font-size: 12px;\n    line-height: 1.5;\n}\n.weather-empty .link {\n    font-size: 11px;\n    color: var(--v6-primary);\n    cursor: pointer;\n}\n.weather-error {\n    display: flex;\n    flex-direction: column;\n    align-items: center;\n    justify-content: center;\n    height: 100%;\n    text-align: center;\n    padding: 24px;\n    color: var(--text-error);\n    gap: 6px;\n}\n.weather-error .big-icon {\n    font-size: 32px;\n}\n.weather-error .msg {\n    font-size: 12px;\n    line-height: 1.5;\n}\n.weather-error .retry {\n    font-size: 11px;\n    color: var(--v6-primary);\n    cursor: pointer;\n    margin-top: 4px;\n}\n`;\n\n// ══════════════════════════════════════\n//  天气图标映射（中文 + WMO代码）\n// ══════════════════════════════════════\n\n/** 中文天气描述 → emoji */\nconst iconMapCN = {\n    '晴': '☀️', '少云': '🌤️', '多云': '⛅', '阴': '☁️',\n    '阵雨': '🌦️', '小雨': '🌧️', '中雨': '🌧️', '大雨': '⛈️',\n    '暴雨': '⛈️', '雷阵雨': '⛈️', '小雪': '🌨️', '中雪': '❄️',\n    '大雪': '❄️', '雾': '🌫️', '霾': '🌫️', '风': '💨',\n    '沙尘': '💨'\n};\n\n/** 英文天气描述 → emoji */\nconst iconMapEN = {\n    'clear': '☀️', 'sunny': '☀️',\n    'partly cloudy': '⛅', 'cloudy': '⛅', 'overcast': '☁️',\n    'fog': '🌫️', 'mist': '🌫️',\n    'light drizzle': '🌦️', 'drizzle': '🌦️',\n    'light rain': '🌧️', 'moderate rain': '🌧️', 'heavy rain': '⛈️',\n    'light showers': '🌦️', 'showers': '🌦️',\n    'snow': '❄️', 'light snow': '🌨️', 'heavy snow': '❄️',\n    'thunderstorm': '⛈️', 'tropical storm': '⛈️',\n    'blowing snow': '💨', 'windy': '💨'\n};\n\n/**\n * WMO 天气现象代码 → { zh, en, emoji }\n * https://open-meteo.com/en/docs#weathervariables\n */\nconst wmoCodeMap = {\n    0:  { zh: '晴',           en: 'Clear sky',            emoji: '☀️' },\n    1:  { zh: '晴间多云',     en: 'Mainly clear',          emoji: '🌤️' },\n    2:  { zh: '部分多云',     en: 'Partly cloudy',         emoji: '🌤️' },\n    3:  { zh: '阴天',         en: 'Overcast',              emoji: '☁️' },\n    45: { zh: '雾',           en: 'Fog',                   emoji: '🌫️' },\n    48: { zh: '雾凇',         en: 'Depositing rime fog',   emoji: '🌫️' },\n    51: { zh: '小毛毛雨',     en: 'Light drizzle',         emoji: '🌦️' },\n    53: { zh: '中毛毛雨',     en: 'Moderate drizzle',      emoji: '🌦️' },\n    55: { zh: '大毛毛雨',     en: 'Dense drizzle',         emoji: '🌧️' },\n    56: { zh: '冻毛毛雨',     en: 'Light freezing drizzle', emoji: '🌧️' },\n    57: { zh: '大冻毛毛雨',   en: 'Dense freezing drizzle',emoji: '🌧️' },\n    61: { zh: '小雨',         en: 'Slight rain',           emoji: '🌧️' },\n    63: { zh: '中雨',         en: 'Moderate rain',         emoji: '🌧️' },\n    65: { zh: '大雨',         en: 'Heavy rain',            emoji: '🌧️' },\n    66: { zh: '冻小雨',       en: 'Light freezing rain',   emoji: '🌧️' },\n    67: { zh: '冻大雨',       en: 'Heavy freezing rain',   emoji: '⛈️' },\n    71: { zh: '小雪',         en: 'Slight snow fall',      emoji: '🌨️' },\n    73: { zh: '中雪',         en: 'Moderate snow fall',    emoji: '❄️' },\n    75: { zh: '大雪',         en: 'Heavy snow fall',       emoji: '❄️' },\n    77: { zh: '雪粒',         en: 'Snow grains',           emoji: '❄️' },\n    80: { zh: '小阵雨',       en: 'Slight rain showers',  emoji: '🌦️' },\n    81: { zh: '中阵雨',       en: 'Moderate rain showers', emoji: '🌧️' },\n    82: { zh: '大阵雨',       en: 'Violent rain showers',  emoji: '⛈️' },\n    85: { zh: '小阵雪',       en: 'Slight snow showers',   emoji: '🌨️' },\n    86: { zh: '大阵雪',       en: 'Heavy snow showers',    emoji: '❄️' },\n    95: { zh: '雷暴',         en: 'Thunderstorm',          emoji: '⛈️' },\n    96: { zh: '雷暴伴冰雹',   en: 'Thunderstorm with hail', emoji: '⛈️' },\n    99: { zh: '强雷暴伴冰雹',en: 'Severe thunderstorm',    emoji: '⛈️' }\n};\n\n/** OWM 天气条件代码 → 中文描述 */\nconst owmConditionMap = {\n    200: '雷暴伴弱雨', 201: '雷暴', 202: '强雷暴', 230: '轻雷暴伴雾淞',\n    231: '雷暴伴雾凇', 232: '强雷暴伴雾凇', 233: '强阵风',\n    300: '小毛毛雨', 301: '中毛毛雨', 302: '大毛毛雨',\n    500: '小雨', 501: '中雨', 502: '大雨', 503: '暴雨', 504: '特大暴雨',\n    511: '冻雨', 520: '小阵雨', 521: '中阵雨', 522: '大阵雨', 531: '强阵雨',\n    600: '小雪', 601: '中雪', 602: '大雪', 603: '暴雪', 620: '小阵雪', 622: '大阵雪',\n    701: '薄雾', 711: '烟雾', 721: '霾', 731: '沙尘', 751: '沙暴',\n    800: '晴', 801: '少云', 802: '多云', 803: '阴天', 804: '阴'\n};\n\nfunction getWeatherIcon(w) {\n    if (!w) return '🌤️';\n    var wl = w.toLowerCase();\n    // 先匹配中文\n    for (var [key, val] of Object.entries(iconMapCN)) {\n        if (w.includes(key)) return val;\n    }\n    // 再匹配英文\n    for (var [key, val] of Object.entries(iconMapEN)) {\n        if (wl.includes(key)) return val;\n    }\n    return '🌤️';\n}\n\nfunction getWmoInfo(code) {\n    var info = wmoCodeMap[code];\n    if (!info) return { zh: '未知', en: 'Unknown (' + code + ')', emoji: '🌤️' };\n    return info;\n}\n\n// ══════════════════════════════════════\n//  统一数据格式（所有平台的输出都转换为此格式）\n// ══════════════════════════════════════\n//\n// {\n//   city: string,\n//   updateTime: string,        // \"14:00\"\n//   temperature: number,       // 25\n//   weatherText: string,       // \"多云\" / \"Partly cloudy\"\n//   humidity: number,          // 65 (%)\n//   windDirection: string,     // \"东南\" / \"NNE\"\n//   windSpeed: string,         // \"≤3级\" / \"7 km/h\"\n//   forecast: [\n//     {\n//       dayLabel: string,      // \"明天\"\n//       weatherText: string,\n//       tempMax: number,\n//       tempMin: number,\n//       emoji: string\n//     }\n//   ]\n// }\n\n// ══════════════════════════════════════\n//  平台 1：高德地图 Amap（原有逻辑不变）\n// ══════════════════════════════════════\n\nasync function _amap_fetchGeo(city, apiKey) {\n    const url = 'https://restapi.amap.com/v3/geocode/geo?address=' + encodeURIComponent(city) + '&key=' + apiKey;\n    const res = await requestUrl({ url, method: 'GET' });\n    const data = res.json;\n    if (!data || data.status !== '1' || !data.geocodes || data.geocodes.length === 0) {\n        throw new Error(t('mod.weather.error.cityNotFound'));\n    }\n    return data.geocodes[0];\n}\n\nasync function _amap_fetchWeather(adcode, apiKey) {\n    const [liveRes, forecastRes] = await Promise.all([\n        requestUrl({ url: 'https://restapi.amap.com/v3/weather/weatherInfo?city=' + adcode + '&key=' + apiKey + '&extensions=base', method: 'GET' }),\n        requestUrl({ url: 'https://restapi.amap.com/v3/weather/weatherInfo?city=' + adcode + '&key=' + apiKey + '&extensions=all', method: 'GET' })\n    ]);\n    const liveData = liveRes.json;\n    const forecastData = forecastRes.json;\n    if (!liveData || liveData.status !== '1' || !liveData.lives || liveData.lives.length === 0) {\n        throw new Error(t('mod.weather.error.fetchFailed') + (liveData && liveData.info ? liveData.info : t('mod.weather.error.unknown')));\n    }\n    return {\n        live: liveData.lives[0],\n        forecast: forecastData && forecastData.status === '1' && forecastData.forecasts ? forecastData.forecasts[0] : null\n    };\n}\n\nasync function fetchAmapWeather(city, apiKey) {\n    const geo = await _amap_fetchGeo(city, apiKey);\n    const adcode = geo.adcode;\n    const cityName = geo.district || geo.city || geo.formatted_address || city;\n\n    const { live, forecast } = await _amap_fetchWeather(adcode, apiKey);\n\n    // 预报\n    var fc = [];\n    if (forecast && forecast.casts && forecast.casts.length > 1) {\n        fc = forecast.casts.slice(1, 4).map(function(day, i) {\n            return {\n                dayLabel: i === 0 ? t('mod.weather.tomorrow') : (i === 1 ? t('mod.weather.dayAfter') : ''),\n                weatherText: day.dayweather + (day.nightweather && day.nightweather !== day.dayweather ? t('mod.weather.turnTo') + day.nightweather : ''),\n                tempMax: parseFloat(day.daytemp) || 0,\n                tempMin: parseFloat(day.nighttemp) || 0,\n                emoji: getWeatherIcon(day.dayweather)\n            };\n        });\n    }\n\n    // 解析更新时间\n    var timeStr = live.reporttime ? live.reporttime.split(' ')[1] || live.reporttime : '';\n\n    return {\n        city: cityName,\n        updateTime: timeStr,\n        temperature: parseFloat(live.temperature) || 0,\n        weatherText: live.weather,\n        humidity: parseFloat(live.humidity) || 0,\n        windDirection: live.winddirection || '--',\n        windSpeed: (live.windpower || '--') + t('mod.weather.scaleUnit'),\n        forecast: fc\n    };\n}\n\n// ══════════════════════════════════════\n//  平台 2：Open-Meteo（无需 API Key）\n// ══════════════════════════════════════\n\nasync function fetchOpenMeteoWeather(city) {\n    // Step 1: 地理编码（使用 Open-Meteo 内置地理编码 API）\n    var geoUrl = 'https://geocoding-api.open-meteo.com/v1/search?name=' + encodeURIComponent(city) + '&count=1&language=' + (typeof getCurrentLang === 'function' && getCurrentLang().startsWith('zh') ? 'zh' : 'en');\n    const geoRes = await requestUrl({ url: geoUrl, method: 'GET' });\n    const geoData = geoRes.json;\n\n    if (!geoData.results || geoData.results.length === 0) {\n        throw new Error(t('mod.weather.error.cityNotFound'));\n    }\n\n    var loc = geoData.results[0];\n    var lat = loc.latitude;\n    var lon = loc.longitude;\n    var cityName = loc.name;\n    // 尝试拼接更完整的城市名\n    if (loc.admin1 && !cityName.includes(loc.admin1)) {\n        cityName = cityName + ', ' + loc.admin1;\n    }\n\n    // Step 2: 获取天气数据\n    var weatherUrl = 'https://api.open-meteo.com/v1/forecast?' +\n        'latitude=' + lat + '&' +\n        'longitude=' + lon + '&' +\n        'current=temperature_2m,relative_humidity_2m,weather_code,wind_speed_10m,wind_direction_10m&' +\n        'daily=weather_code,temperature_2m_max,temperature_2m_min&' +\n        'timezone=auto&' +\n        'forecast_days=4';\n\n    const weatherRes = await requestUrl({ url: weatherUrl, method: 'GET' });\n    var wd = weatherRes.json;\n\n    if (!wd.current || typeof wd.current.weather_code === 'undefined') {\n        throw new Error(t('mod.weather.error.fetchFailed') + 'Open-Meteo');\n    }\n\n    // 当前天气\n    var curWmo = getWmoInfo(wd.current.weather_code);\n    var langKey = _isZh() ? 'zh' : 'en';\n\n    // 风向角度 → 方位文字\n    function degToDirection(deg) {\n        if (deg == null) return '--';\n        var dirs = ['北', '东北', '东', '东南', '南', '西南', '西', '西北'];\n        var dirsEn = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW'];\n        var idx = Math.round(deg / 22.5) % 16;\n        return _isZh() ? dirs[Math.floor(idx / 2)] : dirsEn[idx];\n    }\n\n    // 更新时间（取当前时间的 HH:mm 格式）\n    var now = new Date();\n    var timeStr = ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2);\n\n    // 预报\n    var fc = [];\n    if (wd.daily && wd.daily.time && wd.daily.time.length > 1) {\n        for (var fi = 1; fi < Math.min(4, wd.daily.time.length); fi++) {\n            var fWmo = getWmoInfo(wd.daily.weather_code[fi]);\n            fc.push({\n                dayLabel: fi === 1 ? t('mod.weather.tomorrow') : (fi === 2 ? t('mod.weather.dayAfter') : ''),\n                weatherText: fWmo[langKey],\n                tempMax: wd.daily.temperature_2m_max[fi] != null ? Math.round(wd.daily.temperature_2m_max[fi]) : 0,\n                tempMin: wd.daily.temperature_2m_min[fi] != null ? Math.round(wd.daily.temperature_2m_min[fi]) : 0,\n                emoji: fWmo.emoji\n            });\n        }\n    }\n\n    return {\n        city: cityName,\n        updateTime: timeStr,\n        temperature: Math.round(wd.current.temperature_2m),\n        weatherText: curWmo[langKey],\n        humidity: wd.current.relative_humidity_2m || 0,\n        windDirection: degToDirection(wd.current.wind_direction_10m),\n        windSpeed: (wd.current.wind_speed_10m != null ? Math.round(wd.current.wind_speed_10m) : '--') + ' km/h',\n        forecast: fc\n    };\n}\n\n// ══════════════════════════════════════\n//  平台 3：wttr.in（无需 API Key）\n// ══════════════════════════════════════\n\nasync function fetchWttrWeather(city) {\n    var url = 'https://wttr.in/' + encodeURIComponent(city) + '?format=j1';\n    const res = await requestUrl({ url: url, method: 'GET' });\n    var d = res.json;\n\n    if (!d || !d.current_condition || d.current_condition.length === 0) {\n        throw new Error(t('mod.weather.error.cityNotFound') + ' (wttr.in)');\n    }\n\n    var cc = d.current_condition[0];\n\n    // 天气描述翻译（英文 → 中文）\n    function translateWttrDesc(desc) {\n        if (!_isZh()) return desc;\n        var map = {\n            'Clear': '晴', 'Sunny': '晴', 'Partly cloudy': '部分多云', 'Cloudy': '多云',\n            'overcast': '阴天', 'Mist': '薄雾', 'Fog': '雾',\n            'Light drizzle': '小毛毛雨', 'Patchy rain possible': '局部可能小雨',\n            'Light rain': '小雨', 'Moderate rain': '中雨', 'Heavy rain': '大雨',\n            'Light snow': '小雪', 'Moderate snow': '中雪', 'Heavy snow': '大雪',\n            'Thundery outbreaks possible': '可能有雷暴', 'Blizzard': '暴风雪',\n            'Freezing fog': '冻雾', 'Ice pellets': '冰雹'\n        };\n        for (var [k, v] of Object.entries(map)) {\n            if (desc.toLowerCase().includes(k.toLowerCase())) return v;\n        }\n        return desc;\n    }\n\n    // 更新时间\n    var timeStr = cc.observation_time || '';\n    // wttr.in 时间可能是 \"04:08 PM\" 或 \"16:08\"\n    if (timeStr.match(/^\\d{1,2}:\\d{2}\\s*(AM|PM)$/i)) {\n        // 转换为24小时制\n        var parts = timeStr.replace(/\\s/g, '').split(':');\n        var h = parseInt(parts[0]);\n        var m = parts[1].replace(/AM|PM/i, '');\n        var isPm = /PM/i.test(parts[1]);\n        if (isPm && h !== 12) h += 12;\n        if (!isPm && h === 12) h = 0;\n        timeStr = ('0' + h).slice(-2) + ':' + m;\n    }\n\n    // 预报\n    var fc = [];\n    if (d.weather && d.weather.length > 1) {\n        for (var wi = 1; wi < Math.min(4, d.weather.length); wi++) {\n            var day = d.weather[wi];\n            // 提取当日中间时段的天气描述\n            var fDesc = '';\n            try {\n                if (day.hourly && day.hourly.length > 0) {\n                    var midHour = day.hourly[Math.floor(day.hourly.length / 2)];\n                    if (midHour.weatherDesc && midHour.weatherDesc[0]) {\n                        fDesc = midHour.weatherDesc[0].value;\n                    }\n                }\n            } catch(e) { /* 忽略解析错误 */ }\n\n            fc.push({\n                dayLabel: wi === 1 ? t('mod.weather.tomorrow') : (wi === 2 ? t('mod.weather.dayAfter') : ''),\n                weatherText: translateWttrDesc(fDesc || ((day.mintempC || '') + '°~' + (day.maxtempC || '') + '°')),\n                tempMax: parseFloat(day.maxtempC) || 0,\n                tempMin: parseFloat(day.mintempC) || 0,\n                emoji: getWeatherIcon(cc.weatherDesc && cc.weatherDesc[0] ? cc.weatherDesc[0].value : '')\n            });\n        }\n    }\n\n    return {\n        city: city,\n        updateTime: timeStr,\n        temperature: parseFloat(cc.temp_C) || 0,\n        weatherText: translateWttrDesc(cc.weatherDesc && cc.weatherDesc[0] ? cc.weatherDesc[0].value : ''),\n        humidity: parseFloat(cc.humidity) || 0,\n        windDirection: cc.winddir16Point || '--',\n        windSpeed: (cc.windspeedKmph || '--') + ' km/h',\n        forecast: fc\n    };\n}\n\n// ══════════════════════════════════════\n//  平台 4：OpenWeatherMap（需要 API Key）\n// ══════════════════════════════════════\n\nasync function fetchOWMWeather(city, apiKey) {\n    // Step 1: 当前天气\n    var currentUrl = 'https://api.openweathermap.org/data/2.5/weather?q=' + encodeURIComponent(city) + '&appid=' + apiKey + '&units=metric&lang=' + (_isZh() ? 'zh_cn' : 'en');\n\n    const curRes = await requestUrl({ url: currentUrl, method: 'GET' });\n    var cd = curRes.json;\n\n    if (!cd || cd.cod && cd.cod !== 200) {\n        throw new Error(t('mod.weather.error.cityNotFound') + ' (OWM: ' + (cd && cd.message ? cd.message : '') + ')');\n    }\n\n    // Step 2: 预报\n    var fcData = null;\n    try {\n        var fcUrl = 'https://api.openweathermap.org/data/2.5/forecast?q=' + encodeURIComponent(city) + '&appid=' + apiKey + '&units=metric&lang=' + (_isZh() ? 'zh_cn' : 'en') + '&cnt=20';\n        const fcRes = await requestUrl({ url: fcUrl, method: 'GET' });\n        fcData = fcRes.json;\n    } catch(e) {\n        // 预报获取失败不影响主功能\n        console.log('[DFV-PPT] OWM forecast failed:', e);\n    }\n\n    var conditionId = cd.weather && cd.weather[0] ? cd.weather[0].id : 800;\n    var weatherText = owmConditionMap[conditionId] || (cd.weather && cd.weather[0] ? cd.weather[0].description : '');\n\n    // 风向角度 → 文字\n    function degToDir(deg) {\n        if (deg == null) return '';\n        var dirsZh = ['北', '东北', '东', '东南', '南', '西南', '西', '西北'];\n        var dirsEn = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW'];\n        var idx = Math.round(deg / 22.5) % 16;\n        return _isZh() ? dirsZh[Math.floor(idx / 2)] : dirsEn[idx];\n    }\n\n    // 更新时间\n    var dt = cd.dt ? new Date(cd.dt * 1000) : new Date();\n    var timeStr = ('0' + dt.getHours()).slice(-2) + ':' + ('0' + dt.getMinutes()).slice(-2);\n\n    // 预报处理（按日期分组，取每天中间时段作为当日代表）\n    var fc = [];\n    if (fcData && fcData.list && fcData.list.length > 0) {\n        var dailyGroups = {};\n        for (var li = 0; li < fcData.list.length; li++) {\n            var item = fcData.list[li];\n            var dateStr = item.dt_txt ? item.dt_txt.split(' ')[0] : '';\n            if (!dateStr) continue;\n            if (!dailyGroups[dateStr]) dailyGroups[dateStr] = [];\n            dailyGroups[dateStr].push(item);\n        }\n\n        var dates = Object.keys(dailyGroups);\n        var today = new Date().toISOString().split('T')[0];\n\n        for (var di = 0; di < dates.length && fc.length < 3; di++) {\n            var dstr = dates[di];\n            if (dstr <= today) continue; // 跳过今天及之前的\n            var group = dailyGroups[dstr];\n            var midIdx = Math.floor(group.length / 2);\n            var rep = group[midIdx];\n            var repCond = rep.weather && rep.weather[0] ? rep.weather[0].id : 800;\n            var maxT = -Infinity, minT = Infinity;\n            for (var gi = 0; gi < group.length; gi++) {\n                var t = group[gi].main && group[gi].main.temp_max != null ? group[gi].main.temp_max :\n                         group[gi].main && group[gi].main.temp != null ? group[gi].main.temp : null;\n                if (t != null) {\n                    if (t > maxT) maxT = t;\n                    if (t < minT) minT = t;\n                }\n            }\n            // 如果没有 temp_max/temp_min 字段，从 main.temp 估算\n            if (maxT === -Infinity || minT === Infinity) {\n                maxT = rep.main && rep.main.temp_max != null ? rep.main.temp_max : (rep.main ? rep.main.temp : 0);\n                minT = rep.main && rep.main.temp_min != null ? rep.main.temp_min : (rep.main ? rep.main.temp : 0);\n            }\n\n            fc.push({\n                dayLabel: fc.length === 0 ? t('mod.weather.tomorrow') : (fc.length === 1 ? t('mod.weather.dayAfter') : ''),\n                weatherText: owmConditionMap[repCond] || (rep.weather && rep.weather[0] ? rep.weather[0].description : ''),\n                tempMax: Math.round(maxT),\n                tempMin: Math.round(minT),\n                emoji: getWeatherIcon(weatherText)\n            });\n        }\n    }\n\n    return {\n        city: cd.name || city,\n        updateTime: timeStr,\n        temperature: cd.main ? Math.round(cd.main.temp) : 0,\n        weatherText: weatherText,\n        humidity: cd.main ? Math.round(cd.main.humidity || 0) : 0,\n        windDirection: degToDir(cd.wind && cd.wind.deg),\n        windSpeed: (cd.wind && cd.wind.speed != null ? Math.round(cd.wind.speed) : '--') + (_isZh() ? ' m/s' : ' m/s'),\n        forecast: fc\n    };\n}\n\n// ══════════════════════════════════════\n//  主调度器：根据 provider 选择对应的适配器\n// ══════════════════════════════════════\n\nasync function fetchUnifiedWeather(provider, city, apiKey, customUrl) {\n    switch (provider) {\n        case 'open-meteo':\n            return fetchOpenMeteoWeather(city);\n\n        case 'wttr':\n            return fetchWttrWeather(city);\n\n        case 'openweathermap':\n            if (!apiKey) throw new Error(t('mod.weather.settings.owmNeedKey') || '请填写 OpenWeatherMap API Key');\n            return fetchOWMWeather(city, apiKey);\n\n        case 'custom':\n            if (!customUrl) throw new Error(t('mod.weather.settings.customNeedUrl') || '请填写自定义天气 API 的 URL 模板');\n            return fetchCustomWeather(city, customUrl);\n\n        case 'amap':\n        default:\n            if (!apiKey) throw new Error(t('mod.weather.error.noKey'));\n            return fetchAmapWeather(city, apiKey);\n    }\n}\n\n// ══════════════════════════════════════\n//  平台 5：自定义 URL 模板（高级）\n// ══════════════════════════════════════\n\nasync function fetchCustomWeather(city, templateUrl) {\n    var url = templateUrl.replace(/\\{city\\}/g, encodeURIComponent(city));\n    const res = await requestUrl({ url: url, method: 'GET' });\n    var d = res.json;\n\n    // 自定义模式：尝试智能解析返回的数据\n    // 支持多种常见格式：\n    // 格式A: { temperature: 25, weather: \"多云\", ... }\n    // 格式B: { current: { temp_c: 25, ... }, ... }\n    // 格式C: { data: [{ ... }] }\n    // 格式D: 数组 [{ ... }]\n    // 用户也可以在 render 中看到原始 JSON 进行调试\n\n    var root = Array.isArray(d) ? (d[0] || {}) : d;\n    var current = root.current || root.current_condition ? (root.current_condition ? root.current_condition[0] : root.current) : root;\n    var temp = current.temperature || current.temp || current.temp_C || current.temp_c || current.tempC ||\n               current.main && current.main.temp || current.data && current.data[0] && current.data[0].temp || 0;\n    var humid = current.humidity || current.humidityPercent || current.relative_humidity ||\n                 (current.main && current.main.humidity) || 0;\n    var wtext = current.weather || current.description || current.condition_text ||\n                current.weatherDesc && current.weatherDesc[0] && current.weatherDesc[0].value ||\n                current.weather_text || current.conditions || '';\n    if (Array.isArray(wtext)) wtext = wtext[0] && wtext[0].value ? wtext[0].value : '';\n    var wdir = current.wind_dir || current.windDirection || current.winddir16Point ||\n               (current.wind && current.wind.deg != null ? current.wind.deg + '°' : '') ||\n               (current.wind_dir || '');\n    var wspeed = current.wind_speed || current.windSpeed || current.windspeedKmph || current.windspeedMiles ||\n                  (current.wind && current.wind.speed != null ? current.wind.speed : '') || '';\n    var cname = root.city || root.location || root.name || current.city_name || current.nearest_area && current.nearest_area[0] && current.nearest_area[0].areaName && current.nearest_area[0].areaName[0].value || city;\n    if (Array.isArray(cname)) cname = cname[0] && cname[0].value ? cname[0].value : city;\n\n    // 预报\n    var fc = [];\n    var forecastRoot = root.forecast || root.forecasts || root.daily || (root.weather && Array.isArray(root.weather) ? root.weather.slice(1) : null);\n    if (forecastRoot && Array.isArray(forecastRoot)) {\n        for (var ci = 0; ci < Math.min(3, forecastRoot.length); ci++) {\n            var citem = forecastRoot[ci];\n            var ct = citem.temperature || citem.temp_C || citem.maxtempC || citem.tempMax || 0;\n            var cmin = citem.temp_min || citem.mintempC || citem.tempMin || ct;\n            var cw = citem.weather || citem.condition || citem.weatherDesc && citem.weatherDesc[0] && citem.weatherDesc[0].value || '';\n            fc.push({\n                dayLabel: ci === 0 ? t('mod.weather.tomorrow') : (ci === 1 ? t('mod.weather.dayAfter') : ''),\n                weatherText: String(cw),\n                tempMax: parseFloat(ct) || 0,\n                tempMin: parseFloat(cmin) || 0,\n                emoji: getWeatherIcon(String(cw))\n            });\n        }\n    }\n\n    return {\n        city: String(cname),\n       UpdateTime: current.observation_time || '',\n        temperature: parseFloat(temp) || 0,\n        weatherText: String(wtext),\n        humidity: parseFloat(humid) || 0,\n        windDirection: String(wdir),\n        windSpeed: String(wspeed),\n        forecast: fc\n    };\n}\n\n// ══════════════════════════════════════\n//  渲染函数（只依赖统一数据格式）\n// ══════════════════════════════════════\n\nasync function render(content) {\n    content.empty();\n    var wrap = content.createDiv({ cls: 'weather-wrap' });\n\n    var provider = settings.provider || 'amap';\n    var apiKey = settings.apiKey || '';\n    var city = settings.city || '北京';\n    var customUrl = settings.customApiUrl || '';\n\n    // ═══ 空状态检查（需要 Key 的平台） ═══\n    if ((provider === 'amap' || provider === 'openweathermap') && !apiKey) {\n        var empty = wrap.createDiv({ cls: 'weather-empty' });\n        empty.createEl('div', { text: '🔑', cls: 'big-icon' });\n\n        var tipMsg = provider === 'amap'\n            ? t('mod.weather.error.noKey')\n            : (t('mod.weather.settings.owmNeedKey') || t('mod.weather.settings.owmNeedKeyFallback'));\n        empty.createEl('div', { text: tipMsg, cls: 'tip' });\n\n        var linkUrl = provider === 'amap' ? 'https://lbs.amap.com/' : 'https://openweathermap.org/api';\n        var linkText = provider === 'amap' ? t('mod.weather.freeApply') : (t('mod.weather.owmFreeApply') || '👉 Apply for free');\n        var link = empty.createEl('div', { text: linkText, cls: 'link' });\n        link.addEventListener('click', () => window.open(linkUrl, '_blank'));\n\n        // 显示当前选择的是哪个平台提示\n        var provNames = {\n            amap: t('mod.weather.provider.amap') || 'Amap',\n            'open-meteo': t('mod.weather.provider.openmeteo') || 'Open-Meteo',\n            wttr: t('mod.weather.provider.wttr') || 'wttr.in',\n            openweathermap: t('mod.weather.provider.owm') || 'OpenWeatherMap',\n            custom: t('mod.weather.provider.custom') || 'Custom'\n        };\n        var hint = empty.createEl('div', {\n            text: t('mod.weather.settings.providerHint') + (provNames[provider] || provider),\n            style: 'font-size:10px;color:var(--text-muted);margin-top:8px;'\n        });\n        return;\n    }\n\n    try {\n        // ═══ 统一调度 ═══\n        var result = await fetchUnifiedWeather(provider, city, apiKey, customUrl);\n\n        // ===== 实况区（居中）=====\n        var liveSection = wrap.createDiv({ cls: 'weather-live' });\n        liveSection.createEl('div', { text: getWeatherIcon(result.weatherText), cls: 'weather-emoji' });\n\n        var cityLine = liveSection.createDiv({ cls: 'weather-city-line' });\n        cityLine.createEl('span', { text: result.city, cls: 'weather-city' });\n\n        // 平台标识（可选显示）\n        if (provider !== 'amap') {\n            var provTags = {\n                'open-meteo': 'OM',\n                'wttr': 'WTTR',\n                'openweathermap': 'OWM',\n                'custom': 'CUSTOM'\n            };\n            var tag = cityLine.createEl('span', {\n                text: provTags[provider] || '',\n                style: 'font-size:9px;background:var(--background-modifier-form-field);padding:1px 5px;border-radius:6px;font-weight:400;'\n            });\n        }\n\n        cityLine.createEl('span', { text: result.updateTime || '', cls: 'weather-update-time' });\n\n        liveSection.createEl('div', {\n            cls: 'weather-temp-main',\n            attr: { innerHTML: result.temperature + '<span class=\"unit\">°C</span>' }\n        });\n        liveSection.createEl('div', { text: result.weatherText, cls: 'weather-weather-text' });\n\n        // ===== 详情网格 =====\n        var detailGrid = wrap.createDiv({ cls: 'weather-detail-grid' });\n        var details = [\n            { label: t('mod.weather.humidity'), value: (result.humidity || '--') + '%' },\n            { label: t('mod.weather.windDirection'), value: (result.windDirection || '--') + t('mod.weather.windUnit') },\n            { label: t('mod.weather.windScale'), value: result.windSpeed || '--' }\n        ];\n        details.forEach(function(d) {\n            var cell = detailGrid.createDiv({ cls: 'weather-detail-cell' });\n            cell.createEl('div', { text: d.label, cls: 'label' });\n            cell.createEl('div', { text: d.value, cls: 'value' });\n        });\n\n        // ===== 预报区 =====\n        if (result.forecast && result.forecast.length > 0) {\n            var fWrap = wrap.createDiv({ cls: 'weather-forecast-wrap' });\n            fWrap.createEl('div', { text: t('mod.weather.forecast'), cls: 'weather-forecast-title' });\n\n            var fList = fWrap.createDiv({ cls: 'weather-forecast-list' });\n            result.forecast.forEach(function(day) {\n                var card = fList.createDiv({ cls: 'weather-forecast-card' });\n                card.createEl('div', { text: day.dayLabel, cls: 'day-label' });\n                card.createEl('div', { text: day.emoji || getWeatherIcon(day.weatherText), cls: 'f-emoji' });\n                card.createEl('div', { text: day.weatherText, cls: 'f-desc' });\n                card.createEl('div', {\n                    cls: 'f-temp',\n                    attr: { innerHTML: (day.tempMax || '--') + '°<span class=\"night\"> / ' + (day.tempMin || '--') + '°</span>' }\n                });\n            });\n        }\n\n    } catch (e) {\n        wrap.empty();\n        var err = wrap.createDiv({ cls: 'weather-error' });\n        err.createEl('div', { text: '❌', cls: 'big-icon' });\n        err.createEl('div', { text: e.message || t('mod.weather.loadFailed'), cls: 'msg' });\n        var retry = err.createEl('div', { text: t('mod.weather.retry'), cls: 'retry' });\n        retry.addEventListener('click', function() { render(content); });\n    }\n}\n\n// ══════════════════════════════════════\n//  设置面板（多平台支持）\n// ══════════════════════════════════════\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    var Setting = require('obsidian').Setting;\n\n    containerEl.createEl('h3', { text: t('mod.weather.settings.title') });\n\n    // ─── 平台选择下拉框 ───\n    new Setting(containerEl)\n        .setName(t('mod.weather.settings.provider'))\n        .setDesc(t('mod.weather.settings.providerDesc'))\n        .addDropdown(dd => {\n            dd.addOption('amap', '🗺️ ' + (t('mod.weather.provider.amap') || '高德地图 Amap (中国)'))\n              .addOption('open-meteo', '🌐 ' + (t('mod.weather.provider.openmeteo') || 'Open-Meteo (全球·推荐)'))\n              .addOption('wttr', '🌦️ ' + (t('mod.weather.provider.wttr') || 'wttr.in (极简)'))\n              .addOption('openweathermap', '☁️ ' + (t('mod.weather.provider.owm') || 'OpenWeatherMap (国际)'))\n              .addOption('custom', '⚙️ ' + (t('mod.weather.provider.custom') || '自定义 URL'))\n              .setValue(settings.provider || 'amap')\n              .onChange(async (v) => {\n                  settings.provider = v;\n                  await saveCallback();\n                  // 重新渲染设置面板以更新条件字段\n                  containerEl.empty();\n                  renderSettings(containerEl, plugin, saveCallback);\n              });\n        });\n\n    // ─── 城市名称（所有平台通用）───\n    new Setting(containerEl)\n        .setName(t('mod.weather.settings.city'))\n        .setDesc(t('mod.weather.settings.cityDesc'))\n        .addText(t => t.setPlaceholder(_isZh() ? '北京' : 'London')\n            .setValue(settings.city || '北京')\n            .onChange(async (v) => {\n                settings.city = v.trim();\n                await saveCallback();\n            }));\n\n    var provider = settings.provider || 'amap';\n\n    // ─── 条件字段：仅高德和 OWM 需要 API Key ───\n    if (provider === 'amap' || provider === 'openweathermap') {\n        if (provider === 'amap') {\n            new Setting(containerEl)\n                .setName(t('mod.weather.settings.apiKey'))\n                .setDesc(t('mod.weather.settings.apiKeyDesc'))\n                .addText(t => t.setPlaceholder(_isZh() ? '请输入 API Key' : 'Enter API Key')\n                    .setValue(settings.apiKey || '')\n                    .onChange(async (v) => {\n                        settings.apiKey = v.trim();\n                        await saveCallback();\n                    })\n                );\n            t.inputEl.style.width = '100%';\n        } else {\n            new Setting(containerEl)\n                .setName(t('mod.weather.settings.owmApiKey'))\n                .setDesc(t('mod.weather.settings.owmApiKeyDesc'))\n                .addText(t => t.setPlaceholder(_isZh() ? '请输入 OWM API Key' : 'Enter OWM API Key')\n                    .setValue(settings.apiKey || '')\n                    .onChange(async (v) => {\n                        settings.apiKey = v.trim();\n                        await saveCallback();\n                    }));\n            t.inputEl.style.width = '100%';\n        }\n    }\n\n    // ─── 条件字段：自定义 URL 模板 ───\n    if (provider === 'custom') {\n        new Setting(containerEl)\n            .setName(t('mod.weather.settings.customUrl'))\n            .setDesc(t('mod.weather.settings.customUrlDesc'))\n            .addTextArea(t => {\n                t.setPlaceholder('http://your-weather-api.com/?q={city}&format=json')\n                    .setValue(settings.customApiUrl || '')\n                    .onChange(async (v) => {\n                        settings.customApiUrl = v.trim();\n                        await saveCallback();\n                    });\n                t.inputEl.style.width = '100%';\n                t.inputEl.style.minHeight = '80px';\n                t.inputEl.style.fontFamily = 'monospace';\n                t.inputEl.style.fontSize = '12px';\n            });\n    }\n\n    // ─── 各平台说明信息 ───\n    var helpDiv = containerEl.createDiv({ style: 'margin-top: 12px;padding: 10px;background: var(--background-modifier-form-field);border-radius: 8px;font-size: 11px;line-height: 1.6;color: var(--text-muted);' });\n    var helps = {\n        'amap': t('mod.weather.help.amap') || '高德地图天气服务，覆盖中国全境，数据精准。需要注册并申请免费的 Web 服务 API Key。',\n        'open-meteo': t('mod.weather.help.openmeteo') || '开源免费天气 API，整合了19国国家气象局模型（含中国），全球覆盖。无需 API Key，无调用限制。',\n        'wttr': t('mod.weather.help.wttr') || '极简天气查询服务，直接用城市名即可。无需注册无需 Key。数据来源为 WorldWeatherOnline。',\n        'openweathermap': t('mod.weather.help.owm') || '国际主流天气 API，支持多语言。需免费注册申请 API Key（1000次/天）。',\n        'custom': t('mod.weather.help.custom') || '使用你自己的天气 API 接口。URL 中使用 {city} 作为城市名占位符。返回 JSON 数据将被自动解析。'\n    };\n    helpDiv.setText(helps[provider] || '');\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "web-preview": "/**\n * 网页预览模块 — V17 改造\n * 从 iframe 改为 Electron webview，支持登录态持久化\n * viewport + wrapper + webview 三层架构（与 web-video 一致）\n */\nconst id = 'web-preview';\nconst title = t('mod.webPreview');\nconst icon = '🌐';\n\nconst defaultSettings = {\n    url: 'https://www.baidu.com',\n    zoom: 1,\n    posX: 0,\n    posY: 0\n};\n\nconst styles = `\n.web-preview-toolbar {\n    display: flex;\n    align-items: center;\n    gap: 6px;\n    padding: 8px 12px;\n    border-bottom: 1px solid var(--background-modifier-border);\n    background: var(--background-secondary-alt);\n    flex-wrap: nowrap;\n    overflow: hidden;\n    flex-shrink: 0;\n}\n.web-preview-url {\n    flex: 1;\n    min-width: 80px;\n    padding: 6px 8px;\n    border: 1px solid var(--background-modifier-border);\n    border-radius: 4px;\n    background: var(--background-secondary);\n    color: var(--text-normal);\n    font-size: 12px;\n}\n.web-preview-url:focus {\n    outline: none;\n    border-color: var(--v6-primary);\n}\n.web-preview-btn {\n    padding: 6px 8px;\n    border: 1px solid var(--background-modifier-border);\n    background: var(--background-secondary);\n    border-radius: 4px;\n    cursor: pointer;\n    font-size: 12px;\n    color: var(--text-normal);\n    transition: all 0.2s ease;\n    flex-shrink: 0;\n}\n.web-preview-btn:hover {\n    background: var(--background-modifier-hover);\n    border-color: var(--v6-primary);\n}\n.web-preview-zoom {\n    font-size: 11px;\n    color: var(--text-muted);\n    min-width: 35px;\n    text-align: center;\n    flex-shrink: 0;\n}\n.web-preview-pos-input {\n    width: 45px;\n    padding: 4px 6px;\n    border: 1px solid var(--background-modifier-border);\n    border-radius: 4px;\n    background: var(--background-secondary);\n    color: var(--text-normal);\n    font-size: 11px;\n    text-align: center;\n    flex-shrink: 0;\n}\n.web-preview-pos-input:focus {\n    outline: none;\n    border-color: var(--v6-primary);\n}\n.web-preview-viewport {\n    flex: 1;\n    overflow: hidden;\n    position: relative;\n    background: var(--background-secondary);\n    min-height: 0;\n}\n.web-preview-wrapper {\n    position: absolute;\n    top: 0;\n    left: 0;\n    transform-origin: top left;\n    transition: transform 0.3s ease;\n}\n.web-preview-webview {\n    width: 100%;\n    height: 100%;\n    border: none;\n}\n`;\n\nasync function render(container) {\n    container.empty();\n\n    // 初始化\n    if (!settings.zoom) settings.zoom = 1;\n    if (settings.posY === undefined || settings.posY === null) settings.posY = 0;\n    if (settings.posX === undefined || settings.posX === null) settings.posX = 0;\n\n    let zoom = Number(settings.zoom);\n    if (!isFinite(zoom) || zoom <= 0) zoom = 1;\n\n    container.style.display = 'flex';\n    container.style.flexDirection = 'column';\n    container.style.height = '100%';\n\n    // ── 工具栏 ──\n    const toolbar = container.createDiv({ cls: 'web-preview-toolbar' });\n\n    const urlBar = toolbar.createEl('input', {\n        cls: 'web-preview-url',\n        attr: { type: 'text', value: settings.url, placeholder: t('mod.webPreview.placeholder') }\n    });\n\n    const zoomOutBtn = toolbar.createEl('button', {\n        cls: 'web-preview-btn', text: t('mod.webPreview.zoomOut'), attr: { title: t('mod.webPreview.zoomOutTitle') }\n    });\n    const zoomDisplay = toolbar.createEl('span', {\n        cls: 'web-preview-zoom', text: Math.round(zoom * 100) + '%'\n    });\n    const zoomInBtn = toolbar.createEl('button', {\n        cls: 'web-preview-btn', text: t('mod.webPreview.zoomIn'), attr: { title: t('mod.webPreview.zoomInTitle') }\n    });\n\n    const posYInput = toolbar.createEl('input', {\n        cls: 'web-preview-pos-input',\n        attr: { type: 'number', value: settings.posY, title: t('mod.webPreview.offsetYTitle') }\n    });\n    const posXInput = toolbar.createEl('input', {\n        cls: 'web-preview-pos-input',\n        attr: { type: 'number', value: settings.posX, title: t('mod.webPreview.offsetXTitle') }\n    });\n\n    const refreshBtn = toolbar.createEl('button', {\n        cls: 'web-preview-btn', text: t('mod.webPreview.refresh'), attr: { title: t('mod.webPreview.refreshTitle') }\n    });\n\n    // ── 视口 ──\n    const viewport = container.createDiv({ cls: 'web-preview-viewport' });\n\n    // ── webview 包装器 ──\n    const webviewWrapper = viewport.createDiv({ cls: 'web-preview-wrapper' });\n\n    // ── Electron webview（与 web-video 一致，支持登录态） ──\n    const webview = document.createElement('webview');\n    webview.className = 'web-preview-webview';\n    webview.setAttribute('src', settings.url);\n    // persist: 前缀使 Cookie 持久化，重启 Obsidian 后登录态不丢失\n    webview.setAttribute('partition', 'persist:webpreview-' + (_moduleId || id));\n    webview.setAttribute('preload', '');\n    webview.setAttribute('allowpopups', '');\n    webview.setAttribute('nodeintegration', 'false');\n    webview.setAttribute('webpreferences', 'contextIsolation=true, sandbox=true');\n\n    webviewWrapper.appendChild(webview);\n\n    // ── 缩放和位置 ──\n    const applyTransform = () => {\n        const scale = zoom;\n        const translateX = -settings.posX;\n        const translateY = -settings.posY;\n        webviewWrapper.style.transform = `translate(${translateX}px, ${translateY}px) scale(${scale})`;\n        const containerWidth = viewport.offsetWidth;\n        webviewWrapper.style.width = `${(containerWidth * 2) / scale}px`;\n        webviewWrapper.style.height = `${(containerWidth * 2) / scale}px`;\n    };\n\n    applyTransform();\n\n    // ── 缩放 ──\n    const updateZoom = () => {\n        zoom = Math.max(0.1, Math.min(2, zoom));\n        settings.zoom = zoom;\n        zoomDisplay.textContent = Math.round(zoom * 100) + '%';\n        applyTransform();\n        saveCallback();\n    };\n\n    zoomOutBtn.addEventListener('click', () => { zoom -= 0.1; updateZoom(); });\n    zoomInBtn.addEventListener('click', () => { zoom += 0.1; updateZoom(); });\n\n    // ── 刷新 ──\n    refreshBtn.addEventListener('click', () => {\n        settings.url = urlBar.value;\n        saveCallback();\n        webview.src = urlBar.value;\n    });\n\n    urlBar.addEventListener('keypress', (e) => {\n        if (e.key === 'Enter') {\n            settings.url = urlBar.value;\n            saveCallback();\n            webview.src = urlBar.value;\n        }\n    });\n\n    // ── 位置更新 ──\n    const updatePosition = () => {\n        settings.posX = parseInt(posXInput.value) || 0;\n        settings.posY = parseInt(posYInput.value) || 0;\n        applyTransform();\n        saveCallback();\n    };\n\n    posXInput.addEventListener('change', updatePosition);\n    posYInput.addEventListener('change', updatePosition);\n\n    // ── 注入 CSS 屏蔽广告 ──\n    webview.addEventListener('dom-ready', () => {\n        webview.insertCSS(`\n            .ad, .ads, .advertisement, .popup, .modal-overlay { display: none !important; }\n        `).catch(() => {});\n    });\n\n    // ── 新窗口在内部打开（登录跳转等） ──\n    webview.addEventListener('new-window', (e) => {\n        webview.src = e.url;\n    });\n}\n\nfunction renderSettings(wrapper, plugin, saveCallback) {\n    new Setting(wrapper)\n        .setName(t('mod.webPreview.settings.url'))\n        .setDesc(t('mod.webPreview.settings.urlDesc'))\n        .addText(t => {\n            t.setPlaceholder('https://example.com')\n                .setValue(settings.url || '')\n                .onChange(async (v) => { settings.url = v; await saveCallback(); });\n        });\n\n    new Setting(wrapper)\n        .setName(t('mod.webPreview.settings.zoom'))\n        .setDesc(t('mod.webPreview.settings.zoomDesc'))\n        .addSlider(s => {\n            s.setLimits(0.1, 2, 0.1)\n                .setValue(Number(settings.zoom) || 1)\n                .setDynamicTooltip()\n                .onChange(async (v) => { settings.zoom = v; await saveCallback(); });\n        });\n\n    new Setting(wrapper)\n        .setName(t('mod.webPreview.settings.posX'))\n        .setDesc(t('mod.webPreview.settings.posXDesc'))\n        .addText(t => {\n            t.setValue(String(settings.posX || 0))\n                .onChange(async (v) => { settings.posX = parseInt(v) || 0; await saveCallback(); });\n        });\n\n    new Setting(wrapper)\n        .setName(t('mod.webPreview.settings.posY'))\n        .setDesc(t('mod.webPreview.settings.posYDesc'))\n        .addText(t => {\n            t.setValue(String(settings.posY || 0))\n                .onChange(async (v) => { settings.posY = parseInt(v) || 0; await saveCallback(); });\n        });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "web-video": "/**\n * 网页视频模块 — 从 V13 原样迁移\n * viewport + wrapper + webview 三层架构\n * webview 始终 allowpopups，new-window 直接在内部加载\n */\nconst id = 'web-video';\nconst title = t('mod.webVideo');\nconst icon = '📺';\n\nconst defaultSettings = {\n    url: 'https://www.bilibili.com',\n    zoom: 1,\n    posX: 0,\n    posY: 0\n};\n\nconst styles = `\n.web-video-toolbar {\n    display: flex;\n    align-items: center;\n    gap: 6px;\n    padding: 8px 12px;\n    border-bottom: 1px solid var(--background-modifier-border);\n    background: var(--background-secondary-alt);\n    flex-wrap: nowrap;\n    overflow: hidden;\n    flex-shrink: 0;\n}\n.web-video-url {\n    flex: 1;\n    min-width: 80px;\n    padding: 6px 8px;\n    border: 1px solid var(--background-modifier-border);\n    border-radius: 4px;\n    background: var(--background-secondary);\n    color: var(--text-normal);\n    font-size: 12px;\n}\n.web-video-url:focus {\n    outline: none;\n    border-color: var(--v6-primary);\n}\n.web-video-btn {\n    padding: 6px 8px;\n    border: 1px solid var(--background-modifier-border);\n    background: var(--background-secondary);\n    border-radius: 4px;\n    cursor: pointer;\n    font-size: 12px;\n    color: var(--text-normal);\n    transition: all 0.2s ease;\n    flex-shrink: 0;\n}\n.web-video-btn:hover {\n    background: var(--background-modifier-hover);\n    border-color: var(--v6-primary);\n}\n.web-video-zoom {\n    font-size: 11px;\n    color: var(--text-muted);\n    min-width: 35px;\n    text-align: center;\n    flex-shrink: 0;\n}\n.web-video-pos-input {\n    width: 45px;\n    padding: 4px 6px;\n    border: 1px solid var(--background-modifier-border);\n    border-radius: 4px;\n    background: var(--background-secondary);\n    color: var(--text-normal);\n    font-size: 11px;\n    text-align: center;\n    flex-shrink: 0;\n}\n.web-video-pos-input:focus {\n    outline: none;\n    border-color: var(--v6-primary);\n}\n.web-video-viewport {\n    flex: 1;\n    overflow: hidden;\n    position: relative;\n    background: var(--background-secondary);\n    min-height: 0;\n}\n.web-video-wrapper {\n    position: absolute;\n    top: 0;\n    left: 0;\n    transform-origin: top left;\n    transition: transform 0.3s ease;\n}\n.web-video-webview {\n    width: 100%;\n    height: 100%;\n    border: none;\n}\n`;\n\nasync function render(container) {\n    container.empty();\n\n    // V13 原始初始化\n    if (!settings.zoom) settings.zoom = 1;\n    if (settings.posY === undefined || settings.posY === null) settings.posY = 0;\n    if (settings.posX === undefined || settings.posX === null) settings.posX = 0;\n\n    let zoom = Number(settings.zoom);\n    if (!isFinite(zoom) || zoom <= 0) zoom = 1;\n\n    container.style.display = 'flex';\n    container.style.flexDirection = 'column';\n    container.style.height = '100%';\n\n    // ── 工具栏（V13 原始结构） ──\n    const toolbar = container.createDiv({ cls: 'web-video-toolbar' });\n\n    const urlBar = toolbar.createEl('input', {\n        cls: 'web-video-url',\n        attr: { type: 'text', value: settings.url, placeholder: t('mod.webPreview.placeholder') }\n    });\n\n    const zoomOutBtn = toolbar.createEl('button', {\n        cls: 'web-video-btn', text: t('mod.webPreview.zoomOut'), attr: { title: t('mod.webPreview.zoomOutTitle') }\n    });\n    const zoomDisplay = toolbar.createEl('span', {\n        cls: 'web-video-zoom', text: Math.round(zoom * 100) + '%'\n    });\n    const zoomInBtn = toolbar.createEl('button', {\n        cls: 'web-video-btn', text: t('mod.webPreview.zoomIn'), attr: { title: t('mod.webPreview.zoomInTitle') }\n    });\n\n    const posYInput = toolbar.createEl('input', {\n        cls: 'web-video-pos-input',\n        attr: { type: 'number', value: settings.posY, title: t('mod.webPreview.offsetYTitle') }\n    });\n    const posXInput = toolbar.createEl('input', {\n        cls: 'web-video-pos-input',\n        attr: { type: 'number', value: settings.posX, title: t('mod.webPreview.offsetXTitle') }\n    });\n\n    const refreshBtn = toolbar.createEl('button', {\n        cls: 'web-video-btn', text: t('mod.webPreview.refresh'), attr: { title: t('mod.webPreview.refreshTitle') }\n    });\n\n    // ── 视口（V13: position relative + overflow hidden） ──\n    const viewport = container.createDiv({ cls: 'web-video-viewport' });\n\n    // ── webview 包装器（V13: position absolute，用于 transform） ──\n    const webviewWrapper = viewport.createDiv({ cls: 'web-video-wrapper' });\n\n    // ── Electron webview（V13 原始属性） ──\n    const webview = document.createElement('webview');\n    webview.className = 'web-video-webview';\n    webview.setAttribute('src', settings.url);\n    webview.setAttribute('partition', 'persist:webvideo-' + (_moduleId || id));\n    webview.setAttribute('preload', '');\n    webview.setAttribute('allowpopups', '');\n\n    webview.setAttribute('nodeintegration', 'false');\n    webview.setAttribute('webpreferences', 'contextIsolation=true, sandbox=true');\n\n    webviewWrapper.appendChild(webview);\n\n    // ── 缩放和位置（V13 方案） ──\n    const applyTransform = () => {\n        const scale = zoom;\n        const translateX = -settings.posX;\n        const translateY = -settings.posY;\n        webviewWrapper.style.transform = `translate(${translateX}px, ${translateY}px) scale(${scale})`;\n        const containerWidth = viewport.offsetWidth;\n        webviewWrapper.style.width = `${(containerWidth * 2) / scale}px`;\n        webviewWrapper.style.height = `${(containerWidth * 2) / scale}px`;\n    };\n\n    applyTransform();\n\n    // ── 缩放 ──\n    const updateZoom = () => {\n        zoom = Math.max(0.1, Math.min(2, zoom));\n        settings.zoom = zoom;\n        zoomDisplay.textContent = Math.round(zoom * 100) + '%';\n        applyTransform();\n        saveCallback();\n    };\n\n    zoomOutBtn.addEventListener('click', () => { zoom -= 0.1; updateZoom(); });\n    zoomInBtn.addEventListener('click', () => { zoom += 0.1; updateZoom(); });\n\n    // ── 刷新 ──\n    refreshBtn.addEventListener('click', () => {\n        settings.url = urlBar.value;\n        saveCallback();\n        webview.src = urlBar.value;\n    });\n\n    urlBar.addEventListener('keypress', (e) => {\n        if (e.key === 'Enter') {\n            settings.url = urlBar.value;\n            saveCallback();\n            webview.src = urlBar.value;\n        }\n    });\n\n    // ── 位置更新 ──\n    const updatePosition = () => {\n        settings.posX = parseInt(posXInput.value) || 0;\n        settings.posY = parseInt(posYInput.value) || 0;\n        applyTransform();\n        saveCallback();\n    };\n\n    posXInput.addEventListener('change', updatePosition);\n    posYInput.addEventListener('change', updatePosition);\n\n    // ── 注入 CSS 屏蔽广告（V13 原始逻辑） ──\n    webview.addEventListener('dom-ready', () => {\n        webview.insertCSS(`\n            .ad, .ads, .advertisement, .popup, .modal-overlay { display: none !important; }\n        `).catch(() => {});\n    });\n\n    // ── 新窗口在内部打开（V13 原始逻辑，直接 webview.src = url） ──\n    webview.addEventListener('new-window', (e) => {\n        webview.src = e.url;\n    });\n}\n\nfunction renderSettings(wrapper, plugin, saveCallback) {\n    new Setting(wrapper)\n        .setName(t('mod.webVideo.settings.url'))\n        .setDesc(t('mod.webVideo.settings.urlDesc'))\n        .addText(t => {\n            t.setPlaceholder('https://www.bilibili.com')\n                .setValue(settings.url || '')\n                .onChange(async (v) => { settings.url = v; await saveCallback(); });\n        });\n\n    new Setting(wrapper)\n        .setName(t('mod.webPreview.settings.zoom'))\n        .setDesc(t('mod.webPreview.settings.zoomDesc'))\n        .addSlider(s => {\n            s.setLimits(0.1, 2, 0.1)\n                .setValue(Number(settings.zoom) || 1)\n                .setDynamicTooltip()\n                .onChange(async (v) => { settings.zoom = v; await saveCallback(); });\n        });\n\n    new Setting(wrapper)\n        .setName(t('mod.webPreview.settings.posX'))\n        .setDesc(t('mod.webPreview.settings.posXDesc'))\n        .addText(t => {\n            t.setValue(String(settings.posX || 0))\n                .onChange(async (v) => { settings.posX = parseInt(v) || 0; await saveCallback(); });\n        });\n\n    new Setting(wrapper)\n        .setName(t('mod.webPreview.settings.posY'))\n        .setDesc(t('mod.webPreview.settings.posYDesc'))\n        .addText(t => {\n            t.setValue(String(settings.posY || 0))\n                .onChange(async (v) => { settings.posY = parseInt(v) || 0; await saveCallback(); });\n        });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n",
  "xhs-importer": "/**\n * 小红书导入模块 (v3)\n * 基于原插件小红书 (D:\\Obsidian仓库\\.obsidian\\plugins\\小红书\\main.js) 的已验证逻辑逐一对照重写\n * \n * v3 修复：\n * 1. requestUrl 不再带自定义 headers（原插件：requestUrl({url})，带header会触发部分页面的反爬）\n * 2. 新增视频笔记封面图提取（note.video.image / video.cover）\n * 3. 新增 extractVideoUrl() 提取视频链接\n * 4. extractImages 多重降级：imageList → video cover → DOM img 标签 → regex JSON\n * 5. 诊断日志增强：失败时输出 HTML 片段帮助定位\n */\nconst id = 'xhs-importer';\nconst title = t('mod.xhsImporter');\nconst icon = '📕';\n\nconst defaultSettings = {\n    noteFolder: '收件箱',\n    imageFolder: '附件/XHS',\n    downloadMedia: true,\n    customTags: '小红书'\n};\n\nconst styles = `\n.xhs-importer-wrap { padding: 12px; }\n.xhs-textarea {\n    width: 100%;\n    min-height: 80px;\n    resize: vertical;\n    background: var(--background-modifier-form-field);\n    color: var(--text-normal);\n    border: 1px solid var(--background-modifier-border);\n    border-radius: 6px;\n    padding: 8px 10px;\n    font-size: 12px;\n    font-family: var(--font-interface);\n    box-sizing: border-box;\n}\n.xhs-textarea::placeholder { color: var(--text-muted); }\n.xhs-actions {\n    display: flex;\n    gap: 8px;\n    margin-top: 10px;\n    align-items: center;\n}\n.xhs-btn {\n    padding: 6px 16px;\n    border-radius: 6px;\n    border: none;\n    cursor: pointer;\n    font-size: 12px;\n    background: var(--interactive-accent);\n    color: var(--text-on-accent);\n    transition: opacity 0.15s;\n}\n.xhs-btn:hover { opacity: 0.85; }\n.xhs-btn:disabled { opacity: 0.4; cursor: not-allowed; }\n.xhs-status {\n    font-size: 11px;\n    color: var(--text-muted);\n    flex: 1;\n    text-align: right;\n}\n.xhs-error { color: var(--text-error) !important; }\n.xhs-success { color: var(--text-success) !important; }\n.xhs-preview {\n    margin-top: 10px;\n    padding: 8px;\n    background: var(--background-modifier-form-field);\n    border-radius: 6px;\n    font-size: 12px;\n    color: var(--text-muted);\n    max-height: 120px;\n    overflow-y: auto;\n}\n.xhs-preview-title { font-weight: 600; color: var(--text-normal); margin-bottom: 4px; }\n.xhs-note { margin-top: 8px; padding: 6px 8px; background: var(--background-primary-alt); border-radius: 4px; font-size: 11px; color: var(--text-muted); }\n`;\n\n// ===== 工具函数 =====\n\n/** 从分享文本提取 URL */\nfunction extractURL(text) {\n    var patterns = [\n        /https?:\\/\\/www\\.xiaohongshu\\.com\\/[^\\s]+/,\n        /https?:\\/\\/xhslink\\.com\\/[^\\s]+/,\n        /https?:\\/\\/[^\\s]*xiaohongshu[^\\s]*/i\n    ];\n    for (var i = 0; i < patterns.length; i++) {\n        var m = text.match(patterns[i]);\n        if (m) return m[0].replace(/[.,;!?）)>]+$/, '');\n    }\n    return null;\n}\n\n/** 安全化文件名 */\nfunction sanitizeFilename(name) {\n    return (name || '笔记').replace(/[\\\\/:*?\"<>|]/g, '_').substring(0, 80);\n}\n\n/** 从 URL 提取文件扩展名 */\nfunction getExtensionFromUrl(url, fallback) {\n    try {\n        var u = new URL(url);\n        var match = (u.pathname || '').match(/\\.([a-zA-Z0-9]+)(\\?|$)/);\n        return match ? match[1].toLowerCase() : fallback;\n    } catch (e) {\n        return fallback;\n    }\n}\n\n/** 确保文件夹存在 */\nasync function ensureFolder(folderPath) {\n    var p = (folderPath || '').replace(/\\\\/g, '/').replace(/\\/+$/, '');\n    if (!p) return;\n    try {\n        // app.vault.adapter 不需要 normalizePath，但 createFolder 需要\n        if (!(await app.vault.adapter.exists(p))) {\n            await app.vault.createFolder(p);\n        }\n    } catch (e) {\n        // 忽略文件夹已存在的错误\n        if (e.message && e.message.indexOf('already exists') < 0) {\n            console.warn('[xhs-importer] ensureFolder error:', e.message);\n        }\n    }\n}\n\n/** 唯一文件路径 */\nasync function getUniqueFilePath(folderPath, baseName, ext) {\n    var dir = (folderPath || '').replace(/\\\\/g, '/').replace(/\\/+$/, '');\n    var candidate = dir ? dir + '/' + baseName + '.' + ext : baseName + '.' + ext;\n    var counter = 1;\n    while (await app.vault.adapter.exists(candidate)) {\n        candidate = (dir ? dir + '/' : '') + baseName + '-' + counter + '.' + ext;\n        counter++;\n    }\n    return candidate;\n}\n\n/** 唯一图片路径 */\nasync function getUniqueImagePath(folderPath, baseName, ext) {\n    var dir = (folderPath || '').replace(/\\\\/g, '/').replace(/\\/+$/, '');\n    var candidate = dir ? dir + '/' + baseName + '.' + ext : baseName + '.' + ext;\n    var counter = 1;\n    while (await app.vault.adapter.exists(candidate)) {\n        candidate = (dir ? dir + '/' : '') + baseName + '-' + counter + '.' + ext;\n        counter++;\n    }\n    return candidate;\n}\n\n// ===== __INITIAL_STATE__ 解析（对齐原插件 小红书/main.js） =====\n\n/**\n * 解析小红书页面 HTML 中的 window.__INITIAL_STATE__ JSON\n * 对齐原插件：用 /s (dotAll) 标志让 . 匹配换行符\n */\nfunction parseInitialState(html) {\n    // ★ 对齐原插件：/window\\.__INITIAL_STATE__=(.*?)<\\/script>/s\n    var match = html.match(/window\\.__INITIAL_STATE__\\s*=\\s*({[\\s\\S]*?});?\\s*<\\/script>/);\n    if (!match) {\n        // fallback：不加 { } 限制（兼容非 JSON 格式的变体）\n        match = html.match(/window\\.__INITIAL_STATE__\\s*=\\s*([\\s\\S]*?)<\\/script>/);\n        if (!match) return null;\n    }\n    try {\n        var jsonStr = match[1].trim();\n        // 去除末尾可能的分号\n        if (jsonStr.charAt(jsonStr.length - 1) === ';') jsonStr = jsonStr.slice(0, -1);\n        // ★ 对齐原插件：.replace(/undefined/g, \"null\")\n        jsonStr = jsonStr.replace(/undefined/g, 'null');\n        return JSON.parse(jsonStr);\n    } catch (e) {\n        console.error('[xhs-importer] JSON.parse 失败:', e.message.substring(0, 100));\n        return null;\n    }\n}\n\n/**\n * 从 __INITIAL_STATE__ 提取笔记详情\n * 结构: state.note.noteDetailMap[noteId].note\n */\nfunction getNoteDetail(html) {\n    var state = parseInitialState(html);\n    if (!state || !state.note || !state.note.noteDetailMap) return null;\n    var noteIds = Object.keys(state.note.noteDetailMap);\n    if (noteIds.length === 0) return null;\n    var noteId = noteIds[0];\n    return state.note.noteDetailMap[noteId].note || null;\n}\n\n/**\n * 提取标题（对齐原插件：仅用 <title> 标签，降级到 note.title）\n */\nfunction extractTitle(html) {\n    // ★ 对齐原插件：/<title>(.*?)<\\/title>/\n    var match = html.match(/<title>([\\s\\S]*?)<\\/title>/i);\n    if (match) return match[1].trim().replace(' - 小红书', '').trim();\n    // 降级：__INITIAL_STATE__\n    var note = getNoteDetail(html);\n    if (note && note.title) return note.title;\n    return '小红书笔记';\n}\n\n/**\n * 提取正文内容（对齐原插件：先 DOM 提取 detail-desc，再降级到 note.desc）\n */\nfunction extractContent(html) {\n    // ★ 对齐原插件：先匹配 DOM <div id=\"detail-desc\" class=\"desc\">\n    var domMatch = html.match(/<div[^>]*id=\"detail-desc\"[^>]*class=\"desc\"[^>]*>([\\s\\S]*?)<\\/div>/i)\n                || html.match(/<div[^>]*class=\"desc\"[^>]*id=\"detail-desc\"[^>]*>([\\s\\S]*?)<\\/div>/i);\n    if (domMatch) {\n        var text = domMatch[1]\n            .replace(/<[^>]+>/g, '')\n            .replace(/\\[话题\\]/g, '')\n            .replace(/\\[[^\\]]+\\]/g, '')\n            .trim();\n        if (text) return text;\n    }\n    // ★ 降级：__INITIAL_STATE__ note.desc（对齐原插件）\n    var note = getNoteDetail(html);\n    if (note && note.desc) {\n        return note.desc.replace(/\\[话题\\]/g, '').replace(/\\[[^\\]]+\\]/g, '').trim();\n    }\n    return '';\n}\n\n/**\n * 提取图片 URL 列表（对齐原插件：imageList[].urlDefault）\n * v3：新增视频笔记封面图 + DOM降级 + regex JSON降级\n * \n * 降级链：\n *   1. note.imageList[].urlDefault (普通图文笔记)\n *   2. note.video.* (视频笔记封面)\n *   3. DOM <img> 标签 src (Electron渲染后页面)\n *   4. regex 从HTML中提取图片URL\n */\nfunction extractImages(html) {\n    var note = getNoteDetail(html);\n    \n    // Step 1: 普通图文笔记 imageList\n    if (note && note.imageList && note.imageList.length > 0) {\n        var urls = note.imageList\n            .map(function(img) { return img.urlDefault || img.url || ''; })\n            .filter(function(url) { return url && url.startsWith('http'); });\n        if (urls.length > 0) return urls;\n    }\n    \n    // Step 2: 视频笔记封面图（对齐原插件：视频笔记用第一张图当封面）\n    if (note && note.type === 'video' && note.video) {\n        var v = note.video;\n        var coverUrl = \n            (v.image && (v.image.urlDefault || v.image.url)) ||\n            (v.cover && (v.cover.urlDefault || v.cover.url)) ||\n            (v.media && v.media.image && (v.media.image.urlDefault || v.media.image.url)) ||\n            '';\n        // 补全协议头\n        if (coverUrl && coverUrl.indexOf('//') === 0) coverUrl = 'https:' + coverUrl;\n        if (coverUrl && coverUrl.startsWith('http')) {\n            console.log('[xhs-importer] 视频封面图:', coverUrl);\n            return [coverUrl];\n        }\n    }\n    \n    // Step 3: DOM 提取 — img 标签（小红书 SPA 页面渲染后）\n    // 图片通常在 .swiper-slide img 或 .note-image img 中\n    var imgMatches = [];\n    var imgRe = /<img[^>]+\\bsrc\\s*=\\s*[\"']([^\"']*\\.(?:jpg|jpeg|png|webp|gif)\\??[^\"']*)[\"']/gi;\n    var m;\n    while ((m = imgRe.exec(html)) !== null) {\n        var candidate = m[1];\n        if (candidate.indexOf('xhscdn') > -1 || candidate.indexOf('sns-webpic') > -1 || candidate.indexOf('ci.xiaohongshu') > -1) {\n            if (candidate.indexOf('//') === 0) candidate = 'https:' + candidate;\n            if (candidate.startsWith('http') && imgMatches.indexOf(candidate) < 0) {\n                imgMatches.push(candidate);\n            }\n        }\n    }\n    if (imgMatches.length > 0) {\n        console.log('[xhs-importer] DOM提取图片:', imgMatches.length, '张');\n        return imgMatches;\n    }\n    \n    // Step 4: regex 从 HTML JSON 数据中提取（最后降级）\n    var jsonImgRe = /\"urlDefault\"\\s*:\\s*\"([^\"]+)\"/g;\n    var jsonUrls = [];\n    while ((m = jsonImgRe.exec(html)) !== null) {\n        var u = m[1].replace(/\\\\\\//g, '/');\n        if (u.indexOf('//') === 0) u = 'https:' + u;\n        if (u.startsWith('http') && jsonUrls.indexOf(u) < 0) {\n            jsonUrls.push(u);\n        }\n    }\n    if (jsonUrls.length > 0) {\n        console.log('[xhs-importer] JSON正则提取图片:', jsonUrls.length, '张');\n        return jsonUrls;\n    }\n    \n    return [];\n}\n\n/**\n * 提取视频 URL（对齐原插件：note.video.media.stream.h264[0].masterUrl）\n */\nfunction extractVideoUrl(html) {\n    var note = getNoteDetail(html);\n    if (!note || note.type !== 'video') return '';\n    \n    var v = note.video;\n    if (!v) return '';\n    \n    // 主路径：media.stream.h264\n    if (v.media && v.media.stream && v.media.stream.h264 && v.media.stream.h264.length > 0) {\n        return v.media.stream.h264[0].masterUrl || '';\n    }\n    // 降级：直接 video url\n    return v.url || v.videoUrl || v.playUrl || '';\n}\n\n/**\n * 提取作者（对齐原插件：user.nickname）\n */\nfunction extractAuthor(html) {\n    var note = getNoteDetail(html);\n    if (note && note.user && note.user.nickname) return note.user.nickname;\n    return '';\n}\n\n/**\n * 提取 #话题标签\n */\nfunction extractTags(content) {\n    var tags = [];\n    var re = /#([^\\s#]+)/g;\n    var match;\n    while ((match = re.exec(content)) !== null) {\n        var tag = match[1].replace(/[\\[\\]]/g, '');\n        if (tag && tags.indexOf(tag) < 0) tags.push(tag);\n    }\n    return tags;\n}\n\n/**\n * 判断是否为视频笔记\n */\nfunction isVideoNote(html) {\n    var note = getNoteDetail(html);\n    return !!(note && note.type === 'video');\n}\n\n/**\n * 检查 HTML 是否包含 __INITIAL_STATE__（用于判断请求是否成功）\n */\nfunction looksLikeXHS(html) {\n    return html && html.indexOf('__INITIAL_STATE__') > -1;\n}\n\n// ===== 下载图片（原插件方式：fetch + blob + writeBinary）=====\n\nasync function downloadMediaFile(url, imageFolder, baseName, fallbackExt) {\n    try {\n        var response = await fetch(url);\n        if (!response.ok) {\n            throw new Error('HTTP ' + response.status);\n        }\n        var ext = getExtensionFromUrl(url, fallbackExt);\n        var targetPath = await getUniqueImagePath(imageFolder, baseName, ext);\n        var blob = await response.blob();\n        var bytes = await blob.arrayBuffer();\n        await app.vault.adapter.writeBinary(targetPath, bytes);\n        console.log('[xhs-importer] 图片下载成功: ' + targetPath);\n        return targetPath;\n    } catch (e) {\n        console.error('[xhs-importer] 下载图片失败: ' + url, e.message);\n        return null;\n    }\n}\n\n// ===== 构建 Frontmatter ====\n\nfunction buildFrontmatter(meta) {\n    var tags = [].concat(meta.tags || []);\n    // 添加用户自定义标签\n    var customTags = (settings.customTags || '小红书').split(/[,;\\s]+/).filter(Boolean);\n    customTags.forEach(function(t) { if (tags.indexOf(t) < 0) tags.push(t); });\n\n    var tagStr = tags.map(function(t) { return '\\n  - \"' + t.replace(/\"/g, '\\\\\"') + '\"'; }).join('');\n    var date = '';\n    try { date = moment().format('YYYY-MM-DD'); } catch(e) { date = new Date().toISOString().slice(0, 10); }\n\n    var lines = ['---'];\n    if (meta.title) lines.push('title: \"' + meta.title.replace(/\"/g, '\\\\\"') + '\"');\n    if (meta.author) lines.push('author: \"' + meta.author.replace(/\"/g, '\\\\\"') + '\"');\n    lines.push('source: \"' + (meta.url || '') + '\"');\n    lines.push('date: ' + date);\n    if (tags.length > 0) lines.push('tags:' + tagStr);\n    lines.push('---');\n    return lines.join('\\n');\n}\n\n// ===== 主渲染 =====\n\nasync function render(content) {\n    content.empty();\n    var wrap = content.createDiv({ cls: 'xhs-importer-wrap' });\n\n    var ta = wrap.createEl('textarea', {\n        cls: 'xhs-textarea',\n        attr: { placeholder: t('mod.xhsImporter.placeholder') + '\\n\\n' + t('mod.xhsImporter.autoFetch') }\n    });\n\n    var actions = wrap.createDiv({ cls: 'xhs-actions' });\n    var btn = actions.createEl('button', { cls: 'xhs-btn', text: t('mod.xhsImporter.btn.import') });\n    var status = actions.createEl('span', { cls: 'xhs-status', text: t('mod.xhsImporter.ready') });\n\n    var noteInfo = wrap.createDiv({ cls: 'xhs-note' });\n    noteInfo.style.display = 'none';\n\n    var previewArea = wrap.createDiv({ cls: 'xhs-preview' });\n    previewArea.style.display = 'none';\n\n    var working = false;\n\n    btn.addEventListener('click', async function() {\n        if (working) return;\n        var text = ta.value.trim();\n        if (!text) {\n            status.textContent = t('mod.xhsImporter.error.empty');\n            status.className = 'xhs-status xhs-error';\n            return;\n        }\n\n        working = true;\n        btn.disabled = true;\n        status.textContent = t('mod.xhsImporter.loading.parse');\n        status.className = 'xhs-status';\n\n        try {\n            var url = extractURL(text);\n            if (!url) {\n                status.textContent = t('mod.xhsImporter.error.noLink');\n                status.className = 'xhs-status xhs-error';\n                working = false;\n                btn.disabled = false;\n                return;\n            }\n\n            // 1. 获取页面 HTML（先 requestUrl，失败降级 fetch）\n            status.textContent = t('mod.xhsImporter.loading.fetch');\n            var html;\n            var fetchMethod = 'requestUrl';\n\n            // ★ 方案A：requestUrl（对齐原插件：不带任何自定义 headers）\n            // 原插件只用 requestUrl({url})，加 header 反而会触发部分页面的反爬\n            try {\n                var resp = await requestUrl({ url: url });\n                html = resp.text;\n            } catch (e) {\n                console.warn('[xhs-importer] requestUrl 失败:', e.message);\n                html = null;\n            }\n\n            // 如果 requestUrl 没拿到 __INITIAL_STATE__，方案B：fetch 降级（也不带自定义header）\n            if (!looksLikeXHS(html)) {\n                console.log('[xhs-importer] requestUrl 未获取到 __INITIAL_STATE__，尝试 fetch...');\n                try {\n                    var fetchResp = await fetch(url);\n                    if (fetchResp.ok) {\n                        html = await fetchResp.text();\n                        fetchMethod = 'fetch';\n                    }\n                } catch (e2) {\n                    console.warn('[xhs-importer] fetch 也失败:', e2.message);\n                }\n            }\n\n            // 两个方案都失败\n            if (!html || !looksLikeXHS(html)) {\n                console.error('[xhs-importer] 无法获取有效页面。HTML 长度:', html ? html.length : 0);\n                // 输出前 300 字符帮助调试\n                if (html) console.log('[xhs-importer] HTML 前300字:', html.substring(0, 300));\n                status.textContent = t('mod.xhsImporter.error.fetch');\n                status.className = 'xhs-status xhs-error';\n                working = false;\n                btn.disabled = false;\n                return;\n            }\n\n            console.log('[xhs-importer] 页面获取成功，方式:', fetchMethod, 'HTML 长度:', html.length);\n\n            // 2. 从 HTML 提取结构化数据（对齐原插件：先DOM 后 __INITIAL_STATE__）\n            status.textContent = t('mod.xhsImporter.loading.data');\n            var title = extractTitle(html);\n            var author = extractAuthor(html);\n            var contentText = extractContent(html);\n            var images = extractImages(html);\n            var tags = extractTags(contentText);\n            var isVideo = isVideoNote(html);\n            var videoUrl = isVideo ? extractVideoUrl(html) : '';\n\n            console.log('[xhs-importer] 标题:', title);\n            console.log('[xhs-importer] 作者:', author);\n            console.log('[xhs-importer] 图片数:', images.length);\n            console.log('[xhs-importer] 内容长度:', contentText.length);\n            console.log('[xhs-importer] 话题标签:', tags);\n            console.log('[xhs-importer] 是视频:', isVideo, videoUrl ? 'URL: ' + videoUrl.substring(0, 60) : '');\n\n            // 如果内容仍为空，尝试直接从 meta description 提取（最后的降级）\n            if (!contentText || contentText.length < 5) {\n                var metaMatch = html.match(/<meta[^>]+name=\"description\"[^>]+content=\"([^\"]*)\"/i)\n                             || html.match(/<meta[^>]+content=\"([^\"]*)\"[^>]+name=\"description\"/i);\n                if (metaMatch) {\n                    contentText = metaMatch[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '\"');\n                }\n            }\n\n            // 如果解析仍失败（无 content 且无图片），直接报错，输出诊断信息\n            if ((!contentText || contentText.length < 5) && (!images || images.length === 0)) {\n                console.error('[xhs-importer] 解析失败 - 无内容且无图片');\n                console.log('[xhs-importer] HTML前500字:', html ? html.substring(0, 500) : '(null)');\n                status.textContent = t('mod.xhsImporter.error.parse');\n                status.className = 'xhs-status xhs-error';\n                new Notice(t('mod.xhsImporter.error.cannotParse'));\n                working = false;\n                btn.disabled = false;\n                return;\n            }\n\n            // 仅无文字但有图片时，用标题代替\n            if (!contentText || contentText.length < 5) {\n                contentText = title || t('mod.xhsImporter.xhsNote');\n                if (images.length > 0) {\n                    contentText += '\\n\\n（' + (isVideo ? t('mod.xhsImporter.videoNote') : t('mod.xhsImporter.imageNote')) + t('mod.xhsImporter.mediaNoteSuffix') + images.length + t('mod.xhsImporter.imageCount') + '）';\n                }\n            }\n\n            // 3. 创建文件夹\n            var noteFolder = settings.noteFolder || '收件箱';\n            var imageFolder = settings.imageFolder || '附件/XHS';\n            noteFolder = noteFolder.replace(/\\\\/g, '/').replace(/\\/+$/, '');\n            imageFolder = imageFolder.replace(/\\\\/g, '/').replace(/\\/+$/, '');\n            await ensureFolder(noteFolder);\n            if (settings.downloadMedia !== false) {\n                await ensureFolder(imageFolder);\n            }\n\n            // 4. 下载图片（原插件方式：fetch + blob + writeBinary）\n            var localImagePaths = [];\n            if (settings.downloadMedia !== false && images.length > 0) {\n                status.textContent = t('mod.xhsImporter.loading.downloading') + images.length + t('mod.xhsImporter.loading.imageCount');\n                var sanitizedTitle = sanitizeFilename(title);\n                for (var i = 0; i < images.length; i++) {\n                    status.textContent = t('mod.xhsImporter.loading.downloadImg') + (i + 1) + '/' + images.length + '...';\n                    var imgPath = await downloadMediaFile(\n                        images[i],\n                        imageFolder,\n                        sanitizedTitle + '-' + i,\n                        'jpg'\n                    );\n                    if (imgPath) localImagePaths.push(imgPath);\n                }\n                if (localImagePaths.length > 0) {\n                    status.textContent = t('mod.xhsImporter.loading.downloaded') + localImagePaths.length + t('mod.xhsImporter.imageCount');\n                } else {\n                    status.textContent = t('mod.xhsImporter.loading.downloadFail');\n                }\n            }\n\n            // 5. 构建 Markdown（对齐原插件：视频笔记和普通笔记的格式不同）\n            var frontmatter = buildFrontmatter({\n                title: title,\n                author: author,\n                url: url,\n                tags: tags\n            });\n\n            var mdContent = frontmatter + '\\n\\n';\n            mdContent += '# ' + title + '\\n\\n';\n\n            if (isVideo) {\n                // 视频笔记格式（对齐原插件）\n                if (localImagePaths.length > 0) {\n                    // 封面图（可点击跳转原文）\n                    mdContent += '[![' + localImagePaths[0].split('/').pop() + '](' + localImagePaths[0] + ')](' + url + ')\\n\\n';\n                } else if (images.length > 0) {\n                    mdContent += '[![' + images[0].split('/').pop() + '](' + images[0] + ')](' + url + ')\\n\\n';\n                }\n                if (videoUrl) {\n                    mdContent += '[▶ 视频链接](' + videoUrl + ')\\n\\n';\n                }\n                // 清理正文中的 #话题 以免干扰可读性\n                var cleanedContent = contentText.replace(/#\\S+/g, '').trim();\n                mdContent += cleanedContent + '\\n\\n';\n            } else {\n                // 普通图文笔记格式\n                mdContent += contentText + '\\n\\n';\n\n                // 嵌入图片\n                if (localImagePaths.length > 0) {\n                    mdContent += localImagePaths.map(function(p) {\n                        return '![' + p.split('/').pop() + '](' + p + ')';\n                    }).join('\\n') + '\\n\\n';\n                } else if (images.length > 0) {\n                    // 降级：使用原始 URL\n                    mdContent += images.map(function(u) {\n                        return '![' + u.split('/').pop() + '](' + u + ')';\n                    }).join('\\n') + '\\n\\n';\n                }\n            }\n\n            mdContent += '> 来源: ' + url + '\\n';\n\n            // 6. 保存笔记文件\n            status.textContent = t('mod.xhsImporter.loading.data');\n            var filename = sanitizeFilename(title);\n            var notePath = await getUniqueFilePath(noteFolder, filename, 'md');\n            var createdFile = await app.vault.create(notePath, mdContent);\n\n            // 7. 打开文件\n            await app.workspace.getLeaf(true).openFile(createdFile);\n\n            // 成功反馈\n            var summary = t('mod.xhsImporter.success');\n            if (localImagePaths.length > 0) summary += t('mod.xhsImporter.imagesDownloaded') + localImagePaths.length + t('mod.xhsImporter.imageCount');\n            status.textContent = summary;\n            status.className = 'xhs-status xhs-success';\n            new Notice('📕 ' + summary);\n\n            // 显示预览\n            previewArea.style.display = 'block';\n            previewArea.empty();\n            previewArea.createEl('div', { cls: 'xhs-preview-title', text: title });\n            var infoLines = [t('mod.xhsImporter.preview.author') + (author || t('mod.xhsImporter.unknown'))];\n            infoLines.push(t('mod.xhsImporter.preview.images') + images.length + t('mod.xhsImporter.preview.local') + localImagePaths.length + ')');\n            if (tags.length > 0) infoLines.push(t('mod.xhsImporter.preview.tags') + tags.join(', '));\n            infoLines.push(t('mod.xhsImporter.preview.source') + url);\n            infoLines.forEach(function(line) {\n                previewArea.createEl('div', { text: line, attr: { style: 'font-size:11px;margin-top:4px;' } });\n            });\n            var openLink = previewArea.createEl('a', { text: t('mod.xhsImporter.openNote'), href: '#' });\n            openLink.style.cssText = 'display:inline-block;margin-top:8px;color:var(--interactive-accent);font-size:12px;';\n            openLink.addEventListener('click', function(e) {\n                e.preventDefault();\n                app.workspace.openLinkText(notePath, '', false);\n            });\n\n        } catch (e) {\n            console.error('[xhs-importer] Import error:', e);\n            status.textContent = t('mod.xhsImporter.error.import') + e.message;\n            status.className = 'xhs-status xhs-error';\n            new Notice('📕 ' + t('mod.xhsImporter.error.import') + e.message);\n        } finally {\n            working = false;\n            btn.disabled = false;\n        }\n    });\n}\n\n// ===== 设置面板 =====\n\nfunction renderSettings(containerEl, plugin, saveCallback) {\n    var Setting = require('obsidian').Setting;\n\n    containerEl.empty();\n    containerEl.createEl('h3', { text: t('mod.xhsImporter.settings.title') });\n    containerEl.createEl('p', {\n        text: t('mod.xhsImporter.settings.desc'),\n        attr: { style: 'color:var(--text-muted);font-size:12px;line-height:1.6;' }\n    });\n\n    new Setting(containerEl)\n        .setName(t('mod.xhsImporter.settings.noteFolder'))\n        .setDesc(t('mod.xhsImporter.settings.noteFolderDesc'))\n        .addText(function(t) { return t.setValue(settings.noteFolder || '收件箱').onChange(async function(v) {\n            settings.noteFolder = v;\n            await saveCallback();\n        }); });\n\n    new Setting(containerEl)\n        .setName(t('mod.xhsImporter.settings.imageFolder'))\n        .setDesc(t('mod.xhsImporter.settings.imageFolderDesc'))\n        .addText(function(t) { return t.setValue(settings.imageFolder || '附件/XHS').onChange(async function(v) {\n            settings.imageFolder = v;\n            await saveCallback();\n        }); });\n\n    new Setting(containerEl)\n        .setName(t('mod.xhsImporter.settings.customTags'))\n        .setDesc(t('mod.xhsImporter.settings.customTagsDesc'))\n        .addText(function(t) { return t.setValue(settings.customTags || '小红书').onChange(async function(v) {\n            settings.customTags = v;\n            await saveCallback();\n        }); });\n\n    new Setting(containerEl)\n        .setName(t('mod.xhsImporter.settings.downloadMedia'))\n        .setDesc(t('mod.xhsImporter.settings.downloadMediaDesc'))\n        .addToggle(function(t) { return t.setValue(settings.downloadMedia !== false).onChange(async function(v) {\n            settings.downloadMedia = v;\n            await saveCallback();\n        }); });\n}\n\nmodule.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };\n"
};

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

const DEFAULT_SETTINGS = {
    theme: 'ink',
    layout: {},
    modules: {
        weather: {
            enabled: true,
            provider: 'amap',
            city: '北京',
            apiKey: '',
            customApiUrl: ''
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
        // === 新接入模块 ===
        'autoplay-loop': {
            enabled: true,
            autoplayAudio: true,
            autoplayVideo: true,
            loopAudio: true,
            loopVideo: true,
            muteAutoplayedAudio: false,
            muteAutoplayedVideo: true,
            singlePlaybackAudio: true,
            singlePlaybackVideo: false,
            pauseOutOfViewAudio: true,
            pauseOutOfViewVideo: true
        },
        'code-editor': {
            enabled: true,
            fontSize: 14,
            tabSize: 4,
            theme: 'monokai',
            showLineNumbers: true,
            showMinimap: true,
            wrap: true,
            recentFiles: []
        },
        'data-editor': {
            enabled: true,
            doLoadTxt: true,
            doLoadXml: true,
            doLoadJson: true,
            doLoadYaml: true,
            lineWrapping: true
        },
        'doc-viewer': {
            enabled: true,
            docxEnabled: true,
            pdfEnabled: true,
            defaultZoom: 100
        },
        'excel-to-markdown': {
            enabled: true,
            enabledAutoConvert: true
        },
        'folder-counter': {
            enabled: true,
            showAllNumbers: true,
            addRootFolder: false
        },
        'html-viewer': {
            enabled: true,
            opMode: 'BalanceMode',
            zoomValue: 1
        },
        'image-gallery': {
            enabled: true,
            imgFolder: '',
            sortby: 'mtime',
            sort: 'desc',
            gridCols: 3,
            displayMode: 'square'
        },
        'image-tools': {
            enabled: true,
            autoRename: false,
            resizeWidth: 800,
            quality: 80,
            format: 'webp'
        },
        'media-gallery': {
            enabled: true,
            scanFolder: '',
            sortOrder: 'date-desc',
            gridSize: 200,
            limit: 50,
            displayMode: 'square',
            mediaType: 'all'
        },
        'spreadsheet': {
            enabled: true,
            language: 'ZH',
            isSupportXlsx: true
        },
        'table-resize': {
            enabled: true,
            minColumnWidth: 50
        },
        'url-opener': {
            enabled: true,
            bookmarks: []
        },
        'vault-stats': {
            enabled: true,
            countComments: true,
            pageWords: 300
        },
        // ===== 移动端适配器（全局功能型模块，自检测平台）========
        'mobile-adapter': {
            enabled: true,          // 模块已加载；实际功能由 adaptLayout 控制
            adaptLayout: false,     // 初始值 false，由模块自身在 _detectMobile() 后决定
            portraitCols: 1,
            landscapeCols: 2,
            orientationLock: 'natural',
            minCardWidth: 280
        },
        // ===== 图片画廊 & 媒体画廊（全局功能，代码块处理器）========
        'img-gallery': {
            enabled: true
        },
        'memories': {
            enabled: true
        },
        'xhs-importer': {
            enabled: true,
            noteFolder: '收件箱',
            imageFolder: '附件',
            downloadMedia: true,
            customTags: '小红书'
        },
        // ======== 游戏娱乐模块（默认关闭）========
        'aquarium':      { enabled: false },
        'pixel-garden': { enabled: false },
        'particle-toy': { enabled: false },
        'farm-clicker': { enabled: false },
    },
    // ======== FileViewer 文件查看器扩展开关 ========
    // ★ 每个开关对应 modules/file-viewers/ 下的一个扩展
    // ★ 关闭后对应类型文件将释放给系统默认程序打开（而不是显示错误提示）
    fileViewerExtensions: {
        xlsx: true,     // 表格文件（XLSX/XLS/CSV/ODS）
        docx: true,     // Word 文档（DOCX）
        doc: true,      // 旧版 Word（DOC）
        html: true,     // HTML 网页
        image: true,    // 图片（PNG/JPG/GIF/SVG 等）
        video: true,    // 视频/音频/PDF
        office: true,   // PPT/PPTX 旧版 Office
        text: true      // 纯文本/代码（默认开启，关闭后无法查看任何未注册文件）
    },
    // 实例列表：每个实例 { id: 'weather#1', baseModule: 'weather', label: '天气 1' }
    instances: [],
    instanceCounter: 0,
    moduleOrder: ['weather', 'calendar', 'stats', 'todo', 'recent', 'news', 'directory', 'ai-insight', 'web-preview', 'web-video', 'code-editor', 'data-editor', 'doc-viewer', 'excel-to-markdown', 'html-viewer', 'image-gallery', 'media-gallery', 'spreadsheet', 'url-opener', 'xhs-importer',
        'aquarium', 'pixel-garden', 'particle-toy', 'farm-clicker'],
    headerBg: '',
    showHeader: true,
    cardBgColor: '',
    cardBgOpacity: 0.95,
    categoryCollapsed: {},  // { 'schedule': true, 'viewers': false, ... }
    sectionCollapsed: {},    // { 'file-viewer': false, 'utility': false } — 设置区域折叠状态
    // ======== 语言设置 ========
    language: 'system',       // 界面语言：'zh' | 'en' | 'system' | 'ai' | 'custom_xx'(AI翻译后)
    aiApiKey: '',            // AI 翻译 API Key
    aiApiUrl: '',            // AI 翻译 API 地址
    aiModel: '',             // AI 翻译模型名称
    aiTargetLang: 'auto',   // AI 翻译目标语言：'en' | 'zh' | 'ja' | 'ru' 等 | 'auto' 自动检测
    aiCustomLang: '',        // AI 翻译自定义语言名称（用户手动输入）
    aiQuickLang: '',         // AI 翻译快捷语言选择（下拉）
    aiCustomTranslations: {}, // AI 翻译保存的自定义译文 { langCode: { key: text } }
    customLangName: ''       // 当前自定义语言的显示名称
};

// ===================== 辅助函数 =====================
function formatSize(bytes) {
    if (!bytes || bytes === 0) return '0 B';
    var units = ['B', 'KB', 'MB', 'GB'];
    var i = Math.floor(Math.log(bytes) / Math.log(1024));
    if (i > units.length - 1) i = units.length - 1;
    return (bytes / Math.pow(1024, i)).toFixed(i > 0 ? 1 : 0) + ' ' + units[i];
}

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

        // —— 构建模式：从 BUILTIN_MODULES 直接加载（无需文件系统）——
        if (BUILTIN_MODULES && typeof BUILTIN_MODULES === 'object') {
            for (const [id, encoded] of Object.entries(BUILTIN_MODULES)) {
                try {
                    const code = encoded; // JSON.stringify 已自动转义，直接就是原始代码
                    const mod = this._evalModule(code, id);
                    if (mod && mod.render) {
                        const mid = mod.id || id;
                        this.modules.set(mid, mod);
                        if (mod.defaultSettings && typeof mod.defaultSettings === 'object') {
                            this._moduleDefaults.set(mid, mod.defaultSettings);
                        }
                    }
                } catch (e) {
                    console.warn('[V17] 内置模块 ' + id + ' 加载失败:', e);
                }
            }
            console.log('[V17] 已加载内置模块: ' + [...this.modules.keys()].join(', '));
            return;
        }

        // —— 开发模式：从文件系统加载 modules/ ——
        const modulesDir = this.plugin.manifest.dir + '/modules/';
        let files = [];
        try {
            files = await this.plugin.app.vault.adapter.list(modulesDir);
        } catch (e) {
            console.warn('[V17] 扫描模块目录失败:', e);
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
                    if (mod.defaultSettings && typeof mod.defaultSettings === 'object') {
                        this._moduleDefaults.set(id, mod.defaultSettings);
                    }
                }
            } catch (e) {
                console.warn('[V17] 模块 ' + moduleId + ' 加载失败:', e);
            }
        }

        console.log('[V17] 已加载模块: ' + [...this.modules.keys()].join(', '));
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
            console.error('[V17] 模块解析错误 (' + fallbackId + '):', e);
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
            t,              // ★ i18n 翻译函数（供模块 renderSettings 使用）
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
                text: t('mod.manager.notLoaded') + moduleId + t('mod.manager.notLoaded2'),
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

        // 更新运行时上下文：先清空再赋值，防止前一个模块的属性残留
        const ctx = this.createContext(moduleId);
        for (const key of Object.keys(this._runtimeCtx)) {
            delete this._runtimeCtx[key];
        }
        Object.assign(this._runtimeCtx, ctx);

        try {
            await mod.render(container);
        } catch (e) {
            console.error('[V17] 模块 ' + moduleId + ' 渲染错误:', e);
            container.createEl('div', {
                text: t('mod.manager.renderFailed') + e.message,
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

// ★ 全局功能型模块 — 不需要仪表盘面板，仅在设置中以开关控制
var UTILITY_MODULE_IDS = ['autoplay-loop', 'folder-counter', 'image-tools', 'table-resize', 'vault-stats', 'excel-to-markdown', 'mobile-adapter'];

class DashboardView extends ItemView {
    constructor(leaf, plugin) {
        super(leaf);
        this.plugin = plugin;
        this.draggedCard = null;
        this.dragOffset = { x: 0, y: 0 };
    }

    getViewType() { return VIEW_TYPE; }
    getDisplayText() { return t('app.dashboard'); }
    getIcon() { return 'layout-dashboard'; }

    async onOpen() {
        this.containerEl.empty();
        this.containerEl.addClass('v15-view');

        this.registerDomEvent(document, 'mousemove', (e) => this._onMouseMove(e));
        this.registerDomEvent(document, 'mouseup', (e) => this._onMouseUp(e));

        await this.render();
    }

    async onClose() {
        // ★ 修复：关闭仪表盘时，调用所有已加载模块的 onunload 清理定时器和监听器
        try {
            for (const [moduleId, mod] of this.plugin.moduleManager.modules) {
                if (typeof mod.onunload === 'function') {
                    try { mod.onunload(); } catch (e) { console.error('[V17] onunload error in ' + moduleId + ':', e); }
                }
            }
            // 再清理实例模块
            for (const inst of this.plugin.moduleManager.instances) {
                const baseId = inst.id.split('#')[0];
                const baseMod = this.plugin.moduleManager.getModule(baseId);
                if (baseMod && typeof baseMod.onunload === 'function') {
                    try { baseMod.onunload(); } catch (e) {}
                }
            }
        } catch (e) {
            console.error('[V17] onClose cleanup error:', e);
        }
    }

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

        root.style.setProperty('--background-primary', this._hexToRgba(theme.bg, opacity));
        root.style.setProperty('--text-normal', theme.text);
        root.style.setProperty('--text-muted', theme.muted);
        root.style.setProperty('--interactive-accent', theme.primary);
    }

    _hexToRgba(hex, alpha) {
        // ★ 兼容 hex / rgb / rgba 输入，始终应用透明度
        if (!hex || typeof hex !== 'string') return `rgba(255,255,255,${alpha})`;
        if (hex.startsWith('rgba')) {
            // 提取现有 rgba 的 rgb 部分，替换 alpha
            var m = hex.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*[\d.]+)?\s*\)/);
            if (m) return `rgba(${m[1]},${m[2]},${m[3]},${alpha})`;
            return hex;
        }
        if (hex.startsWith('rgb(')) {
            var m2 = hex.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
            if (m2) return `rgba(${m2[1]},${m2[2]},${m2[3]},${alpha})`;
            return hex;
        }
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

        // ★ 文件点击双层防护：
        // 第1层：canvas 级 preventDefault() 阻止浏览器/Electron 默认打开文件行为
        // 第2层：openLinkText 路由到 FileViewer，内部根据扩展名选择渲染引擎
        //        （SheetJS/mammoth/iframe/纯文本）
        const FILE_ITEM_SELECTOR = '.ss-file-item, .dv-file-item, .htmlv-file-item,' +
            ' .ce-file-item, .de-file-item, .mg-item,' +
            ' .ig-thumb, .uo-bm-item, .fc-folder-header';
        const killDefault = function(evt) {
            const card = evt.target.closest('.v6-card');
            if (!card) return;
            const fileItem = evt.target.closest(FILE_ITEM_SELECTOR);
            if (!fileItem) return;
            evt.preventDefault();
        };
        canvas.addEventListener('mousedown', killDefault, true);
        canvas.addEventListener('click', killDefault, true);

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
        const renderQueue = [];

        // 先按 moduleOrder 收集
        for (const moduleId of moduleOrder) {
            if (rendered.has(moduleId)) continue;
            if (!loadedIds.includes(moduleId) && !instanceIds.includes(moduleId)) continue;
            // ★ 跳过全局功能型模块（不需要仪表盘面板）
            var baseId = moduleId.indexOf('#') > 0 ? moduleId.substring(0, moduleId.indexOf('#')) : moduleId;
            if (UTILITY_MODULE_IDS.indexOf(baseId) !== -1) continue;
            const modSettings = this.plugin.settings.modules[moduleId];
            if (!modSettings || modSettings.enabled === false) continue;
            const mod = this.plugin.moduleManager.resolveModule(moduleId);
            if (!mod) continue;
            renderQueue.push({ moduleId, mod });
            rendered.add(moduleId);
        }

        // 收集剩余基础模块
        for (const moduleId of loadedIds) {
            if (rendered.has(moduleId)) continue;
            // ★ 跳过全局功能型模块（不需要仪表盘面板）
            if (UTILITY_MODULE_IDS.indexOf(moduleId) !== -1) continue;
            const modSettings = this.plugin.settings.modules[moduleId];
            if (!modSettings || modSettings.enabled === false) continue;
            const mod = this.plugin.moduleManager.getModule(moduleId);
            if (!mod) continue;
            renderQueue.push({ moduleId, mod });
            rendered.add(moduleId);
        }

        // 收集剩余实例
        for (const inst of instances) {
            if (rendered.has(inst.id)) continue;
            const modSettings = this.plugin.settings.modules[inst.id];
            if (!modSettings || modSettings.enabled === false) continue;
            const mod = this.plugin.moduleManager.resolveModule(inst.id);
            if (!mod) continue;
            renderQueue.push({ moduleId: inst.id, mod });
            rendered.add(inst.id);
        }

        // ★ 修复：串行异步批量渲染，每批2个模块，每批之间 yield 150ms 给UI线程呼吸
        // 之前的 setTimeout 方案有两个致命问题：
        // 1. renderModule 未 await → 26个模块并发抢占 _runtimeCtx，状态互相覆盖
        // 2. 前4个模块完全同步阻塞 → UI帧冻结
        const BATCH_SIZE = 2;
        const BATCH_DELAY = 150;

        // 加载进度指示器
        const progressEl = canvas.createDiv({ cls: 'v15-loading-progress' });
        progressEl.innerHTML = t('dashboard.loading');
        progressEl.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:9999;background:var(--background-primary);border:1px solid var(--background-modifier-border);border-radius:8px;padding:8px 16px;font-size:12px;color:var(--text-muted);box-shadow:0 2px 8px rgba(0,0,0,0.1);pointer-events:none;';

        for (let i = 0; i < renderQueue.length; i++) {
            const { moduleId, mod } = renderQueue[i];

            // 更新进度
            progressEl.innerHTML = t('dashboard.loadingProgress') + (i + 1) + '/' + renderQueue.length;

            // 同步创建卡片容器（轻量DOM操作，不会卡）
            const card = this._createModuleCard(canvas, moduleId, mod);

            // ★ 关键：await 确保模块渲染完成后再继续下一个
            // 避免多个模块同时抢占共享的 _runtimeCtx
            const content = card.querySelector('.v6-card-content');
            if (content) {
                const t0 = performance.now();
                await this.plugin.moduleManager.renderModule(moduleId, content);
                const t1 = performance.now();
                const dur = (t1 - t0).toFixed(0);
                if (dur > 500) console.warn('[V17] SLOW MODULE: ' + moduleId + ' took ' + dur + 'ms');
                // 进度条更新
                progressEl.innerHTML = t('dashboard.loadingProgress') + (i + 1) + '/' + renderQueue.length + ' (' + dur + 'ms) ' + moduleId;
            }

            // 给 UI 线程喘息：每 BATCH_SIZE 个模块后 yield
            if ((i + 1) % BATCH_SIZE === 0 && i < renderQueue.length - 1) {
                await new Promise(resolve => setTimeout(resolve, BATCH_DELAY));
            }
        }

        // 加载完成
        progressEl.innerHTML = t('dashboard.ready');
        setTimeout(() => { if (progressEl.isConnected) progressEl.remove(); }, 1500);
    }

    // 抽出轻量卡片创建（仅 DOM 搭建，不做渲染）
    _createModuleCard(canvas, moduleId, mod) {
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

        const isInstance = moduleId.indexOf('#') > 0;
        const instanceInfo = isInstance ? this.plugin.moduleManager.getInstanceInfo(moduleId) : null;

        const cardHeader = card.createDiv({ cls: 'v6-card-header' });
        const titleArea = cardHeader.createDiv({ cls: 'v6-card-title' });

        if (isInstance && instanceInfo) {
            const baseMod = this.plugin.moduleManager.getModule(instanceInfo.baseModule);
            titleArea.createEl('span', { text: (baseMod ? baseMod.icon : '📦') + ' ' + instanceInfo.label, cls: 'v6-card-label' });
        } else {
            titleArea.createEl('span', { text: mod.icon || '📦', cls: 'v6-card-icon' });
            titleArea.createEl('span', { text: mod.title || moduleId, cls: 'v6-card-label' });
        }

        const content = card.createDiv({ cls: 'v6-card-content' });
        content.style.overflow = 'auto';
        content.style.height = 'calc(100% - 50px)';

        const refreshBtn = cardHeader.createEl('button', {
            cls: 'v6-card-btn', attr: { title: t('dashboard.refresh') }
        });
        refreshBtn.innerHTML = '↺';
        refreshBtn.addEventListener('click', async (e) => {
            e.stopPropagation();
            content.empty();
            await this.plugin.moduleManager.renderModule(moduleId, content);
        });

        if (isInstance) {
            const removeBtn = cardHeader.createEl('button', {
                cls: 'v6-card-btn', attr: { title: t('dashboard.removeInstance') }
            });
            removeBtn.innerHTML = '✕';
            removeBtn.style.color = 'var(--text-error)';
            removeBtn.addEventListener('click', async (e) => {
                e.stopPropagation();
                await this._removeInstance(moduleId);
            });
        }

        cardHeader.addEventListener('mousedown', (e) => this._onDragStart(e, card));

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
            if (card.contains(e.target)) { setTimeout(saveSize, 50); }
        };
        document.addEventListener('mouseup', globalMouseUp);
        this.register(() => document.removeEventListener('mouseup', globalMouseUp));

        return card;
    }

    _renderHeader(parent) {
        const header = parent.createDiv({ cls: 'v15-header' });

        const left = header.createDiv({ cls: 'v15-header-left' });
        left.createEl('span', { text: '🏠', cls: 'v15-header-icon' });
        left.createEl('span', { text: t('app.dashboard'), cls: 'v15-header-title' });

        const right = header.createDiv({ cls: 'v15-header-right' });

        // ★ 新增：添加板块按钮
        const addBtn = right.createEl('button', {
            cls: 'v15-header-btn',
            attr: { title: t('dashboard.addSection') }
        });
        addBtn.innerHTML = '➕';
        addBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this._showAddMenu(addBtn);
        });

        // ★ 智能自动排序按钮：货架排列算法，保留每个模块的当前尺寸，自动换行紧凑排列
        const sortBtn = right.createEl('button', {
            cls: 'v15-header-btn',
            attr: { title: t('dashboard.smartSort') }
        });
        sortBtn.innerHTML = '📐';
        sortBtn.addEventListener('click', async (e) => {
            e.stopPropagation();
            await this._autoSortLayout();
        });

        // 主题切换
        const themeBtn = right.createEl('button', {
            cls: 'v15-header-btn',
            attr: { title: t('dashboard.switchTheme') }
        });
        themeBtn.innerHTML = '🎨';
        themeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this._showThemeMenu(themeBtn);
        });

        // 刷新
        const refreshBtn = right.createEl('button', {
            cls: 'v15-header-btn',
            attr: { title: t('dashboard.refresh') }
        });
        refreshBtn.innerHTML = '🔄';
        refreshBtn.addEventListener('click', () => this.render());

        // 设置
        const settingsBtn = right.createEl('button', {
            cls: 'v15-header-btn',
            attr: { title: t('dashboard.settings') }
        });
        settingsBtn.innerHTML = '⚙️';
        settingsBtn.addEventListener('click', () => {
            try {
                this.plugin.app.setting.open();
            } catch (e) {
                console.warn('[V17] 打开设置失败:', e);
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
            text: t('dashboard.addSectionTitle'),
            attr: { style: 'padding:4px 8px;font-size:11px;color:var(--text-muted);font-weight:bold;' }
        });

        const allModules = this.plugin.moduleManager.getAllModules();
        allModules.forEach(mod => {
            const baseId = mod.id;
            if (!baseId) return;
            // ★ 跳过全局功能型模块（不需要仪表盘面板）
            if (UTILITY_MODULE_IDS.indexOf(baseId) !== -1) return;
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
            new Notice(t('dashboard.moduleNotLoaded') + baseModule + t('dashboard.moduleNotLoaded2'));
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
        new Notice(t('dashboard.added') + label);
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
        new Notice(t('dashboard.removed') + label);
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

    // 兼容旧调用链：完整渲染单个卡片（创建 DOM + 触发模块 render）
    // 用于 _addInstance / _showThemeMenu 等非 render() 主循环的路径
    async renderModuleCard(canvas, moduleId, mod) {
        const card = this._createModuleCard(canvas, moduleId, mod);
        const content = card.querySelector('.v6-card-content');
        if (content) {
            await this.plugin.moduleManager.renderModule(moduleId, content);
        }
        return card;
    }

    // ★ 智能自动排序：货架排列（Shelf Packing）算法
    // 按行打包：每行从左到右排列模块，换行时下一个模块放在上一行最高模块下方。
    // 这样保证左右方向是连续的（按 moduleOrder 顺序），上下方向自动对齐。
    // 同一行的模块顶部对齐，不同行之间紧凑排列。
    async _autoSortLayout() {
        var MARGIN = 20;
        var GAP = 12;
        var MAX_ROW_WIDTH = 1800; // 画布最大宽度

        // 获取当前显示的所有模块
        var loadedIds = this.plugin.moduleManager.getLoadedModuleIds();
        var instances = this.plugin.settings.instances || [];
        var moduleOrder = this.plugin.settings.moduleOrder || [];

        var items = [];
        var seen = {};
        var self = this;

        function addItem(id) {
            if (seen[id]) return;
            // ★ 跳过全局功能型模块（不需要仪表盘面板）
            var baseId = id.indexOf('#') > 0 ? id.substring(0, id.indexOf('#')) : id;
            if (UTILITY_MODULE_IDS.indexOf(baseId) !== -1) return;
            var modSettings = self.plugin.settings.modules[id];
            if (!modSettings || modSettings.enabled === false) return;
            if (!loadedIds.includes(id) && !instances.some(function(i) { return i.id === id; })) return;

            var layout = self.plugin.settings.layout[id];
            var w = (layout && layout.width) ? layout.width : 280;
            var h = (layout && layout.height) ? layout.height : 250;
            w = Math.max(200, Math.min(w, 800));
            h = Math.max(150, Math.min(h, 600));

            items.push({ id: id, width: w, height: h });
            seen[id] = true;
        }

        moduleOrder.forEach(function(id) { addItem(id); });
        loadedIds.forEach(function(id) { addItem(id); });
        instances.forEach(function(inst) { addItem(inst.id); });

        if (items.length === 0) {
            new Notice(t('dashboard.nothingToSort'));
            return;
        }

        // === 货架排列算法 ===
        // rowX: 当前行已占用的X位置
        // rowY: 当前行的Y位置
        // rowHeight: 当前行的最大高度
        var rowX = MARGIN;
        var rowY = MARGIN;
        var rowHeight = 0;
        var layout = {};

        for (var i = 0; i < items.length; i++) {
            var item = items[i];

            // 检查是否需要换行（当前行的剩余宽度不够放此模块）
            if (rowX + item.width > MAX_ROW_WIDTH && rowX > MARGIN) {
                // 换行
                rowY += rowHeight + GAP;
                rowX = MARGIN;
                rowHeight = 0;
            }

            // 放置模块
            layout[item.id] = {
                x: Math.round(rowX),
                y: Math.round(rowY),
                width: item.width,
                height: item.height
            };

            // 更新行状态
            rowX += item.width + GAP;
            if (item.height > rowHeight) rowHeight = item.height;
        }

        self.plugin.settings.layout = layout;
        self.plugin.settings.moduleOrder = items.map(function(it) { return it.id; });
        await self.plugin.saveSettings();

        new Notice(t('dashboard.sorted') + items.length + t('dashboard.sortedSuffix'));
        self.render();
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

// ★ 引用 06-dashboard-view.js 中声明的 UTILITY_MODULE_IDS
//    这些模块不需要仪表盘面板，仅在设置中以开关控制

class DashboardSettingTab extends PluginSettingTab {
    constructor(app, plugin) {
        super(app, plugin);
        this.plugin = plugin;
        this._currentModuleId = null;
    }

    display() {
        const { containerEl } = this;
        var targetModuleId = this._currentModuleId;

        containerEl.empty();

        // 注入行内设置容器的样式（不会被 empty() 清除的在 head 中，这里用 data-style 标识防重复）
        if (!document.getElementById('v15-inline-settings-css')) {
            const cssEl = document.createElement('style');
            cssEl.id = 'v15-inline-settings-css';
            cssEl.textContent = `
                .v15-inline-settings-container {
                    margin: -4px 0 16px 24px;
                    padding: 0;
                    border-left: 3px solid var(--interactive-accent);
                    border-radius: 0 6px 6px 0;
                    background: var(--background-secondary);
                    overflow: hidden;
                }
                .v15-inline-settings-container .v15-module-settings-wrapper {
                    padding: 12px 16px;
                }
                .v15-inline-settings-container .v15-module-settings-wrapper h3 {
                    margin-top: 0;
                    font-size: 14px;
                }
                .v15-inline-settings-container .setting-item {
                    border-top: none;
                    padding: 8px 0;
                }
            `;
            document.head.appendChild(cssEl);
        }

        containerEl.createEl('h2', { text: t('settings.title') });

        this._renderLanguageSettings(containerEl);
        this._renderAppearanceSettings(containerEl);
        this._renderModuleToggles(containerEl);
        this._renderFileViewerSettings(containerEl);
        this._renderUtilityToggles(containerEl);
        this._renderInstanceManager(containerEl);

        // 配置注入到对应模块行下方（手风琴式），而非页面底部
        if (targetModuleId) {
            var inlineTarget = containerEl.querySelector('[data-inline-settings="' + targetModuleId + '"]');
            if (inlineTarget) {
                inlineTarget.style.display = 'block';
                inlineTarget.empty();
                this._renderModuleSettings(inlineTarget, targetModuleId);
            }
        }

        // 设置面板底部固定展示打赏（可通过模块设置关闭）
        this._renderDonateSection(containerEl);
    }

    // ======== AI 翻译预览弹窗 ========
    _showTranslationPreview(result, targetLang) {
        var self = this;
        var modal = new Modal(this.plugin.app);
        modal.titleEl.setText(t('settings.aiPreview.title') + ' — ' + targetLang);
        modal.contentEl.style.maxHeight = '60vh';
        modal.contentEl.style.overflowY = 'auto';

        var info = modal.contentEl.createDiv({ attr: { style: 'margin-bottom:12px;padding:8px;background:var(--background-secondary);border-radius:6px;font-size:13px;' } });
        info.setText(t('settings.aiPreview.translated') + result.total + t('settings.aiPreview.translatedSuffix') + result.usedLang);

        // 翻译结果列表
        var list = modal.contentEl.createDiv({ attr: { style: 'max-height:40vh;overflow-y:auto;' } });
        var entries = [];
        for (var k in result.translations) {
            entries.push({ key: k, text: result.translations[k] });
        }

        entries.forEach(function(e) {
            var row = list.createDiv({ attr: { style: 'display:flex;gap:8px;padding:4px 0;border-bottom:1px solid var(--background-modifier-border);font-size:12px;' } });
            row.createEl('span', { text: e.key, attr: { style: 'flex:0 0 200px;color:var(--text-muted);font-family:monospace;word-break:break-all;' } });
            row.createEl('span', { text: '→', attr: { style: 'flex:0 0 20px;text-align:center;' } });
            row.createEl('span', { text: e.text, attr: { style: 'flex:1;color:var(--text-normal);word-break:break-all;' } });
        });

        // 按钮栏
        var btnBar = modal.contentEl.createDiv({ attr: { style: 'display:flex;gap:8px;margin-top:12px;justify-content:flex-end;' } });

        var cancelBtn = btnBar.createEl('button', { text: t('settings.aiPreview.cancel'), attr: { style: 'padding:6px 16px;border:1px solid var(--background-modifier-border);border-radius:6px;background:var(--background-secondary);cursor:pointer;' } });
        cancelBtn.addEventListener('click', () => modal.close());

        var applyBtn = btnBar.createEl('button', { text: t('settings.aiPreview.apply'), attr: { style: 'padding:6px 16px;border:none;border-radius:6px;background:var(--interactive-accent);color:var(--text-on-accent);cursor:pointer;font-weight:600;' } });
        applyBtn.addEventListener('click', async () => {
            var lang = targetLang;
            var langCode = 'custom_' + lang;
            // 保存到 settings（持久化）—— 用 custom_xx 作 key
            if (!self.plugin.settings.aiCustomTranslations) {
                self.plugin.settings.aiCustomTranslations = {};
            }
            self.plugin.settings.aiCustomTranslations[langCode] = result.translations;
            self.plugin.settings.language = langCode;
            self.plugin.settings.customLangName = lang;
            // 运行时生效：_currentLang 和 _customTranslations 都用 custom_xx 格式
            setLanguage(langCode);
            setCustomTranslations(langCode, result.translations);
            await self.plugin.saveSettings();
            modal.close();
            self.plugin.refreshView();
            new Notice('✅ ' + t('settings.aiPreview.applySuccess') + ' [' + lang + ']', 5000);
            self.display();
        });

        modal.open();
    }

    // ======== 语言设置（新增，放在最顶部） ========
    _renderLanguageSettings(containerEl) {
        var self = this;
        containerEl.createEl('h3', { text: t('settings.language') });

        new Setting(containerEl)
            .setName(t('settings.language'))
            .setDesc(t('settings.language.desc'))
            .addDropdown(d => {
                d.addOption('zh', t('settings.lang.zh'))
                 .addOption('en', t('settings.lang.en'))
                 .addOption('system', t('settings.lang.system'))
                 .addOption('ai', t('settings.lang.ai'));
                // ★ 动态添加已保存的自定义语言（如 custom_ja, custom_ru 等）
                var saved = this.plugin.settings.aiCustomTranslations || {};
                Object.keys(saved).forEach(function(langKey) {
                    if (langKey.startsWith('custom_')) {
                        var name = langKey.substring(7).toUpperCase();
                        d.addOption(langKey, '⚡ ' + name + ' (AI)');
                    }
                });
                d.setValue(this.plugin.settings.language || 'system')
                 .onChange(async (v) => {
                    this.plugin.settings.language = v;
                    setLanguage(v);
                    await this.plugin.saveSettings();
                    this.plugin.refreshView();
                    // AI 模式：提示配置 API Key
                    if (v === 'ai') {
                        var hasKey = this.plugin.settings.aiApiKey && this.plugin.settings.aiApiKey.length > 0;
                        if (!hasKey) {
                            new Notice('AI ' + t('settings.language') + ': ' + t('settings.aiKey.desc'), 8000);
                        }
                    }
                    // 立即生效：重新渲染设置面板
                    this.display();
                });
            });

        // AI 翻译配置面板（仅在选择 AI 时展开）
        var aiSection = containerEl.createDiv({ cls: 'v15-inline-settings-container' });
        var isAiMode = (this.plugin.settings.language === 'ai');
        if (!isAiMode) {
            aiSection.style.display = 'none';
        } else {
            aiSection.style.marginTop = '-8px';
        }

        var aiWrapper = aiSection.createDiv({ cls: 'v15-module-settings-wrapper' });

        // ★ 说明文字：告诉用户可以用任何 AI 模型翻译任何语言（包括方言）
        var hintEl = aiWrapper.createDiv({
            cls: 'v15-i18n-hint',
            attr: {
                style: 'margin-bottom:12px;padding:10px 14px;background:var(--background-secondary);border-radius:8px;border-left:3px solid var(--interactive-accent);font-size:12px;line-height:1.6;color:var(--text-muted);'
            }
        });
        hintEl.innerHTML = '<strong>🌍 ' + t('settings.aiCustomLang.helpTitle') + '</strong><br>' +
            t('settings.aiCustomLang.helpLine1') + '<br>' +
            t('settings.aiCustomLang.helpLine2') + '<br>' +
            t('settings.aiCustomLang.helpLine3');

        // AI API Key
        new Setting(aiWrapper)
            .setName(t('settings.aiKey'))
            .setDesc(t('settings.aiKey.desc'))
            .addText(txt => {
                txt.setPlaceholder(t('settings.aiKey.placeholder'))
                    .setValue(this.plugin.settings.aiApiKey || '')
                    .onChange(async (v) => {
                        this.plugin.settings.aiApiKey = v;
                        await this.plugin.saveSettings();
                    });
            });

        // AI API URL
        new Setting(aiWrapper)
            .setName(t('settings.aiApiUrl'))
            .setDesc(t('settings.aiApiUrl.desc'))
            .addText(txt => {
                txt.setPlaceholder(t('settings.aiApiUrl.placeholder'))
                    .setValue(this.plugin.settings.aiApiUrl || '')
                    .onChange(async (v) => {
                        this.plugin.settings.aiApiUrl = v;
                        await this.plugin.saveSettings();
                    });
            });

        // AI Model
        new Setting(aiWrapper)
            .setName(t('settings.aiModel'))
            .setDesc(t('settings.aiModel.desc'))
            .addText(txt => {
                txt.setPlaceholder(t('settings.aiModel.placeholder'))
                    .setValue(this.plugin.settings.aiModel || '')
                    .onChange(async (v) => {
                        this.plugin.settings.aiModel = v;
                        await this.plugin.saveSettings();
                    });
            });

        // ★ 自定义语言输入（核心：让用户输入任意语言名）
        new Setting(aiWrapper)
            .setName(t('settings.aiCustomLang'))
            .setDesc(t('settings.aiCustomLang.desc'))
            .addText(txt => {
                txt.setPlaceholder(t('settings.aiCustomLang.placeholder'))
                    .setValue(this.plugin.settings.aiCustomLang || '')
                    .onChange(async (v) => {
                        this.plugin.settings.aiCustomLang = v;
                        await this.plugin.saveSettings();
                    });
            });

        // ★ 预设语言快捷选项
        new Setting(aiWrapper)
            .setName(t('settings.aiQuickLang'))
            .setDesc(t('settings.aiQuickLang.desc'))
            .addDropdown(d => {
                d.addOption('', '— ' + t('settings.aiQuickLang.none') + ' —')
                 .addOption('en', 'English')
                 .addOption('zh', t('settings.lang.zh'))
                 .addOption('ja', '日本語')
                 .addOption('ko', '한국어')
                 .addOption('ru', 'Русский')
                 .addOption('de', 'Deutsch')
                 .addOption('fr', 'Français')
                 .addOption('es', 'Español')
                 .addOption('pt', 'Português')
                 .addOption('it', 'Italiano')
                 .addOption('ar', 'العربية')
                 .addOption('hi', 'हिन्दी')
                 .addOption('th', 'ไทย')
                 .addOption('vi', 'Tiếng Việt')
                 .addOption('id', 'Bahasa Indonesia')
                 .addOption('nl', 'Nederlands')
                 .addOption('pl', 'Polski')
                 .addOption('uk', 'Українська')
                 .setValue(this.plugin.settings.aiQuickLang || '')
                 .onChange(async (v) => {
                    // 同步到自定义语言输入框
                    if (v && !this.plugin.settings.aiCustomLang) {
                        this.plugin.settings.aiCustomLang = v;
                    }
                    this.plugin.settings.aiQuickLang = v;
                    await this.plugin.saveSettings();
                    this.display();
                 });
            });

        // ★ 翻译按钮
        new Setting(aiWrapper)
            .setName(t('settings.aiTranslateBtn'))
            .setDesc(t('settings.aiTranslateBtn.desc'))
            .addButton(b => {
                b.setButtonText(t('settings.aiTranslateBtn'))
                 .setCta()
                 .onClick(async () => {
                    var key = this.plugin.settings.aiApiKey;
                    if (!key) {
                        new Notice('❌ ' + t('settings.aiKey') + ' required', 5000);
                        return;
                    }
                    b.setButtonText(t('settings.aiTranslating'));
                    b.setDisabled(true);

                    // 确定目标语言：优先用自定义输入，其次用快捷下拉
                    var customLang = (this.plugin.settings.aiCustomLang || '').trim();
                    var quickLang = this.plugin.settings.aiQuickLang || '';
                    var targetLang = customLang || quickLang || 'auto';

                    try {
                        var result = await aiTranslateAll(
                            key,
                            this.plugin.settings.aiApiUrl,
                            this.plugin.settings.aiModel,
                            targetLang
                        );
                        if (result.success) {
                            if (result.total === 0) {
                                // 0条目：给出详细诊断信息
                                var diagMsg = t('settings.aiTranslateFailed')
                                    + (getCurrentLang() === 'en' || getCurrentLang() === 'ai'
                                        ? '0 entries translated. Possible causes:\n1. AI returned non-standard format (check console for raw response)\n2. API URL may need full path (e.g. https://your-host/v1/chat/completions)\n3. Model returned unexpected format'
                                        : '翻译了0条目。可能原因：\n1. AI返回格式不符（请查看控制台日志了解原始响应）\n2. API地址需要完整路径（如 https://your-host/v1/chat/completions）\n3. 模型返回了意外格式');
                                new Notice('⚠️ ' + diagMsg, 8000);
                            } else {
                                self._showTranslationPreview(result, targetLang);
                            }
                        } else {
                            new Notice('❌ ' + t('settings.aiTranslateFailed') + (result.error || ''), 5000);
                        }
                    } catch(e) {
                        new Notice('❌ ' + t('settings.aiTranslateFailed') + e.message, 5000);
                    }
                    b.setButtonText(t('settings.aiTranslateBtn'));
                    b.setDisabled(false);
                 });
            });
    }

    _renderAppearanceSettings(containerEl) {
        containerEl.createEl('h3', { text: t('settings.appearance') });

        new Setting(containerEl)
            .setName(t('settings.theme'))
            .setDesc(t('settings.theme.desc'))
            .addDropdown(d => {
                Object.entries(THEMES).forEach(([id, t2]) => d.addOption(id, t2.name));
                d.setValue(this.plugin.settings.theme)
                    .onChange(async (v) => {
                        this.plugin.settings.theme = v;
                        await this.plugin.saveSettings();
                        this.plugin.refreshView();
                    });
            });

        new Setting(containerEl)
            .setName(t('settings.showHeader'))
            .setDesc(t('settings.showHeader.desc'))
            .addToggle(t2 => {
                t2.setValue(this.plugin.settings.showHeader !== false)
                    .onChange(async (v) => {
                        this.plugin.settings.showHeader = v;
                        await this.plugin.saveSettings();
                        this.plugin.refreshView();
                    });
            });

        new Setting(containerEl)
            .setName(t('settings.cardBgColor'))
            .setDesc(t('settings.cardBgColor.desc'))
            .addText(t2 => {
                t2.setPlaceholder(t('settings.cardBgColor.placeholder'))
                    .setValue(this.plugin.settings.cardBgColor || '')
                    .onChange(async (v) => {
                        this.plugin.settings.cardBgColor = v;
                        await this.plugin.saveSettings();
                        this.plugin.refreshView();
                    });
            });

        new Setting(containerEl)
            .setName(t('settings.cardBgOpacity'))
            .setDesc(t('settings.cardBgOpacity.desc'))
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
            .setName(t('settings.resetLayout'))
            .setDesc(t('settings.resetLayout.desc'))
            .addButton(b => {
                b.setButtonText(t('settings.resetBtn')).setWarning()
                    .onClick(async () => {
                        this.plugin.settings.layout = {};
                        await this.plugin.saveSettings();
                        this.plugin.refreshView();
                        new Notice(t('settings.layoutReset'));
                    });
            });

        // ======== 设置存档：导入/导出 ========
        const backupSection = containerEl.createDiv({
            attr: { style: 'margin-top:24px;padding-top:20px;border-top:1px solid var(--background-modifier-border);' }
        });
        backupSection.createEl('h3', { text: t('settings.backup') });

        new Setting(backupSection)
            .setName(t('settings.export'))
            .setDesc(t('settings.export.desc'))
            .addButton(b => {
                b.setButtonText(t('settings.exportBtn'))
                    .setCta()
                    .onClick(() => {
                        const data = JSON.stringify(this.plugin.settings, null, 2);
                        const blob = new Blob([data], { type: 'application/json' });
                        const url = URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = 'obsidian-dashboard-settings-' + new Date().toISOString().slice(0, 10) + '.json';
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        URL.revokeObjectURL(url);
                        new Notice(t('settings.exportSuccess'));
                    });
            });

        new Setting(backupSection)
            .setName(t('settings.import'))
            .setDesc(t('settings.import.desc'))
            .addButton(b => {
                b.setButtonText(t('settings.importBtn'))
                    .onClick(() => {
                        const input = document.createElement('input');
                        input.type = 'file';
                        input.accept = '.json';
                        input.onchange = async (e) => {
                            const file = e.target.files[0];
                            if (!file) return;
                            try {
                                const text = await file.text();
                                const imported = JSON.parse(text);
                                // 简单校验：至少要有 theme 和 modules 字段
                                if (!imported.modules && !imported.theme) {
                                    new Notice(t('settings.importInvalid'), 5000);
                                    return;
                                }
                                // 合并导入设置
                                Object.assign(this.plugin.settings, imported);
                                // 如果导入设置中有语言配置，立即生效
                                if (imported.language) {
                                    setLanguage(imported.language);
                                }
                                await this.plugin.saveSettings();
                                this.plugin.refreshView();
                                new Notice(t('settings.importSuccess'), 5000);
                            } catch (err) {
                                new Notice(t('settings.importFailed') + err.message, 5000);
                            }
                        };
                        input.click();
                    });
            });
    }

    _renderModuleToggles(containerEl) {
        var self = this;
        containerEl.createEl('h3', { text: t('settings.moduleManage') });

        const loadedModules = this.plugin.moduleManager.getAllModules();
        if (loadedModules.length === 0) {
            containerEl.createEl('p', {
                text: t('settings.noModules'),
                attr: { style: 'color: var(--text-muted); font-size: 13px;' }
            });
            return;
        }

        // ============ 分类体系 ============
        var CATEGORIES = [
            { id: 'schedule', icon: '📅', name: t('settings.cat.schedule'), modules: ['calendar', 'stats', 'todo', 'recent'] },
            { id: 'viewers', icon: '👁️', name: t('settings.cat.viewers'),   modules: ['doc-viewer', 'spreadsheet', 'html-viewer', 'code-editor', 'data-editor', 'ppt-viewer'] },
            { id: 'notes',   icon: '📝', name: t('settings.cat.notes'),     modules: [] },
            { id: 'files',   icon: '📂', name: t('settings.cat.files'),     modules: ['directory'] },
            { id: 'media',   icon: '🎬', name: t('settings.cat.media'),     modules: ['image-gallery', 'media-gallery'] },
            { id: 'web',     icon: '🌍', name: t('settings.cat.web'),       modules: ['web-preview', 'web-video', 'url-opener', 'news', 'weather'] },
            { id: 'ai',      icon: '🤖', name: t('settings.cat.ai'),        modules: ['ai-insight', 'xhs-importer'] },
            { id: 'games',  icon: '🎮', name: t('settings.cat.games'),     modules: ['aquarium', 'pixel-garden', 'particle-toy', 'farm-clicker'] }
        ];

        // 构建 moduleId -> mod 的映射
        var modMap = {};
        loadedModules.forEach(function(mod) { if (mod.id) modMap[mod.id] = mod; });

        // 追踪未分类模块
        var uncategorized = [];
        var categorizedIds = new Set();
        CATEGORIES.forEach(function(cat) { cat.modules.forEach(function(mid) { categorizedIds.add(mid); }); });
        loadedModules.forEach(function(mod) {
            if (mod.id && !categorizedIds.has(mod.id) && UTILITY_MODULE_IDS.indexOf(mod.id) === -1) uncategorized.push(mod.id);
        });
        if (uncategorized.length > 0) {
            CATEGORIES.push({ id: 'other', icon: '📦', name: t('settings.cat.other'), modules: uncategorized });
        }

        // 渲染分类折叠面板（带记忆）
        var collapsedState = self.plugin.settings.categoryCollapsed || {};

        CATEGORIES.forEach(function(cat) {
            // 筛选出实际存在的模块
            var existingMods = cat.modules.filter(function(mid) { return modMap[mid]; });
            if (existingMods.length === 0) return;

            // 读记忆：该分类是否折叠（默认展开）
            var isCollapsed = collapsedState[cat.id] === true;
            var categoryEl = containerEl.createDiv({ cls: 'v15-category' });
            if (!isCollapsed) categoryEl.classList.add('open');

            // 分类头部
            var header = categoryEl.createDiv({ cls: 'v15-cat-header' });
            header.createSpan({ text: cat.icon, cls: 'v15-cat-icon' });
            header.createSpan({ text: cat.name, cls: 'v15-cat-label' });
            header.createSpan({ text: existingMods.length, cls: 'v15-cat-count' });
            var arrow = header.createSpan({ text: '▶', cls: 'v15-cat-arrow' });

            header.addEventListener('click', function() {
                categoryEl.classList.toggle('open');
                // 保存折叠状态
                var isNowCollapsed = !categoryEl.classList.contains('open');
                if (!self.plugin.settings.categoryCollapsed) self.plugin.settings.categoryCollapsed = {};
                self.plugin.settings.categoryCollapsed[cat.id] = isNowCollapsed;
                self.plugin.saveSettings();
            });

            // 分类内容
            var body = categoryEl.createDiv({ cls: 'v15-cat-body' });

            existingMods.forEach(function(moduleId) {
                var mod = modMap[moduleId];
                var modSettings = self.plugin.settings.modules[moduleId] || {};

                // 直接用 Setting API 创建
                try {
                    new Setting(body)
                        .setName((mod.icon || '📦') + ' ' + (mod.title || moduleId))
                        .setDesc(mod.id)
                        .addToggle(function(t2) {
                            t2.setValue(modSettings.enabled !== false)
                                .onChange(async function(v) {
                                    if (!self.plugin.settings.modules[moduleId]) {
                                        self.plugin.settings.modules[moduleId] = {};
                                    }
                                    self.plugin.settings.modules[moduleId].enabled = v;
                                    await self.plugin.saveSettings();
                                    self.plugin.refreshView();
                                });
                        })
                        .addButton(function(b) {
                            b.setButtonText(t('settings.configureBtn'))
                                .onClick(function() {
                                    // ★ 手风琴式 toggle：不调 display()，直接在当前行下方展开/收拢配置面板
                                    var newId = self._currentModuleId === moduleId ? null : moduleId;

                                    // 收拢旧的
                                    if (self._currentModuleId && self._currentModuleId !== moduleId) {
                                        var oldEl = containerEl.querySelector('[data-inline-settings="' + self._currentModuleId + '"]');
                                        if (oldEl) { oldEl.style.display = 'none'; oldEl.empty(); }
                                    }

                                    self._currentModuleId = newId;

                                    // 展开或收拢当前
                                    var target = containerEl.querySelector('[data-inline-settings="' + moduleId + '"]');
                                    if (newId && target) {
                                        target.style.display = 'block';
                                        target.empty();
                                        self._renderModuleSettings(target, newId);
                                        // 轻量滚动：有需要才把配置面板带入视口，block:'nearest' 不跳滚
                                        target.scrollIntoView({ block: 'nearest', behavior: 'instant' });
                                    } else if (target) {
                                        target.style.display = 'none';
                                        target.empty();
                                    }
                                });
                        });
                } catch(e) {
                    console.warn('[V17] 渲染模块设置项失败:', moduleId, e);
                }
                // 行内隐藏容器：点击"配置"时在此展开，而非滚到底部
                body.createDiv({
                    cls: 'v15-inline-settings-container',
                    attr: { 'data-inline-settings': moduleId, style: 'display:none;' }
                });
            });
        });
    }

    _renderFileViewerSettings(containerEl) {
        var self = this;
        var collapsedState = this.plugin.settings.sectionCollapsed || {};
        var isCollapsed = collapsedState['file-viewer'] === true;

        var categoryEl = containerEl.createDiv({ cls: 'v15-category' });
        if (!isCollapsed) categoryEl.classList.add('open');

        var header = categoryEl.createDiv({ cls: 'v15-cat-header' });
        header.createSpan({ text: '📂', cls: 'v15-cat-icon' });
        header.createSpan({ text: t('settings.fileViewer'), cls: 'v15-cat-label' });
        header.createSpan({ text: '8' + '', cls: 'v15-cat-count' });
        var arrow = header.createSpan({ text: '▶', cls: 'v15-cat-arrow' });

        header.addEventListener('click', function() {
            categoryEl.classList.toggle('open');
            if (!self.plugin.settings.sectionCollapsed) self.plugin.settings.sectionCollapsed = {};
            self.plugin.settings.sectionCollapsed['file-viewer'] = !categoryEl.classList.contains('open');
            self.plugin.saveSettings();
        });

        var body = categoryEl.createDiv({ cls: 'v15-cat-body' });

        var fvSettings = this.plugin.settings.fileViewerExtensions || {};

        var fvGroups = [
            { key: 'xlsx',  name: t('settings.fv.ext.xlsx'),  desc: t('settings.fv.ext.xlsx.desc') },
            { key: 'docx',  name: t('settings.fv.ext.docx'),  desc: t('settings.fv.ext.docx.desc') },
            { key: 'doc',   name: t('settings.fv.ext.doc'),   desc: t('settings.fv.ext.doc.desc') },
            { key: 'html',  name: t('settings.fv.ext.html'),  desc: t('settings.fv.ext.html.desc') },
            { key: 'image', name: t('settings.fv.ext.image'), desc: t('settings.fv.ext.image.desc') },
            { key: 'video', name: t('settings.fv.ext.video'), desc: t('settings.fv.ext.video.desc') },
            { key: 'office',name: t('settings.fv.ext.office'),desc: t('settings.fv.ext.office.desc') },
            { key: 'text',  name: t('settings.fv.ext.text'),  desc: t('settings.fv.ext.text.desc') }
        ];

        fvGroups.forEach(function(g) {
            new Setting(body)
                .setName(g.name)
                .setDesc(g.desc)
                .addToggle(function(t2) {
                    t2.setValue(fvSettings[g.key] !== false)
                        .onChange(async function(v) {
                            self.plugin.settings.fileViewerExtensions[g.key] = v;
                            await self.plugin.saveSettings();
                        });
                });
        });
    }

    _renderUtilityToggles(containerEl) {
        var self = this;
        var collapsedState = this.plugin.settings.sectionCollapsed || {};
        var isCollapsed = collapsedState['utility'] === true;

        var categoryEl = containerEl.createDiv({ cls: 'v15-category' });
        if (!isCollapsed) categoryEl.classList.add('open');

        var header = categoryEl.createDiv({ cls: 'v15-cat-header' });
        header.createSpan({ text: '⚡', cls: 'v15-cat-icon' });
        header.createSpan({ text: t('settings.utility'), cls: 'v15-cat-label' });
        header.createSpan({ text: '7' + '', cls: 'v15-cat-count' });
        var arrow = header.createSpan({ text: '▶', cls: 'v15-cat-arrow' });

        header.addEventListener('click', function() {
            categoryEl.classList.toggle('open');
            if (!self.plugin.settings.sectionCollapsed) self.plugin.settings.sectionCollapsed = {};
            self.plugin.settings.sectionCollapsed['utility'] = !categoryEl.classList.contains('open');
            self.plugin.saveSettings();
        });

        var body = categoryEl.createDiv({ cls: 'v15-cat-body' });

        var utilityModules = [
            { id: 'autoplay-loop',     name: t('settings.ut.autoplay'),        desc: t('settings.ut.autoplay.desc') },
            { id: 'folder-counter',    name: t('settings.ut.folderCounter'),   desc: t('settings.ut.folderCounter.desc') },
            { id: 'excel-to-markdown', name: t('settings.ut.excelToMd'),       desc: t('settings.ut.excelToMd.desc') },
            { id: 'img-gallery',       name: t('settings.ut.imgGallery'),      desc: t('settings.ut.imgGallery.desc') },
            { id: 'memories',          name: t('settings.ut.memories'),        desc: t('settings.ut.memories.desc') },
            { id: 'table-resize',      name: t('settings.ut.tableResize'),     desc: t('settings.ut.tableResize.desc') },
            { id: 'vault-stats',       name: t('settings.ut.vaultStats'),      desc: t('settings.ut.vaultStats.desc') },
            { id: 'image-tools',       name: t('settings.ut.imageTools'),      desc: t('settings.ut.imageTools.desc') },
            { id: 'mobile-adapter',    name: t('settings.ut.mobileAdapter'),   desc: t('settings.ut.mobileAdapter.desc') }
        ];

        // 有独立配置面板的 utility 模块（有 renderSettings()）
        var configurableUtility = ['mobile-adapter'];

        utilityModules.forEach(function(um) {
            var modSettings = self.plugin.settings.modules[um.id] || {};
            var setting = new Setting(body)
                .setName(um.name)
                .setDesc(um.desc)
                .addToggle(function(t2) {
                    t2.setValue(modSettings.enabled !== false)
                        .onChange(async function(v) {
                            if (!self.plugin.settings.modules[um.id]) {
                                self.plugin.settings.modules[um.id] = { enabled: true };
                            }
                            self.plugin.settings.modules[um.id].enabled = v;
                            await self.plugin.saveSettings();
                            self.plugin.refreshView();
                        });
                });

            // 对有独立配置的模块额外添加"配置"按钮
            if (configurableUtility.indexOf(um.id) !== -1) {
                var umId = um.id;
                setting.addButton(function(b) {
                    b.setButtonText(t('settings.configureBtn'))
                        .onClick(function() {
                            var newId = self._currentModuleId === umId ? null : umId;

                            // 收拢旧的
                            if (self._currentModuleId && self._currentModuleId !== umId) {
                                var oldEl = body.querySelector('[data-inline-settings="' + self._currentModuleId + '"]');
                                if (oldEl) { oldEl.style.display = 'none'; oldEl.empty(); }
                            }

                            self._currentModuleId = newId;

                            // 展开或收拢当前
                            var target = body.querySelector('[data-inline-settings="' + umId + '"]');
                            if (newId && target) {
                                target.style.display = 'block';
                                target.empty();
                                self._renderModuleSettings(target, newId);
                                target.scrollIntoView({ block: 'nearest', behavior: 'instant' });
                            } else if (target) {
                                target.style.display = 'none';
                                target.empty();
                            }
                        });
                });

                // 内联隐藏容器（手风琴展开区域）
                body.createDiv({
                    cls: 'v15-inline-settings-container',
                    attr: { 'data-inline-settings': um.id, style: 'display:none;' }
                });
            }
        });
    }

    _renderInstanceManager(containerEl) {
        const instances = this.plugin.settings.instances || [];
        if (instances.length === 0) return;

        containerEl.createEl('h3', { text: t('settings.instanceManage') });

        containerEl.createEl('p', {
            text: t('settings.instanceDesc'),
            attr: { style: 'color: var(--text-muted); font-size: 12px; margin-bottom: 12px;' }
        });

        instances.forEach(inst => {
            const baseMod = this.plugin.moduleManager.getModule(inst.baseModule);
            const modSettings = this.plugin.settings.modules[inst.id] || {};

            new Setting(containerEl)
                .setName((baseMod ? baseMod.icon : '📦') + ' ' + inst.label)
                .setDesc(t('settings.instanceType') + (baseMod ? baseMod.title : inst.baseModule) + ' (ID: ' + inst.id + ')')
                .addToggle(t2 => {
                    t2.setValue(modSettings.enabled !== false)
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
                    b.setButtonText(t('settings.configureBtn'))
                        .onClick(() => {
                            var newId = this._currentModuleId === inst.id ? null : inst.id;

                            if (this._currentModuleId && this._currentModuleId !== inst.id) {
                                var oldEl = containerEl.querySelector('[data-inline-settings="' + this._currentModuleId + '"]');
                                if (oldEl) { oldEl.style.display = 'none'; oldEl.empty(); }
                            }

                            this._currentModuleId = newId;

                            var target = containerEl.querySelector('[data-inline-settings="' + inst.id + '"]');
                            if (newId && target) {
                                target.style.display = 'block';
                                target.empty();
                                this._renderModuleSettings(target, newId);
                                target.scrollIntoView({ block: 'nearest', behavior: 'instant' });
                            } else if (target) {
                                target.style.display = 'none';
                                target.empty();
                            }
                        });
                })
                .addButton(b => {
                    b.setButtonText(t('settings.instanceDelete')).setWarning()
                        .onClick(async () => {
                            instances.splice(instances.indexOf(inst), 1);
                            this.plugin.settings.instances = instances;
                            delete this.plugin.settings.modules[inst.id];
                            delete this.plugin.settings.layout[inst.id];
                            if (this.plugin.settings.moduleOrder) {
                                this.plugin.settings.moduleOrder = this.plugin.settings.moduleOrder.filter(id2 => id2 !== inst.id);
                            }
                            if (this._currentModuleId === inst.id) this._currentModuleId = null;
                            await this.plugin.saveSettings();
                            this.plugin.refreshView();
                            new Notice(t('settings.instanceDeleted') + inst.label);
                            this.display();
                        });
                });
                // 行内隐藏容器
                containerEl.createDiv({
                    cls: 'v15-inline-settings-container',
                    attr: { 'data-inline-settings': inst.id, style: 'display:none;' }
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
        wrapper.createEl('h3', { text: t('settings.moduleSettings') + displayTitle + t('settings.moduleSettingsSuffix') });

        const saveCallback = async () => {
            await this.plugin.saveSettings();
        };

        try {
            mod.renderSettings(wrapper, this.plugin, saveCallback);
        } catch (e) {
            console.error('[V17] 模块 ' + moduleId + ' 设置渲染失败:', e);
            wrapper.createEl('p', { text: t('settings.moduleSettingsFailed') + e.message, attr: { style: 'color: var(--text-muted);' } });
        }
    }

    _renderDonateSection(containerEl) {
        const section = containerEl.createDiv({
            attr: {
                style: 'margin-top:32px;padding-top:20px;border-top:2px dashed var(--background-modifier-border);'
            }
        });

        section.createEl('h3', {
            text: t('settings.donate'),
            attr: { style: 'text-align:center;margin-bottom:12px;' }
        });

        // 三栏布局：爱发电 | PayPal+Ko-fi | 微信二维码
        const donateRow = section.createDiv({
            attr: {
                style: 'display:flex;justify-content:center;align-items:stretch;gap:24px;flex-wrap:wrap;margin-top:8px;'
            }
        });

        // 打赏二维码（图床外链）
        const qrSrc = 'https://img-reg-ab.imagency.cn/e/19467f4b916c082ee6ef3b9d81aa9ecb.png';

        // 第一栏：中文打赏（爱发电）
        const cnBox = donateRow.createDiv({
            attr: {
                style: 'flex:0 0 auto;text-align:center;background:var(--background-secondary);border-radius:12px;padding:16px;border:1px solid var(--background-modifier-border);display:flex;flex-direction:column;justify-content:center;align-items:center;min-width:200px;min-height:274px;'
            }
        });
        cnBox.createEl('div', {
            text: t('settings.donateWeChat'),
            attr: { style: 'font-size:14px;font-weight:600;margin-bottom:10px;color:var(--text-normal);' }
        });
        cnBox.createEl('div', {
            text: 'ifdian.net',
            attr: { style: 'font-size:40px;font-weight:700;color:var(--interactive-accent);margin-bottom:12px;' }
        });
        cnBox.createEl('a', {
            text: t('settings.donateCoffee'),
            attr: {
                href: 'https://ifdian.net/a/liamzy2021',
                target: '_blank',
                rel: 'noopener',
                style: 'display:inline-block;padding:10px 24px;background:var(--interactive-accent);color:var(--text-on-accent);border-radius:8px;text-decoration:none;font-size:15px;font-weight:600;cursor:pointer;'
            }
        });

        // 第二栏：PayPal + Ko-fi（国际打赏）
        const intlBox = donateRow.createDiv({
            attr: {
                style: 'flex:0 0 auto;text-align:center;background:var(--background-secondary);border-radius:12px;padding:16px;border:1px solid var(--background-modifier-border);display:flex;flex-direction:column;justify-content:center;align-items:center;min-width:200px;min-height:274px;gap:12px;'
            }
        });
        intlBox.createEl('div', {
            text: 'Support',
            attr: { style: 'font-size:14px;font-weight:600;margin-bottom:6px;color:var(--text-normal);' }
        });
        // PayPal 按钮
        intlBox.createEl('a', {
            text: t('settings.donatePayPal'),
            attr: {
                href: 'https://paypal.me/lilirenzy',
                target: '_blank',
                rel: 'noopener',
                style: 'display:inline-block;padding:10px 20px;background:#003087;color:#fff;border-radius:8px;text-decoration:none;font-size:14px;font-weight:600;cursor:pointer;width:160px;box-sizing:border-box;'
            }
        });
        // Ko-fi 按钮
        intlBox.createEl('a', {
            text: t('settings.donateKofi'),
            attr: {
                href: 'https://ko-fi.com/liamzy',
                target: '_blank',
                rel: 'noopener',
                style: 'display:inline-block;padding:10px 20px;background:#FF5E5B;color:#fff;border-radius:8px;text-decoration:none;font-size:14px;font-weight:600;cursor:pointer;width:160px;box-sizing:border-box;'
            }
        });

        // 第三栏：微信赞赏二维码（最后）
        const qrBox = donateRow.createDiv({
            attr: {
                style: 'flex:0 0 auto;text-align:center;background:var(--background-secondary);border-radius:12px;padding:16px;border:1px solid var(--background-modifier-border);display:flex;flex-direction:column;justify-content:center;align-items:center;min-width:200px;min-height:274px;'
            }
        });
        qrBox.createEl('div', {
            text: t('settings.donateWeChatQR'),
            attr: { style: 'font-size:14px;font-weight:600;margin-bottom:10px;color:var(--text-normal);' }
        });
        // 微信赞赏码
        qrBox.createEl('img', {
            attr: {
                src: qrSrc,
                onerror: "this.style.display='none';this.nextSibling.style.display='block';",
                style: 'width:200px;height:200px;object-fit:contain;border-radius:8px;background:#fff;display:block;margin:0 auto;'
            }
        });
        qrBox.createEl('div', {
            text: t('settings.qrLoadFailed'),
            attr: { style: 'display:none;color:var(--text-muted);font-size:12px;' }
        });
    }
}

// ============================================================
// 09-i18n.js — 中英双语国际化系统
// ============================================================
// 语言选项：
//   'zh'      — 中文
//   'en'      — English
//   'system'  — 跟随系统（中文系统→中文，其他→英文）
//   'ai'      — AI 翻译模式（选择目标语言 → 调用 AI 翻译 → 预览 → 保存）
//   'custom_xx' — AI 翻译后保存的自定义语言（如 'custom_ru'）
// ============================================================

// -------------------- 翻译字典 --------------------
const I18N_DICT = {
    // ===== 通用 =====
    'app.dashboard':                   { zh: '仪表盘主页',           en: 'Dashboard' },
    'app.settings':                    { zh: '设置',                 en: 'Settings' },

    // ===== 设置面板标题 =====
    'settings.title':                  { zh: '仪表盘主页 V17 设置',  en: 'Dashboard V17 Settings' },
    'settings.appearance':             { zh: '外观',                 en: 'Appearance' },
    'settings.theme':                  { zh: '主题',                 en: 'Theme' },
    'settings.theme.desc':             { zh: '选择仪表盘主题风格',   en: 'Choose dashboard theme style' },
    'settings.showHeader':             { zh: '显示顶栏',             en: 'Show Top Bar' },
    'settings.showHeader.desc':        { zh: '显示或隐藏仪表盘顶部工具栏', en: 'Show or hide the dashboard top toolbar' },
    'settings.cardBgColor':            { zh: '卡片背景颜色',          en: 'Card Background Color' },
    'settings.cardBgColor.desc':       { zh: '自定义卡片背景色，留空则使用主题默认', en: 'Custom card background color, leave empty for theme default' },
    'settings.cardBgColor.placeholder':{ zh: '#1a1a1a 或 #ffffff',  en: '#1a1a1a or #ffffff' },
    'settings.cardBgOpacity':         { zh: '卡片背景透明度',        en: 'Card Background Opacity' },
    'settings.cardBgOpacity.desc':    { zh: '0 = 完全透明，1 = 完全不透明', en: '0 = fully transparent, 1 = fully opaque' },
    'settings.resetLayout':            { zh: '重置布局',             en: 'Reset Layout' },
    'settings.resetLayout.desc':       { zh: '清除所有模块的位置和尺寸设置，恢复默认布局', en: 'Clear all module position and size settings, restore default layout' },
    'settings.resetBtn':               { zh: '重置',                 en: 'Reset' },
    'settings.layoutReset':            { zh: '布局已重置',           en: 'Layout has been reset' },

    // ===== 语言设置 =====
    'settings.language':               { zh: '语言',                 en: 'Language' },
    'settings.language.desc':          { zh: '选择界面语言',          en: 'Select interface language' },
    'settings.lang.zh':                { zh: '中文',                 en: '中文' },
    'settings.lang.en':                { zh: 'English',              en: 'English' },
    'settings.lang.system':            { zh: '跟随系统',             en: 'Follow System' },
    'settings.lang.ai':                { zh: 'AI 翻译',             en: 'AI Translation' },
    'settings.aiKey':                  { zh: 'AI API Key',           en: 'AI API Key' },
    'settings.aiKey.desc':             { zh: '用于 AI 翻译的 API 密钥（明文显示）', en: 'API key for AI translation (displayed in plain text)' },
    'settings.aiKey.placeholder':      { zh: 'sk-...',              en: 'sk-...' },
    'settings.aiApiUrl':               { zh: 'AI API 地址',          en: 'AI API URL' },
    'settings.aiApiUrl.desc':          { zh: 'AI 翻译接口地址，留空使用 OpenAI 默认地址', en: 'AI translation API URL, leave empty for OpenAI default' },
    'settings.aiApiUrl.placeholder':   { zh: 'https://api.openai.com/v1/chat/completions', en: 'https://api.openai.com/v1/chat/completions' },
    'settings.aiModel':                { zh: 'AI 模型',              en: 'AI Model' },
    'settings.aiModel.desc':           { zh: '用于翻译的模型名称',    en: 'Model name for translation' },
    'settings.aiModel.placeholder':    { zh: 'gpt-3.5-turbo',       en: 'gpt-3.5-turbo' },
    'settings.aiTranslating':          { zh: '正在翻译中...',          en: 'Translating...' },
    'settings.aiTranslateBtn':         { zh: '翻译并保存',            en: 'Translate & Save' },
    'settings.aiTranslateSuccess':     { zh: '翻译完成！已保存自定义翻译', en: 'Translation complete! Custom translations saved.' },
    'settings.aiTranslateFailed':      { zh: '翻译失败：',            en: 'Translation failed: ' },
    'settings.aiTargetLang':           { zh: '目标语言',              en: 'Target Language' },
    'settings.aiTargetLang.desc':      { zh: '选择要翻译成的语言，或选择「自动检测」', en: 'Select target language, or "Auto-detect"' },
    'settings.aiTargetLang.auto':      { zh: '自动检测',              en: 'Auto-detect' },
    'settings.aiCustomLang':           { zh: '自定义语言',            en: 'Custom Language' },
    'settings.aiCustomLang.desc':      { zh: '输入任意语言名称（如：泰语、乌克兰语、斯瓦希里语、广东话、四川话）', en: 'Enter any language name (e.g.: Thai, Ukrainian, Swahili, Cantonese, Sichuanese)' },
    'settings.aiCustomLang.placeholder': { zh: '例：泰语 / 广东话 / 闽南语', en: 'e.g.: Thai / Cantonese / Hokkien' },
    'settings.aiCustomLang.helpTitle': { zh: '支持任意语言和方言',    en: 'Supports any language & dialect' },
    'settings.aiCustomLang.helpLine1':  { zh: '• 使用你喜欢的 AI 模型（OpenAI / DeepSeek / Moonshot 等），将界面翻译成任意语言', en: '• Use any AI model (OpenAI / DeepSeek / Moonshot / ...) to translate into any language' },
    'settings.aiCustomLang.helpLine2':  { zh: '• 支持方言和地方话：广东话、四川话、闽南语、上海话、客家话、潮汕话...', en: '• Supports dialects: Cantonese, Sichuanese, Hokkien, Shanghainese, Hakka, Teochew...' },
    'settings.aiCustomLang.helpLine3':  { zh: '• 只要你的 AI 模型能理解，什么语言都可以：克林贡语、精灵语、世界语、BL 语... 🖖', en: '• If your AI understands it, it works: Klingon, Elvish, Esperanto, ... 🖖' },
    'settings.aiQuickLang':             { zh: '快捷选择',              en: 'Quick Select' },
    'settings.aiQuickLang.desc':        { zh: '从预设语言中选择，或在上面输入框手动填写', en: 'Pick from presets, or type manually above' },
    'settings.aiQuickLang.none':        { zh: '自定义输入',            en: 'Custom input' },
    'settings.aiTranslateBtn.desc':    { zh: '翻译完成后会弹出预览，确认后生效', en: 'Preview will appear after translation, confirm to apply' },
    'settings.aiPreview.title':        { zh: 'AI 翻译预览',          en: 'AI Translation Preview' },
    'settings.aiPreview.translated':   { zh: '已翻译 ',               en: 'Translated ' },
    'settings.aiPreview.translatedSuffix': { zh: ' 个条目 → ',        en: ' entries → ' },
    'settings.aiPreview.apply':        { zh: '应用翻译',              en: 'Apply Translation' },
    'settings.aiPreview.cancel':       { zh: '取消',                  en: 'Cancel' },
    'settings.aiPreview.applySuccess': { zh: '翻译已应用！界面语言已切换为 ', en: 'Translation applied! Interface language switched to ' },

    // ===== 设置存档 =====
    'settings.backup':                 { zh: '📦 设置存档',          en: '📦 Settings Archive' },
    'settings.export':                 { zh: '导出设置',             en: 'Export Settings' },
    'settings.export.desc':            { zh: '将当前所有设置保存为 JSON 文件，升级插件后可导入恢复', en: 'Save all current settings as a JSON file for restoration after plugin upgrade' },
    'settings.exportBtn':              { zh: '导出设置',             en: 'Export Settings' },
    'settings.exportSuccess':          { zh: '✅ 设置已导出',         en: '✅ Settings exported' },
    'settings.import':                 { zh: '导入设置',             en: 'Import Settings' },
    'settings.import.desc':            { zh: '选择之前导出的 JSON 文件恢复设置（会覆盖当前配置）', en: 'Select a previously exported JSON file to restore settings (overwrites current config)' },
    'settings.importBtn':              { zh: '导入设置',             en: 'Import Settings' },
    'settings.importInvalid':          { zh: '❌ 无效的设置文件，缺少必要字段', en: '❌ Invalid settings file, missing required fields' },
    'settings.importSuccess':          { zh: '✅ 设置已导入，请重新打开设置面板查看', en: '✅ Settings imported, please reopen settings panel to view' },
    'settings.importFailed':           { zh: '❌ 导入失败：',         en: '❌ Import failed: ' },

    // ===== 模块管理 =====
    'settings.moduleManage':           { zh: '模块管理',             en: 'Module Manager' },
    'settings.noModules':              { zh: '未找到任何模块文件，请检查 modules/ 目录', en: 'No module files found, please check the modules/ directory' },
    'settings.cat.schedule':           { zh: '日程与任务',            en: 'Schedule & Tasks' },
    'settings.cat.viewers':            { zh: '查看器（点开即用，免配置）', en: 'Viewers (Click to use, no config)' },
    'settings.cat.notes':              { zh: '笔记与写作',            en: 'Notes & Writing' },
    'settings.cat.files':              { zh: '文件与管理',            en: 'Files & Management' },
    'settings.cat.media':              { zh: '图片与媒体',            en: 'Images & Media' },
    'settings.cat.web':                { zh: '网络与信息',            en: 'Web & Information' },
    'settings.cat.ai':                 { zh: 'AI 与导入',            en: 'AI & Import' },
    'settings.cat.other':              { zh: '其他模块',             en: 'Other Modules' },
    'settings.configureBtn':           { zh: '配置',                 en: 'Configure' },

    // ===== 文件查看器 =====
    'settings.fileViewer':             { zh: '文件查看器',            en: 'File Viewer' },
    'settings.fv.ext.xlsx':           { zh: '表格文件',              en: 'Spreadsheet' },
    'settings.fv.ext.xlsx.desc':      { zh: 'XLSX / XLS / CSV / ODS', en: 'XLSX / XLS / CSV / ODS' },
    'settings.fv.ext.docx':           { zh: 'Word 文档',             en: 'Word Document' },
    'settings.fv.ext.docx.desc':      { zh: 'DOCX（mammoth.js 解析）', en: 'DOCX (mammoth.js parsing)' },
    'settings.fv.ext.doc':            { zh: '旧版 Word',             en: 'Legacy Word' },
    'settings.fv.ext.doc.desc':       { zh: 'DOC 97-2003（docstream 解析）', en: 'DOC 97-2003 (docstream parsing)' },
    'settings.fv.ext.html':           { zh: 'HTML 网页',             en: 'HTML Page' },
    'settings.fv.ext.html.desc':      { zh: 'iframe 安全渲染',         en: 'iframe secure rendering' },
    'settings.fv.ext.image':          { zh: '图片预览',              en: 'Image Preview' },
    'settings.fv.ext.image.desc':     { zh: 'PNG / JPG / GIF / SVG / WebP 等', en: 'PNG / JPG / GIF / SVG / WebP etc.' },
    'settings.fv.ext.video':          { zh: '视频音频+PDF',           en: 'Video/Audio+PDF' },
    'settings.fv.ext.video.desc':     { zh: 'MP4 / WebM / MP3 / PDF', en: 'MP4 / WebM / MP3 / PDF' },
    'settings.fv.ext.office':         { zh: 'PPT 演示文稿',          en: 'PowerPoint' },
    'settings.fv.ext.office.desc':    { zh: 'PPT / PPTX（文字模式预览）', en: 'PPT / PPTX (text-only preview)' },
    'settings.fv.ext.text':           { zh: '纯文本/代码',           en: 'Plain Text/Code' },
    'settings.fv.ext.text.desc':      { zh: 'TXT / JSON / XML / JS / PY 等', en: 'TXT / JSON / XML / JS / PY etc.' },

    // ===== 表格查看器工具栏 =====
    'fv.table.decimals':               { zh: '小数位',               en: 'Decimals' },
    'fv.table.auto':                   { zh: '自动',                 en: 'Auto' },

    // ===== 全局功能 =====
    'settings.utility':                { zh: '全局功能',              en: 'Global Features' },
    'settings.ut.autoplay':           { zh: '自动播放',              en: 'Auto Play' },
    'settings.ut.autoplay.desc':      { zh: '开启后，笔记中的 .mp4/.mp3 自动播放', en: 'When enabled, .mp4/.mp3 in notes autoplay' },
    'settings.ut.folderCounter':      { zh: '文件夹计数器',          en: 'Folder Counter' },
    'settings.ut.folderCounter.desc': { zh: '文件浏览器中显示文件夹内文件数量', en: 'Show file count in folder within file explorer' },
    'settings.ut.excelToMd':          { zh: 'Excel 转表格',          en: 'Excel to Table' },
    'settings.ut.excelToMd.desc':     { zh: '粘贴 Excel 内容时自动转为 Markdown 表格', en: 'Auto-convert pasted Excel content to Markdown table' },
    'settings.ut.tableResize':        { zh: '表格列宽调整',          en: 'Table Column Resize' },
    'settings.ut.tableResize.desc':   { zh: 'Markdown 表格支持拖拽调整列宽', en: 'Markdown tables support drag-to-resize columns' },
    'settings.ut.vaultStats':         { zh: '笔记统计',              en: 'Vault Stats' },
    'settings.ut.vaultStats.desc':    { zh: '统计整个知识库的文件数、字数等', en: 'Count files, words, etc. across the entire vault' },
    'settings.ut.imageTools':         { zh: '图片处理',              en: 'Image Tools' },
    'settings.ut.imageTools.desc':    { zh: '图片格式转换、压缩、重命名（右键菜单）', en: 'Image format conversion, compression, rename (right-click menu)' },
    'settings.ut.imgGallery':         { zh: '图片画廊',              en: 'Image Gallery' },
    'settings.ut.imgGallery.desc':    { zh: '通过 ```t 代码块显示图片画廊（水平网格/垂直瀑布流）', en: 'Display image gallery via ```t code block (grid/waterfall)' },
    'settings.ut.memories':           { zh: '媒体画廊',              en: 'Media Gallery' },
    'settings.ut.memories.desc':      { zh: '通过 ```s 代码块显示媒体画廊（图片/视频/音频）', en: 'Display media gallery via ```s code block (images/videos/audio)' },

    'settings.ut.mobileAdapter':      { zh: '移动端适配器',          en: 'Mobile Adapter' },
    'settings.ut.mobileAdapter.desc': { zh: '自检测移动端并启用响应式布局，支持屏幕方向锁定', en: 'Auto-detect mobile and enable responsive layout, supports screen orientation lock' },

    // ===== 实例管理 =====
    'settings.instanceManage':         { zh: '实例管理',              en: 'Instance Manager' },
    'settings.instanceDesc':           { zh: '以下是通过 ➕ 按钮添加的额外板块实例，可在此管理', en: 'Extra dashboard instances added via ➕ button, managed here' },
    'settings.instanceType':           { zh: '类型: ',                en: 'Type: ' },
    'settings.instanceDelete':         { zh: '删除',                  en: 'Delete' },
    'settings.instanceDeleted':        { zh: '已删除: ',              en: 'Deleted: ' },
    'settings.moduleSettings':         { zh: '⚙️ ',                   en: '⚙️ ' },
    'settings.moduleSettingsSuffix':   { zh: ' 设置',                 en: ' Settings' },
    'settings.moduleSettingsFailed':   { zh: '设置加载失败: ',         en: 'Settings load failed: ' },

    // ===== 打赏 =====
    'settings.donate':                 { zh: '☕ 支持开发者',         en: '☕ Support Developer' },
    'settings.donateWeChat':           { zh: '🇨🇳 爱发电',             en: '🇨🇳 Afdian' },
    'settings.donateCoffee':           { zh: '☕ 请喝杯咖啡',          en: '☕ Buy me a coffee' },
    'settings.donateWeChatQR':         { zh: '🇨🇳 微信赞赏',          en: '🇨🇳 WeChat Tip' },
    'settings.donateGitHub':           { zh: '⭐ 关注我的项目',       en: '⭐ Star & Follow' },
    'settings.donatePayPal':           { zh: '💳 PayPal 捐赠',        en: '💳 Donate via PayPal' },
    'settings.donateKofi':             { zh: '☕ Ko-fi 支持',          en: '☕ Buy me a coffee' },
    'settings.qrLoadFailed':           { zh: '二维码加载失败',        en: 'QR code load failed' },

    // ===== 仪表盘视图 =====
    'dashboard.loading':               { zh: '⏳ 加载中...',          en: '⏳ Loading...' },
    'dashboard.loadingProgress':       { zh: '⏳ ',                   en: '⏳ ' },
    'dashboard.ready':                 { zh: '✅ 就绪',              en: '✅ Ready' },
    'dashboard.refresh':               { zh: '刷新',                 en: 'Refresh' },
    'dashboard.removeInstance':        { zh: '移除此板块',            en: 'Remove this section' },
    'dashboard.addSection':            { zh: '添加板块',              en: 'Add Section' },
    'dashboard.addSectionTitle':       { zh: '添加板块（所有模块均可添加多个）', en: 'Add Section (all modules support multiple instances)' },
    'dashboard.smartSort':            { zh: '智能排序 — 保留每个模块的当前尺寸，紧凑排列。放大/缩小的模块会自动找到合适的位置', en: 'Smart Sort — preserves each module\'s current size, compact arrangement' },
    'dashboard.switchTheme':           { zh: '切换主题',              en: 'Switch Theme' },
    'dashboard.settings':              { zh: '设置',                 en: 'Settings' },
    'dashboard.moduleNotLoaded':       { zh: '模块 ',                en: 'Module ' },
    'dashboard.moduleNotLoaded2':      { zh: ' 未加载',              en: ' not loaded' },
    'dashboard.added':                 { zh: '已添加: ',              en: 'Added: ' },
    'dashboard.removed':               { zh: '已移除: ',              en: 'Removed: ' },
    'dashboard.nothingToSort':         { zh: '没有可排序的模块',       en: 'No modules to sort' },
    'dashboard.sorted':                { zh: '已货架排列 ',           en: 'Shelf-sorted ' },
    'dashboard.sortedSuffix':          { zh: ' 个模块（行列对齐，保留各自尺寸）', en: ' modules (grid-aligned, sizes preserved)' },

    // ===== 模块通用 =====
    'mod.settings':                    { zh: ' 设置',                 en: ' Settings' },
    'mod.weather':                     { zh: '天气',                  en: 'Weather' },
    'mod.calendar':                    { zh: '日历',                  en: 'Calendar' },
    'mod.todo':                        { zh: '待办事项',              en: 'To-Do' },
    'mod.stats':                       { zh: '笔记统计',              en: 'Note Stats' },
    'mod.recent':                      { zh: '最近文件',              en: 'Recent Files' },
    'mod.aiInsight':                  { zh: 'AI洞察',               en: 'AI Insight' },
    'mod.autoplay':                    { zh: '自动播放',              en: 'Auto Play' },
    'mod.folderCounter':               { zh: '文件夹统计',            en: 'Folder Counter' },
    'mod.directory':                   { zh: '目录',                  en: 'Directory' },
    'mod.imageTools':                  { zh: '图片处理',              en: 'Image Tools' },
    'mod.excelToMd':                   { zh: 'Excel转表格',          en: 'Excel to Table' },
    'mod.imageGallery':                { zh: '图片画廊',              en: 'Image Gallery' },
    'mod.docViewer':                   { zh: '文档查看器',            en: 'Document Viewer' },
    'mod.spreadsheet':                 { zh: '表格查看器',            en: 'Spreadsheet Viewer' },
    'mod.codeEditor':                  { zh: '代码编辑器',            en: 'Code Editor' },
    'mod.dataEditor':                  { zh: '数据编辑器',            en: 'Data Editor' },
    'mod.mobileAdapter':               { zh: '移动端适配器',          en: 'Mobile Adapter' },
    'mod.mobileAdapter.title':         { zh: '移动端适配器',          en: 'Mobile Adapter' },
    'mod.mobileAdapter.mobile':        { zh: '📱 移动端',             en: '📱 Mobile' },
    'mod.mobileAdapter.desktop':       { zh: '🖥️ 桌面端',             en: '🖥️ Desktop' },
    'mod.mobileAdapter.mainSwitch':    { zh: '自适应布局',            en: 'Adaptive Layout' },
    'mod.mobileAdapter.adaptLayout':   { zh: '响应式网格（覆盖拖拽位置）', en: 'Responsive grid (overrides drag positions)' },
    'mod.mobileAdapter.colSettings':   { zh: '列数设置',              en: 'Column Settings' },
    'mod.mobileAdapter.portraitCols':  { zh: '竖屏列数',              en: 'Portrait Columns' },
    'mod.mobileAdapter.landscapeCols': { zh: '横屏列数',              en: 'Landscape Columns' },
    'mod.mobileAdapter.orientationLock': { zh: '屏幕方向锁定',        en: 'Orientation Lock' },
    'mod.mobileAdapter.lockMode':      { zh: '锁定模式',              en: 'Lock Mode' },
    'mod.mobileAdapter.lock.natural':  { zh: '🔄 跟随重力',           en: '🔄 Follow Gravity' },
    'mod.mobileAdapter.lock.portrait': { zh: '📱 强制竖屏',           en: '📱 Force Portrait' },
    'mod.mobileAdapter.lock.landscape':{ zh: '📺 强制横屏',           en: '📺 Force Landscape' },
    'mod.mobileAdapter.status.engine': { zh: '引擎：',                en: 'Engine: ' },
    'mod.mobileAdapter.status.running':{ zh: '运行中 ✅',             en: 'Running ✅' },
    'mod.mobileAdapter.status.stopped':{ zh: '已停止',                en: 'Stopped' },
    'mod.mobileAdapter.status.screen': { zh: '屏幕：',                en: 'Screen: ' },
    'mod.mobileAdapter.status.orient': { zh: '方向：',                en: 'Orientation: ' },
    'mod.mobileAdapter.status.portrait': { zh: '竖屏',               en: 'Portrait' },
    'mod.mobileAdapter.status.landscape':{ zh: '横屏',               en: 'Landscape' },
    'mod.mobileAdapter.status.cols':   { zh: '当前列数：',            en: 'Columns: ' },
    'mod.mobileAdapter.settings.title':{ zh: '移动端适配设置',        en: 'Mobile Adapter Settings' },
    'mod.mobileAdapter.settings.desc': { zh: '自动检测移动端并启用响应式布局，支持屏幕方向锁定。移动端默认开启，桌面端默认关闭。', en: 'Auto-detect mobile and enable responsive layout with orientation lock. Default ON for mobile, OFF for desktop.' },

    // ===== v12 设置面板（renderSettings）=====
    'mod.mobileAdapter.setup.mobileActive':  { zh: '📱 当前为移动端，适配已自动启用。修改参数后保存生效。', en: '📱 Mobile detected, adapter auto-enabled. Changes apply after save.' },
    'mod.mobileAdapter.setup.desktopPreview':{ zh: '🖥️ 桌面端。此处修改会在手机打开时生效。', en: '🖥️ Desktop. Changes here will apply when opened on mobile.' },
    'mod.mobileAdapter.setup.portraitTitle': { zh: '📱 竖屏设置（竖拿手机）', en: '📱 Portrait (phone upright)' },
    'mod.mobileAdapter.setup.landscapeTitle':{ zh: '🔄 横屏设置（横拿手机）', en: '🔄 Landscape (phone sideways)' },
    'mod.mobileAdapter.setup.generalTitle':  { zh: '⚙️ 通用', en: '⚙️ General' },
    'mod.mobileAdapter.setup.cols':          { zh: '列数', en: 'Columns' },
    'mod.mobileAdapter.setup.colsDesc':      { zh: '默认 ', en: 'Default ' },
    'mod.mobileAdapter.setup.cardWidth':     { zh: '卡片宽度 (px)', en: 'Card Width (px)' },
    'mod.mobileAdapter.setup.cardWidthDesc': { zh: '0=自动等分, 默认 180', en: '0=Auto divide, default 180' },
    'mod.mobileAdapter.setup.cardHeight':    { zh: '卡片高度 (px)', en: 'Card Height (px)' },
    'mod.mobileAdapter.setup.cardHeightDesc':{ zh: '0=用 vh 比例, 默认 0', en: '0=Use vh ratio, default 0' },
    'mod.mobileAdapter.setup.cardHeightVh':  { zh: '卡片高度 (vh)', en: 'Card Height (vh)' },
    'mod.mobileAdapter.setup.cardHeightVhDesc':{ zh: 'px=0 时生效, 默认 28', en: 'Active when px=0, default 28' },
    'mod.mobileAdapter.setup.contentScale':  { zh: '内容缩放比例', en: 'Content Scale' },
    'mod.mobileAdapter.setup.contentScaleDesc':{ zh: '0.4-1.0, 默认 0.75。越小内容越紧凑', en: '0.4-1.0, default 0.75. Smaller = more compact' },
    'mod.mobileAdapter.setup.applyBtn':      { zh: '立即应用', en: 'Apply Now' },
    'mod.mobileAdapter.setup.applied':       { zh: '✅ 已应用', en: '✅ Applied' },
    'mod.mobileAdapter.setup.hiddenTitle':   { zh: '🙈 已隐藏模块', en: '🙈 Hidden Modules' },
    'mod.mobileAdapter.setup.hiddenEmpty':   { zh: '暂无隐藏模块。在手机端长按卡片即可隐藏。', en: 'No hidden modules. Long-press a card on mobile to hide it.' },
    'mod.mobileAdapter.setup.restoreBtn':    { zh: '恢复', en: 'Restore' },
    'mod.mobileAdapter.setup.tip':           { zh: '💡 尺寸建议：宽度 0 = 自动按列等分；填具体 px 则固定。内容缩放 0.7-0.8 适合大多数场景。', en: '💡 Tip: Width 0 = auto divide; enter px for fixed size. Scale 0.7-0.8 works for most.' },

    // ===== toggle + label =====
    'mod.mobileAdapter.hideHeader':          { zh: '隐藏标题栏', en: 'Hide Header' },
    'mod.mobileAdapter.hideHeaderDesc':      { zh: '不显示卡片标题，节省手机空间', en: 'Hide card headers to save screen space' },

    'mod.mobileAdapter.hideNavbar':          { zh: '隐藏顶部菜单', en: 'Hide Top Menu' },
    'mod.mobileAdapter.hideNavbarDesc':      { zh: '隐藏移动端顶部标题和按钮（底部导航栏不受影响）', en: 'Hide top header title & buttons on mobile (bottom bar unaffected)' },
    'mod.mobileAdapter.status.mobileEnabled':  { zh: '📱 移动端（已启用）', en: '📱 Mobile (enabled)' },
    'mod.mobileAdapter.status.desktopDisabled':{ zh: '🖥️ 桌面端（手机打开时自动启用）', en: '🖥️ Desktop (auto-enable on mobile)' },

    // ===== 长按隐藏按钮 =====
    'mod.mobileAdapter.hideModuleBtn':          { zh: '隐藏此模块', en: 'Hide This Module' },

    'mod.htmlViewer':                  { zh: 'HTML查看器',           en: 'HTML Viewer' },
    'mod.urlOpener':                   { zh: '网址导航',              en: 'URL Navigator' },
    'mod.mediaGallery':                { zh: '媒体画廊',              en: 'Media Gallery' },
    'mod.vaultStats':                  { zh: '笔记统计',              en: 'Vault Stats' },
    'mod.tableResize':                 { zh: '表格列宽',              en: 'Table Resize' },
    'mod.xhsImporter':                 { zh: '小红书导入',            en: 'XHS Importer' },
    'mod.webPreview':                   { zh: '网页预览',              en: 'Web Preview' },
    'mod.webVideo':                    { zh: '网页视频',              en: 'Web Video' },
    'mod.news':                        { zh: '新闻',                  en: 'News' },
    'mod.news.settings.title':         { zh: '资讯模块设置',           en: 'News Module Settings' },
    'mod.news.settings.default_feed':  { zh: '默认订阅源',            en: 'Default Feed' },
    'mod.news.settings.default_feed_desc': { zh: '打开时默认显示的 RSS 分类', en: 'Default RSS category to show on open' },
    'mod.news.settings.custom_feeds':  { zh: '自定义 RSS 源',          en: 'Custom RSS Feeds' },
    'mod.news.settings.custom_feeds_desc': { zh: '添加你自己的 RSS 订阅，格式：中文名 | English Name | URL', en: 'Add your own RSS feeds. Format: Chinese Label | English Label | URL' },
    'mod.news.settings.zh_label':      { zh: '中文名',                en: 'ZH Label' },
    'mod.news.settings.en_label':      { zh: '英文名',                en: 'EN Label' },
    'mod.news.settings.add_feed':      { zh: '+ 添加 RSS 源',         en: '+ Add RSS Feed' },
    'mod.news.title':                  { zh: '资讯',                  en: 'News' },
    'mod.pptViewer':                   { zh: 'PPT演示文稿',           en: 'PPT Viewer' },

    // ===== 天气模块 =====
    'mod.weather.error.cityNotFound':  { zh: '城市未找到，请检查城市名称或 API Key', en: 'City not found, please check city name or API Key' },
    'mod.weather.error.fetchFailed':   { zh: '实时天气获取失败: ',     en: 'Failed to fetch weather: ' },
    'mod.weather.error.unknown':       { zh: '未知',                   en: 'Unknown' },
    'mod.weather.windUnit':            { zh: '风',                    en: '' },
    'mod.weather.scaleUnit':           { zh: '级',                    en: '' },
    'mod.weather.turnTo':              { zh: '转',                    en: ' → ' },
    'mod.weather.error.noKey':         { zh: '请先在模块设置中填写高德地图 API Key', en: 'Please set Amap API Key in module settings first' },
    'mod.weather.freeApply':          { zh: '👉 免费申请',            en: '👉 Apply for free' },
    'mod.weather.owmFreeApply':       { zh: '👉 免费申请',            en: '👉 Apply for free' },
    'mod.weather.settings.owmNeedKeyFallback': { zh: '请先在模块设置中填写 OpenWeatherMap API Key', en: 'Please set OpenWeatherMap API Key in module settings first' },
    'mod.weather.humidity':            { zh: '湿度',                  en: 'Humidity' },
    'mod.weather.windDirection':      { zh: '风向',                  en: 'Wind Dir' },
    'mod.weather.windScale':          { zh: '风力',                  en: 'Wind Scale' },
    'mod.weather.forecast':            { zh: '未来预报',              en: 'Forecast' },
    'mod.weather.tomorrow':            { zh: '明天',                  en: 'Tomorrow' },
    'mod.weather.dayAfter':            { zh: '后天',                  en: 'Day After' },
    'mod.weather.loadFailed':         { zh: '天气加载失败',           en: 'Weather load failed' },
    'mod.weather.retry':              { zh: '点击重试',              en: 'Click to retry' },
    'mod.weather.settings.title':      { zh: '天气模块设置',           en: 'Weather Module Settings' },
    'mod.weather.settings.city':       { zh: '城市',                  en: 'City' },
    'mod.weather.settings.cityDesc':   { zh: '输入城市名称（如：北京、上海、深圳）', en: 'Enter city name (e.g.: Beijing, Shanghai, Shenzhen)' },
    'mod.weather.settings.apiKey':     { zh: '高德地图 API Key',       en: 'Amap API Key' },
    'mod.weather.settings.apiKeyDesc': { zh: '免费申请：https://lbs.amap.com/', en: 'Apply free: https://lbs.amap.com/' },
    // --- V16 多平台新增 ---
    'mod.weather.settings.provider':       { zh: '天气数据源',            en: 'Weather Provider' },
    'mod.weather.settings.providerDesc':   { zh: '选择天气数据的获取平台',   en: 'Select weather data source platform' },
    'mod.weather.settings.providerHint':   { zh: '当前平台: ',             en: 'Current provider: ' },
    'mod.weather.settings.owmApiKey':      { zh: 'OpenWeatherMap API Key',  en: 'OpenWeatherMap API Key' },
    'mod.weather.settings.owmApiKeyDesc':  { zh: '免费申请：https://openweathermap.org/api (1000次/天)', en: 'Apply free: https://openweathermap.org/api (1000/day)' },
    'mod.weather.settings.customUrl':      { zh: '自定义 API 地址',        en: 'Custom API URL' },
    'mod.weather.settings.customUrlDesc':  { zh: '使用 {city} 作为城市名占位符，返回 JSON 格式数据。支持多种常见 JSON 结构自动解析。', en: 'Use {city} as placeholder for city name. Returns JSON data with auto-detection of common formats.' },
    // 平台名称
    'mod.weather.provider.amap':          { zh: '高德地图 Amap (中国)',     en: 'Amap (China)' },
    'mod.weather.provider.openmeteo':      { zh: 'Open-Meteo (全球·推荐)',  en: 'Open-Meteo (Global·Recommended)' },
    'mod.weather.provider.wttr':           { zh: 'wttr.in (极简)',          en: 'wttr.in (Simple)' },
    'mod.weather.provider.owm':            { zh: 'OpenWeatherMap (国际)',   en: 'OpenWeatherMap (International)' },
    'mod.weather.provider.custom':         { zh: '自定义 URL',             en: 'Custom URL' },
    // 错误提示
    'mod.weather.settings.owmNeedKey':     { zh: '请先在模块设置中填写 OpenWeatherMap API Key', en: 'Please set OpenWeatherMap API Key in module settings first' },
    'mod.weather.settings.customNeedUrl':  { zh: '请填写自定义天气 API 的 URL 模板（{city} 占位符）', en: 'Please enter custom weather API URL template ({city} placeholder)' },
    // 帮助文案
    'mod.weather.help.amap':              { zh: '高德地图天气服务，覆盖中国全境，数据精准。需要注册并申请免费的 Web 服务 API Key。', en: 'Amap weather service, covers all of China with precise data. Requires free API key registration.' },
    'mod.weather.help.openmeteo':          { zh: '开源免费天气 API，整合了19国国家气象局模型（含中国 CMA），全球覆盖。无需 API Key，无调用次数限制。', en: 'Open-source free weather API, powered by 19 national weather services incl. China. No API key needed, no rate limits.' },
    'mod.weather.help.wttr':              { zh: '极简天气查询服务，直接用城市名即可。无需注册、无需 API Key。数据来源 WorldWeatherOnline。', en: 'Ultra-simple weather service, just use city name directly. No registration or key needed.' },
    'mod.weather.help.owm':               { zh: '国际主流天气 API，支持多语言返回。需免费注册申请 API Key，免费版限制 1000 次/天。', en: 'Popular international weather API, supports multiple languages. Free registration required, 1000 requests/day on free tier.' },
    'mod.weather.help.custom':             { zh: '使用你自己的天气 API 接口。URL 中使用 {city} 作为城市名占位符，系统会自动尝试解析 JSON 返回数据。', en: 'Use your own weather API endpoint. Use {city} as city name placeholder. JSON response will be auto-parsed.' },

    // ===== 日历模块 =====
    'mod.calendar.settings.title':     { zh: '日历模块设置',           en: 'Calendar Module Settings' },
    'mod.calendar.settings.lunar':     { zh: '显示农历',              en: 'Show Lunar Calendar' },
    'mod.calendar.settings.lunarDesc': { zh: '在每天下方显示农历日期',   en: 'Show lunar date below each day' },
    'mod.calendar.settings.holiday':   { zh: '显示节日/节气',         en: 'Show Festivals/Solar Terms' },
    'mod.calendar.settings.holidayDesc': { zh: '在节日和节气当天显示标注', en: 'Show markers on festival and solar term days' },

    // ===== 待办模块 =====
    'mod.todo.placeholder':            { zh: '添加新待办，按 Enter 确认...', en: 'Add new to-do, press Enter to confirm...' },
    'mod.todo.error.readFailed':       { zh: '读取失败: ',             en: 'Read failed: ' },
    'mod.todo.filter.all':             { zh: '全部',                  en: 'All' },
    'mod.todo.filter.todo':            { zh: '待办',                  en: 'To-Do' },
    'mod.todo.filter.done':            { zh: '完成',                  en: 'Done' },
    'mod.todo.progress':              { zh: '已完成 ',                en: 'Completed ' },
    'mod.todo.progressSuffix':         { zh: '，还剩 ',               en: ', ' },
    'mod.todo.progressSuffix2':        { zh: ' 项',                   en: ' remaining' },
    'mod.todo.empty.today':            { zh: '今天没有待办，加油！',   en: 'No to-dos for today, keep going!' },
    'mod.todo.empty.done':             { zh: '还没有完成的事项',        en: 'No completed items yet' },
    'mod.todo.settings.title':         { zh: '待办模块设置',           en: 'To-Do Module Settings' },
    'mod.todo.settings.folder':        { zh: '待办文件夹',             en: 'To-Do Folder' },
    'mod.todo.settings.folderDesc':    { zh: '存放待办 Markdown 文件的文件夹路径（相对于 Vault 根目录）', en: 'Folder path for to-do Markdown files (relative to Vault root)' },

    // ===== 笔记统计模块 =====
    'mod.stats.loading':               { zh: '⏳ 统计中...',          en: '⏳ Counting...' },
    'mod.stats.error.loadFailed':      { zh: '加载失败: ',             en: 'Load failed: ' },
    'mod.stats.totalNotes':            { zh: '📄 笔记总数',           en: '📄 Total Notes' },
    'mod.stats.totalWords':            { zh: '✏️ 总字数',             en: '✏️ Total Words' },
    'mod.stats.folders':               { zh: '📁 文件夹',             en: '📁 Folders' },
    'mod.stats.avgWords':              { zh: '📊 平均字数',           en: '📊 Avg Words' },
    'mod.stats.folderRank':            { zh: '📂 文件夹排行',         en: '📂 Folder Ranking' },
    'mod.stats.settings.title':        { zh: '统计模块设置',           en: 'Stats Module Settings' },
    'mod.stats.settings.showCount':    { zh: '显示笔记数量',           en: 'Show note count' },
    'mod.stats.settings.showWords':    { zh: '显示字数统计',           en: 'Show word count' },

    // ===== 最近文件模块 =====
    'mod.recent.justNow':              { zh: '刚刚',                  en: 'Just now' },
    'mod.recent.minutesAgo':           { zh: '分钟前',                en: ' min ago' },
    'mod.recent.hoursAgo':             { zh: '小时前',                en: ' hours ago' },
    'mod.recent.yesterday':            { zh: '昨天',                  en: 'Yesterday' },
    'mod.recent.daysAgo':              { zh: '天前',                 en: ' days ago' },
    'mod.recent.empty':                { zh: '暂无文件',              en: 'No files' },
    'mod.recent.rootDir':              { zh: '根目录',                en: 'Root' },
    'mod.recent.error.loadFailed':     { zh: '加载失败: ',             en: 'Load failed: ' },
    'mod.recent.settings.title':       { zh: '最近文件设置',           en: 'Recent Files Settings' },
    'mod.recent.settings.maxFiles':    { zh: '显示数量',              en: 'Display Count' },
    'mod.recent.settings.maxFilesDesc': { zh: '最多显示多少个最近修改的文件', en: 'Maximum number of recently modified files to display' },

    // ===== AI洞察模块 =====
    'mod.ai.error.noKey':              { zh: '请先在模块设置中配置 AI API 密钥', en: 'Please configure AI API key in module settings first' },
    'mod.ai.btn.analyze':              { zh: '🔍 分析最近笔记',        en: '🔍 Analyze Recent Notes' },
    'mod.ai.btn.clearCache':           { zh: '🗑️ 清除缓存',           en: '🗑️ Clear Cache' },
    'mod.ai.lastAnalysis':            { zh: '上次分析：',             en: 'Last analysis: ' },
    'mod.ai.analysisAt':              { zh: '分析于：',               en: 'Analyzed at: ' },
    'mod.ai.analyzing':               { zh: '🤔 正在分析笔记内容，请稍候...', en: '🤔 Analyzing notes, please wait...' },
    'mod.ai.empty':                    { zh: '暂无笔记可分析',          en: 'No notes to analyze' },
    'mod.ai.cacheCleared':            { zh: '缓存已清除，点击「分析最近笔记」重新分析', en: 'Cache cleared, click "Analyze Recent Notes" to re-analyze' },
    'mod.ai.needKey':                 { zh: '⚙️ 请先在模块设置中填写 AI API 密钥，再点击「分析最近笔记」', en: '⚙️ Please fill in AI API key in module settings first, then click "Analyze Recent Notes"' },
    'mod.ai.settings.title':           { zh: 'AI洞察模块设置',         en: 'AI Insight Module Settings' },
    'mod.ai.settings.apiKey':          { zh: 'API Key',               en: 'API Key' },
    'mod.ai.settings.apiKeyDesc':     { zh: 'OpenAI 或兼容接口的 API 密钥（明文显示）', en: 'OpenAI or compatible API key (displayed in plain text)' },
    'mod.ai.settings.apiUrl':          { zh: 'API URL',               en: 'API URL' },
    'mod.ai.settings.apiUrlDesc':     { zh: '留空使用 OpenAI 默认地址；使用其他兼容接口（如 deepseek、moonshot）请填入对应地址', en: 'Leave empty for OpenAI default; enter URL for compatible APIs (e.g. deepseek, moonshot)' },
    'mod.ai.settings.model':           { zh: '模型',                 en: 'Model' },
    'mod.ai.settings.modelDesc':      { zh: '选择或输入模型名称',      en: 'Select or enter model name' },
    'mod.ai.settings.temperature':     { zh: '温度',                 en: 'Temperature' },
    'mod.ai.settings.temperatureDesc': { zh: '越低越保守（0.0），越高越创意（1.0）', en: 'Lower = more conservative (0.0), higher = more creative (1.0)' },
    'mod.ai.settings.requestDelay':    { zh: '请求延迟',              en: 'Request Delay' },
    'mod.ai.settings.requestDelayDesc': { zh: '在此实例触发 AI 请求前的额外等待时间（秒），用于错开多个 AI 板块的并发请求', en: 'Extra wait time (seconds) before triggering AI request, to stagger concurrent AI requests' },
    'mod.ai.settings.globalInterval':  { zh: '全局最小间隔',          en: 'Global Min Interval' },
    'mod.ai.settings.globalIntervalDesc': { zh: '所有 AI 洞察实例之间的最小请求间隔（毫秒），防止触发 API 频率限制', en: 'Minimum request interval (ms) between all AI insight instances to prevent API rate limiting' },
    // AI 错误提示
    'mod.ai.error.invalidKey':        { zh: 'API 密钥无效，请检查设置',        en: 'Invalid API key, please check settings' },
    'mod.ai.error.invalidUrl':        { zh: 'API 地址无效，请检查 URL',         en: 'Invalid API URL, please check URL' },
    'mod.ai.error.rateLimited':       { zh: '请求频率过高，请稍后再试',         en: 'Rate limited, please try again later' },
    'mod.ai.error.callFailed':       { zh: 'AI 调用失败: ',                  en: 'AI call failed: ' },
    'mod.ai.error.apiError':          { zh: 'API返回错误',                    en: 'API returned error' },
    'mod.ai.error.parseFailed':      { zh: '无法解析 AI 响应格式',          en: 'Cannot parse AI response format' },
    'mod.ai.settings.customOption':   { zh: '自定义...',                      en: 'Custom...' },
    'mod.ai.settings.customModelPlaceholder': { zh: '自定义模型名',            en: 'Custom model name' },

    // ===== 自动播放模块 =====
    'mod.autoplay.title':              { zh: '🔊 媒体自动播放控制',     en: '🔊 Media Autoplay Control' },
    'mod.autoplay.running':            { zh: '● 运行中',              en: '● Running' },
    'mod.autoplay.stopped':            { zh: '○ 已停止',              en: '○ Stopped' },
    'mod.autoplay.on':                 { zh: '开',                   en: 'On' },
    'mod.autoplay.off':                { zh: '关',                   en: 'Off' },
    'mod.autoplay.grp.video':          { zh: '🎬 视频',              en: '🎬 Video' },
    'mod.autoplay.grp.audio':          { zh: '🔊 音频',              en: '🔊 Audio' },
    'mod.autoplay.grp.advanced':       { zh: '⚙ 高级',               en: '⚙ Advanced' },
    'mod.autoplay.opt.autoplay':       { zh: '自动播放',              en: 'Autoplay' },
    'mod.autoplay.opt.automute':       { zh: '自动静音',              en: 'Auto Mute' },
    'mod.autoplay.opt.loop':           { zh: '循环播放',              en: 'Loop' },
    'mod.autoplay.opt.pauseOut':       { zh: '离开视野暂停',          en: 'Pause on Leave View' },
    'mod.autoplay.opt.singleAudio':    { zh: '同时只播放一个音频',    en: 'Single Audio Playback' },
    'mod.autoplay.opt.singleVideo':    { zh: '同时只播放一个视频',    en: 'Single Video Playback' },
    'mod.autoplay.status':             { zh: '当前页面: ',             en: 'Current page: ' },
    'mod.autoplay.statusVideo':        { zh: ' 个视频, ',             en: ' videos, ' },
    'mod.autoplay.statusAudio':        { zh: ' 个音频 | ',            en: ' audio | ' },
    'mod.autoplay.engineRunning':      { zh: '引擎: 运行中',          en: 'Engine: Running' },
    'mod.autoplay.engineStopped':      { zh: '引擎: 已停止',          en: 'Engine: Stopped' },
    'mod.autoplay.settings.title':     { zh: '自动播放设置',          en: 'Autoplay Settings' },
    'mod.autoplay.settings.desc':      { zh: '使用 IntersectionObserver 检测页面上的 video/audio 元素，自动控制播放、静音、循环和离开视野暂停。点击总开关启用引擎，设置项即时生效。', en: 'Uses IntersectionObserver to detect video/audio elements, auto-controls playback, mute, loop, and pause-on-leave. Toggle the main switch to enable.' },

    // ===== 文件夹计数器 =====
    'mod.folderCounter.title':         { zh: '📁 文件夹笔记统计',      en: '📁 Folder Note Stats' },
    'mod.folderCounter.refresh':       { zh: '🔄 刷新',              en: '🔄 Refresh' },
    'mod.folderCounter.empty':         { zh: '📭 库中没有文件夹',      en: '📭 No folders in vault' },
    'mod.folderCounter.total':         { zh: '共 ',                  en: ' ' },
    'mod.folderCounter.totalFolders':   { zh: ' 个文件夹, ',           en: ' folders, ' },
    'mod.folderCounter.totalNotes':    { zh: ' 篇笔记',               en: ' notes' },
    'mod.folderCounter.noteCount':     { zh: ' 篇',                   en: ' notes' },
    'mod.folderCounter.limit20':       { zh: '限制只展示前20个子文件夹避免卡顿', en: 'Limited to top 20 subfolders to avoid lag' },
    'mod.folderCounter.more':          { zh: '... 还有 ',             en: '... ' },
    'mod.folderCounter.moreSuffix':    { zh: ' 个子文件夹',           en: ' more subfolders' },
    'mod.folderCounter.settings.title': { zh: '文件夹统计设置',        en: 'Folder Counter Settings' },
    'mod.folderCounter.settings.desc': { zh: '递归统计库中所有文件夹的笔记数量，顶级文件夹默认展开，点击箭头查看子文件夹统计。点击文件夹可在文件浏览器中定位。', en: 'Recursively count notes in all folders. Top-level folders expanded by default. Click folders to locate in file explorer.' },

    // ===== 目录模块 =====
    'mod.directory.empty':             { zh: '📁 请在设置中添加文件夹路径', en: '📁 Please add folder paths in settings' },
    'mod.directory.error.notFound':    { zh: '文件夹不存在: ',         en: 'Folder not found: ' },
    'mod.directory.fileCount':         { zh: ' 个文件',                en: ' files' },
    'mod.directory.settings.title':    { zh: '目录模块设置',           en: 'Directory Module Settings' },
    'mod.directory.settings.desc':     { zh: '添加 Vault 中的文件夹路径（相对路径，如：笔记/日记）', en: 'Add folder paths in the vault (relative path, e.g.: Notes/Diary)' },

    // ===== 图片处理模块 =====
    'mod.imageTools.format':           { zh: '格式:',                 en: 'Format:' },
    'mod.imageTools.width':            { zh: '宽度:',                 en: 'Width:' },
    'mod.imageTools.quality':          { zh: '质量%:',                en: 'Quality%:' },
    'mod.imageTools.dropHint':         { zh: '拖放图片到下方区域，自动按设置转换格式和尺寸', en: 'Drop images below to auto-convert format and size per settings' },
    'mod.imageTools.dropZone':         { zh: '📥 拖放图片到此处 / 或点击选择文件', en: '📥 Drop images here / or click to select files' },
    'mod.imageTools.processing':       { zh: '处理中...',              en: 'Processing...' },
    'mod.imageTools.done':             { zh: '处理完成: ',             en: 'Done: ' },
    'mod.imageTools.success':          { zh: ' 成功, ',               en: ' success, ' },
    'mod.imageTools.failed':           { zh: ' 失败',                 en: ' failed' },
    'mod.imageTools.saved':            { zh: '图片处理完成: ',         en: 'Image processing done: ' },
    'mod.imageTools.savedSuffix':      { zh: ' 张已保存到库根目录',     en: ' images saved to vault root' },
    'mod.imageTools.error.convert':    { zh: '转换失败',              en: 'Conversion failed' },
    'mod.imageTools.error.load':       { zh: '图片加载失败',          en: 'Image load failed' },
    'mod.imageTools.settings.title':   { zh: '图片处理设置',          en: 'Image Tools Settings' },
    'mod.imageTools.settings.desc':    { zh: '拖放图片到模块面板，自动按设置的格式、宽度和质量进行转换。处理后的图片保存到库根目录（避免覆盖原文件）。', en: 'Drop images to auto-convert per format/width/quality settings. Processed images saved to vault root.' },
    'mod.imageTools.settings.hint':    { zh: '💡 WebP格式体积最小，JPEG兼容性最好，PNG适合需要透明度的图片', en: '💡 WebP is smallest, JPEG has best compatibility, PNG is best for transparency' },

    // ===== Excel转表格模块 =====
    'mod.excelToMd.autoOn':            { zh: '全局自动转换：开启',     en: 'Global auto-convert: ON' },
    'mod.excelToMd.autoOff':           { zh: '全局自动转换：关闭',     en: 'Global auto-convert: OFF' },
    'mod.excelToMd.placeholder':       { zh: '在此 Ctrl+V 粘贴Excel数据（也可直接在笔记中粘贴，会自动转换）...', en: 'Ctrl+V paste Excel data here (or paste directly in notes for auto-convert)...' },
    'mod.excelToMd.btn.insert':        { zh: '📝 插入当前笔记',        en: '📝 Insert to Current Note' },
    'mod.excelToMd.btn.copy':          { zh: '📋 复制Markdown',        en: '📋 Copy Markdown' },
    'mod.excelToMd.btn.clear':         { zh: '🗑 清空',               en: '🗑 Clear' },
    'mod.excelToMd.autoConvertHint':   { zh: '开启全局自动转换后，在任意笔记中粘贴Excel表格数据即可自动转换。', en: 'With global auto-convert on, paste Excel data in any note for auto-conversion.' },
    'mod.excelToMd.waiting':           { zh: '等待粘贴Excel数据...',    en: 'Waiting for Excel data paste...' },
    'mod.excelToMd.detected':          { zh: '✓ 已识别为表格数据（',   en: '✓ Table data detected (' },
    'mod.excelToMd.detectedRows':      { zh: '行）',                  en: ' rows)' },
    'mod.excelToMd.notDetected':       { zh: '❌ 未能识别为Excel表格数据...', en: '❌ Could not detect as Excel table data...' },
    'mod.excelToMd.inserted':          { zh: '✓ 已插入当前笔记！',     en: '✓ Inserted to current note!' },
    'mod.excelToMd.copied':            { zh: '✓ 已复制到剪贴板！',     en: '✓ Copied to clipboard!' },
    'mod.excelToMd.nothingToCopy':     { zh: '请先在面板中粘贴Excel数据', en: 'Please paste Excel data in the panel first' },
    'mod.excelToMd.nothingToCopy2':    { zh: '没有可复制的内容',       en: 'Nothing to copy' },
    'mod.excelToMd.settings.title':    { zh: 'Excel转表格 设置',      en: 'Excel to Table Settings' },
    'mod.excelToMd.settings.auto':     { zh: '启用全局自动转换',        en: 'Enable Global Auto-Convert' },
    'mod.excelToMd.settings.autoDesc': { zh: '开启后，在任意笔记中 Ctrl+V 粘贴 Excel 数据时自动转为 Markdown 表格。关闭后仅支持在模块面板内手动粘贴转换。', en: 'When enabled, Ctrl+V paste Excel data in any note auto-converts to Markdown table. When off, only manual paste in panel works.' },
    'mod.excelToMd.settings.hint1':    { zh: '💡 提示：支持列对齐语法（表头加 ^c 居中, ^r 右对齐, ^l 左对齐）', en: '💡 Tip: Supports column alignment syntax (^c center, ^r right, ^l left)' },
    'mod.excelToMd.settings.hint2':    { zh: '📋 使用方式：在 WPS/Excel 中选中表格 → Ctrl+C → 在 Obsidian 笔记中 Ctrl+V → 自动转换！', en: '📋 Usage: Select table in WPS/Excel → Ctrl+C → Ctrl+V in Obsidian note → auto-convert!' },
    'mod.excelToMd.copyFailed':       { zh: '复制失败，请手动选中上方文本复制', en: 'Copy failed, please manually select the text above to copy' },
    'mod.excelToMd.settings.desc':    { zh: '启用全局自动转换后，在任意笔记中 Ctrl+V 粘贴 Excel 数据时自动转为 Markdown 表格。关闭后仅支持在模块面板内手动粘贴转换。', en: 'When enabled, Ctrl+V paste Excel data in any note auto-converts to Markdown table. When off, only manual paste in panel works.' },

    // ===== 图片画廊模块 =====
    'mod.imageGallery.display':        { zh: '展示:',                 en: 'Display:' },
    'mod.imageGallery.folder':         { zh: '文件夹:',               en: 'Folder:' },
    'mod.imageGallery.cols':           { zh: '列数:',                 en: 'Columns:' },
    'mod.imageGallery.sort':           { zh: '排序:',                 en: 'Sort:' },
    'mod.imageGallery.mode.square':    { zh: '正方形',                en: 'Square' },
    'mod.imageGallery.mode.masonry':   { zh: '瀑布流',                en: 'Masonry' },
    'mod.imageGallery.mode.smart':     { zh: '智能',                  en: 'Smart' },
    'mod.imageGallery.folderHint':     { zh: '请输入照片路径',          en: 'Enter photo folder path' },
    'mod.imageGallery.sort.mtime':     { zh: '修改时间',              en: 'Modified Time' },
    'mod.imageGallery.sort.name':      { zh: '文件名',                en: 'File Name' },
    'mod.imageGallery.sort.size':      { zh: '文件大小',              en: 'File Size' },
    'mod.imageGallery.sort.desc':      { zh: '↓ 降序',               en: '↓ Descending' },
    'mod.imageGallery.sort.asc':       { zh: '↑ 升序',               en: '↑ Ascending' },
    'mod.imageGallery.refresh':        { zh: '🔄 刷新',              en: '🔄 Refresh' },
    'mod.imageGallery.count':          { zh: ' 张',                   en: ' images' },
    'mod.imageGallery.empty.folder':   { zh: '📭 文件夹 "',            en: '📭 No images in folder "' },
    'mod.imageGallery.empty.folder2':  { zh: '"中没有图片',            en: '"' },
    'mod.imageGallery.empty.vault':    { zh: '📭 库中没有图片文件',    en: '📭 No image files in vault' },
    'mod.imageGallery.formats':        { zh: '支持 PNG, JPG, JPEG, GIF, WebP, BMP, SVG', en: 'Supports PNG, JPG, JPEG, GIF, WebP, BMP, SVG' },
    'mod.imageGallery.settings.title': { zh: '图片画廊设置',          en: 'Image Gallery Settings' },
    'mod.imageGallery.settings.desc':  { zh: '扫描指定文件夹中的图片，支持三种展示模式：正方形网格、瀑布流（保持原始比例）、智能自适应。点击图片打开灯箱，支持键盘导航（← → ESC）、鼠标滚轮切换。', en: 'Scan images in specified folder, supports 3 display modes. Click to open lightbox with keyboard navigation.' },
    'mod.imageGallery.showCount':      { zh: '显示图片计数条',         en: 'Show image count bar' },
    'mod.imageGallery.showCount.desc': { zh: '在工具栏右侧显示「N张图片」计数', en: 'Show "N images" counter on the right side of toolbar' },

    // ===== 文档查看器 =====
    'mod.docViewer.btn.files':         { zh: '📂 文件列表',           en: '📂 File List' },
    'mod.docViewer.btn.refresh':       { zh: '🔄 刷新',              en: '🔄 Refresh' },
    'mod.docViewer.btn.view':          { zh: '📖 查看选中',           en: '📖 View Selected' },
    'mod.docViewer.hint':              { zh: '📄 选择文件后点击"查看选中"', en: '📄 Select a file then click "View Selected"' },
    'mod.docViewer.hintFormats':       { zh: '支持: .docx, .pdf',      en: 'Supports: .docx, .pdf' },
    'mod.docViewer.empty':             { zh: '📭 库中没有文档文件（.docx / .pdf）', en: '📭 No document files in vault (.docx / .pdf)' },
    'mod.docViewer.loading':           { zh: '⏳ 加载中...',          en: '⏳ Loading...' },
    'mod.docViewer.selected':          { zh: '已选中: ',              en: 'Selected: ' },
    'mod.docViewer.notFound':          { zh: '文件未找到: ',          en: 'File not found: ' },
    'mod.docViewer.error.mammoth':     { zh: 'Word解析库(mammoth.js)加载失败', en: 'Word parsing library (mammoth.js) load failed' },
    'mod.docViewer.error.network':     { zh: '请检查网络连接后刷新重试', en: 'Please check network connection and retry' },
    'mod.docViewer.error.corrupt':     { zh: '该 DOCX 文件内部引用可能损坏', en: 'This DOCX file may have internal reference corruption' },
    'mod.docViewer.error.corruptHint': { zh: '建议用 WPS 另存为新格式', en: 'Try re-saving with WPS' },
    'mod.docViewer.pdfReader':         { zh: 'PDF 阅读器',            en: 'PDF Reader' },
    'mod.docViewer.docxRendered':      { zh: 'DOCX 已渲染',           en: 'DOCX Rendered' },
    'mod.docViewer.error.load':        { zh: '加载失败: ',             en: 'Load failed: ' },
    'mod.docViewer.settings.title':    { zh: '文档查看器 设置',       en: 'Document Viewer Settings' },
    'mod.docViewer.settings.desc':     { zh: '自动扫描库中的文档文件（.docx .pdf），在面板内渲染预览。DOCX 通过 mammoth.js 转换保留格式，PDF 通过浏览器内置阅读器渲染。', en: 'Auto-scan document files (.docx .pdf) in vault for in-panel preview. DOCX via mammoth.js, PDF via browser reader.' },
    'mod.docViewer.settings.hint':     { zh: '✅ 单击选中 → 点击"📖 查看选中"/双击 → 面板内渲染（不会调 WPS）', en: '✅ Click to select → click "📖 View Selected"/double-click → renders in panel (won\'t open WPS)' },

    // ===== 表格查看器 =====
    'mod.spreadsheet.btn.files':       { zh: '📂 文件列表',           en: '📂 File List' },
    'mod.spreadsheet.btn.refresh':     { zh: '🔄 刷新',              en: '🔄 Refresh' },
    'mod.spreadsheet.btn.view':        { zh: '📊 查看选中',           en: '📊 View Selected' },
    'mod.spreadsheet.hint':            { zh: '📊 选择文件后点击"查看选中"', en: '📊 Select a file then click "View Selected"' },
    'mod.spreadsheet.hintFormats':     { zh: '支持: .xlsx .xls .csv .ods', en: 'Supports: .xlsx .xls .csv .ods' },
    'mod.spreadsheet.empty':           { zh: '📭 库中没有表格文件',    en: '📭 No spreadsheet files in vault' },
    'mod.spreadsheet.loading':         { zh: '⏳ 加载中...',          en: '⏳ Loading...' },
    'mod.spreadsheet.selected':        { zh: '已选中: ',              en: 'Selected: ' },
    'mod.spreadsheet.notFound':        { zh: '文件未找到: ',          en: 'File not found: ' },
    'mod.spreadsheet.error.xlsx':      { zh: 'XLSX解析库加载失败，请检查网络连接。', en: 'XLSX parsing library load failed, please check network.' },
    'mod.spreadsheet.error.csvOk':     { zh: 'CSV文件可直接查看。',     en: 'CSV files can be viewed directly.' },
    'mod.spreadsheet.emptyTable':      { zh: '表格为空',              en: 'Table is empty' },
    'mod.spreadsheet.dimensions':      { zh: ' 行 × ',                en: ' rows × ' },
    'mod.spreadsheet.dimensions2':     { zh: ' 列',                   en: ' cols' },
    'mod.spreadsheet.settings.title':  { zh: '表格查看器 设置',       en: 'Spreadsheet Viewer Settings' },
    'mod.spreadsheet.settings.desc':   { zh: '自动扫描库中的表格文件（.xlsx .xls .csv .ods），在面板内渲染为HTML表格。支持多工作表切换。点击选中 + 点击"查看选中"加载。', en: 'Auto-scan spreadsheet files in vault, render as HTML tables in panel. Supports multi-sheet switching.' },
    'mod.spreadsheet.settings.hint1':  { zh: '✅ 单击选中文件 → 点击"📊 查看选中" → 面板内渲染（不会调 WPS）', en: '✅ Click to select → click "📊 View Selected" → renders in panel (won\'t open WPS)' },
    'mod.spreadsheet.settings.hint2':  { zh: '💡 也可以双击文件名直接加载', en: '💡 Double-click a file name to load directly' },

    // ===== 代码编辑器 =====
    'mod.codeEditor.btn.files':        { zh: '📂 文件列表',           en: '📂 File List' },
    'mod.codeEditor.btn.new':          { zh: '📝 新建',              en: '📝 New' },
    'mod.codeEditor.btn.save':         { zh: '💾 保存',              en: '💾 Save' },
    'mod.codeEditor.btn.openObs':      { zh: '🔍 Obsidian中打开',     en: '🔍 Open in Obsidian' },
    'mod.codeEditor.hint':             { zh: '选择一个代码文件开始编辑', en: 'Select a code file to start editing' },
    'mod.codeEditor.hintFormats':      { zh: '支持所有文本格式',       en: 'Supports all text formats' },
    'mod.codeEditor.empty':            { zh: '📭 库中没有代码文件',    en: '📭 No code files in vault' },
    'mod.codeEditor.line':             { zh: '行: ',                 en: 'Ln: ' },
    'mod.codeEditor.modified':         { zh: '● 已修改',             en: '● Modified' },
    'mod.codeEditor.saved':            { zh: '已保存',               en: 'Saved' },
    'mod.codeEditor.confirmDiscard':   { zh: '当前文件有未保存的修改，要放弃修改吗？', en: 'Current file has unsaved changes. Discard changes?' },
    'mod.codeEditor.nothingToSave':    { zh: '没有修改需要保存',       en: 'No changes to save' },
    'mod.codeEditor.savedNotice':      { zh: '已保存: ',              en: 'Saved: ' },
    'mod.codeEditor.saveFailed':       { zh: '保存失败: ',            en: 'Save failed: ' },
    'mod.codeEditor.selectFirst':      { zh: '请先选择一个文件',        en: 'Please select a file first' },
    'mod.codeEditor.useObsidianNew':   { zh: '新建文件请使用Obsidian原生功能(右键→新建笔记)', en: 'Please use Obsidian\'s native feature to create new files (right-click → New Note)' },
    'mod.codeEditor.settings.title':   { zh: '代码编辑器设置',         en: 'Code Editor Settings' },
    'mod.codeEditor.settings.desc':    { zh: '支持编辑库中所有文本格式的代码文件（JavaScript、Python、CSS、HTML、JSON、YAML等）。支持行号显示、Tab缩进、Ctrl+S保存。如需高级IDE功能（语法高亮、自动补全），请点击"Obsidian中打开"使用原生编辑器。', en: 'Edit all text-format code files in vault. Supports line numbers, Tab indent, Ctrl+S save. For advanced IDE features, click "Open in Obsidian".' },

    // ===== 数据编辑器 =====
    'mod.dataEditor.btn.files':        { zh: '📂 文件列表',           en: '📂 File List' },
    'mod.dataEditor.btn.format':       { zh: '🔧 格式化',            en: '🔧 Format' },
    'mod.dataEditor.btn.validate':     { zh: '✅ 验证',              en: '✅ Validate' },
    'mod.dataEditor.btn.copy':         { zh: '📋 复制',              en: '📋 Copy' },
    'mod.dataEditor.btn.refresh':      { zh: '🔄 刷新',              en: '🔄 Refresh' },
    'mod.dataEditor.hint':             { zh: '选择一个数据文件查看',    en: 'Select a data file to view' },
    'mod.dataEditor.hintFormats':      { zh: 'JSON / YAML / XML / TXT', en: 'JSON / YAML / XML / TXT' },
    'mod.dataEditor.empty':            { zh: '📭 库中没有数据文件',    en: '📭 No data files in vault' },
    'mod.dataEditor.notFound':         { zh: '文件不存在',            en: 'File not found' },
    'mod.dataEditor.readFailed':       { zh: '读取失败: ',            en: 'Read failed: ' },
    'mod.dataEditor.jsonError':        { zh: 'JSON格式错误: ',        en: 'JSON format error: ' },
    'mod.dataEditor.yamlNotSupport':   { zh: 'YAML格式化暂不支持，请使用原样查看', en: 'YAML formatting not supported, please view as-is' },
    'mod.dataEditor.typeNotSupport':   { zh: '该文件类型暂不支持格式化', en: 'This file type does not support formatting' },
    'mod.dataEditor.jsonValid':        { zh: '✓ JSON有效',           en: '✓ JSON valid' },
    'mod.dataEditor.jsonInvalid':      { zh: '✗ JSON无效: ',         en: '✗ JSON invalid: ' },
    'mod.dataEditor.copied':           { zh: '已复制到剪贴板',        en: 'Copied to clipboard' },
    'mod.dataEditor.copyFailed':       { zh: '复制失败',              en: 'Copy failed' },
    'mod.dataEditor.autoSaveFailed':   { zh: '自动保存失败',          en: 'Auto-save failed' },
    'mod.dataEditor.settings.title':   { zh: '数据编辑器设置',         en: 'Data Editor Settings' },
    'mod.dataEditor.settings.desc':    { zh: '浏览和编辑库中的JSON/YAML/XML/TXT等数据文件。支持JSON格式化保存、语法高亮、JSON有效性验证。', en: 'Browse and edit JSON/YAML/XML/TXT data files. Supports JSON formatting, syntax highlighting, and JSON validation.' },

    // ===== HTML查看器 =====
    'mod.htmlViewer.btn.files':        { zh: '📂 文件列表',           en: '📂 File List' },
    'mod.htmlViewer.btn.mode':         { zh: '模式:',                en: 'Mode:' },
    'mod.htmlViewer.btn.refresh':      { zh: '🔄 刷新',              en: '🔄 Refresh' },
    'mod.htmlViewer.mode.text':        { zh: '纯文本',               en: 'Plain Text' },
    'mod.htmlViewer.mode.safe':        { zh: '安全浏览',             en: 'Safe Browse' },
    'mod.htmlViewer.mode.trust':       { zh: '完全信任',             en: 'Full Trust' },
    'mod.htmlViewer.hint':             { zh: '选择一个HTML文件开始预览', en: 'Select an HTML file to start preview' },
    'mod.htmlViewer.empty':            { zh: '📭 库中没有HTML文件',   en: '📭 No HTML files in vault' },
    'mod.htmlViewer.emptyContent':     { zh: '文件内容为空',          en: 'File content is empty' },
    'mod.htmlViewer.modeLabel':        { zh: '模式: ',               en: 'Mode: ' },
    'mod.htmlViewer.settings.title':   { zh: 'HTML查看器设置',        en: 'HTML Viewer Settings' },
    'mod.htmlViewer.settings.mode':    { zh: '模式',                 en: 'Mode' },
    'mod.htmlViewer.settings.desc':    { zh: '说明',                 en: 'Description' },
    'mod.htmlViewer.settings.text':    { zh: '显示HTML源代码，最安全', en: 'Show HTML source code, safest' },
    'mod.htmlViewer.settings.safe':    { zh: '在iframe中渲染，移除脚本和事件，保留样式和图片', en: 'Render in iframe, remove scripts/events, keep styles and images' },
    'mod.htmlViewer.settings.trust':   { zh: '完整渲染，允许脚本执行（仅用于可信HTML文件）', en: 'Full rendering, allows script execution (for trusted HTML only)' },
    'mod.htmlViewer.settings.hint':    { zh: '💡 模式可在模块面板中随时切换。库中的HTML文件会自动扫描并可在文件列表中选取预览。', en: '💡 Mode can be switched anytime in panel. HTML files auto-scanned for preview.' },
    'mod.htmlViewer.error.load':       { zh: '⚠ 无法读取文件: ',         en: '⚠ Cannot read file: ' },

    // ===== 网址导航 =====
    'mod.urlOpener.placeholder':       { zh: '输入网址...',            en: 'Enter URL...' },
    'mod.urlOpener.btn.go':            { zh: '前往',                  en: 'Go' },
    'mod.urlOpener.btn.bookmark':      { zh: '⭐ 收藏',              en: '⭐ Bookmark' },
    'mod.urlOpener.btn.back':          { zh: '后退',                  en: 'Back' },
    'mod.urlOpener.btn.forward':       { zh: '前进',                  en: 'Forward' },
    'mod.urlOpener.btn.refresh':       { zh: '刷新',                  en: 'Refresh' },
    'mod.urlOpener.btn.external':      { zh: '↗ 外部',               en: '↗ External' },
    'mod.urlOpener.btn.bookmarks':     { zh: '📑 书签',              en: '📑 Bookmarks' },
    'mod.urlOpener.hint':              { zh: '输入网址或选择书签开始浏览', en: 'Enter URL or select a bookmark to start' },
    'mod.urlOpener.hintXFrame':        { zh: '部分网站可能因X-Frame-Options限制无法在iframe中显示', en: 'Some sites may not display in iframe due to X-Frame-Options' },
    'mod.urlOpener.noBookmarks':       { zh: '暂无书签，浏览网页时点击⭐收藏', en: 'No bookmarks yet, click ⭐ while browsing to bookmark' },
    'mod.urlOpener.bookmarkExists':    { zh: '该书签已存在',           en: 'Bookmark already exists' },
    'mod.urlOpener.bookmarkAdded':     { zh: '书签已添加',             en: 'Bookmark added' },
    'mod.urlOpener.settings.title':    { zh: '网址导航设置',           en: 'URL Navigator Settings' },
    'mod.urlOpener.settings.desc':     { zh: '在仪表盘中输入网址，使用内嵌浏览器查看。支持前进/后退/刷新，可收藏常用网址。注意：部分网站因X-Frame-Options策略无法在iframe中显示（如百度、淘宝等）。', en: 'Enter URLs in dashboard for embedded browsing. Supports navigation and bookmarks. Some sites may not display in iframe.' },
    'mod.urlOpener.settings.hint':     { zh: '💡 如遇到无法显示的网站，可点击"↗ 外部"按钮在系统默认浏览器中打开。', en: '💡 For sites that won\'t display, click "↗ External" to open in system browser.' },

    // ===== 媒体画廊 =====
    'mod.mediaGallery.dropZone':             { zh: '拖放媒体文件到此处上传',  en: 'Drop media files here to upload' },
    'mod.mediaGallery.settings.title':       { zh: '媒体画廊设置',            en: 'Media Gallery Settings' },
    'mod.mediaGallery.showUploadZone':       { zh: '显示拖放上传区',          en: 'Show upload drop zone' },
    'mod.mediaGallery.showUploadZone.desc':  { zh: '在画廊顶部显示文件拖放上传区域', en: 'Show drag-and-drop upload zone at the top of the gallery' },

    // ===== 仓库统计 =====
    'mod.vaultStats.title':            { zh: '📈 仓库统计',           en: '📈 Vault Stats' },
    'mod.vaultStats.refresh':          { zh: '刷新',                  en: 'Refresh' },
    'mod.vaultStats.loading':          { zh: '正在统计中...',          en: 'Counting...' },
    'mod.vaultStats.totalNotes':       { zh: '📄 笔记总数',           en: '📄 Total Notes' },
    'mod.vaultStats.totalWords':       { zh: '✏️ 总字数',             en: '✏️ Total Words' },
    'mod.vaultStats.folderCount':      { zh: '📁 文件夹数',           en: '📁 Folders' },
    'mod.vaultStats.recent7d':         { zh: '🕐 7天内修改',          en: '🕐 Modified in 7d' },
    'mod.vaultStats.estPages':         { zh: '📐 估算页数',           en: '📐 Est. Pages' },
    'mod.vaultStats.vaultSize':        { zh: '💾 仓库大小',           en: '💾 Vault Size' },
    'mod.vaultStats.largestNote':      { zh: '最大的笔记',            en: 'Largest Note' },
    'mod.vaultStats.currentNote':      { zh: '📝 当前笔记',           en: '📝 Current Note' },
    'mod.vaultStats.fileName':         { zh: '📄 文件名',             en: '📄 File Name' },
    'mod.vaultStats.wordCount':        { zh: '✏️ 字数',              en: '✏️ Words' },
    'mod.vaultStats.charCount':        { zh: '🔤 字符数',             en: '🔤 Characters' },
    'mod.vaultStats.lineCount':        { zh: '📏 行数',              en: '📏 Lines' },
    'mod.vaultStats.fileCount':        { zh: ' 篇',                   en: ' files' },
    'mod.vaultStats.settings.title':   { zh: '笔记统计设置',           en: 'Vault Stats Settings' },
    'mod.vaultStats.settings.desc':    { zh: '统计注释内容',           en: 'Count comment content' },
    'mod.vaultStats.settings.pageWords': { zh: '每页字数',             en: 'Words per page' },
    'mod.vaultStats.settings.pageWordsDesc': { zh: '用于估算总页数（默认 300 字/页）', en: 'Used to estimate total pages (default 300 words/page)' },

    // ===== 表格列宽 =====
    'mod.tableResize.title':           { zh: '📐 表格列宽拖拽调整',    en: '📐 Table Column Resize' },
    'mod.tableResize.minWidth':        { zh: '最小列宽 (px)',          en: 'Min Column Width (px)' },
    'mod.tableResize.btn.inject':      { zh: '🔧 注入拖拽手柄',        en: '🔧 Inject Drag Handles' },
    'mod.tableResize.btn.remove':      { zh: '🗑 移除手柄',           en: '🗑 Remove Handles' },
    'mod.tableResize.status':          { zh: '页面表格: ',             en: 'Page tables: ' },
    'mod.tableResize.statusInjected':  { zh: ' | 已注入: ',            en: ' | Injected: ' },
    'mod.tableResize.statusHandles':   { zh: ' | 手柄数: ',            en: ' | Handles: ' },
    'mod.tableResize.injected':        { zh: '已注入拖拽手柄到 ',      en: 'Injected drag handles to ' },
    'mod.tableResize.injectedSuffix':  { zh: ' 个表格',               en: ' tables' },
    'mod.tableResize.removed':         { zh: '已移除所有拖拽手柄',     en: 'All drag handles removed' },
    'mod.tableResize.handleTitle':     { zh: '拖拽调整列宽',           en: 'Drag to resize column' },
    'mod.tableResize.settings.title':  { zh: '表格列宽调整设置',       en: 'Table Resize Settings' },
    'mod.tableResize.settings.desc':   { zh: '点击"注入拖拽手柄"后，阅读模式下所有表格的表头单元格右侧会出现拖拽手柄，可以拖拽调整列宽。调整后的列宽在当前会话中保持，切换页面后需重新注入。', en: 'After clicking "Inject Drag Handles", drag handles appear on table headers in reading mode for column width adjustment.' },
    'mod.tableResize.settings.hint':   { zh: '💡 此功能作用于Obsidian阅读模式下的markdown渲染表格，编辑模式下的表格不受影响。', en: '💡 This affects tables in Obsidian reading mode; editing mode tables are not affected.' },

    // ===== 小红书导入 =====
    'mod.xhsImporter.placeholder':      { zh: '粘贴小红书分享文本或链接...', en: 'Paste XHS share text or link...' },
    'mod.xhsImporter.autoFetch':       { zh: '自动抓取全文+图片（从页面__INITIAL_STATE__解析）', en: 'Auto-fetch full text + images (parsed from page __INITIAL_STATE__)' },
    'mod.xhsImporter.btn.import':      { zh: '📕 导入笔记',           en: '📕 Import Note' },
    'mod.xhsImporter.ready':           { zh: '就绪',                  en: 'Ready' },
    'mod.xhsImporter.error.empty':     { zh: '请粘贴分享内容',         en: 'Please paste share content' },
    'mod.xhsImporter.loading.parse':   { zh: '正在解析...',           en: 'Parsing...' },
    'mod.xhsImporter.error.noLink':    { zh: '未找到有效链接，请粘贴包含链接的分享文本', en: 'No valid link found, please paste share text containing a link' },
    'mod.xhsImporter.loading.fetch':   { zh: '正在获取笔记页面...',    en: 'Fetching note page...' },
    'mod.xhsImporter.loading.data':    { zh: '正在解析笔记数据...',    en: 'Parsing note data...' },
    'mod.xhsImporter.error.fetch':     { zh: '无法获取小红书页面（可能需要登录或网络问题）', en: 'Cannot fetch XHS page (may need login or network issue)' },
    'mod.xhsImporter.error.parse':     { zh: '页面解析失败，可能该笔记需要登录才能查看。请尝试在浏览器中打开后复制分享文本。', en: 'Page parse failed, the note may need login. Try opening in browser and copy share text.' },
    // 设置面板
    'mod.xhsImporter.settings.title':       { zh: '小红书导入设置',       en: 'XHS Import Settings' },
    'mod.xhsImporter.settings.desc':        { zh: '粘贴小红书分享链接导入笔记。自动从页面 __INITIAL_STATE__ 解析结构化内容（标题/正文/图片/话题标签），用 fetch() 下载原图到本地。', en: 'Paste XHS share link to import notes. Auto-parses structured content (title/body/images/tags) from __INITIAL_STATE__, downloads original images via fetch().' },
    'mod.xhsImporter.settings.noteFolder':      { zh: '笔记保存文件夹',       en: 'Note Folder' },
    'mod.xhsImporter.settings.noteFolderDesc':  { zh: '导入的笔记保存到此文件夹', en: 'Save imported notes to this folder' },
    'mod.xhsImporter.settings.imageFolder':     { zh: '图片保存文件夹',       en: 'Image Folder' },
    'mod.xhsImporter.settings.imageFolderDesc': { zh: '下载的图片保存到此文件夹', en: 'Save downloaded images to this folder' },
    'mod.xhsImporter.settings.customTags':      { zh: '自定义标签',           en: 'Custom Tags' },
    'mod.xhsImporter.settings.customTagsDesc':  { zh: '导入笔记时添加的标签（逗号或空格分隔）', en: 'Tags to add when importing (comma or space separated)' },
    'mod.xhsImporter.settings.downloadMedia':       { zh: '下载图片',       en: 'Download Images' },
    'mod.xhsImporter.settings.downloadMediaDesc':   { zh: '导入时自动下载笔记中的图片到本地', en: 'Auto download images in note to local' },
    // 导入流程
    'mod.xhsImporter.success':              { zh: '导入成功！',           en: 'Import successful!' },
    'mod.xhsImporter.imagesDownloaded':     { zh: ' 已下载 ',             en: ' Downloaded ' },
    'mod.xhsImporter.imageCount':           { zh: ' 张图片',              en: ' images' },
    'mod.xhsImporter.openNote':             { zh: '📂 打开笔记',          en: '📂 Open Note' },
    'mod.xhsImporter.error.import':         { zh: '导入失败: ',           en: 'Import failed: ' },
    'mod.xhsImporter.xhsNote':              { zh: '小红书笔记',            en: 'XHS Note' },
    'mod.xhsImporter.unknown':              { zh: '未知',                 en: 'Unknown' },
    'mod.xhsImporter.preview.author':       { zh: '作者: ',               en: 'Author: ' },
    'mod.xhsImporter.preview.images':       { zh: '图像: ',               en: 'Images: ' },
    'mod.xhsImporter.preview.local':        { zh: ' 张（已下载: ',        en: ' (downloaded: ' },
    'mod.xhsImporter.preview.tags':         { zh: '标签: ',               en: 'Tags: ' },
    'mod.xhsImporter.preview.source':       { zh: '来源: ',               en: 'Source: ' },
    'mod.xhsImporter.videoNote':            { zh: '视频',                 en: 'Video' },
    'mod.xhsImporter.imageNote':            { zh: '图片',                 en: 'Image' },
    'mod.xhsImporter.mediaNoteSuffix':      { zh: '笔记，共 ',            en: ' note, ' },
    'mod.xhsImporter.error.cannotParse':    { zh: '📕 无法解析笔记内容，请确认链接有效或尝试在浏览器中打开', en: '📕 Cannot parse note content, please check link is valid or try opening in browser' },
    'mod.xhsImporter.loading.downloading':  { zh: '正在下载 ',             en: 'Downloading ' },
    'mod.xhsImporter.loading.imageCount':   { zh: ' 张图片...',           en: ' images...' },

    // ===== 网页预览/视频 =====
    'mod.webPreview.placeholder':      { zh: '网址...',               en: 'URL...' },
    'mod.webPreview.zoomOut':          { zh: '➖',                    en: '➖' },
    'mod.webPreview.zoomIn':           { zh: '➕',                    en: '➕' },
    'mod.webPreview.refresh':          { zh: '🔄',                    en: '🔄' },
    'mod.webPreview.zoomOutTitle':     { zh: '缩小',                 en: 'Zoom Out' },
    'mod.webPreview.zoomInTitle':      { zh: '放大',                 en: 'Zoom In' },
    'mod.webPreview.refreshTitle':     { zh: '刷新',                 en: 'Refresh' },
    'mod.webPreview.offsetYTitle':     { zh: '向下偏移',              en: 'Offset Down' },
    'mod.webPreview.offsetXTitle':     { zh: '向右偏移',              en: 'Offset Right' },
    'mod.webPreview.settings.url':     { zh: '预览网址',              en: 'Preview URL' },
    'mod.webPreview.settings.urlDesc': { zh: '使用 Electron webview 打开，支持登录态持久化', en: 'Opens with Electron webview, supports persistent login' },
    'mod.webPreview.settings.zoom':    { zh: '默认缩放',              en: 'Default Zoom' },
    'mod.webPreview.settings.zoomDesc': { zh: '初始缩放比例（0.1 ~ 2.0）', en: 'Initial zoom ratio (0.1 ~ 2.0)' },
    'mod.webPreview.settings.posX':   { zh: '水平偏移 (X)',          en: 'Horizontal Offset (X)' },
    'mod.webPreview.settings.posXDesc': { zh: '向右偏移像素值',        en: 'Rightward offset in pixels' },
    'mod.webPreview.settings.posY':   { zh: '垂直偏移 (Y)',          en: 'Vertical Offset (Y)' },
    'mod.webPreview.settings.posYDesc': { zh: '向下偏移像素值',        en: 'Downward offset in pixels' },

    'mod.webVideo.settings.url':     { zh: '视频网址',              en: 'Video URL' },
    'mod.webVideo.settings.urlDesc': { zh: '使用 Electron webview 打开', en: 'Opens with Electron webview' },

    // ===== PPT查看器 =====
    'mod.pptViewer':                  { zh: 'PPT演示文稿',           en: 'PPT Viewer' },
    'mod.pptViewer.btn.files':        { zh: '📂 文件列表',           en: '📂 File List' },
    'mod.pptViewer.btn.refresh':      { zh: '🔄 刷新',              en: '🔄 Refresh' },
    'mod.pptViewer.btn.view':         { zh: '📊 查看选中',           en: '📊 View Selected' },
    'mod.pptViewer.hint':             { zh: '📊 选择文件后点击"查看选中"', en: '📊 Select a file then click "View Selected"' },
    'mod.pptViewer.hintFormats':      { zh: '支持: .pptx, .ppt',      en: 'Supports: .pptx, .ppt' },
    'mod.pptViewer.empty':            { zh: '📭 库中没有PPT文件（.pptx / .ppt）', en: '📭 No PPT files in vault (.pptx / .ppt)' },
    'mod.pptViewer.loading':          { zh: '⏳ 解析中...',         en: '⏳ Parsing...' },
    'mod.pptViewer.selected':         { zh: '已选中: ',              en: 'Selected: ' },
    'mod.pptViewer.notFound':         { zh: '文件未找到: ',          en: 'File not found: ' },
    'mod.pptViewer.error.jszip':      { zh: 'JSZip解压库加载失败，请检查网络连接。', en: 'JSZip library load failed, please check network.' },
    'mod.pptViewer.error.cfb':        { zh: 'CFB解析库加载失败，旧版PPT解析不可用。', en: 'CFB library load failed, legacy PPT parsing unavailable.' },
    'mod.pptViewer.pptxRendered':     { zh: 'PPTX 已渲染（文字模式）',  en: 'PPTX Rendered (text mode)' },
    'mod.pptViewer.pptRendered':      { zh: 'PPT 已渲染（近似文本提取）', en: 'PPT Rendered (approximate text extraction)' },
    'mod.pptViewer.error.load':       { zh: '加载失败: ',             en: 'Load failed: ' },
    'mod.pptViewer.settings.title':   { zh: 'PPT查看器 设置',        en: 'PPT Viewer Settings' },
    'mod.pptViewer.settings.desc':    { zh: '自动扫描库中的PPT文件（.pptx .ppt），在面板内以纯文本模式预览幻灯片内容。PPTX通过JSZip解压+XML解析，PPT通过CFB二进制流提取。仅显示文字，不含图片/形状/动画。', en: 'Auto-scan PPT files (.pptx .ppt) in vault for in-panel text-only preview. PPTX via JSZip+XML parsing, PPT via CFB binary stream extraction. Text only, no images/shapes/animations.' },
    'mod.pptViewer.settings.hint':    { zh: '✅ 单击选中 → 点击"📊 查看选中"/双击 → 面板内渲染（不会调 PowerPoint）', en: '✅ Click to select → click "📊 View Selected"/double-click → renders in panel (won\'t open PowerPoint)' },

    // ===== 水族箱模块 =====
    'mod.aquarium':                    { zh: '水族箱',               en: 'Aquarium' },
    'mod.aquarium.addFish':            { zh: '+ 鱼',                 en: '+ Fish' },
    'mod.aquarium.addBubble':          { zh: '+ 气泡',               en: '+ Bubble' },
    'mod.aquarium.addPlant':           { zh: '+ 植物',               en: '+ Plant' },
    'mod.aquarium.addFood':            { zh: '+ 投食',               en: '+ Food' },
    'mod.aqu.decor':                   { zh: '+ 装饰',               en: '+ Decor' },
    'mod.aquarium.clear':             { zh: '清除',                en: 'Clear' },
    'mod.aquarium.reset':             { zh: '重置',                en: 'Reset' },
    'mod.aquarium.resetConfirm':      { zh: '重置整个水族箱？',       en: 'Reset entire aquarium?' },
    'mod.aquarium.resetData':         { zh: '重置存档数据',         en: 'Reset Saved Data' },
    'mod.aquarium.stats':             { zh: '水族箱统计',           en: 'Stats' },
    'mod.aquarium.fish':              { zh: '鱼',                  en: 'Fish' },
    'mod.aquarium.plants':            { zh: '植物',                en: 'Plants' },
    'mod.aquarium.bubbles':           { zh: '气泡源',              en: 'Bubbles' },
    'mod.aquarium.color':             { zh: '颜色',                en: 'Color' },
    'mod.aquarium.age':               { zh: '年龄',                en: 'Age' },
    'mod.aquarium.fed':              { zh: '喂食次数',            en: 'Fed' },
    'mod.aquarium.min':              { zh: '分钟',                en: 'min' },
    'mod.aquarium.themeLabel':        { zh: '主题',                en: 'Theme' },
    'mod.aquarium.sandLabel':         { zh: '沙底',                en: 'Sand' },
    'mod.aquarium.nightMode':         { zh: '夜间模式',            en: 'Night Mode' },

    // ===== 像素花园模块 =====
    'mod.pixelGarden':                { zh: '像素花园',             en: 'Pixel Garden' },
    'mod.garden.plant':               { zh: '种植',                en: 'Plant' },
    'mod.garden.waterAll':            { zh: '全部浇水',            en: 'Water All' },
    'mod.garden.harvest':             { zh: '收获',                en: 'Harvest' },
    'mod.garden.shop':                { zh: '种子商店',            en: 'Shop' },
    'mod.garden.seedShop':            { zh: '种子商店',            en: 'Seed Shop' },
    'mod.garden.close':               { zh: '关闭',                en: 'Close' },
    'mod.garden.flower':             { zh: '花朵',                en: 'Flower' },
    'mod.garden.tomato':             { zh: '番茄',                en: 'Tomato' },
    'mod.garden.carrot':             { zh: '胡萝卜',              en: 'Carrot' },
    'mod.garden.sunflower':          { zh: '向日葵',              en: 'Sunflower' },
    'mod.garden.cactus':             { zh: '仙人掌',              en: 'Cactus' },
    'mod.garden.mushroom':           { zh: '蘑菇',                en: 'Mushroom' },

    // ===== 电子宠物模块 =====
    'mod.pet':                        { zh: '电子宠物',             en: 'Pet' },
    'mod.pet.feed':                   { zh: '喂食',                en: 'Feed' },
    'mod.pet.play':                   { zh: '玩耍',                en: 'Play' },
    'mod.pet.clean':                  { zh: '清洁',                en: 'Clean' },
    'mod.pet.sleep':                  { zh: '睡眠',                en: 'Sleep' },
    'mod.pet.cat':                    { zh: '猫咪',                en: 'Cat' },
    'mod.pet.dog':                    { zh: '狗狗',                en: 'Dog' },
    'mod.pet.bunny':                  { zh: '兔子',                en: 'Bunny' },
    'mod.pet.bear':                   { zh: '小熊',                en: 'Bear' },
    'mod.pet.fox':                    { zh: '狐狸',                en: 'Fox' },
    'mod.pet.chick':                  { zh: '小鸡',                en: 'Chick' },
    'mod.pet.stageEgg':              { zh: '蛋',                  en: 'Egg' },
    'mod.pet.stageBaby':             { zh: '幼崽',                en: 'Baby' },
    'mod.pet.stageChild':             { zh: '成长',                en: 'Child' },
    'mod.pet.stageAdult':             { zh: '成年',                en: 'Adult' },
    'mod.pet.grewUp':                 { zh: '长大了！',             en: 'Grew up!' },

    // ===== 粒子玩具模块 =====
    'mod.particleToy':                { zh: '粒子玩具',             en: 'Particle Toy' },
    'mod.pt.attract':                 { zh: '吸引',                en: 'Attract' },
    'mod.pt.repel':                   { zh: '排斥',                en: 'Repel' },
    'mod.pt.flow':                    { zh: '流动',                en: 'Flow' },
    'mod.pt.firework':                { zh: '烟花',                en: 'Firework' },
    'mod.pt.snow':                    { zh: '雪花',                en: 'Snow' },
    'mod.pt.galaxy':                  { zh: '星系',                en: 'Galaxy' },
    'mod.pt.rainbow':                 { zh: '彩虹',                en: 'Rainbow' },

    // ===== 农场点击模块 =====
    'mod.farmClicker':                { zh: '农场点击',             en: 'Farm Clicker' },
    'mod.farm.wheat':                 { zh: '小麦',                en: 'Wheat' },
    'mod.farm.carrot':                { zh: '胡萝卜',              en: 'Carrot' },
    'mod.farm.tomato':                { zh: '番茄',                en: 'Tomato' },
    'mod.farm.corn':                  { zh: '玉米',                en: 'Corn' },
    'mod.farm.strawberry':            { zh: '草莓',                en: 'Strawberry' },
    'mod.farm.pumpkin':               { zh: '南瓜',                en: 'Pumpkin' },
    'mod.farm.coins':                 { zh: '金币',                en: 'Coins' },
    'mod.farm.perSec':                { zh: '每秒',                en: '/sec' },
    'mod.farm.harvestAll':            { zh: '全部收获',            en: 'Harvest All' },
    'mod.farm.upgrades':              { zh: '升级',                en: 'Upgrades' },
    'mod.farm.upgradeShop':           { zh: '升级商店',            en: 'Upgrade Shop' },
    'mod.farm.upgClick':              { zh: '点击力',              en: 'Click Power' },
    'mod.farm.upgAuto':               { zh: '自动速度',            en: 'Auto Speed' },
    'mod.farm.upgYield':              { zh: '产量加成',            en: 'Yield Bonus' },
    'mod.farm.upgGolden':             { zh: '金手',                en: 'Golden Touch' },
    'mod.farm.ready':                 { zh: '可收获!',              en: 'READY!' },
    'mod.farm.growing':               { zh: '生长中...',           en: 'Growing...' },

    // ===== 游戏分类 =====
    'settings.cat.games':             { zh: '游戏娱乐',             en: 'Games & Fun' },

    // ===== 模块管理器 =====
    'mod.manager.notLoaded':           { zh: '模块 "',                en: 'Module "' },
    'mod.manager.notLoaded2':          { zh: '" 未加载',              en: '" not loaded' },
    'mod.manager.renderFailed':        { zh: '渲染失败: ',            en: 'Render failed: ' },

    // ===== 文件查看器（main.js） =====
    'fileViewer.title':                { zh: '文件查看器',              en: 'File Viewer' },
    'fileViewer.openExternal':         { zh: '用外部程序打开 ',          en: 'Open with External App ' },
    'fileViewer.docBadge':             { zh: '旧版 Word 97-2003 (.doc)', en: 'Legacy Word 97-2003 (.doc)' },
    'fileViewer.parsedWith':           { zh: ' — ',                   en: ' — ' },
    'fileViewer.parsedMethod':         { zh: ' 解析',                 en: ' parsed' },
    'fileViewer.author':               { zh: ' | 作者: ',              en: ' | Author: ' },
    'fileViewer.title':                { zh: ' | 标题: ',              en: ' | Title: ' },
    'fileViewer.cannotOpen':           { zh: '无法打开: ',            en: 'Cannot open: ' },
    'fileViewer.cannotOpenFile':       { zh: '无法打开文件: ',        en: 'Cannot open file: ' },
    'fileViewer.readFailed':           { zh: '读取失败: ',             en: 'Read failed: ' },
    'fileViewer.cannotReadFile':       { zh: '无法读取文件: ',         en: 'Cannot read file: ' },
    'fileViewer.cannotOpen':           { zh: '无法打开: ',             en: 'Cannot open: ' },

    // ===== 主文件菜单 =====
    'main.ribbon':                     { zh: '仪表盘主页',            en: 'Dashboard' },
    'main.command':                    { zh: '打开仪表盘主页',         en: 'Open Dashboard' },
};

// -------------------- 语言检测 --------------------
function detectSystemLanguage() {
    try {
        const lang = (navigator.language || navigator.userLanguage || 'en').toLowerCase();
        // 中文系统：zh, zh-cn, zh-tw, zh-hk 等
        return lang.startsWith('zh') ? 'zh' : 'en';
    } catch (e) {
        return 'en';
    }
}

// -------------------- AI 翻译缓存 --------------------
// 结构: { 'zh': { key: translatedText }, 'en': { key: translatedText } }
var _aiTranslationCache = {};
var _aiTranslating = false;

// -------------------- i18n 核心 --------------------
var _currentLang = 'zh'; // 当前语言代码
var _customTranslations = {}; // 用户自定义翻译（如 AI 翻译结果）{ langCode: { key: text } }

function t(key) {
    var entry = I18N_DICT[key];
    if (!entry) return key;

    // 1. 优先使用自定义翻译（AI 翻译保存的结果）
    // _customTranslations key 格式与 _currentLang 一致（如 'custom_ja'）
    if (_customTranslations[_currentLang] && _customTranslations[_currentLang][key]) {
        return _customTranslations[_currentLang][key];
    }

    // 2. 解析实际语言
    var lang = _currentLang;
    var isCustom = lang && lang.startsWith('custom_');
    if (lang === 'system') lang = detectSystemLanguage();
    // ai 模式 → 固定用英文显示设置面板（AI配置界面应保持英文，避免歧义）
    if (lang === 'ai') lang = 'en';
    // custom_xx → 提取 xx 作为显示语言（先查自定义翻译，上面已处理，这里是 fallback）
    if (isCustom) lang = lang.substring(7);

    if (entry[lang]) return entry[lang];
    // 回退：根据语言类型决定 fallback 顺序
    var sys = detectSystemLanguage();
    var preferZh;
    if (isCustom) {
        // 自定义语言：优先英文（通用语言），再中文
        preferZh = false;
    } else {
        // 内置语言：按系统语言偏好
        preferZh = (lang === 'zh') || (lang !== 'en' && sys === 'zh');
    }
    if (preferZh) {
        if (entry.zh) return entry.zh;
        if (entry.en) return entry.en;
    } else {
        if (entry.en) return entry.en;
        if (entry.zh) return entry.zh;
    }
    return key;
}

function getCurrentLang() { return _currentLang; }

function getResolvedLang() {
    if (_currentLang === 'system') return detectSystemLanguage();
    if (_currentLang === 'ai') return 'en'; // AI模式下固定英文显示
    if (_currentLang && _currentLang.startsWith('custom_')) return _currentLang.substring(7);
    return _currentLang;
}

function setLanguage(lang) {
    _currentLang = lang;
}

// 设置自定义翻译（AI 翻译后保存）
// lang 应为完整语言代码（如 'custom_ja'），与 _currentLang 格式一致
function setCustomTranslations(lang, translations) {
    _customTranslations[lang] = translations;
    console.log('[i18n] setCustomTranslations: ' + lang + ' entries=' + Object.keys(translations).length);
}

// -------------------- AI 翻译功能 ====================
async function aiTranslateAll(apiKey, apiUrl, model, targetLang) {
    if (!apiKey) return { success: false, error: 'API Key is required' };
    if (_aiTranslating) return { success: false, error: 'Translation in progress' };
    _aiTranslating = true;

    // 目标语言（用户输入的）
    var langCode = (targetLang || 'auto').trim();
    if (langCode === 'auto' || !langCode) langCode = detectSystemLanguage();

    // 源语言：与系统语言相反（中→英 或 英→中）
    var sysLang = detectSystemLanguage();
    var sourceLang = (sysLang === 'zh') ? 'zh' : 'en';

    console.log('[i18n] aiTranslateAll START: langCode=' + langCode + ' sourceLang=' + sourceLang + ' totalDictKeys=' + Object.keys(I18N_DICT).length);

    try {
        // 收集所有需要翻译的条目（遍历字典）
        var toTranslate = [];
        var allKeys = Object.keys(I18N_DICT);
        for (var i = 0; i < allKeys.length; i++) {
            var key = allKeys[i];
            var entry = I18N_DICT[key];
            if (!entry) continue;

            // 取源语言文本
            var src = entry[sourceLang];
            if (!src) continue;
            // 跳过纯符号/空字符串
            if (typeof src !== 'string' || src.length < 2) continue;

            // 源=目标则跳过
            if (langCode === sourceLang) continue;

            // 已有内置翻译则跳过
            if (entry[langCode]) continue;

            toTranslate.push({ key: key, src: src });
        }

        console.log('[i18n] toTranslate count:', toTranslate.length);

        if (toTranslate.length === 0) {
            _aiTranslating = false;
            return { success: true, total: 0, usedLang: langCode, translations: {} };
        }

        // 分批翻译（每次20条）
        var BATCH_SIZE = 20;
        var allTrans = {};

        for (var b = 0; b < toTranslate.length; b += BATCH_SIZE) {
            var batch = toTranslate.slice(b, b + BATCH_SIZE);
            var prompt = buildPrompt(batch, langCode, sourceLang);

            try {
                var resp = await callAiApi(apiKey, apiUrl, model, prompt);
                if (resp) {
                    console.log('[i18n] batch ' + b + ' raw resp (first 200 chars):', resp.substring(0, 200));
                    // 去掉 Markdown 代码块标记
                    var cleanResp = resp.replace(/```[a-z]*\n?/gi, '').replace(/```/g, '');
                    var lines = cleanResp.split('\n');
                    for (var l = 0; l < lines.length; l++) {
                        var line = lines[l].trim();
                        if (!line) continue;
                        var sep = line.indexOf('|||');
                        if (sep < 0) continue;
                        var k = line.substring(0, sep).trim();
                        var v = line.substring(sep + 3).trim();
                        // 清洗 key：去掉序号（"1. "）、反引号、多余空格
                        k = k.replace(/^\d+[\.\)]\s*/, '').replace(/^`+|`+$/g, '').trim();
                        // 验证 key 是否在字典中
                        if (k && v && I18N_DICT[k]) allTrans[k] = v;
                        else if (k && v) {
                            console.log('[i18n] key not in dict (skipping):', k);
                        }
                    }
                }
            } catch (err) {
                console.warn('[i18n] batch err:', err);
            }
        }

        _aiTranslating = false;
        console.log('[i18n] translation done, count:', Object.keys(allTrans).length);
        return { success: true, total: Object.keys(allTrans).length, usedLang: langCode, translations: allTrans };
    } catch (e) {
        _aiTranslating = false;
        console.error('[i18n] aiTranslateAll error:', e);
        return { success: false, error: e.message || String(e) };
    }
}

function buildPrompt(batch, target, source) {
    // 语言别名（让 AI 更准确理解）
    var alias = {
        en: 'English', zh: 'Simplified Chinese', ja: 'Japanese', ko: 'Korean',
        ru: 'Russian', de: 'German', fr: 'French', es: 'Spanish', pt: 'Portuguese',
        it: 'Italian', ar: 'Arabic', hi: 'Hindi', th: 'Thai', vi: 'Vietnamese',
        id: 'Indonesian', nl: 'Dutch', pl: 'Polish', uk: 'Ukrainian'
    };
    var tName = alias[target] || target;
    var sName = alias[source] || source;

    var lines = [];
    for (var i = 0; i < batch.length; i++) {
        lines.push(batch[i].key + ' ||| ' + batch[i].src);
    }
    return 'Translate the following UI strings from ' + sName + ' to ' + tName
        + '. Return one line per entry in format: key ||| translation. '
        + 'Keep key unchanged. Only output translations, no explanation.\n\n'
        + lines.join('\n');
}

async function callAiApi(apiKey, apiUrl, model, prompt) {
    // 留空时使用 OpenAI 默认地址
    var url = (apiUrl || 'https://api.openai.com/v1/chat/completions').trim();

    // 智能补全：如果用户只填了域名（如 https://api.example.com 或 https://api.example.com/）
    // 自动追加 /v1/chat/completions；如果已有完整路径（如 /v1/... /v3/...）则不修改
    if (apiUrl) {
        try {
            var parsed = new URL(url);
            var path = (parsed.pathname || '/').replace(/\/+$/, '');
            if (path === '' || path === '/') {
                url = url.replace(/\/+$/, '') + '/v1/chat/completions';
            }
        } catch(e) { /* 解析失败，原样使用 */ }
    }

    var mdl = model || 'gpt-3.5-turbo';

    console.log('[i18n] callAiApi url=' + url + ' model=' + mdl);

    var resp = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + apiKey
        },
        body: JSON.stringify({
            model: mdl,
            messages: [
                { role: 'system', content: 'You are a professional translator for software UI.' },
                { role: 'user', content: prompt }
            ],
            temperature: 0.3,
            max_tokens: 8000
        })
    });

    if (!resp.ok) {
        var text = await resp.text();
        throw new Error('API error ' + resp.status + ': ' + text);
    }

    var data = await resp.json();
    if (data.choices && data.choices.length > 0 && data.choices[0].message) {
        return data.choices[0].message.content;
    }
    return null;
}

// -------------------- 导出 --------------------
window.t = t;
window.getCurrentLang = getCurrentLang;
window.getResolvedLang = getResolvedLang;
window.setLanguage = setLanguage;
window.setCustomTranslations = setCustomTranslations;
window.detectSystemLanguage = detectSystemLanguage;
window.aiTranslateAll = aiTranslateAll;
window.I18N_DICT = I18N_DICT;

// ===================== Gallery Processors (img-gallery + memories) =====================
var __GALLERY_BUILTIN__ = true;

(function(global) {
    'use strict';

    var VALID_EXTENSIONS = ['jpeg','jpg','gif','png','webp','tiff','tif','avif','bmp'];

    // ---- i18n ----
    function getLang(plugin) {
        try {
            if (typeof window.getResolvedLang === 'function') {
                var rl = window.getResolvedLang();
                if (rl && rl.startsWith('zh')) return 'zh';
                if (rl) return 'en';
            }
            var lang = plugin && plugin.settings && plugin.settings.language;
            if (!lang || lang === 'system') {
                var nav = (navigator.language || navigator.userLanguage || 'en').toLowerCase();
                return nav.startsWith('zh') ? 'zh' : 'en';
            }
            if (lang === 'ai' || lang.startsWith('custom_')) return 'en';
            return lang.startsWith('zh') ? 'zh' : 'en';
        } catch(e) {}
        return 'en';
    }

    var S = {
        zh: {
            menuEdit:'\u270f\ufe0f \u7f16\u8f91\u4ee3\u7801\u5757',
            menuDelete:'\ud83d\uddd1\ufe0f \u5220\u9664\u4ee3\u7801\u5757',
            menuLayout:'\ud83d\udcd0 \u81ea\u5b9a\u4e49\u6392\u7248',
            menuSize:'\ud83d\udccf \u81ea\u5b9a\u4e49\u5927\u5c0f',
            menuColumns:'\ud83d\udd22 \u81ea\u5b9a\u4e49\u6bcf\u884c\u6570\u91cf',
            titleLayout:'\u81ea\u5b9a\u4e49\u6392\u7248',
            titleSize:'\u81ea\u5b9a\u4e49\u5927\u5c0f',
            titleColumns:'\u6bcf\u884c\u56fe\u7247\u6570\u91cf',
            labelWidth:'\u56fe\u7247\u5bbd\u5ea6 (px\uff0c0=\u81ea\u9002\u5e94):',
            labelHeight:'\u56fe\u7247\u9ad8\u5ea6 (px):',
            labelColumns:'\u6bcf\u884c\u663e\u793a\u56fe\u7247\u6570\u91cf:',
            apply:'\u5e94\u7528',
            cancel:'\u53d6\u6d88',
            perRow:'\u4e2a/\u884c',
            noEditor:'\u672a\u627e\u5230\u7f16\u8f91\u5668',
            noBlock:'\u672a\u627e\u5230\u4ee3\u7801\u5757',
            jumped:'\u5df2\u8df3\u8f6c\u5230\u4ee3\u7801\u5757',
            deleted:'\u5df2\u5220\u9664\u4ee3\u7801\u5757',
            failOp:'\u64cd\u4f5c\u5931\u8d25: ',
            appliedLayout:'\u5df2\u5e94\u7528\u6392\u7248: ',
            appliedSize:'\u5df2\u8bbe\u7f6e\u5927\u5c0f: ',
            widthSuffix:' (\u5bbd: ',
            appliedCols:'\u5df2\u8bbe\u7f6e\u6bcf\u884c\u663e\u793a: ',
            imgCount:' \u5f20\u56fe\u7247',
            errNoPath:'[\u753b\u5eca] \u8bf7\u6307\u5b9a\u6587\u4ef6\u5939\u8def\u5f84',
            errNoFolder:'[\u753b\u5eca] \u6587\u4ef6\u5939\u4e0d\u5b58\u5728: ',
            errEmpty:'(\u8be5\u6587\u4ef6\u5939\u6ca1\u6709\u56fe\u7247)',
            errRender:'[\u753b\u5eca\u9519\u8bef] ',
            layoutHoriz:'\u6c34\u5e73\u6392\u5217',
            layoutVert:'\u5782\u76f4\u7011\u5e03\u6d41',
            layoutGrid:'\u7f51\u683c\u5e03\u5c40',
            insertMenu:'\ud83d\uddbc\ufe0f \u63d2\u5165\u56fe\u7247\u753b\u5eca\u5230\u7b14\u8bb0',
            insertDone:'\u5df2\u63d2\u5165\u56fe\u7247\u753b\u5eca\u4ee3\u7801\u5757',
            openNote:'\u8bf7\u5148\u6253\u5f00\u4e00\u4e2a\u7b14\u8bb0',
            menuSpacing:'\u2699 \u95f4\u8ddd\u8bbe\u7f6e',
            titleSpacing:'\u95f4\u8ddd\u8bbe\u7f6e',
            labelSpacingLeft:'\u5de6\u4fa7\u8ddd\u79bb (px):',
            labelSpacingRight:'\u53f3\u4fa7\u8ddd\u79bb (px):',
            labelItemGap:'\u95f4\u8ddd (px):',
            appliedSpacing:'\u5df2\u5e94\u7528\u95f4\u8ddd\u8bbe\u7f6e',
            labelSmartCenter:'\u667a\u80fd\u5c45\u4e2d\uff1a\u52fe\u9009\u540e\u5185\u5bb9\u81ea\u52a8\u5c45\u4e2d',
            btnSmartCenter:'\u667a\u80fd\u5c45\u4e2d',
            btnApplyAll:'\u5e94\u7528\u5230\u6240\u6709\u7b14\u8bb0',
            titleApplyAll:'\u5e94\u7528\u5230\u6240\u6709\u7b14\u8bb0',
            confirmApplyAll:'\u786e\u5b9a\u5c06\u5f53\u524d\u95f4\u8ddd\u8bbe\u7f6e\u5e94\u7528\u5230\u6240\u6709 t \u4ee3\u7801\u5757\u5417\uff1f',
            appliedAll:'\u5df2\u66f4\u65b0 {0} \u4e2a\u7b14\u8bb0\u4e2d\u7684 {1} \u4e2a t \u4ee3\u7801\u5757',
            scanningNotes:'\u6b63\u5728\u626b\u63cf\u7b14\u8bb0...',
            noBlocksFound:'\u672a\u627e\u5230 t \u4ee3\u7801\u5757',
        },
        en: {
            menuEdit:'\u270f\ufe0f Edit Code Block',
            menuDelete:'\ud83d\uddd1\ufe0f Delete Code Block',
            menuLayout:'\ud83d\udcd0 Custom Layout',
            menuSize:'\ud83d\udccf Custom Size',
            menuColumns:'\ud83d\udd22 Custom Columns',
            titleLayout:'Custom Layout',
            titleSize:'Custom Size',
            titleColumns:'Images per Row',
            labelWidth:'Image width (px, 0=auto):',
            labelHeight:'Image height (px):',
            labelColumns:'Images per row:',
            apply:'Apply',
            cancel:'Cancel',
            perRow:'/row',
            noEditor:'Editor not found',
            noBlock:'Code block not found',
            jumped:'Jumped to code block',
            deleted:'Code block deleted',
            failOp:'Action failed: ',
            appliedLayout:'Layout applied: ',
            appliedSize:'Size set: ',
            widthSuffix:' (width: ',
            appliedCols:'Columns set: ',
            imgCount:' images',
            errNoPath:'[Gallery] Please specify folder path',
            errNoFolder:'[Gallery] Folder not found: ',
            errEmpty:'(No images in this folder)',
            errRender:'[Gallery Error] ',
            layoutHoriz:'Horizontal',
            layoutVert:'Vertical Waterfall',
            layoutGrid:'Grid',
            insertMenu:'\ud83d\uddbc\ufe0f Insert Image Gallery to Note',
            insertDone:'Image gallery block inserted',
            openNote:'Please open a note first',
            menuSpacing:'\u2699 Spacing',
            titleSpacing:'Spacing Settings',
            labelSpacingLeft:'Left padding (px):',
            labelSpacingRight:'Right padding (px):',
            labelItemGap:'Item gap (px):',
            appliedSpacing:'Spacing applied',
            labelSmartCenter:'Smart Center: auto-center content',
            btnSmartCenter:'Smart Center',
            btnApplyAll:'Apply to All Notes',
            titleApplyAll:'Apply to All Notes',
            confirmApplyAll:'Apply current spacing to ALL t code blocks?',
            appliedAll:'Updated {1} t code blocks in {0} notes',
            scanningNotes:'Scanning notes...',
            noBlocksFound:'No t code blocks found',
        }
    };

    function tr(plugin, key) {
        var lang = getLang(plugin);
        return (S[lang] && S[lang][key]) || (S['en'][key]) || key;
    }

    // ---- parse source ----
    function parseSource(source) {
        source = source.trim();
        var p = { path:'', type:'horizontal', columns:4, radius:8, gutter:12, height:200, width:0, spacingLeft:0, spacingRight:0, itemGap:0, smartCenter:false };
        if (!source) return p;
        var fl = source.split('\n')[0].trim();
        if (fl.indexOf(':') === -1 || fl.charAt(0) === '/' || fl.charAt(0) === '\\') {
            var parts = fl.split('|');
            p.path = parts[0].trim();
            if (parts[1]) p.type = parts[1].trim();
            if (parts[2] && !isNaN(parseInt(parts[2]))) p.columns = parseInt(parts[2]);
            if (parts[3] && !isNaN(parseInt(parts[3]))) p.height = parseInt(parts[3]);
            if (parts[4] && !isNaN(parseInt(parts[4]))) p.width = parseInt(parts[4]);
            if (parts[5] && !isNaN(parseInt(parts[5]))) p.spacingLeft = parseInt(parts[5]);
            if (parts[6] && !isNaN(parseInt(parts[6]))) p.spacingRight = parseInt(parts[6]);
            if (parts[7] && !isNaN(parseInt(parts[7]))) p.itemGap = parseInt(parts[7]);
            if (parts[8] && parts[8].trim().toLowerCase()==='true') p.smartCenter = true;
            return p;
        }
        var lines = source.split('\n');
        for (var i = 0; i < lines.length; i++) {
            var ln = lines[i].trim();
            if (!ln || ln.charAt(0) === '#') continue;
            var idx = ln.indexOf(':');
            if (idx === -1) continue;
            var k = ln.substring(0,idx).trim().toLowerCase();
            var v = ln.substring(idx+1).trim();
            if (k==='path') p.path=v;
            else if (k==='type') p.type=v;
            else if (k==='columns'||k==='col') p.columns=parseInt(v)||4;
            else if (k==='radius') p.radius=parseInt(v)||8;
            else if (k==='gutter') p.gutter=parseInt(v)||12;
            else if (k==='height') p.height=parseInt(v)||200;
            else if (k==='width') p.width=parseInt(v)||0;
            else if (k==='spacingleft') p.spacingLeft=parseInt(v)||0;
            else if (k==='spacingright') p.spacingRight=parseInt(v)||0;
            else if (k==='itemgap') p.itemGap=parseInt(v)||0;
            else if (k==='smartcenter') p.smartCenter = (v==='true'||v==='1');
        }
        return p;
    }

    // ============================================================
    //  LOCATE CODE BLOCK IN EDITOR (robust version)
    //  Strategy: search for ```tag\n marker, then find closing ```
    //  Works in BOTH Live Preview and Reading mode
    //  Returns { editor, lineStart, lineEnd } or null
    // ============================================================
    function findMarkdownView(app) {
        var obs;
        try { obs = require('obsidian'); } catch(e) { return null; }
        var view = null;
        try { view = app.workspace.getActiveViewOfType(obs.MarkdownView); } catch(e) {}
        if (!view) {
            try {
                app.workspace.iterateAllLeaves(function(leaf) {
                    if (!view && leaf.view instanceof obs.MarkdownView) view = leaf.view;
                });
            } catch(e) {}
        }
        return view;
    }

    function locateBlock(app, el, sourceText, langTag) {
        var view = findMarkdownView(app);
        if (!view || !view.editor) return null;

        var editor = view.editor;
        var doc;
        try { doc = editor.getValue(); } catch(e) { return null; }

        var openMark = '```' + langTag;
        var closeMark = '\n```';

        // Find ALL occurrences of ```tag and try each one
        var searchPos = 0;
        while (true) {
            var oPos = doc.indexOf(openMark, searchPos);
            if (oPos === -1) break;

            // Must be followed by newline (real code block fence)
            var charAfter = doc.charAt(oPos + openMark.length);
            if (charAfter !== '\n' && charAfter !== '\r') {
                searchPos = oPos + 1;
                continue;
            }

            var bodyStart = oPos + openMark.length + 1; // skip past ```t\n
            var cPos = doc.indexOf(closeMark, bodyStart);
            if (cPos === -1) break;

            var body = doc.substring(bodyStart, cPos).trim();

            // Match strategy:
            // 1. Exact match (user hasn't edited)
            // 2. First-line match (user may have added params)
            // 3. Path match (most reliable fallback)
            var sourceFirstLine = sourceText.trim().split('\n')[0];
            var bodyFirstLine = body.split('\n')[0];

            if (body === sourceText.trim() ||
                bodyFirstLine === sourceFirstLine ||
                body.indexOf(sourceFirstLine) !== -1 ||
                sourceText.trim().indexOf(bodyFirstLine) !== -1) {

                var ls = doc.substring(0, oPos).split('\n').length - 1;
                var le = doc.substring(0, cPos).split('\n').length;

                // Strict validation
                if (typeof ls === 'number' && typeof le === 'number' &&
                    !isNaN(ls) && !isNaN(le) && le > ls && ls >= 0) {
                    return { editor: editor, lineStart: ls, lineEnd: le };
                }
            }

            searchPos = cPos + 3;
        }
        return null;
    }

    // ---- get active editor for right-click insert ----
    function getEditorForInsert(app) {
        var obs;
        try { obs = require('obsidian'); } catch(e) { return null; }
        try {
            var v = app.workspace.getActiveViewOfType(obs.MarkdownView);
            if (v) return v.editor;
        } catch(e) {}
        try {
            var leaf = app.workspace.getMostRecentLeaf();
            if (leaf && leaf.view instanceof obs.MarkdownView) return leaf.view.editor;
        } catch(e) {}
        var f = null;
        try {
            app.workspace.iterateAllLeaves(function(l) {
                if (!f && l.view instanceof obs.MarkdownView) f = l.view.editor;
            });
        } catch(e) {}
        return f;
    }

    // ---- write new source to editor (replace inner content) ----
    function applyNewSource(newLines, editor, lineStart, lineEnd) {
        var innerS = lineStart + 1;
        var innerE = lineEnd - 1;
        var lastLn = '';
        try { lastLn = editor.getLine(innerE) || ''; } catch(e) { lastLn = ''; }
        editor.replaceRange(
            newLines.join('\n'),
            { line: innerS, ch: 0 },
            { line: innerE,   ch: lastLn.length }
        );
    }

    // ============================================================
    //  TOOLBAR (gear button + dropdown menu)
    // ============================================================
    function attachToolbar(containerEl, el, sourceText, langTag, plugin) {
        var bar = document.createElement('div');
        bar.className = 'v6-gal-toolbar';
        bar.style.cssText = 'position:absolute;top:6px;right:6px;z-index:20;display:flex;gap:2px;opacity:0;transition:opacity .15s;';
        containerEl.style.position = 'relative';
        containerEl.appendChild(bar);

        var btn = document.createElement('button');
        btn.innerHTML = '&#9881;';
        btn.style.cssText = 'width:26px;height:22px;border:none;border-radius:4px;background:rgba(0,0,0,.55);color:#fff;font-size:13px;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;line-height:1;';
        bar.appendChild(btn);

        containerEl.addEventListener('mouseenter', function() { bar.style.opacity='1'; });
        containerEl.addEventListener('mouseleave', function() { bar.style.opacity='0'; });

        var menu = null;
        btn.addEventListener('click', function(e) {
            e.stopPropagation();
            if (menu && menu.parentNode) { menu.remove(); menu=null; return; }
            menu = document.createElement('div');
            menu.style.cssText = 'position:absolute;top:100%;right:0;margin-top:3px;background:var(--background-secondary);border:1px solid var(--background-modifier-border);border-radius:6px;padding:4px 0;min-width:160px;box-shadow:0 4px 16px rgba(0,0,0,.25);z-index:100001;font-family:var(--font-interface,system-ui);font-size:13px;color:var(--text-normal);';
            bar.appendChild(menu);

            var items = [
                { k:'menuEdit',    a:'edit'    },
                { k:'menuDelete',  a:'delete'  },
                { k:'menuLayout',  a:'layout'  },
                { k:'menuSize',    a:'size'    },
                { k:'menuColumns', a:'columns' },
                { k:'menuSpacing', a:'spacing' }
            ];
            items.forEach(function(it) {
                var mi = document.createElement('div');
                mi.textContent = tr(plugin, it.k);
                mi.style.cssText = 'padding:8px 14px;cursor:pointer;white-space:nowrap;';
                mi.addEventListener('mouseenter', function() { mi.style.background='var(--background-modifier-hover)'; });
                mi.addEventListener('mouseleave', function() { mi.style.background=''; });
                mi.addEventListener('click', function() {
                    menu.remove(); menu=null;
                    handleAction(el, it.a, langTag, sourceText, plugin);
                });
                menu.appendChild(mi);
            });

            setTimeout(function() {
                document.addEventListener('mousedown', function cl(ev) {
                    if(!menu||!menu.parentNode){document.removeEventListener('mousedown',cl);return;}
                    if(!menu.contains(ev.target)){menu.remove();menu=null;document.removeEventListener('mousedown',cl);}
                });
            }, 0);
        });
    }

    // ============================================================
    //  ACTION HANDLER (edit / delete / layout / size / columns)
    // ============================================================
    function handleAction(el, action, langTag, sourceText, plugin) {
        var obs;
        try { obs = require('obsidian'); } catch(e) { return; }
        try {
            var loc = locateBlock(plugin.app, el, sourceText, langTag);
            if (!loc || !loc.editor) {
                new obs.Notice(tr(plugin,'noEditor'));
                return;
            }
            var editor = loc.editor;
            var ls = loc.lineStart;
            var le = loc.lineEnd;

            // Extra safety: validate line numbers
            if (typeof ls !== 'number' || typeof le !== 'number' || isNaN(ls) || isNaN(le)) {
                new obs.Notice(tr(plugin,'noBlock'));
                return;
            }

            if (action === 'edit') {
                editor.setCursor({ line: ls + 1, ch: 0 });
                try { editor.scrollIntoView({ from: { line: ls, ch: 0 }, to: { line: ls + 1, ch: 0 } }, true); } catch(se) {}
                editor.focus();
                new obs.Notice(tr(plugin,'jumped'));

            } else if (action === 'delete') {
                var endTxt = '';
                try { endTxt = editor.getLine(le) || ''; } catch(e) {}
                editor.replaceRange('', {line:ls,ch:0}, {line:le,ch:endTxt.length});
                new obs.Notice(tr(plugin,'deleted'));

            } else if (action === 'layout') {
                showLayoutDialog(sourceText, langTag, editor, ls, le, plugin);
            } else if (action === 'size') {
                showSizeDialog(sourceText, langTag, editor, ls, le, plugin);
            } else if (action === 'columns') {
                showColumnsDialog(sourceText, langTag, editor, ls, le, plugin);
            } else if (action === 'spacing') {
                showSpacingDialog(sourceText, langTag, editor, ls, le, plugin);
            }
        } catch(err) {
            console.warn('[Gallery] Action error:', err);
            try { new obs.Notice(tr(plugin,'failOp') + (err.message||err)); } catch(e2) {}
        }
    }

    // ---- DIALOG: Layout ----
    function showLayoutDialog(src, tag, ed, ls, le, pl) {
        var obs = require('obsidian');
        var ov = mkOverlay();
        var dlg = mkDlg();
        dlg.appendChild(mkTitle(tr(pl,'titleLayout')));
        var curType = 'horizontal';
        var fl = src.split('\n')[0].trim();
        if (fl.indexOf('|') !== -1) { var tp = fl.split('|')[1]; if(tp) curType=tp.trim(); }

        var opts = [
            {v:'horizontal', lk:'layoutHoriz'},
            {v:'vertical',   lk:'layoutVert'},
            {v:'grid',       lk:'layoutGrid'}
        ];
        opts.forEach(function(o) {
            var row = mkOptRow(curType===o.v, tr(pl,o.lk), o.v);
            row.addEventListener('click', function(){
                var rb = row.querySelector('input');
                if (rb) rb.checked = true;
                dlg.querySelectorAll('[data-lr]').forEach(function(r){r.style.borderColor='var(--background-modifier-border)';r.style.background='';});
                row.style.borderColor='var(--interactive-accent)';
                row.style.background='var(--background-modifier-hover)';
            });
            dlg.appendChild(row);
        });
        dlg.appendChild(mkBtnRow(
            mkCancelBtn(tr(pl,'cancel'),function(){ov.remove();}),
            mkApplyBtn(tr(pl,'apply'),function(){
                var sel=dlg.querySelector('input[name="lr"]:checked');
                if(!sel){ov.remove();return;}
                var nt=sel.value;
                var ln=src.split('\n');
                if(ln.length>0){
                    var ps=ln[0].split('|');
                    if(ps.length>=2)ps[1]=nt;else ps.push(nt);
                    ln[0]=ps.join('|');
                }
                applyNewSource(ln,ed,ls,le); ov.remove();
                new obs.Notice(tr(pl,'appliedLayout')+nt);
            })
        ));
        ov.appendChild(dlg);
        ov.addEventListener('click',function(e){if(e.target===ov)ov.remove();});
        document.body.appendChild(ov);
    }

    // ---- DIALOG: Size ----
    function showSizeDialog(src, tag, ed, ls, le, pl) {
        var obs = require('obsidian');
        var ov = mkOverlay();
        var dlg = mkDlg(); dlg.style.minWidth='360px';
        dlg.appendChild(mkTitle(tr(pl,'titleSize')));
        var w=0,h=200;
        var fl=src.split('\n')[0].trim();
        if(fl.indexOf('|')!==-1){var pp=fl.split('|');if(pp[3])h=parseInt(pp[3])||200;if(pp[4])w=parseInt(pp[4])||0;}

        var wi=addNumField(tr(pl,'labelWidth'),w,0,undefined,dlg);
        var hi=addNumField(tr(pl,'labelHeight'),h,50,1000,dlg);

        var pr=document.createElement('div');
        pr.style.cssText='display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;';
        [150,200,250,300].forEach(function(p){
            var b=document.createElement('button');
            b.textContent=p+'px';
            b.style.cssText='padding:6px 12px;border:1px solid var(--background-modifier-border);border-radius:4px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:12px;';
            b.addEventListener('click',function(){hi.value=p;});
            pr.appendChild(b);
        }); dlg.appendChild(pr);
        dlg.appendChild(mkBtnRow(
            mkCancelBtn(tr(pl,'cancel'),function(){ov.remove();}),
            mkApplyBtn(tr(pl,'apply'),function(){
                var nw=parseInt(wi.value)||0;
                var nh=Math.max(50,Math.min(1000,parseInt(hi.value)||h));
                var ln=src.split('\n');
                if(ln.length>0){
                    var ps=ln[0].split('|');
                    while(ps.length<5)ps.push('');
                    ps[3]=String(nh); ps[4]=nw>0?String(nw):'';
                    ln[0]=ps.join('|');
                }
                applyNewSource(ln,ed,ls,le); ov.remove();
                var msg=tr(pl,'appliedSize')+nh+'px';
                if(nw>0) msg+=tr(pl,'widthSuffix')+nw+'px)';
                new obs.Notice(msg);
            })
        )); ov.appendChild(dlg);
        ov.addEventListener('click',function(e){if(e.target===ov)ov.remove();});
        document.body.appendChild(ov);
    }

    // ---- DIALOG: Columns ----
    function showColumnsDialog(src, tag, ed, ls, le, pl) {
        var obs = require('obsidian');
        var ov = mkOverlay();
        var dlg = mkDlg();
        dlg.appendChild(mkTitle(tr(pl,'titleColumns')));
        var cols=4;
        var fl=src.split('\n')[0].trim();
        if(fl.indexOf('|')!==-1){var pp=fl.split('|');if(pp[2])cols=parseInt(pp[2])||4;}

        var lb=document.createElement('label');
        lb.textContent=tr(pl,'labelColumns');
        lb.style.cssText='display:block;font-size:12px;color:var(--text-muted);margin-bottom:6px;';
        dlg.appendChild(lb);
        var inp=document.createElement('input');
        inp.type='number';inp.min=1;inp.max=20;inp.value=cols;
        inp.style.cssText='width:100%;padding:8px 10px;border:1px solid var(--background-modifier-border);border-radius:5px;background:var(--background-secondary);color:var(--text-normal);font-size:14px;box-sizing:border-box;margin-bottom:16px;';
        dlg.appendChild(inp);

        var pr2=document.createElement('div');
        pr2.style.cssText='display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;';
        [2,3,4,5,6,8].forEach(function(p){
            var b=document.createElement('button');
            b.textContent=p+tr(pl,'perRow');
            b.style.cssText='padding:6px 12px;border:1px solid var(--background-modifier-border);border-radius:4px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:12px;';
            b.addEventListener('click',function(){inp.value=p;});
            pr2.appendChild(b);
        }); dlg.appendChild(pr2);
        dlg.appendChild(mkBtnRow(
            mkCancelBtn(tr(pl,'cancel'),function(){ov.remove();}),
            mkApplyBtn(tr(pl,'apply'),function(){
                var nc=Math.max(1,Math.min(20,parseInt(inp.value)||cols));
                var ln=src.split('\n');
                if(ln.length>0){
                    var ps=ln[0].split('|');
                    while(ps.length<3)ps.push('');
                    ps[2]=String(nc); ln[0]=ps.join('|');
                }
                applyNewSource(ln,ed,ls,le); ov.remove();
                new obs.Notice(tr(pl,'appliedCols')+nc+tr(pl,'perRow'));
            })
        )); ov.appendChild(dlg);
        ov.addEventListener('click',function(e){if(e.target===ov)ov.remove();});
        document.body.appendChild(ov);
    }

    // ---- DIALOG: Spacing ----
    function showSpacingDialog(src, tag, ed, ls, le, pl) {
        var obs = require('obsidian');
        var ov = mkOverlay();
        var dlg = mkDlg(); dlg.style.minWidth='360px';
        dlg.appendChild(mkTitle(tr(pl,'titleSpacing')));
        var sl=0, sr=0, ig=0;
        var fl=src.split('\n')[0].trim();
        if(fl.indexOf('|')!==-1){var pp=fl.split('|');if(pp[5])sl=parseInt(pp[5])||0;if(pp[6])sr=parseInt(pp[6])||0;if(pp[7])ig=parseInt(pp[7])||0;}

        var li=addNumField(tr(pl,'labelSpacingLeft'),sl,0,undefined,dlg);
        var ri=addNumField(tr(pl,'labelSpacingRight'),sr,0,undefined,dlg);
        var gi=addNumField(tr(pl,'labelItemGap'),ig,0,undefined,dlg);

        var pr=document.createElement('div');
        pr.style.cssText='display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;';
        [0,5,10,15,20,30].forEach(function(p){
            var b=document.createElement('button');
            b.textContent=p+'px';
            b.style.cssText='padding:6px 12px;border:1px solid var(--background-modifier-border);border-radius:4px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:12px;';
            b.addEventListener('click',function(){li.value=p;ri.value=p;gi.value=p;});
            pr.appendChild(b);
        }); dlg.appendChild(pr);
        var scCb=document.createElement('div');scCb.style.cssText='margin-bottom:16px;display:flex;align-items:center;gap:8px;';
        var scBox=document.createElement('input');scBox.type='checkbox';scBox.style.cssText='width:16px;height:16px;cursor:pointer;';
        scCb.appendChild(scBox);
        var scLabel=document.createElement('span');scLabel.textContent=tr(pl,'labelSmartCenter');
        scLabel.style.cssText='font-size:13px;color:var(--text-normal);cursor:pointer;';
        scCb.appendChild(scLabel);
        dlg.appendChild(scCb);
        dlg.appendChild(mkBtnRow(
            mkCancelBtn(tr(pl,'cancel'),function(){ov.remove();}),
            mkApplyBtn(tr(pl,'apply'),function(){
                var nsl=parseInt(li.value)||0;
                var nsr=parseInt(ri.value)||0;
                var nig=parseInt(gi.value)||0;
                var ln=src.split('\n');
                if(ln.length>0){
                    var ps=ln[0].split('|');
                    while(ps.length<8)ps.push('');
                    ps[5]=String(nsl);ps[6]=String(nsr);ps[7]=String(nig);
                    ln[0]=ps.join('|');
                }
                applyNewSource(ln,ed,ls,le); ov.remove();
                new obs.Notice(tr(pl,'appliedSpacing'));
            })
        )); ov.appendChild(dlg);
        ov.addEventListener('click',function(e){if(e.target===ov)ov.remove();});
        document.body.appendChild(ov);
    }

    // ============================================================
    //  DIALOG HELPERS
    // ============================================================
    function mkOverlay() {
        var d=document.createElement('div');
        d.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.35);z-index:100002;display:flex;align-items:center;justify-content:center;';
        return d;
    }
    function mkDlg() {
        var d=document.createElement('div');
        d.style.cssText='background:var(--background-primary);border:1px solid var(--background-modifier-border);border-radius:10px;padding:20px 24px;min-width:320px;box-shadow:0 12px 48px rgba(0,0,0,.35);font-family:var(--font-interface,system-ui);';
        return d;
    }
    function mkTitle(t){
        var d=document.createElement('div');
        d.textContent=t;
        d.style.cssText='font-size:15px;font-weight:600;margin-bottom:16px;color:var(--text-normal);';
        return d;
    }
    function mkBtnRow(){
        var r=document.createElement('div');
        r.style.cssText='display:flex;gap:8px;margin-top:16px;justify-content:flex-end;';
        for(var i=0;i<arguments.length;i++){r.appendChild(arguments[i]);}
        return r;
    }
    function mkCancelBtn(txt,fn){
        var b=document.createElement('button');
        b.textContent=txt;
        b.style.cssText='padding:8px 16px;border:1px solid var(--background-modifier-border);border-radius:5px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:12px;';
        b.addEventListener('click',fn); return b;
    }
    function mkApplyBtn(txt,fn){
        var b=document.createElement('button');
        b.textContent=txt;
        b.style.cssText='padding:8px 18px;border:none;border-radius:5px;background:var(--interactive-accent);color:#fff;cursor:pointer;font-size:12px;font-weight:600;';
        b.addEventListener('click',fn); return b;
    }
    function mkOptRow(active,labelKey,val){
        var r=document.createElement('div');
        r.setAttribute('data-lr','1');
        r.style.cssText='padding:10px;margin-bottom:8px;border:1px solid '+(active?'var(--interactive-accent)':'var(--background-modifier-border)')+';border-radius:6px;cursor:pointer;display:flex;align-items:center;gap:10px;'+(active?'background:var(--background-modifier-hover);':'');
        var rb=document.createElement('input');
        rb.type='radio';rb.name='lr';rb.checked=active;rb.value=val||'';
        r.appendChild(rb);
        var sp=document.createElement('span');
        sp.textContent=labelKey;
        r.appendChild(sp); return r;
    }
    function addNumField(labelKey,val,min,max,parent){
        var lb=document.createElement('label');
        lb.textContent=labelKey;
        lb.style.cssText='display:block;font-size:12px;color:var(--text-muted);margin-bottom:6px;';
        parent.appendChild(lb);
        var inp=document.createElement('input');
        inp.type='number';inp.value=val;
        if(min!==undefined)inp.min=min;
        if(max!==undefined)inp.max=max;
        inp.style.cssText='width:100%;padding:8px 10px;border:1px solid var(--background-modifier-border);border-radius:5px;background:var(--background-secondary);color:var(--text-normal);font-size:14px;box-sizing:border-box;margin-bottom:12px;';
        parent.appendChild(inp); return inp;
    }

    // ---- LIGHTBOX ----
    function openLightbox(src,alt){
        var ex=document.querySelector('.v6-lightbox-overlay');
        if(ex)ex.remove();
        var ov=document.createElement('div');
        ov.className='v6-lightbox-overlay';
        ov.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.85);z-index:100000;display:flex;align-items:center;justify-content:center;cursor:pointer;';
        var img=document.createElement('img');
        img.src=src;img.alt=alt||'';
        img.style.cssText='max-width:90vw;max-height:85vh;object-fit:contain;border-radius:8px;box-shadow:0 8px 40px rgba(0,0,0,.5);user-select:none;';
        var cap=document.createElement('div');
        cap.textContent=alt||'';
        cap.style.cssText='color:#ccc;margin-top:14px;font-size:13px;text-align:center;font-family:system-ui;pointer-events:none;';
        var box=document.createElement('div');
        box.style.cssText='display:flex;flex-direction:column;align-items:center;';
        box.appendChild(img);box.appendChild(cap);
        ov.appendChild(box);
        ov.addEventListener('click',function(){ov.remove();});
        document.addEventListener('keydown',function onEsc(ev){
            if(ev.key==='Escape'){ov.remove();document.removeEventListener('keydown',onEsc);}
        });
        document.body.appendChild(ov);
    }

    // ---- SHOW COUNT HELPER ----
    function shouldShowCount(plugin) {
        try {
            var m = plugin.settings.modules;
            if (m && m['img-gallery'] && m['img-gallery'].showCount === false) return false;
        } catch(e) {}
        return true; // default: show
    }

    // ============================================================
    //  MAIN SETUP
    // ============================================================
    function setupImgGalleryProcessor(plugin) {
        try {
            var obs = require('obsidian');
            var normalizePath = obs.normalizePath;
            var app = plugin.app;

            // --- Context menu: right-click image/folder -> Insert Gallery ---
            plugin.registerEvent(
                app.workspace.on('file-menu', function(menu, file) {
                    var isFolder = file instanceof obs.TFolder;
                    var isImg = file instanceof obs.TFile && VALID_EXTENSIONS.indexOf(file.extension.toLowerCase()) !== -1;
                    if (!isFolder && !isImg) return;
                    menu.addItem(function(item) {
                        item.setTitle(tr(plugin,'insertMenu')).setIcon('image')
                            .onClick(function() {
                                var tp = isFolder ? file.path : file.parent.path;
                                var ed = getEditorForInsert(app);
                                if (!ed) { new obs.Notice(tr(plugin,'openNote')); return; }
                                var cur = ed.getCursor();
                                ed.replaceRange('\n```t\n' + tp + '\n```\n', cur);
                                new obs.Notice(tr(plugin,'insertDone'));
                            });
                    });
                })
            );

            // --- Register ```t code block processor ---
            plugin.registerMarkdownCodeBlockProcessor('t', function(source, el, ctx) {
                var Child = function(c, src) {
                    obs.MarkdownRenderChild.call(this, c);
                    this.src = src;
                    this.el = el;
                };
                Child.prototype = Object.create(obs.MarkdownRenderChild.prototype);
                Child.prototype.constructor = Child;

                Child.prototype.onload = function() {
                    var self = this;
                    try {
                        var container = self.containerEl;
                        container.empty();
                        container.className = 'v6-img-gallery';
                        attachToolbar(container, self.el, self.src, 't', plugin);

                        var params = parseSource(self.src);
                        var targetPath = normalizePath(params.path);

                        if (!targetPath) {
                            container.createEl('div', { cls:'v6-ig-error', text: tr(plugin,'errNoPath') });
                            return;
                        }

                        var folder = app.vault.getAbstractFileByPath(targetPath);
                        if (!folder || !(folder instanceof obs.TFolder)) {
                            container.createEl('div', { cls:'v6-ig-error', text: tr(plugin,'errNoFolder') + targetPath });
                            return;
                        }

                        var imageFiles = [];
                        folder.children.forEach(function(f) {
                            if (f instanceof obs.TFile && VALID_EXTENSIONS.indexOf(f.extension.toLowerCase()) !== -1)
                                imageFiles.push(f);
                        });

                        if (imageFiles.length === 0) {
                            container.createEl('div', { cls:'v6-ig-empty', text: tr(plugin,'errEmpty') });
                            return;
                        }

                        var isVertical = (params.type==='vertical'||params.type==='waterfall'||params.type==='masonry');
                        var galleryWrap = container.createEl('div', { cls: isVertical ? 'v6-ig-vertical' : 'v6-ig-horizontal' });
                        var totalCount = imageFiles.length;
                        var cols = Math.max(1, params.columns);
                        var gutter = params.gutter;
                        var h = params.height;
                        var w = params.width;
                        var rad = params.radius;

                        if (!isVertical) {
                            // Horizontal/Grid: CSS Grid guarantees exact column count
                            var itemGap = params.itemGap > 0 ? params.itemGap : gutter;
                            if(params.smartCenter){
                                galleryWrap.style.cssText = 'display:flex;flex-wrap:wrap;justify-content:center;gap:'+itemGap+'px;padding-left:'+params.spacingLeft+'px;padding-right:'+params.spacingRight+'px;';
                            } else {
                                galleryWrap.style.cssText = 'display:grid;grid-template-columns:repeat('+cols+',1fr);gap:'+itemGap+'px;padding-left:'+params.spacingLeft+'px;padding-right:'+params.spacingRight+'px;';
                            }
                            imageFiles.forEach(function(imgFile) {
                                var rp = app.vault.getResourcePath(imgFile);
                                var item = document.createElement('div');
                                item.style.cssText = 'overflow:hidden;border-radius:'+rad+'px;cursor:pointer;height:'+h+'px;';
                                if(w>0)item.style.width=w+'px';
                                var img=document.createElement('img');
                                img.src=rp;img.alt=imgFile.basename;
                                img.style.cssText='width:100%;height:100%;object-fit:cover;display:block;transition:transform .25s;';
                                img.draggable=false;
                                img.addEventListener('mouseenter',function(){img.style.transform='scale(1.03)';});
                                img.addEventListener('mouseleave',function(){img.style.transform='';});
                                item.appendChild(img);
                                galleryWrap.appendChild(item);
                                (function(r,n){item.addEventListener('click',function(){openLightbox(r,n);});})(rp,imgFile.basename);
                            });
                        } else {
                            // Vertical waterfall: CSS columns
                            var itemGapV = params.itemGap > 0 ? params.itemGap : gutter;
                            if(params.smartCenter){
                                galleryWrap.style.cssText = 'display:flex;flex-wrap:wrap;justify-content:center;gap:'+itemGapV+'px;padding-left:'+params.spacingLeft+'px;padding-right:'+params.spacingRight+'px;';
                            } else {
                                galleryWrap.style.cssText = 'column-count:'+Math.min(cols,6)+';column-gap:'+itemGapV+'px;padding-left:'+params.spacingLeft+'px;padding-right:'+params.spacingRight+'px;';
                            }
                            imageFiles.forEach(function(vf) {
                                var vr = app.vault.getResourcePath(vf);
                                var item = document.createElement('div');
                                item.style.cssText = 'break-inside:avoid;margin-bottom:'+gutter+'px;border-radius:'+rad+'px;overflow:hidden;cursor:pointer;display:inline-block;width:100%;';
                                var img=document.createElement('img');
                                img.src=vr;img.alt=vf.basename;
                                img.style.cssText='width:100%;display:block;border-radius:'+rad+'px;transition:transform .25s;';
                                img.draggable=false;
                                img.addEventListener('mouseenter',function(){img.style.transform='scale(1.03)';});
                                img.addEventListener('mouseleave',function(){img.style.transform='';});
                                item.appendChild(img);
                                galleryWrap.appendChild(item);
                                (function(r,n){item.addEventListener('click',function(){openLightbox(r,n);});})(vr,vf.basename);
                            });
                        }

                        // Show media count (controlled by setting)
                        if (shouldShowCount(plugin)) {
                            container.createEl('div', { cls:'v6-ig-count', text: totalCount + tr(plugin,'imgCount') });
                        }
                    } catch(renderErr) {
                        console.warn('[img-gallery] render error:', renderErr);
                        self.containerEl.textContent = tr(plugin,'errRender') + (renderErr.message||renderErr);
                    }
                };

                ctx.addChild(new Child(el, source));
            });

            console.log('[Dashboard] image gallery processor registered (code: t)');
        } catch(e) {
            console.error('[Dashboard] setupImgGalleryProcessor failed:', e);
        }
    }

    if (typeof global !== 'undefined') global.__setupImgGalleryProcessor = setupImgGalleryProcessor;
    else if (typeof window !== 'undefined') window.__setupImgGalleryProcessor = setupImgGalleryProcessor;
    if (typeof module !== 'undefined' && module.exports) module.exports = setupImgGalleryProcessor;

})(typeof globalThis !== 'undefined' ? globalThis : (typeof window !== 'undefined' ? window : typeof global !== 'undefined' ? global : this));

(function(global) {
    'use strict';

    var IE={jpeg:1,jpg:1,gif:1,png:1,webp:1,tiff:1,tif:1,avif:1,bmp:1,svg:1,ico:1,heic:1,heif:1};
    var VE={mp4:1,webm:1,ogg:1,mov:1,mkv:1,avi:1,wmv:1,m4v:1,'3gp':1,flv:1};
    var AE={mp3:1,wav:1,ogg:1,flac:1,aac:1,m4a:1,wma:1,opus:1,aiff:1};
    var AM={};[IE,VE,AE].forEach(function(o){Object.keys(o).forEach(function(k){AM[k]=1;});});

    function getLang(plugin) {
        try {
            if (typeof window.getResolvedLang === 'function') {
                var rl = window.getResolvedLang();
                if (rl && rl.startsWith('zh')) return 'zh';
                if (rl) return 'en';
            }
            var lang = plugin && plugin.settings && plugin.settings.language;
            if (!lang || lang === 'system') {
                var nav = (navigator.language || navigator.userLanguage || 'en').toLowerCase();
                return nav.startsWith('zh') ? 'zh' : 'en';
            }
            if (lang === 'ai' || lang.startsWith('custom_')) return 'en';
            return lang.startsWith('zh') ? 'zh' : 'en';
        } catch(e) {}
        return 'en';
    }

    var S={
        zh:{
            menuEdit:'\u270f\ufe0f \u7f16\u8f91\u4ee3\u7801\u5757',
            menuDelete:'\ud83d\uddd1\ufe0f \u5220\u9664\u4ee3\u7801\u5757',
            menuLayout:'\ud83d\udcd0 \u81ea\u5b9a\u4e49\u6392\u7248',
            menuSize:'\ud83d\udccf \u81ea\u5b9a\u4e49\u5927\u5c0f',
            apply:'\u5e94\u7528',cancel:'\u53d6\u6d88',
            titleLayout:'\u81ea\u5b9a\u4e49\u6392\u7248',titleSize:'\u5a92\u4f53\u5927\u5c0f',
            labelSize:'\u5a92\u4f53\u683c\u5b50\u5927\u5c0f (px):',
            layoutGrid:'\u7f51\u683c\u5e03\u5c40',layoutList:'\u5217\u8868\u5e03\u5c40',layoutFull:'\u5168\u5bbd\u5e03\u5c40',
            noEditor:'\u672a\u627e\u5230\u7f16\u8f91\u5668',noBlock:'\u672a\u627e\u5230\u4ee3\u7801\u5757',
            jumped:'\u5df2\u8df3\u8f6c\u5230\u4ee3\u7801\u5757',deleted:'\u5df2\u5220\u9664\u4ee3\u7801\u5757',
            failOp:'\u64cd\u4f5c\u5931\u8d25: ',
            appliedLayout:'\u5df2\u5e94\u7528\u6392\u7248: ',appliedSize:'\u5df8\u8bbe\u7f6e\u5927\u5c0f: ',
            errNoPath:'[\u5a92\u4f53\u753b\u5eca] \u8bf7\u6307\u5b9a\u8def\u5f84',
            errEmpty:'(\u672a\u627e\u5230\u5a92\u4f53\u6587\u4ef6)',
            errRender:'[\u5a92\u4f53\u753b\u5eca\u9519\u8bef] ',
            mediaCount:' \u4e2a\u5a92\u4f53',
            insertMenu:'\ud83c\udfa6 \u63d2\u5165\u5a92\u4f53\u753b\u5eca\u5230\u7b14\u8bb0',
            insertDone:'\u5df2\u63d2\u5165\u5a92\u4f53\u753b\u5eca\u4ee3\u7801\u5757',
            openNote:'\u8bf7\u5148\u6253\u5f00\u4e00\u4e2a\u7b14\u8bb0',
            menuSpacing:'\u2699 \u95f4\u8ddd\u8bbe\u7f6e',
            titleSpacing:'\u95f4\u8ddd\u8bbe\u7f6e',
            labelSpacingLeft:'\u5de6\u4fa7\u8ddd\u79bb (px):',
            labelSpacingRight:'\u53f3\u4fa7\u8ddd\u79bb (px):',
            labelItemGap:'\u95f4\u8ddd (px):',
            labelSmartCenter:'\u667a\u80fd\u5c45\u4e2d\uff1a\u52fe\u9009\u540e\u5185\u5bb9\u81ea\u52a8\u5c45\u4e2d',
            btnSmartCenter:'\u667a\u80fd\u5c45\u4e2d',
            btnApplyAll:'\u5e94\u7528\u5230\u6240\u6709\u7b14\u8bb0',
            titleApplyAll:'\u5e94\u7528\u5230\u6240\u6709\u7b14\u8bb0',
            confirmApplyAll:'\u786e\u5b9a\u5c06\u5f53\u524d\u95f4\u8ddd\u8bbe\u7f6e\u5e94\u7528\u5230\u6240\u6709 s \u4ee3\u7801\u5757\u5417\uff1f',
            appliedAll:'\u5df2\u66f4\u65b0 {0} \u4e2a\u7b14\u8bb0\u4e2d\u7684 {1} \u4e2a s \u4ee3\u7801\u5757',
            scanningNotes:'\u6b63\u5728\u626b\u63cf\u7b14\u8bb0...',
            noBlocksFound:'\u672a\u627e\u5230 s \u4ee3\u7801\u5757',
            appliedSpacing:'\u5df2\u5e94\u7528\u95f4\u8ddd\u8bbe\u7f6e',
        },
        en:{
            menuEdit:'\u270f\ufe0f Edit Code Block',
            menuDelete:'\ud83d\uddd1\ufe0f Delete Code Block',
            menuLayout:'\ud83d\udcd0 Custom Layout',
            menuSize:'\ud83d\udccf Custom Size',
            apply:'Apply',cancel:'Cancel',
            titleLayout:'Custom Layout',titleSize:'Media Size',
            labelSize:'Media cell size (px):',
            layoutGrid:'Grid',layoutList:'List',layoutFull:'Full Width',
            noEditor:'Editor not found',noBlock:'Code block not found',
            jumped:'Jumped to code block',deleted:'Code block deleted',
            failOp:'Action failed: ',
            appliedLayout:'Layout applied: ',appliedSize:'Size set: ',
            errNoPath:'[Media Gallery] Please specify a path',
            errEmpty:'(No media files found)',
            errRender:'[Media Gallery Error] ',
            mediaCount:' media files',
            insertMenu:'\ud83c\udfa6 Insert Media Gallery to Note',
            insertDone:'Media gallery block inserted',
            openNote:'Please open a note first',
            menuSpacing:'\u2699 Spacing',
            titleSpacing:'Spacing Settings',
            labelSpacingLeft:'Left padding (px):',
            labelSpacingRight:'Right padding (px):',
            labelItemGap:'Item gap (px):',
            appliedSpacing:'Spacing applied',
            labelSmartCenter:'Smart Center: auto-center content',
            btnSmartCenter:'Smart Center',
            btnApplyAll:'Apply to All Notes',
            titleApplyAll:'Apply to All Notes',
            confirmApplyAll:'Apply current spacing to ALL s code blocks?',
            appliedAll:'Updated {1} s code blocks in {0} notes',
            scanningNotes:'Scanning notes...',
            noBlocksFound:'No s code blocks found',
        }
    };

    function tr(plugin,key){var l=getLang(plugin);return(S[l]&&S[l][key])||(S['en'][key])||key;}

    function getMediaType(ext){
        ext=ext.toLowerCase();if(IE[ext])return'image';if(VE[ext])return'video';if(AE[ext])return'audio';return null;
    }

    // ---- parse source ----
    function parseSource(source){
        source=source.trim();var p={paths:[],sort:'name',type:'grid',size:220,limit:0,spacingLeft:0,spacingRight:0,itemGap:0,smartCenter:false};
        if(!source)return p;
        var fl=source.split('\n')[0].trim();
        if(fl.indexOf(':')===-1||fl.charAt(0)==='/'||fl.charAt(0)==='\\'){
            var parts=fl.split('|');p.paths.push(parts[0].trim());
            if(parts[1]){var tok=parts[1].trim().toLowerCase();
                if(tok==='full'||tok==='list'||tok==='grid')p.type=tok;
                else if(!isNaN(parseInt(tok)))p.size=parseInt(tok);}
            if(parts[2]&&!isNaN(parseInt(parts[2])))p.size=parseInt(parts[2]);
            if(parts[3]&&!isNaN(parseInt(parts[3])))p.spacingLeft=parseInt(parts[3]);
            if(parts[4]&&!isNaN(parseInt(parts[4])))p.spacingRight=parseInt(parts[4]);
            if(parts[5]&&!isNaN(parseInt(parts[5])))p.itemGap=parseInt(parts[5]);
            if(parts[6]&&parts[6].trim().toLowerCase()==='true')p.smartCenter=true;
            return p;
        }
        var lines=source.split('\n');
        for(var i=0;i<lines.length;i++){
            var ln=lines[i].trim();if(!ln||ln.charAt(0)==='#')continue;
            var ix=ln.indexOf(':');if(ix===-1)continue;
            var k=ln.substring(0,ix).trim().toLowerCase(),v=ln.substring(ix+1).trim();
            if(k==='path'||k==='paths')v.split(',').forEach(function(x){x=x.trim();if(x)p.paths.push(x);});
            else if(k==='sort')p.sort=v;
            else if(k==='type')p.type=v;
            else if(k==='size')p.size=parseInt(v)||220;
            else if(k==='limit'||k==='max')p.limit=parseInt(v)||0;
            else if(k==='spacingleft')p.spacingLeft=parseInt(v)||0;
            else if(k==='spacingright')p.spacingRight=parseInt(v)||0;
            else if(k==='itemgap')p.itemGap=parseInt(v)||0;
            else if(k==='smartcenter')p.smartCenter=(v==='true'||v==='1');
        }
        return p;
    }

    // ============================================================
    //  LOCATE CODE BLOCK (same robust approach as img-gallery)
    // ============================================================
    function findMarkdownView(app){
        var obs;try{obs=require('obsidian');}catch(e){return null;}
        var v=null;
        try{v=app.workspace.getActiveViewOfType(obs.MarkdownView);}catch(e){}
        if(!v){try{app.workspace.iterateAllLeaves(function(l){if(!v&&l.view instanceof obs.MarkdownView)v=l.view;});}catch(e){}
        }return v;
    }

    function locateBlock(app,el,sourceText,langTag){
        var view=findMarkdownView(app);
        if(!view||!view.editor)return null;
        var editor=view.editor;
        var doc;try{doc=editor.getValue();}catch(e){return null;}
        var openMark='```'+langTag,closeMark='\n```';
        var sp=0;
        while(true){
            var oPos=doc.indexOf(openMark,sp);if(oPos===-1)break;
            var ca=doc.charAt(oPos+openMark.length);
            if(ca!=='\n'&&ca!=='\r'){sp=oPos+1;continue;}
            var bodyStart=oPos+openMark.length+1;
            var cPos=doc.indexOf(closeMark,bodyStart);if(cPos===-1)break;
            var body=doc.substring(bodyStart,cPos).trim();
            var sFL=sourceText.trim().split('\n')[0],bFL=body.split('\n')[0];
            if(body===sourceText.trim()||bFL===sFL||body.indexOf(sFL)!==-1||sourceText.trim().indexOf(bFL)!==-1){
                var ls=doc.substring(0,oPos).split('\n').length-1;
                var le=doc.substring(0,cPos).split('\n').length;
                if(typeof ls==='number'&&typeof le==='number'&&!isNaN(ls)&&!isNaN(le)&&le>ls&&ls>=0)
                    return{editor:editor,lineStart:ls,lineEnd:le};
            }
            sp=cPos+3;
        }
        return null;
    }

    function getEditorForInsert(app){
        var obs;try{obs=require('obsidian');}catch(e){return null;}
        try{var v=app.workspace.getActiveViewOfType(obs.MarkdownView);if(v)return v.editor;}catch(e){}
        try{var lf=app.workspace.getMostRecentLeaf();if(lf&&lf.view instanceof obs.MarkdownView)return lf.view.editor;}catch(e){}
        var f=null;
        try{app.workspace.iterateAllLeaves(function(l){if(!f&&l.view instanceof obs.MarkdownView)f=l.view.editor;});}catch(e){}
        return f;
    }

    function applyNewSource(lines,ed,ls,le){
        var is=ls+1,ie=le-1,lastLn='';
        try{lastLn=ed.getLine(ie)||'';}catch(e){}
        ed.replaceRange(lines.join('\n'),{line:is,ch:0},{line:ie,ch:lastLn.length});
    }

    // ---- toolbar ----
    function attachToolbar(containerEl,el,src,tag,plugin){
        var bar=document.createElement('div');
        bar.className='v6-gal-toolbar';
        bar.style.cssText='position:absolute;top:6px;right:6px;z-index:20;display:flex;gap:2px;opacity:0;transition:opacity .15s;';
        containerEl.style.position='relative';containerEl.appendChild(bar);
        var btn=document.createElement('button');
        btn.innerHTML='&#9881;';
        btn.style.cssText='width:26px;height:22px;border:none;border-radius:4px;background:rgba(0,0,0,.55);color:#fff;font-size:13px;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;line-height:1;';
        bar.appendChild(btn);
        containerEl.addEventListener('mouseenter',function(){bar.style.opacity='1';});
        containerEl.addEventListener('mouseleave',function(){bar.style.opacity='0';});

        var menu=null;
        btn.addEventListener('click',function(e){
            e.stopPropagation();
            if(menu&&menu.parentNode){menu.remove();menu=null;return;}
            menu=document.createElement('div');
            menu.style.cssText='position:absolute;top:100%;right:0;margin-top:3px;background:var(--background-secondary);border:1px solid var(--background-modifier-border);border-radius:6px;padding:4px 0;min-width:170px;box-shadow:0 4px 16px rgba(0,0,0,.25);z-index:100001;font-family:var(--font-interface,system-ui);font-size:13px;color:var(--text-normal);';
            bar.appendChild(menu);

            var items=[{k:'menuEdit',a:'edit'},{k:'menuDelete',a:'delete'},{k:'menuLayout',a:'layout'},{k:'menuSize',a:'size'},{k:'menuSpacing',a:'spacing'}];
            items.forEach(function(it){
                var mi=document.createElement('div');
                mi.textContent=tr(plugin,it.k);
                mi.style.cssText='padding:8px 14px;cursor:pointer;white-space:nowrap;';
                mi.addEventListener('mouseenter',function(){mi.style.background='var(--background-modifier-hover)';});
                mi.addEventListener('mouseleave',function(){mi.style.background='';});
                mi.addEventListener('click',function(){menu.remove();menu=null;handleAction(el,it.a,tag,src,plugin);});
                menu.appendChild(mi);
            });
            setTimeout(function(){
                document.addEventListener('mousedown',function cl(ev){
                    if(!menu||!menu.parentNode){document.removeEventListener('mousedown',cl);return;}
                    if(!menu.contains(ev.target)){menu.remove();menu=null;document.removeEventListener('mousedown',cl);}
                });
            },0);
        });
    }

    // ---- action handler ----
    function handleAction(el,action,tag,src,plugin){
        var obs;try{obs=require('obsidian');}catch(e){return;}
        try{
            var loc=locateBlock(plugin.app,el,src,tag);
            if(!loc||!loc.editor){new obs.Notice(tr(plugin,'noEditor'));return;}
            var editor=loc.editor,ls=loc.lineStart,le=loc.lineEnd;
            if(typeof ls!=='number'||typeof le!=='number'||isNaN(ls)||isNaN(le)){
                new obs.Notice(tr(plugin,'noBlock'));return;}
            if(action==='edit'){
                editor.setCursor({line:ls+1,ch:0});
                try{editor.scrollIntoView({from:{line:ls,ch:0},to:{line:ls+1,ch:0}},true);}catch(se){}editor.focus();
                new obs.Notice(tr(plugin,'jumped'));
            }else if(action==='delete'){
                var et='';try{et=editor.getLine(le)||'';}catch(e){}
                editor.replaceRange('',{line:ls,ch:0},{line:le,ch:et.length});
                new obs.Notice(tr(plugin,'deleted'));
            }else if(action==='layout')
                showLayoutDialog(src,tag,editor,ls,le,plugin);
            else if(action==='size')
                showSizeDialog(src,tag,editor,ls,le,plugin);
            else if(action==='spacing')
                showSpacingDialog(src,tag,editor,ls,le,plugin);
        }catch(err){
            console.warn('[Media] Action error:',err);
            try{new obs.Notice(tr(plugin,'failOp')+(err.message||err));}catch(e2){}
        }
    }

    // ============================================================
    //  DIALOG HELPERS
    // ============================================================
    function mkO(){var d=document.createElement('div');d.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.35);z-index:100002;display:flex;align-items:center;justify-content:center;';return d;}
    function mkD(){var d=document.createElement('div');d.style.cssText='background:var(--background-primary);border:1px solid var(--background-modifier-border);border-radius:10px;padding:20px 24px;min-width:320px;box-shadow:0 12px 48px rgba(0,0,0,.35);font-family:var(--font-interface,system-ui);';return d;}
    function mkT(t){var d=document.createElement('div');d.textContent=t;d.style.cssText='font-size:15px;font-weight:600;margin-bottom:16px;color:var(--text-normal);';return d;}
    function mkBR(){var r=document.createElement('div');r.style.cssText='display:flex;gap:8px;margin-top:16px;justify-content:flex-end;';for(var i=0;i<arguments.length;i++)r.appendChild(arguments[i]);return r;}
    function mkCB(t,f){var b=document.createElement('button');b.textContent=t;b.style.cssText='padding:8px 16px;border:1px solid var(--background-modifier-border);border-radius:5px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:12px;';b.addEventListener('click',f);return b;}
    function mkAB(t,f){var b=document.createElement('button');b.textContent=t;b.style.cssText='padding:8px 18px;border:none;border-radius:5px;background:var(--interactive-accent);color:#fff;cursor:pointer;font-size:12px;font-weight:600;';b.addEventListener('click',f);return b;}

    function showLayoutDialog(src,tag,ed,ls,le,pl){
        var obs=require('obsidian'),ov=mkO(),dlg=mkD();
        dlg.appendChild(mkT(tr(pl,'titleLayout')));
        var curT='grid',fl=src.split('\n')[0].trim();
        if(fl.indexOf('|')!==-1){var p=fl.split('|')[1];if(p){var tk=p.trim().toLowerCase();if(tk==='full'||tk==='list'||tk==='grid')curT=tk;}}
        [{v:'grid',lk:'layoutGrid'},{v:'list',lk:'layoutList'},{v:'full',lk:'layoutFull'}].forEach(function(opt){
            var a=curT===opt.v,row=document.createElement('div');
            row.setAttribute('data-lr','1');
            row.style.cssText='padding:10px;margin-bottom:8px;border:1px solid '+(a?'var(--interactive-accent)':'var(--background-modifier-border)')+';border-radius:6px;cursor:pointer;display:flex;align-items:center;gap:10px;'+(a?'background:var(--background-modifier-hover);':'');
            var rb=document.createElement('input');rb.type='radio';rb.name='lr';rb.checked=a;rb.value=opt.v;row.appendChild(rb);
            var sp=document.createElement('span');sp.textContent=tr(pl,opt.lk);sp.style.cssText='font-size:13px;color:var(--text-normal);';
            row.appendChild(sp);
            row.addEventListener('click',function(){rb.checked=true;dlg.querySelectorAll('[data-lr]').forEach(function(r){r.style.borderColor='var(--background-modifier-border)';r.style.background='';});row.style.borderColor='var(--interactive-accent)';row.style.background='var(--background-modifier-hover)';});
            dlg.appendChild(row);
        });
        dlg.appendChild(mkBR(mkCB(tr(pl,'cancel'),function(){ov.remove();}),mkAB(tr(pl,'apply'),function(){
            var sel=dlg.querySelector('input[name="lr"]:checked');if(!sel){ov.remove();return;}
            var nt=sel.value,ln=src.split('\n');
            if(ln.length>0){var ps=ln[0].split('|');if(ps.length>=2){var tk2=ps[1].trim().toLowerCase();if(isNaN(parseInt(tk2))){ps[1]=nt;}else{ps.splice(1,0,nt);}ln[0]=ps.join('|');}}
            applyNewSource(ln,ed,ls,le);ov.remove();new obs.Notice(tr(pl,'appliedLayout')+nt);
        })));
        ov.appendChild(dlg);ov.addEventListener('click',function(e){if(e.target===ov)ov.remove();});document.body.appendChild(ov);
    }

    function showSizeDialog(src,tag,ed,ls,le,pl){
        var obs=require('obsidian'),ov=mkO(),dlg=mkD();
        dlg.appendChild(mkT(tr(pl,'titleSize')));
        var size=220,fl=src.split('\n')[0].trim();
        if(fl.indexOf('|')!==-1){var pp=fl.split('|');if(pp[1]&&!isNaN(parseInt(pp[1])))size=parseInt(pp[1]);else if(pp[2]&&!isNaN(parseInt(pp[2])))size=parseInt(pp[2]);}
        var lb=document.createElement('label');lb.textContent=tr(pl,'labelSize');
        lb.style.cssText='display:block;font-size:12px;color:var(--text-muted);margin-bottom:6px;';dlg.appendChild(lb);
        var inp=document.createElement('input');inp.type='number';inp.value=size;inp.min=80;inp.max=800;
        inp.style.cssText='width:100%;padding:8px 10px;border:1px solid var(--background-modifier-border);border-radius:5px;background:var(--background-secondary);color:var(--text-normal);font-size:14px;box-sizing:border-box;margin-bottom:16px;';dlg.appendChild(inp);
        var pr=document.createElement('div');pr.style.cssText='display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;';
        [150,180,220,280,320].forEach(function(p){
            var b=document.createElement('button');b.textContent=p+'px';
            b.style.cssText='padding:6px 12px;border:1px solid var(--background-modifier-border);border-radius:4px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:12px;';
            b.addEventListener('click',function(){inp.value=p;});pr.appendChild(b);
        });dlg.appendChild(pr);
        dlg.appendChild(mkBR(mkCB(tr(pl,'cancel'),function(){ov.remove();}),mkAB(tr(pl,'apply'),function(){
            var ns=Math.max(80,Math.min(800,parseInt(inp.value)||size)),ln=src.split('\n');
            if(ln.length>0){var ps=ln[0].split('|');if(ps.length>=2){var tk3=ps[1].trim().toLowerCase();if(tk3==='full'||tk3==='list'||tk3==='grid'){while(ps.length<3)ps.push('');ps[2]=String(ns);}else{ps[1]=String(ns);}}else{ps.push(String(ns));}ln[0]=ps.join('|');}
            applyNewSource(ln,ed,ls,le);ov.remove();new obs.Notice(tr(pl,'appliedSize')+ns+'px');
        })));ov.appendChild(dlg);ov.addEventListener('click',function(e){if(e.target===ov)ov.remove();});document.body.appendChild(ov);
    }

    function showSpacingDialog(src,tag,ed,ls,le,pl){
        var obs=require('obsidian'),ov=mkO(),dlg=mkD();dlg.style.minWidth='360px';
        dlg.appendChild(mkT(tr(pl,'titleSpacing')));
        // read current spacing from key-value lines
        var sl=0,sr=0,ig=0;
        var lines=src.split('\n');
        for(var i=0;i<lines.length;i++){
            var ln=lines[i].trim();
            if(!ln||ln.charAt(0)==='#')continue;
            var ix=ln.indexOf(':');if(ix===-1)continue;
            var k=ln.substring(0,ix).trim().toLowerCase(),v=ln.substring(ix+1).trim();
            if(k==='spacingleft')sl=parseInt(v)||0;
            if(k==='spacingright')sr=parseInt(v)||0;
            if(k==='itemgap')ig=parseInt(v)||0;
        }
        var lb1=document.createElement('label');lb1.textContent=tr(pl,'labelSpacingLeft');
        lb1.style.cssText='display:block;font-size:12px;color:var(--text-muted);margin-bottom:6px;';dlg.appendChild(lb1);
        var li=document.createElement('input');li.type='number';li.value=sl;li.min=0;
        li.style.cssText='width:100%;padding:8px 10px;border:1px solid var(--background-modifier-border);border-radius:5px;background:var(--background-secondary);color:var(--text-normal);font-size:14px;box-sizing:border-box;margin-bottom:12px;';dlg.appendChild(li);
        var lb2=document.createElement('label');lb2.textContent=tr(pl,'labelSpacingRight');
        lb2.style.cssText='display:block;font-size:12px;color:var(--text-muted);margin-bottom:6px;';dlg.appendChild(lb2);
        var ri=document.createElement('input');ri.type='number';ri.value=sr;ri.min=0;
        ri.style.cssText='width:100%;padding:8px 10px;border:1px solid var(--background-modifier-border);border-radius:5px;background:var(--background-secondary);color:var(--text-normal);font-size:14px;box-sizing:border-box;margin-bottom:12px;';dlg.appendChild(ri);
        var lb3=document.createElement('label');lb3.textContent=tr(pl,'labelItemGap');
        lb3.style.cssText='display:block;font-size:12px;color:var(--text-muted);margin-bottom:6px;';dlg.appendChild(lb3);
        var gi=document.createElement('input');gi.type='number';gi.value=ig;gi.min=0;
        gi.style.cssText='width:100%;padding:8px 10px;border:1px solid var(--background-modifier-border);border-radius:5px;background:var(--background-secondary);color:var(--text-normal);font-size:14px;box-sizing:border-box;margin-bottom:16px;';dlg.appendChild(gi);
        var pr=document.createElement('div');pr.style.cssText='display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;';
        [0,5,10,15,20].forEach(function(p){
            var b=document.createElement('button');b.textContent=p+'px';
            b.style.cssText='padding:6px 12px;border:1px solid var(--background-modifier-border);border-radius:4px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:12px;';
            b.addEventListener('click',function(){li.value=p;ri.value=p;gi.value=p;});pr.appendChild(b);
        });dlg.appendChild(pr);
        var scCb=document.createElement('div');scCb.style.cssText='margin-bottom:16px;display:flex;align-items:center;gap:8px;';
        var scBox=document.createElement('input');scBox.type='checkbox';
        scBox.style.cssText='width:16px;height:16px;cursor:pointer;';
        scCb.appendChild(scBox);
        var scLabel=document.createElement('span');scLabel.textContent=tr(pl,'labelSmartCenter');
        scLabel.style.cssText='font-size:13px;color:var(--text-normal);cursor:pointer;';
        scCb.appendChild(scLabel);
        dlg.appendChild(scCb);
        dlg.appendChild(mkBR(mkCB(tr(pl,'cancel'),function(){ov.remove();}),mkAB(tr(pl,'apply'),function(){
            var nsl=parseInt(li.value)||0,nsr=parseInt(ri.value)||0,nig=parseInt(gi.value)||0;
            var newLines=[];
            var foundLeft=false,foundRight=false,foundGap=false;
            lines.forEach(function(ln){
                var trl=ln.trim();
                if(trl&&trl.charAt(0)!=='#'){
                    var ix=trl.indexOf(':');
                    if(ix!==-1){
                        var k=trl.substring(0,ix).trim().toLowerCase();
                        if(k==='spacingleft'){newLines.push('spacingleft: '+nsl);foundLeft=true;return;}
                        if(k==='spacingright'){newLines.push('spacingright: '+nsr);foundRight=true;return;}
                        if(k==='itemgap'){newLines.push('itemgap: '+nig);foundGap=true;return;}
                    }
                }
                newLines.push(ln);
            });
            if(!foundLeft)newLines.push('spacingleft: '+nsl);
            if(!foundRight)newLines.push('spacingright: '+nsr);
            if(!foundGap)newLines.push('itemgap: '+nig);
            applyNewSource(newLines,ed,ls,le);ov.remove();
            new obs.Notice(tr(pl,'appliedSpacing'));
        })));ov.appendChild(dlg);
        ov.addEventListener('click',function(e){if(e.target===ov)ov.remove();});document.body.appendChild(ov);
    }

    // ---- lightbox ----
    function openLightbox(src,name){
        var ex=document.querySelector('.v6-lightbox-overlay');if(ex)ex.remove();
        var ov=document.createElement('div');ov.className='v6-lightbox-overlay';
        ov.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.88);z-index:100000;display:flex;align-items:center;justify-content:center;cursor:pointer;';
        var content;
        if(/\.(mp4|webm|ogg|mov)$/i.test(src)){
            content=document.createElement('video');content.src=src;content.controls=true;content.autoplay=true;content.style.cssText='max-width:90vw;max-height:85vh;border-radius:8px;';
        }else if(/\.(mp3|wav|ogg|flac|m4a)$/i.test(src)){
            var box=document.createElement('div');box.style.cssText='display:flex;flex-direction:column;align-items:center;gap:16px;';
            var audio=document.createElement('audio');audio.src=src;audio.controls=true;audio.style.cssText='width:min(400px,80vw);';box.appendChild(audio);
            var lb=document.createElement('div');lb.textContent=name||'';lb.style.cssText='color:#ccc;font-size:14px;';box.appendChild(lb);
            content=box;
        }else{
            content=document.createElement('img');content.src=src;content.alt=name||'';content.style.cssText='max-width:90vw;max-height:85vh;object-fit:contain;border-radius:8px;';
        }
        if(!(content instanceof HTMLDivElement)){
            var wrapper=document.createElement('div');wrapper.style.cssText='display:flex;flex-direction:column;align-items:center;';wrapper.appendChild(content);
            if(content.tagName==='IMG'){var cap=document.createElement('div');cap.textContent=name||'';cap.style.cssText='color:#ccc;margin-top:12px;font-size:13px;';wrapper.appendChild(cap);}
            content=wrapper;
        }
        ov.appendChild(content);
        ov.addEventListener('click',function(e){if(e.target===ov||e.target===content)ov.remove();});
        document.addEventListener('keydown',function onEsc(ev){if(ev.key==='Escape'&&ov.parentNode){ov.remove();document.removeEventListener('keydown',onEsc);}});
        document.body.appendChild(ov);
    }

    // ---- show count helper ----
    function shouldShowCount(plugin) {
        try { var m = plugin.settings.modules; if (m && m['memories'] && m['memories'].showCount === false) return false; } catch(e) {}
        return true;
    }

    // ============================================================
    //  MAIN SETUP
    // ============================================================
    function setupMemoriesProcessor(plugin) {
        try {
            var obs = require('obsidian');
            var normalizePath = obs.normalizePath;
            var app = plugin.app;

            // --- context menu ---
            plugin.registerEvent(
                app.workspace.on('file-menu', function(menu,file){
                    var isFolder=file instanceof obs.TFolder;
                    var isMedia=file instanceof obs.TFile&&getMediaType(file.extension)!==null;
                    if(!isFolder&&!isMedia)return;
                    menu.addItem(function(item){
                        item.setTitle(tr(plugin,'insertMenu')).setIcon('film')
                            .onClick(function(){
                                var tp=isFolder?file.path:file.parent.path;
                                var ed=getEditorForInsert(app);
                                if(!ed){new obs.Notice(tr(plugin,'openNote'));return;}
                                var cur=ed.getCursor();
                                ed.replaceRange('\n```s\n'+tp+'\n```\n',cur);
                                new obs.Notice(tr(plugin,'insertDone'));
                            });
                    });
                })
            );

            // --- register ```s code block processor ---
            plugin.registerMarkdownCodeBlockProcessor('s', function(source, el, ctx) {
                var Child=function(c,src){
                    obs.MarkdownRenderChild.call(this,c);this.src=src;this.el=el;
                };
                Child.prototype=Object.create(obs.MarkdownRenderChild.prototype);
                Child.prototype.constructor=Child;

                Child.prototype.onload=function(){
                    var self=this;
                    try{
                        var container=self.containerEl;container.empty();container.className='v6-memories-gallery';
                        attachToolbar(container,self.el,self.src,'s',plugin);

                        var params=parseSource(self.src);
                        if(!params.paths||params.paths.length===0){
                            container.createEl('div',{cls:'v6-mem-error',text:tr(plugin,'errNoPath')});return;
                        }
                        var allFiles=[];
                        params.paths.forEach(function(rp){
                            var tp=normalizePath(rp),af=app.vault.getAbstractFileByPath(tp);if(!af)return;
                            if(af instanceof obs.TFolder)af.children.forEach(function(c){if(c instanceof obs.TFile&&getMediaType(c.extension))allFiles.push(c);});
                            else if(af instanceof obs.TFile&&getMediaType(af.extension))allFiles.push(af);
                        });
                        if(allFiles.length===0){container.createEl('div',{cls:'v6-mem-empty',text:tr(plugin,'errEmpty')});return;}
                        // sort
                        if(params.sort==='date-desc'||params.sort==='newest'||params.sort==='modified')
                            allFiles.sort(function(a,b){return b.stat.mtime-a.stat.mtime;});
                        else if(params.sort==='date-asc'||params.sort==='oldest')
                            allFiles.sort(function(a,b){return a.stat.mtime-b.stat.mtime;});
                        else if(params.sort==='name-desc'||params.sort==='reverse')
                            allFiles.sort(function(a,b){return b.name.localeCompare(a.name);});
                        else allFiles.sort(function(a,b){return a.name.localeCompare(b.name);});
                        if(params.limit>0&&allFiles.length>params.limit)allFiles=allFiles.slice(0,params.limit);

                        var isFull=params.type==='full',isList=params.type==='list',sz=params.size;

                        var grid=container.createEl('div',{
                            cls:isList?'v6-mem-list':'v6-mem-grid',
                            attr:{style:!isList?(params.smartCenter?'display:flex;flex-wrap:wrap;justify-content:center;gap:'+(params.itemGap||10)+'px;padding-left:'+(params.spacingLeft||0)+'px;padding-right:'+(params.spacingRight||0)+'px;':'display:grid;grid-template-columns:repeat(auto-fill,minmax('+sz+'px,1fr));gap:'+(params.itemGap||10)+'px;padding-left:'+(params.spacingLeft||0)+'px;padding-right:'+(params.spacingRight||0)+'px;'):''}
                        });

                        allFiles.forEach(function(file){
                            var rp=app.vault.getResourcePath(file),mt=getMediaType(file.extension);
                            if(isList){
                                var row=document.createElement('div');
                                row.style.cssText='display:flex;align-items:center;gap:10px;padding:8px 12px;border-bottom:1px solid var(--background-modifier-border);cursor:pointer;';
                                var tsz=Math.min(sz,60),thumb=document.createElement('div');
                                thumb.style.cssText='width:'+tsz+'px;height:'+tsz+'px;border-radius:6px;overflow:hidden;background:var(--background-secondary);flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:20px;color:var(--text-muted);';
                                if(mt==='image'){var im=document.createElement('img');im.src=rp;im.style.cssText='width:100%;height:100%;object-fit:cover;display:block;';thumb.appendChild(im);}
                                else thumb.textContent=mt==='video'?'\u25b6':'\u266a';
                                var info=document.createElement('span');info.textContent=file.name;
                                info.style.cssText='font-size:13px;color:var(--text-normal);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;';
                                row.appendChild(thumb);row.appendChild(info);grid.appendChild(row);
                                (function(r,n){row.addEventListener('click',function(){openLightbox(r,n);});})(rp,file.name);
                            }else{
                                var item=document.createElement('div');
                                item.style.cssText='border-radius:8px;overflow:hidden;background:var(--background-secondary);cursor:pointer;'+(isFull?'':'aspect-ratio:1;');
                                if(isFull)item.style.minHeight=sz+'px';
                                if(mt==='image'){
                                    var im2=document.createElement('img');im2.src=rp;im2.alt=file.name;
                                    im2.style.cssText='width:100%;height:100%;object-fit:cover;display:block;transition:transform .25s;';
                                    im2.draggable=false;im2.loading='lazy';item.appendChild(im2);
                                    item.addEventListener('mouseenter',function(){im2.style.transform='scale(1.04)';});
                                    item.addEventListener('mouseleave',function(){im2.style.transform='';});
                                }else if(mt==='video'){
                                    item.style.display='flex';item.style.alignItems='center';item.style.justifyContent='center';
                                    var vl=document.createElement('div');vl.textContent='\u25b6 VIDEO';
                                    vl.style.cssText='color:var(--text-muted);font-size:11px;font-weight:600;';item.appendChild(vl);
                                }else{
                                    item.style.display='flex';item.style.alignItems='center';item.style.justifyContent='center';
                                    var al=document.createElement('div');al.textContent='\u266a AUDIO';
                                    al.style.cssText='color:var(--text-muted);font-size:11px;font-weight:600;';item.appendChild(al);
                                }
                                (function(r,n){item.addEventListener('click',function(){openLightbox(r,n);});})(rp,file.name);
                                grid.appendChild(item);
                            }
                        });

                        // Show media count (controlled by setting)
                        if (shouldShowCount(plugin)) {
                            container.createEl('div',{cls:'v6-mem-count',text:allFiles.length+tr(plugin,'mediaCount')});
                        }
                    }catch(re){
                        console.warn('[memories] render error:',re);
                        self.containerEl.textContent=tr(plugin,'errRender')+(re.message||re);
                    }
                };

                ctx.addChild(new Child(el,source));
            });

            console.log('[Dashboard] media gallery processor registered (code: s)');
        } catch(e) {
            console.error('[Dashboard] setupMemoriesProcessor failed:', e);
        }
    }

    if (typeof global !== 'undefined') global.__setupMemoriesProcessor = setupMemoriesProcessor;
    else if (typeof window !== 'undefined') window.__setupMemoriesProcessor = setupMemoriesProcessor;
    if (typeof module !== 'undefined' && module.exports) module.exports = setupMemoriesProcessor;

})(typeof globalThis !== 'undefined' ? globalThis : (typeof window !== 'undefined' ? window : typeof global !== 'undefined' ? global : this));

// 工具函数：ArrayBuffer → Base64
function arrayBufferToBase64(buffer) {
    var binary = '';
    var bytes = new Uint8Array(buffer);
    for (var i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

class DashboardPlugin extends Plugin {
    async onload() {
        await this.loadSettings();

        // ★ 初始化语言（必须在所有 UI 之前）
        var lang = this.settings.language || 'system';
        setLanguage(lang);
        // 恢复 AI 自定义翻译（key 格式为 'custom_xx'，与 _currentLang 一致）
        if (lang.startsWith('custom_') && this.settings.aiCustomTranslations) {
            var saved = this.settings.aiCustomTranslations[lang];
            if (saved) {
                setCustomTranslations(lang, saved);
            }
        }

        // ★ 设置全局引用
        // __DBFV_PLUGIN__：供 DashboardFileViewer.canAcceptExtension 访问 settings
        // __PLUGIN_ID__：供库加载器读取 libs/ 本地库文件（离线可用）
        __DBFV_PLUGIN__ = this;
        window.__PLUGIN_ID__ = this.manifest.id;

        // ★ 动态加载 FileViewer 扩展（dev模式从磁盘读取，release模式已内联无需操作）
        try { await initFileViewers(this); } catch(e) { console.warn('[Dashboard] FV 加载异常:', e); }

        // ★ SheetJS 和 mammoth.js 预加载（requestUrl + new Function，后台异步不阻塞）
        startViewerLibPreload();

        this.moduleManager = new ModuleManager(this);

        await this.initModuleLayouts();

        this.registerView(VIEW_TYPE, (leaf) => new DashboardView(leaf, this));

        // ★ 文件查看器：接管非 md 文件扩展名，在 Obsidian 页签内直接预览
        // SheetJS → xlsx/xls/csv 渲染为表格；mammoth.js → docx 渲染为 HTML
        // HTML 直接 iframe；doc/ppt 旧版格式展示升级提示
        try {
            this.registerView(FILE_VIEWER_TYPE, (leaf) => new DashboardFileViewer(leaf, this));
        } catch(e) {
            console.warn('[Dashboard] FileViewer registration failed:', e.message);
        }
        var silentExtensions = [
            // 文档类
            'doc', 'docx', 'pdf', 'odt', 'rtf',
            // 表格类（SheetJS 渲染）
            'xlsx', 'xls', 'xlsm', 'csv', 'ods',
            // 演示
            'ppt', 'pptx',
            // 网页类（iframe 渲染）
            'html', 'htm',
            // 媒体类（video/audio HTML5 标签）
            'mp4', 'webm', 'mov', 'avi', 'mkv', 'm4v',
            'mp3', 'wav', 'ogg', 'flac', 'm4a', 'aac',
            // 数据/配置文件
            'json', 'yaml', 'yml', 'xml', 'txt', 'toml', 'ini', 'cfg', 'env',
            // 代码文件
            'js', 'ts', 'jsx', 'tsx', 'css', 'scss', 'less',
            'py', 'rb', 'java', 'go', 'rs', 'c', 'cpp', 'cs', 'sh', 'bat', 'ps1'
        ];
        var plugin = this;
        silentExtensions.forEach(function(ext) {
            try {
                plugin.registerExtensions([ext], FILE_VIEWER_TYPE);
            } catch(e) {
                // 扩展名冲突（其他插件已注册），静默跳过
            }
        });

        this.addRibbonIcon('layout-dashboard', t('app.dashboard'), () => this.activateView());

        this.addCommand({
            id: 'open-dashboard',
            name: t('main.command'),
            callback: () => this.activateView()
        });

        this.addSettingTab(new DashboardSettingTab(this.app, this));

        this.app.workspace.onLayoutReady(() => {
            this.activateView();
        });

        // ============ 后台功能：功能性插件（不显示面板，后台运作）============

        // 1. 文件夹计数器：在文件浏览器中显示文件/文件夹数量
        this._setupFolderCounter();

        // 2. 表格列宽调整：为笔记中的表格添加拖拽调整列宽功能
        this._setupTableResize();

        // 3. 自动播放引擎：使用 IntersectionObserver 检测 video/audio 并自动播放
        this._setupAutoplayEngine();

        // 4. 图片处理右键菜单：为图片文件右键菜单添加格式转换/压缩选项
        this._setupImageTools();

        // 6. Excel转表格：全局粘贴拦截器（因已标记为 UTILITY_MODULE，
        //    render() 不被调用，需在此手动注册 paste 监听）
        this._setupExcelToMarkdown();

        // 7. Gallery 处理器（img-gallery + memories）：
        //    dev 模式：从 src/utils/ 动态加载 → 改代码只需 Ctrl+R
        //    release 模式：build 时已内联到 main.js
        try { await initGalleryProcessors(this); } catch(e) {
            console.warn('[Dashboard] Gallery processors init failed:', e);
        }

        // 7b. 图片画廊：注册 ```img-gallery 代码块处理器（根据设置开关）
        try {
            var igSet = this.settings.modules['img-gallery'];
            if (!igSet || igSet.enabled !== false) {
                if (typeof window.__setupImgGalleryProcessor === 'function') {
                    window.__setupImgGalleryProcessor(this);
                }
            }
        } catch(e) {
            console.warn('[Dashboard] img-gallery processor init failed:', e);
        }

        // 8. 媒体画廊：注册 ```memories 代码块处理器（根据设置开关）
        try {
            var memSet = this.settings.modules['memories'];
            if (!memSet || memSet.enabled !== false) {
                if (typeof window.__setupMemoriesProcessor === 'function') {
                    window.__setupMemoriesProcessor(this);
                }
            }
        } catch(e) {
            console.warn('[Dashboard] memories processor init failed:', e);
        }
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

        // 分配默认布局（基础模块按 5 列排列，跳过全局功能型模块 + 已关闭的模块）
        var self = this;
        var dashboardOnlyIds = loadedIds.filter(function(id) {
            if (UTILITY_MODULE_IDS.indexOf(id) !== -1) return false;
            // ★ 跳过已关闭的模块 — 不分配布局位置
            var ms = self.settings.modules[id];
            if (ms && ms.enabled === false) return false;
            return true;
        });
        dashboardOnlyIds.forEach((moduleId, index) => {
            if (!this.settings.layout[moduleId]) {
                const COLS = 5;
                const CARD_W = 280;
                const CARD_H = 250;
                const GAP_X = 12;
                const GAP_Y = 12;
                const MARGIN = 20;
                const col = index % COLS;
                const row = Math.floor(index / COLS);
                this.settings.layout[moduleId] = {
                    x: MARGIN + col * (CARD_W + GAP_X),
                    y: MARGIN + row * (CARD_H + GAP_Y),
                    width: CARD_W,
                    height: CARD_H
                };
                changed = true;
            }
        });

        // 实例布局：5 列排列，放在基础模块下方
        instances.forEach((inst, index) => {
            if (!this.settings.layout[inst.id]) {
                const COLS = 5;
                const CARD_W = 280;
                const CARD_H = 250;
                const GAP_X = 12;
                const GAP_Y = 12;
                const MARGIN = 20;
                const totalBaseRows = Math.ceil(dashboardOnlyIds.length / COLS);
                const col = index % COLS;
                const row = totalBaseRows + Math.floor(index / COLS);
                this.settings.layout[inst.id] = {
                    x: MARGIN + col * (CARD_W + GAP_X),
                    y: MARGIN + row * (CARD_H + GAP_Y),
                    width: CARD_W,
                    height: CARD_H
                };
                changed = true;
            }
        });

        // ★ 清理已关闭模块的布局条目（避免空占位）
        Object.keys(this.settings.layout).forEach(moduleId => {
            var modSet = this.settings.modules[moduleId];
            if (modSet && modSet.enabled === false) {
                delete this.settings.layout[moduleId];
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
        // 清理后台功能
        if (this._folderCounterObserver) { this._folderCounterObserver.disconnect(); }
        if (this._fcObserver) { this._fcObserver.disconnect(); }
        if (this._fcTimeout) { clearTimeout(this._fcTimeout); }
        if (this._autoplayObserver) { this._autoplayObserver.disconnect(); }
        if (this._autoplayInterval) { clearInterval(this._autoplayInterval); }
        if (this._systemTray) {
            try { this._systemTray.destroy(); } catch(e) {}
        }
        this.app.workspace.detachLeavesOfType(VIEW_TYPE);
    }

    // ============ 后台功能：Excel 转 Markdown 表格 ============
    // ★ excel-to-markdown 是 UTILITY_MODULE，其 render() 已被屏蔽。
    //    但它的全局粘贴拦截器是关键功能，必须在此手动注册。
    _setupExcelToMarkdown() {
        var plugin = this;
        var modId = 'excel-to-markdown';

        // 获取模块设置（如果不存在则用默认值）
        function getModSettings() {
            return (plugin.settings.modules && plugin.settings.modules[modId])
                || { enabled: true, enabledAutoConvert: true };
        }

        // 判断剪贴板数据是否为Excel表格
        function isExcelData(text) {
            if (!text) return false;
            var trimmed = text.trim();
            if (!trimmed) return false;
            if (trimmed.indexOf('\t') === -1) return false;
            var firstLine = trimmed.split(/[\n\r]/)[0] || '';
            return firstLine.split('\t').length >= 2;
        }

        // 单元格内换行处理
        function replaceIntraCellNewline(data) {
            return data.replace(/"([^\t]*(?<=[^\r])\n[^\t]*)"/g, function(match) {
                return match.slice(1, -1).replace(/""/g, '"').replace(/\n/g, '<br/>');
            });
        }

        // 列对齐
        var COL_ALIGN_REGEX = /^(\^[lcr])/i;

        function getColumnWidthsAndAlignments(rows) {
            var colAlignments = [];
            var columnWidths = rows[0].map(function(col, i) {
                var align = 'l';
                var m = col.match(COL_ALIGN_REGEX);
                if (m) {
                    align = m[1][1].toLowerCase();
                    rows[0][i] = col.replace(COL_ALIGN_REGEX, '');
                }
                colAlignments.push(align);
                return Math.max.apply(null, rows.map(function(r) { return String(r[i] || '').length; }));
            });
            return { columnWidths: columnWidths, colAlignments: colAlignments };
        }

        function excelToMarkdown(rawData) {
            var data = rawData.trim();
            if (!data) return null;
            data = replaceIntraCellNewline(data);
            var rows = data.split(/[\n\u0085\u2028\u2029]|\r\n?/g).map(function(r) { return r.split('\t'); });
            if (!rows[0] || rows[0].length < 2) return null;
            rows = rows.filter(function(r) { return r.some(function(c) { return c.trim(); }); });
            if (rows.length === 0) return null;
            var sizes = getColumnWidthsAndAlignments(rows);
            var colWidths = sizes.columnWidths;
            var colAlignments = sizes.colAlignments;
            var mdRows = rows.map(function(row) {
                return '| ' + row.map(function(col, i) {
                    return String(col).replace(/\|/g, '\\|') + ' '.repeat(Math.max(0, colWidths[i] - String(col).length + 1));
                }).join(' | ') + ' |';
            });
            var ALIGN_MAP = { l: ' ', r: ':', c: ':' };
            var ALIGN_POST = { l: ' ', r: '', c: ':' };
            var alignRow = '|' + colWidths.map(function(w, i) {
                var a = colAlignments[i] || 'l';
                return ALIGN_MAP[a] + '-'.repeat(w + 2) + ALIGN_POST[a];
            }).join('|') + '|';
            mdRows.splice(1, 0, alignRow);
            return mdRows.join('\n');
        }

        // 注册全局粘贴拦截器
        function registerPasteHandler() {
            if (plugin._excelPasteRef) {
                try { plugin.app.workspace.offref(plugin._excelPasteRef); } catch(e) {}
                plugin._excelPasteRef = null;
            }
            try {
                plugin._excelPasteRef = plugin.app.workspace.on('editor-paste', function(evt, editor, view) {
                    var ms = getModSettings();
                    if (!ms.enabledAutoConvert) return;
                    var clipboardData = evt.clipboardData;
                    if (!clipboardData) return;
                    var text = clipboardData.getData('text/plain');
                    if (!isExcelData(text)) return;
                    var mdTable = excelToMarkdown(text);
                    if (!mdTable) return;
                    evt.preventDefault();
                    try {
                        editor.replaceSelection(mdTable + '\n');
                        new Notice('✓ Excel数据已自动转为Markdown表格');
                    } catch(e2) {
                        console.error('[Dashboard] Excel->Markdown 插入失败:', e2);
                    }
                });
            } catch(e) {
                console.error('[Dashboard] 注册 Excel 粘贴处理失败:', e);
            }
        }

        // 首次加载：自动注册（如果开关开启）
        var initMs = getModSettings();
        if (initMs.enabled && initMs.enabledAutoConvert !== false) {
            registerPasteHandler();
        }

        // ★ 当用户在设置中切换开关时，插件会重新调用 _setup 方法，
        //    但更简单的方案：在 _renderUtilityToggles 切换时触发 refreshView，
        //    这里我们只在首次启用时注册。后续切换由模块自己处理。
        //    同时我们注册一个事件：每当 settings 变化时检查是否需要重新注册。
        plugin._excelPasteRef = plugin._excelPasteRef || null;
    }

    // ============ 后台功能：文件夹计数器 ============
    // ★ 仿原插件 file-explorer-note-count：
    //    1. 启动时预扫描所有文件夹计数
    //    2. 设置 data-count 属性到 .nav-folder-title
    //    3. CSS ::after 伪元素显示计数
    //    4. MutationObserver + layout-change 监听 DOM 变化
    _setupFolderCounter() {
        var plugin = this;

        // ★ 注入样式（仿原插件风格）
        if (!document.getElementById('__dfc_styles__')) {
            var styleEl = document.createElement('style');
            styleEl.id = '__dfc_styles__';
            styleEl.textContent = [
                '.nav-folder-title[data-count]::after {',
                '  content: attr(data-count);',
                '  display: inline-block;',
                '  font-size: 11px;',
                '  margin-left: 6px;',
                '  padding: 1px 5px;',
                '  border-radius: 3px;',
                '  background: var(--background-modifier-hover);',
                '  color: var(--text-muted);',
                '  line-height: 1.4;',
                '  order: 2;',
                '  white-space: nowrap;',
                '}',
                '.nav-folder-title-content { flex-grow: 1; }',
                '.nav-folder-title { display: flex; align-items: center; }',
                '.nav-folder-title .nav-folder-collapse-indicator { order: 3; }',
            ].join('\n');
            document.head.appendChild(styleEl);
        }

        // ★ 预扫描：从 vault 中一次性获取所有文件夹及其文件/子目录数
        var folderCountMap = {};

        function preScanVault() {
            folderCountMap = {};
            try {
                // ★ 关键修复：用 getFiles() 替代 getAllLoadedFiles()
                // getAllLoadedFiles() 只返回已加载到元数据缓存的页面，
                // 未浏览过的文件夹中的文件不会被计入，导致统计错误。
                // getFiles() 返回 vault 中所有文件（含图片/PDF 等非 md 文件）
                var allFiles = plugin.app.vault.getFiles();
                allFiles.forEach(function(file) {
                    // 确定父路径
                    var fpath = file.path || '';
                    var idx = fpath.lastIndexOf('/');
                    var parentPath = (idx > 0) ? fpath.substring(0, idx) : '/';
                    if (!folderCountMap[parentPath]) {
                        folderCountMap[parentPath] = { files: 0, dirs: 0 };
                    }
                    folderCountMap[parentPath].files++;
                });

                // ★ 同时遍历所有文件夹自身，获取子目录计数
                var allFolders = plugin.app.vault.getAllLoadedFiles().filter(function(item) {
                    return item.children !== undefined;
                });
                allFolders.forEach(function(folder) {
                    var fpath = folder.path || '';
                    var idx = fpath.lastIndexOf('/');
                    var parentPath = (idx > 0) ? fpath.substring(0, idx) : '/';
                    if (!folderCountMap[parentPath]) {
                        folderCountMap[parentPath] = { files: 0, dirs: 0 };
                    }
                    folderCountMap[parentPath].dirs++;
                    // 确保文件夹自身也有条目（目录本身不计 files，只计子目录中的 files）
                    var selfPath = fpath || '/';
                    if (!folderCountMap[selfPath]) {
                        folderCountMap[selfPath] = { files: 0, dirs: 0 };
                    }
                });

                console.log('[Dashboard] 文件夹计数器扫描完成, 文件夹数:', Object.keys(folderCountMap).length);
            } catch(e) {
                console.log('[Dashboard] 文件夹计数器预扫描失败:', e.message);
            }
        }

        function getCountText(folderPath) {
            if (folderPath === '') folderPath = '/';
            var entry = folderCountMap[folderPath];
            if (!entry || (entry.files <= 0 && entry.dirs <= 0)) return null;
            // ★ 纯文字显示：仅显示文件数，目录数在子目录展开时更清晰
            return String(entry.files);
        }

        // ★ 更新 DOM：遍历所有文件夹标题，设置 data-count
        function updateCounts() {
            try {
                // 检查功能是否启用
                if (!plugin.settings.modules['folder-counter'] || plugin.settings.modules['folder-counter'].enabled === false) {
                    document.querySelectorAll('.nav-folder-title[data-count]').forEach(function(el) {
                        el.removeAttribute('data-count');
                    });
                    return;
                }

                // 每30秒重新扫描
                if (!plugin._fcLastScan || (Date.now() - plugin._fcLastScan) > 30000) {
                    preScanVault();
                    plugin._fcLastScan = Date.now();
                }

                document.querySelectorAll('.nav-files-container').forEach(function(container) {
                    container.querySelectorAll('.nav-folder-title').forEach(function(folderEl) {
                        var parentFolder = folderEl.closest('.nav-folder');
                        if (!parentFolder) return;
                        // ★ 修复：优先从 .nav-folder-title 读取 data-path
                        // .nav-folder 上的 data-path 在某些 Obsidian 版本中可能为空
                        var fp = folderEl.getAttribute('data-path') || parentFolder.getAttribute('data-path') || '';
                        // 如果仍然是空，尝试从该父链上的 .nav-folder 元素拼接路径
                        if (!fp) {
                            var pathParts = [];
                            var cur = parentFolder;
                            while (cur && cur.classList.contains('nav-folder')) {
                                var dp = cur.getAttribute('data-path');
                                if (dp && dp !== '/') pathParts.unshift(dp);
                                cur = cur.parentElement ? cur.parentElement.closest('.nav-folder') : null;
                            }
                            fp = pathParts.length > 0 ? pathParts.join('/') : '';
                        }
                        if (!fp || fp === '/') {
                            if (parentFolder.classList.contains('nav-folder-root')) fp = '/';
                        }
                        // debug: 打印前几个路径
                        if (!plugin._fcDebugged) {
                            console.log('[Dashboard] FC debug: fp="' + fp + '", mapKeys:', Object.keys(folderCountMap).slice(0, 5).join(', '));
                            plugin._fcDebugged = true;
                        }
                        var countText = getCountText(fp);
                        if (countText) {
                            folderEl.setAttribute('data-count', countText);
                        } else {
                            folderEl.removeAttribute('data-count');
                        }
                    });
                });
            } catch(e) {
                console.log('[Dashboard] 文件夹计数器更新失败:', e.message);
            }
        }

        // 初始化扫描
        preScanVault();
        plugin._fcLastScan = Date.now();

        // MutationObserver
        var fo = new MutationObserver(function() {
            if (plugin._fcTimeout) clearTimeout(plugin._fcTimeout);
            plugin._fcTimeout = setTimeout(updateCounts, 300);
        });

        var startObserving = function() {
            var target = document.querySelector('.nav-files-container');
            if (target) {
                fo.observe(target, { childList: true, subtree: true });
                updateCounts();
            }
        };

        setTimeout(startObserving, 1500);

        plugin.registerEvent(
            plugin.app.workspace.on('layout-change', function() {
                preScanVault();
                plugin._fcLastScan = Date.now();
                setTimeout(startObserving, 500);
            })
        );

        var fcInterval = setInterval(function() {
            preScanVault();
            plugin._fcLastScan = Date.now();
            updateCounts();
        }, 15000);
        plugin.registerInterval(fcInterval);

        plugin._fcObserver = fo;
        plugin._fcTimeout = null;
        plugin._fcLastScan = Date.now();
    }

    // ============ 后台功能：表格列宽调整 ============
    _setupTableResize() {
        var plugin = this;

        // 监听渲染后的 markdown，为表格添加列宽调整
        var observer = new MutationObserver(function(mutations) {
            mutations.forEach(function(mutation) {
                mutation.addedNodes.forEach(function(node) {
                    if (node.nodeType !== 1) return;
                    // 查找新添加的表格
                    var tables = node.querySelectorAll ? node.querySelectorAll('table') : [];
                    if (node.tagName === 'TABLE') tables = [node].concat(Array.from(tables));

                    tables.forEach(function(table) {
                        if (table.hasAttribute('data-tcr-setup')) return;
                        table.setAttribute('data-tcr-setup', '1');
                        plugin._enableTableResize(table);
                    });
                });
            });
        });

        observer.observe(document.body, { childList: true, subtree: true });
        this._folderCounterObserver = observer; // 复用 observer 引用
    }

    _enableTableResize(table) {
        var headerRow = table.querySelector('thead tr') || table.querySelector('tr');
        if (!headerRow) return;

        var headerCells = headerRow.querySelectorAll('th, td');
        if (headerCells.length === 0) return;

        // 计算可用宽度
        var availWidth = table.getBoundingClientRect().width || table.parentElement.clientWidth || 800;
        var cellWidth = Math.floor(availWidth / headerCells.length);

        // 设置列宽样式
        headerCells.forEach(function(cell) {
            cell.style.width = cellWidth + 'px';
            cell.style.minWidth = '50px';
            cell.style.position = 'relative';

            // 拖拽手柄
            var handle = document.createElement('div');
            handle.className = 'tcr-resize-handle';
            handle.style.cssText = 'position:absolute;right:0;top:0;bottom:0;width:5px;cursor:col-resize;z-index:1;';
            handle.addEventListener('mousedown', function(e) {
                e.preventDefault();
                e.stopPropagation();
                var startX = e.clientX;
                var startW = cell.offsetWidth;

                function onMove(ev) {
                    var dx = ev.clientX - startX;
                    var newW = Math.max(50, startW + dx);
                    cell.style.width = newW + 'px';
                }
                function onUp() {
                    document.removeEventListener('mousemove', onMove);
                    document.removeEventListener('mouseup', onUp);
                    // 应用到该列所有行
                    var colIdx = Array.from(headerRow.children).indexOf(cell);
                    table.querySelectorAll('tr').forEach(function(row) {
                        var cells = row.querySelectorAll('th, td');
                        if (cells[colIdx]) cells[colIdx].style.width = cell.style.width;
                    });
                }
                document.addEventListener('mousemove', onMove);
                document.addEventListener('mouseup', onUp);
            });
            cell.appendChild(handle);
        });
    }

    // ============ 后台功能：自动播放引擎 ============
    _setupAutoplayEngine() {
        var plugin = this;
        var settings = plugin.settings.modules['autoplay-loop'] || DEFAULT_SETTINGS.modules['autoplay-loop'];

        // 如果自动播放被禁用，不启动
        if (settings.enabled === false) return;

        var visibleMedia = new Map();

        function applyToElement(el) {
            var tag = el.tagName.toLowerCase();
            var isVideo = tag === 'video';
            var isAudio = tag === 'audio';

            if (isVideo && settings.autoplayVideo) {
                el.setAttribute('autoplay', '');
                if (settings.muteAutoplayedVideo) el.muted = true;
                if (settings.loopVideo) el.loop = true;
                try { el.play(); } catch(e) {}
            }
            if (isAudio && settings.autoplayAudio) {
                el.setAttribute('autoplay', '');
                if (settings.muteAutoplayedAudio) el.muted = true;
                if (settings.loopAudio) el.loop = true;
                try { el.play(); } catch(e) {}
            }
        }

        var observer = new IntersectionObserver(function(entries) {
            entries.forEach(function(entry) {
                var el = entry.target;
                var tag = el.tagName.toLowerCase();

                if (entry.isIntersecting && entry.intersectionRatio >= 0.3) {
                    visibleMedia.set(el, entry.intersectionRatio);
                    applyToElement(el);
                } else {
                    visibleMedia.delete(el);
                    if (tag === 'video' && settings.pauseOutOfViewVideo) el.pause();
                    if (tag === 'audio' && settings.pauseOutOfViewAudio) el.pause();
                }
            });
        }, { threshold: [0, 0.25, 0.5, 0.75, 1] });

        // 扫描媒体元素
        function scan() {
            document.querySelectorAll('video, audio').forEach(function(el) {
                if (!el.dataset.apEngineObserved) {
                    el.dataset.apEngineObserved = '1';
                    observer.observe(el);
                    applyToElement(el);
                }
            });
        }

        scan();
        var scanInterval = setInterval(scan, 2000);
        plugin._autoplayObserver = observer;
        plugin._autoplayInterval = scanInterval;
    }

    // ============ 后台功能：图片处理右键菜单 ============
    _setupImageTools() {
        var plugin = this;

        var imgExts = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'ico', 'tif', 'tiff', 'heic', 'avif'];

        plugin.registerEvent(
            plugin.app.workspace.on('file-menu', function(menu, file) {
                if (!plugin.settings.modules['image-tools'] || plugin.settings.modules['image-tools'].enabled === false) return;
                if (imgExts.indexOf(file.extension.toLowerCase()) === -1) return;

                var modSettings = plugin.settings.modules['image-tools'] || {};

                // ====== 分隔 ======
                menu.addSeparator();

                // ====== 在新窗口中打开 ======
                menu.addItem(function(item) {
                    item.setTitle('在新窗口中打开')
                        .setIcon('popup-open')
                        .onClick(async function() {
                            var filePath = plugin.app.vault.adapter.getFullPath(file.path);
                            require('electron').shell.openPath(filePath);
                        });
                });

                // ====== 复制图片 ======
                menu.addItem(function(item) {
                    item.setTitle('复制图片')
                        .setIcon('copy')
                        .onClick(async function() {
                            try {
                                var buff = await plugin.app.vault.readBinary(file);
                                var blob = new Blob([buff], { type: 'image/' + file.extension.toLowerCase() });
                                await navigator.clipboard.write([
                                    new ClipboardItem({ [blob.type]: blob })
                                ]);
                                new Notice('已复制图片: ' + file.name);
                            } catch(e) {
                                new Notice('复制失败: ' + e.message);
                            }
                        });
                });

                // ====== 复制为Base64编码图片 ======
                menu.addItem(function(item) {
                    item.setTitle('复制为Base64编码图片')
                        .setIcon('code-2')
                        .onClick(async function() {
                            try {
                                var buff = await plugin.app.vault.readBinary(file);
                                var base64 = arrayBufferToBase64(buff);
                                var mimeType = 'image/' + (file.extension.toLowerCase() === 'jpg' ? 'jpeg' : file.extension.toLowerCase());
                                var dataUri = 'data:' + mimeType + ';base64,' + base64;
                                await navigator.clipboard.writeText(dataUri);
                                new Notice('已复制 Base64 编码: ' + file.name);
                            } catch(e) {
                                new Notice('Base64复制失败: ' + e.message);
                            }
                        });
                });

                // ====== 格式转换（自定义参数） ======
                menu.addSeparator();
                menu.addItem(function(item) {
                    item.setTitle('格式转换/压缩（自定义参数）')
                        .setIcon('wrench')
                        .onClick(async function() {
                            await plugin._openImageConvertModal(file, modSettings);
                        });
                });

                // ====== 快捷预设 ======
                menu.addItem(function(item) {
                    item.setTitle('快捷转换：→ WebP (质量75)')
                        .setIcon('image-file')
                        .onClick(async function() {
                            await plugin._convertImageLocal(file, 'webp', 75, null);
                        });
                });

                menu.addItem(function(item) {
                    item.setTitle('快捷转换：→ JPEG (质量85)')
                        .setIcon('image-file')
                        .onClick(async function() {
                            await plugin._convertImageLocal(file, 'jpeg', 85, null);
                        });
                });

                menu.addItem(function(item) {
                    item.setTitle('快捷转换：→ PNG')
                        .setIcon('image-file')
                        .onClick(async function() {
                            await plugin._convertImageLocal(file, 'png', 100, null);
                        });
                });

                // ====== 旋转/翻转 ======
                menu.addSeparator();
                menu.addItem(function(item) {
                    item.setTitle('旋转/翻转')
                        .setIcon('rotate-cw')
                        .setDisabled(true);
                });

                menu.addItem(function(item) {
                    item.setTitle('  顺时针旋转 90°')
                        .setIcon('rotate-cw')
                        .onClick(async function() {
                            await plugin._modifyImageLocal(file, 'rotate', { degrees: 90 });
                        });
                });

                menu.addItem(function(item) {
                    item.setTitle('  逆时针旋转 90°')
                        .setIcon('rotate-ccw')
                        .onClick(async function() {
                            await plugin._modifyImageLocal(file, 'rotate', { degrees: -90 });
                        });
                });

                menu.addItem(function(item) {
                    item.setTitle('  旋转 180°')
                        .setIcon('rotate-ccw')
                        .onClick(async function() {
                            await plugin._modifyImageLocal(file, 'rotate', { degrees: 180 });
                        });
                });

                menu.addItem(function(item) {
                    item.setTitle('  水平翻转')
                        .setIcon('flip-horizontal')
                        .onClick(async function() {
                            await plugin._modifyImageLocal(file, 'flip', { direction: 'horizontal' });
                        });
                });

                menu.addItem(function(item) {
                    item.setTitle('  垂直翻转')
                        .setIcon('flip-vertical')
                        .onClick(async function() {
                            await plugin._modifyImageLocal(file, 'flip', { direction: 'vertical' });
                        });
                });
            })
        );
    }

    // ★ 图片处理模态框（自定义参数）
    async _openImageConvertModal(file, modSettings) {
        var plugin = this;
        var Modal = require('obsidian').Modal;

        var modal = new Modal(plugin.app);
        modal.titleEl.setText('图片处理 - ' + file.name);

        var contentEl = modal.contentEl;
        contentEl.style.cssText = 'padding:16px;min-width:360px;';

        // 预览
        var previewSection = contentEl.createDiv({ attr: { style: 'text-align:center;margin-bottom:14px;' } });
        previewSection.createEl('div', { text: '📸 ' + file.name, attr: { style: 'font-weight:600;margin-bottom:8px;color:var(--text-normal);' } });

        // 目标格式
        var formatSection = contentEl.createDiv({ attr: { style: 'margin-bottom:12px;' } });
        formatSection.createEl('div', { text: '目标格式：', attr: { style: 'font-size:13px;font-weight:600;margin-bottom:4px;color:var(--text-normal);' } });
        var formatRow = formatSection.createDiv({ attr: { style: 'display:flex;gap:8px;flex-wrap:wrap;' } });
        var formats = [
            { value: 'webp', label: 'WebP' },
            { value: 'jpeg', label: 'JPEG' },
            { value: 'png', label: 'PNG' },
            { value: 'same', label: '保持原格式' }
        ];
        var selectedFormat = modSettings.outputFormat || 'same';
        var formatRadios = [];
        formats.forEach(function(f) {
            var btn = formatRow.createEl('label', { attr: { style: 'font-size:12px;cursor:pointer;padding:4px 10px;border:1px solid var(--background-modifier-border);border-radius:4px;' + (selectedFormat === f.value ? 'background:var(--v6-primary);color:#fff;border-color:var(--v6-primary);' : '') } });
            var radio = btn.createEl('input', { attr: { type: 'radio', name: 'imgFormat', value: f.value, style: 'margin-right:4px;' } });
            if (f.value === selectedFormat) radio.checked = true;
            btn.appendText(f.label);
            formatRadios.push(radio);
            btn.addEventListener('click', function(e) {
                e.preventDefault();
                radio.checked = true;
                formatRadios.forEach(function(r, i) {
                    var lbl = r.parentElement;
                    if (r === radio) {
                        lbl.style.background = 'var(--v6-primary)';
                        lbl.style.color = '#fff';
                        lbl.style.borderColor = 'var(--v6-primary)';
                    } else {
                        lbl.style.background = '';
                        lbl.style.color = '';
                        lbl.style.borderColor = 'var(--background-modifier-border)';
                    }
                });
            });
        });

        // 质量
        var qualitySection = contentEl.createDiv({ attr: { style: 'margin-bottom:12px;' } });
        qualitySection.createEl('div', { text: '输出质量（1-100）：', attr: { style: 'font-size:13px;font-weight:600;margin-bottom:4px;color:var(--text-normal);' } });
        var qualityRow = qualitySection.createDiv({ attr: { style: 'display:flex;align-items:center;gap:8px;' } });
        var qualitySlider = qualityRow.createEl('input', { attr: { type: 'range', min: '1', max: '100', value: String(modSettings.quality || 75), style: 'flex:1;' } });
        var qualityLabel = qualityRow.createEl('span', { text: String(modSettings.quality || 75), attr: { style: 'min-width:30px;text-align:center;font-weight:600;font-size:14px;' } });
        qualitySlider.addEventListener('input', function() { qualityLabel.textContent = qualitySlider.value; });

        // 最大宽度
        var widthSection = contentEl.createDiv({ attr: { style: 'margin-bottom:12px;' } });
        widthSection.createEl('div', { text: '最大宽度（px，0=不限制）：', attr: { style: 'font-size:13px;font-weight:600;margin-bottom:4px;color:var(--text-normal);' } });
        var widthInput = widthSection.createEl('input', { attr: { type: 'number', value: String(modSettings.resizeWidth || 0), min: '0', max: '10000', step: '100', style: 'width:100%;box-sizing:border-box;padding:6px 8px;border:1px solid var(--background-modifier-border);border-radius:4px;background:var(--background-modifier-form-field);color:var(--text-normal);' } });

        // 是否覆盖原文件
        var overwriteSection = contentEl.createDiv({ attr: { style: 'margin-bottom:14px;' } });
        var overwriteLabel = overwriteSection.createEl('label', { attr: { style: 'font-size:13px;cursor:pointer;display:flex;align-items:center;gap:6px;color:var(--text-muted);' } });
        var overwriteCheckbox = overwriteLabel.createEl('input', { attr: { type: 'checkbox', checked: true, style: 'margin:0;' } });
        overwriteLabel.appendText('直接修改原文件（覆盖）');
        overwriteLabel.createEl('div', { text: '取消勾选则另存为新文件', attr: { style: 'font-size:11px;color:var(--text-faint);margin-left:22px;' } });

        // 操作按钮
        var btnRow = contentEl.createDiv({ attr: { style: 'display:flex;gap:10px;justify-content:flex-end;margin-top:16px;' } });

        var cancelBtn = btnRow.createEl('button', { text: '取消', attr: { style: 'padding:8px 16px;border:1px solid var(--background-modifier-border);border-radius:4px;background:var(--background-secondary);color:var(--text-normal);cursor:pointer;font-size:13px;' } });
        cancelBtn.addEventListener('click', function() { modal.close(); });

        var processBtn = btnRow.createEl('button', { text: '开始处理', attr: { style: 'padding:8px 20px;border:none;border-radius:4px;background:var(--v6-primary);color:#fff;cursor:pointer;font-size:13px;font-weight:600;' } });
        processBtn.addEventListener('click', async function() {
            modal.close();
            var fmt = formatRadios.find(function(r) { return r.checked; });
            var targetFormat = fmt ? fmt.value : 'same';
            if (targetFormat === 'same') targetFormat = file.extension.toLowerCase();
            var quality = parseInt(qualitySlider.value) || 75;
            var maxWidth = parseInt(widthInput.value) || 0;
            var overwrite = overwriteCheckbox.checked;

            await plugin._processImageLocal(file, targetFormat, quality, maxWidth, overwrite);
        });

        modal.open();
    }

    // ★ 核心处理函数：直接修改本地文件
    async _processImageLocal(file, targetFormat, quality, maxWidth, overwrite) {
        var plugin = this;
        try {
            new Notice('⏳ 正在处理图片...');
            var arrayBuffer = await plugin.app.vault.readBinary(file);
            var blob = new Blob([arrayBuffer]);
            var img = new Image();
            var url = URL.createObjectURL(blob);

            img.onload = async function() {
                URL.revokeObjectURL(url);
                var w = img.width, h = img.height;

                // 如果设置了最大宽度且当前宽度超过，等比缩小
                if (maxWidth > 0 && w > maxWidth) {
                    h = Math.round(h * maxWidth / w);
                    w = maxWidth;
                }

                var canvas = document.createElement('canvas');
                canvas.width = w;
                canvas.height = h;
                var ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, w, h);

                var mimeType = 'image/' + (targetFormat === 'jpeg' ? 'jpeg' : targetFormat === 'same' ? (file.extension === 'jpg' ? 'jpeg' : file.extension) : targetFormat);
                var ext = targetFormat === 'jpeg' ? 'jpg' : targetFormat === 'same' ? file.extension : targetFormat;

                canvas.toBlob(async function(outBlob) {
                    var outBuf = await outBlob.arrayBuffer();

                    if (overwrite) {
                        // ★ 直接覆盖原文件
                        await plugin.app.vault.adapter.writeBinary(
                            plugin.app.vault.adapter.getFullPath(file.path),
                            outBuf
                        );
                        new Notice('✅ 已修改: ' + file.name + ' (' + w + 'x' + h + ', ' + ext.toUpperCase() + ', q=' + quality + ')');
                    } else {
                        // 另存为新文件
                        var baseName = file.basename;
                        var newPath = file.parent ? file.parent.path + '/' + baseName + '.' + ext : baseName + '.' + ext;
                        var counter = 1;
                        while (plugin.app.vault.getAbstractFileByPath(newPath)) {
                            newPath = file.parent ? file.parent.path + '/' + baseName + '_' + counter + '.' + ext : baseName + '_' + counter + '.' + ext;
                            counter++;
                        }
                        await plugin.app.vault.createBinary(newPath, outBuf);
                        new Notice('✅ 已另存: ' + newPath.split('/').pop() + ' (' + w + 'x' + h + ', ' + ext.toUpperCase() + ')');
                    }
                }, mimeType, quality / 100);
            };
            img.onerror = function() {
                URL.revokeObjectURL(url);
                new Notice('❌ 图片加载失败: ' + file.name);
            };
            img.src = url;
        } catch(e) {
            new Notice('❌ 处理失败: ' + e.message);
            console.error('[Dashboard] 图片处理失败:', e);
        }
    }

    // ★ 快捷转换（格式转换，覆盖原文件或另存）
    async _convertImageLocal(file, targetFormat, quality, maxWidth) {
        // 默认覆盖原文件
        await this._processImageLocal(file, targetFormat, quality, maxWidth || 0, true);
    }

    // ★ 图片修改（旋转/翻转，直接覆盖本地文件）
    async _modifyImageLocal(file, operation, params) {
        var plugin = this;
        try {
            new Notice('⏳ 正在处理图片...');
            var arrayBuffer = await plugin.app.vault.readBinary(file);
            var blob = new Blob([arrayBuffer]);
            var img = new Image();
            var url = URL.createObjectURL(blob);

            img.onload = async function() {
                URL.revokeObjectURL(url);
                var canvas = document.createElement('canvas');
                var ctx = canvas.getContext('2d');

                if (operation === 'rotate') {
                    var rad = params.degrees * Math.PI / 180;
                    var sin = Math.abs(Math.sin(rad)), cos = Math.abs(Math.cos(rad));
                    canvas.width = Math.ceil(img.width * cos + img.height * sin);
                    canvas.height = Math.ceil(img.width * sin + img.height * cos);
                    ctx.translate(canvas.width / 2, canvas.height / 2);
                    ctx.rotate(rad);
                    ctx.drawImage(img, -img.width / 2, -img.height / 2);
                } else if (operation === 'flip') {
                    canvas.width = img.width;
                    canvas.height = img.height;
                    if (params.direction === 'horizontal') {
                        ctx.translate(img.width, 0);
                        ctx.scale(-1, 1);
                    } else {
                        ctx.translate(0, img.height);
                        ctx.scale(1, -1);
                    }
                    ctx.drawImage(img, 0, 0);
                }

                var ext = file.extension.toLowerCase();
                var mimeType = ext === 'png' ? 'image/png' : ext === 'webp' ? 'image/webp' : 'image/jpeg';

                canvas.toBlob(async function(outBlob) {
                    var outBuf = await outBlob.arrayBuffer();
                    // ★ 直接覆盖原文件
                    await plugin.app.vault.adapter.writeBinary(
                        plugin.app.vault.adapter.getFullPath(file.path),
                        outBuf
                    );
                    var opLabel = operation === 'rotate' ? ('旋转' + params.degrees + '°') : (params.direction === 'horizontal' ? '水平翻转' : '垂直翻转');
                    new Notice('✅ 已修改: ' + file.name + ' (' + opLabel + ')');
                }, mimeType, 0.92);
            };
            img.onerror = function() {
                URL.revokeObjectURL(url);
                new Notice('❌ 图片加载失败: ' + file.name);
            };
            img.src = url;
        } catch(e) {
            new Notice('❌ 处理失败: ' + e.message);
            console.error('[Dashboard] 图片处理失败:', e);
        }
    }

    // ============ 后台功能：系统托盘 ============
    // ★ 参照原插件 tray(退出改为后台托盘)，完整实现：
    //    - 关闭窗口时隐藏到托盘（后台运行）
    //    - 托盘图标（使用 Obsidian 原版 16x16 图标）
    //    - 托盘 tooltip（支持 {{vault}} 占位符）
    //    - 右键菜单：显示/隐藏/退出
    //    - 全局快捷键切换窗口焦点
    //    - 启动时隐藏/任务栏图标隐藏
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

        // ★ 游戏娱乐模块强制默认关闭（覆盖已保存数据）
        var GAME_MODULE_IDS = ['aquarium', 'pixel-garden', 'particle-toy', 'farm-clicker'];
        GAME_MODULE_IDS.forEach(function(gmid) {
            if (this.settings.modules[gmid]) {
                this.settings.modules[gmid].enabled = false;
            }
        }.bind(this));

        // ★ 语言设置初始化
        if (!this.settings.language) this.settings.language = 'system';
        setLanguage(this.settings.language);

        if (!this.settings.layout) this.settings.layout = {};
        if (!this.settings.categoryCollapsed) this.settings.categoryCollapsed = {};
        if (!this.settings.sectionCollapsed) this.settings.sectionCollapsed = {};
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
