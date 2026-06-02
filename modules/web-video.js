/**
 * 网页视频模块 — 从 V13 原样迁移
 * viewport + wrapper + webview 三层架构
 * webview 始终 allowpopups，new-window 直接在内部加载
 */
const id = 'web-video';
const title = '网页视频';
const icon = '📺';

const defaultSettings = {
    url: 'https://www.bilibili.com',
    zoom: 1,
    posX: 0,
    posY: 0
};

const styles = `
.web-video-toolbar {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 8px 12px;
    border-bottom: 1px solid var(--background-modifier-border);
    background: var(--background-secondary-alt);
    flex-wrap: nowrap;
    overflow: hidden;
    flex-shrink: 0;
}
.web-video-url {
    flex: 1;
    min-width: 80px;
    padding: 6px 8px;
    border: 1px solid var(--background-modifier-border);
    border-radius: 4px;
    background: var(--background-secondary);
    color: var(--text-normal);
    font-size: 12px;
}
.web-video-url:focus {
    outline: none;
    border-color: var(--v6-primary);
}
.web-video-btn {
    padding: 6px 8px;
    border: 1px solid var(--background-modifier-border);
    background: var(--background-secondary);
    border-radius: 4px;
    cursor: pointer;
    font-size: 12px;
    color: var(--text-normal);
    transition: all 0.2s ease;
    flex-shrink: 0;
}
.web-video-btn:hover {
    background: var(--background-modifier-hover);
    border-color: var(--v6-primary);
}
.web-video-zoom {
    font-size: 11px;
    color: var(--text-muted);
    min-width: 35px;
    text-align: center;
    flex-shrink: 0;
}
.web-video-pos-input {
    width: 45px;
    padding: 4px 6px;
    border: 1px solid var(--background-modifier-border);
    border-radius: 4px;
    background: var(--background-secondary);
    color: var(--text-normal);
    font-size: 11px;
    text-align: center;
    flex-shrink: 0;
}
.web-video-pos-input:focus {
    outline: none;
    border-color: var(--v6-primary);
}
.web-video-viewport {
    flex: 1;
    overflow: hidden;
    position: relative;
    background: var(--background-secondary);
    min-height: 0;
}
.web-video-wrapper {
    position: absolute;
    top: 0;
    left: 0;
    transform-origin: top left;
    transition: transform 0.3s ease;
}
.web-video-webview {
    width: 100%;
    height: 100%;
    border: none;
}
`;

async function render(container) {
    container.empty();

    // V13 原始初始化
    if (!settings.zoom) settings.zoom = 1;
    if (settings.posY === undefined || settings.posY === null) settings.posY = 0;
    if (settings.posX === undefined || settings.posX === null) settings.posX = 0;

    let zoom = Number(settings.zoom);
    if (!isFinite(zoom) || zoom <= 0) zoom = 1;

    container.style.display = 'flex';
    container.style.flexDirection = 'column';
    container.style.height = '100%';

    // ── 工具栏（V13 原始结构） ──
    const toolbar = container.createDiv({ cls: 'web-video-toolbar' });

    const urlBar = toolbar.createEl('input', {
        cls: 'web-video-url',
        attr: { type: 'text', value: settings.url, placeholder: '网址...' }
    });

    const zoomOutBtn = toolbar.createEl('button', {
        cls: 'web-video-btn', text: '➖', attr: { title: '缩小' }
    });
    const zoomDisplay = toolbar.createEl('span', {
        cls: 'web-video-zoom', text: Math.round(zoom * 100) + '%'
    });
    const zoomInBtn = toolbar.createEl('button', {
        cls: 'web-video-btn', text: '➕', attr: { title: '放大' }
    });

    const posYInput = toolbar.createEl('input', {
        cls: 'web-video-pos-input',
        attr: { type: 'number', value: settings.posY, title: '向下偏移' }
    });
    const posXInput = toolbar.createEl('input', {
        cls: 'web-video-pos-input',
        attr: { type: 'number', value: settings.posX, title: '向右偏移' }
    });

    const refreshBtn = toolbar.createEl('button', {
        cls: 'web-video-btn', text: '🔄', attr: { title: '刷新' }
    });

    // ── 视口（V13: position relative + overflow hidden） ──
    const viewport = container.createDiv({ cls: 'web-video-viewport' });

    // ── webview 包装器（V13: position absolute，用于 transform） ──
    const webviewWrapper = viewport.createDiv({ cls: 'web-video-wrapper' });

    // ── Electron webview（V13 原始属性） ──
    const webview = document.createElement('webview');
    webview.className = 'web-video-webview';
    webview.setAttribute('src', settings.url);
    webview.setAttribute('partition', 'persist:webvideo-' + (_moduleId || id));
    webview.setAttribute('preload', '');
    webview.setAttribute('allowpopups', '');

    webview.setAttribute('nodeintegration', 'false');
    webview.setAttribute('webpreferences', 'contextIsolation=true, sandbox=true');

    webviewWrapper.appendChild(webview);

    // ── 缩放和位置（V13 方案） ──
    const applyTransform = () => {
        const scale = zoom;
        const translateX = -settings.posX;
        const translateY = -settings.posY;
        webviewWrapper.style.transform = `translate(${translateX}px, ${translateY}px) scale(${scale})`;
        const containerWidth = viewport.offsetWidth;
        webviewWrapper.style.width = `${(containerWidth * 2) / scale}px`;
        webviewWrapper.style.height = `${(containerWidth * 2) / scale}px`;
    };

    applyTransform();

    // ── 缩放 ──
    const updateZoom = () => {
        zoom = Math.max(0.1, Math.min(2, zoom));
        settings.zoom = zoom;
        zoomDisplay.textContent = Math.round(zoom * 100) + '%';
        applyTransform();
        saveCallback();
    };

    zoomOutBtn.addEventListener('click', () => { zoom -= 0.1; updateZoom(); });
    zoomInBtn.addEventListener('click', () => { zoom += 0.1; updateZoom(); });

    // ── 刷新 ──
    refreshBtn.addEventListener('click', () => {
        settings.url = urlBar.value;
        saveCallback();
        webview.src = urlBar.value;
    });

    urlBar.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            settings.url = urlBar.value;
            saveCallback();
            webview.src = urlBar.value;
        }
    });

    // ── 位置更新 ──
    const updatePosition = () => {
        settings.posX = parseInt(posXInput.value) || 0;
        settings.posY = parseInt(posYInput.value) || 0;
        applyTransform();
        saveCallback();
    };

    posXInput.addEventListener('change', updatePosition);
    posYInput.addEventListener('change', updatePosition);

    // ── 注入 CSS 屏蔽广告（V13 原始逻辑） ──
    webview.addEventListener('dom-ready', () => {
        webview.insertCSS(`
            .ad, .ads, .advertisement, .popup, .modal-overlay { display: none !important; }
        `).catch(() => {});
    });

    // ── 新窗口在内部打开（V13 原始逻辑，直接 webview.src = url） ──
    webview.addEventListener('new-window', (e) => {
        webview.src = e.url;
    });
}

function renderSettings(wrapper, plugin, saveCallback) {
    new Setting(wrapper)
        .setName('视频网址')
        .setDesc('使用 Electron webview 打开')
        .addText(t => {
            t.setPlaceholder('https://www.bilibili.com')
                .setValue(settings.url || '')
                .onChange(async (v) => { settings.url = v; await saveCallback(); });
        });

    new Setting(wrapper)
        .setName('默认缩放')
        .setDesc('初始缩放比例（0.1 ~ 2.0）')
        .addSlider(s => {
            s.setLimits(0.1, 2, 0.1)
                .setValue(Number(settings.zoom) || 1)
                .setDynamicTooltip()
                .onChange(async (v) => { settings.zoom = v; await saveCallback(); });
        });

    new Setting(wrapper)
        .setName('水平偏移 (X)')
        .setDesc('向右偏移像素值')
        .addText(t => {
            t.setValue(String(settings.posX || 0))
                .onChange(async (v) => { settings.posX = parseInt(v) || 0; await saveCallback(); });
        });

    new Setting(wrapper)
        .setName('垂直偏移 (Y)')
        .setDesc('向下偏移像素值')
        .addText(t => {
            t.setValue(String(settings.posY || 0))
                .onChange(async (v) => { settings.posY = parseInt(v) || 0; await saveCallback(); });
        });
}

module.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };
