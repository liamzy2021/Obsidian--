/**
 * 网页预览模块 — 从 V13 原样迁移
 * viewport + wrapper + iframe 三层架构
 * iframe sandbox 始终全权限，无任何链接拦截
 */
const id = 'web-preview';
const title = '网页预览';
const icon = '🌐';

const defaultSettings = {
    url: 'https://www.baidu.com',
    zoom: 1,
    posX: 0,
    posY: 0
};

const styles = `
.web-preview-toolbar {
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
.web-preview-url {
    flex: 1;
    min-width: 80px;
    padding: 6px 8px;
    border: 1px solid var(--background-modifier-border);
    border-radius: 4px;
    background: var(--background-secondary);
    color: var(--text-normal);
    font-size: 12px;
}
.web-preview-url:focus {
    outline: none;
    border-color: var(--v6-primary);
}
.web-preview-btn {
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
.web-preview-btn:hover {
    background: var(--background-modifier-hover);
    border-color: var(--v6-primary);
}
.web-preview-zoom {
    font-size: 11px;
    color: var(--text-muted);
    min-width: 35px;
    text-align: center;
    flex-shrink: 0;
}
.web-preview-pos-input {
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
.web-preview-pos-input:focus {
    outline: none;
    border-color: var(--v6-primary);
}
.web-preview-viewport {
    flex: 1;
    overflow: hidden;
    position: relative;
    background: var(--background-secondary);
    min-height: 0;
}
.web-preview-wrapper {
    position: absolute;
    top: 0;
    left: 0;
    transform-origin: top left;
    transition: transform 0.3s ease;
}
.web-preview-iframe {
    width: 100%;
    height: 100%;
    border: none;
    background: white;
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
    const toolbar = container.createDiv({ cls: 'web-preview-toolbar' });

    const urlBar = toolbar.createEl('input', {
        cls: 'web-preview-url',
        attr: { type: 'text', value: settings.url, placeholder: '网址...' }
    });

    const zoomOutBtn = toolbar.createEl('button', {
        cls: 'web-preview-btn', text: '➖', attr: { title: '缩小' }
    });
    const zoomDisplay = toolbar.createEl('span', {
        cls: 'web-preview-zoom', text: Math.round(zoom * 100) + '%'
    });
    const zoomInBtn = toolbar.createEl('button', {
        cls: 'web-preview-btn', text: '➕', attr: { title: '放大' }
    });

    const posYInput = toolbar.createEl('input', {
        cls: 'web-preview-pos-input',
        attr: { type: 'number', value: settings.posY, title: '向下偏移' }
    });
    const posXInput = toolbar.createEl('input', {
        cls: 'web-preview-pos-input',
        attr: { type: 'number', value: settings.posX, title: '向右偏移' }
    });

    const refreshBtn = toolbar.createEl('button', {
        cls: 'web-preview-btn', text: '🔄', attr: { title: '刷新' }
    });

    // ── 视口（V13: position relative + overflow hidden） ──
    const viewport = container.createDiv({ cls: 'web-preview-viewport' });

    // ── iframe 包装器（V13: position absolute，用于 transform） ──
    const iframeWrapper = viewport.createDiv({ cls: 'web-preview-wrapper' });

    // ── iframe（V13 原始 sandbox） ──
    const iframe = iframeWrapper.createEl('iframe', {
        cls: 'web-preview-iframe',
        attr: {
            src: settings.url,
            sandbox: 'allow-scripts allow-same-origin allow-forms allow-popups'
        }
    });

    // ── 缩放和位置（V13 方案） ──
    const applyTransform = () => {
        const scale = zoom;
        const translateX = -settings.posX;
        const translateY = -settings.posY;
        iframeWrapper.style.transform = `translate(${translateX}px, ${translateY}px) scale(${scale})`;
        // iframe 宽高根据缩放调整，确保缩小后能看到更多内容
        const containerWidth = viewport.offsetWidth;
        iframeWrapper.style.width = `${(containerWidth * 2) / scale}px`;
        iframeWrapper.style.height = `${(containerWidth * 2) / scale}px`;
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
        iframe.src = urlBar.value;
    });

    urlBar.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            settings.url = urlBar.value;
            saveCallback();
            iframe.src = urlBar.value;
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
}

function renderSettings(wrapper, plugin, saveCallback) {
    new Setting(wrapper)
        .setName('预览网址')
        .setDesc('嵌入的网页地址')
        .addText(t => {
            t.setPlaceholder('https://example.com')
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
