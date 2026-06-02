/**
 * 目录模块 V15
 * 格式：V14（含 id/styles/renderSettings）
 * 功能：树形目录，折叠/展开，文件图标，点击打开
 * 特性：展开状态持久化到 settings.expandedNodes（使用 child.path 作为 key）
 */
const id = 'directory';
const title = '目录';
const icon = '📂';

const defaultSettings = {
    folders: [],
    expandedNodes: []
};

const styles = `
.dir-tree { padding: 4px 0; }
.dir-root { margin-bottom: 8px; }
.dir-root-node {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 6px 8px;
    border-radius: 6px;
    cursor: default;
    font-weight: 600;
    font-size: 13px;
    color: var(--text-normal);
    background: var(--background-modifier-form-field);
}
.dir-root-label { flex: 1; }
.dir-count {
    font-size: 10px;
    color: var(--text-muted);
    background: var(--background-secondary);
    padding: 1px 6px;
    border-radius: 10px;
}
.dir-node { }
.dir-node-header {
    display: flex;
    align-items: center;
    gap: 4px;
    padding: 3px 6px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 12px;
    color: var(--text-normal);
    transition: background 0.15s;
}
.dir-node-header:hover {
    background: var(--background-modifier-hover);
}
.dir-toggle {
    width: 14px;
    text-align: center;
    font-size: 9px;
    color: var(--text-muted);
    cursor: pointer;
    flex-shrink: 0;
}
.dir-icon { flex-shrink: 0; }
.dir-label { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.dir-children { padding-left: 14px; }
.dir-children.collapsed { display: none; }
.dir-empty {
    text-align: center;
    padding: 24px;
    color: var(--text-muted);
    font-size: 13px;
}
`;

const FILE_ICONS = {
    'md': '📝', 'markdown': '📝',
    'png': '🖼️', 'jpg': '🖼️', 'jpeg': '🖼️', 'gif': '🖼️', 'webp': '🖼️', 'svg': '🖼️',
    'pdf': '📄',
    'doc': '📘', 'docx': '📘',
    'xls': '📗', 'xlsx': '📗',
    'ppt': '📙', 'pptx': '📙',
    'mp3': '🎵', 'wav': '🎵', 'flac': '🎵',
    'mp4': '🎬', 'mov': '🎬', 'mkv': '🎬',
    'zip': '📦', 'rar': '📦', '7z': '📦',
    'txt': '📃', 'csv': '📊', 'json': '🔧', 'js': '🔧', 'ts': '🔧', 'py': '🐍'
};

function getFileIcon(file) {
    const ext = (file.extension || '').toLowerCase();
    return FILE_ICONS[ext] || '📄';
}

function countFiles(folder) {
    if (!folder.children) return 0;
    let count = 0;
    folder.children.forEach(child => {
        count += child.children ? countFiles(child) : 1;
    });
    return count;
}

function renderFolder(container, folder, saveCallback) {
    if (!folder.children) return;

    const sorted = [...folder.children].sort((a, b) => {
        if (a.children && !b.children) return -1;
        if (!a.children && b.children) return 1;
        return a.name.localeCompare(b.name, 'zh-CN');
    });

    sorted.forEach(child => {
        const node = container.createDiv({ cls: 'dir-node' });

        if (child.children !== undefined) {
            // === 子文件夹 ===
            // 用 child.path 作为持久化 key（Obsidian 提供的完整路径，绝对可靠）
            const nodePath = child.path;
            const isExpanded = settings.expandedNodes && settings.expandedNodes.includes(nodePath);

            const header = node.createDiv({ cls: 'dir-node-header' });
            const toggle = header.createEl('span', { text: isExpanded ? '▼' : '▶', cls: 'dir-toggle' });
            header.createEl('span', { text: '📁', cls: 'dir-icon' });
            header.createEl('span', { text: child.name, cls: 'dir-label' });
            const cnt = countFiles(child);
            if (cnt > 0) header.createEl('span', { text: String(cnt), cls: 'dir-count' });

            const childContainer = node.createDiv({ cls: 'dir-children' + (isExpanded ? '' : ' collapsed') });

            // 若已展开，递归渲染子内容
            if (isExpanded) {
                renderFolder(childContainer, child, saveCallback);
            }

            header.addEventListener('click', async () => {
                const nowCollapsed = !childContainer.hasClass('collapsed');
                childContainer.toggleClass('collapsed', nowCollapsed);
                toggle.textContent = nowCollapsed ? '▶' : '▼';

                // 持久化展开状态
                if (!settings.expandedNodes) settings.expandedNodes = [];
                if (nowCollapsed) {
                    settings.expandedNodes = settings.expandedNodes.filter(p => p !== nodePath);
                } else {
                    if (!settings.expandedNodes.includes(nodePath)) {
                        settings.expandedNodes.push(nodePath);
                    }
                    // 展开时若子内容为空则渲染
                    if (childContainer.childElementCount === 0) {
                        renderFolder(childContainer, child, saveCallback);
                    }
                }

                // 调试日志
                console.log('[directory] 展开状态变更:', nodePath, nowCollapsed ? '折叠' : '展开', 'expandedNodes:', settings.expandedNodes);

                try {
                    await saveCallback();
                    console.log('[directory] 保存成功');
                } catch (e) {
                    console.error('[directory] 保存失败:', e);
                }
            });
        } else {
            // === 文件 ===
            const header = node.createDiv({ cls: 'dir-node-header' });
            header.createEl('span', { cls: 'dir-toggle' }); // 占位
            header.createEl('span', { text: getFileIcon(child), cls: 'dir-icon' });
            header.createEl('span', { text: child.name, cls: 'dir-label' });

            header.addEventListener('click', () => {
                app.workspace.openLinkText(child.path, '', false);
            });
        }
    });
}

async function render(content) {
    content.empty();

    // 确保 expandedNodes 已初始化
    if (!settings.expandedNodes) {
        settings.expandedNodes = [];
        console.log('[directory] 初始化 expandedNodes 为空数组');
    }
    console.log('[directory] 当前 expandedNodes:', settings.expandedNodes);

    const container = content.createDiv({ cls: 'dir-tree' });
    const folders = settings.folders || [];

    if (folders.length === 0) {
        container.createEl('div', {
            cls: 'dir-empty',
            text: '📁 请在设置中添加文件夹路径'
        });
        return;
    }

    for (const folderPath of folders) {
        const folder = app.vault.getAbstractFileByPath(folderPath);
        if (!folder || folder.children === undefined) {
            const errNode = container.createDiv({ cls: 'dir-root' });
            const errHeader = errNode.createDiv({ cls: 'dir-root-node' });
            errHeader.createEl('span', { text: '⚠️' });
            errHeader.createEl('span', {
                text: `文件夹不存在: ${folderPath}`,
                cls: 'dir-root-label',
                attr: { style: 'color: var(--text-muted);' }
            });
            continue;
        }

        const rootNode = container.createDiv({ cls: 'dir-root' });
        const rootHeader = rootNode.createDiv({ cls: 'dir-root-node' });
        rootHeader.createEl('span', { text: '📁' });
        rootHeader.createEl('span', { text: folder.name || folderPath, cls: 'dir-root-label' });
        const totalFiles = countFiles(folder);
        rootHeader.createEl('span', { text: totalFiles + ' 个文件', cls: 'dir-count' });

        const childContainer = rootNode.createDiv({ cls: 'dir-children' });
        renderFolder(childContainer, folder, async () => {
            console.log('[directory] 调用 saveSettings...');
            await plugin.saveSettings();
        });
    }
}

function renderSettings(containerEl, plugin, saveCallback) {
    const { Setting } = require('obsidian');

    containerEl.createEl('h3', { text: '目录模块设置' });
    containerEl.createEl('p', {
        text: '添加 Vault 中的文件夹路径（相对路径，如：笔记/日记）',
        attr: { style: 'font-size: 12px; color: var(--text-muted); margin: 0 0 8px;' }
    });

    // 初始化
    if (!settings.folders) settings.folders = [];
    if (!settings.expandedNodes) settings.expandedNodes = [];

    // 添加文件夹
    const addSetting = new Setting(containerEl)
        .setName('添加文件夹')
        .setDesc('输入文件夹路径后点击添加');

    let tempPath = '';
    addSetting.addText(t => {
        t.setPlaceholder('例如：笔记/日记')
            .onChange(v => { tempPath = v; });
    });
    addSetting.addButton(b => {
        b.setButtonText('添加')
            .setCta()
            .onClick(async () => {
                const path = tempPath.trim();
                if (!path) return new Notice('路径不能为空');
                if (settings.folders.includes(path)) return new Notice('已存在');
                const folder = app.vault.getAbstractFileByPath(path);
                if (!folder) return new Notice(`文件夹不存在: ${path}`);
                settings.folders.push(path);
                await saveCallback();
                containerEl.querySelectorAll('.dir-path-setting').forEach(el => el.remove());
                renderFolderList();
            });
    });

    // 已有文件夹列表
    const renderFolderList = () => {
        if (!settings.folders || settings.folders.length === 0) return;
        settings.folders.forEach((path, index) => {
            const s = new Setting(containerEl)
                .setName('📁 ' + path)
                .addButton(b => {
                    b.setButtonText('移除').setWarning()
                        .onClick(async () => {
                            settings.folders.splice(index, 1);
                            // 清理该文件夹相关的展开记录
                            if (settings.expandedNodes) {
                                settings.expandedNodes = settings.expandedNodes.filter(p => !p.startsWith(path + '/'));
                            }
                            await saveCallback();
                            containerEl.querySelectorAll('.dir-path-setting').forEach(el => el.remove());
                            renderFolderList();
                        });
                });
            s.settingEl.addClass('dir-path-setting');
        });
    };
    renderFolderList();
}

module.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };
