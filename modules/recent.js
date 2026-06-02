/**
 * 最近文件模块 V15
 * 格式：V14（含 id/styles/renderSettings）
 * 功能：V11/V14 一致（最近修改文件列表，相对时间，点击打开）
 */
const id = 'recent';
const title = '最近文件';
const icon = '🕐';

const defaultSettings = {
    maxFiles: 10
};

const styles = `/* 最近文件模块样式已在 styles.css 中定义 */`;

function formatTime(timestamp) {
    const diff = Date.now() - timestamp;
    const minutes = Math.floor(diff / 60000);
    const hours = Math.floor(diff / 3600000);
    const days = Math.floor(diff / 86400000);
    if (minutes < 1) return '刚刚';
    if (minutes < 60) return minutes + '分钟前';
    if (hours < 24) return hours + '小时前';
    if (days === 1) return '昨天';
    if (days < 7) return days + '天前';
    return moment(timestamp).format('MM-DD');
}

async function render(content) {
    content.empty();

    const container = content.createDiv({ cls: 'recent-container' });
    const maxFiles = settings.maxFiles || 10;

    try {
        const files = app.vault.getMarkdownFiles()
            .sort((a, b) => b.stat.mtime - a.stat.mtime)
            .slice(0, maxFiles);

        if (files.length === 0) {
            container.createEl('div', { text: '暂无文件', cls: 'recent-empty' });
            return;
        }

        files.forEach(file => {
            const item = container.createDiv({ cls: 'recent-item' });
            item.createEl('div', { text: '📝', cls: 'recent-icon' });

            const info = item.createEl('div', { cls: 'recent-info' });
            info.createEl('div', { text: file.basename, cls: 'recent-title' });

            const pathParts = file.path.split('/');
            pathParts.pop();
            const folderPath = pathParts.join('/') || '根目录';
            info.createEl('div', { text: folderPath, cls: 'recent-path' });

            item.createEl('div', { text: formatTime(file.stat.mtime), cls: 'recent-time' });

            item.addEventListener('click', () => {
                app.workspace.openLinkText(file.path, '', false);
            });
        });

    } catch (e) {
        container.createEl('div', {
            text: '加载失败: ' + e.message,
            attr: { style: 'padding: 20px; text-align: center; color: var(--text-muted); font-size: 12px;' }
        });
    }
}

function renderSettings(containerEl, plugin, saveCallback) {
    const { Setting } = require('obsidian');

    containerEl.createEl('h3', { text: '最近文件设置' });

    new Setting(containerEl)
        .setName('显示数量')
        .setDesc('最多显示多少个最近修改的文件')
        .addSlider(s => {
            s.setLimits(5, 30, 5)
                .setValue(settings.maxFiles || 10)
                .setDynamicTooltip()
                .onChange(async (v) => {
                    settings.maxFiles = v;
                    await saveCallback();
                });
        });
}

module.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };
