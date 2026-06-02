/**
 * 统计模块 V15
 * 格式：V14（含 id/styles/renderSettings）
 * 功能：V11 完整版（笔记数/总字数/文件夹数/平均字数 + 文件夹排行Top5 带进度条）
 */
const id = 'stats';
const title = '笔记统计';
const icon = '📈';

const defaultSettings = {
    showFileCount: true,
    showWordCount: true
};

const styles = `/* 统计模块样式已在 styles.css 中定义 */`;

function formatNumber(num) {
    if (num >= 10000) return (num / 10000).toFixed(1) + '万';
    return num.toLocaleString();
}

async function render(content) {
    content.empty();

    const container = content.createDiv({ cls: 'stats-container' });

    // 加载提示
    const loading = container.createEl('div', {
        text: '⏳ 统计中...',
        attr: { style: 'grid-column: 1/-1; text-align: center; padding: 20px; color: var(--text-muted); font-size: 13px;' }
    });

    try {
        const files = app.vault.getMarkdownFiles();

        let totalWords = 0;
        const folderCount = new Set();
        const folderFiles = {};

        for (const file of files) {
            try {
                const fileContent = await app.vault.read(file);
                // 移除 YAML frontmatter 和 Markdown 符号再统计字符数
                const clean = fileContent
                    .replace(/^---[\s\S]*?---\n?/, '')
                    .replace(/```[\s\S]*?```/g, '')
                    .replace(/`[^`]*`/g, '')
                    .replace(/[#*\[\]>!\-_~|]/g, '');
                totalWords += clean.replace(/\s+/g, '').length;
            } catch (e) { /* 忽略单文件读取失败 */ }

            const parts = file.path.split('/');
            if (parts.length > 1) {
                folderCount.add(parts[0]);
                folderFiles[parts[0]] = (folderFiles[parts[0]] || 0) + 1;
            }
        }

        const avgWords = files.length > 0 ? Math.round(totalWords / files.length) : 0;
        const topFolders = Object.entries(folderFiles)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 5);
        const maxCount = topFolders.length > 0 ? topFolders[0][1] : 1;

        // 清空加载提示
        container.empty();

        // 四个统计卡片
        const showFileCount = settings.showFileCount !== false;
        const showWordCount = settings.showWordCount !== false;

        const items = [];
        if (showFileCount) {
            items.push({ icon: '📄', value: files.length, label: '笔记总数' });
        }
        if (showWordCount) {
            items.push({ icon: '✏️', value: totalWords, label: '总字数' });
        }
        items.push({ icon: '📁', value: folderCount.size, label: '文件夹' });
        if (showWordCount) {
            items.push({ icon: '📊', value: avgWords, label: '平均字数' });
        }

        items.forEach(item => {
            const itemEl = container.createDiv({ cls: 'stats-item' });
            itemEl.createEl('div', { text: item.icon, cls: 'stats-icon' });
            itemEl.createEl('div', { text: formatNumber(item.value), cls: 'stats-value' });
            itemEl.createEl('div', { text: item.label, cls: 'stats-label' });
        });

        // 文件夹排行（带进度条）
        if (topFolders.length > 0) {
            const rankDiv = container.createDiv({ cls: 'stats-rank' });
            rankDiv.createEl('div', { text: '📂 文件夹排行', cls: 'stats-rank-title' });

            topFolders.forEach((folder, index) => {
                const rankItem = rankDiv.createDiv({ cls: 'stats-rank-item' });
                rankItem.createEl('span', {
                    text: ['🥇','🥈','🥉','4️⃣','5️⃣'][index] || String(index + 1)
                });

                const info = rankItem.createDiv({ cls: 'stats-rank-info' });
                info.createEl('div', { text: folder[0], cls: 'stats-rank-name' });

                const barWrap = info.createDiv({ cls: 'stats-rank-bar-wrap' });
                const bar = barWrap.createDiv({ cls: 'stats-rank-bar' });
                const pct = Math.round((folder[1] / maxCount) * 100);
                bar.style.width = pct + '%';

                rankItem.createEl('span', { text: folder[1] + ' 篇', cls: 'stats-rank-count' });
            });
        }

    } catch (e) {
        container.empty();
        container.createEl('div', {
            text: '加载失败: ' + e.message,
            attr: { style: 'grid-column: 1/-1; text-align: center; padding: 20px; color: var(--text-muted); font-size: 12px;' }
        });
    }
}

function renderSettings(containerEl, plugin, saveCallback) {
    const { Setting } = require('obsidian');

    containerEl.createEl('h3', { text: '统计模块设置' });

    new Setting(containerEl)
        .setName('显示笔记数量')
        .addToggle(t => {
            t.setValue(settings.showFileCount !== false)
                .onChange(async (v) => {
                    settings.showFileCount = v;
                    await saveCallback();
                });
        });

    new Setting(containerEl)
        .setName('显示字数统计')
        .addToggle(t => {
            t.setValue(settings.showWordCount !== false)
                .onChange(async (v) => {
                    settings.showWordCount = v;
                    await saveCallback();
                });
        });
}

module.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };
