/**
 * 新闻模块 V15 - AI HOT RSS (全新UI)
 * 格式：V14（含 id/styles/renderSettings）
 */
const id = 'news';
const title = '资讯';
const icon = '🔥';

const defaultSettings = {
    source: 'aihot',
    pageSize: 10
};

const styles = `
/* Tab 栏 */
.aihot-tabs {
    display: flex;
    gap: 4px;
    padding: 10px 12px 6px;
    border-bottom: 1px solid var(--background-modifier-border);
}
.aihot-tab {
    flex: 1;
    padding: 5px 4px;
    border: none;
    background: transparent;
    border-radius: 6px;
    cursor: pointer;
    font-size: 12px;
    color: var(--text-muted);
    font-weight: 500;
    transition: all 0.2s ease;
    text-align: center;
}
.aihot-tab:hover {
    background: var(--background-modifier-hover);
    color: var(--text-normal);
}
.aihot-tab.active {
    background: var(--v6-primary);
    color: white;
}

/* 文章卡片 */
.aihot-card {
    padding: 12px;
    display: flex;
    flex-direction: column;
    height: calc(100% - 80px);
}
.aihot-source-badge {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    font-size: 10px;
    font-weight: 600;
    color: var(--v6-primary);
    background: var(--v6-primary);
    opacity: 0.15;
    padding: 2px 8px;
    border-radius: 10px;
    margin-bottom: 8px;
    width: fit-content;
}
.aihot-source-badge span {
    opacity: 6;
    color: var(--v6-primary);
}
.aihot-article-title {
    font-size: 15px;
    font-weight: 600;
    color: var(--text-normal);
    line-height: 1.45;
    margin-bottom: 8px;
    display: -webkit-box;
    -webkit-line-clamp: 3;
    -webkit-box-orient: vertical;
    overflow: hidden;
}
.aihot-article-meta {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 11px;
    color: var(--text-muted);
    margin-bottom: 10px;
}
.aihot-article-meta .dot {
    width: 3px;
    height: 3px;
    border-radius: 50%;
    background: var(--text-muted);
    opacity: 0.5;
}
.aihot-article-body {
    flex: 1;
    overflow: auto;
    background: var(--background-modifier-form-field);
    border-radius: 8px;
    padding: 10px 12px;
    margin-bottom: 10px;
}
.aihot-article-body p {
    font-size: 13px;
    color: var(--text-normal);
    line-height: 1.65;
    margin: 0;
    display: -webkit-box;
    -webkit-line-clamp: 8;
    -webkit-box-orient: vertical;
    overflow: hidden;
}

/* 操作区 */
.aihot-actions {
    display: flex;
    gap: 8px;
    margin-bottom: 10px;
}
.aihot-btn {
    flex: 1;
    padding: 8px;
    border: 1px solid var(--background-modifier-border);
    background: var(--background-secondary);
    border-radius: 6px;
    cursor: pointer;
    font-size: 12px;
    color: var(--text-normal);
    text-align: center;
    transition: all 0.15s;
}
.aihot-btn:hover {
    background: var(--background-modifier-hover);
}
.aihot-btn.primary {
    background: var(--v6-primary);
    border-color: var(--v6-primary);
    color: white;
}
.aihot-btn.primary:hover {
    opacity: 0.9;
}

/* 导航栏 */
.aihot-footer {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding-top: 8px;
    border-top: 1px solid var(--background-modifier-border);
}
.aihot-footer-btn {
    padding: 5px 10px;
    border: none;
    background: transparent;
    border-radius: 6px;
    cursor: pointer;
    font-size: 11px;
    color: var(--text-muted);
    transition: all 0.15s;
}
.aihot-footer-btn:hover:not(:disabled) {
    background: var(--background-modifier-hover);
    color: var(--text-normal);
}
.aihot-footer-btn:disabled {
    opacity: 0.3;
    cursor: not-allowed;
}
.aihot-footer-counter {
    font-size: 11px;
    color: var(--text-muted);
    font-weight: 500;
    font-variant-numeric: tabular-nums;
}

/* 状态 */
.v5-loading {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100%;
    gap: 8px;
    color: var(--text-muted);
    font-size: 13px;
}
.v5-error {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100%;
    text-align: center;
    padding: 24px 16px;
    gap: 8px;
    color: var(--text-error);
}
.v5-error .err-title {
    font-size: 13px;
    font-weight: 600;
}
.v5-error .err-detail {
    font-size: 11px;
    color: var(--text-muted);
    line-height: 1.5;
    max-width: 100%;
    word-break: break-all;
}
.v5-error .err-retry {
    margin-top: 4px;
    padding: 6px 16px;
    border: none;
    background: var(--v6-primary);
    border-radius: 6px;
    cursor: pointer;
    font-size: 12px;
    color: white;
}
.v5-warning {
    padding: 10px 12px;
    font-size: 11px;
    color: var(--v6-primary);
    background: var(--v6-primary);
    opacity: 0.1;
    border-radius: 6px;
    margin: 8px 12px;
}
.v5-warning span {
    opacity: 10;
    color: var(--v6-primary);
}
`;

const RSS_FEEDS = {
    '精选': 'https://aihot.virxact.com/feed.xml',
    '全部': 'https://aihot.virxact.com/feed/all.xml',
    '日报': 'https://aihot.virxact.com/feed/daily.xml'
};

if (!window._v15NewsState) {
    window._v15NewsState = {
        currentFeed: '精选',
        currentIndex: 0,
        cachedData: null,
        currentItems: null
    };
}

function parseRSS_DOM(text) {
    if (typeof DOMParser === 'undefined') throw new Error('DOMParser 不可用');
    const parser = new DOMParser();
    const xml = parser.parseFromString(text, 'application/xml');
    const parseError = xml.querySelector('parsererror');
    if (parseError) throw new Error('DOMParser 解析 XML 出错');

    const items = [];
    xml.querySelectorAll('item').forEach(item => {
        const getText = (sel) => {
            const el = item.querySelector(sel);
            return el ? el.textContent.trim() : '';
        };
        const description = getText('content\\:encoded') || getText('content:encoded') || getText('description');
        const author = getText('dc\\:creator') || getText('dc:creator') || getText('author');
        items.push({
            title: getText('title'),
            link: getText('link'),
            description: description,
            pubDate: getText('pubDate'),
            author: author
        });
    });
    if (items.length === 0) throw new Error('未找到 item 节点');
    return items;
}

function parseRSS_Regex(text) {
    const items = [];
    const itemMatches = text.match(/<item[\s\S]*?<\/item>/gi);
    if (!itemMatches || itemMatches.length === 0) throw new Error('正则未匹配到 item');

    itemMatches.forEach(itemBlock => {
        const getTag = (tag) => {
            const re = new RegExp('<' + tag + '(?:\\s[^>]*)?>([\\s\\S]*?)<\\/' + tag + '>', 'i');
            const m = itemBlock.match(re);
            return m ? m[1].replace(/<!\[CDATA\[|\]\]>/g, '').trim() : '';
        };
        items.push({
            title: getTag('title'),
            link: getTag('link'),
            description: getTag('content:encoded') || getTag('description'),
            pubDate: getTag('pubDate'),
            author: getTag('dc:creator') || getTag('author')
        });
    });
    return items;
}

function parseRSS(text) {
    try { return parseRSS_DOM(text); }
    catch (e) { return parseRSS_Regex(text); }
}

function isValidXML(text) {
    const t = text.trim();
    return t.startsWith('<?xml') || t.startsWith('<rss') || t.startsWith('<feed');
}

function formatTime(pubDate) {
    if (!pubDate) return '';
    try {
        const m = moment(pubDate);
        if (m.isValid()) return m.fromNow();
    } catch (e) {}
    return pubDate;
}

function stripHtml(html) {
    if (!html) return '';
    return html
        .replace(/<script[^>]*>.*?<\/script>/gi, '')
        .replace(/<style[^>]*>.*?<\/style>/gi, '')
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

async function render(content) {
    const state = window._v15NewsState;
    const feedUrl = RSS_FEEDS[state.currentFeed];

    content.empty();
    const loading = content.createDiv({ cls: 'v5-loading' });
    loading.createEl('div', { text: '🔥', attr: { style: 'font-size: 28px;' } });
    loading.createEl('div', { text: '加载 AI HOT...' });

    try {
        const res = await requestUrl({
            url: feedUrl,
            method: 'GET',
            headers: {
                'Accept': 'application/rss+xml, application/xml, text/xml, */*',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
            }
        });

        if (res.status !== 200) {
            throw new Error('HTTP ' + res.status + (res.text ? ': ' + res.text.substring(0, 80) : ''));
        }

        const rssText = res.text;
        if (!rssText) throw new Error('响应内容为空');

        if (!isValidXML(rssText)) {
            const preview = rssText.substring(0, 120).replace(/\s+/g, ' ');
            throw new Error('返回的不是 RSS/XML。\n前120字符: ' + preview);
        }

        const items = parseRSS(rssText);
        if (!items || items.length === 0) throw new Error('解析成功但无内容');

        state.cachedData = items;
        state.currentItems = items;
        state.currentIndex = 0;

        content.empty();
        renderUI(content, state);
        updateArticle(content, state);

    } catch (e) {
        content.empty();

        if (state.cachedData && state.cachedData.length > 0) {
            const warning = content.createDiv({ cls: 'v5-warning' });
            warning.createEl('span', { text: '⚠️ 网络异常，显示缓存内容' });
            state.currentItems = state.cachedData;
            state.currentIndex = 0;
            renderUI(content, state);
            updateArticle(content, state);
            return;
        }

        const err = content.createDiv({ cls: 'v5-error' });
        err.createEl('div', { text: '❌', attr: { style: 'font-size: 28px;' } });
        err.createEl('div', { text: '加载失败', cls: 'err-title' });
        err.createEl('div', { text: e.message || '未知错误', cls: 'err-detail' });
        const retry = err.createEl('button', { text: '重新加载', cls: 'err-retry' });
        retry.addEventListener('click', () => render(content));
    }
}

function renderUI(content, state) {
    // Tab 栏
    const tabs = content.createDiv({ cls: 'aihot-tabs' });
    Object.keys(RSS_FEEDS).forEach(feedName => {
        const btn = tabs.createEl('button', {
            text: feedName,
            cls: 'aihot-tab' + (state.currentFeed === feedName ? ' active' : '')
        });
        btn.addEventListener('click', () => {
            state.currentFeed = feedName;
            state.currentIndex = 0;
            state.cachedData = null;
            state.currentItems = null;
            render(content);
        });
    });

    // 文章卡片
    const card = content.createDiv({ cls: 'aihot-card' });

    const badge = card.createDiv({ cls: 'aihot-source-badge' });
    badge.createEl('span', { text: 'AI HOT' });

    card.createEl('h3', { cls: 'aihot-article-title', attr: { 'data-role': 'title' } });

    const meta = card.createDiv({ cls: 'aihot-article-meta' });
    meta.createEl('span', { attr: { 'data-role': 'author' } });
    meta.createEl('span', { cls: 'dot' });
    meta.createEl('span', { attr: { 'data-role': 'time' } });

    const body = card.createDiv({ cls: 'aihot-article-body' });
    body.createEl('p', { attr: { 'data-role': 'desc' } });

    // 操作按钮
    const actions = card.createDiv({ cls: 'aihot-actions' });
    const readBtn = actions.createEl('button', { text: '查看原文 →', cls: 'aihot-btn primary' });
    readBtn.addEventListener('click', () => {
        const item = state.currentItems[state.currentIndex];
        if (item && item.link) window.open(item.link, '_blank');
    });

    // 底部导航
    const footer = card.createDiv({ cls: 'aihot-footer' });
    const prevBtn = footer.createEl('button', { text: '← 上一条', cls: 'aihot-footer-btn', attr: { 'data-role': 'prev' } });
    prevBtn.addEventListener('click', () => {
        if (state.currentIndex > 0) {
            state.currentIndex--;
            updateArticle(content, state);
        }
    });

    footer.createEl('span', { cls: 'aihot-footer-counter', attr: { 'data-role': 'counter' } });

    const nextBtn = footer.createEl('button', { text: '下一条 →', cls: 'aihot-footer-btn', attr: { 'data-role': 'next' } });
    nextBtn.addEventListener('click', () => {
        if (state.currentIndex < state.currentItems.length - 1) {
            state.currentIndex++;
            updateArticle(content, state);
        }
    });
}

function updateArticle(content, state) {
    const items = state.currentItems;
    if (!items || items.length === 0) return;

    const item = items[state.currentIndex] || items[0];

    const titleEl = content.querySelector('[data-role="title"]');
    if (titleEl) titleEl.textContent = item.title || '无标题';

    const authorEl = content.querySelector('[data-role="author"]');
    if (authorEl) authorEl.textContent = item.author || 'AI HOT';

    const timeEl = content.querySelector('[data-role="time"]');
    if (timeEl) timeEl.textContent = formatTime(item.pubDate);

    const descEl = content.querySelector('[data-role="desc"]');
    if (descEl) {
        const text = stripHtml(item.description);
        descEl.textContent = text.substring(0, 400) + (text.length >= 400 ? '...' : '');
    }

    const prevBtn = content.querySelector('[data-role="prev"]');
    const nextBtn = content.querySelector('[data-role="next"]');
    const counterEl = content.querySelector('[data-role="counter"]');

    if (prevBtn) prevBtn.disabled = state.currentIndex === 0;
    if (nextBtn) nextBtn.disabled = state.currentIndex >= items.length - 1;
    if (counterEl) counterEl.textContent = (state.currentIndex + 1) + ' / ' + items.length;
}

function renderSettings(containerEl, plugin, saveCallback) {
    const { Setting } = require('obsidian');

    containerEl.createEl('h3', { text: '资讯模块设置' });

    new Setting(containerEl)
        .setName('默认订阅源')
        .setDesc('打开时默认显示的 RSS 源')
        .addDropdown(d => {
            Object.keys(RSS_FEEDS).forEach(name => d.addOption(name, name));
            d.setValue(settings.defaultFeed || '精选')
                .onChange(async (v) => {
                    settings.defaultFeed = v;
                    window._v15NewsState.currentFeed = v;
                    await saveCallback();
                });
        });
}

module.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };
