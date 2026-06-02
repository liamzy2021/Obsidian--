/**
 * AI洞察模块 V15
 * 格式：V14（含 id/styles/renderSettings/defaultSettings）
 * 功能：V11 完整版（分析最近5篇笔记 + 调用 AI API + 格式化显示 + 当天缓存）
 * 新增：全局请求节流 + 实例独立缓存 + 可配置请求延迟
 */
const id = 'ai-insight';
const title = 'AI洞察';
const icon = '💡';

const defaultSettings = {
    apiKey: '',
    apiUrl: 'https://api.openai.com/v1/chat/completions',
    model: 'gpt-3.5-turbo',
    temperature: 0.7,
    requestDelay: 0
};

const styles = `/* AI洞察模块样式已在 styles.css 中定义 */`;

// 全局 AI 请求节流器（跨实例共享，避免同时触发多个 AI 请求）
if (!window._v15AIThrottle) {
    window._v15AIThrottle = {
        lastRequestTime: 0,
        minInterval: 2000, // 默认最小间隔 2 秒
        async waitForTurn(extraDelayMs = 0) {
            const now = Date.now();
            const nextAvailable = this.lastRequestTime + this.minInterval;
            const waitTime = Math.max(0, nextAvailable - now) + extraDelayMs;
            if (waitTime > 0) {
                await new Promise(r => setTimeout(r, waitTime));
            }
            this.lastRequestTime = Date.now();
        }
    };
}

// 实例级缓存（以 settings 对象为 key，确保每个实例独立缓存）
if (!window._v15AICaches) {
    window._v15AICaches = new Map();
}

function getInstanceCache() {
    let state = window._v15AICaches.get(settings);
    if (!state) {
        state = { lastDate: null, analysisResult: null };
        window._v15AICaches.set(settings, state);
    }
    return state;
}

async function getRecentNotes(limit = 5) {
    const files = app.vault.getMarkdownFiles()
        .sort((a, b) => b.stat.mtime - a.stat.mtime)
        .slice(0, limit);

    const notes = [];
    for (const file of files) {
        try {
            const content = await app.vault.read(file);
            const cleanContent = content
                .replace(/^---[\s\S]*?---\n?/, '')
                .replace(/```[\s\S]*?```/g, '')
                .trim();
            notes.push({
                title: file.basename,
                content: cleanContent.substring(0, 600),
                path: file.path
            });
        } catch (e) { /* ignore */ }
    }
    return notes;
}

async function analyzeWithAI(notes) {
    const apiKey = settings.apiKey || '';
    const apiModel = settings.model || 'gpt-3.5-turbo';
    const temperature = settings.temperature || 0.7;

    let apiUrl = settings.apiUrl || 'https://api.openai.com/v1/chat/completions';
    if (apiUrl && !apiUrl.includes('/v1/') && !apiUrl.includes('/chat')) {
        apiUrl = apiUrl.replace(/\/$/, '') + '/v1/chat/completions';
    }

    if (!apiKey) throw new Error('请先在模块设置中配置 AI API 密钥');

    const prompt = `请分析以下笔记内容，提供：
1. 主题总结（2-3句话）
2. 关键知识点提取（3-5个）
3. 建议的关联方向或行动

笔记内容：
${notes.map((n, i) => `${i + 1}. 《${n.title}》\n${n.content}`).join('\n\n')}`;

    try {
        const response = await requestUrl({
            url: apiUrl,
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + apiKey
            },
            body: JSON.stringify({
                model: apiModel,
                messages: [{ role: 'user', content: prompt }],
                temperature: parseFloat(temperature)
            })
        });

        let data = response;
        if (response.text) {
            try { data = JSON.parse(response.text); } catch (e) { return response.text; }
        }
        if (typeof data === 'object' && data.json) data = data.json;

        if (data.choices?.[0]?.message?.content) return data.choices[0].message.content;
        if (data.content) return data.content;
        if (data.response) return data.response;
        if (data.text) return data.text;
        if (data.result) return data.result;
        if (data.error) throw new Error(data.error.message || 'API返回错误');

        throw new Error('无法解析 AI 响应格式');
    } catch (e) {
        if (e.message.includes('401')) throw new Error('API 密钥无效，请检查设置');
        if (e.message.includes('404')) throw new Error('API 地址无效，请检查 URL');
        if (e.message.includes('429')) throw new Error('请求频率过高，请稍后再试');
        throw new Error('AI 调用失败: ' + e.message);
    }
}

function displayContent(resultArea, text) {
    resultArea.empty();
    const lines = text.split('\n').filter(l => l.trim());
    lines.forEach(line => {
        if (line.startsWith('###')) {
            resultArea.createEl('h4', {
                text: line.replace(/^###\s*/, ''),
                attr: { style: 'margin: 10px 0 5px; font-size: 13px; color: var(--v6-primary);' }
            });
        } else if (line.startsWith('##')) {
            resultArea.createEl('h3', {
                text: line.replace(/^##\s*/, ''),
                attr: { style: 'margin: 12px 0 6px; font-size: 14px; color: var(--v6-primary);' }
            });
        } else if (line.startsWith('- ') || line.startsWith('* ')) {
            resultArea.createEl('div', {
                text: '• ' + line.substring(2),
                attr: { style: 'margin: 5px 0; padding-left: 10px; font-size: 13px;' }
            });
        } else if (/^\d+\./.test(line)) {
            resultArea.createEl('div', {
                text: line,
                attr: { style: 'margin: 5px 0; padding-left: 6px; font-size: 13px;' }
            });
        } else {
            resultArea.createEl('p', {
                text: line,
                attr: { style: 'margin: 6px 0; font-size: 13px; line-height: 1.7;' }
            });
        }
    });
}

async function render(content) {
    const state = getInstanceCache();
    const today = moment().format('YYYY-MM-DD');

    content.empty();
    const container = content.createDiv({ cls: 'ai-insight-container' });

    // 工具栏
    const toolbar = container.createDiv({ cls: 'ai-insight-toolbar' });
    const analyzeBtn = toolbar.createEl('button', { text: '🔍 分析最近笔记', cls: 'ai-insight-btn' });
    const clearBtn = toolbar.createEl('button', { text: '🗑️ 清除缓存', cls: 'ai-insight-btn secondary' });

    // 结果区域
    const resultArea = container.createDiv({ cls: 'ai-insight-response' });

    // 时间戳
    const dateEl = container.createDiv({ cls: 'ai-insight-date' });
    if (state.lastDate) dateEl.textContent = `上次分析：${state.lastDate}`;

    const doAnalyze = async () => {
        resultArea.empty();
        resultArea.createEl('div', {
            cls: 'ai-insight-loading',
            text: '🤔 正在分析笔记内容，请稍候...'
        });
        analyzeBtn.disabled = true;

        try {
            // 请求节流：等待轮到自己的回合
            const extraDelay = (Number(settings.requestDelay) || 0) * 1000;
            await window._v15AIThrottle.waitForTurn(extraDelay);

            const notes = await getRecentNotes(5);
            if (notes.length === 0) {
                resultArea.empty();
                resultArea.createEl('div', { cls: 'ai-insight-empty', text: '暂无笔记可分析' });
                analyzeBtn.disabled = false;
                return;
            }

            const result = await analyzeWithAI(notes);
            state.analysisResult = result;
            state.lastDate = today;
            dateEl.textContent = `分析于：${today}`;
            displayContent(resultArea, result);
        } catch (e) {
            resultArea.empty();
            resultArea.createEl('div', {
                cls: 'ai-insight-error',
                text: e.message
            });
        } finally {
            analyzeBtn.disabled = false;
        }
    };

    analyzeBtn.addEventListener('click', doAnalyze);
    clearBtn.addEventListener('click', () => {
        state.analysisResult = null;
        state.lastDate = null;
        resultArea.empty();
        resultArea.createEl('div', { cls: 'ai-insight-empty', text: '缓存已清除，点击「分析最近笔记」重新分析' });
        dateEl.textContent = '';
    });

    // 有缓存直接显示，无缓存自动触发分析
    if (state.lastDate === today && state.analysisResult) {
        displayContent(resultArea, state.analysisResult);
        dateEl.textContent = `分析于：${today}`;
    } else if (settings.apiKey) {
        doAnalyze();
    } else {
        resultArea.createEl('div', {
            cls: 'ai-insight-empty',
            text: '⚙️ 请先在模块设置中填写 AI API 密钥，再点击「分析最近笔记」'
        });
    }
}

function renderSettings(containerEl, plugin, saveCallback) {
    const { Setting } = require('obsidian');

    containerEl.createEl('h3', { text: 'AI洞察模块设置' });

    new Setting(containerEl)
        .setName('API Key')
        .setDesc('OpenAI 或兼容接口的 API 密钥（明文显示）')
        .addText(t => {
            t.setPlaceholder('sk-...')
                .setValue(settings.apiKey || '')
                .onChange(async (v) => {
                    settings.apiKey = v.trim();
                    await saveCallback();
                });
            t.inputEl.style.width = '100%';
        });

    new Setting(containerEl)
        .setName('API URL')
        .setDesc('留空使用 OpenAI 默认地址；使用其他兼容接口（如 deepseek、moonshot）请填入对应地址')
        .addText(t => {
            t.setPlaceholder('https://api.openai.com/v1/chat/completions')
                .setValue(settings.apiUrl || '')
                .onChange(async (v) => {
                    settings.apiUrl = v.trim();
                    await saveCallback();
                });
            t.inputEl.style.width = '100%';
        });

    new Setting(containerEl)
        .setName('模型')
        .setDesc('选择或输入模型名称')
        .addDropdown(d => {
            d.addOption('gpt-3.5-turbo', 'GPT-3.5 Turbo')
                .addOption('gpt-4o-mini', 'GPT-4o Mini')
                .addOption('gpt-4o', 'GPT-4o')
                .addOption('deepseek-chat', 'DeepSeek Chat')
                .addOption('moonshot-v1-8k', 'Moonshot v1-8k')
                .addOption('custom', '自定义...');

            const knownModels = ['gpt-3.5-turbo', 'gpt-4o-mini', 'gpt-4o', 'deepseek-chat', 'moonshot-v1-8k'];
            const currentModel = settings.model || 'gpt-3.5-turbo';
            d.setValue(knownModels.includes(currentModel) ? currentModel : 'custom')
                .onChange(async (v) => {
                    if (v !== 'custom') {
                        settings.model = v;
                        await saveCallback();
                    }
                });
        })
        .addText(t => {
            t.setPlaceholder('自定义模型名')
                .setValue(['gpt-3.5-turbo', 'gpt-4o-mini', 'gpt-4o', 'deepseek-chat', 'moonshot-v1-8k'].includes(settings.model || 'gpt-3.5-turbo') ? '' : (settings.model || ''))
                .onChange(async (v) => {
                    if (v.trim()) {
                        settings.model = v.trim();
                        await saveCallback();
                    }
                });
        });

    new Setting(containerEl)
        .setName('温度')
        .setDesc('越低越保守（0.0），越高越有创意（1.0）')
        .addSlider(s => {
            s.setLimits(0, 1, 0.1)
                .setValue(settings.temperature || 0.7)
                .setDynamicTooltip()
                .onChange(async (v) => {
                    settings.temperature = v;
                    await saveCallback();
                });
        });

    new Setting(containerEl)
        .setName('请求延迟')
        .setDesc('在此实例触发 AI 请求前的额外等待时间（秒），用于错开多个 AI 板块的并发请求')
        .addSlider(s => {
            s.setLimits(0, 10, 0.5)
                .setValue(Number(settings.requestDelay) || 0)
                .setDynamicTooltip()
                .onChange(async (v) => {
                    settings.requestDelay = v;
                    await saveCallback();
                });
        });

    new Setting(containerEl)
        .setName('全局最小间隔')
        .setDesc('所有 AI 洞察实例之间的最小请求间隔（毫秒），防止触发 API 频率限制')
        .addText(t => {
            t.setPlaceholder('2000')
                .setValue(String(window._v15AIThrottle ? window._v15AIThrottle.minInterval : 2000))
                .onChange(async (v) => {
                    const val = parseInt(v);
                    if (window._v15AIThrottle && isFinite(val) && val >= 0) {
                        window._v15AIThrottle.minInterval = val;
                    }
                    await saveCallback();
                });
        });
}

module.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };
