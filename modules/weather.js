/**
 * 天气模块 V15 - 高德地图API (居中排版)
 * 格式：V14（含 id/styles/renderSettings）
 * 功能：地理编码 + 实时天气 + 3天预报
 */
const id = 'weather';
const title = '天气';
const icon = '🌤️';

const defaultSettings = {
    city: '北京',
    apiKey: ''
};

const styles = `
.weather-wrap {
    padding: 0;
    height: 100%;
    display: flex;
    flex-direction: column;
}
/* 顶部实况区 - 居中 */
.weather-live {
    padding: 16px 14px 12px;
    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
    gap: 4px;
}
.weather-emoji {
    font-size: 52px;
    line-height: 1;
}
.weather-city-line {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-top: 4px;
}
.weather-city {
    font-size: 15px;
    font-weight: 600;
    color: var(--v6-text);
}
.weather-update-time {
    font-size: 10px;
    color: var(--v6-muted);
    background: var(--background-modifier-form-field);
    padding: 1px 6px;
    border-radius: 10px;
}
.weather-temp-main {
    font-size: 36px;
    font-weight: 700;
    color: var(--v6-primary);
    line-height: 1.1;
}
.weather-temp-main .unit {
    font-size: 20px;
    font-weight: 400;
    margin-left: 1px;
}
.weather-weather-text {
    font-size: 13px;
    color: var(--text-muted);
}
/* 实况详情网格 */
.weather-detail-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 6px;
    padding: 0 14px 10px;
}
.weather-detail-cell {
    background: var(--background-modifier-form-field);
    border-radius: 8px;
    padding: 8px 6px;
    text-align: center;
}
.weather-detail-cell .label {
    font-size: 10px;
    color: var(--text-muted);
    margin-bottom: 2px;
}
.weather-detail-cell .value {
    font-size: 13px;
    font-weight: 600;
    color: var(--text-normal);
}
/* 预报区 */
.weather-forecast-wrap {
    flex: 1;
    overflow: auto;
    padding: 0 14px 10px;
}
.weather-forecast-title {
    font-size: 11px;
    color: var(--text-muted);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 6px;
    padding-left: 2px;
}
.weather-forecast-list {
    display: flex;
    flex-direction: column;
    gap: 6px;
}
.weather-forecast-card {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 8px 10px;
    background: var(--background-modifier-form-field);
    border-radius: 8px;
}
.weather-forecast-card .day-label {
    width: 32px;
    font-size: 11px;
    font-weight: 600;
    color: var(--text-muted);
    text-align: center;
}
.weather-forecast-card .f-emoji {
    font-size: 22px;
    flex-shrink: 0;
}
.weather-forecast-card .f-desc {
    flex: 1;
    font-size: 12px;
    color: var(--text-normal);
}
.weather-forecast-card .f-temp {
    font-size: 12px;
    font-weight: 600;
    color: var(--v6-primary);
    text-align: right;
    white-space: nowrap;
}
.weather-forecast-card .f-temp .night {
    font-size: 10px;
    color: var(--text-muted);
    font-weight: 400;
}
/* 错误/空状态 */
.weather-empty {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100%;
    text-align: center;
    padding: 24px;
    color: var(--text-muted);
    gap: 8px;
}
.weather-empty .big-icon {
    font-size: 40px;
    opacity: 0.6;
}
.weather-empty .tip {
    font-size: 12px;
    line-height: 1.5;
}
.weather-empty .link {
    font-size: 11px;
    color: var(--v6-primary);
    cursor: pointer;
}
.weather-error {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100%;
    text-align: center;
    padding: 24px;
    color: var(--text-error);
    gap: 6px;
}
.weather-error .big-icon {
    font-size: 32px;
}
.weather-error .msg {
    font-size: 12px;
    line-height: 1.5;
}
.weather-error .retry {
    font-size: 11px;
    color: var(--v6-primary);
    cursor: pointer;
    margin-top: 4px;
}
`;

const iconMap = {
    '晴': '☀️', '少云': '🌤️', '多云': '⛅', '阴': '☁️',
    '阵雨': '🌦️', '小雨': '🌧️', '中雨': '🌧️', '大雨': '⛈️',
    '暴雨': '⛈️', '雷阵雨': '⛈️', '小雪': '🌨️', '中雪': '❄️',
    '大雪': '❄️', '雾': '🌫️', '霾': '🌫️', '风': '💨',
    '沙尘': '💨'
};

function getWeatherIcon(w) {
    if (!w) return '🌤️';
    for (const [key, val] of Object.entries(iconMap)) {
        if (w.includes(key)) return val;
    }
    return '🌤️';
}

async function fetchGeo(city, apiKey) {
    const url = 'https://restapi.amap.com/v3/geocode/geo?address=' + encodeURIComponent(city) + '&key=' + apiKey;
    const res = await requestUrl({ url, method: 'GET' });
    const data = res.json;
    if (!data || data.status !== '1' || !data.geocodes || data.geocodes.length === 0) {
        throw new Error('城市未找到，请检查城市名称或 API Key');
    }
    return data.geocodes[0];
}

async function fetchWeather(adcode, apiKey) {
    const [liveRes, forecastRes] = await Promise.all([
        requestUrl({ url: 'https://restapi.amap.com/v3/weather/weatherInfo?city=' + adcode + '&key=' + apiKey + '&extensions=base', method: 'GET' }),
        requestUrl({ url: 'https://restapi.amap.com/v3/weather/weatherInfo?city=' + adcode + '&key=' + apiKey + '&extensions=all', method: 'GET' })
    ]);

    const liveData = liveRes.json;
    const forecastData = forecastRes.json;

    if (!liveData || liveData.status !== '1' || !liveData.lives || liveData.lives.length === 0) {
        throw new Error('实时天气获取失败: ' + (liveData && liveData.info ? liveData.info : '未知'));
    }

    return {
        live: liveData.lives[0],
        forecast: forecastData && forecastData.status === '1' && forecastData.forecasts ? forecastData.forecasts[0] : null
    };
}

async function render(content) {
    content.empty();
    const wrap = content.createDiv({ cls: 'weather-wrap' });

    const apiKey = settings.apiKey || '';
    const city = settings.city || '北京';

    if (!apiKey) {
        const empty = wrap.createDiv({ cls: 'weather-empty' });
        empty.createEl('div', { text: '🔑', cls: 'big-icon' });
        empty.createEl('div', { text: '请先在模块设置中填写高德地图 API Key', cls: 'tip' });
        const link = empty.createEl('div', { text: '👉 免费申请', cls: 'link' });
        link.addEventListener('click', () => window.open('https://lbs.amap.com/', '_blank'));
        return;
    }

    try {
        const geo = await fetchGeo(city, apiKey);
        const adcode = geo.adcode;
        const cityName = geo.district || geo.city || geo.formatted_address || city;

        const { live, forecast } = await fetchWeather(adcode, apiKey);

        // ===== 实况区（居中）=====
        const liveSection = wrap.createDiv({ cls: 'weather-live' });
        liveSection.createEl('div', { text: getWeatherIcon(live.weather), cls: 'weather-emoji' });

        const cityLine = liveSection.createDiv({ cls: 'weather-city-line' });
        cityLine.createEl('span', { text: cityName, cls: 'weather-city' });
        cityLine.createEl('span', { text: live.reporttime ? live.reporttime.split(' ')[1] || live.reporttime : '', cls: 'weather-update-time' });

        liveSection.createEl('div', {
            cls: 'weather-temp-main',
            attr: { innerHTML: live.temperature + '<span class="unit">°C</span>' }
        });
        liveSection.createEl('div', { text: live.weather, cls: 'weather-weather-text' });

        // ===== 详情网格 =====
        const detailGrid = wrap.createDiv({ cls: 'weather-detail-grid' });
        const details = [
            { label: '湿度', value: (live.humidity || '--') + '%' },
            { label: '风向', value: (live.winddirection || '--') + '风' },
            { label: '风力', value: (live.windpower || '--') + '级' }
        ];
        details.forEach(d => {
            const cell = detailGrid.createDiv({ cls: 'weather-detail-cell' });
            cell.createEl('div', { text: d.label, cls: 'label' });
            cell.createEl('div', { text: d.value, cls: 'value' });
        });

        // ===== 预报区 =====
        if (forecast && forecast.casts && forecast.casts.length > 1) {
            const fWrap = wrap.createDiv({ cls: 'weather-forecast-wrap' });
            fWrap.createEl('div', { text: '未来预报', cls: 'weather-forecast-title' });

            const fList = fWrap.createDiv({ cls: 'weather-forecast-list' });
            forecast.casts.slice(1, 4).forEach((day, i) => {
                const card = fList.createDiv({ cls: 'weather-forecast-card' });
                const label = i === 0 ? '明天' : (i === 1 ? '后天' : (day.week ? '周' + ['日','一','二','三','四','五','六'][day.week] : ''));
                card.createEl('div', { text: label, cls: 'day-label' });
                card.createEl('div', { text: getWeatherIcon(day.dayweather), cls: 'f-emoji' });
                card.createEl('div', { text: day.dayweather + (day.nightweather && day.nightweather !== day.dayweather ? '转' + day.nightweather : ''), cls: 'f-desc' });
                card.createEl('div', {
                    cls: 'f-temp',
                    attr: { innerHTML: (day.daytemp || '--') + '°<span class="night"> / ' + (day.nighttemp || '--') + '°</span>' }
                });
            });
        }

    } catch (e) {
        wrap.empty();
        const err = wrap.createDiv({ cls: 'weather-error' });
        err.createEl('div', { text: '❌', cls: 'big-icon' });
        err.createEl('div', { text: e.message || '天气加载失败', cls: 'msg' });
        const retry = err.createEl('div', { text: '点击重试', cls: 'retry' });
        retry.addEventListener('click', () => render(content));
    }
}

function renderSettings(containerEl, plugin, saveCallback) {
    const { Setting } = require('obsidian');

    containerEl.createEl('h3', { text: '天气模块设置' });

    new Setting(containerEl)
        .setName('城市')
        .setDesc('输入城市名称（如：北京、上海、深圳）')
        .addText(t => {
            t.setPlaceholder('北京')
                .setValue(settings.city || '北京')
                .onChange(async (v) => {
                    settings.city = v.trim();
                    await saveCallback();
                });
        });

    new Setting(containerEl)
        .setName('高德地图 API Key')
        .setDesc('免费申请：https://lbs.amap.com/')
        .addText(t => {
            t.setPlaceholder('请输入 API Key')
                .setValue(settings.apiKey || '')
                .onChange(async (v) => {
                    settings.apiKey = v.trim();
                    await saveCallback();
                });
            t.inputEl.style.width = '100%';
        });
}

module.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };
