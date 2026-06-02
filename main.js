/**
 * AI Dashboard V17
 * 底层：V14 自由拖拽/resize 布局 + 完全模块化架构
 * 功能：V11 完整功能迁移 + 无限实例化系统（所有模块默认可克隆）
 * 主题：V11 8个精美主题
 */

const { Plugin, ItemView, Setting, PluginSettingTab, Modal, Notice, setIcon, requestUrl, moment } = require('obsidian');

const VIEW_TYPE = 'ai-dashboard-v15';

const BUILTIN_MODULES = {
    'weather': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
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
            
        }
    },
    'calendar': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
            /**
             * 日历模块 V15
             * 格式：V14（含 id/styles/renderSettings）
             * 功能：V11 完整版（月历翻页 + 农历 + 节日 + 节气 + 天干地支）
             */
            const id = 'calendar';
            const title = '日历';
            const icon = '📅';
            
            const defaultSettings = {
                showLunar: true,
                showHoliday: true
            };
            
            const styles = `/* 日历模块样式已在 styles.css 中定义 */`;
            
            // ===== 农历工具 =====
            const LUNAR_INFO = [
                0x04bd8,0x04ae0,0x0a570,0x054d5,0x0d260,0x0d950,0x16554,0x056a0,0x09ad0,0x055d2,
                0x04ae0,0x0a5b6,0x0a4d0,0x0d250,0x1d255,0x0b540,0x0d6a0,0x0ada2,0x095b0,0x14977,
                0x04970,0x0a4b0,0x0b4b5,0x06a50,0x06d40,0x1ab54,0x02b60,0x09570,0x052f2,0x04970,
                0x06566,0x0d4a0,0x0ea50,0x06e95,0x05ad0,0x02b60,0x186e3,0x092e0,0x1c8d7,0x0c950,
                0x0d4a0,0x1d8a6,0x0b550,0x056a0,0x1a5b4,0x025d0,0x092d0,0x0d2b2,0x0a950,0x0b557,
                0x06ca0,0x0b550,0x15355,0x04da0,0x0a5b0,0x14573,0x052b0,0x0a9a8,0x0e950,0x06aa0,
                0x0aea6,0x0ab50,0x04b60,0x0aae4,0x0a570,0x05260,0x0f263,0x0d950,0x05b57,0x056a0,
                0x096d0,0x04dd5,0x04ad0,0x0a4d0,0x0d4d4,0x0d250,0x0d558,0x0b540,0x0b6a0,0x195a6,
                0x095b0,0x049b0,0x0a974,0x0a4b0,0x0b27a,0x06a50,0x06d40,0x0af46,0x0ab60,0x09570,
                0x04af5,0x04970,0x064b0,0x074a3,0x0ea50,0x06aa0,0x0a6b6,0x056a0,0x02b40,0x0acb6,
                0x0a940,0x0a950,0x0b4a6,0x0b550,0x0d2a0,0x11d25,0x0d960,0x05954,0x056a0,0x0aba0,
                0x1a3c5,0x09250,0x0a950,0x0b965,0x0aa40,0x0bccd,0x0b550,0x04b60,0x0a576,0x0a520,
                0x0dd45,0x0d950,0x056a0,0x14ad5,0x055d0,0x0a9b0,0x14b75,0x04970,0x0a4b0,0x0e950,
                0x06b60,0x0b4b5,0x05ab0,0x02b40,0x1ab60,0x096d5,0x095b0,0x049b0,0x0a4b0,0x0b8a6
            ];
            
            const TG = ['甲','乙','丙','丁','戊','己','庚','辛','壬','癸'];
            const DZ = ['子','丑','寅','卯','辰','巳','午','未','申','酉','戌','亥'];
            const ANIMALS = ['鼠','牛','虎','兔','龙','蛇','马','羊','猴','鸡','狗','猪'];
            const LUNAR_MONTHS = ['正','二','三','四','五','六','七','八','九','十','十一','十二'];
            const LUNAR_DAYS = ['初一','初二','初三','初四','初五','初六','初七','初八','初九','初十',
                '十一','十二','十三','十四','十五','十六','十七','十八','十九','二十',
                '廿一','廿二','廿三','廿四','廿五','廿六','廿七','廿八','廿九','三十'];
            
            function lYearDays(y) {
                let i, sum = 348;
                for (i = 0x8000; i > 0x8; i >>= 1) {
                    sum += (LUNAR_INFO[y - 1900] & i) ? 1 : 0;
                }
                return sum + leapDays(y);
            }
            function leapMonth(y) { return LUNAR_INFO[y - 1900] & 0xf; }
            function leapDays(y) {
                if (leapMonth(y)) {
                    return (LUNAR_INFO[y - 1900] & 0x10000) ? 30 : 29;
                }
                return 0;
            }
            function monthDays(y, m) {
                return (LUNAR_INFO[y - 1900] & (0x10000 >> m)) ? 30 : 29;
            }
            
            function solarToLunar(sYear, sMonth, sDay) {
                let y, m, d, leapYear = false;
                let dayCyclical, monthCyclical;
                
                const baseDate = new Date(1900, 0, 31);
                const objDate = new Date(sYear, sMonth - 1, sDay);
                let offset = Math.round((objDate - baseDate) / 86400000);
                
                let i;
                for (i = 1900; i < 2100 && offset > 0; i++) {
                    let daysInYear = lYearDays(i);
                    offset -= daysInYear;
                }
                if (offset < 0) {
                    offset += lYearDays(--i);
                }
                
                y = i;
                const leap = leapMonth(y);
                leapYear = false;
                
                for (i = 1; i < 13 && offset > 0; i++) {
                    if (leap > 0 && i === leap + 1 && !leapYear) {
                        --i;
                        leapYear = true;
                        d = leapDays(y);
                    } else {
                        d = monthDays(y, i);
                    }
                    if (leapYear && i === leap + 1) leapYear = false;
                    offset -= d;
                }
                
                if (offset === 0 && leap > 0 && i === leap + 1) {
                    if (leapYear) {
                        leapYear = false;
                    } else {
                        leapYear = true;
                        --i;
                    }
                }
                if (offset < 0) {
                    offset += d;
                    --i;
                }
                
                m = i;
                d = offset + 1;
                
                const cyclicalYear = y - 1900 + 36;
                const gan = TG[cyclicalYear % 10];
                const zhi = DZ[cyclicalYear % 12];
                const animal = ANIMALS[cyclicalYear % 12];
                
                return {
                    year: y,
                    month: m,
                    day: d,
                    isLeap: leapYear,
                    ganZhi: gan + zhi,
                    animal,
                    monthStr: (leapYear ? '闰' : '') + LUNAR_MONTHS[m - 1] + '月',
                    dayStr: LUNAR_DAYS[d - 1]
                };
            }
            
            // 节气表（每年近似，精度够用）
            const SOLAR_TERMS = {
                '1-6': '小寒', '1-20': '大寒',
                '2-4': '立春', '2-19': '雨水',
                '3-6': '惊蛰', '3-21': '春分',
                '4-5': '清明', '4-20': '谷雨',
                '5-6': '立夏', '5-21': '小满',
                '6-6': '芒种', '6-21': '夏至',
                '7-7': '小暑', '7-23': '大暑',
                '8-7': '立秋', '8-23': '处暑',
                '9-8': '白露', '9-23': '秋分',
                '10-8': '寒露', '10-23': '霜降',
                '11-7': '立冬', '11-22': '小雪',
                '12-7': '大雪', '12-22': '冬至'
            };
            
            // 法定节假日
            const HOLIDAYS = {
                '1-1': '元旦',
                '2-14': '情人节',
                '3-8': '妇女节',
                '3-12': '植树节',
                '4-4': '清明',
                '4-5': '清明',
                '5-1': '劳动节',
                '5-4': '青年节',
                '6-1': '儿童节',
                '7-1': '建党节',
                '8-1': '建军节',
                '9-9': '重阳',
                '10-1': '国庆节',
                '10-2': '国庆节',
                '10-3': '国庆节',
                '11-11': '双十一',
                '12-25': '圣诞节'
            };
            
            // 农历节日
            const LUNAR_FESTIVALS = {
                '1-1': '春节',
                '1-15': '元宵',
                '5-5': '端午',
                '7-7': '七夕',
                '7-15': '中元',
                '8-15': '中秋',
                '9-9': '重阳',
                '12-30': '除夕',
                '12-29': '除夕'
            };
            
            function getDayInfo(year, month, day) {
                const solarKey = `${month}-${day}`;
                if (HOLIDAYS[solarKey]) return { text: HOLIDAYS[solarKey], isHoliday: true };
                
                const termKey = solarKey;
                if (SOLAR_TERMS[termKey]) return { text: SOLAR_TERMS[termKey], isHoliday: false };
                
                try {
                    const lunar = solarToLunar(year, month, day);
                    const lunarKey = `${lunar.month}-${lunar.day}`;
                    if (LUNAR_FESTIVALS[lunarKey]) return { text: LUNAR_FESTIVALS[lunarKey], isHoliday: true };
                    return { text: lunar.dayStr, isHoliday: false };
                } catch (e) {
                    return { text: '', isHoliday: false };
                }
            }
            
            // 全局状态
            if (!window._v15CalState) {
                window._v15CalState = {
                    year: new Date().getFullYear(),
                    month: new Date().getMonth() + 1
                };
            }
            
            async function render(content) {
                const state = window._v15CalState;
                content.empty();
            
                const container = content.createDiv({ cls: 'calendar-container' });
            
                const today = new Date();
                const todayY = today.getFullYear();
                const todayM = today.getMonth() + 1;
                const todayD = today.getDate();
            
                let { year, month } = state;
            
                // 天干地支年份信息
                try {
                    const lunarYear = solarToLunar(year, month, 1);
                    const yearInfo = container.createDiv({ cls: 'calendar-year-info' });
                    yearInfo.textContent = `${lunarYear.ganZhi}年 · ${lunarYear.animal}年`;
                } catch (e) {}
            
                // 导航栏
                const nav = container.createDiv({ cls: 'calendar-nav' });
                const prevBtn = nav.createEl('button', { cls: 'calendar-nav-btn', text: '‹' });
                const titleEl = nav.createEl('span', {
                    cls: 'calendar-title',
                    text: `${year}年${month}月`
                });
                const todayBtn = nav.createEl('button', { cls: 'calendar-nav-btn', text: '今', attr: { style: 'font-size: 11px; width: 28px;' } });
                const nextBtn = nav.createEl('button', { cls: 'calendar-nav-btn', text: '›' });
            
                prevBtn.addEventListener('click', () => {
                    state.month--;
                    if (state.month < 1) { state.month = 12; state.year--; }
                    render(content);
                });
                nextBtn.addEventListener('click', () => {
                    state.month++;
                    if (state.month > 12) { state.month = 1; state.year++; }
                    render(content);
                });
                todayBtn.addEventListener('click', () => {
                    state.year = todayY;
                    state.month = todayM;
                    render(content);
                });
            
                // 星期头
                const weekdays = container.createDiv({ cls: 'calendar-weekdays' });
                ['日','一','二','三','四','五','六'].forEach(d => {
                    weekdays.createEl('div', { cls: 'calendar-weekday', text: d });
                });
            
                // 构建日期格子
                const grid = container.createDiv({ cls: 'calendar-grid' });
                const firstDay = new Date(year, month - 1, 1).getDay();
                const daysInMonth = new Date(year, month, 0).getDate();
                const daysInPrevMonth = new Date(year, month - 1, 0).getDate();
            
                // 补充上月
                for (let i = firstDay - 1; i >= 0; i--) {
                    const d = daysInPrevMonth - i;
                    const cell = grid.createDiv({ cls: 'calendar-day other-month' });
                    cell.createEl('div', { cls: 'calendar-day-num', text: String(d) });
                    cell.createEl('div', { cls: 'calendar-lunar', text: '' });
                }
            
                // 当月日期
                for (let d = 1; d <= daysInMonth; d++) {
                    const isToday = year === todayY && month === todayM && d === todayD;
                    const dow = new Date(year, month - 1, d).getDay();
                    const isWeekend = dow === 0 || dow === 6;
            
                    let cls = 'calendar-day';
                    if (isToday) cls += ' today';
                    if (isWeekend) cls += ' weekend';
            
                    const cell = grid.createDiv({ cls });
                    cell.createEl('div', { cls: 'calendar-day-num', text: String(d) });
            
                    // 农历 / 节日 / 节气
                    const showLunar = settings.showLunar !== false;
                    const showHoliday = settings.showHoliday !== false;
            
                    if (showLunar || showHoliday) {
                        const dayInfo = getDayInfo(year, month, d);
                        const lunarEl = cell.createEl('div', {
                            cls: dayInfo.isHoliday ? 'calendar-holiday' : 'calendar-lunar',
                            text: dayInfo.text
                        });
                    }
                }
            
                // 补充下月
                const totalCells = firstDay + daysInMonth;
                const remaining = totalCells % 7 === 0 ? 0 : 7 - (totalCells % 7);
                for (let d = 1; d <= remaining; d++) {
                    const cell = grid.createDiv({ cls: 'calendar-day other-month' });
                    cell.createEl('div', { cls: 'calendar-day-num', text: String(d) });
                    cell.createEl('div', { cls: 'calendar-lunar', text: '' });
                }
            }
            
            function renderSettings(containerEl, plugin, saveCallback) {
                const { Setting } = require('obsidian');
            
                containerEl.createEl('h3', { text: '日历模块设置' });
            
                new Setting(containerEl)
                    .setName('显示农历')
                    .setDesc('在每天下方显示农历日期')
                    .addToggle(t => {
                        t.setValue(settings.showLunar !== false)
                            .onChange(async (v) => {
                                settings.showLunar = v;
                                await saveCallback();
                            });
                    });
            
                new Setting(containerEl)
                    .setName('显示节日/节气')
                    .setDesc('在节日和节气当天显示标注')
                    .addToggle(t => {
                        t.setValue(settings.showHoliday !== false)
                            .onChange(async (v) => {
                                settings.showHoliday = v;
                                await saveCallback();
                            });
                    });
            }
            
            module.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };
            
        }
    },
    'directory': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
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
            
        }
    },
    'news': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
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
            
        }
    },
    'recent': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
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
            
        }
    },
    'stats': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
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
            
        }
    },
    'todo': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
            /**
             * 待办模块 V15
             * 格式：V14（含 id/styles/renderSettings）
             * 功能：V11 完整版（增删改查 + 双击编辑 + 筛选 + 进度统计 + 读写 Markdown 文件）
             */
            const id = 'todo';
            const title = '待办事项';
            const icon = '✅';
            
            const defaultSettings = {
                folder: '待办'
            };
            
            const styles = `/* 待办模块样式已在 styles.css 中定义 */`;
            
            // 全局筛选状态
            if (!window._v15TodoState) {
                window._v15TodoState = { filter: 'all' };
            }
            
            function parseTodos(content) {
                const todos = [];
                content.split('\n').forEach(line => {
                    const matchActive = line.match(/^\s*- \[ \] (.*)$/);
                    const matchDone = line.match(/^\s*- \[x\] (.*)$/i);
                    if (matchActive) todos.push({ text: matchActive[1].trim(), completed: false, rawLine: line });
                    else if (matchDone) todos.push({ text: matchDone[1].trim(), completed: true, rawLine: line });
                });
                return todos;
            }
            
            async function ensureTodoFile(folder, filename) {
                const today = moment().format('YYYY-MM-DD');
                let file = app.vault.getAbstractFileByPath(filename);
                if (!file) {
                    const folderExists = app.vault.getAbstractFileByPath(folder);
                    if (!folderExists) {
                        await app.vault.createFolder(folder);
                    }
                    await app.vault.create(filename, `# ${today} 待办事项\n\n`);
                    file = app.vault.getAbstractFileByPath(filename);
                }
                return file;
            }
            
            async function addTodo(filename, text) {
                const file = app.vault.getAbstractFileByPath(filename);
                const c = await app.vault.read(file);
                await app.vault.modify(file, c + `- [ ] ${text}\n`);
            }
            
            async function toggleTodo(filename, todo) {
                const file = app.vault.getAbstractFileByPath(filename);
                const c = await app.vault.read(file);
                const lines = c.split('\n');
                const idx = lines.findIndex(l => l === todo.rawLine);
                if (idx >= 0) {
                    lines[idx] = todo.completed
                        ? lines[idx].replace(/- \[x\]/i, '- [ ]')
                        : lines[idx].replace('- [ ]', '- [x]');
                    await app.vault.modify(file, lines.join('\n'));
                }
            }
            
            async function deleteTodo(filename, todo) {
                const file = app.vault.getAbstractFileByPath(filename);
                const c = await app.vault.read(file);
                const lines = c.split('\n');
                const idx = lines.findIndex(l => l === todo.rawLine);
                if (idx >= 0) {
                    lines.splice(idx, 1);
                    await app.vault.modify(file, lines.join('\n'));
                }
            }
            
            async function editTodo(filename, todo, newText) {
                const file = app.vault.getAbstractFileByPath(filename);
                const c = await app.vault.read(file);
                const lines = c.split('\n');
                const idx = lines.findIndex(l => l === todo.rawLine);
                if (idx >= 0) {
                    const prefix = todo.completed ? '- [x] ' : '- [ ] ';
                    lines[idx] = prefix + newText;
                    await app.vault.modify(file, lines.join('\n'));
                }
            }
            
            async function render(content) {
                const state = window._v15TodoState;
                content.empty();
            
                const folder = settings.folder || '待办';
                const today = moment().format('YYYY-MM-DD');
                const filename = `${folder}/${today}.md`;
            
                const container = content.createDiv({ cls: 'todo-container' });
            
                // 输入区域
                const inputArea = container.createDiv({ cls: 'todo-input-area' });
                const inputWrapper = inputArea.createDiv({ cls: 'todo-input-wrapper' });
                inputWrapper.createDiv({ cls: 'todo-input-icon', text: '⭕' });
                const input = inputWrapper.createEl('input', {
                    cls: 'todo-input',
                    attr: { placeholder: '添加新待办，按 Enter 确认...' }
                });
                input.addEventListener('keypress', async (e) => {
                    if (e.key === 'Enter' && input.value.trim()) {
                        await ensureTodoFile(folder, filename);
                        await addTodo(filename, input.value.trim());
                        input.value = '';
                        render(content);
                    }
                });
            
                let todos = [];
                try {
                    await ensureTodoFile(folder, filename);
                    const file = app.vault.getAbstractFileByPath(filename);
                    const fileContent = await app.vault.read(file);
                    todos = parseTodos(fileContent);
                } catch (e) {
                    container.createEl('div', { text: '读取失败: ' + e.message, attr: { style: 'padding: 10px; color: var(--text-muted); font-size: 12px;' } });
                    return;
                }
            
                const completed = todos.filter(t => t.completed).length;
                const total = todos.length;
            
                // 筛选栏
                const filterArea = container.createDiv({ cls: 'todo-filter-area' });
                [
                    { key: 'all', label: `全部 ${total}` },
                    { key: 'active', label: `待办 ${total - completed}` },
                    { key: 'done', label: `完成 ${completed}` }
                ].forEach(f => {
                    const btn = filterArea.createEl('button', {
                        cls: 'todo-filter-btn' + (state.filter === f.key ? ' active' : ''),
                        text: f.label
                    });
                    btn.addEventListener('click', () => {
                        state.filter = f.key;
                        render(content);
                    });
                });
            
                // 进度提示
                if (total > 0) {
                    const progress = container.createDiv({ cls: 'todo-progress' });
                    progress.textContent = `已完成 ${completed} / ${total}，还剩 ${total - completed} 项`;
                }
            
                // 列表区域
                const listArea = container.createDiv({ cls: 'todo-list-area' });
            
                const filtered = todos.filter(t => {
                    if (state.filter === 'active') return !t.completed;
                    if (state.filter === 'done') return t.completed;
                    return true;
                });
            
                if (filtered.length === 0) {
                    const empty = listArea.createDiv({ cls: 'todo-empty' });
                    empty.createEl('div', { text: '📝', cls: 'todo-empty-icon' });
                    empty.createEl('div', {
                        text: state.filter === 'done' ? '还没有完成的事项' : '今天没有待办，加油！',
                        cls: 'todo-empty-text'
                    });
                    return;
                }
            
                filtered.forEach((todo) => {
                    const item = listArea.createDiv({ cls: 'todo-item' + (todo.completed ? ' completed' : '') });
            
                    const checkbox = item.createDiv({ cls: 'todo-checkbox' + (todo.completed ? ' checked' : '') });
                    if (todo.completed) checkbox.textContent = '✓';
            
                    const textEl = item.createEl('div', { text: todo.text, cls: 'todo-text' });
                    const deleteBtn = item.createEl('div', { text: '✕', cls: 'todo-delete' });
            
                    // 点击勾选/取消
                    checkbox.addEventListener('click', async (e) => {
                        e.stopPropagation();
                        await toggleTodo(filename, todo);
                        render(content);
                    });
            
                    // 双击编辑
                    textEl.addEventListener('dblclick', (e) => {
                        e.stopPropagation();
                        const editInput = item.createEl('input', {
                            cls: 'todo-text-edit',
                            attr: { value: todo.text }
                        });
                        textEl.remove();
                        editInput.select();
                        editInput.addEventListener('blur', async () => {
                            const newText = editInput.value.trim();
                            if (newText && newText !== todo.text) {
                                await editTodo(filename, todo, newText);
                            }
                            render(content);
                        });
                        editInput.addEventListener('keypress', async (e) => {
                            if (e.key === 'Enter') {
                                editInput.blur();
                            }
                        });
                        editInput.addEventListener('keydown', (e) => {
                            if (e.key === 'Escape') render(content);
                        });
                    });
            
                    // 删除
                    deleteBtn.addEventListener('click', async (e) => {
                        e.stopPropagation();
                        await deleteTodo(filename, todo);
                        render(content);
                    });
                });
            }
            
            function renderSettings(containerEl, plugin, saveCallback) {
                const { Setting } = require('obsidian');
            
                containerEl.createEl('h3', { text: '待办模块设置' });
            
                new Setting(containerEl)
                    .setName('待办文件夹')
                    .setDesc('存放待办 Markdown 文件的文件夹路径（相对于 Vault 根目录）')
                    .addText(t => {
                        t.setPlaceholder('待办')
                            .setValue(settings.folder || '待办')
                            .onChange(async (v) => {
                                settings.folder = v.trim() || '待办';
                                await saveCallback();
                            });
                    });
            }
            
            module.exports = { id, title, icon, defaultSettings, styles, render, renderSettings };
            
        }
    },
    'web-preview': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
            /**
             * 网页预览模块 — V17 改造
             * 从 iframe 改为 Electron webview，支持登录态持久化
             * viewport + wrapper + webview 三层架构（与 web-video 一致）
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
            .web-preview-webview {
                width: 100%;
                height: 100%;
                border: none;
            }
            `;
            
            async function render(container) {
                container.empty();
            
                // 初始化
                if (!settings.zoom) settings.zoom = 1;
                if (settings.posY === undefined || settings.posY === null) settings.posY = 0;
                if (settings.posX === undefined || settings.posX === null) settings.posX = 0;
            
                let zoom = Number(settings.zoom);
                if (!isFinite(zoom) || zoom <= 0) zoom = 1;
            
                container.style.display = 'flex';
                container.style.flexDirection = 'column';
                container.style.height = '100%';
            
                // ── 工具栏 ──
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
            
                // ── 视口 ──
                const viewport = container.createDiv({ cls: 'web-preview-viewport' });
            
                // ── webview 包装器 ──
                const webviewWrapper = viewport.createDiv({ cls: 'web-preview-wrapper' });
            
                // ── Electron webview（与 web-video 一致，支持登录态） ──
                const webview = document.createElement('webview');
                webview.className = 'web-preview-webview';
                webview.setAttribute('src', settings.url);
                // persist: 前缀使 Cookie 持久化，重启 Obsidian 后登录态不丢失
                webview.setAttribute('partition', 'persist:webpreview-' + (_moduleId || id));
                webview.setAttribute('preload', '');
                webview.setAttribute('allowpopups', '');
                webview.setAttribute('nodeintegration', 'false');
                webview.setAttribute('webpreferences', 'contextIsolation=true, sandbox=true');
            
                webviewWrapper.appendChild(webview);
            
                // ── 缩放和位置 ──
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
            
                // ── 注入 CSS 屏蔽广告 ──
                webview.addEventListener('dom-ready', () => {
                    webview.insertCSS(`
                        .ad, .ads, .advertisement, .popup, .modal-overlay { display: none !important; }
                    `).catch(() => {});
                });
            
                // ── 新窗口在内部打开（登录跳转等） ──
                webview.addEventListener('new-window', (e) => {
                    webview.src = e.url;
                });
            }
            
            function renderSettings(wrapper, plugin, saveCallback) {
                new Setting(wrapper)
                    .setName('预览网址')
                    .setDesc('使用 Electron webview 打开，支持登录态持久化')
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
            
        }
    },
    'web-video': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
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
            
        }
    },
    'ai-insight': function(module, exports, require, app, plugin, moment, Notice, requestUrl, setIcon, _runtimeCtx) {
        with (_runtimeCtx) {
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
            
        }
    }
};


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

        for (const [moduleId, fn] of Object.entries(BUILTIN_MODULES)) {
            try {
                const mod = this._evalModule(fn, moduleId);
                if (mod && mod.render) {
                    const id = mod.id || moduleId;
                    this.modules.set(id, mod);
                    // 捕获模块导出的默认设置
                    if (mod.defaultSettings && typeof mod.defaultSettings === 'object') {
                        this._moduleDefaults.set(id, mod.defaultSettings);
                    }
                }
            } catch (e) {
                console.warn('[V17] 模块 ' + moduleId + ' 加载失败:', e);
            }
        }

        console.log('[V17] 已加载内联模块: ' + [...this.modules.keys()].join(', '));
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

    _evalModule(fn, fallbackId) {
        try {
            const moduleExports = {};
            const module = { exports: moduleExports };
            const exports = moduleExports;
            const _require = (pkg) => {
                if (pkg === 'obsidian') return require('obsidian');
                throw new Error('Unknown module: ' + pkg);
            };

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
            console.error('[V17] 模块执行错误 (' + fallbackId + '):', e);
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
            console.error('[V17] 模块 ' + moduleId + ' 渲染错误:', e);
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
    getDisplayText() { return '仪表盘主页'; }
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
        left.createEl('span', { text: '仪表盘主页', cls: 'v15-header-title' });

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

        containerEl.createEl('h2', { text: '仪表盘主页 V17 设置' });

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
            console.error('[V17] 模块 ' + moduleId + ' 设置渲染失败:', e);
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

        // 打赏二维码（图床外链）
        const qrSrc = 'https://img-reg-ab.imagency.cn/e/19467f4b916c082ee6ef3b9d81aa9ecb.png';

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
                text: '二维码加载失败',
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

        this.addRibbonIcon('layout-dashboard', '仪表盘主页', () => this.activateView());

        this.addCommand({
            id: 'open-dashboard',
            name: '打开仪表盘主页',
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