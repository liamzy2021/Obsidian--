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
