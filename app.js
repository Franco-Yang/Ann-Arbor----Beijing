// ============================================
// 全局配置
// ============================================
const STORAGE_KEY_SCHEDULE = 'schedule_data';
const STORAGE_KEY_STATUS = 'weekly_status';
const STORAGE_KEY_MESSAGES = 'mom_messages';
const EXCEL_FILE_PATH = 'schedule.xlsx';

// 星期映射
const DAY_MAP_CN = {
    'Monday': '周一', 'Tuesday': '周二', 'Wednesday': '周三',
    'Thursday': '周四', 'Friday': '周五', 'Saturday': '周六', 'Sunday': '周日'
};

// ============================================
// 初始化
// ============================================
document.addEventListener('DOMContentLoaded', () => {
    updateClock();
    loadAllData();
    setInterval(updateClock, 30000);  // 每30秒更新时间
    setInterval(checkCurrentActivity, 30000);
});

// ============================================
// 1. 时间模块
// ============================================
function updateClock() {
    const now = new Date();
    
    // 安娜堡时间
    const aaTime = now.toLocaleTimeString('zh-CN', {
        timeZone: 'America/Detroit', hour12: false,
        hour: '2-digit', minute: '2-digit'
    });
    const aaDate = now.toLocaleDateString('zh-CN', {
        timeZone: 'America/Detroit',
        month: 'long', day: 'numeric', weekday: 'long'
    });
    
    // 北京时间
    const bjTime = now.toLocaleTimeString('zh-CN', {
        timeZone: 'Asia/Shanghai', hour12: false,
        hour: '2-digit', minute: '2-digit'
    });
    const bjDate = now.toLocaleDateString('zh-CN', {
        timeZone: 'Asia/Shanghai',
        month: 'long', day: 'numeric', weekday: 'long'
    });
    
    document.getElementById('annArborTime').textContent = aaTime;
    document.getElementById('annArborDate').textContent = aaDate;
    document.getElementById('beijingTime').textContent = bjTime;
    document.getElementById('beijingDate').textContent = bjDate;
}

function getAnnArborDay() {
    return new Date().toLocaleDateString('en-US', { timeZone: 'America/Detroit', weekday: 'long' });
}

function getAnnArborTimeString() {
    return new Date().toLocaleTimeString('en-US', {
        timeZone: 'America/Detroit', hour12: false,
        hour: '2-digit', minute: '2-digit'
    });
}

// ============================================
// 2. 数据加载
// ============================================
function loadAllData() {
    // 尝试从localStorage加载（你在admin页面更新后会存在这里）
    const savedSchedule = localStorage.getItem(STORAGE_KEY_SCHEDULE);
    
    if (savedSchedule) {
        // 有缓存数据，直接使用
        window.scheduleData = JSON.parse(savedSchedule);
        renderAll();
    } else {
        // 没有缓存，尝试加载Excel文件
        loadFromExcel();
    }
    
    // 加载状态和留言
    renderWeeklyStatus();
    renderMessages();
}

// 从Excel文件加载日程
async function loadFromExcel() {
    try {
        const response = await fetch(EXCEL_FILE_PATH);
        if (!response.ok) throw new Error('Excel文件未找到');
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet);
        
        window.scheduleData = parseExcelRows(rows);
        localStorage.setItem(STORAGE_KEY_SCHEDULE, JSON.stringify(window.scheduleData));
        renderAll();
        
    } catch (error) {
        console.warn('无法加载Excel，使用默认示例数据:', error.message);
        window.scheduleData = getDefaultSchedule();
        localStorage.setItem(STORAGE_KEY_SCHEDULE, JSON.stringify(window.scheduleData));
        renderAll();
    }
}

// 解析Excel行数据
// Excel格式：Day | Start Time | End Time | Activity | Location
function parseExcelRows(rows) {
    const data = {};
    
    rows.forEach(row => {
        // 兼容中英文列名
        const day = row['Day'] || row['星期'] || row['day'];
        if (!day) return;
        
        const dayKey = day.trim();
        if (!data[dayKey]) data[dayKey] = [];
        
        data[dayKey].push({
            start: String(row['Start Time'] || row['开始时间'] || '').trim(),
            end: String(row['End Time'] || row['结束时间'] || '').trim(),
            activity: String(row['Activity'] || row['活动'] || row['activity'] || '').trim(),
            location: String(row['Location'] || row['地点'] || row['location'] || '').trim()
        });
    });
    
    // 按开始时间排序
    Object.keys(data).forEach(day => {
        data[day].sort((a, b) => a.start.localeCompare(b.start));
    });
    
    return data;
}

// 默认日程（仅当Excel加载失败时使用）
function getDefaultSchedule() {
    return {
        'Monday': [
            { start: '08:30', end: '09:30', activity: '早饭', location: '宿舍' },
            { start: '09:30', end: '11:00', activity: '上课 MATH 215', location: 'Mason Hall' },
            { start: '11:00', end: '12:30', activity: '自习', location: '图书馆' },
            { start: '12:30', end: '13:30', activity: '午饭', location: '食堂' },
            { start: '13:30', end: '15:00', activity: '上课 PHYSICS 140', location: 'Randall Lab' },
            { start: '15:00', end: '17:00', activity: '健身', location: '健身房' },
            { start: '17:30', end: '18:30', activity: '晚饭', location: '食堂' },
            { start: '19:00', end: '22:00', activity: '写作业', location: '宿舍' },
            { start: '23:00', end: '23:30', activity: '洗漱睡觉', location: '宿舍' }
        ]
    };
}

// ============================================
// 3. 渲染模块
// ============================================
function renderAll() {
    renderTodaySchedule();
    checkCurrentActivity();
}

function renderTodaySchedule() {
    const day = getAnnArborDay();
    const schedule = window.scheduleData[day] || [];
    const container = document.getElementById('todaySchedule');
    
    if (schedule.length === 0) {
        container.innerHTML = '<div style="color:#999;padding:10px 0;">今天没有安排</div>';
        return;
    }
    
    const currentTime = getAnnArborTimeString();
    
    container.innerHTML = schedule.map(item => {
        const isCurrent = currentTime >= item.start && currentTime < item.end;
        return `
            <div class="schedule-item ${isCurrent ? 'current' : ''}">
                <span class="schedule-time">${item.start} - ${item.end}</span>
                <span class="schedule-activity">${item.activity}</span>
                ${item.location ? `<span class="schedule-location">${item.location}</span>` : ''}
                ${isCurrent ? '<span class="current-indicator">← 现在</span>' : ''}
            </div>
        `;
    }).join('');
}

function checkCurrentActivity() {
    const day = getAnnArborDay();
    const schedule = window.scheduleData[day] || [];
    const currentTime = getAnnArborTimeString();
    const container = document.getElementById('currentActivity');
    
    // 找当前活动
    const current = schedule.find(item => currentTime >= item.start && currentTime < item.end);
    
    if (current) {
        container.innerHTML = `${current.activity}${current.location ? ` · ${current.location}` : ''}`;
        return;
    }
    
    // 找下一个活动
    const next = schedule.find(item => currentTime < item.start);
    
    if (next) {
        const bjTime = convertToBeijing(next.start);
        container.innerHTML = `休息中，接下来 ${next.start} ${next.activity}（北京${bjTime}）`;
        return;
    }
    
    container.innerHTML = '今天的安排都结束了';
}

function convertToBeijing(timeStr) {
    const [h, m] = timeStr.split(':').map(Number);
    let bjH = h + 12;
    const label = bjH >= 24 ? '次日' : '';
    if (bjH >= 24) bjH -= 24;
    return `${label}${String(bjH).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
}

function renderWeeklyStatus() {
    const status = localStorage.getItem(STORAGE_KEY_STATUS);
    const container = document.getElementById('weeklyStatus');
    
    if (status) {
        container.innerHTML = status;
    } else {
        container.innerHTML = '<span style="color:#999;">本周还没有更新</span>';
    }
}

// ============================================
// 4. 留言模块
// ============================================
function sendMessage() {
    const input = document.getElementById('messageInput');
    const text = input.value.trim();
    
    if (!text) return;
    
    const messages = JSON.parse(localStorage.getItem(STORAGE_KEY_MESSAGES) || '[]');
    const now = new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' });
    
    messages.push({
        text: text,
        time: now,
        read: false
    });
    
    localStorage.setItem(STORAGE_KEY_MESSAGES, JSON.stringify(messages));
    input.value = '';
    
    // 显示提示
    const hint = document.getElementById('messageHint');
    hint.textContent = '已发送 ✓';
    setTimeout(() => hint.textContent = '', 2000);
    
    renderMessages();
}

function renderMessages() {
    const messages = JSON.parse(localStorage.getItem(STORAGE_KEY_MESSAGES) || '[]');
    const container = document.getElementById('recentMessages');
    
    if (messages.length === 0) {
        container.innerHTML = '';
        return;
    }
    
    // 显示最近3条
    const recent = messages.slice(-3).reverse();
    container.innerHTML = recent.map(msg => `
        <div class="message-item">
            <div class="message-text">${msg.text}</div>
            <div class="message-meta">${msg.time}</div>
        </div>
    `).join('');
}

function toggleAllMessages() {
    const container = document.getElementById('allMessages');
    const messages = JSON.parse(localStorage.getItem(STORAGE_KEY_MESSAGES) || '[]');
    
    if (container.style.display === 'none') {
        if (messages.length === 0) {
            container.innerHTML = '<div style="color:#999;">暂无留言</div>';
        } else {
            container.innerHTML = messages.slice().reverse().map((msg, i) => `
                <div class="message-item">
                    <div class="message-text">${msg.text}</div>
                    <div class="message-meta">${msg.time}${msg.read ? ' · 已读' : ''}</div>
                </div>
            `).join('');
        }
        container.style.display = 'block';
    } else {
        container.style.display = 'none';
    }
}
