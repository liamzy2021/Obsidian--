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
