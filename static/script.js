/**
 * MD → Word  |  Frontend Logic
 * 处理编辑器交互、实时预览、文件拖拽和转换请求
 */

// ============================================================
// DOM 元素
// ============================================================
const markdownInput = document.getElementById('markdownInput');
const previewContent = document.getElementById('previewContent');
const btnConvert = document.getElementById('btnConvert');
const btnPaste = document.getElementById('btnPaste');
const btnClear = document.getElementById('btnClear');
const btnSample = document.getElementById('btnSample');
const btnTheme = document.getElementById('btnTheme');
const fileName = document.getElementById('fileName');
const charCount = document.getElementById('charCount');
const lineCount = document.getElementById('lineCount');
const dropOverlay = document.getElementById('dropOverlay');
const loadingOverlay = document.getElementById('loadingOverlay');
const toastContainer = document.getElementById('toastContainer');

// ============================================================
// 示例 Markdown
// ============================================================
const SAMPLE_MARKDOWN = `# Markdown 转 Word 文档演示

## 文本格式

这是一段普通文本，其中包含 **粗体文字**、*斜体文字* 和 \`行内代码\`。

你还可以使用 ~~删除线~~ 来标记已完成的内容。

## 列表

### 无序列表

- 🎯 支持标题转换（H1 - H6）
- 📝 支持有序和无序列表
  - 支持嵌套列表
  - 多级嵌套也没问题
- 💻 代码块支持语法高亮标签
- 📊 表格渲染美观

### 有序列表

1. 第一步：粘贴 Markdown 文本
2. 第二步：点击"转换并下载"
3. 第三步：打开生成的 Word 文档

## 代码块

\`\`\`python
def hello_world():
    """一个简单的 Python 函数"""
    name = "Gemini"
    print(f"Hello from {name}!")
    return True
\`\`\`

\`\`\`javascript
// JavaScript 示例
const convert = async (markdown) => {
    const response = await fetch('/convert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ markdown })
    });
    return response.blob();
};
\`\`\`

## 表格

| 功能 | 状态 | 说明 |
|------|------|------|
| 标题 | ✅ 已支持 | H1-H6 各级标题 |
| 列表 | ✅ 已支持 | 有序、无序、嵌套 |
| 代码块 | ✅ 已支持 | 带语言标签 |
| 表格 | ✅ 已支持 | 完美渲染 |
| 引用 | ✅ 已支持 | 左侧竖线样式 |

## 引用

> 💡 这是一段引用文字。Markdown 转 Word 工具可以将引用块转换为 Word 中带左边框的段落样式。

## 分隔线

---

## 链接

访问 [GitHub](https://github.com) 获取更多信息。

---

*由 MD → Word 转换工具生成*
`;

// ============================================================
// LaTeX 数学表达式转 Unicode（前端预览用）
// ============================================================
const GREEK_MAP = {
    '\\\\alpha': 'α', '\\\\beta': 'β', '\\\\gamma': 'γ', '\\\\delta': 'δ',
    '\\\\epsilon': 'ε', '\\\\varepsilon': 'ε', '\\\\zeta': 'ζ', '\\\\eta': 'η',
    '\\\\theta': 'θ', '\\\\iota': 'ι', '\\\\kappa': 'κ', '\\\\lambda': 'λ',
    '\\\\mu': 'μ', '\\\\nu': 'ν', '\\\\xi': 'ξ', '\\\\pi': 'π',
    '\\\\rho': 'ρ', '\\\\sigma': 'σ', '\\\\tau': 'τ', '\\\\upsilon': 'υ',
    '\\\\phi': 'φ', '\\\\chi': 'χ', '\\\\psi': 'ψ', '\\\\omega': 'ω',
    '\\\\Gamma': 'Γ', '\\\\Delta': 'Δ', '\\\\Theta': 'Θ', '\\\\Lambda': 'Λ',
    '\\\\Xi': 'Ξ', '\\\\Pi': 'Π', '\\\\Sigma': 'Σ', '\\\\Phi': 'Φ',
    '\\\\Psi': 'Ψ', '\\\\Omega': 'Ω',
};

const MATH_SYM_MAP = {
    '\\\\times': '×', '\\\\div': '÷', '\\\\pm': '±', '\\\\mp': '∓',
    '\\\\cdot': '·', '\\\\leq': '≤', '\\\\le': '≤', '\\\\geq': '≥', '\\\\ge': '≥',
    '\\\\neq': '≠', '\\\\ne': '≠', '\\\\approx': '≈', '\\\\equiv': '≡',
    '\\\\infty': '∞', '\\\\partial': '∂', '\\\\nabla': '∇',
    '\\\\sum': '∑', '\\\\prod': '∏', '\\\\int': '∫',
    '\\\\forall': '∀', '\\\\exists': '∃', '\\\\in': '∈', '\\\\notin': '∉',
    '\\\\subset': '⊂', '\\\\supset': '⊃', '\\\\cup': '∪', '\\\\cap': '∩',
    '\\\\emptyset': '∅', '\\\\Rightarrow': '⇒', '\\\\Leftarrow': '⇐',
    '\\\\rightarrow': '→', '\\\\leftarrow': '←', '\\\\to': '→',
    '\\\\ldots': '…', '\\\\cdots': '⋯', '\\\\sqrt': '√',
    '\\\\left': '', '\\\\right': '', '\\\\quad': ' ', '\\\\qquad': '  ',
    '\\\\,': ' ', '\\\\;': ' ', '\\\\!': '',
};

const UNICODE_FRACS = {
    '1/2': '½', '1/3': '⅓', '2/3': '⅔', '1/4': '¼', '3/4': '¾',
    '1/5': '⅕', '2/5': '⅖', '3/5': '⅗', '4/5': '⅘',
    '1/6': '⅙', '5/6': '⅚', '1/7': '⅐', '1/8': '⅛',
    '3/8': '⅜', '5/8': '⅝', '7/8': '⅞', '1/9': '⅑', '1/10': '⅒',
};

const SUP_MAP = { '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', '+': '⁺', '-': '⁻', 'n': 'ⁿ', 'i': 'ⁱ', 'x': 'ˣ' };
const SUB_MAP = { '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉', '+': '₊', '-': '₋', 'a': 'ₐ', 'e': 'ₑ', 'i': 'ᵢ', 'n': 'ₙ', 'x': 'ₓ' };

function convertLatexContent(s) {
    // \text{...}, \mathrm{...}
    s = s.replace(/\\(?:text|mathrm|mathbf|mathit)\s*\{([^{}]+)\}/g, '$1');
    // \frac{a}{b} → HTML fraction display
    s = s.replace(/\\frac\s*\{([^{}]+)\}\s*\{([^{}]+)\}/g, (_, n, d) => {
        return `<span class="math-frac"><span class="frac-num">${n.trim()}</span><span class="frac-den">${d.trim()}</span></span>`;
    });
    // \sqrt{x}
    s = s.replace(/\\sqrt\s*\{([^{}]+)\}/g, '√$1');
    // ^{...} superscript
    s = s.replace(/\^\{([^{}]+)\}/g, (_, c) => [...c].map(ch => SUP_MAP[ch] || ch).join(''));
    s = s.replace(/\^([0-9a-zA-Z])/g, (_, c) => SUP_MAP[c] || `^${c}`);
    // _{...} subscript
    s = s.replace(/_\{([^{}]+)\}/g, (_, c) => [...c].map(ch => SUB_MAP[ch] || ch).join(''));
    s = s.replace(/_([0-9a-zA-Z])/g, (_, c) => SUB_MAP[c] || `_${c}`);
    // Greek letters (sorted by length desc)
    for (const [tex, uni] of Object.entries(GREEK_MAP).sort((a, b) => b[0].length - a[0].length)) {
        s = s.replaceAll(tex.replace(/\\\\/g, '\\'), uni);
    }
    // Math symbols (sorted by length desc)
    for (const [tex, uni] of Object.entries(MATH_SYM_MAP).sort((a, b) => b[0].length - a[0].length)) {
        s = s.replaceAll(tex.replace(/\\\\/g, '\\'), uni);
    }
    return s.replace(/\s+/g, ' ').trim();
}

function convertLatex(text) {
    if (!text.includes('$')) return text;
    // $$...$$ display math
    text = text.replace(/\$\$(.+?)\$\$/gs, (_, c) => convertLatexContent(c));
    // $...$ inline math
    text = text.replace(/(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)/g, (_, c) => convertLatexContent(c));
    return text;
}

// ============================================================
// 简单的 Markdown → HTML 渲染（用于预览）
// ============================================================
function renderMarkdown(md) {
    if (!md.trim()) return '';

    let html = md;

    // 转义 HTML 特殊字符
    html = html.replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');

    // 代码块（必须先于 LaTeX 处理，防止代码块内容被转换）
    const codeBlocks = [];
    html = html.replace(/```(\w*)\n([\s\S]*?)```/g, (match, lang, code) => {
        const placeholder = `⟦CODE_BLOCK_${codeBlocks.length}⟧`;
        codeBlocks.push(`<pre><code class="language-${lang}">${code.trim()}</code></pre>`);
        return placeholder;
    });

    // 行内代码保护
    const inlineCodes = [];
    html = html.replace(/`([^`]+)`/g, (match, code) => {
        const placeholder = `⟦INLINE_CODE_${inlineCodes.length}⟧`;
        inlineCodes.push(`<code>${code}</code>`);
        return placeholder;
    });

    // LaTeX 数学表达式转换（在 HTML 转义之后，避免 HTML 标签被转义）
    html = convertLatex(html);

    // 恢复代码块
    codeBlocks.forEach((block, i) => {
        html = html.replace(`⟦CODE_BLOCK_${i}⟧`, block);
    });


    // 表格
    html = html.replace(/^\|(.+)\|\s*\n\|[-| :]+\|\s*\n((?:\|.+\|\s*\n?)*)/gm, (match, header, body) => {
        const headers = header.split('|').map(h => h.trim()).filter(h => h);
        const rows = body.trim().split('\n').map(row =>
            row.split('|').map(c => c.trim()).filter(c => c)
        );

        let table = '<table><thead><tr>';
        headers.forEach(h => { table += `<th>${h}</th>`; });
        table += '</tr></thead><tbody>';
        rows.forEach(row => {
            table += '<tr>';
            row.forEach(cell => { table += `<td>${cell}</td>`; });
            table += '</tr>';
        });
        table += '</tbody></table>';
        return table;
    });

    // 标题
    html = html.replace(/^######\s+(.+)$/gm, '<h6>$1</h6>');
    html = html.replace(/^#####\s+(.+)$/gm, '<h5>$1</h5>');
    html = html.replace(/^####\s+(.+)$/gm, '<h4>$1</h4>');
    html = html.replace(/^###\s+(.+)$/gm, '<h3>$1</h3>');
    html = html.replace(/^##\s+(.+)$/gm, '<h2>$1</h2>');
    html = html.replace(/^#\s+(.+)$/gm, '<h1>$1</h1>');

    // 水平线
    html = html.replace(/^---+$/gm, '<hr>');

    // 引用
    html = html.replace(/^&gt;\s+(.+)$/gm, '<blockquote>$1</blockquote>');

    // 合并连续的 blockquote
    html = html.replace(/<\/blockquote>\n<blockquote>/g, '<br>');

    // 无序列表（支持 * 和 - 两种前缀）
    // 必须在粗体/斜体处理之前，避免 * 被误匹配
    html = html.replace(/^(\s*)[\*\-]\s+(.+)$/gm, (match, indent, content) => {
        const level = Math.floor(indent.length / 2);
        const bullets = ['•', '○', '■', '◦', '▪'];
        const bullet = bullets[Math.min(level, bullets.length - 1)];
        const marginLeft = level * 24;
        return `<li class="ul-item" style="margin-left:${marginLeft}px"><span class="bullet">${bullet}</span> ${content}</li>`;
    });

    // 有序列表 — 保留原始数字编号
    html = html.replace(/^(\s*)(\d+)\.\s+(.+)$/gm, (match, indent, num, content) => {
        const level = Math.floor(indent.length / 2);
        const marginLeft = level * 24;
        return `<li class="ol-item" style="margin-left:${marginLeft}px"><span class="ol-num">${num}.</span> ${content}</li>`;
    });

    // 包裹连续的 li 为 ul
    html = html.replace(/((?:<li[^>]*>.*<\/li>\n?)+)/g, '<ul class="md-list">$1</ul>');

    // 恢复行内代码
    inlineCodes.forEach((code, i) => {
        html = html.replace(`⟦INLINE_CODE_${i}⟧`, code);
    });

    // 粗体和斜体（在列表处理之后，避免和 * 列表标记冲突）
    html = html.replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>');
    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');

    // 删除线
    html = html.replace(/~~(.+?)~~/g, '<del>$1</del>');

    // 链接
    html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>');

    // 段落：将剩余的非空行包裹为 <p>
    html = html.replace(/^(?!<[a-z]|$)(.+)$/gm, '<p>$1</p>');

    // 清理多余的空行
    html = html.replace(/\n{3,}/g, '\n\n');

    return html;
}

// ============================================================
// 更新预览
// ============================================================
let previewTimer = null;

function updatePreview() {
    const md = markdownInput.value;

    if (!md.trim()) {
        previewContent.innerHTML = `
            <div class="preview-placeholder">
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" opacity="0.3">
                    <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>
                    <polyline points="14,2 14,8 20,8"/>
                    <line x1="16" y1="13" x2="8" y2="13"/>
                    <line x1="16" y1="17" x2="8" y2="17"/>
                </svg>
                <p>在左侧输入 Markdown 文本<br>这里将显示实时预览</p>
            </div>`;
        return;
    }

    previewContent.innerHTML = renderMarkdown(md);
}

function debouncedPreview() {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(updatePreview, 150);
}

// ============================================================
// 更新统计
// ============================================================
function updateStats() {
    const text = markdownInput.value;
    charCount.textContent = `${text.length} 字符`;
    lineCount.textContent = `${text.split('\n').length} 行`;
}

// ============================================================
// 事件监听
// ============================================================

// 编辑器输入
markdownInput.addEventListener('input', () => {
    updateStats();
    debouncedPreview();
});

// Tab 键支持
markdownInput.addEventListener('keydown', (e) => {
    if (e.key === 'Tab') {
        e.preventDefault();
        const start = markdownInput.selectionStart;
        const end = markdownInput.selectionEnd;
        markdownInput.value = markdownInput.value.substring(0, start) + '    ' + markdownInput.value.substring(end);
        markdownInput.selectionStart = markdownInput.selectionEnd = start + 4;
        updateStats();
        debouncedPreview();
    }
});

// 粘贴按钮
btnPaste.addEventListener('click', async () => {
    try {
        const text = await navigator.clipboard.readText();
        markdownInput.value = text;
        updateStats();
        updatePreview();
        showToast('已从剪贴板粘贴', 'success');
    } catch (err) {
        showToast('无法访问剪贴板，请手动粘贴 (Ctrl+V)', 'error');
    }
});

// 清空按钮
btnClear.addEventListener('click', () => {
    if (markdownInput.value.trim() && !confirm('确定要清空所有内容吗？')) return;
    markdownInput.value = '';
    updateStats();
    updatePreview();
    showToast('已清空', 'info');
});

// 全选预览按钮
const btnSelectAllPreview = document.getElementById('btnSelectAllPreview');
if (btnSelectAllPreview) {
    btnSelectAllPreview.addEventListener('click', () => {
        const previewContentWrapper = document.getElementById('previewContent');
        if (!previewContentWrapper || !previewContentWrapper.textContent.trim()) {
            showToast('预览区为空', 'info');
            return;
        }

        const selection = window.getSelection();
        const range = document.createRange();
        range.selectNodeContents(previewContentWrapper);
        selection.removeAllRanges();
        selection.addRange(range);

        showToast('预览内容已全选', 'success');
    });
}

// 示例按钮
btnSample.addEventListener('click', () => {
    markdownInput.value = SAMPLE_MARKDOWN;
    updateStats();
    updatePreview();
    showToast('已加载示例', 'success');
});

// 主题切换
btnTheme.addEventListener('click', () => {
    const current = document.documentElement.getAttribute('data-theme');
    const next = current === 'light' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('theme', next);
    showToast(`已切换到${next === 'light' ? '亮色' : '暗色'}主题`, 'info');
});

// 加载保存的主题
const savedTheme = localStorage.getItem('theme');
if (savedTheme) {
    document.documentElement.setAttribute('data-theme', savedTheme);
}

// 转换按钮
btnConvert.addEventListener('click', async () => {
    const md = markdownInput.value.trim();
    if (!md) {
        showToast('请先输入 Markdown 文本', 'error');
        return;
    }

    const name = fileName.value.trim() || '文档';

    loadingOverlay.classList.add('active');
    btnConvert.disabled = true;

    try {
        const response = await fetch('/convert', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                markdown: md,
                filename: name
            })
        });

        const result = await response.json();

        if (!response.ok) {
            throw new Error(result.error || '转换失败');
        }

        // 检测是否在 pywebview 桌面应用中
        if (window.pywebview && window.pywebview.api) {
            // 桌面应用 → 弹出 macOS 原生保存对话框
            const saveResult = await window.pywebview.api.save_file(
                result.download_id,
                result.filename
            );
            if (saveResult && saveResult.success) {
                showToast(`✅ 文件已保存到: ${saveResult.path}`, 'success');
            } else {
                const errMsg = (saveResult && saveResult.error) || '保存取消';
                if (errMsg !== '用户取消保存') {
                    showToast(`⚠️ ${errMsg}`, 'error');
                }
            }
        } else {
            // 普通浏览器 → 通过隐藏 iframe 触发下载
            const downloadName = encodeURIComponent(result.filename);
            const downloadUrl = `/download/${result.download_id}?name=${downloadName}`;
            const iframe = document.createElement('iframe');
            iframe.style.display = 'none';
            iframe.src = downloadUrl;
            document.body.appendChild(iframe);
            setTimeout(() => {
                document.body.removeChild(iframe);
            }, 5000);
            showToast('✅ Word 文档已生成并下载！', 'success');
        }
    } catch (err) {
        showToast(`❌ ${err.message}`, 'error');
    } finally {
        loadingOverlay.classList.remove('active');
        btnConvert.disabled = false;
    }
});

// ============================================================
// 文件拖拽
// ============================================================
let dragCounter = 0;

document.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragCounter++;
    if (dragCounter === 1) {
        dropOverlay.classList.add('active');
    }
});

document.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dragCounter--;
    if (dragCounter === 0) {
        dropOverlay.classList.remove('active');
    }
});

document.addEventListener('dragover', (e) => {
    e.preventDefault();
});

document.addEventListener('drop', async (e) => {
    e.preventDefault();
    dragCounter = 0;
    dropOverlay.classList.remove('active');

    const files = e.dataTransfer?.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    if (!file.name.endsWith('.md') && !file.name.endsWith('.markdown') && !file.name.endsWith('.txt')) {
        showToast('请拖入 .md 或 .txt 文件', 'error');
        return;
    }

    try {
        const text = await file.text();
        markdownInput.value = text;
        // 自动设置文件名
        const baseName = file.name.replace(/\.(md|markdown|txt)$/, '');
        fileName.value = baseName;
        updateStats();
        updatePreview();
        showToast(`已导入: ${file.name}`, 'success');
    } catch (err) {
        showToast('文件读取失败', 'error');
    }
});

// ============================================================
// 面板拖拽调整
// ============================================================
const panelDivider = document.getElementById('panelDivider');
const editorPanel = document.querySelector('.panel-editor');
const previewPanel = document.querySelector('.panel-preview');
let isResizing = false;

panelDivider.addEventListener('mousedown', (e) => {
    isResizing = true;
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
});

document.addEventListener('mousemove', (e) => {
    if (!isResizing) return;
    const container = document.querySelector('.editor-container');
    const rect = container.getBoundingClientRect();
    const ratio = (e.clientX - rect.left) / rect.width;
    const clamped = Math.max(0.2, Math.min(0.8, ratio));
    editorPanel.style.flex = `${clamped}`;
    previewPanel.style.flex = `${1 - clamped}`;
});

document.addEventListener('mouseup', () => {
    if (isResizing) {
        isResizing = false;
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
    }
});

// ============================================================
// Toast 通知
// ============================================================
function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    toastContainer.appendChild(toast);

    setTimeout(() => {
        toast.style.animation = 'toastOut 0.3s ease-in forwards';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// ============================================================
// 快捷键
// ============================================================
document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + Enter: 转换
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        btnConvert.click();
    }
    // Ctrl/Cmd + Shift + V: 智能粘贴
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'V') {
        e.preventDefault();
        btnPaste.click();
    }
});

// ============================================================
// 初始化
// ============================================================
updateStats();
