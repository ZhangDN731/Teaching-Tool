/* ================================================================
 * Python 程序结构（顺序 · 分支 · for 循环 · while 循环）
 * 支持用户在代码框中自定义 Python 代码，内置轻量解释器模拟器：
 * 赋值 / print / if-elif-else / for range / while / 算术·比较·逻辑运算
 * ================================================================ */

/* 模块专属样式：仅作用于 #stage 画布；大屏字号放大，变量面板纵向排列 */
if (!document.getElementById('mod-pystruct-css')) {
  document.head.insertAdjacentHTML('beforeend', `<style id="mod-pystruct-css">
  #stage .pystruct-wrap .ps-code{font-size:clamp(13px,1.8vw,17px);line-height:1.9;padding:clamp(8px,1.4vw,12px) clamp(8px,1.2vw,12px)}
  #stage .pystruct-wrap .ps-cols{display:flex;gap:clamp(8px,1.4vw,14px);flex-wrap:wrap;margin-top:clamp(6px,1vw,10px)}
  #stage .pystruct-wrap .ps-panel{flex:1 1 300px;min-width:250px}
  #stage .pystruct-wrap .ps-panel .zone-label{margin-bottom:6px}
  #stage .pystruct-wrap .var-rack{flex-direction:column;align-items:stretch;gap:clamp(6px,1vw,8px);margin:0;min-height:clamp(60px,9vw,90px)}
  #stage .pystruct-wrap .var-pill{display:flex;flex-direction:row;align-items:center;gap:clamp(8px,1.2vw,12px);width:100%;padding:clamp(6px,1vw,9px) clamp(10px,1.4vw,14px);font-size:clamp(13px,1.8vw,17px)}
  #stage .pystruct-wrap .vp-name{color:var(--accent);font-weight:700;min-width:clamp(70px,9vw,100px)}
  #stage .pystruct-wrap .vp-eq{color:var(--text-3)}
  #stage .pystruct-wrap .vp-val{font-family:var(--mono);font-weight:700;color:var(--text);flex:1;word-break:break-all}
  #stage .pystruct-wrap .vp-type{font-size:clamp(10px,1.3vw,12.5px);color:var(--text-3);background:var(--surface-3);border-radius:6px;padding:2px clamp(7px,1vw,10px);font-family:var(--mono);flex:none}
  #stage .pystruct-wrap .var-pill.hot{animation:psHotFlash .7s ease 2}
  #stage .pystruct-wrap .var-pill.var-new{animation:psPopIn .3s ease}
  @keyframes psHotFlash{0%,100%{background:#fff;border-color:var(--border)}50%{background:var(--orange-soft);border-color:var(--orange)}}
  @keyframes psPopIn{from{opacity:0;transform:translateY(6px) scale(.94)}to{opacity:1;transform:none}}
  #stage .pystruct-wrap .ps-console{font-size:clamp(12px,1.6vw,15px);min-height:clamp(60px,9vw,90px)}
  #stage .pystruct-wrap .ps-banner{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:clamp(6px,1vw,8px) 0 0;font-size:clamp(12px,1.6vw,15px)}
  #stage .pystruct-wrap .ps-cond{background:var(--orange-soft);border:1px solid var(--orange);color:#B45309;border-radius:9px;padding:clamp(6px,1vw,9px) clamp(12px,1.6vw,16px);font-family:var(--mono);font-weight:700}
  #stage .pystruct-wrap .ps-loop{background:var(--accent-soft);border:1px solid var(--accent);color:var(--accent);border-radius:9px;padding:clamp(6px,1vw,9px) clamp(12px,1.6vw,16px);font-family:var(--mono);font-weight:700}
  #stage .pystruct-wrap .ps-empty{opacity:.4;font-size:clamp(12px,1.6vw,15px);padding:clamp(8px,1.4vw,12px)}
  /* 仅在本模块设置项存在时约束外层控制条，避免 flex-shrink:0 导致横向溢出 */
  #settingsRowWrap:has(.ps-setting-root){flex:0 0 auto;width:clamp(320px,52vw,640px);min-width:0;max-width:100%}
  #settingsRow:has(.ps-setting-root){display:flex;flex-direction:column;gap:8px;width:100%;min-width:0;max-width:100%}
  #ps-code{font-family:var(--mono);font-size:clamp(13px,1.8vw,16px);line-height:1.7;width:100%;min-width:0;box-sizing:border-box;display:block;min-height:clamp(150px,22vw,210px);padding:clamp(10px,1.6vw,14px);border:1.5px solid var(--border);border-radius:10px;outline:none;resize:vertical;background:#fff;color:var(--text)}
  #ps-code:focus{border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-soft)}
  </style>`);
}

MOD.pystruct = {
  name: 'Python程序结构', sub: '顺序结构 · 分支结构 · 循环结构',
  badge: '代码 · 流程', tag: '代码', cat: 'code',
  color: '#F97316', soft: '#FFF7ED', type: '顺序 · if 分支 · for/while 循环',
  desc: '在代码框中编写 Python 代码，内置模拟器逐步执行：顺序结构逐行执行、分支结构按条件选择路径、循环结构重复执行代码块；变量状态与控制台输出实时可见。',
  icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h6v6H4zM14 4h6v6h-6zM4 14h6v6H4zM14 14h6v6h-6z"/><path d="M10 7h4M10 17h4M7 10v4M17 10v4"/></svg>',
  setDesc: '可直接修改代码框中的示例，或点击预设按钮切换结构；改动后自动生成演示步骤。',
  stageDesc: '蓝色高亮 = 当前执行行，灰色 = 已执行，半透明 = 被跳过；变量面板常驻显示，橙色闪烁 = 刚变化。',
  initDesc: '点击「自动演示」或「下一步」，逐步观看程序执行过程；可在上方代码框中自定义 Python 代码。',
  currentSub: 'seq',
  subTabs: [
    { id: 'seq', label: '顺序结构', method: '逐行执行' },
    { id: 'if', label: '分支结构', method: 'if-else' },
    { id: 'for', label: 'for 循环', method: 'range 遍历' },
    { id: 'while', label: 'while 循环', method: '条件循环' }
  ],
  PRESETS: {
    seq: `a = 5
b = 3
c = a + b
print(c)`,
    if: `a = 10
if a > 5:
    print('big')
else:
    print('small')`,
    for: `s = 0
for i in range(3):
    s = s + i
print(s)`,
    while: `a = 0
while a < 3:
    a = a + 1
print(a)`
  },
  _code: '',
  _run: null,
  steps: [],
  canRegenerate: true,

  theory() {
    return `<div class="rule"><b>顺序结构</b>：代码从上到下依次执行，是最基本的程序结构，没有分支，也没有跳转。</div>
    <div class="rule"><b>分支结构</b>：用 <span class="mono">if / elif / else</span> 根据条件选择执行不同的代码块，条件为 True 才执行对应分支。</div>
    <div class="rule"><b>循环结构</b>：<span class="mono">for</span> 遍历序列（如 <span class="mono">range</span>），<span class="mono">while</span> 在条件为 True 时重复执行代码块，条件为 False 时退出。</div>
    <div class="rule"><b>缩进</b>：Python 用缩进来表示代码块归属，同一代码块缩进必须一致（本示例统一缩进 4 个空格）。</div>
    <div class="rule"><b>print()</b>：输出函数，把内容显示到控制台；逗号分隔多个输出项，输出时以空格隔开。</div>
    <div class="rule"><b>自定义代码</b>：可直接在代码框中修改示例，或点击预设按钮切换结构；支持赋值、print、if/elif/else、for range、while 与基本算术、比较、逻辑运算。</div>`;
  },

  settingsHTML() {
    const code = this._code || this.PRESETS[this.currentSub] || this.PRESETS.seq;
    const labelMap = { seq: '顺序结构', if: '分支结构', for: 'for 循环', while: 'while 循环' };
    const presetBtns = Object.keys(this.PRESETS).map(k =>
      `<button class="btn" type="button" data-preset="${k}">${labelMap[k]}</button>`).join('');
    return `<div class="ps-setting-root" style="flex:1 1 auto;min-width:0;width:100%;display:flex;flex-direction:column;gap:8px">
      <div class="setting" style="width:100%">
        <label>自定义 Python 代码（可直接修改，自动生成演示步骤）</label>
        <textarea id="ps-code" rows="7" spellcheck="false">${esc(code)}</textarea>
      </div>
      <div class="inline-setting" style="width:100%;min-width:0;margin-right:0">
        <div class="setting" style="width:100%">
          <label>预设示例 / 操作</label>
          <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center">
            ${presetBtns}
            <button class="btn primary" type="button" id="ps-run">执行演示</button>
          </div>
        </div>
      </div>
    </div>`;
  },
  afterSettings() {
    const ta = $('#ps-code');
    if (ta) {
      ta.oninput = () => { this._code = ta.value; };
      ta.onchange = () => { this._code = ta.value; Engine.reload(); };
    }
    $$('[data-preset]').forEach(b => b.onclick = () => {
      this._code = this.PRESETS[b.dataset.preset];
      this.currentSub = b.dataset.preset;
      Engine.reload();
    });
    /* 同步子页面 tab 高亮（reload 不会重建 tab 条） */
    $$('#subTabs .sub-tab').forEach(x => x.classList.toggle('active', x.dataset.sub === this.currentSub));
    const run = $('#ps-run');
    if (run) run.onclick = () => {
      this._code = ta ? ta.value : this._code;
      Engine.regenerate();
      Engine.reset();
      Engine.auto();
    };
  },
  onSubChange(sub) {
    this._code = this.PRESETS[sub] || this._code;
  },

  /* ==================== 基础工具 ==================== */
  tf(v) { return v ? 'True' : 'False'; },
  truthy(v) { return v !== false && v !== null && v !== 0 && v !== '' && !Number.isNaN(v); },
  typeOf(v) {
    if (v === null) return 'None';
    if (typeof v === 'boolean') return 'bool';
    if (typeof v === 'string') return 'str';
    if (Number.isInteger(v)) return 'int';
    return 'float';
  },
  fmtVal(v) {
    if (typeof v === 'boolean') return v ? 'True' : 'False';
    if (v === null) return 'None';
    if (typeof v === 'string') return `"${v}"`;
    if (Number.isInteger(v)) return String(v);
    return String(Math.round(v * 100000) / 100000);
  },
  rawVal(v) {
    if (typeof v === 'string') return v;
    return this.fmtVal(v);
  },
  _cmp(a, op, b) {
    switch (op) {
      case '==': return a === b; case '!=': return a !== b;
      case '>': return a > b; case '<': return a < b;
      case '>=': return a >= b; case '<=': return a <= b;
    }
    throw new Error('不支持的比较运算符：' + op);
  },
  _arith(a, op, b) {
    switch (op) {
      case '+': return a + b; case '-': return a - b; case '*': return a * b;
      case '/': return a / b; case '//': return Math.floor(a / b);
      case '%': return a % b; case '**': return Math.pow(a, b);
    }
    throw new Error('不支持的算术运算符：' + op);
  },

  /* ==================== 表达式解析与求值 ==================== */
  _tokenize(src) {
    const s = String(src).trim();
    const toks = [];
    let i = 0;
    while (i < s.length) {
      const c = s[i];
      if (/\s/.test(c)) { i++; continue; }
      if (c === '"' || c === "'") {
        const q = c; let j = i + 1, buf = '';
        while (j < s.length && s[j] !== q) { buf += s[j]; j++; }
        if (j >= s.length) throw new Error('字符串引号不配对');
        toks.push({ t: 'str', v: buf }); i = j + 1; continue;
      }
      let m = s.slice(i).match(/^\d+(\.\d+)?/);
      if (m) { toks.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
      m = s.slice(i).match(/^[A-Za-z_][A-Za-z0-9_]*/);
      if (m) { toks.push({ t: 'id', v: m[0] }); i += m[0].length; continue; }
      const two = s.slice(i, i + 2);
      if (['//', '**', '==', '!=', '<=', '>='].includes(two)) { toks.push({ t: 'op', v: two }); i += 2; continue; }
      if ('+-*/%<>(),'.includes(c)) { toks.push({ t: 'op', v: c }); i++; continue; }
      throw new Error(`无法识别的符号：${c}`);
    }
    return toks;
  },
  evalExpr(src, vars) {
    const toks = this._tokenize(src);
    if (!toks.length) throw new Error('表达式为空');
    let p = 0;
    const peek = () => toks[p];
    const next = () => toks[p++];
    const eat = op => { const k = peek(); if (k && k.t === 'op' && k.v === op) { p++; return true; } return false; };

    const parseOr = () => {
      let v = parseAnd();
      while (peek() && peek().t === 'id' && peek().v === 'or') { next(); v = this.truthy(v) || this.truthy(parseAnd()); }
      return v;
    };
    const parseAnd = () => {
      let v = parseNot();
      while (peek() && peek().t === 'id' && peek().v === 'and') { next(); v = this.truthy(v) && this.truthy(parseNot()); }
      return v;
    };
    const parseNot = () => {
      if (peek() && peek().t === 'id' && peek().v === 'not') { next(); return !this.truthy(parseNot()); }
      return parseCmp();
    };
    const parseCmp = () => {
      let v = parseAdd();
      while (peek() && peek().t === 'op' && ['==', '!=', '>', '<', '>=', '<='].includes(peek().v)) {
        const op = next().v; const r = parseAdd(); v = this._cmp(v, op, r);
      }
      return v;
    };
    const parseAdd = () => {
      let v = parseMul();
      while (peek() && peek().t === 'op' && (peek().v === '+' || peek().v === '-')) {
        const op = next().v; const r = parseMul(); v = op === '+' ? v + r : v - r;
      }
      return v;
    };
    const parseMul = () => {
      let v = parseUnary();
      while (peek() && peek().t === 'op' && ['*', '/', '//', '%', '**'].includes(peek().v)) {
        const op = next().v; const r = parseUnary(); v = this._arith(v, op, r);
      }
      return v;
    };
    const parseUnary = () => {
      if (peek() && peek().t === 'op' && (peek().v === '-' || peek().v === '+')) {
        const op = next().v; const v = parseUnary(); return op === '-' ? -v : v;
      }
      return parseAtom();
    };
    const parseAtom = () => {
      const k = next();
      if (!k) throw new Error('表达式不完整，请检查括号与运算符');
      if (k.t === 'num') return k.v;
      if (k.t === 'str') return k.v;
      if (k.t === 'id') {
        if (k.v === 'True') return true;
        if (k.v === 'False') return false;
        if (k.v === 'None') return null;
        if (eat('(')) {
          if (!eat(')')) { parseOr(); while (eat(',')) parseOr(); if (!eat(')')) throw new Error('括号不匹配'); }
          throw new Error(`不支持的函数调用：${k.v}()`);
        }
        if (k.v in vars) return vars[k.v];
        throw new Error(`变量 ${k.v} 未定义`);
      }
      if (k.t === 'op' && k.v === '(') { const v = parseOr(); if (!eat(')')) throw new Error('括号不匹配'); return v; }
      throw new Error('表达式中存在无法识别的符号');
    };

    const root = parseOr();
    if (p < toks.length) throw new Error('表达式格式不正确');
    return root;
  },
  /* 把表达式中的变量替换为当前值，生成可读的代入文本，如 a + b → 5 + 3 */
  substExpr(src, vars) {
    return this._tokenize(src).map(k => {
      if (k.t === 'id' && k.v in vars) return this.fmtVal(vars[k.v]);
      if (k.t === 'str') return `'${k.v}'`;
      return k.v;
    }).join(' ');
  },
  evalRange(args, vars) {
    let start = 0, stop = 0, step = 1;
    if (args.length === 1) stop = this.evalExpr(args[0], vars);
    else if (args.length === 2) { start = this.evalExpr(args[0], vars); stop = this.evalExpr(args[1], vars); }
    else if (args.length === 3) { start = this.evalExpr(args[0], vars); stop = this.evalExpr(args[1], vars); step = this.evalExpr(args[2], vars); }
    else throw new Error('range 参数个数应为 1~3 个');
    const arr = [];
    if (step > 0) { for (let v = start; v < stop; v += step) arr.push(v); }
    else { for (let v = start; v > stop; v += step) arr.push(v); }
    return arr;
  },
  /* 按顶层逗号拆分 print 参数（忽略字符串与括号内的逗号） */
  splitArgs(s) {
    const out = []; let depth = 0, buf = '', q = null;
    for (let i = 0; i < s.length; i++) {
      const c = s[i];
      if (q) { buf += c; if (c === q) q = null; continue; }
      if (c === '"' || c === "'") { q = c; buf += c; continue; }
      if (c === '(' || c === '[') depth++;
      if (c === ')' || c === ']') depth--;
      if (c === ',' && depth === 0) { out.push(buf.trim()); buf = ''; continue; }
      buf += c;
    }
    if (buf.trim()) out.push(buf.trim());
    return out;
  },

  /* ==================== 代码高亮 ==================== */
  highlightLine(text) {
    let out = ''; const s = text; let i = 0; let q = null;
    while (i < s.length) {
      const c = s[i];
      if (!q && c === '#') { out += `<span class="cm">${esc(s.slice(i))}</span>`; break; }
      if (c === ' ' || c === '\t') { out += c; i++; continue; }
      if (c === '"' || c === "'") {
        q = c; let j = i + 1, buf = '';
        while (j < s.length && s[j] !== q) { buf += s[j]; j++; }
        out += `<span class="str">${q}${esc(buf)}${q}</span>`; q = null; i = j + 1; continue;
      }
      if (/[0-9]/.test(c)) {
        let j = i, num = '';
        while (j < s.length && /[0-9.]/.test(s[j])) { num += s[j]; j++; }
        out += `<span class="num">${num}</span>`; i = j; continue;
      }
      if (/[A-Za-z_]/.test(c)) {
        let j = i, id = '';
        while (j < s.length && /[A-Za-z0-9_]/.test(s[j])) { id += s[j]; j++; }
        if (id === 'print' || id === 'range') out += `<span class="fn">${id}</span>`;
        else if (['if', 'elif', 'else', 'for', 'in', 'while', 'and', 'or', 'not', 'True', 'False', 'None'].includes(id)) out += `<span class="kw">${id}</span>`;
        else out += esc(id);
        i = j; continue;
      }
      out += esc(c); i++;
    }
    return out;
  },

  /* ==================== 轻量解释器：把代码编译为步骤快照 ==================== */
  buildRun(code) {
    const self = this;
    if (!code || !code.trim()) throw new Error('代码框为空：请先输入 Python 代码，或点击预设示例按钮。');
    const rawLines = code.split(/\r?\n/);
    const stripInlineComment = line => {
      let q = null;
      for (let i = 0; i < line.length; i++) {
        const c = line[i];
        if (q) { if (c === q) q = null; continue; }
        if (c === '"' || c === "'") { q = c; continue; }
        if (c === '#') return line.slice(0, i);
      }
      return line;
    };
    const meta = rawLines.map((raw, idx) => {
      const cleaned = stripInlineComment(raw);
      const text = cleaned.trim();
      const lead = raw.match(/^[ \t]*/)[0];
      const indent = lead.replace(/\t/g, '    ').length;
      return { raw, text, indent, idx, blank: text === '', comment: raw.trim().startsWith('#') };
    });
    const execLines = meta.filter(m => !m.blank && !m.comment);
    if (!execLines.length) throw new Error('代码中没有可执行语句。');

    /* ---------- 解析为带缩进层级的语句树 ---------- */
    let pi = 0;
    function parseBlock(minIndent) {
      const out = [];
      while (pi < execLines.length && execLines[pi].indent > minIndent) {
        const ln = execLines[pi];
        const t = ln.text;
        if (/^if\b/.test(t)) out.push(parseIf(ln));
        else if (/^for\b/.test(t)) out.push(parseFor(ln));
        else if (/^while\b/.test(t)) out.push(parseWhile(ln));
        else { out.push({ kind: 'stmt', line: ln }); pi++; }
      }
      return out;
    }
    function parseIf(header) {
      const cond = header.text.replace(/^if\s+/, '').replace(/:\s*$/, '').trim();
      if (!cond) throw new Error(`第 ${header.idx + 1} 行：if 后缺少条件表达式`);
      const branches = [{ cond, line: header, body: [] }];
      pi++;
      branches[0].body = parseBlock(header.indent);
      while (pi < execLines.length && execLines[pi].indent === header.indent) {
        const t = execLines[pi].text;
        if (/^elif\b/.test(t)) {
          const c = t.replace(/^elif\s+/, '').replace(/:\s*$/, '').trim();
          const l = execLines[pi];
          pi++;
          branches.push({ cond: c, line: l, body: parseBlock(l.indent) });
        } else if (/^else\s*:?\s*$/.test(t)) {
          const l = execLines[pi];
          pi++;
          branches.push({ cond: null, line: l, body: parseBlock(l.indent) });
          break;
        } else break;
      }
      return { kind: 'if', branches };
    }
    function parseFor(header) {
      const m = header.text.match(/^for\s+([A-Za-z_]\w*)\s+in\s+range\s*\(([^)]*)\)\s*:?\s*$/);
      if (!m) throw new Error(`第 ${header.idx + 1} 行：for 语句格式应为 for 变量 in range(起, 止):`);
      pi++;
      const body = parseBlock(header.indent);
      return { kind: 'for', line: header, var: m[1], args: m[2].split(',').map(x => x.trim()).filter(Boolean), body };
    }
    function parseWhile(header) {
      const cond = header.text.replace(/^while\s+/, '').replace(/:\s*$/, '').trim();
      if (!cond) throw new Error(`第 ${header.idx + 1} 行：while 后缺少条件表达式`);
      pi++;
      const body = parseBlock(header.indent);
      return { kind: 'while', line: header, cond, body };
    }
    const tree = parseBlock(-1);
    if (pi < execLines.length) throw new Error(`第 ${execLines[pi].idx + 1} 行：缩进层级有误，请检查代码块缩进`);

    /* ---------- 执行并录制步骤快照 ---------- */
    const vars = {};
    const out = [];
    const done = new Set();
    const skip = new Set();
    const steps = [];
    let prevSnapVars = {};
    const sameVal = (a, b) => (typeof a === 'object' ? JSON.stringify(a) === JSON.stringify(b) : a === b);

    const snapshot = extra => {
      const hot = [], fresh = [];
      Object.keys(vars).forEach(k => {
        if (!(k in prevSnapVars)) fresh.push(k);
        else if (!sameVal(prevSnapVars[k], vars[k])) hot.push(k);
      });
      const s = Object.assign({
        cur: -1, done: [...done], skip: [...skip],
        vars: JSON.parse(JSON.stringify(vars)), hot, fresh,
        out: [...out], cond: null, loop: null, desc: ''
      }, extra);
      prevSnapVars = JSON.parse(JSON.stringify(vars));
      steps.push(s);
    };

    function skipStmt(st) {
      skip.add(st.line.idx);
      if (st.body) st.body.forEach(skipStmt);
    }
    function execBlock(stmts) { stmts.forEach(execStmt); }
    function execStmt(st) {
      if (st.kind === 'if') execIf(st);
      else if (st.kind === 'for') execFor(st);
      else if (st.kind === 'while') execWhile(st);
      else execSimple(st.line);
    }

    function execSimple(line) {
      const t = line.text;
      /* print 输出语句 */
      const pm = t.match(/^print\s*\((.*)\)\s*$/s);
      if (pm) {
        const vals = self.splitArgs(pm[1]).map(p => self.evalExpr(p, vars));
        const text = vals.map(v => self.rawVal(v)).join(' ');
        out.push(text);
        done.add(line.idx);
        snapshot({ cur: line.idx, desc: `第 ${line.idx + 1} 行：调用 print 输出 <b class="c-op">${esc(text)}</b>。` });
        return;
      }
      /* 复合赋值 += -= *= /= */
      const am = t.match(/^([A-Za-z_]\w*)\s*(\+=|-=|\*=|\/=)\s*(.+)$/);
      if (am) {
        const name = am[1], op = am[2], rhs = am[3];
        if (!(name in vars)) throw new Error(`第 ${line.idx + 1} 行：变量 ${name} 尚未赋值，不能使用 ${op}`);
        const old = vars[name], rv = self.evalExpr(rhs, vars);
        let nv;
        if (op === '+=') nv = old + rv;
        else if (op === '-=') nv = old - rv;
        else if (op === '*=') nv = old * rv;
        else nv = old / rv;
        vars[name] = nv;
        done.add(line.idx);
        snapshot({
          cur: line.idx,
          desc: `第 ${line.idx + 1} 行：<b class="c-op">${esc(name)} ${op} ${esc(rhs)}</b>，即 ${esc(name)} = ${esc(self.fmtVal(old))} ${op.slice(0, 1)} ${esc(self.fmtVal(rv))} = <b class="c-op">${esc(self.fmtVal(nv))}</b>。`
        });
        return;
      }
      /* 普通赋值 = */
      const sm = t.match(/^([A-Za-z_]\w*)\s*=\s*(.+)$/);
      if (sm) {
        const name = sm[1], rhs = sm[2];
        const val = self.evalExpr(rhs, vars);
        const isSimple = !/[+\-*/%()]/.test(rhs);
        const subst = isSimple ? null : self.substExpr(rhs, vars);
        vars[name] = val;
        done.add(line.idx);
        const desc = isSimple
          ? `第 ${line.idx + 1} 行：赋值 <b class="c-op">${esc(name)} = ${esc(self.fmtVal(val))}</b>。`
          : `第 ${line.idx + 1} 行：计算 <b class="c-op">${esc(name)} = ${esc(rhs)} = ${esc(subst)} = ${esc(self.fmtVal(val))}</b>。`;
        snapshot({ cur: line.idx, desc });
        return;
      }
      /* 裸表达式（仅做语法校验） */
      self.evalExpr(t, vars);
      done.add(line.idx);
      snapshot({ cur: line.idx, desc: `第 ${line.idx + 1} 行：执行语句 <b class="c-op">${esc(t)}</b>。` });
    }

    function execIf(st) {
      let chosen = -1;
      st.branches.forEach((br, bi) => {
        if (chosen >= 0) { skipStmt({ line: br.line, body: br.body }); return; }
        if (br.cond === null) {
          done.add(br.line.idx);
          snapshot({ cur: br.line.idx, desc: `前面条件均不成立，进入 else 分支。` });
          chosen = bi;
          execBlock(br.body);
          return;
        }
        const ev = self.evalExpr(br.cond, vars);
        const truth = self.truthy(ev);
        const sub = self.substExpr(br.cond, vars);
        if (truth) {
          done.add(br.line.idx);
          snapshot({
            cur: br.line.idx,
            cond: `${sub} = ${self.tf(truth)}`,
            desc: `第 ${br.line.idx + 1} 行：判断条件 <b class="c-op">${esc(br.cond)}</b> → ${esc(sub)} = <b class="c-op">${self.tf(truth)}</b>，条件成立，执行该分支。`
          });
          chosen = bi;
          execBlock(br.body);
        } else {
          skip.add(br.line.idx);
          br.body.forEach(skipStmt);
          snapshot({
            cur: br.line.idx,
            cond: `${sub} = ${self.tf(truth)}`,
            desc: `第 ${br.line.idx + 1} 行：判断条件 <b class="c-op">${esc(br.cond)}</b> → ${esc(sub)} = <b class="c-op">${self.tf(truth)}</b>，条件不成立，跳过该分支。`
          });
        }
      });
    }

    function execFor(st) {
      const arr = self.evalRange(st.args, vars);
      const total = arr.length;
      arr.forEach((v, k) => {
        vars[st.var] = v;
        done.add(st.line.idx);
        snapshot({
          cur: st.line.idx,
          loop: `第 ${k + 1}/${total} 轮 · 循环变量 ${st.var} = ${self.fmtVal(v)}`,
          desc: `第 ${st.line.idx + 1} 行：进入第 ${k + 1}/${total} 轮循环，<b class="c-op">${esc(st.var)} = ${esc(self.fmtVal(v))}</b>。`
        });
        execBlock(st.body);
      });
      if (total === 0) {
        done.add(st.line.idx);
        snapshot({ cur: st.line.idx, loop: 'range 为空，循环体执行 0 轮', desc: `range 序列为空，循环体不执行。` });
      }
    }

    function execWhile(st) {
      let iter = 0;
      for (;;) {
        if (iter > 100) throw new Error('while 循环已超过 100 轮，可能是死循环，演示已终止，请检查循环条件。');
        const ev = self.evalExpr(st.cond, vars);
        const truth = self.truthy(ev);
        const sub = self.substExpr(st.cond, vars);
        done.add(st.line.idx);
        snapshot({
          cur: st.line.idx,
          cond: `${sub} = ${self.tf(truth)}`,
          desc: truth
            ? `第 ${st.line.idx + 1} 行：条件 <b class="c-op">${esc(st.cond)}</b> → ${esc(sub)} = <b class="c-op">True</b>，进入循环体。`
            : `第 ${st.line.idx + 1} 行：条件 <b class="c-op">${esc(st.cond)}</b> → ${esc(sub)} = <b class="c-op">False</b>，退出循环。`
        });
        if (!truth) break;
        execBlock(st.body);
        iter++;
      }
    }

    execBlock(tree);
    snapshot({ cur: -1, desc: `程序执行完毕：共输出 ${out.length} 行结果，变量状态已全部记录在左侧监控面板。` });
    return { meta, steps };
  },

  /* ==================== 画布渲染 ==================== */
  _labelMap: {
    seq: '顺序结构 · 逐行执行',
    if: '分支结构 · if-elif-else 判断',
    for: 'for 循环 · range 遍历',
    while: 'while 循环 · 条件循环'
  },
  codeBlock(meta, snap) {
    const doneSet = snap.done, skipSet = snap.skip;
    return meta.map(m => {
      if (m.blank) return `<div class="l" style="padding:0 8px;opacity:.35"> </div>`;
      let cls = '', style = 'padding:0 8px';
      if (snap.cur === m.idx) cls = 'cur';
      else if (doneSet.indexOf(m.idx) >= 0) cls = 'done';
      else if (skipSet.indexOf(m.idx) >= 0) style = 'padding:0 8px;opacity:.3';
      const body = m.comment ? `<span class="cm">${esc(m.raw)}</span>` : this.highlightLine(m.raw);
      return `<div class="l ${cls}" style="${style}">${body}</div>`;
    }).join('');
  },
  varPanel(vars, hot, fresh) {
    const keys = Object.keys(vars);
    if (!keys.length) return `<div class="ps-empty">（尚未执行，暂无变量）</div>`;
    return keys.map(k => {
      const v = vars[k];
      const cls = hot.indexOf(k) >= 0 ? 'hot' : (fresh.indexOf(k) >= 0 ? 'var-new' : '');
      return `<div class="var-pill ${cls}">
        <span class="vp-name">${esc(k)}</span>
        <span class="vp-eq">=</span>
        <span class="vp-val">${this.fmtVal(v)}</span>
        <span class="vp-type">${this.typeOf(v)}</span>
      </div>`;
    }).join('');
  },
  consoleOut(out) {
    if (!out.length) return `<span style="opacity:.4">（暂无输出）</span>`;
    return out.map(l => `<div class="cl write-in">${esc(l)}</div>`).join('');
  },
  renderStep(i) {
    const run = this._run;
    const s = run.steps[i];
    const label = this._labelMap[this.currentSub] || 'Python 程序执行';
    $('#stage').innerHTML = `<div class="pystruct-wrap">
      <div class="stage-hd" style="margin-bottom:4px"><span>${label}</span><span class="state">第 ${i + 1}/${run.steps.length} 步</span></div>
      <div class="code-block ps-code">${this.codeBlock(run.meta, s)}</div>
      ${s.cond ? `<div class="ps-banner"><span class="ps-cond">条件：${esc(s.cond)}</span></div>` : ''}
      ${s.loop ? `<div class="ps-banner"><span class="ps-loop">${esc(s.loop)}</span></div>` : ''}
      <div class="ps-cols">
        <div class="ps-panel">
          <div class="zone-label">变量监控面板（实时）</div>
          <div class="var-rack">${this.varPanel(s.vars, s.hot, s.fresh)}</div>
        </div>
        <div class="ps-panel">
          <div class="zone-label">控制台输出</div>
          <div class="console ps-console">${this.consoleOut(s.out)}</div>
        </div>
      </div>
    </div>`;
  },
  initStage() {
    const code = this._code || this.PRESETS[this.currentSub] || this.PRESETS.seq;
    this._code = code;
    const meta = code.split(/\r?\n/).map((raw, idx) => ({
      raw, text: raw.trim(), idx, blank: raw.trim() === '', comment: raw.trim().startsWith('#')
    }));
    const label = this._labelMap[this.currentSub] || 'Python 程序执行';
    $('#stage').innerHTML = `<div class="pystruct-wrap">
      <div class="stage-hd" style="margin-bottom:4px"><span>${label}</span><span class="state">准备就绪</span></div>
      <div class="code-block ps-code">${meta.map(m =>
        `<div class="l" style="padding:0 8px">${m.blank ? ' ' : this.highlightLine(m.raw)}</div>`).join('')}</div>
      <div class="ps-cols">
        <div class="ps-panel">
          <div class="zone-label">变量监控面板（实时）</div>
          <div class="var-rack"><div class="ps-empty">（点击「下一步」或「自动演示」开始执行）</div></div>
        </div>
        <div class="ps-panel">
          <div class="zone-label">控制台输出</div>
          <div class="console ps-console"><span style="opacity:.4">（暂无输出）</span></div>
        </div>
      </div>
    </div>`;
  },
  generate() {
    const code = (this._code || this.PRESETS[this.currentSub] || '').trim();
    this._code = code;
    const run = this.buildRun(code);
    this._run = run;
    this.steps = run.steps.map((s, i) => ({
      desc: s.desc,
      draw: () => this.renderStep(i)
    }));
  }
};
