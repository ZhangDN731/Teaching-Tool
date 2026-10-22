/* ================================================================
 * Python 程序结构（顺序 · 分支 · for 循环 · while 循环）
 * ================================================================ */
MOD.pystruct = {
  name: 'Python程序结构', sub: '顺序结构 · 分支结构 · 循环结构',
  badge: '代码 · 流程', tag: '代码', cat: 'code',
  color: '#F97316', soft: '#FFF7ED', type: '顺序 · if 分支 · for/while 循环',
  desc: '逐步演示 Python 三种基本程序结构的执行流程：顺序结构逐行执行、分支结构按条件选择路径、循环结构重复执行代码块。',
  icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h6v6H4zM14 4h6v6h-6zM4 14h6v6H4zM14 14h6v6h-6z"/><path d="M10 7h4M10 17h4M7 10v4M17 10v4"/></svg>',
  setDesc: '演示代码自动生成，无需额外参数。',
  stageDesc: '蓝色高亮 = 当前执行行，灰色 = 已执行；实时显示变量状态和控制台输出。',
  initDesc: '点击「自动演示」，逐步观看程序的执行过程。',
  currentSub: 'seq',
  subTabs: [
    { id: 'seq', label: '顺序结构', method: '逐行执行' },
    { id: 'if', label: '分支结构', method: 'if-else' },
    { id: 'for', label: 'for 循环', method: 'range 累加' },
    { id: 'while', label: 'while 循环', method: '条件计数' }
  ],
  theory() {
    return `<div class="rule"><b>顺序结构</b>：代码从上到下依次执行，是最基本的程序结构。</div>
    <div class="rule"><b>分支结构</b>：用 if / elif / else 根据条件选择执行不同的代码块，条件为 True 才执行对应分支。</div>
    <div class="rule"><b>循环结构</b>：for 遍历序列（如 range），while 在条件为 True 时重复执行代码块。</div>
    <div class="rule"><b>缩进</b>：Python 用缩进来表示代码块归属，同一代码块缩进必须一致。</div>
    <div class="rule"><b>print()</b>：输出函数，把内容显示到控制台；逗号分隔多个输出项。</div>`;
  },
  settingsHTML() {
    return `<div class="inline-setting"><div class="setting"><label>程序代码</label><span class="val">自动生成</span></div></div>`;
  },
  afterSettings() { /* 无动态参数；切换子页面由 tab 触发 Engine.reload() */ },
  /* ---------- 通用画布辅助 ---------- */
  tf(v) { return v ? 'True' : 'False'; },
  codeLines(lines, st) {
    const done = st.done || [], skip = st.skip || [];
    return lines.map((l, i) => {
      let cls = '', style = 'style="padding:0 8px"';
      if (done.indexOf(i) >= 0) cls = 'done';
      else if (st.cur === i) cls = 'cur';
      else if (skip.indexOf(i) >= 0) style = 'style="padding:0 8px;opacity:.35"';
      return `<div class="l ${cls}" ${style}>${l || ' '}</div>`;
    }).join('');
  },
  varRack(vars, hotKeys) {
    hotKeys = hotKeys || [];
    return Object.entries(vars).map(([k, v], i) => {
      const disp = typeof v === 'string' ? `"${esc(v)}"` : v;
      return `<span class="var-pill write-in ${hotKeys.indexOf(k) >= 0 ? 'hot' : ''}" style="padding:3px 9px;animation-delay:${i * 0.08}s"><b>${k}</b>=${disp}</span>`;
    }).join('') || `<span style="opacity:.4;font-size:12px">（无变量）</span>`;
  },
  consoleLines(out) {
    if (!out || !out.length) return `<span style="opacity:.4">（无输出）</span>`;
    return out.map((l, i) => `<div class="cl write-in" style="animation-delay:${i * 0.06}s">${esc(l)}</div>`).join('');
  },
  /* ========== seq：顺序结构（逐行赋值、计算、输出） ========== */
  genSeq() {
    const a = randInt(3, 9), b = randInt(2, 8), c = a + b;
    const lines = [
      `a = ${a}`,
      `b = ${b}`,
      `c = a + b`,
      `print("a + b =", c)`
    ];
    const steps = [
      { desc: `第 1 行：赋值 <b class="c-op">a = ${a}</b>。`, vars: { a }, hot: ['a'], out: [], cur: 0, done: [] },
      { desc: `第 2 行：赋值 <b class="c-op">b = ${b}</b>。`, vars: { a, b }, hot: ['b'], out: [], cur: 1, done: [0] },
      { desc: `第 3 行：计算 <b class="c-op">c = a + b = ${a} + ${b} = ${c}</b>。`, vars: { a, b, c }, hot: ['c'], out: [], cur: 2, done: [0, 1] },
      { desc: `第 4 行：print 输出 <b class="c-op">a + b = ${c}</b>。`, vars: { a, b, c }, hot: [], out: [`a + b = ${c}`], cur: 3, done: [0, 1, 2] },
      { desc: `顺序结构执行完毕：代码从上到下依次执行了 4 行，没有分支，也没有跳转。`, vars: { a, b, c }, hot: [], out: [`a + b = ${c}`], cur: -1, done: [0, 1, 2, 3] }
    ];
    return { lines, steps };
  },
  /* ========== if：分支结构（if-else 条件判断） ========== */
  genIf() {
    const score = randInt(45, 98);
    const pass = score >= 60;
    const lines = [
      `score = ${score}`,
      `if score >= 60:`,
      `    print("及格")`,
      `else:`,
      `    print("不及格")`
    ];
    const steps = [];
    steps.push({ desc: `第 1 行：赋值 <b class="c-op">score = ${score}</b>。`, vars: { score }, hot: ['score'], out: [], cur: 0, done: [] });
    steps.push({ desc: `第 2 行：判断条件 <b class="c-op">score >= 60</b> → ${score} >= 60 = <b class="c-op">${this.tf(pass)}</b>。`, vars: { score }, hot: ['score'], out: [], cur: 1, done: [0], cond: `${score} >= 60 = ${this.tf(pass)}` });
    if (pass) {
      steps.push({ desc: `条件为 True，执行 if 分支第 3 行：<b class="c-op">print("及格")</b>。`, vars: { score }, hot: [], out: ['及格'], cur: 2, done: [0, 1] });
      steps.push({ desc: `只有条件为 False 才执行 else，因此第 4、5 行被跳过；程序结束，最终输出「及格」。`, vars: { score }, hot: [], out: ['及格'], cur: -1, done: [0, 1, 2], skip: [3, 4] });
    } else {
      steps.push({ desc: `条件为 False，跳过 if 分支第 3 行（不执行 print("及格")）。`, vars: { score }, hot: [], out: [], cur: 1, done: [0], skip: [2] });
      steps.push({ desc: `进入 else 分支，执行第 5 行：<b class="c-op">print("不及格")</b>。`, vars: { score }, hot: [], out: ['不及格'], cur: 4, done: [0, 1], skip: [2, 3] });
      steps.push({ desc: `程序结束；最终输出「不及格」。`, vars: { score }, hot: [], out: ['不及格'], cur: -1, done: [0, 1, 4], skip: [2, 3] });
    }
    return { lines, steps };
  },
  /* ========== for：for 循环（range 遍历累加） ========== */
  genFor() {
    const n = randInt(3, 5);
    const lines = [
      `total = 0`,
      `for i in range(1, ${n + 1}):`,
      `    total = total + i`,
      `print("总和 =", total)`
    ];
    const steps = [];
    let total = 0;
    steps.push({ desc: `第 1 行：初始化 <b class="c-op">total = 0</b>。`, vars: { total }, hot: ['total'], out: [], cur: 0, done: [] });
    for (let i = 1; i <= n; i++) {
      const before = total;
      total += i;
      steps.push({ desc: `第 2 行：进入循环，本轮 <b class="c-op">i = ${i}</b>（共 ${n} 轮）。`, vars: { total: before, i }, hot: ['i'], out: [], cur: 1, done: [0] });
      steps.push({ desc: `第 3 行：<b class="c-op">total = total + i = ${before} + ${i} = ${total}</b>。`, vars: { total, i }, hot: ['total', 'i'], out: [], cur: 2, done: [0, 1] });
    }
    steps.push({ desc: `循环结束，i 已遍历完 range(1, ${n + 1})；第 4 行 print 输出 <b class="c-op">总和 = ${total}</b>。`, vars: { total }, hot: ['total'], out: [`总和 = ${total}`], cur: 3, done: [0, 1, 2] });
    steps.push({ desc: `程序结束：累加结果 <b class="c-op">1 + 2 + … + ${n} = ${total}</b>。`, vars: { total }, hot: [], out: [`总和 = ${total}`], cur: -1, done: [0, 1, 2, 3] });
    return { lines, steps };
  },
  /* ========== while：while 循环（条件循环计数倒数） ========== */
  genWhile() {
    const n = randInt(3, 4);
    const lines = [
      `count = ${n}`,
      `while count > 0:`,
      `    print(count)`,
      `    count = count - 1`,
      `print("结束")`
    ];
    const steps = [];
    let count = n;
    const out = [];
    steps.push({ desc: `第 1 行：初始化 <b class="c-op">count = ${n}</b>。`, vars: { count }, hot: ['count'], out: [], cur: 0, done: [] });
    while (count > 0) {
      const curCount = count;
      steps.push({ desc: `第 2 行：判断条件 count > 0 → ${curCount} > 0 = <b class="c-op">True</b>，进入循环体。`, vars: { count: curCount }, hot: ['count'], out: [...out], cur: 1, done: [0], cond: `${curCount} > 0 = True` });
      steps.push({ desc: `第 3 行：print(count) 输出 <b class="c-op">${curCount}</b>。`, vars: { count: curCount }, hot: ['count'], out: [...out, String(curCount)], cur: 2, done: [0, 1] });
      count--;
      steps.push({ desc: `第 4 行：count = count - 1 = <b class="c-op">${count}</b>。`, vars: { count }, hot: ['count'], out: [...out, String(curCount)], cur: 3, done: [0, 1, 2] });
      out.push(String(curCount));
    }
    steps.push({ desc: `第 2 行：判断条件 count > 0 → ${count} > 0 = <b class="c-op">False</b>，退出循环。`, vars: { count }, hot: ['count'], out: [...out], cur: 1, done: [0, 1, 2, 3], cond: `${count} > 0 = False` });
    steps.push({ desc: `第 5 行：print("结束")，输出 <b class="c-op">结束</b>。`, vars: { count }, hot: [], out: [...out, '结束'], cur: 4, done: [0, 1, 2, 3] });
    steps.push({ desc: `程序结束：循环共执行 ${n} 轮，先依次输出 ${out.join(' ')}，最后打印「结束」。`, vars: { count }, hot: [], out: [...out, '结束'], cur: -1, done: [0, 1, 2, 3, 4] });
    return { lines, steps };
  },
  /* 初始化演示画布（切换子页面时清空旧内容） */
  initStage() {
    const labelMap = {
      seq: '顺序结构 · 逐行执行',
      if: '分支结构 · if-else 判断',
      for: 'for 循环 · range 累加',
      while: 'while 循环 · 条件计数'
    };
    const codeMap = {
      seq: 'a = 5 → b = 3 → c = a + b → print(...)',
      if: 'if score >= 60: … else: …',
      for: 'for i in range(1, n + 1): total = total + i',
      while: 'while count > 0: … count = count - 1'
    };
    const sub = this.currentSub;
    $('#stage').innerHTML = `<div class="stage-hd" style="margin-bottom:3px"><span>${labelMap[sub] || 'Python 程序执行'}</span><span class="state">准备就绪</span></div>
      <div class="step-title" style="margin-bottom:3px"><span class="st-dot"></span>代码：${codeMap[sub] || ''}</div>
      <div style="text-align:center;padding:clamp(10px,2.5vw,18px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.9vw,16px);margin-bottom:6px">点击「自动演示」或「下一步」逐步执行程序</div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">蓝色 = 当前执行行 · 灰色 = 已执行 · 变量实时更新 · 控制台逐行输出</div>
      </div>`;
  },
  generate() {
    let r;
    switch (this.currentSub) {
      case 'if': r = this.genIf(); break;
      case 'for': r = this.genFor(); break;
      case 'while': r = this.genWhile(); break;
      default: r = this.genSeq();
    }
    const labelMap = {
      seq: '顺序结构 · 逐行执行',
      if: '分支结构 · if-else 判断',
      for: 'for 循环 · range 累加',
      while: 'while 循环 · 条件计数'
    };
    const label = labelMap[this.currentSub] || 'Python 程序执行';
    this.steps = r.steps.map((s, i) => ({
      desc: s.desc,
      draw: () => {
        const code = this.codeLines(r.lines, s);
        const vars = this.varRack(s.vars || {}, s.hot || []);
        const out = this.consoleLines(s.out || []);
        $('#stage').innerHTML = `<div class="stage-hd" style="margin-bottom:3px"><span>${label}</span><span class="state">第 ${i + 1}/${r.steps.length} 步</span></div>
          <div class="code-block" style="padding:5px 8px;line-height:1.5;margin-bottom:6px">${code}</div>
          ${s.cond ? `<div class="exp-inline" style="margin:0 0 6px;padding:2px 8px;line-height:1.4"><span class="term cur">${s.cond}</span></div>` : ''}
          <div style="display:flex;gap:8px;flex-wrap:wrap">
            <div style="flex:1;min-width:150px">
              <div class="zone-label" style="margin-bottom:2px">变量</div>
              <div class="var-rack" style="margin:0;min-height:26px;gap:6px">${vars}</div>
            </div>
            <div style="flex:1;min-width:150px">
              <div class="zone-label" style="margin-bottom:2px">控制台输出</div>
              <div class="console" style="margin:0;min-height:32px;padding:5px 10px;line-height:1.5">${out}</div>
            </div>
          </div>`;
      }
    }));
  }
};
