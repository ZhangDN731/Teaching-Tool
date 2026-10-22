/* ================================================================
 * Python 表达式（常量 · 变量 · 运算符 · 基本函数 · 优先级逐步化简）
 * ================================================================ */
MOD.pyexpr = {
  name: 'Python表达式', sub: '常量 · 变量 · 运算符 · 基本函数 · 优先级逐步化简',
  badge: '代码 · 求值', tag: '代码', cat: 'code',
  color: '#F59E0B', soft: '#FFF4DB', type: '优先级 · 常量变量 · 运算符 · 内置函数',
  desc: '演示 Python 表达式按优先级逐步化简求值：算术、函数与幂、括号与函数、比较与逻辑、取余与长度，每种类型独立子页面。',
  icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>',
  setDesc: '变量自动随机生成，无需额外参数。',
  stageDesc: '橙色 = 当前正在计算的子表达式，绿色 = 已经算好的部分。',
  initDesc: '点击「自动演示」或「下一步」，观看 Python 表达式按运算优先级逐步化简求值。',
  currentSub: 'arith',
  subTabs: [
    { id: 'arith', label: '算术表达式', method: '先乘后加' },
    { id: 'func_pow', label: '函数与幂', method: 'abs · **' },
    { id: 'paren', label: '括号与函数', method: '括号优先' },
    { id: 'logic', label: '比较与逻辑', method: '比较 → and' },
    { id: 'mod_len', label: '取余与长度', method: 'len · %' }
  ],
  theory() {
    return `<div class="rule"><b>常量</b>：固定不变的值，如数字 3、字符串 "hello"、布尔 True/False。</div>
    <div class="rule"><b>变量</b>：存储数据的名字，如 a=3，使用时取其当前值代入表达式。</div>
    <div class="rule"><b>运算符优先级</b>（从高到低）：括号 () → 函数调用 → ** 幂 → * / // % → + - → 比较 == != &gt; &lt; → not → and → or。</div>
    <div class="rule"><b>算术运算符</b>：+ 加、- 减、* 乘、/ 除、// 整除、% 取余、** 幂。</div>
    <div class="rule"><b>比较运算符</b>：== 等于、!= 不等于、&gt; 大于、&lt; 小于、&gt;= 大于等于、&lt;= 小于等于，结果为 True 或 False。</div>
    <div class="rule"><b>逻辑运算符</b>：and 与（两边都为 True 才为 True）、or 或、not 非。</div>
    <div class="rule"><b>基本函数</b>：abs(x) 绝对值、len(x) 长度、max(a,b) 最大、min(a,b) 最小、int(x) 转整数、str(x) 转字符串。</div>`;
  },
  settingsHTML() {
    return `<div class="inline-setting"><div class="setting"><label>变量取值</label><span class="val">自动随机生成</span></div></div>`;
  },
  afterSettings() { /* 无动态参数；切换子页面由 tab 触发 Engine.reload() */ },
  /* ---------- 通用画布辅助 ---------- */
  typeChars(text) {
    return [...String(text)].map((ch, i) => `<span class="char-in" style="animation-delay:${i * 0.05}s">${esc(ch)}</span>`).join('');
  },
  varRack(vars, hotKeys) {
    hotKeys = hotKeys || [];
    return `<div class="var-rack">` + Object.entries(vars).map(([k, v], i) => {
      const disp = typeof v === 'string' ? `"${esc(v)}"` : v;
      return `<span class="var-pill write-in ${hotKeys.indexOf(k) >= 0 ? 'hot' : ''}" style="animation-delay:${i * 0.08}s"><b>${k}</b>=${disp}</span>`;
    }).join('') + `</div>`;
  },
  frame(label, state, title) {
    return `<div class="stage-hd"><span>${label}</span><span class="state">${state}</span></div>
      <div class="step-title"><span class="st-dot"></span>${title}</div>`;
  },
  exprCode(text, done) {
    const body = done ? esc(text) : this.typeChars(text);
    return `<div class="code-block" style="margin-bottom:8px"><div class="l ${done ? 'done' : ''}">>>> ${body}${done ? '' : '<span class="caret pen"></span>'}</div></div>`;
  },
  /* ========== arith：算术表达式 a + b * 2（先乘后加） ========== */
  genArith() {
    const a = randInt(3, 9), b = randInt(2, 9);
    const prod = b * 2, sum = a + prod;
    const label = '算术表达式 · a + b * 2';
    const steps = [];
    steps.push({
      desc: `定义变量：<b class="c-op">a = ${a}</b>，<b class="c-op">b = ${b}</b>；待求表达式为 <b>a + b * 2</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '准备', '变量与表达式')
          + this.varRack({ a, b })
          + this.exprCode('a + b * 2');
      }
    });
    steps.push({
      desc: `把变量替换成数值：<b class="c-op">${a} + ${b} * 2</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '代入变量', '代入数值')
          + this.varRack({ a, b })
          + this.exprCode(`${a} + ${b} * 2`)
          + `<div class="exp-inline"><span class="term cur">${a} + ${b} * 2</span></div>`;
      }
    });
    steps.push({
      desc: `乘法优先级高于加法，先算 <b class="c-op">${b} * 2 = ${prod}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '先算乘法', '按优先级化简')
          + this.varRack({ a, b })
          + this.exprCode(`${a} + ${b} * 2`, true)
          + `<div class="exp-inline"><span class="term done">${a}</span><span class="op">+</span><span class="term cur">${b} * 2</span><span class="op">=</span><span class="term cur write-in">${prod}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${a}</span><span class="op">+</span><span class="term done">${prod}</span></div>`;
      }
    });
    steps.push({
      desc: `再算加法：<b class="c-op">${a} + ${prod} = ${sum}</b>，表达式最终结果为 <b class="c-op">${sum}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '完成', '化简到最终值')
          + this.varRack({ a, b })
          + this.exprCode(`${a} + ${b} * 2`, true)
          + `<div class="exp-inline"><span class="term done">${a}</span><span class="op">+</span><span class="term done">${prod}</span><span class="op">=</span><span class="term cur write-in">${sum}</span></div>
          <div class="write-result">${a} + ${b} * 2 = ${sum}</div>`;
      }
    });
    return steps;
  },
  /* ========== func_pow：函数与幂 abs(c) + a ** 2 ========== */
  genFuncPow() {
    const a = randInt(3, 9), c = randInt(-9, -1);
    const absC = Math.abs(c), pw = a * a, sum = absC + pw;
    const label = '函数与幂 · abs(c) + a ** 2';
    const steps = [];
    steps.push({
      desc: `定义变量：<b class="c-op">a = ${a}</b>，<b class="c-op">c = ${c}</b>；待求表达式为 <b>abs(c) + a ** 2</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '准备', '变量与表达式')
          + this.varRack({ a, c })
          + this.exprCode('abs(c) + a ** 2');
      }
    });
    steps.push({
      desc: `把变量替换成数值：<b class="c-op">abs(${c}) + ${a} ** 2</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '代入变量', '代入数值')
          + this.varRack({ a, c })
          + this.exprCode(`abs(${c}) + ${a} ** 2`)
          + `<div class="exp-inline"><span class="term cur">abs(${c}) + ${a} ** 2</span></div>`;
      }
    });
    steps.push({
      desc: `先算函数调用：<b class="c-op">abs(${c}) = ${absC}</b>（绝对值）。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '先算函数', '按优先级化简')
          + this.varRack({ a, c })
          + this.exprCode(`abs(${c}) + ${a} ** 2`, true)
          + `<div class="exp-inline"><span class="term cur">abs(${c})</span><span class="op">=</span><span class="term cur write-in">${absC}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${absC}</span><span class="op">+</span><span class="term">${a} ** 2</span></div>`;
      }
    });
    steps.push({
      desc: `再算幂运算：<b class="c-op">${a} ** 2 = ${pw}</b>（${a} 的平方）。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '再算幂', '按优先级化简')
          + this.varRack({ a, c })
          + this.exprCode(`abs(${c}) + ${a} ** 2`, true)
          + `<div class="exp-inline"><span class="term done">${absC}</span><span class="op">+</span><span class="term cur">${a} ** 2</span><span class="op">=</span><span class="term cur write-in">${pw}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${absC}</span><span class="op">+</span><span class="term done">${pw}</span></div>`;
      }
    });
    steps.push({
      desc: `最后算加法：<b class="c-op">${absC} + ${pw} = ${sum}</b>，表达式最终结果为 <b class="c-op">${sum}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '完成', '化简到最终值')
          + this.varRack({ a, c })
          + this.exprCode(`abs(${c}) + ${a} ** 2`, true)
          + `<div class="exp-inline"><span class="term done">${absC}</span><span class="op">+</span><span class="term done">${pw}</span><span class="op">=</span><span class="term cur write-in">${sum}</span></div>
          <div class="write-result">abs(${c}) + ${a} ** 2 = ${sum}</div>`;
      }
    });
    return steps;
  },
  /* ========== paren：括号与函数 (a + b) * 3 - max(a,b) ========== */
  genParen() {
    const a = randInt(2, 9), b = randInt(2, 9);
    const inside = a + b, mx = Math.max(a, b), prod = inside * 3, ans = prod - mx;
    const label = '括号与函数 · (a + b) * 3 - max(a,b)';
    const steps = [];
    steps.push({
      desc: `定义变量：<b class="c-op">a = ${a}</b>，<b class="c-op">b = ${b}</b>；待求表达式为 <b>(a + b) * 3 - max(a,b)</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '准备', '变量与表达式')
          + this.varRack({ a, b })
          + this.exprCode('(a + b) * 3 - max(a,b)');
      }
    });
    steps.push({
      desc: `把变量替换成数值：<b class="c-op">(${a} + ${b}) * 3 - max(${a},${b})</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '代入变量', '代入数值')
          + this.varRack({ a, b })
          + this.exprCode(`(${a} + ${b}) * 3 - max(${a},${b})`)
          + `<div class="exp-inline"><span class="term cur">(${a} + ${b}) * 3 - max(${a},${b})</span></div>`;
      }
    });
    steps.push({
      desc: `括号优先级最高，先算括号内：<b class="c-op">${a} + ${b} = ${inside}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '先算括号', '括号优先')
          + this.varRack({ a, b })
          + this.exprCode(`(${a} + ${b}) * 3 - max(${a},${b})`, true)
          + `<div class="exp-inline"><span class="term cur">(${a} + ${b})</span><span class="op">=</span><span class="term cur write-in">${inside}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${inside} * 3</span><span class="op">-</span><span class="term">max(${a},${b})</span></div>`;
      }
    });
    steps.push({
      desc: `再算函数：<b class="c-op">max(${a},${b}) = ${mx}</b>（取较大值）。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '再算函数', '括号优先')
          + this.varRack({ a, b })
          + this.exprCode(`(${a} + ${b}) * 3 - max(${a},${b})`, true)
          + `<div class="exp-inline"><span class="term done">${inside} * 3</span><span class="op">-</span><span class="term cur">max(${a},${b})</span><span class="op">=</span><span class="term cur write-in">${mx}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${inside} * 3</span><span class="op">-</span><span class="term done">${mx}</span></div>`;
      }
    });
    steps.push({
      desc: `接着算乘法：<b class="c-op">${inside} * 3 = ${prod}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '再算乘法', '按优先级化简')
          + this.varRack({ a, b })
          + this.exprCode(`(${a} + ${b}) * 3 - max(${a},${b})`, true)
          + `<div class="exp-inline"><span class="term cur">${inside} * 3</span><span class="op">=</span><span class="term cur write-in">${prod}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${prod}</span><span class="op">-</span><span class="term done">${mx}</span></div>`;
      }
    });
    steps.push({
      desc: `最后算减法：<b class="c-op">${prod} - ${mx} = ${ans}</b>，表达式最终结果为 <b class="c-op">${ans}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '完成', '化简到最终值')
          + this.varRack({ a, b })
          + this.exprCode(`(${a} + ${b}) * 3 - max(${a},${b})`, true)
          + `<div class="exp-inline"><span class="term done">${prod}</span><span class="op">-</span><span class="term done">${mx}</span><span class="op">=</span><span class="term cur write-in">${ans}</span></div>
          <div class="write-result">(${a} + ${b}) * 3 - max(${a},${b}) = ${ans}</div>`;
      }
    });
    return steps;
  },
  /* ========== logic：比较与逻辑 a > b and c != 0 ========== */
  genLogic() {
    const a = randInt(3, 9), b = randInt(2, 9);
    const c = pick([-4, -3, -2, -1, 1, 2, 3, 4]);
    const p1 = a > b, p2 = c !== 0, r = p1 && p2;
    const tf = v => v ? 'True' : 'False';
    const label = '比较与逻辑 · a > b and c != 0';
    const steps = [];
    steps.push({
      desc: `定义变量：<b class="c-op">a = ${a}</b>，<b class="c-op">b = ${b}</b>，<b class="c-op">c = ${c}</b>；待求表达式为 <b>a > b and c != 0</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '准备', '变量与表达式')
          + this.varRack({ a, b, c })
          + this.exprCode('a > b and c != 0');
      }
    });
    steps.push({
      desc: `把变量替换成数值：<b class="c-op">${a} > ${b} and ${c} != 0</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '代入变量', '代入数值')
          + this.varRack({ a, b, c })
          + this.exprCode(`${a} > ${b} and ${c} != 0`)
          + `<div class="exp-inline"><span class="term cur">${a} > ${b} and ${c} != 0</span></div>`;
      }
    });
    steps.push({
      desc: `先做比较：<b class="c-op">${a} > ${b} = ${tf(p1)}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '比较运算', '先算比较')
          + this.varRack({ a, b, c })
          + this.exprCode(`${a} > ${b} and ${c} != 0`, true)
          + `<div class="exp-inline"><span class="term cur">${a} > ${b}</span><span class="op">=</span><span class="term cur write-in">${tf(p1)}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${tf(p1)}</span><span class="op">and</span><span class="term">${c} != 0</span></div>`;
      }
    });
    steps.push({
      desc: `再做比较：<b class="c-op">${c} != 0 = ${tf(p2)}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '比较运算', '先算比较')
          + this.varRack({ a, b, c })
          + this.exprCode(`${a} > ${b} and ${c} != 0`, true)
          + `<div class="exp-inline"><span class="term done">${tf(p1)}</span><span class="op">and</span><span class="term cur">${c} != 0</span><span class="op">=</span><span class="term cur write-in">${tf(p2)}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${tf(p1)}</span><span class="op">and</span><span class="term done">${tf(p2)}</span></div>`;
      }
    });
    steps.push({
      desc: `最后算逻辑 and：<b class="c-op">${tf(p1)} and ${tf(p2)} = ${tf(r)}</b>，表达式最终结果为 <b class="c-op">${tf(r)}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '完成', '逻辑运算')
          + this.varRack({ a, b, c })
          + this.exprCode(`${a} > ${b} and ${c} != 0`, true)
          + `<div class="exp-inline"><span class="term done">${tf(p1)}</span><span class="op">and</span><span class="term done">${tf(p2)}</span><span class="op">=</span><span class="term cur write-in">${tf(r)}</span></div>
          <div class="write-result">${a} > ${b} and ${c} != 0 = ${tf(r)}</div>`;
      }
    });
    return steps;
  },
  /* ========== mod_len：取余与长度 len(s) + a % b ========== */
  genModLen() {
    const s = pick(['hello', 'world', 'python', 'code', 'data']);
    const a = randInt(11, 20), b = randInt(3, 7);
    const len = s.length, rem = a % b, ans = len + rem;
    const label = '取余与长度 · len(s) + a % b';
    const steps = [];
    steps.push({
      desc: `定义变量：<b class="c-op">s = "${s}"</b>，<b class="c-op">a = ${a}</b>，<b class="c-op">b = ${b}</b>；待求表达式为 <b>len(s) + a % b</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '准备', '变量与表达式')
          + this.varRack({ s, a, b })
          + this.exprCode('len(s) + a % b');
      }
    });
    steps.push({
      desc: `把变量替换成数值：<b class="c-op">len("${s}") + ${a} % ${b}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '代入变量', '代入数值')
          + this.varRack({ s, a, b })
          + this.exprCode(`len("${s}") + ${a} % ${b}`)
          + `<div class="exp-inline"><span class="term cur">len("${s}") + ${a} % ${b}</span></div>`;
      }
    });
    steps.push({
      desc: `先算函数：<b class="c-op">len("${s}") = ${len}</b>（字符串长度）。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '先算长度', '函数优先')
          + this.varRack({ s, a, b })
          + this.exprCode(`len("${s}") + ${a} % ${b}`, true)
          + `<div class="exp-inline"><span class="term cur">len("${s}")</span><span class="op">=</span><span class="term cur write-in">${len}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${len}</span><span class="op">+</span><span class="term">${a} % ${b}</span></div>`;
      }
    });
    steps.push({
      desc: `再算取余：<b class="c-op">${a} % ${b} = ${rem}</b>（余数）。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '再算取余', '% 与 * / 同级')
          + this.varRack({ s, a, b })
          + this.exprCode(`len("${s}") + ${a} % ${b}`, true)
          + `<div class="exp-inline"><span class="term done">${len}</span><span class="op">+</span><span class="term cur">${a} % ${b}</span><span class="op">=</span><span class="term cur write-in">${rem}</span></div>
          <div class="exp-inline" style="margin-top:6px"><span class="term done">${len}</span><span class="op">+</span><span class="term done">${rem}</span></div>`;
      }
    });
    steps.push({
      desc: `最后算加法：<b class="c-op">${len} + ${rem} = ${ans}</b>，表达式最终结果为 <b class="c-op">${ans}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame(label, '完成', '化简到最终值')
          + this.varRack({ s, a, b })
          + this.exprCode(`len("${s}") + ${a} % ${b}`, true)
          + `<div class="exp-inline"><span class="term done">${len}</span><span class="op">+</span><span class="term done">${rem}</span><span class="op">=</span><span class="term cur write-in">${ans}</span></div>
          <div class="write-result">len("${s}") + ${a} % ${b} = ${ans}</div>`;
      }
    });
    return steps;
  },
  /* 初始化演示画布（切换子页面时清空旧内容） */
  initStage() {
    const labelMap = {
      arith: '算术表达式 · 先乘后加',
      func_pow: '函数与幂 · abs / **',
      paren: '括号与函数 · 括号优先',
      logic: '比较与逻辑 · and',
      mod_len: '取余与长度 · len / %'
    };
    const exprMap = {
      arith: 'a + b * 2',
      func_pow: 'abs(c) + a ** 2',
      paren: '(a + b) * 3 - max(a,b)',
      logic: 'a > b and c != 0',
      mod_len: 'len(s) + a % b'
    };
    const sub = this.currentSub;
    $('#stage').innerHTML = `<div class="stage-hd"><span>${labelMap[sub] || 'Python 表达式求值'}</span><span class="state">准备就绪</span></div>
      <div class="step-title"><span class="st-dot"></span>表达式：${exprMap[sub] || ''}</div>
      <div style="text-align:center;padding:clamp(18px,3.5vw,30px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.9vw,16px);margin-bottom:8px">点击「自动演示」或「下一步」开始逐步化简</div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">橙色 = 当前计算项 · 绿色 = 已算好 · 灰色 = 运算符</div>
      </div>`;
  },
  generate() {
    switch (this.currentSub) {
      case 'func_pow': this.steps = this.genFuncPow(); break;
      case 'paren': this.steps = this.genParen(); break;
      case 'logic': this.steps = this.genLogic(); break;
      case 'mod_len': this.steps = this.genModLen(); break;
      default: this.steps = this.genArith();
    }
  }
};
