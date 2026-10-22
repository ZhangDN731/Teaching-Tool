/* ================================================================
 * Python 表达式（数据类型 · 运算符 · 运算符优先级 · 常见函数）
 * 子页面：datatype 数据介绍 / operator 运算符 / precedence 优先级 / function 常见函数
 * ================================================================ */

/* 模块专属样式：仅作用于 #stage 画布，不影响其他模块；大屏字号放大 */
if (!document.getElementById('pyexpr-css')) {
  document.head.insertAdjacentHTML('beforeend', `<style id="pyexpr-css">
#stage .code-block{font-size:clamp(13px,1.8vw,17px)}
#stage .console{font-size:clamp(14px,2vw,18px)}
#stage .py-type-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:clamp(8px,1.3vw,12px);margin:10px 0}
#stage .py-tcard{background:#fff;border:1.5px solid var(--border);border-radius:14px;padding:clamp(10px,1.5vw,14px);transition:all .2s;opacity:.5}
#stage .py-tcard.done{opacity:1;border-color:var(--green);background:var(--green-soft)}
#stage .py-tcard.cur{opacity:1;border-color:var(--orange);background:#FFF8EE;box-shadow:var(--shadow-hover);transform:translateY(-2px)}
#stage .py-tname{font-family:var(--mono);font-weight:700;color:var(--accent);font-size:clamp(13px,1.8vw,17px);margin-bottom:4px}
#stage .py-tval{font-family:var(--mono);color:var(--text);font-size:clamp(12px,1.7vw,16px);margin:2px 0;word-break:break-all}
#stage .py-ttype{font-family:var(--mono);color:var(--text-3);font-size:clamp(11px,1.5vw,13.5px)}
#stage .py-tdesc{font-size:clamp(11px,1.5vw,13.5px);color:var(--text-2);margin-top:6px;line-height:1.6}
#stage .py-cat-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:clamp(8px,1.3vw,12px);margin:10px 0}
#stage .py-cat-card{background:#fff;border:1.5px solid var(--border);border-radius:14px;padding:clamp(10px,1.5vw,14px);transition:all .2s;opacity:.5}
#stage .py-cat-card.done{opacity:1;border-color:var(--green);background:var(--green-soft)}
#stage .py-cat-card.cur{opacity:1;border-color:var(--orange);background:#FFF8EE;box-shadow:var(--shadow-hover);transform:translateY(-2px)}
#stage .py-cat-name{font-weight:700;color:var(--text);font-size:clamp(13px,1.7vw,16px);margin-bottom:4px}
#stage .py-cat-method{font-family:var(--mono);color:var(--text-3);font-size:clamp(11px,1.5vw,13.5px)}
#stage .py-op{color:var(--accent);font-family:var(--mono)}
#stage .py-fn-rack{display:flex;flex-wrap:wrap;gap:8px;margin:10px 0}
#stage .py-fn-chip{font-family:var(--mono);font-weight:700;font-size:clamp(11px,1.5vw,13.5px);border-radius:9px;padding:clamp(5px,.9vw,8px) clamp(10px,1.5vw,14px);background:var(--accent-soft);color:var(--accent);border:1px solid var(--border);opacity:.5;transition:all .15s}
#stage .py-fn-chip.done{opacity:1;background:var(--green-soft);color:#15803D;border-color:var(--green)}
#stage .py-fn-chip.cur{opacity:1;background:var(--orange-soft);color:#B45309;border-color:var(--orange);transform:scale(1.05)}
#stage .py-fn-card{background:#fff;border:1.5px solid var(--border);border-radius:14px;padding:clamp(12px,1.8vw,18px);margin:10px 0;border-left:4px solid var(--accent)}
#stage .py-fn-name{font-family:var(--mono);font-weight:700;color:var(--accent);font-size:clamp(14px,2vw,18px);margin-bottom:6px}
#stage .py-fn-purpose{font-size:clamp(12px,1.6vw,15px);color:var(--text-2);line-height:1.7}
#stage .py-chain{margin-top:12px;background:#fff;border:1px solid var(--border);border-radius:12px;padding:clamp(10px,1.8vw,16px)}
#stage .py-chain-line{font-family:var(--mono);font-size:clamp(13px,1.8vw,17px);line-height:2.1;word-break:break-all}
#stage .py-chain-hd{font-family:var(--font);font-weight:700;color:var(--text-3);font-size:clamp(11px,1.4vw,13px);margin-right:8px}
#stage .expr-lit{color:#F78C6C;font-weight:700}
#stage .expr-op{color:#94A3B8;margin:0 4px}
#stage .expr-cur{background:var(--orange);color:#fff;border-radius:6px;padding:1px 8px;font-weight:700}
#stage .expr-done{color:var(--green);font-weight:700;background:var(--green-soft);border-radius:6px;padding:1px 8px}
#stage .py-prec td.op{font-family:var(--mono);font-weight:700;color:var(--accent)}
@media (max-width:768px){
  #stage .py-type-grid{grid-template-columns:repeat(2,1fr)}
  #stage .py-cat-grid{grid-template-columns:repeat(2,1fr)}
}
</style>`);
}

MOD.pyexpr = {
  name: 'Python表达式', sub: '数据类型 · 运算符 · 运算符优先级 · 常见内置函数',
  badge: '代码 · 演示', tag: '代码', cat: 'code',
  color: '#F59E0B', soft: '#FFF4DB', type: '数据类型 · 运算符 · 优先级 · 内置函数',
  desc: '分四个子页面系统讲解 Python 基础：八种内置数据类型、六类运算符、运算符优先级逐步化简、十三个常见内置函数，均配有示例代码与运行结果。',
  icon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>',
  setDesc: '在上方选择要浏览的类型 / 运算符类别 / 函数；优先级页面可选择预设表达式或自行输入简单表达式。',
  stageDesc: '橙色 = 当前正在讲解 / 计算的内容，绿色 = 已完成，灰色 = 待浏览。',
  initDesc: '点击「自动演示」或「下一步」，逐个浏览数据类型、运算符、优先级化简与内置函数。',
  currentSub: 'datatype',
  subTabs: [
    { id: 'datatype', label: '数据类型', method: 'int · str · list …' },
    { id: 'operator', label: '运算符', method: '算术 · 比较 · 逻辑 …' },
    { id: 'precedence', label: '运算符优先级', method: '逐步化简' },
    { id: 'function', label: '常见函数', method: 'print · len · abs …' }
  ],
  theory() {
    return `<div class="rule"><b>数据类型</b>：Python 内置 int（整数）、float（浮点数）、str（字符串）、bool（布尔值）、list（列表）、dict（字典）、tuple（元组）、set（集合）八种常用类型；用 <span class="mono">type(x)</span> 可以判断变量的类型。</div>
    <div class="rule"><b>运算符</b>：算术（+ - * / // % **）、比较（== != &gt; &lt; &gt;= &lt;=）、逻辑（and or not）、赋值（= += -= *= /=）、成员（in not in）、身份（is is not）六类。</div>
    <div class="rule"><b>优先级</b>（从高到低）：<span class="mono">**</span> → <span class="mono">* / // %</span> → <span class="mono">+ -</span> → 比较 → <span class="mono">not</span> → <span class="mono">and</span> → <span class="mono">or</span> → 赋值；圆括号内的运算最先执行。</div>
    <div class="rule"><b>内置函数</b>：print 输出、type 查类型、len 求长度、int / float / str 类型转换、input 读取输入、range 生成序列、abs 绝对值、round 四舍五入、max / min 取最值、sum 求和。</div>`;
  },
  settingsHTML() {
    const sub = this.currentSub;
    if (sub === 'datatype') {
      const cur = $('#pe-type') ? $('#pe-type').value : (this._type || 'all');
      return `<div class="inline-setting"><div class="setting"><label>数据类型</label>
        <select id="pe-type">
          <option value="all" ${cur === 'all' ? 'selected' : ''}>全部浏览（8 种）</option>
          ${this.TYPES.map(t => `<option value="${t.id}" ${cur === t.id ? 'selected' : ''}>${t.name.split('（')[0]}</option>`).join('')}
        </select></div></div>`;
    }
    if (sub === 'operator') {
      const cur = $('#pe-op') ? $('#pe-op').value : (this._op || 'all');
      return `<div class="inline-setting"><div class="setting"><label>运算符类别</label>
        <select id="pe-op">
          <option value="all" ${cur === 'all' ? 'selected' : ''}>全部浏览（6 类）</option>
          ${this.OP_CATS.map(c => `<option value="${c.id}" ${cur === c.id ? 'selected' : ''}>${c.name}</option>`).join('')}
        </select></div></div>`;
    }
    if (sub === 'function') {
      const cur = $('#pe-fn') ? $('#pe-fn').value : (this._fn || 'all');
      return `<div class="inline-setting"><div class="setting"><label>内置函数</label>
        <select id="pe-fn">
          <option value="all" ${cur === 'all' ? 'selected' : ''}>全部浏览（13 个）</option>
          ${this.FUNCS.map(f => `<option value="${f.id}" ${cur === f.id ? 'selected' : ''}>${f.name}</option>`).join('')}
        </select></div></div>`;
    }
    if (sub === 'precedence') {
      const curP = $('#pe-preset') ? $('#pe-preset').value : (this._preset || 'p1');
      const curE = $('#pe-expr') ? $('#pe-expr').value : (this._customExpr || '');
      return `<div class="inline-setting"><div class="setting"><label>示例表达式</label>
        <select id="pe-preset">
          <option value="p1" ${curP === 'p1' ? 'selected' : ''}>综合：2 + 3 * 4 &gt; 10 and not False</option>
          <option value="p2" ${curP === 'p2' ? 'selected' : ''}>逻辑：not True and False</option>
          <option value="p3" ${curP === 'p3' ? 'selected' : ''}>算术：2 + 3 * 4</option>
          <option value="p4" ${curP === 'p4' ? 'selected' : ''}>比较：3 * 4 &gt; 10 or 2 &lt; 5</option>
          <option value="custom" ${curP === 'custom' ? 'selected' : ''}>自定义输入</option>
        </select></div></div>
        <div class="inline-setting"><div class="setting"><label>自定义表达式</label>
        <input type="text" id="pe-expr" value="${esc(curE)}" placeholder="如 2 + 3 * 4 > 10" style="width:clamp(150px,24vw,220px)"></div></div>`;
    }
    return '';
  },
  afterSettings() {
    const bind = id => { const el = $('#' + id); if (el) el.onchange = () => Engine.reload(); };
    ['pe-type', 'pe-op', 'pe-fn', 'pe-preset', 'pe-expr'].forEach(bind);
  },

  /* ==================== 数据定义 ==================== */
  TYPES: [
    { id: 'int', name: 'int（整数）', sample: '42', code1: `x = <span class="num">42</span>`, typeStr: "<class 'int'>",
      desc: '表示整数，可正可负，没有小数部分，如 0、-7、2026。' },
    { id: 'float', name: 'float（浮点数）', sample: '3.14', code1: `x = <span class="num">3.14</span>`, typeStr: "<class 'float'>",
      desc: '表示带小数部分的实数，用于小数运算，如 2.0、-0.5。' },
    { id: 'str', name: 'str（字符串）', sample: '"hello"', code1: `x = <span class="str">"hello"</span>`, typeStr: "<class 'str'>",
      desc: '由引号括起来的文本，单引号、双引号均可，用于表示文字内容。' },
    { id: 'bool', name: 'bool（布尔值）', sample: 'True', code1: `x = <span class="kw">True</span>`, typeStr: "<class 'bool'>",
      desc: '只有 True（真）和 False（假）两个取值，常用于条件判断。' },
    { id: 'list', name: 'list（列表）', sample: '[1, 2, 3]', code1: `x = [<span class="num">1</span>, <span class="num">2</span>, <span class="num">3</span>]`, typeStr: "<class 'list'>",
      desc: '有序、可变的序列，用方括号表示，元素之间用逗号分隔。' },
    { id: 'dict', name: 'dict（字典）', sample: '{"name": "Tom", "age": 18}', code1: `x = {<span class="str">"name"</span>: <span class="str">"Tom"</span>, <span class="str">"age"</span>: <span class="num">18</span>}`, typeStr: "<class 'dict'>",
      desc: '键值对集合，用花括号表示，通过“键”来访问对应的“值”。' },
    { id: 'tuple', name: 'tuple（元组）', sample: '(1, 2, 3)', code1: `x = (<span class="num">1</span>, <span class="num">2</span>, <span class="num">3</span>)`, typeStr: "<class 'tuple'>",
      desc: '有序、不可变的序列，用圆括号表示，创建之后不能修改元素。' },
    { id: 'set', name: 'set（集合）', sample: '{1, 2, 3}', code1: `x = {<span class="num">1</span>, <span class="num">2</span>, <span class="num">3</span>}`, typeStr: "<class 'set'>",
      desc: '无序、不重复的元素集合，用花括号表示，会自动去除重复项。' }
  ],
  OP_CATS: [
    { id: 'arith', name: '算术运算符', method: '+ - * / // % **',
      rows: [
        ['+', '加法', '3 + 4', '7'],
        ['-', '减法', '9 - 5', '4'],
        ['*', '乘法', '3 * 4', '12'],
        ['/', '除法（结果为小数）', '7 / 2', '3.5'],
        ['//', '整除（取商的整数部分）', '5 // 2', '2'],
        ['%', '取余（余数）', '7 % 3', '1'],
        ['**', '幂运算（乘方）', '2 ** 3', '8']
      ],
      code: [
        `<span class="fn">print</span>(<span class="num">3</span> + <span class="num">4</span>)`,
        `<span class="fn">print</span>(<span class="num">7</span> / <span class="num">2</span>)`,
        `<span class="fn">print</span>(<span class="num">5</span> // <span class="num">2</span>)`,
        `<span class="fn">print</span>(<span class="num">2</span> ** <span class="num">3</span>)`
      ],
      out: ['7', '3.5', '2', '8'],
      note: '算术运算符用于数值计算：/ 除法结果一定是小数；// 整除只保留商的整数部分；% 取余数；** 表示乘方（a 的 b 次方）。' },
    { id: 'cmp', name: '比较运算符', method: '== != > < >= <=',
      rows: [
        ['==', '等于', '3 == 3', 'True'],
        ['!=', '不等于', '3 != 5', 'True'],
        ['>', '大于', '5 > 3', 'True'],
        ['<', '小于', '3 < 5', 'True'],
        ['>=', '大于等于', '4 >= 4', 'True'],
        ['<=', '小于等于', '5 <= 3', 'False']
      ],
      code: [
        `<span class="fn">print</span>(<span class="num">3</span> == <span class="num">3</span>)`,
        `<span class="fn">print</span>(<span class="num">5</span> &lt;= <span class="num">3</span>)`
      ],
      out: ['True', 'False'],
      note: '比较运算符比较两边的值是否成立，成立结果为 True，不成立为 False；注意 == 是判断相等，单个 = 是赋值。' },
    { id: 'logic', name: '逻辑运算符', method: 'and · or · not',
      rows: [
        ['and', '与：两边都为 True 才为 True', 'True and False', 'False'],
        ['or', '或：一边为 True 即为 True', 'True or False', 'True'],
        ['not', '非：把结果取反', 'not True', 'False']
      ],
      code: [
        `<span class="fn">print</span>(<span class="kw">True</span> <span class="kw">and</span> <span class="kw">False</span>)`,
        `<span class="fn">print</span>(<span class="kw">True</span> <span class="kw">or</span> <span class="kw">False</span>)`,
        `<span class="fn">print</span>(<span class="kw">not</span> <span class="kw">True</span>)`
      ],
      out: ['False', 'True', 'False'],
      note: 'and 两边全真才真；or 两边一真即真；not 将 True 变为 False、False 变为 True。' },
    { id: 'assign', name: '赋值运算符', method: '= · += · -= …',
      rows: [
        ['=', '直接赋值', 'x = 5', 'x = 5'],
        ['+=', '加后赋值', 'x += 3', 'x = 8'],
        ['-=', '减后赋值', 'x -= 2', 'x = 6'],
        ['*=', '乘后赋值', 'x *= 4', 'x = 24'],
        ['/=', '除后赋值', 'x /= 2', 'x = 12.0']
      ],
      code: [
        `x = <span class="num">5</span>`,
        `x += <span class="num">3</span>  <span class="cm"># x = 5 + 3 = 8</span>`,
        `x *= <span class="num">2</span>  <span class="cm"># x = 8 × 2 = 16</span>`,
        `<span class="fn">print</span>(x)  <span class="cm"># 16</span>`
      ],
      out: ['16'],
      note: '= 把右边的值赋给变量；+=、-=、*=、/= 是“先运算再赋值”的简写，例如 x += 3 等价于 x = x + 3。' },
    { id: 'member', name: '成员运算符', method: 'in · not in',
      rows: [
        ['in', '在序列中则为 True', '3 in [1, 2, 3]', 'True'],
        ['not in', '不在序列中则为 True', '4 not in [1, 2, 3]', 'True']
      ],
      code: [
        `<span class="fn">print</span>(<span class="num">3</span> <span class="kw">in</span> [<span class="num">1</span>, <span class="num">2</span>, <span class="num">3</span>])`,
        `<span class="fn">print</span>(<span class="num">4</span> <span class="kw">not in</span> [<span class="num">1</span>, <span class="num">2</span>, <span class="num">3</span>])`
      ],
      out: ['True', 'True'],
      note: 'in 判断某个元素是否在列表、字符串等序列中；not in 判断是否不在序列中。' },
    { id: 'identity', name: '身份运算符', method: 'is · is not',
      rows: [
        ['is', '是否为同一对象', 'a is b', 'True'],
        ['is not', '是否不是同一对象', 'a is c', 'False']
      ],
      code: [
        `a = [<span class="num">1</span>, <span class="num">2</span>]`,
        `b = a            <span class="cm"># b 引用同一对象</span>`,
        `c = [<span class="num">1</span>, <span class="num">2</span>]  <span class="cm"># 内容相同，新对象</span>`,
        `<span class="fn">print</span>(a <span class="kw">is</span> b)`,
        `<span class="fn">print</span>(a <span class="kw">is</span> c)`
      ],
      out: ['True', 'False'],
      note: 'is 判断两个变量是否指向同一个对象（身份相同）；内容相同的两个新对象并不相同，判断值是否相等要用 ==。' }
  ],
  FUNCS: [
    { id: 'print', name: 'print(x)', example: '输出文字与多个值',
      purpose: '把内容输出到控制台；多个参数用逗号分隔，输出时以空格隔开。',
      code: [`<span class="fn">print</span>(<span class="str">"Hello"</span>, <span class="str">"Python"</span>)`],
      out: ['Hello Python'] },
    { id: 'type', name: 'type(x)', example: '查看 3.14 的类型',
      purpose: '返回参数的数据类型，常用于判断变量是什么类型。',
      code: [`<span class="fn">print</span>(<span class="fn">type</span>(<span class="num">3.14</span>))`],
      out: ["<class 'float'>"] },
    { id: 'len', name: 'len(x)', example: '字符串长度',
      purpose: '返回序列的长度：字符串字符数、列表元素个数等。',
      code: [`<span class="fn">print</span>(<span class="fn">len</span>(<span class="str">"hello"</span>))`],
      out: ['5'] },
    { id: 'int', name: 'int(x)', example: '转换为整数',
      purpose: '把参数转换为整数；小数直接去掉小数部分，字符串数字可转为整数。',
      code: [`<span class="fn">print</span>(<span class="fn">int</span>(<span class="num">3.9</span>), <span class="fn">int</span>(<span class="str">"18"</span>))`],
      out: ['3 18'] },
    { id: 'float', name: 'float(x)', example: '转换为浮点数',
      purpose: '把参数转换为浮点数（小数形式）。',
      code: [`<span class="fn">print</span>(<span class="fn">float</span>(<span class="num">5</span>))`],
      out: ['5.0'] },
    { id: 'str', name: 'str(x)', example: '转字符串后拼接',
      purpose: '把参数转换为字符串，便于和文字进行拼接。',
      code: [`<span class="fn">print</span>(<span class="fn">str</span>(<span class="num">100</span>) + <span class="str">"分"</span>)`],
      out: ['100分'] },
    { id: 'input', name: 'input(提示)', example: '读取键盘输入（模拟输入“小明”）',
      purpose: '显示提示文字并读取用户从键盘输入的内容；本演示模拟输入姓名“小明”。',
      code: [
        `name = <span class="fn">input</span>(<span class="str">"请输入姓名："</span>)`,
        `<span class="fn">print</span>(<span class="str">"你好，"</span> + name)`
      ],
      out: ['请输入姓名：小明', '你好，小明'] },
    { id: 'range', name: 'range(起, 止)', example: '1 到 5 的序列',
      purpose: '生成从“起”到“止-1”的整数序列（含头不含尾），常用于 for 循环。',
      code: [`<span class="fn">print</span>(<span class="fn">list</span>(<span class="fn">range</span>(<span class="num">1</span>, <span class="num">6</span>)))`],
      out: ['[1, 2, 3, 4, 5]'] },
    { id: 'abs', name: 'abs(x)', example: '绝对值',
      purpose: '返回参数的绝对值（负数取正）。',
      code: [`<span class="fn">print</span>(<span class="fn">abs</span>(-<span class="num">3</span>), <span class="fn">abs</span>(<span class="num">4.5</span>))`],
      out: ['3 4.5'] },
    { id: 'round', name: 'round(x, n)', example: '保留两位小数',
      purpose: '把 x 四舍五入，保留 n 位小数。',
      code: [`<span class="fn">print</span>(<span class="fn">round</span>(<span class="num">3.14159</span>, <span class="num">2</span>))`],
      out: ['3.14'] },
    { id: 'max', name: 'max(…)', example: '取最大值',
      purpose: '返回所有参数中的最大值。',
      code: [`<span class="fn">print</span>(<span class="fn">max</span>(<span class="num">1</span>, <span class="num">5</span>, <span class="num">3</span>))`],
      out: ['5'] },
    { id: 'min', name: 'min(…)', example: '取最小值',
      purpose: '返回所有参数中的最小值。',
      code: [`<span class="fn">print</span>(<span class="fn">min</span>(<span class="num">4</span>, <span class="num">2</span>, <span class="num">9</span>))`],
      out: ['2'] },
    { id: 'sum', name: 'sum(序列)', example: '列表求和',
      purpose: '返回序列中所有元素相加的和。',
      code: [`<span class="fn">print</span>(<span class="fn">sum</span>([<span class="num">1</span>, <span class="num">2</span>, <span class="num">3</span>, <span class="num">4</span>]))`],
      out: ['10'] }
  ],
  PRESETS: {
    p1: { expr: '2 + 3 * 4 > 10 and not False', note: '算术 → 比较 → not → and，综合演练' },
    p2: { expr: 'not True and False', note: 'not 优先级高于 and' },
    p3: { expr: '2 + 3 * 4', note: '先乘除后加减' },
    p4: { expr: '3 * 4 > 10 or 2 < 5', note: '比较运算先于 or' }
  },
  PREC_ROWS: [
    { level: 1, ops: '**', desc: '幂运算（乘方）' },
    { level: 2, ops: '* / // %', desc: '乘、除、整除、取余' },
    { level: 3, ops: '+ -', desc: '加、减' },
    { level: 4, ops: '== != > < >= <=', desc: '比较运算（结果为 True / False）' },
    { level: 5, ops: 'not', desc: '逻辑非' },
    { level: 6, ops: 'and', desc: '逻辑与' },
    { level: 7, ops: 'or', desc: '逻辑或' },
    { level: 8, ops: '= += -= *= /= …', desc: '赋值运算（最低，仅供参考）' }
  ],

  /* ==================== 通用画布辅助 ==================== */
  frame(label, state, title) {
    return `<div class="stage-hd"><span>${label}</span><span class="state">${state}</span></div>
      <div class="step-title"><span class="st-dot"></span>${title}</div>`;
  },
  codeBlock(lines, curIdx) {
    curIdx = curIdx == null ? -1 : curIdx;
    return `<div class="code-block">` + lines.map((l, i) =>
      `<div class="l ${i === curIdx ? 'cur' : ''}">${l}</div>`).join('') + `</div>`;
  },
  consoleOut(lines) {
    return `<div class="console">` + lines.map(l => `<div class="cl">${esc(l)}</div>`).join('') + `</div>`;
  },

  /* ==================== 子页面一：数据类型介绍 ==================== */
  typeGrid(activeIdx, focus, focusState) {
    return `<div class="py-type-grid">` + this.TYPES.map((t, i) => {
      let cls = 'dim';
      if (focus) { if (t.id === focus) cls = focusState || 'cur'; }
      else if (i < activeIdx) cls = 'done';
      else if (i === activeIdx) cls = 'cur';
      return `<div class="py-tcard ${cls} write-in" style="animation-delay:${(i * 0.04).toFixed(2)}s">
        <div class="py-tname">${t.name}</div>
        <div class="py-tval">${esc(t.sample)}</div>
        <div class="py-ttype">${esc(t.typeStr)}</div>
        <div class="py-tdesc">${t.desc}</div>
      </div>`;
    }).join('') + `</div>`;
  },
  genDatatype() {
    const selEl = $('#pe-type');
    const sel = selEl ? selEl.value : (this._type || 'all');
    const valid = this.TYPES.some(t => t.id === sel) ? sel : 'all';
    this._type = valid;
    const focus = valid === 'all' ? null : valid;
    const list = valid === 'all' ? this.TYPES : this.TYPES.filter(t => t.id === valid);
    const steps = [];

    steps.push({
      desc: `Python 内置了 <b class="c-op">8 种</b>常见数据类型：整数、浮点数、字符串、布尔值、列表、字典、元组、集合。下一步逐个查看每种类型的示例值与 <span class="mono">type()</span> 判断结果。`,
      draw: () => {
        $('#stage').innerHTML = this.frame('数据类型 · 总览', '准备', '8 种内置数据类型')
          + this.typeGrid(-1, focus, 'cur')
          + `<div class="annot" style="margin-top:10px">使用 <span class="mono">type(x)</span> 可以返回变量 x 的数据类型，是判断变量类型最常用的方法。</div>`;
      }
    });

    list.forEach((t, k) => {
      const allIdx = this.TYPES.indexOf(t);
      steps.push({
        desc: `查看 <b class="c-op">${t.name}</b>：示例值 <span class="mono">${esc(t.sample)}</span>，执行 <span class="mono">type(x)</span> 得到 <b>${esc(t.typeStr)}</b>。${t.desc}`,
        draw: () => {
          $('#stage').innerHTML = this.frame('数据类型 · ' + t.name, `第 ${k + 1}/${list.length} 步`, t.name)
            + this.typeGrid(allIdx, focus, focus ? 'done' : null)
            + `<div class="two-col" style="margin-top:8px">
                <div>${this.codeBlock([t.code1, `<span class="fn">print</span>(<span class="fn">type</span>(x))`], 1)}</div>
                <div>${this.consoleOut([t.typeStr])}</div>
              </div>
              <div class="annot" style="margin-top:8px">${t.desc}</div>`;
        }
      });
    });
    return steps;
  },

  /* ==================== 子页面二：运算符介绍 ==================== */
  opGrid(activeIdx, focus, focusState) {
    return `<div class="py-cat-grid">` + this.OP_CATS.map((c, i) => {
      let cls = 'dim';
      if (focus) { if (c.id === focus) cls = focusState || 'cur'; }
      else if (i < activeIdx) cls = 'done';
      else if (i === activeIdx) cls = 'cur';
      return `<div class="py-cat-card ${cls} write-in" style="animation-delay:${(i * 0.04).toFixed(2)}s">
        <div class="py-cat-name">${c.name}</div>
        <div class="py-cat-method">${esc(c.method)}</div>
      </div>`;
    }).join('') + `</div>`;
  },
  genOperator() {
    const selEl = $('#pe-op');
    const sel = selEl ? selEl.value : (this._op || 'all');
    const valid = this.OP_CATS.some(c => c.id === sel) ? sel : 'all';
    this._op = valid;
    const focus = valid === 'all' ? null : valid;
    const list = valid === 'all' ? this.OP_CATS : this.OP_CATS.filter(c => c.id === valid);
    const steps = [];

    steps.push({
      desc: `Python 运算符分为 <b class="c-op">六类</b>：算术、比较、逻辑、赋值、成员、身份。下一步逐类查看运算符表格、示例代码与运行结果。`,
      draw: () => {
        $('#stage').innerHTML = this.frame('运算符 · 总览', '准备', '六类运算符')
          + this.opGrid(-1, focus, 'cur')
          + `<div class="annot" style="margin-top:10px">每一类运算符都配有运算符表格、示例代码与控制台运行结果，便于对照记忆。</div>`;
      }
    });

    list.forEach((cat, k) => {
      const allIdx = this.OP_CATS.indexOf(cat);
      steps.push({
        desc: `<b class="c-op">${cat.name}</b>：${cat.note}`,
        draw: () => {
          $('#stage').innerHTML = this.frame('运算符 · ' + cat.name, `第 ${k + 1}/${list.length} 步`, esc(cat.method))
            + this.opGrid(allIdx, focus, focus ? 'done' : null)
            + `<div class="tbl-wrap" style="margin-top:8px"><table class="tbl">
                <colgroup><col style="width:18%"><col style="width:34%"><col style="width:24%"><col style="width:24%"></colgroup>
                <thead><tr><th>运算符</th><th>含义</th><th>示例</th><th>结果</th></tr></thead>
                <tbody>${cat.rows.map(r =>
                  `<tr><td><b class="py-op">${esc(r[0])}</b></td><td class="l">${r[1]}</td><td class="l" style="font-family:var(--mono)">${esc(r[2])}</td><td><b>${esc(r[3])}</b></td></tr>`).join('')}</tbody>
              </table></div>
              <div class="two-col" style="margin-top:8px">
                <div>${this.codeBlock(cat.code, -1)}</div>
                <div>${this.consoleOut(cat.out)}</div>
              </div>
              <div class="annot" style="margin-top:8px">${cat.note}</div>`;
        }
      });
    });
    return steps;
  },

  /* ==================== 子页面三：运算符优先级介绍 ==================== */
  _rank(op) {
    const map = { '**': 8, '*': 7, '/': 7, '//': 7, '%': 7, '+': 6, '-': 6, '>': 5, '<': 5, '>=': 5, '<=': 5, '==': 5, '!=': 5, 'not': 4, 'and': 3, 'or': 2 };
    return map[op] || 0;
  },
  _needsParen(child, parent) {
    if (!child || child.type !== 'bin') return false;
    const cr = this._rank(child.op), pr = this._rank(parent.op);
    if (cr < pr) return true;
    if (cr === pr && parent.op === '**' && parent.right === child) return true;
    return false;
  },
  _fmtVal(v) {
    if (typeof v === 'boolean') return v ? 'True' : 'False';
    if (typeof v === 'number') {
      if (Number.isInteger(v)) return String(v);
      return String(Math.round(v * 100000) / 100000);
    }
    return String(v);
  },
  _exprSrc(node) {
    if (node.type === 'lit') return this._fmtVal(node.v);
    if (node.type === 'unary') return `${node.op} ${this._exprSrc(node.operand)}`;
    const lp = this._needsParen(node.left, node) ? '(' : '';
    const lq = lp ? ')' : '';
    const rp = this._needsParen(node.right, node) ? '(' : '';
    const rq = rp ? ')' : '';
    return `${lp}${this._exprSrc(node.left)}${lq} ${node.op} ${rp}${this._exprSrc(node.right)}${rq}`;
  },
  /* 把已完成子节点替换为结果后的当前子表达式文本（用于化简行与说明） */
  _subExpr(node, doneSet) {
    if (node.type === 'lit') return this._fmtVal(node.v);
    if (doneSet.has(node)) return this._fmtVal(node.value);
    if (node.type === 'unary') return `${node.op} ${this._subExpr(node.operand, doneSet)}`;
    const lp = this._needsParen(node.left, node) ? '(' : '';
    const lq = lp ? ')' : '';
    const rp = this._needsParen(node.right, node) ? '(' : '';
    const rq = rp ? ')' : '';
    return `${lp}${this._subExpr(node.left, doneSet)}${lq} ${node.op} ${rp}${this._subExpr(node.right, doneSet)}${rq}`;
  },
  _renderExpr(node, doneSet, curNode) {
    if (node.type === 'lit') return `<span class="expr-lit">${this._fmtVal(node.v)}</span>`;
    if (node === curNode) return `<span class="expr-cur">${esc(this._exprSrc(node))}</span>`;
    if (doneSet.has(node)) return `<span class="expr-done">${this._fmtVal(node.value)}</span>`;
    if (node.type === 'unary') {
      return `<span class="expr-op">${esc(node.op)}</span> ${this._renderExpr(node.operand, doneSet, curNode)}`;
    }
    const lp = this._needsParen(node.left, node) ? '(' : '';
    const lq = lp ? ')' : '';
    const rp = this._needsParen(node.right, node) ? '(' : '';
    const rq = rp ? ')' : '';
    return `${lp}${this._renderExpr(node.left, doneSet, curNode)}${lq} <span class="expr-op">${esc(node.op)}</span> ${rp}${this._renderExpr(node.right, doneSet, curNode)}${rq}`;
  },
  _binNote(op) {
    const map = {
      '**': '幂运算优先级最高，先算 **（乘方）',
      '*': '乘法优先级高于加法，先算 *',
      '/': '除法优先级高于加法，先算 /',
      '//': '整除与乘法同级，先算 //',
      '%': '取余与乘法同级，先算 %',
      '+': '乘法算完后，再算 +（加法）',
      '-': '乘法算完后，再算 -（减法）',
      '==': '比较运算优先级低于加减，先判断 ==（等于）',
      '!=': '比较运算，先判断 !=（不等于）',
      '>': '比较运算，先判断 >（大于）',
      '<': '比较运算，先判断 <（小于）',
      '>=': '比较运算，先判断 >=（大于等于）',
      '<=': '比较运算，先判断 <=（小于等于）',
      'and': '比较与 not 算完后，最后算逻辑 and（两边都真才真）',
      'or': 'and 算完后，最后算逻辑 or（一边为真即为真）'
    };
    return map[op] || '按优先级完成本步运算';
  },
  _unaryNote(op) {
    if (op === 'not') return '比较运算之后，先算 not（取反）';
    if (op === '-') return '负号优先，先取负号';
    return '按优先级完成本步运算';
  },
  _tokenize(src) {
    const s = String(src).trim();
    const toks = [];
    let i = 0;
    while (i < s.length) {
      const c = s[i];
      if (/\s/.test(c)) { i++; continue; }
      let m = s.slice(i).match(/^\d+(\.\d+)?/);
      if (m) { toks.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
      m = s.slice(i).match(/^[A-Za-z_][A-Za-z0-9_]*/);
      if (m) {
        const w = m[0];
        if (w === 'and') toks.push({ t: 'kw', v: 'and' });
        else if (w === 'or') toks.push({ t: 'kw', v: 'or' });
        else if (w === 'not') toks.push({ t: 'kw', v: 'not' });
        else if (w === 'True') toks.push({ t: 'bool', v: true });
        else if (w === 'False') toks.push({ t: 'bool', v: false });
        else throw new Error('无法识别的标识符：' + w);
        i += w.length; continue;
      }
      const two = s.slice(i, i + 2);
      if (two === '**' || two === '//' || two === '==' || two === '!=' || two === '>=' || two === '<=') {
        toks.push({ t: 'op', v: two }); i += 2; continue;
      }
      if ('+-*/%<>()'.includes(c)) { toks.push({ t: 'op', v: c }); i++; continue; }
      throw new Error('无法识别的字符：' + c);
    }
    if (!toks.length) throw new Error('表达式为空，请输入一个简单表达式');
    return toks;
  },
  _parse(src) {
    const toks = this._tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const next = () => toks[p++];

    const parseOr = () => {
      let node = parseAnd();
      while (peek() && peek().t === 'kw' && peek().v === 'or') { next(); const r = parseAnd(); node = { type: 'bin', op: 'or', left: node, right: r }; }
      return node;
    };
    const parseAnd = () => {
      let node = parseNot();
      while (peek() && peek().t === 'kw' && peek().v === 'and') { next(); const r = parseNot(); node = { type: 'bin', op: 'and', left: node, right: r }; }
      return node;
    };
    const parseNot = () => {
      if (peek() && peek().t === 'kw' && peek().v === 'not') { next(); const o = parseNot(); return { type: 'unary', op: 'not', operand: o }; }
      return parseCmp();
    };
    const parseCmp = () => {
      let node = parseAdd();
      if (peek() && peek().t === 'op' && ['==', '!=', '>', '<', '>=', '<='].includes(peek().v)) {
        const op = next().v; const r = parseAdd(); node = { type: 'bin', op, left: node, right: r };
      }
      return node;
    };
    const parseAdd = () => {
      let node = parseMul();
      while (peek() && peek().t === 'op' && (peek().v === '+' || peek().v === '-')) {
        const op = next().v; const r = parseMul(); node = { type: 'bin', op, left: node, right: r };
      }
      return node;
    };
    const parseMul = () => {
      let node = parsePow();
      while (peek() && peek().t === 'op' && ['*', '/', '//', '%'].includes(peek().v)) {
        const op = next().v; const r = parsePow(); node = { type: 'bin', op, left: node, right: r };
      }
      return node;
    };
    const parsePow = () => {
      const node = parseUnary();
      if (peek() && peek().t === 'op' && peek().v === '**') { next(); const r = parsePow(); return { type: 'bin', op: '**', left: node, right: r }; }
      return node;
    };
    const parseUnary = () => {
      if (peek() && peek().t === 'op' && peek().v === '-') { next(); const o = parsePow(); return { type: 'unary', op: '-', operand: o }; }
      return parsePrimary();
    };
    const parsePrimary = () => {
      const tk = next();
      if (!tk) throw new Error('表达式不完整');
      if (tk.t === 'num') return { type: 'lit', v: tk.v };
      if (tk.t === 'bool') return { type: 'lit', v: tk.v };
      if (tk.t === 'op' && tk.v === '(') {
        const node = parseOr();
        const close = next();
        if (!close || close.t !== 'op' || close.v !== ')') throw new Error('括号不匹配');
        return node;
      }
      throw new Error('表达式格式不正确');
    };

    const root = parseOr();
    if (p < toks.length) throw new Error('表达式格式不正确');
    return root;
  },
  _eval(node, trace) {
    if (node.type === 'lit') return node.v;
    if (node.type === 'unary') {
      const v = this._eval(node.operand, trace);
      const r = node.op === 'not' ? !v : -v;
      node.value = r;
      trace.push({ node, kind: node.type, op: node.op, result: r });
      return r;
    }
    const lv = this._eval(node.left, trace);
    const rv = this._eval(node.right, trace);
    let r;
    switch (node.op) {
      case '+': r = lv + rv; break;
      case '-': r = lv - rv; break;
      case '*': r = lv * rv; break;
      case '/': r = lv / rv; break;
      case '//': r = Math.floor(lv / rv); break;
      case '%': r = lv % rv; break;
      case '**': r = Math.pow(lv, rv); break;
      case '==': r = lv === rv; break;
      case '!=': r = lv !== rv; break;
      case '>': r = lv > rv; break;
      case '<': r = lv < rv; break;
      case '>=': r = lv >= rv; break;
      case '<=': r = lv <= rv; break;
      case 'and': r = lv && rv; break;
      case 'or': r = lv || rv; break;
      default: throw new Error('不支持的运算符：' + node.op);
    }
    node.value = r;
    trace.push({ node, kind: node.type, op: node.op, result: r });
    return r;
  },
  genPrecedence() {
    const presetEl = $('#pe-preset');
    let preset = presetEl ? presetEl.value : (this._preset || 'p1');
    if (preset !== 'custom' && !this.PRESETS[preset]) preset = 'p1';
    this._preset = preset;

    let src;
    if (preset === 'custom') {
      const e = $('#pe-expr');
      src = (e ? e.value : this._customExpr || '').trim();
      this._customExpr = src;
      if (!src) throw new Error('请在输入框中填写一个简单表达式（如 2 + 3 * 4），或改回预设示例。');
    } else {
      src = this.PRESETS[preset].expr;
    }

    const root = this._parse(src);
    const trace = [];
    const final = this._eval(root, trace);

    const steps = [];

    /* 第 1 步：优先级表 */
    steps.push({
      desc: `Python 运算符<b class="c-op">优先级从高到低</b>为：<span class="mono">**</span> → <span class="mono">* / // %</span> → <span class="mono">+ -</span> → 比较 → <span class="mono">not</span> → <span class="mono">and</span> → <span class="mono">or</span> → 赋值；同一级运算符从左到右计算。`,
      draw: () => {
        $('#stage').innerHTML = this.frame('运算符优先级', '知识梳理', '优先级表（从高到低）')
          + `<div class="tbl-wrap"><table class="tbl py-prec">
              <colgroup><col style="width:14%"><col style="width:36%"><col style="width:50%"></colgroup>
              <thead><tr><th>级别</th><th>运算符</th><th>说明</th></tr></thead>
              <tbody>${this.PREC_ROWS.map(r =>
                `<tr><td><span class="chip ${r.level <= 3 ? 'cy' : r.level <= 5 ? 'o' : 'dim'}" style="font-size:clamp(10px,1.3vw,12px)">L${r.level}</span></td>
                 <td class="op">${esc(r.ops)}</td><td class="l">${r.desc}</td></tr>`).join('')}
              </tbody>
            </table></div>
            <div class="annot" style="margin-top:8px">注：圆括号 <span class="mono">()</span> 优先级最高；赋值运算符只用于修改变量，不参与表达式求值。</div>`;
      }
    });

    /* 第 2 步：待化简表达式 */
    steps.push({
      desc: `待化简表达式为 <b class="c-op">${esc(src)}</b>${preset === 'custom' ? '' : '（' + this.PRESETS[preset].note + '）'}；接下来从<b>最高优先级</b>的子表达式开始，一步步把结果代回原表达式。`,
      draw: () => {
        $('#stage').innerHTML = this.frame('优先级化简', '准备', '原始表达式')
          + this.codeBlock([`>>> ${esc(src)}`], -1)
          + `<div class="console"><div class="cl" style="color:#94A3B8"># 按优先级从高到低逐步化简…</div></div>`
          + `<div class="annot" style="margin-top:8px">化简原则：先算优先级最高的子表达式，把结果代回原表达式，再算下一级，直到剩下最终结果。</div>`;
      }
    });

    /* 逐步化简 */
    trace.forEach((t, i) => {
      const prev = trace.slice(0, i);
      const doneSet = new Set(prev.map(x => x.node));
      const subExpr = this._subExpr(t.node, doneSet);
      const note = t.kind === 'unary' ? this._unaryNote(t.op) : this._binNote(t.op);
      steps.push({
        desc: `第 ${i + 1} 步：${esc(note)}，即 <b class="c-op">${esc(subExpr)} = ${this._fmtVal(t.result)}</b>，把结果代回原表达式。`,
        draw: () => {
          const afterSet = new Set(doneSet);
          afterSet.add(t.node);
          const before = this._renderExpr(root, doneSet, t.node);
          const after = this._renderExpr(root, afterSet, null);
          const doneRows = prev.map((x, j) => {
            const ds = new Set(trace.slice(0, j).map(y => y.node));
            return `<div class="cp-row"><span class="w" style="background:var(--green-soft);color:#15803D">${esc(this._subExpr(x.node, ds))}</span><span class="eq">=</span><span class="sum" style="background:var(--green-soft);color:#15803D;border-color:var(--green)">${this._fmtVal(x.result)}</span></div>`;
          }).join('');
          $('#stage').innerHTML = this.frame('优先级化简', `第 ${i + 1}/${trace.length} 步`, '逐步替换结果')
            + this.codeBlock([`>>> ${esc(src)}`], -1)
            + `<div class="calc-process">
                <div class="cp-label">化简过程（绿色 = 已完成，橙色 = 当前计算）</div>
                ${doneRows}
                <div class="cp-row"><span class="w active">${esc(subExpr)}</span><span class="eq">=</span><span class="sum">${this._fmtVal(t.result)}</span></div>
              </div>
              <div class="py-chain">
                <div class="py-chain-line"><span class="py-chain-hd">当前表达式：</span>${before}</div>
                <div class="py-chain-line"><span class="py-chain-hd">替换后：</span>${after}</div>
              </div>`;
        }
      });
    });

    /* 最终结果 */
    steps.push({
      desc: `表达式化简完成，最终结果为 <b class="c-op">${this._fmtVal(final)}</b>。`,
      draw: () => {
        $('#stage').innerHTML = this.frame('优先级化简', '完成', '最终结果')
          + this.codeBlock([`>>> ${esc(src)}`], -1)
          + this.consoleOut([this._fmtVal(final)])
          + `<div class="result-box">${esc(src)} = ${this._fmtVal(final)}</div>`;
      }
    });

    return steps;
  },

  /* ==================== 子页面四：常见函数介绍 ==================== */
  fnRack(activeIdx, focus, focusState) {
    return `<div class="py-fn-rack">` + this.FUNCS.map((f, i) => {
      let cls = 'dim';
      if (focus) { if (f.id === focus) cls = focusState || 'cur'; }
      else if (i < activeIdx) cls = 'done';
      else if (i === activeIdx) cls = 'cur';
      return `<span class="py-fn-chip ${cls} write-in" style="animation-delay:${(i * 0.03).toFixed(2)}s">${f.name}</span>`;
    }).join('') + `</div>`;
  },
  genFunction() {
    const selEl = $('#pe-fn');
    const sel = selEl ? selEl.value : (this._fn || 'all');
    const valid = this.FUNCS.some(f => f.id === sel) ? sel : 'all';
    this._fn = valid;
    const focus = valid === 'all' ? null : valid;
    const list = valid === 'all' ? this.FUNCS : this.FUNCS.filter(f => f.id === valid);
    const steps = [];

    steps.push({
      desc: `Python 常用内置函数共 <b class="c-op">13 个</b>：print、type、len、int、float、str、input、range、abs、round、max、min、sum。下一步逐个查看功能说明、示例代码与运行结果。`,
      draw: () => {
        $('#stage').innerHTML = this.frame('常见函数 · 总览', '准备', '13 个常用内置函数')
          + this.fnRack(-1, focus, 'cur')
          + `<div class="annot" style="margin-top:10px">每个函数都配有功能说明、示例代码与控制台运行结果；点击「下一步」逐个浏览。</div>`;
      }
    });

    list.forEach((f, k) => {
      const allIdx = this.FUNCS.indexOf(f);
      steps.push({
        desc: `<b class="c-op">${f.name}</b>：${f.purpose}`,
        draw: () => {
          $('#stage').innerHTML = this.frame('常见函数 · ' + f.name, `第 ${k + 1}/${list.length} 步`, f.example)
            + this.fnRack(allIdx, focus, focus ? 'done' : null)
            + `<div class="py-fn-card write-in">
                <div class="py-fn-name">${f.name}</div>
                <div class="py-fn-purpose">${f.purpose}</div>
              </div>
              <div class="two-col" style="margin-top:8px">
                <div>${this.codeBlock(f.code, -1)}</div>
                <div>${this.consoleOut(f.out)}</div>
              </div>`;
        }
      });
    });
    return steps;
  },

  /* 初始化演示画布（切换子页面时清空旧内容） */
  initStage() {
    const labelMap = {
      datatype: '数据类型 · type() 判断',
      operator: '运算符 · 六类运算',
      precedence: '运算符优先级 · 逐步化简',
      function: '常见内置函数 · 示例演示'
    };
    const tab = this.subTabs.find(t => t.id === this.currentSub);
    $('#stage').innerHTML = `<div class="stage-hd"><span>${labelMap[this.currentSub] || 'Python 表达式'}</span><span class="state">准备就绪</span></div>
      <div class="step-title"><span class="st-dot"></span>${tab ? tab.label : ''}</div>
      <div style="text-align:center;padding:clamp(18px,3.5vw,30px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.9vw,16px);margin-bottom:8px">点击「自动演示」或「下一步」开始浏览</div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">橙色 = 当前项 · 绿色 = 已完成 · 灰色 = 待浏览</div>
      </div>`;
  },
  generate() {
    switch (this.currentSub) {
      case 'operator': this.steps = this.genOperator(); break;
      case 'precedence': this.steps = this.genPrecedence(); break;
      case 'function': this.steps = this.genFunction(); break;
      default: this.steps = this.genDatatype();
    }
  }
};
