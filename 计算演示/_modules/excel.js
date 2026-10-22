/* ================================================================
 * Excel 公式
 * 子页面：ref_rel 相对引用 / ref_abs 绝对引用 / ref_mix 混合引用 / func 统计函数
 * 固定复制路径：B2 → D4（行偏移 +2，列偏移 +2）
 * ================================================================ */
MOD.excel = {
  name:'Excel公式', sub:'相对 · 绝对 · 混合引用复制 · 统计函数',
  badge:'表格 · 引用', tag:'表格', cat:'net',
  color:'#22C55E', soft:'#E8FBF0', type:'相对/绝对/混合引用复制 · SUM/AVERAGE/MAX/MIN',
  desc:'演示相对、绝对、混合引用在公式复制时的行列变化，并逐步计算 SUM/AVERAGE/MAX/MIN 统计函数。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/></svg>',
  setDesc:'引用类型固定由 B2 复制到 D4；统计函数可随机或手动选择。',
  stageDesc:'橙色=源单元格/当前计算，绿色=结果，观察复制时引用地址的偏移规律。',
  initDesc:'点击「自动演示」，观看公式引用变化或函数逐步计算。',
  currentSub:'ref_rel',
  subTabs:[
    {id:'ref_rel', label:'相对引用', method:'行列都偏移'},
    {id:'ref_abs', label:'绝对引用', method:'行列都锁定'},
    {id:'ref_mix', label:'混合引用', method:'锁列或锁行'},
    {id:'func',    label:'统计函数', method:'逐步求值'},
    {id:'func_type', label:'函数与数据类型', method:'SUM对数值/日期/逻辑值'}
  ],
  theory(){
    return `<div class="rule"><b>相对引用</b>（如 A1）：复制时行号和列号都会随目标位置偏移，向右复制列 +1，向下复制行 +1。</div>
    <div class="rule"><b>绝对引用</b>（如 $A$1）：复制时行号和列号都<b>不变</b>，$ 同时锁定行和列。</div>
    <div class="rule"><b>混合引用 $A1</b>：列绝对、行相对——列锁死不变，行随目标位置偏移。</div>
    <div class="rule"><b>混合引用 A$1</b>：列相对、行绝对——行锁死不变，列随目标位置偏移。</div>
    <div class="rule"><b>移动</b>公式（剪切粘贴）：无论哪种引用，引用地址都不变；只有<b>复制</b>时相对部分才会偏移。</div>
    <div class="rule"><b>常用函数</b>：SUM 求和、AVERAGE 平均值、MAX 最大值、MIN 最小值。</div>
    <div class="rule"><b>SUM 与数据类型</b>：区域引用 SUM(A1:A3) 会<b>忽略</b>单元格中的文本和逻辑值；直接写参数 SUM(TRUE,5) 时 TRUE=1、FALSE=0、文本数字会被转换；日期本质是序列号，会被当作数值相加。</div>`;
  },
  /* ---------- 参数区 ---------- */
  settingsHTML(){
    if(this.currentSub === 'func'){
      const old = $('#xl-func');
      const cur = old ? old.value : (this._func || 'auto');
      return `<div class="inline-setting"><div class="setting"><label>统计函数</label>
        <select id="xl-func">
          <option value="auto" ${cur==='auto'?'selected':''}>自动随机</option>
          <option value="SUM" ${cur==='SUM'?'selected':''}>SUM 求和</option>
          <option value="AVERAGE" ${cur==='AVERAGE'?'selected':''}>AVERAGE 平均值</option>
          <option value="MAX" ${cur==='MAX'?'selected':''}>MAX 最大值</option>
          <option value="MIN" ${cur==='MIN'?'selected':''}>MIN 最小值</option>
        </select></div></div>`;
    }
    return `<div class="inline-setting"><div class="setting"><label>复制路径</label>
      <div class="annot" style="padding:4px 8px">固定由 <b>B2</b> 复制到 <b>D4</b>（行 +2、列 +2）</div></div></div>`;
  },
  afterSettings(){
    const el = $('#xl-func');
    if(el) el.onchange = ()=>{ this._func = el.value; Engine.reload(); };
  },
  colLetter(n){ return String.fromCharCode(64 + n); },
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.05}s">${esc(ch)}</span>`).join('');
  },

  /* ---------- 引用复制：5×5 网格 ---------- */
  drawRefGrid(srcCell, dstCell, srcFormula, dstFormula, highlight, done, note){
    const cols = ['A','B','C','D','E'];
    let html = `<div class="stage-hd"><span>引用复制演示</span><span class="state">${srcCell} → ${dstCell}</span></div>
      <div class="step-title"><span class="st-dot"></span>公式复制 · 观察引用地址变化</div>`;
    html += `<div class="xl-grid" style="grid-template-columns:34px repeat(5,1fr)">`;
    html += `<div class="xl-cell hd"></div>` + cols.map(c=>`<div class="xl-cell hd">${c}</div>`).join('');
    for(let r=1; r<=5; r++){
      html += `<div class="xl-cell hd">${r}</div>`;
      for(let c=1; c<=5; c++){
        const cell = this.colLetter(c) + r;
        let cls = '', content = '';
        if(cell === srcCell){ cls = 'sel'; content = '=' + this.typeChars(srcFormula); }
        else if(cell === dstCell){ cls = done ? 'done' : 'cur'; content = dstFormula ? '=' + this.typeChars(dstFormula) : '?'; }
        html += `<div class="xl-cell ${cls}">${content}</div>`;
      }
    }
    html += `</div>`;
    html += `<div class="xl-formula">源 ${srcCell}: =${this.typeChars(srcFormula)} ${dstFormula ? `→ 目标 ${dstCell}: =<span class="write-in">${this.typeChars(dstFormula)}</span>` : ''}</div>`;
    if(note) html += `<div class="annot" style="text-align:center;margin-top:2px">${note}</div>`;
    $('#stage').innerHTML = html;
  },

  /* ---------- 单个引用类型的复制演示 ---------- */
  genRefCase(type, intro){
    const srcR = 2, srcC = 2;   // B2
    const dstR = 4, dstC = 4;   // D4
    const offR = dstR - srcR;   // +2 行
    const offC = dstC - srcC;   // +2 列
    const srcCell = this.colLetter(srcC) + srcR;
    const dstCell = this.colLetter(dstC) + dstR;
    let ref, dstRef, rule;
    if(type === 'rel'){
      ref = 'A1';
      dstRef = this.colLetter(1 + offC) + (1 + offR);
      rule = `列 A + ${offC} = <b class="c-op">${this.colLetter(1 + offC)}</b>，行 1 + ${offR} = <b class="c-op">${1 + offR}</b>`;
    }else if(type === 'abs'){
      ref = '$A$1';
      dstRef = '$A$1';
      rule = `$ 同时锁定列 A、锁定行 1，复制后<b>都不变</b>，仍为 <b class="c-op">$A$1</b>`;
    }else if(type === 'mixcol'){
      ref = '$A1';
      dstRef = '$A' + (1 + offR);
      rule = `列 A 前有 $ → <b>不变</b>；行 1 无 $ → 1 + ${offR} = <b class="c-op">${1 + offR}</b>`;
    }else{
      ref = 'A$1';
      dstRef = this.colLetter(1 + offC) + '$1';
      rule = `列 A 无 $ → A + ${offC} = <b class="c-op">${this.colLetter(1 + offC)}</b>；行 1 前有 $ → <b>不变</b>`;
    }
    const steps = [];
    // 1. 源单元格公式
    steps.push({ desc:`源单元格 <b class="c-op">${srcCell}</b> 的公式 =<b class="c-op">${ref}</b>，将其<b>复制</b>到 ${dstCell}（行偏移 +${offR}，列偏移 +${offC}）。`,
      draw:()=>{ this.drawRefGrid(srcCell, dstCell, ref, '', srcCell); }});
    // 2. 规则说明
    steps.push({ desc:intro,
      draw:()=>{ this.drawRefGrid(srcCell, dstCell, ref, '', dstCell); }});
    // 3. 推导目标引用
    steps.push({ desc:rule,
      draw:()=>{ this.drawRefGrid(srcCell, dstCell, ref, dstRef, dstCell); }});
    // 4. 复制完成
    steps.push({ desc:`复制完成：${srcCell} =${ref} → ${dstCell} =<b class="c-op">${dstRef}</b>`,
      draw:()=>{ this.drawRefGrid(srcCell, dstCell, ref, dstRef, dstCell, true); }});
    return steps;
  },

  /* ========== ① ref_rel：相对引用 ========== */
  genRefRel(){
    return this.genRefCase('rel', '相对引用：列号、行号都随目标位置一起偏移，偏移量 = 目标 − 源。');
  },
  /* ========== ② ref_abs：绝对引用 ========== */
  genRefAbs(){
    return this.genRefCase('abs', '绝对引用：$ 同时锁定列和行，复制后列号、行号都保持不变。');
  },
  /* ========== ③ ref_mix：混合引用（两种类型同页对比） ========== */
  genRefMix(){
    const mixCol = this.genRefCase('mixcol', '第一种 混合引用 $A1：<b>列绝对</b>（列锁死）、行相对（行偏移）。');
    const mixRow = this.genRefCase('mixrow', '第二种 混合引用 A$1：列相对（列偏移）、<b>行绝对</b>（行锁死）。');
    return mixCol.concat(mixRow);
  },

  /* ---------- 统计函数：数据网格 ---------- */
  drawFuncGrid(data, curIdx, funcName, result, done, calcHTML){
    let html = `<div class="stage-hd"><span>${funcName} 函数</span><span class="state">A1:A${data.length}</span></div>
      <div class="step-title"><span class="st-dot"></span>${funcName} 逐步求值 · 数据区域 A1:A${data.length}</div>`;
    html += `<div class="xl-grid" style="grid-template-columns:34px repeat(${data.length},1fr)">`;
    html += `<div class="xl-cell hd"></div>` + data.map((_,i)=>`<div class="xl-cell hd">${String.fromCharCode(65+i)}</div>`).join('');
    html += `<div class="xl-cell hd">1</div>` + data.map((v,i)=>`<div class="xl-cell ${i<curIdx?'done':i===curIdx?'cur':''} ${i===curIdx?'write-in':''}">${v}</div>`).join('');
    html += `</div>`;
    html += `<div class="xl-formula">=${funcName}(A1:A${data.length})${result ? ` = <b style="color:#86EFAC" class="write-in">${result}</b>` : ''}</div>`;
    if(calcHTML) html += calcHTML;
    if(done) html += `<div class="write-result">${funcName}(A1:A${data.length}) = ${result}</div>`;
    $('#stage').innerHTML = html;
  },

  /* ========== ④ func：统计函数逐步求值 ========== */
  genFunc(){
    const el = $('#xl-func');
    let f = el ? el.value : (this._func || 'auto');
    if(!f || f === 'auto') f = pick(['SUM','AVERAGE','MAX','MIN']);
    this._func = f;
    const datasets = {
      SUM:[3,7,2,8,5],
      AVERAGE:[4,8,6,9,3],
      MAX:[5,12,3,9,7],
      MIN:[8,3,10,5,2]
    };
    const data = datasets[f];
    const range = 'A1:A' + data.length;
    const steps = [];

    /* 计算过程卡片辅助：生成带步骤序号的累加/比较行 */
    const procLine = (idx, total, text, isCur)=>`
      <div class="cp-sum-line" style="margin-bottom:5px;${isCur?'background:#FFF4DB;border-color:var(--orange)':''}">
        <span style="color:var(--text-3);font-size:.85em">${idx+1}.</span> ${text}
      </div>`;

    // 1. 函数与数据区域
    steps.push({ desc:`函数 =<b class="c-op">${f}(${range})</b>，数据区域 ${range} = [${data.join(', ')}]，共 ${data.length} 个值。`,
      draw:()=>{ this.drawFuncGrid(data, -1, f, '', false,
        `<div class="calc-process"><div class="cp-label">待计算数据</div>
          <div class="cp-sum-line">[${data.join(', ')}] · 共 ${data.length} 个数</div></div>`); }});

    if(f === 'SUM'){
      let acc = 0;
      const history = [];
      data.forEach((v,i)=>{
        const oldAcc = acc;
        acc += v;
        history.push({old:oldAcc, add:v, new:acc});
        const snap = acc;
        const calcHTML = `<div class="calc-process">
          <div class="cp-label">逐步累加过程（第 ${i+1}/${data.length} 步）</div>
          ${history.map((h,j)=>procLine(j, data.length,
            `${h.old} <span class="op">+</span> ${h.add} <span class="eq">=</span> <span class="sum">${h.new}</span>`, j===i)).join('')}
        </div>`;
        steps.push({ desc:`第 ${i+1} 个值 <b class="c-op">${v}</b>：累计 ${oldAcc} + ${v} = <b class="c-op">${snap}</b>`,
          draw:()=>{ this.drawFuncGrid(data, i, f, String(snap), false, calcHTML); }});
      });
      const finalCalc = `<div class="calc-process">
        <div class="cp-label">完整求和式</div>
        <div class="cp-sum-line">${data.join(' + ')} <span class="eq">=</span> <span class="sum">${acc}</span></div>
      </div>`;
      steps.push({ desc:`SUM 求和结果 = <b class="c-op">${acc}</b>`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, String(acc), true, finalCalc); }});

    }else if(f === 'AVERAGE'){
      const sum = data.reduce((a,b)=>a+b,0);
      const avg = sum / data.length;
      // 2. 逐步求和
      let acc = 0;
      const sumHistory = [];
      data.forEach((v,i)=>{
        const oldAcc = acc;
        acc += v;
        sumHistory.push({old:oldAcc, add:v, new:acc});
      });
      const sumCalc = `<div class="calc-process">
        <div class="cp-label">第一步：求和（${data.length} 个数相加）</div>
        ${sumHistory.map((h,j)=>procLine(j, data.length,
          `${h.old} <span class="op">+</span> ${h.add} <span class="eq">=</span> <span class="sum">${h.new}</span>`, j===data.length-1)).join('')}
        <div class="cp-sum-line" style="margin-top:6px;background:#EEF3FF">
          总和 = <span style="color:var(--accent);font-weight:700">${sum}</span>
        </div>
      </div>`;
      steps.push({ desc:`先求和：${data.join(' + ')} = <b class="c-op">${sum}</b>（共 ${data.length} 个数相加）`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, sum + '（和）', false, sumCalc); }});
      // 3. 除以个数
      const divCalc = `<div class="calc-process">
        <div class="cp-label">第二步：求平均值 = 总和 ÷ 数据个数</div>
        <div class="cp-sum-line" style="margin-bottom:6px">
          总和 <span class="eq">=</span> <span style="color:var(--accent);font-weight:700">${sum}</span>
        </div>
        <div class="cp-sum-line" style="margin-bottom:6px">
          数据个数 <span class="eq">=</span> <span style="color:var(--accent);font-weight:700">${data.length}</span>
        </div>
        <div class="cp-sum-line" style="background:#FFF4DB">
          ${sum} <span class="op">÷</span> ${data.length} <span class="eq">=</span> <span class="sum">${avg.toFixed(2)}</span>
        </div>
      </div>`;
      steps.push({ desc:`再除以个数 ${data.length}：${sum} ÷ ${data.length} = <b class="c-op">${avg.toFixed(2)}</b>`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, avg.toFixed(2), true, divCalc); }});

    }else if(f === 'MAX'){
      let mx = data[0];
      const cmpHistory = [{val:data[0], result:data[0], note:'初始值'}];
      steps.push({ desc:`初始最大值 = 第一个值 <b class="c-op">${data[0]}</b>`,
        draw:()=>{ this.drawFuncGrid(data, 0, f, String(mx), false,
          `<div class="calc-process"><div class="cp-label">第 1 步：初始最大值</div>
            <div class="cp-sum-line">max = <span class="sum">${data[0]}</span>（取第一个值）</div></div>`); }});
      for(let i=1; i<data.length; i++){
        const oldMx = mx;
        const bigger = data[i] > oldMx;
        if(bigger) mx = data[i];
        cmpHistory.push({val:data[i], compared:oldMx, result:mx, bigger});
        const snap = mx;
        const calcHTML = `<div class="calc-process">
          <div class="cp-label">逐步比较（第 ${i+1}/${data.length} 步）</div>
          ${cmpHistory.slice(1).map((h,j)=>procLine(j, data.length,
            `max(${h.compared}, ${h.val}) <span class="eq">=</span> <span class="sum">${h.result}</span>${h.bigger?' <span style="color:var(--orange);font-size:.8em">↑ 更新</span>':' <span style="color:var(--text-3);font-size:.8em">保持</span>'}`,
            j===i-1)).join('')}
        </div>`;
        steps.push({ desc:`比较 ${data[i]} 与当前最大 ${oldMx}：${bigger ? `<b class="c-op">${data[i]} 更大，更新最大值为 ${data[i]}</b>` : `${data[i]} 不更大，保持 ${oldMx}`}`,
          draw:()=>{ this.drawFuncGrid(data, i, f, String(snap), false, calcHTML); }});
      }
      const finalCalc = `<div class="calc-process">
        <div class="cp-label">比较过程汇总</div>
        <div class="cp-sum-line">max(${data.join(', ')}) <span class="eq">=</span> <span class="sum">${mx}</span></div>
      </div>`;
      steps.push({ desc:`MAX 最大值 = <b class="c-op">${mx}</b>`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, String(mx), true, finalCalc); }});

    }else{ // MIN
      let mn = data[0];
      const cmpHistory = [{val:data[0], result:data[0], note:'初始值'}];
      steps.push({ desc:`初始最小值 = 第一个值 <b class="c-op">${data[0]}</b>`,
        draw:()=>{ this.drawFuncGrid(data, 0, f, String(mn), false,
          `<div class="calc-process"><div class="cp-label">第 1 步：初始最小值</div>
            <div class="cp-sum-line">min = <span class="sum">${data[0]}</span>（取第一个值）</div></div>`); }});
      for(let i=1; i<data.length; i++){
        const oldMn = mn;
        const smaller = data[i] < oldMn;
        if(smaller) mn = data[i];
        cmpHistory.push({val:data[i], compared:oldMn, result:mn, smaller});
        const snap = mn;
        const calcHTML = `<div class="calc-process">
          <div class="cp-label">逐步比较（第 ${i+1}/${data.length} 步）</div>
          ${cmpHistory.slice(1).map((h,j)=>procLine(j, data.length,
            `min(${h.compared}, ${h.val}) <span class="eq">=</span> <span class="sum">${h.result}</span>${h.smaller?' <span style="color:var(--orange);font-size:.8em">↓ 更新</span>':' <span style="color:var(--text-3);font-size:.8em">保持</span>'}`,
            j===i-1)).join('')}
        </div>`;
        steps.push({ desc:`比较 ${data[i]} 与当前最小 ${oldMn}：${smaller ? `<b class="c-op">${data[i]} 更小，更新最小值为 ${data[i]}</b>` : `${data[i]} 不更小，保持 ${oldMn}`}`,
          draw:()=>{ this.drawFuncGrid(data, i, f, String(snap), false, calcHTML); }});
      }
      const finalCalc = `<div class="calc-process">
        <div class="cp-label">比较过程汇总</div>
        <div class="cp-sum-line">min(${data.join(', ')}) <span class="eq">=</span> <span class="sum">${mn}</span></div>
      </div>`;
      steps.push({ desc:`MIN 最小值 = <b class="c-op">${mn}</b>`,
        draw:()=>{ this.drawFuncGrid(data, data.length, f, String(mn), true, finalCalc); }});
    }
    return steps;
  },

  /* ========== ⑤ func_type：SUM 与数据类型对比 ========== */
  genFuncType(){
    const steps = [];
    /* 三组对比数据 */
    const groupNum = { label:'纯数值', cells:[
      {addr:'A1', val:'3',   type:'num',   seq:3},
      {addr:'A2', val:'7',   type:'num',   seq:7},
      {addr:'A3', val:'2',   type:'num',   seq:2}
    ], sum:12 };
    const groupDate = { label:'含日期', cells:[
      {addr:'B1', val:'2024-01-01', type:'date', seq:45292},
      {addr:'B2', val:'5',          type:'num',  seq:5},
      {addr:'B3', val:'3',          type:'num',  seq:3}
    ], sum:45300 };
    const groupMixed = { label:'含逻辑值/文本', cells:[
      {addr:'C1', val:'TRUE',  type:'bool',   seq:1, ignored:true},
      {addr:'C2', val:'FALSE', type:'bool',   seq:0, ignored:true},
      {addr:'C3', val:'"5"',   type:'text',   seq:5, ignored:true},
      {addr:'C4', val:'3',     type:'num',    seq:3}
    ], sum:3 };

    /* 渲染一组数据的小表格 */
    const miniGrid = (group, highlightIdx, showSeq)=>{
      let html = `<div style="flex:1;min-width:0">
        <div style="font-size:clamp(10px,1.2vw,12px);font-weight:700;color:var(--text-3);margin-bottom:4px;text-align:center">${group.label}</div>
        <div class="xl-grid" style="grid-template-columns:30px 1fr">`;
      group.cells.forEach((c,i)=>{
        const cls = i===highlightIdx ? 'cur' : (showSeq && !c.ignored ? 'done' : '');
        const display = showSeq && c.type==='date' ? c.seq : c.val;
        html += `<div class="xl-cell hd" style="font-size:10px">${c.addr}</div>
          <div class="xl-cell ${cls}" style="${c.ignored?'opacity:.45;text-decoration:line-through':''}">${display}</div>`;
      });
      html += `</div></div>`;
      return html;
    };

    // 1. 三组数据并排展示
    steps.push({ desc:`同样是 SUM 函数，数据类型不同，结果大不相同。对比三组数据：纯数值、含日期、含逻辑值/文本。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>SUM 与数据类型</span><span class="state">对比三组数据</span></div>
          <div class="step-title"><span class="st-dot"></span>三组数据并排对比（区域引用 SUM(范围)）</div>
          <div style="display:flex;gap:10px;flex-wrap:wrap">
            ${miniGrid(groupNum, -1, false)}
            ${miniGrid(groupDate, -1, false)}
            ${miniGrid(groupMixed, -1, false)}
          </div>
          <div class="annot" style="text-align:center;margin-top:8px">C1/C2 逻辑值、C3 文本数字在区域引用中会被 SUM 忽略</div>`;
      }});

    // 2. 第一组：纯数值
    steps.push({ desc:`第一组「纯数值」：SUM(A1:A3) = 3 + 7 + 2 = <b class="c-op">12</b>，所有值都参与计算。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>SUM 与数据类型</span><span class="state">第1组：纯数值</span></div>
          <div class="step-title"><span class="st-dot"></span>纯数值 — 全部参与计算</div>
          ${miniGrid(groupNum, -1, true)}
          <div class="calc-process">
            <div class="cp-label">SUM(A1:A3) 计算过程</div>
            <div class="cp-sum-line" style="margin-bottom:5px">3 <span class="op">+</span> 7 <span class="op">+</span> 2 <span class="eq">=</span> <span class="sum">12</span></div>
            <div class="cp-sum-line" style="background:#E8FBF0">数值型数据 → <b style="color:#15803D">全部参与求和</b></div>
          </div>`;
      }});

    // 3. 第二组：含日期
    steps.push({ desc:`第二组「含日期」：日期在 Excel 中本质是序列号（2024-01-01 = 45292），SUM 会把日期当数值相加：45292 + 5 + 3 = <b class="c-op">45300</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>SUM 与数据类型</span><span class="state">第2组：含日期</span></div>
          <div class="step-title"><span class="st-dot"></span>日期 = 序列号，参与求和</div>
          ${miniGrid(groupDate, -1, true)}
          <div class="calc-process">
            <div class="cp-label">SUM(B1:B3) 计算过程（日期转序列号）</div>
            <div class="cp-sum-line" style="margin-bottom:5px">2024-01-01 → 序列号 <b style="color:var(--accent)">45292</b></div>
            <div class="cp-sum-line" style="margin-bottom:5px">45292 <span class="op">+</span> 5 <span class="op">+</span> 3 <span class="eq">=</span> <span class="sum">45300</span></div>
            <div class="cp-sum-line" style="background:#E8FBF0">日期型数据 → <b style="color:#15803D">转序列号后参与求和</b></div>
          </div>`;
      }});

    // 4. 第三组：含逻辑值/文本（区域引用）
    steps.push({ desc:`第三组「含逻辑值/文本」：区域引用 SUM(C1:C4) 中，TRUE/FALSE 逻辑值和 "5" 文本数字都被<b>忽略</b>，只有数值 3 参与计算，结果 = <b class="c-op">3</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>SUM 与数据类型</span><span class="state">第3组：含逻辑值/文本</span></div>
          <div class="step-title"><span class="st-dot"></span>区域引用 — 逻辑值和文本被忽略</div>
          ${miniGrid(groupMixed, -1, true)}
          <div class="calc-process">
            <div class="cp-label">SUM(C1:C4) 计算过程（区域引用）</div>
            <div class="cp-sum-line" style="margin-bottom:5px;opacity:.45;text-decoration:line-through">TRUE → 被忽略 &nbsp; FALSE → 被忽略 &nbsp; "5" → 被忽略</div>
            <div class="cp-sum-line" style="margin-bottom:5px">只有数值 3 参与 → <span class="sum">3</span></div>
            <div class="cp-sum-line" style="background:#FEF2F2">区域引用中的逻辑值/文本 → <b style="color:#DC2626">被 SUM 忽略</b></div>
          </div>`;
      }});

    // 5. 对比总结：直接参数 vs 区域引用
    steps.push({ desc:`关键区别：<b>区域引用</b> SUM(A1:A4) 忽略逻辑值和文本；但<b>直接写参数</b> SUM(TRUE, FALSE, "5", 3) 时，TRUE=1、FALSE=0、"5"转数值5，结果 = 1+0+5+3 = <b class="c-op">9</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>SUM 与数据类型</span><span class="state">关键区别</span></div>
          <div class="step-title"><span class="st-dot"></span>区域引用 vs 直接参数 — 结果完全不同</div>
          <div style="display:flex;gap:10px;flex-wrap:wrap">
            <div style="flex:1;min-width:200px">
              <div class="calc-process" style="margin:0">
                <div class="cp-label">区域引用 SUM(C1:C4)</div>
                <div class="cp-sum-line" style="margin-bottom:5px;opacity:.45">TRUE FALSE "5" 被忽略</div>
                <div class="cp-sum-line" style="background:#FEF2F2">结果 = <b style="color:#DC2626">3</b></div>
              </div>
            </div>
            <div style="flex:1;min-width:200px">
              <div class="calc-process" style="margin:0">
                <div class="cp-label">直接参数 SUM(TRUE,FALSE,"5",3)</div>
                <div class="cp-sum-line" style="margin-bottom:5px">TRUE=1, FALSE=0, "5"→5</div>
                <div class="cp-sum-line" style="margin-bottom:5px">1 + 0 + 5 + 3 = 9</div>
                <div class="cp-sum-line" style="background:#E8FBF0">结果 = <b style="color:#15803D">9</b></div>
              </div>
            </div>
          </div>
          <div class="write-result">同样的数据，引用方式不同 → 结果不同！区域引用忽略逻辑值/文本，直接参数会转换它们。</div>`;
      }});
    return steps;
  },

  /* ---------- 初始化画布 ---------- */
  initStage(){
    const sub = this.currentSub;
    const info = this.subTabs.find(t=>t.id===sub) || {};
    const labelMap = { ref_rel:'相对引用复制', ref_abs:'绝对引用复制', ref_mix:'混合引用复制', func:'统计函数逐步求值', func_type:'SUM与数据类型对比' };
    $('#stage').innerHTML = `<div class="stage-hd"><span>${labelMap[sub]||'Excel 公式'}</span><span class="state">准备就绪</span></div>
      <div style="text-align:center;padding:clamp(16px,3vw,28px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.8vw,16px);margin-bottom:8px">使用 <b style="color:#22C55E">${info.method||''}</b></div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">点击「自动演示」或「下一步」开始推导</div>
      </div>`;
  },
  generate(){
    switch(this.currentSub){
      case 'ref_rel': this.steps = this.genRefRel(); break;
      case 'ref_abs': this.steps = this.genRefAbs(); break;
      case 'ref_mix': this.steps = this.genRefMix(); break;
      case 'func':    this.steps = this.genFunc(); break;
      case 'func_type': this.steps = this.genFuncType(); break;
      default: this.steps = this.genRefRel();
    }
  }
};
