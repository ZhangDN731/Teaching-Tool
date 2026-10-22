/* ================================================================
 * ⑨ 队列 Queue（普通队列 FIFO · 双端队列 Deque）
 * 子页面：normal 普通队列 / deque 双端队列
 * ================================================================ */
MOD.queue = {
  name:'队列 Queue', sub:'入队 enqueue · 出队 dequeue · 双端队列',
  badge:'结构 · FIFO', tag:'队列', cat:'ds',
  color:'#14B8A6', soft:'#E6FAF7', type:'普通队列 · 双端队列 · 先进先出',
  desc:'演示普通队列（先进先出 FIFO：队尾入、队头出）与双端队列（左右两端均可进出）的逐步操作过程。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="7" width="5" height="10" rx="1.5"/><rect x="9" y="7" width="5" height="10" rx="1.5"/><rect x="16" y="7" width="5" height="10" rx="1.5"/><path d="M4.5 4v3M11.5 4v3M18.5 4v3M4.5 17v3M11.5 17v3M18.5 17v3"/></svg>',
  setDesc:'',
  stageDesc:'普通队列队头在左、队尾在右；橙色=当前操作元素，绿色=已在队列中。',
  initDesc:'点击「自动演示」或「下一步」，观看队列的进出操作。',
  currentSub:'normal',
  subTabs:[
    {id:'normal', label:'普通队列', method:'enqueue / dequeue · FIFO'},
    {id:'deque',  label:'双端队列', method:'append / appendleft · pop / popleft'}
  ],
  theory(){
    return `<div class="rule"><b>队列（Queue）</b>：线性数据结构，一端（队尾）插入，另一端（队头）删除。</div>
    <div class="rule"><b>先进先出 FIFO</b>：First In First Out，最先入队的元素最先出队。</div>
    <div class="rule"><b>enqueue（入队）</b>：在队尾添加元素，队列长度 +1；<b>dequeue（出队）</b>：从队头移除元素，长度 -1。</div>
    <div class="rule"><b>双端队列（Deque）</b>：两端都可插入和删除。append（右端入）、appendleft（左端入）、pop（右端出）、popleft（左端出）。</div>
    <div class="rule">应用：打印任务排队、广度优先搜索 BFS、消息队列、售票窗口排队。</div>`;
  },
  settingsHTML(){
    const cur = $('#q-n') ? $('#q-n').value : '5';
    return `<div class="inline-setting"><div class="setting"><label>操作步数</label>
      <select id="q-n">
        <option value="4" ${cur==='4'?'selected':''}>4 步</option>
        <option value="5" ${cur==='5'?'selected':''}>5 步</option>
        <option value="6" ${cur==='6'?'selected':''}>6 步</option>
      </select></div></div>`;
  },
  afterSettings(){ $('#q-n').onchange=()=>Engine.reload(); },

  /* 逐字出现动画辅助 */
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.06}s">${esc(ch)}</span>`).join('');
  },

  /* 统一画布渲染：o = {title, state, queue, mode, curVal, popped, log, expr, result, resultLabel, endLabels} */
  renderQueue(o){
    const subName=this.currentSub==='deque'?'双端队列':'普通队列';
    const size=o.queue.length;
    const labels=o.endLabels||{left:'队头',right:'队尾'};

    /* 队列方块：DOM 顺序 队头(左) → 队尾(右) */
    const chips=o.queue.map((v,i)=>{
      const isHead=i===0;
      const isTail=i===size-1;
      let isCur=false;
      if(o.mode==='insert-left'&&isHead&&v===o.curVal) isCur=true;
      if(o.mode==='insert-right'&&isTail&&v===o.curVal) isCur=true;
      const cls=isCur?'o':'g';
      const anim=isCur?' write-in':'';
      return `<div class="chip ${cls}${anim}" style="position:relative;min-width:clamp(38px,5.5vw,46px);height:clamp(24px,3.2vw,28px);padding:2px 8px;font-size:12px;${isCur?'animation-delay:.1s':''}">${esc(v)}</div>`;
    }).join('');
    const bodyHtml=size===0
      ? '<div style="color:#CBD5E1;font-size:12px;font-family:var(--mono)">（空队列）</div>'
      : chips;

    /* 出队“离开”幽灵方块：左端离开 / 右端离开 */
    let ghostLeft='', ghostRight='';
    if(o.mode==='remove-left'&&o.popped!==null){
      ghostLeft=`<div class="chip o write-in" style="position:relative;min-width:clamp(38px,5.5vw,46px);height:clamp(24px,3.2vw,28px);padding:2px 8px;font-size:12px">${esc(o.popped)}</div><span style="color:#F59E0B;font-size:10.5px;font-family:var(--mono);font-weight:700;white-space:nowrap">离开 ←</span>`;
    }
    if(o.mode==='remove-right'&&o.popped!==null){
      ghostRight=`<span style="color:#F59E0B;font-size:10.5px;font-family:var(--mono);font-weight:700;white-space:nowrap">→ 离开</span><div class="chip o write-in" style="position:relative;min-width:clamp(38px,5.5vw,46px);height:clamp(24px,3.2vw,28px);padding:2px 8px;font-size:12px">${esc(o.popped)}</div>`;
    }

    /* 操作表达式（逐字写入 + 笔形光标） */
    let exprHtml='';
    if(o.expr){
      const e=o.expr;
      const nameMap={enqueue:`enqueue(${e.v})`,dequeue:'dequeue()',append:`append(${e.v})`,appendleft:`appendleft(${e.v})`,pop:'pop()',popleft:'popleft()'};
      const sideMap={enqueue:`队尾 ← ${e.v}`,dequeue:`返回 ${e.v}`,append:`右端 ← ${e.v}`,appendleft:`左端 ← ${e.v}`,pop:`返回 ${e.v}（右端）`,popleft:`返回 ${e.v}（左端）`};
      const sizeLabel=(e.kind==='enqueue'||e.kind==='dequeue')?'长度':'大小';
      const sizeTxt=e.prev!==undefined?`${sizeLabel} ${e.prev} → ${e.size}`:`${sizeLabel} ${e.before} → ${e.size}`;
      exprHtml=`<span class="term cur">${this.typeChars(nameMap[e.kind])}</span><span class="op">=</span><span class="term">${sideMap[e.kind]}</span><span class="op">·</span><span class="term done">${sizeTxt}</span><span class="caret pen"></span>`;
    }

    /* 完成态分组框 + 结果框 */
    let resultHtml='';
    if(o.result){
      const r=o.result;
      const grps=r.map((v,i)=>{
        const lbl=r.length===1?'':i===0?labels.left:i===r.length-1?labels.right:'';
        return `<span class="grp done"><span class="grp-label">${lbl}</span>${esc(v)}</span>`;
      }).join(r.length>1?'<span class="arrow">→</span>':'');
      resultHtml=`<div class="group-box" style="justify-content:center;margin:4px 0">${grps}</div>
        <div class="write-result" style="margin-top:3px;padding:4px 10px">${o.resultLabel||'队头 → 队尾'}：${r.length?r.join(' , '):'（空队列）'} · 共 ${r.length} 个元素</div>`;
    }

    $('#stage').innerHTML=`<div class="stage-hd"><span>队列 Queue · ${subName}</span><span class="state">${o.state||`长度 ${size}`}</span></div>
      <div class="step-title"><span class="st-dot"></span>${o.title||''}</div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:nowrap">
        <div style="flex:0 0 58px;min-height:44px;display:flex;justify-content:center;align-items:center;gap:4px;flex-wrap:wrap">${ghostLeft}</div>
        <div style="flex:1;min-width:0">
          <div style="display:flex;justify-content:space-between;font-size:9.5px;font-weight:700;color:#14B8A6;padding:0 4px 3px">
            <span>${labels.left}</span><span>${labels.right}</span>
          </div>
          <div style="display:flex;gap:6px;align-items:center;justify-content:center;background:#fff;border:1px solid var(--border);border-radius:10px;padding:10px;min-height:52px;overflow-x:auto">${bodyHtml}</div>
        </div>
        <div style="flex:0 0 58px;min-height:44px;display:flex;justify-content:center;align-items:center;gap:4px;flex-wrap:wrap">${ghostRight}</div>
      </div>
      <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:stretch;margin-top:6px">
        <div style="flex:1 1 150px;min-width:140px">
          <div class="zone-label" style="margin-bottom:0">操作日志</div>
          <div class="ds-log" style="margin-top:2px;line-height:1.6">${o.log||''}</div>
        </div>
        ${exprHtml?`<div style="flex:1 1 150px;min-width:140px"><div class="zone-label" style="margin-bottom:0">操作表达式</div><div class="exp-inline" style="margin-top:2px">${exprHtml}</div></div>`:''}
      </div>
      ${resultHtml}`;
  },

  /* ========== 子页面 1：普通队列 enqueue / dequeue（FIFO） ========== */
  genNormal(){
    const n=+$('#q-n').value||5;
    const ops=[]; let size=0;
    for(let i=0;i<n;i++){
      if(size===0||size<2){ ops.push({op:'enqueue',val:randInt(1,9)}); size++; }
      else if(size>=5){ ops.push({op:'dequeue'}); size--; }
      else if(Math.random()<0.55){ ops.push({op:'enqueue',val:randInt(1,9)}); size++; }
      else { ops.push({op:'dequeue'}); size--; }
    }
    /* 回放得到每一步快照，保证画面与描述一致 */
    const q=[]; const replay=[];
    ops.forEach(o=>{
      if(o.op==='enqueue'){ q.push(o.val); replay.push({op:'enqueue',val:o.val,snap:[...q]}); }
      else { const v=q.shift(); replay.push({op:'dequeue',val:v,snap:[...q]}); }
    });
    const steps=[];
    steps.push({desc:`初始状态：<b class="c-op">空队列</b>，队头与队尾重合；新元素从<b>队尾</b>进入，从<b>队头</b>离开（FIFO）。`,
      draw:()=>{ this.renderQueue({title:'普通队列', state:'空队列 · 准备就绪', queue:[], mode:'ready', curVal:null, popped:null, log:'（尚未开始操作）', expr:null, endLabels:{left:'队头',right:'队尾'}}); }});
    replay.forEach((r,k)=>{
      if(r.op==='enqueue'){
        const prev=r.snap.length-1;
        steps.push({desc:`第 ${k+1} 步：<b class="c-op">enqueue(${r.val})</b>，在队尾加入 ${r.val}；队列长度 <b>${prev}</b> → <b>${r.snap.length}</b>（+1）。`,
          draw:()=>{ this.renderQueue({title:'入队 enqueue', state:`长度 ${r.snap.length}`, queue:r.snap, mode:'insert-right', curVal:r.val, popped:null,
            log:`<b>enqueue(${r.val})</b>：队尾新增 ${r.val}<br>队列长度 ${prev} → ${r.snap.length}（+1）`,
            expr:{kind:'enqueue', v:r.val, prev:prev, size:r.snap.length}, endLabels:{left:'队头',right:'队尾'}}); }});
      }else{
        const before=r.snap.length+1;
        steps.push({desc:`第 ${k+1} 步：<b class="c-op">dequeue()</b>，从队头移除 <b class="c-op">${r.val}</b>——先进先出，最先入队的最先出队；长度 <b>${before}</b> → <b>${r.snap.length}</b>（-1）。`,
          draw:()=>{ this.renderQueue({title:'出队 dequeue', state:`长度 ${r.snap.length}`, queue:r.snap, mode:'remove-left', curVal:null, popped:r.val,
            log:`<b>dequeue()</b>：返回队头元素 ${r.val}<br>队列长度 ${before} → ${r.snap.length}（-1）`,
            expr:{kind:'dequeue', v:r.val, before:before, size:r.snap.length}, endLabels:{left:'队头',right:'队尾'}}); }});
      }
    });
    const finalSnap=replay.length?replay[replay.length-1].snap:[];
    steps.push({desc:`普通队列操作结束：队头 → 队尾共 <b class="c-op">${finalSnap.length}</b> 个元素${finalSnap.length?`（${finalSnap.join('、')}）`:'（空队列）'}；先入队的元素始终排在队头，体现 <b class="c-op">FIFO 先进先出</b>。`,
      draw:()=>{ this.renderQueue({title:'普通队列完成', state:`共 ${finalSnap.length} 个`, queue:finalSnap, mode:'done', curVal:null, popped:null,
        log:`普通队列执行完毕，队头 → 队尾：${finalSnap.join('、')}`, expr:null, result:finalSnap, resultLabel:'队头 → 队尾', endLabels:{left:'队头',right:'队尾'}}); }});
    return steps;
  },

  /* ========== 子页面 2：双端队列 append / appendleft / pop / popleft ========== */
  genDeque(){
    const n=+$('#q-n').value||5;
    const ops=[]; let size=0;
    for(let i=0;i<n;i++){
      if(size===0){ ops.push({op:pick(['append','appendleft']),val:randInt(1,9)}); size++; }
      else if(size>=5){ ops.push({op:pick(['pop','popleft'])}); size--; }
      else {
        const r=Math.random();
        if(r<0.28){ ops.push({op:'append',val:randInt(1,9)}); size++; }
        else if(r<0.56){ ops.push({op:'appendleft',val:randInt(1,9)}); size++; }
        else if(r<0.78){ ops.push({op:'pop'}); size--; }
        else { ops.push({op:'popleft'}); size--; }
      }
    }
    /* 回放得到每一步快照，保证画面与描述一致 */
    const q=[]; const replay=[];
    ops.forEach(o=>{
      if(o.op==='append'){ q.push(o.val); replay.push({op:'append',val:o.val,snap:[...q]}); }
      else if(o.op==='appendleft'){ q.unshift(o.val); replay.push({op:'appendleft',val:o.val,snap:[...q]}); }
      else if(o.op==='pop'){ const v=q.pop(); replay.push({op:'pop',val:v,snap:[...q]}); }
      else { const v=q.shift(); replay.push({op:'popleft',val:v,snap:[...q]}); }
    });
    const steps=[];
    steps.push({desc:`初始状态：<b class="c-op">空双端队列</b>。左右两端都可以插入和删除：append（右端入）、appendleft（左端入）、pop（右端出）、popleft（左端出）。`,
      draw:()=>{ this.renderQueue({title:'双端队列', state:'空队列 · 准备就绪', queue:[], mode:'ready', curVal:null, popped:null, log:'（尚未开始操作）', expr:null, endLabels:{left:'左端',right:'右端'}}); }});
    replay.forEach((r,k)=>{
      const prev=r.snap.length-1, before=r.snap.length+1;
      let mode='done', curVal=null, popped=null, title='', log='', expr=null;
      if(r.op==='append'){
        mode='insert-right'; curVal=r.val; title='右端入 append';
        log=`<b>append(${r.val})</b>：右端新增 ${r.val}<br>队列大小 ${prev} → ${r.snap.length}（+1）`;
        expr={kind:'append', v:r.val, prev:prev, size:r.snap.length};
      }else if(r.op==='appendleft'){
        mode='insert-left'; curVal=r.val; title='左端入 appendleft';
        log=`<b>appendleft(${r.val})</b>：左端新增 ${r.val}<br>队列大小 ${prev} → ${r.snap.length}（+1）`;
        expr={kind:'appendleft', v:r.val, prev:prev, size:r.snap.length};
      }else if(r.op==='pop'){
        mode='remove-right'; popped=r.val; title='右端出 pop';
        log=`<b>pop()</b>：返回右端元素 ${r.val}<br>队列大小 ${before} → ${r.snap.length}（-1）`;
        expr={kind:'pop', v:r.val, before:before, size:r.snap.length};
      }else{
        mode='remove-left'; popped=r.val; title='左端出 popleft';
        log=`<b>popleft()</b>：返回左端元素 ${r.val}<br>队列大小 ${before} → ${r.snap.length}（-1）`;
        expr={kind:'popleft', v:r.val, before:before, size:r.snap.length};
      }
      const descMap={
        append:`第 ${k+1} 步：<b class="c-op">append(${r.val})</b>，在右端加入 ${r.val}；队列大小 <b>${prev}</b> → <b>${r.snap.length}</b>（+1）。`,
        appendleft:`第 ${k+1} 步：<b class="c-op">appendleft(${r.val})</b>，在左端加入 ${r.val}；队列大小 <b>${prev}</b> → <b>${r.snap.length}</b>（+1）。`,
        pop:`第 ${k+1} 步：<b class="c-op">pop()</b>，从右端移除 <b class="c-op">${r.val}</b>；队列大小 <b>${before}</b> → <b>${r.snap.length}</b>（-1）。`,
        popleft:`第 ${k+1} 步：<b class="c-op">popleft()</b>，从左端移除 <b class="c-op">${r.val}</b>；队列大小 <b>${before}</b> → <b>${r.snap.length}</b>（-1）。`
      };
      steps.push({desc:descMap[r.op],
        draw:()=>{ this.renderQueue({title:title, state:`大小 ${r.snap.length}`, queue:r.snap, mode:mode, curVal:curVal, popped:popped, log:log, expr:expr, endLabels:{left:'左端',right:'右端'}}); }});
    });
    const finalSnap=replay.length?replay[replay.length-1].snap:[];
    steps.push({desc:`双端队列操作结束：左端 → 右端共 <b class="c-op">${finalSnap.length}</b> 个元素${finalSnap.length?`（${finalSnap.join('、')}）`:'（空队列）'}；两端均可进出，元素顺序由每次插入的端决定。`,
      draw:()=>{ this.renderQueue({title:'双端队列完成', state:`共 ${finalSnap.length} 个`, queue:finalSnap, mode:'done', curVal:null, popped:null,
        log:`双端队列执行完毕，左端 → 右端：${finalSnap.join('、')}`, expr:null, result:finalSnap, resultLabel:'左端 → 右端', endLabels:{left:'左端',right:'右端'}}); }});
    return steps;
  },

  /* 切换子页面时清空画布，显示空队列与准备状态 */
  initStage(){
    const map={
      normal:{label:'普通队列 · enqueue 入队 / dequeue 出队（FIFO）', title:'普通队列', left:'队头', right:'队尾'},
      deque: {label:'双端队列 · 左右两端均可进出', title:'双端队列', left:'左端', right:'右端'}
    };
    const m=map[this.currentSub]||map.normal;
    $('#stage').innerHTML=`<div class="stage-hd"><span>队列 Queue</span><span class="state">空队列 · 准备就绪</span></div>
      <div class="step-title"><span class="st-dot"></span>${m.title}</div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:nowrap">
        <div style="flex:0 0 58px"></div>
        <div style="flex:1;min-width:0">
          <div style="display:flex;justify-content:space-between;font-size:9.5px;font-weight:700;color:#14B8A6;padding:0 4px 3px"><span>${m.left}</span><span>${m.right}</span></div>
          <div style="display:flex;gap:6px;align-items:center;justify-content:center;background:#fff;border:1px solid var(--border);border-radius:10px;padding:10px;min-height:52px">
            <div style="color:#CBD5E1;font-size:12px;font-family:var(--mono)">（空队列）</div>
          </div>
        </div>
        <div style="flex:0 0 58px"></div>
      </div>
      <div style="margin-top:6px">
        <div class="zone-label" style="margin-bottom:0">操作日志</div>
        <div class="ds-log" style="margin-top:2px;line-height:1.6">（尚未开始操作）</div>
        <div style="margin-top:4px;color:#94A3B8;font-size:clamp(10.5px,1.3vw,12px);line-height:1.6">当前模式：<b style="color:#14B8A6">${m.label}</b>，点击「下一步」开始演示</div>
      </div>`;
  },

  generate(){
    switch(this.currentSub){
      case 'normal': this.steps=this.genNormal(); break;
      case 'deque':  this.steps=this.genDeque(); break;
      default:       this.steps=this.genNormal();
    }
  }
};
