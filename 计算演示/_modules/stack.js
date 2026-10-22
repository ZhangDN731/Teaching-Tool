/* ================================================================
 * ⑧ 栈与队列（合并模块）
 *   1) 栈 Stack（LIFO）：栈顶一端进出 push / pop
 *   2) 标准队列 Queue（FIFO）：队尾入 enqueue / 队头出 dequeue
 *   3) 双端入 · 一端出队列：左右均可入，仅左端出
 *   4) 一端入 · 两端出队列：仅右端入，左右均可出
 * 交互式模块：输入元素值后点击按钮自主操作，自有状态管理，
 * 每次操作记入日志，支持空栈 / 空队列下溢提示。
 * ================================================================ */

/* ---------- 注入模块专用样式（大屏幕放大 + 进出动画） ---------- */
if(!document.getElementById('ds-stack-queue-style')){
  document.head.insertAdjacentHTML('beforeend', `<style id="ds-stack-queue-style">
    /* 操作按钮：大触控热区 */
    .ds-op{min-height:clamp(38px,5vw,46px);font-size:clamp(12.5px,1.6vw,14.5px);gap:6px}
    /* 栈 chip：纵向大尺寸 */
    .ds-chip-v{position:relative;min-width:clamp(64px,9vw,86px);min-height:clamp(44px,6vw,56px);
      font-size:clamp(14px,2vw,18px);padding:clamp(6px,1vw,9px) clamp(10px,1.6vw,14px)}
    /* 队列 chip：横向大尺寸 */
    .ds-chip-h{min-width:clamp(58px,8vw,78px);min-height:clamp(38px,5vw,50px);
      font-size:clamp(14px,2vw,18px);padding:clamp(6px,1vw,9px) clamp(11px,1.8vw,15px)}
    /* 栈容器：底部封闭、顶部开放 */
    .ds-stack-area{display:flex;flex-direction:column-reverse;gap:clamp(6px,1vw,8px);
      align-items:center;justify-content:flex-end;min-height:clamp(210px,28vw,270px);
      max-height:clamp(260px,36vw,330px);overflow-y:auto;padding:clamp(30px,4vw,40px) 14px 12px;
      background:#fff;border:1px solid var(--border);border-bottom:4px solid #8B5CF6;
      border-radius:0 0 12px 12px;scrollbar-width:thin}
    /* 队列容器 */
    .ds-queue-area{display:flex;flex-direction:row;gap:clamp(7px,1.2vw,10px);align-items:center;
      justify-content:center;flex-wrap:wrap;min-height:clamp(64px,9vw,84px);padding:12px;
      background:#fff;border:1px solid var(--border);border-radius:12px}
    .ds-empty{color:#CBD5E1;font-family:var(--mono);font-size:clamp(13px,1.8vw,16px)}
    /* 侧端标签（栈顶 / 栈底 / 队头 / 队尾） */
    .side-tag{position:absolute;font-size:clamp(9.5px,1.2vw,11px);font-weight:700;white-space:nowrap}
    .side-top{left:100%;top:50%;transform:translateY(-50%);margin-left:12px;color:#8B5CF6}
    .side-bottom{right:100%;top:50%;transform:translateY(-50%);margin-right:12px;color:#94A3B8}
    /* 幽灵动画槽位 */
    .ds-ghost-v-row{min-height:clamp(46px,7vw,64px);display:flex;align-items:center;justify-content:center;gap:10px;margin-bottom:6px}
    .ds-ghost-v{display:flex;align-items:center;gap:10px}
    .ds-ghost-h{display:flex;align-items:center;gap:6px}
    .ghost-tag{font-size:clamp(10px,1.3vw,12px);font-family:var(--mono);font-weight:700;color:#F59E0B;white-space:nowrap}
    .ds-ghost-slot{flex:0 0 clamp(70px,10vw,96px);min-height:clamp(64px,9vw,84px);display:flex;align-items:center;justify-content:center}
    /* 队列左右箭头标注行 */
    .ds-q-row{display:flex;gap:clamp(8px,1.2vw,12px);align-items:stretch}
    .ds-q-labels{display:flex;justify-content:space-between;font-size:clamp(10px,1.3vw,12px);font-weight:700;color:#8B5CF6;padding:0 6px 5px}
    /* 出栈 / 出队消失动画 */
    @keyframes dsOutUp{from{opacity:1;transform:translateY(0) scale(1)}to{opacity:0;transform:translateY(-36px) scale(.82)}}
    @keyframes dsOutLeft{from{opacity:1;transform:translateX(0)}to{opacity:0;transform:translateX(-32px)}}
    @keyframes dsOutRight{from{opacity:1;transform:translateX(0)}to{opacity:0;transform:translateX(32px)}}
    .ds-out-up{animation:dsOutUp .55s ease forwards}
    .ds-out-left{animation:dsOutLeft .55s ease forwards}
    .ds-out-right{animation:dsOutRight .55s ease forwards}
    /* 日志条目 */
    .ds-log-line{animation:exIn .25s}
  </style>`);
}

/* ---------- 四种模式配置 ---------- */
const DS_MODES = {
  stack: {
    axis:'v',
    title:'栈 Stack · LIFO 后进先出',
    rule:'push / pop · 同一端（栈顶）进出',
    hint:'元素只能在栈顶一端插入和删除：最后进入栈的元素最先弹出（后进先出 LIFO）。',
    stageHint:'栈底在下、封闭；栈顶在上、开放。橙色=刚入栈元素，绿色=已在栈中；只能在栈顶 push / pop。',
    labels:{top:'栈顶', bottom:'栈底'},
    empty:'（空栈）', sizeWord:'大小',
    emptyWarn:'空栈，无法执行 pop()——发生栈下溢（underflow），请先入栈元素。',
    ops:[
      {id:'push', verb:'push', label:'入栈 push', kind:'insert', side:'top', primary:true},
      {id:'pop',  verb:'pop',  label:'出栈 pop', kind:'remove', side:'top', primary:false}
    ]
  },
  queue: {
    axis:'h',
    title:'标准队列 Queue · FIFO 先进先出',
    rule:'enqueue / dequeue · 队尾进、队头出',
    hint:'元素从队尾（右端）入队，从队头（左端）出队：最先进入队列的元素最先离开（先进先出 FIFO）。',
    stageHint:'队头在左、队尾在右。橙色=刚入队元素，绿色=已在队列中；队尾入、队头出（FIFO）。',
    labels:{left:'队头', right:'队尾'},
    empty:'（空队列）', sizeWord:'长度',
    emptyWarn:'空队列，无法执行 dequeue()——发生队列下溢（underflow），请先入队元素。',
    ops:[
      {id:'enq', verb:'enqueue', label:'入队 enqueue', kind:'insert', side:'right', primary:true},
      {id:'deq', verb:'dequeue', label:'出队 dequeue', kind:'remove', side:'left', primary:false}
    ]
  },
  'dual-in': {
    axis:'h',
    title:'双端入 · 一端出队列',
    rule:'左右两端均可入 · 仅左端出队',
    hint:'元素可以从左端或右端入队，但只能从左端出队；先进入左端一侧的元素先离开。',
    stageHint:'左端可入可出，右端仅可入。橙色=刚入队元素，绿色=已在队列中；仅左端出队。',
    labels:{left:'左端', right:'右端'},
    empty:'（空队列）', sizeWord:'大小',
    emptyWarn:'空队列，无法执行左端出队——发生队列下溢（underflow），请先入队元素。',
    ops:[
      {id:'inL', verb:'appendleft', label:'左端入队', kind:'insert', side:'left', primary:true},
      {id:'inR', verb:'append',     label:'右端入队', kind:'insert', side:'right', primary:true},
      {id:'outL',verb:'popleft',    label:'左端出队', kind:'remove', side:'left', primary:false}
    ]
  },
  'dual-out': {
    axis:'h',
    title:'一端入 · 两端出队列',
    rule:'仅右端入 · 左右两端均可出',
    hint:'元素只能从右端入队，但可以从左端或右端出队：入队端固定在右端，出队端开放为两端。',
    stageHint:'右端仅可入，左右两端均可出。橙色=刚入队元素，绿色=已在队列中；可选择左端或右端出队。',
    labels:{left:'左端', right:'右端'},
    empty:'（空队列）', sizeWord:'大小',
    emptyWarn:'空队列，无法执行出队——发生队列下溢（underflow），请先入队元素。',
    ops:[
      {id:'inR', verb:'append',   label:'右端入队', kind:'insert', side:'right', primary:true},
      {id:'outL',verb:'popleft',  label:'左端出队', kind:'remove', side:'left', primary:false},
      {id:'outR',verb:'pop',      label:'右端出队', kind:'remove', side:'right', primary:false}
    ]
  }
};

MOD.stack = {
  name:'栈与队列',
  sub:'栈 LIFO · 队列 FIFO · 双端进出 · 自主操作演示',
  badge:'结构 · 线性表', tag:'栈与队列', cat:'ds',
  color:'#8B5CF6', soft:'#F3EEFF',
  type:'栈 · 队列 · 双端队列 · 后进先出 · 先进先出',
  desc:'在同一个模块中对比演示栈（后进先出 LIFO）与三种队列（先进先出 FIFO、双端入一端出、一端入两端出）：输入元素值后点击按钮自主操作，实时观察元素进出顺序、状态变化与操作日志。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2.5" y="3" width="8" height="3.4" rx="1"/><rect x="2.5" y="8.2" width="8" height="3.4" rx="1"/><rect x="2.5" y="13.4" width="8" height="3.4" rx="1"/><path d="M6.5 17.4v3M6.5 20.4l-1.7-1.7M6.5 20.4l1.7-1.7"/><rect x="14" y="8" width="3.2" height="8" rx="1"/><rect x="18.8" y="8" width="3.2" height="8" rx="1"/><path d="M11.8 12h1.4M21.7 12h-1.4"/></svg>',
  setDesc:'',
  stageDesc:'栈底封闭、栈顶开放；队头在左、队尾在右；橙色=刚操作元素，绿色=已在结构中。',
  initDesc:'在上方输入元素值（数字或字符串），点击「入栈 / 入队」「出栈 / 出队」按钮自主操作；切换顶部标签可在栈与三种队列结构之间切换，切换后自动清空元素。',
  canRegenerate:false,
  currentSub:'stack',
  subTabs:[
    {id:'stack',     label:'栈（LIFO）',        method:'push / pop · 后进先出'},
    {id:'queue',     label:'标准队列（FIFO）',  method:'队尾入 · 队头出 · 先进先出'},
    {id:'dual-in',   label:'双端入 · 一端出',   method:'左右入 · 仅左端出'},
    {id:'dual-out',  label:'一端入 · 两端出',   method:'仅右端入 · 左右出'}
  ],

  theory(){
    return `<div class="rule"><b>栈（Stack）</b>：线性数据结构，只能在一端（<b>栈顶</b>）进行插入和删除，另一端为<b>栈底</b>，栈底封闭。</div>
    <div class="rule"><b>后进先出 LIFO</b>：Last In First Out，最后入栈的元素最先出栈；<code>push(v)</code> 入栈、<code>pop()</code> 出栈；空栈再执行 pop 会发生<b>栈下溢</b>。</div>
    <div class="rule"><b>队列（Queue）</b>：一端（<b>队尾</b>）插入、另一端（<b>队头</b>）删除；<code>enqueue(v)</code> 入队、<code>dequeue()</code> 出队。</div>
    <div class="rule"><b>先进先出 FIFO</b>：First In First Out，最先入队的元素最先出队；空队列再执行 dequeue 会发生<b>队列下溢</b>。</div>
    <div class="rule"><b>双端变体</b>：「双端入 · 一端出」——左右两端均可入队，仅左端出队；「一端入 · 两端出」——仅右端入队，左右两端均可出队。</div>
    <div class="rule">应用：函数调用栈、括号匹配、浏览器后退按钮；排队叫号、广度优先搜索 BFS、消息队列。</div>`;
  },

  /* ================= 设置区：输入框 + 模式按钮组 ================= */
  settingsHTML(){
    const meta=this.meta();
    const btns=meta.ops.map(op=>
      `<button type="button" class="btn ds-op ${op.primary?'primary':''}" data-op="${op.id}">${op.label}</button>`).join('');
    return `<div class="inline-setting"><div class="setting"><label>元素值（数字 / 字符串）</label>
        <input type="text" id="ds-val" placeholder="如 10、A、X" autocomplete="off" style="width:clamp(110px,16vw,170px)"></div></div>
      ${btns}
      <button type="button" class="btn ghost" id="ds-clear">清空</button>`;
  },
  afterSettings(){
    const meta=this.meta();
    const input=$('#ds-val');
    $$('.ds-op').forEach(b=>{ b.onclick=()=>this.doOp(b.dataset.op); });
    if(input){
      input.onkeydown=e=>{
        if(e.key==='Enter'){
          const primary=meta.ops.find(o=>o.primary&&o.kind==='insert')||meta.ops[0];
          this.doOp(primary.id);
        }
      };
    }
    const clr=$('#ds-clear');
    if(clr) clr.onclick=()=>this.resetStage();
    if(input) input.focus();
  },

  /* ================= 自有状态管理 ================= */
  meta(){ return DS_MODES[this.currentSub]||DS_MODES.stack; },
  initState(){ this._data=[]; this._log=[]; this._last=null; this._warn=null; this._stepNo=0; },
  resetState(){ this.initState(); },
  resetStage(){
    this.resetState();
    this.renderStage();
    const ex=$('#explainBox'); if(ex) ex.innerHTML=this.initDesc;
  },
  normalizeVal(raw){
    if(/^-?\d+$/.test(raw)) return parseInt(raw,10);
    if(/^-?\d+\.\d+$/.test(raw)) return parseFloat(raw);
    return raw;
  },

  /* ---------- 核心操作 ---------- */
  doOp(opId){
    const meta=this.meta();
    const op=meta.ops.find(o=>o.id===opId);
    if(!op) return;
    const input=$('#ds-val');
    const raw=input?input.value.trim():'';
    this._warn=null;

    if(op.kind==='insert'){
      if(!raw){
        this._warn='请先在输入框中填写要插入的元素值（数字或字符串），再点击入栈 / 入队按钮。';
        this.renderStage(); this.explainWarn(); return;
      }
      const val=this.normalizeVal(raw);
      const before=this._data.length;
      if(op.side==='left') this._data.unshift(val); else this._data.push(val);
      const after=this._data.length;
      this._stepNo++;
      this._last={op:op,kind:'insert',side:op.side,val:val,before:before,after:after,seq:this._stepNo};
      this._log.push({html:`<b>${op.verb}(${esc(val)})</b>：${meta.labels[op.side]}新增 ${esc(val)} · ${meta.sizeWord} ${before} → ${after}（+1）`});
      this.renderStage(); this.explainOp();
      if(input){ input.value=''; input.focus(); }
    }else{
      if(this._data.length===0){
        this._warn=meta.emptyWarn;
        this.renderStage(); this.explainWarn(); return;
      }
      const val=op.side==='left'?this._data.shift():this._data.pop();
      const before=this._data.length+1, after=this._data.length;
      this._stepNo++;
      this._last={op:op,kind:'remove',side:op.side,val:val,before:before,after:after,seq:this._stepNo};
      this._log.push({html:`<b>${op.verb}()</b>：返回${meta.labels[op.side]}元素 ${esc(val)} · ${meta.sizeWord} ${before} → ${after}（-1）`});
      this.renderStage(); this.explainOp();
      if(input) input.focus();
    }
  },

  /* ---------- 说明区 ---------- */
  explainOp(){
    const l=this._last; if(!l) return;
    const meta=this.meta();
    let txt;
    if(l.kind==='insert'){
      txt=`第 ${l.seq} 步：<b class="c-op">${l.op.verb}(${esc(l.val)})</b>，在${meta.labels[l.side]}新增 <b>${esc(l.val)}</b>；${meta.sizeWord} <b>${l.before}</b> → <b>${l.after}</b>（+1）。`;
    }else{
      txt=`第 ${l.seq} 步：<b class="c-op">${l.op.verb}()</b>，从${meta.labels[l.side]}移除 <b class="c-op">${esc(l.val)}</b>；${meta.sizeWord} <b>${l.before}</b> → <b>${l.after}</b>（-1）。`;
    }
    const ex=$('#explainBox'); if(ex) ex.innerHTML=txt;
  },
  explainWarn(){
    const ex=$('#explainBox');
    if(ex) ex.innerHTML=`<div class="err"><div class="err-t">无法执行操作</div>${esc(this._warn||'')}</div>`;
  },

  /* ================= 画布渲染 ================= */
  renderStage(){
    const meta=this.meta();
    const N=this._data.length;
    const sd=$('#stageDesc'); if(sd) sd.textContent=meta.stageHint;
    const stateTxt=N===0?(meta.axis==='v'?'空栈 · 准备就绪':'空队列 · 准备就绪'):`${meta.sizeWord} ${N}`;
    const structureHtml=meta.axis==='v'?this.stackBox(meta):this.queueBox(meta);
    const exprHtml=this.exprLine(meta);
    const logHtml=this.logHtml();
    const warnHtml=this._warn?`<div class="err" style="margin-top:8px"><div class="err-t">提示</div>${esc(this._warn)}</div>`:'';

    $('#stage').innerHTML=`<div class="stage-hd"><span>${meta.title}</span><span class="state">${stateTxt}</span></div>
      <div class="step-title" style="margin-bottom:6px"><span class="st-dot"></span>${meta.rule}</div>
      <div class="annot" style="margin:0 0 8px;line-height:1.7">${meta.hint}</div>
      ${structureHtml}
      <div style="display:flex;gap:clamp(10px,1.6vw,14px);flex-wrap:wrap;align-items:stretch;margin-top:10px">
        <div style="flex:1 1 280px;min-width:230px">
          <div class="zone-label" style="margin-bottom:0">操作日志（共 ${this._stepNo} 步）</div>
          <div class="ds-log" style="margin-top:2px;line-height:1.8">${logHtml}</div>
        </div>
        ${exprHtml?`<div style="flex:1 1 260px;min-width:210px">
          <div class="zone-label" style="margin-bottom:0">最近操作表达式</div>
          <div class="exp-inline" style="margin-top:2px;min-height:34px">${exprHtml}</div>
        </div>`:''}
      </div>
      ${warnHtml}`;
  },

  /* ---------- 栈结构（纵向 · 底封闭顶开放） ---------- */
  stackBox(meta){
    const size=this._data.length;
    let ghost='';
    if(this._last&&this._last.kind==='remove'&&this._last.side==='top'){
      ghost=`<div class="ds-ghost-v">
        <span class="ghost-tag">弹出 ↑</span>
        <div class="chip o ds-chip-v ds-out-up">${esc(this._last.val)}</div>
      </div>`;
    }
    let chips;
    if(size===0){
      chips=`<div class="ds-empty">${meta.empty}</div>`;
    }else{
      chips=this._data.map((v,i)=>{
        const isTop=i===size-1, isBottom=i===0;
        const isCur=this._last&&this._last.kind==='insert'&&this._last.side==='top'&&isTop;
        const cls=isCur?'o':'g', anim=isCur?' write-in':'';
        const topTag=isTop?`<span class="side-tag side-top">${meta.labels.top} · 操作端</span>`:'';
        const botTag=isBottom&&size>1?`<span class="side-tag side-bottom">${meta.labels.bottom}</span>`:'';
        return `<div class="chip ds-chip-v ${cls}${anim}">${esc(v)}${topTag}${botTag}</div>`;
      }).join('');
    }
    return `<div style="flex:1 1 280px;min-width:240px">
      <div class="zone-label" style="margin-bottom:4px">${meta.labels.top}在上（开放操作端） · ${meta.labels.bottom}在下（封闭端）</div>
      ${ghost?`<div class="ds-ghost-v-row">${ghost}</div>`:'<div class="ds-ghost-v-row"></div>'}
      <div class="ds-stack-area">${chips}</div>
    </div>`;
  },

  /* ---------- 队列结构（横向 · 左端 → 右端） ---------- */
  queueBox(meta){
    let ghostL='', ghostR='';
    if(this._last&&this._last.kind==='remove'){
      if(this._last.side==='left') ghostL=`<div class="ds-ghost-h"><div class="chip o ds-chip-h ds-out-left">${esc(this._last.val)}</div><span class="ghost-tag">← 离开</span></div>`;
      if(this._last.side==='right') ghostR=`<div class="ds-ghost-h"><span class="ghost-tag">离开 →</span><div class="chip o ds-chip-h ds-out-right">${esc(this._last.val)}</div></div>`;
    }
    let chips;
    if(this._data.length===0){
      chips=`<div class="ds-empty">${meta.empty}</div>`;
    }else{
      const size=this._data.length;
      chips=this._data.map((v,i)=>{
        const isHead=i===0, isTail=i===size-1;
        let isCur=false;
        if(this._last&&this._last.kind==='insert'){
          if(this._last.side==='left'&&isHead) isCur=true;
          if(this._last.side==='right'&&isTail) isCur=true;
        }
        const cls=isCur?'o':'g', anim=isCur?' write-in':'';
        return `<div class="chip ds-chip-h ${cls}${anim}">${esc(v)}</div>`;
      }).join('');
    }
    return `<div class="ds-q-row">
      <div class="ds-ghost-slot">${ghostL}</div>
      <div style="flex:1;min-width:0">
        <div class="ds-q-labels"><span>${meta.labels.left}</span><span>${meta.labels.right}</span></div>
        <div class="ds-queue-area">${chips}</div>
      </div>
      <div class="ds-ghost-slot">${ghostR}</div>
    </div>`;
  },

  /* ---------- 最近操作表达式 ---------- */
  exprLine(meta){
    const l=this._last; if(!l) return '';
    if(l.kind==='insert'){
      const api=`${l.op.verb}(${esc(l.val)})`;
      return `<span class="term cur">${this.typeChars(api)}</span><span class="op">=</span><span class="term">${meta.labels[l.side]} ← ${esc(l.val)}</span><span class="op">·</span><span class="term done">${meta.sizeWord} ${l.before} → ${l.after}（+1）</span><span class="caret pen"></span>`;
    }
    return `<span class="term cur">${this.typeChars(l.op.verb+'()')}</span><span class="op">=</span><span class="term cur">返回 ${esc(l.val)}（${meta.labels[l.side]}）</span><span class="op">·</span><span class="term done">${meta.sizeWord} ${l.before} → ${l.after}（-1）</span><span class="caret pen"></span>`;
  },

  /* ---------- 操作日志 ---------- */
  logHtml(){
    if(!this._log.length) return '<span style="opacity:.45">（尚未开始操作：先在上方输入元素值，再点击入栈 / 入队或出栈 / 出队按钮）</span>';
    return this._log.map((e,i)=>{
      const latest=i===this._log.length-1;
      return `<div class="ds-log-line" style="opacity:${latest?1:.7}">${e.html}</div>`;
    }).join('');
  },

  /* ---------- 逐字出现动画辅助 ---------- */
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.06}s">${esc(ch)}</span>`).join('');
  },

  /* ================= Engine 钩子 ================= */
  initStage(){ this.resetState(); this.renderStage(); },
  generate(){
    this.resetState();
    this.steps=[{desc:this.initDesc, draw:()=>this.renderStage()}];
  }
};
