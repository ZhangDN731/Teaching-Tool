/* ================================================================
 * 字符编码（ASCII / BCD / 区位码→国标码→机内码 / 字形码容量）
 * 子页面：ascii · bcd · hanzi · font
 * 重构要点：
 *  1. 步骤累加渲染：每一步 draw() 重绘“已完成 + 当前”完整状态，上一步可回退
 *  2. BCD 固定骨架 + 累积补位，消除整屏重置导致的抖动/跳脱
 *  3. 区位/国标/机内三码卡片同屏，当前步骤高亮
 *  4. 大屏幕字号放大：计算过程 clamp(16px,2.2vw,22px)，关键数字 clamp(18px,2.8vw,28px)
 * ================================================================ */
(function injectEncCSS(){
  if(document.getElementById('enc-mod-css')) return;
  document.head.insertAdjacentHTML('beforeend', `
  <style id="enc-mod-css">
    /* ---------- 字符编码模块自定义样式 ---------- */
    .enc-flow{display:flex;flex-wrap:wrap;align-items:stretch;justify-content:center;gap:clamp(6px,1vw,10px);margin:clamp(6px,1vw,10px) 0}
    .enc-flow .enc-node{flex:1 1 130px;max-width:230px;display:flex;flex-direction:column;gap:5px;align-items:center;background:#fff;border:1.5px solid var(--border);border-radius:14px;padding:clamp(8px,1.3vw,12px) clamp(10px,1.5vw,14px);transition:all .2s}
    .enc-flow .enc-node .enc-nl{font-family:var(--font);font-size:clamp(11px,1.4vw,13px);color:var(--text-3);font-weight:700;letter-spacing:.3px}
    .enc-flow .enc-node .enc-nv{font-family:var(--mono);font-size:clamp(16px,2.4vw,24px);font-weight:800;color:var(--text-2);letter-spacing:1px}
    .enc-flow .enc-node.done{border-color:var(--green);background:var(--green-soft)}
    .enc-flow .enc-node.done .enc-nv{color:#15803D}
    .enc-flow .enc-node.cur{border-color:var(--orange);background:#FFF4DB;box-shadow:0 4px 14px rgba(245,158,11,.18)}
    .enc-flow .enc-node.cur .enc-nv{color:#B45309}
    .enc-flow .enc-node.locked{opacity:.55}
    .enc-flow .enc-node.locked .enc-nv{color:#CBD5E1}
    .enc-flow .enc-arrow{display:flex;align-items:center;color:#94A3B8;font-size:clamp(16px,2.2vw,22px);font-weight:700}

    .enc-card{background:#fff;border:1.5px solid var(--border);border-radius:var(--radius-lg);padding:clamp(12px,2vw,20px);margin-bottom:clamp(10px,2vw,16px);position:relative;transition:all .2s}
    .enc-card.done{border-color:var(--green);background:linear-gradient(135deg,#F0FDF4,#fff 60%)}
    .enc-card.cur{border-color:var(--orange);background:linear-gradient(135deg,#FFF8EE,#fff 60%);box-shadow:0 6px 20px rgba(245,158,11,.12)}
    .enc-card.locked{opacity:.55;background:#F8FAFC;border-style:dashed}
    .enc-card .enc-card-hd{display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap;margin-bottom:clamp(8px,1.4vw,12px)}
    .enc-card .enc-card-title{font-size:clamp(14px,1.9vw,18px);font-weight:700;color:var(--text)}
    .enc-card .enc-card-sub{font-size:clamp(11px,1.4vw,13px);color:var(--text-3);font-weight:600;margin-top:2px}
    .enc-badge{display:inline-flex;align-items:center;gap:5px;border-radius:999px;padding:3px 12px;font-size:clamp(10px,1.3vw,12.5px);font-weight:700;font-family:var(--mono);white-space:nowrap}
    .enc-badge.done{background:var(--green-soft);color:#15803D}
    .enc-badge.cur{background:var(--orange);color:#fff}
    .enc-badge.locked{background:#F1F5F9;color:#94A3B8}

    .enc-big{font-family:var(--mono);font-size:clamp(20px,3.4vw,32px);font-weight:800;text-align:center;letter-spacing:2px;color:var(--text);margin:clamp(6px,1.2vw,10px) 0;line-height:1.6}
    .enc-card.done .enc-big{color:#15803D}
    .enc-card.cur .enc-big{color:#B45309}

    .enc-calc{font-family:var(--mono);font-size:clamp(16px,2.2vw,22px);font-weight:700;line-height:2.1;text-align:center;color:var(--text)}
    .enc-calc .enc-op{color:#94A3B8;margin:0 6px;font-weight:400}
    .enc-calc .enc-eq{color:var(--accent);margin:0 6px}
    .enc-calc .enc-res{color:#B45309;background:#FFF4DB;border:1.5px solid var(--orange);border-radius:8px;padding:2px 12px}
    .enc-card.done .enc-calc .enc-res{color:#15803D;background:#E8FBF0;border-color:var(--green)}
    .enc-calc-row{display:flex;flex-wrap:wrap;align-items:center;justify-content:center;gap:clamp(6px,1vw,10px);margin:clamp(6px,1.1vw,10px) 0}
    .enc-sub{font-size:clamp(12px,1.6vw,15px);color:var(--text-3);text-align:center;font-family:var(--font);margin-top:4px;line-height:1.7}

    /* 二进制大网格（8 位一组） */
    .enc-bitgrid{display:grid;grid-template-columns:repeat(8,1fr);gap:clamp(4px,.9vw,8px);max-width:580px;margin:clamp(8px,1.4vw,12px) auto}
    .enc-bitgrid .enc-bit{background:#fff;border:1.5px solid var(--border);border-radius:10px;text-align:center;font-family:var(--mono);font-size:clamp(18px,2.8vw,28px);font-weight:800;padding:clamp(8px,1.4vw,12px) 0;color:var(--text-2)}
    .enc-bitgrid .enc-bit.one{color:var(--accent);background:var(--accent-soft);border-color:var(--accent)}
    .enc-bitgrid .enc-bit.zero{color:#CBD5E1}
    .enc-bitgrid.weight .enc-bit{font-size:clamp(10px,1.4vw,13px);font-weight:700;color:var(--text-3);background:var(--surface-3);border-style:dashed;padding:4px 0}
    .enc-bitgroup{display:flex;flex-wrap:wrap;gap:clamp(10px,1.8vw,16px);justify-content:center;margin:clamp(8px,1.4vw,12px) 0}
    .enc-bitgroup .enc-b4{display:grid;grid-template-columns:repeat(4,1fr);gap:clamp(4px,.8vw,6px);background:#F8FAFC;border:1px solid var(--border);border-radius:12px;padding:clamp(6px,1vw,10px);min-width:150px}
    .enc-bitgroup .enc-b4 .enc-bit{font-size:clamp(16px,2.4vw,24px);padding:clamp(6px,1vw,10px) 0;border-radius:8px}
    .enc-b4-label{text-align:center;font-family:var(--font);font-size:clamp(11px,1.4vw,13px);font-weight:700;color:var(--text-3);margin-bottom:6px}

    /* 三码同屏卡片网格 */
    .enc-tri{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:clamp(10px,1.8vw,16px);margin:clamp(6px,1.2vw,10px) 0}
    .enc-tri .enc-card{margin-bottom:0}
    @media (min-width:768px){ .enc-tri{grid-template-columns:repeat(3,1fr)} }

    .enc-chip-row{display:flex;flex-wrap:wrap;gap:clamp(6px,1vw,10px);justify-content:center;margin:clamp(6px,1.1vw,10px) 0}
    .enc-chip{font-family:var(--mono);font-size:clamp(15px,2.1vw,21px);font-weight:800;border-radius:10px;padding:clamp(6px,1vw,10px) clamp(12px,1.8vw,18px);border:1.5px solid var(--border);background:#fff;color:var(--text-2)}
    .enc-chip.src{border-color:var(--accent);background:var(--accent-soft);color:var(--accent)}
    .enc-chip.done{border-color:var(--green);background:var(--green-soft);color:#15803D}
    .enc-chip.cur{border-color:var(--orange);background:#FFF4DB;color:#B45309}
    .enc-chip.empty{color:#CBD5E1;border-style:dashed;background:#F8FAFC}

    .enc-placeholder{text-align:center;color:#94A3B8;font-size:clamp(12px,1.6vw,15px);padding:clamp(12px,2vw,18px) 0;font-family:var(--font)}
    .enc-result-strip{font-family:var(--mono);font-size:clamp(15px,2.3vw,22px);font-weight:800;color:var(--accent);background:linear-gradient(135deg,#EEF3FF,#E0F2FE);border-radius:12px;padding:clamp(10px,1.6vw,14px) clamp(12px,2vw,18px);text-align:center;margin-top:clamp(8px,1.4vw,12px);border:1px solid #C7D2FE;line-height:1.9;min-height:clamp(46px,7vw,64px);display:flex;align-items:center;justify-content:center;flex-wrap:wrap;gap:clamp(6px,1vw,10px)}
    .enc-result-strip .enc-op{color:#94A3B8;font-weight:400}
    .enc-result-strip.locked{background:#F8FAFC;border:1px dashed var(--border);color:#94A3B8;font-family:var(--font);font-size:clamp(12px,1.6vw,15px);font-weight:600}
    .enc-mini{font-family:var(--mono);font-size:clamp(14px,2vw,18px);color:var(--text-2);background:#F8FAFC;border:1px solid var(--border);border-radius:12px;padding:clamp(8px,1.3vw,12px);margin:clamp(6px,1.1vw,10px) 0;line-height:2}
    .enc-mini .enc-h{color:var(--text-3);font-size:clamp(10px,1.3vw,12.5px);font-family:var(--font);font-weight:700;margin-bottom:4px}
    .enc-note{display:inline-block;background:#F8FAFC;border:1px dashed var(--border);border-radius:10px;padding:6px 12px;font-size:clamp(12px,1.6vw,15px);color:var(--text-3);font-family:var(--font);margin:6px auto}
    .enc-table{font-family:var(--mono);font-size:clamp(15px,2.2vw,22px);margin:clamp(6px,1.2vw,10px) auto;max-width:560px}
    .enc-table th,.enc-table td{padding:clamp(7px,1.1vw,11px) clamp(8px,1.4vw,13px);font-size:inherit}
    .enc-table th.cur{background:#FFF4DB;color:#B45309;font-weight:800}
  </style>`);
})();

MOD.enc = {
  name:'字符编码', sub:'ASCII · BCD · 区位码国标码机内码 · 字形码',
  badge:'编码 · 转换', tag:'编码', cat:'bit',
  color:'#8B5CF6', soft:'#F3EEFF', type:'ASCII · BCD · 区位码 · 国标码 · 机内码 · 字形码',
  desc:'涵盖 ASCII 码值查询、BCD 编码、汉字区位码→国标码→机内码转换、字形码容量计算；步骤累加演示，转换链路与三码结果同屏可见。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6h16M4 12h10M4 18h7M16 15l5 5M16 20l5-5"/></svg>',
  setDesc:'',
  stageDesc:'橙色=当前计算项，绿色=已完成结果；已完成步骤保留同屏，新步骤逐步累加显示。',
  initDesc:'点击「自动演示」或「下一步」，观看字符编码的逐步推导过程。',
  currentSub:'ascii',
  subTabs:[
    {id:'ascii',label:'ASCII码值',method:'字符→十进制→8位二进制→十六进制'},
    {id:'bcd',label:'BCD编码',method:'逐位4位二进制·累积拼接'},
    {id:'hanzi',label:'区位→国标→机内',method:'三码同屏·+2020H·+8080H'},
    {id:'font',label:'字形码容量',method:'点数→字节→总容量'}
  ],
  theory(){
    return `<div class="rule"><b>ASCII</b>：美国信息交换标准代码，用 1 字节（8 位二进制）表示 128 个字符，含英文、数字、符号和控制符；同一字符有十进制、二进制、十六进制三种写法。</div>
    <div class="rule"><b>BCD 码</b>：二进制编码的十进制，用 4 位二进制（8421 权）表示 1 位十进制数（0~9），常用于数字显示；各位二进制依次拼接即得 BCD 码。</div>
    <div class="rule"><b>区位码</b>：汉字在 GB2312 中的位置，由区号（1~94）和位号（1~94）组成，是十进制表示。</div>
    <div class="rule"><b>国标码</b> = 区位码（区、位分别）+ 2020H。即区号转十六进制后 +20H，位号 +20H。</div>
    <div class="rule"><b>机内码</b> = 国标码 + 8080H（即区、位分别 +80H），计算机内部存储和处理汉字时使用，占 2 字节。</div>
    <div class="rule"><b>字形码</b>：汉字字形的点阵数据。16×16 点阵 = 16×16÷8 = 32 字节/字；24×24 = 72 字节/字。</div>
    <div class="rule">口诀：<b>区位 + 2020H = 国标，国标 + 8080H = 机内</b>（等价于区位 + A0A0H = 机内）。</div>`;
  },
  settingsHTML(){
    const sub=this.currentSub;
    if(sub==='ascii'){
      const curChar=$('#e-ascii-custom')?$('#e-ascii-custom').value:(this._customChar||'');
      return `<div class="inline-setting"><div class="setting"><label>输入字符</label>
        <input type="text" id="e-ascii-custom" value="${esc(curChar)}" placeholder="如 A、a、5、?" maxlength="1" style="width:clamp(60px,10vw,80px)"></div></div>
        <div class="inline-setting"><div class="setting"><label>或随机类型</label>
        <select id="e-ascii">
          <option value="upper">大写字母</option>
          <option value="lower">小写字母</option>
          <option value="digit">数字</option>
          <option value="symbol">符号</option>
        </select></div></div>`;
    }
    if(sub==='bcd'){
      const curNum=$('#e-bcd-custom')?$('#e-bcd-custom').value:(this._customNum||'');
      return `<div class="inline-setting"><div class="setting"><label>输入十进制数</label>
        <input type="text" id="e-bcd-custom" value="${esc(curNum)}" placeholder="如 256" maxlength="4" style="width:clamp(70px,12vw,90px)"></div></div>`;
    }
    if(sub==='hanzi'){
      const curQ=$('#e-hanzi-q')?$('#e-hanzi-q').value:(this._customQ||'');
      const curW=$('#e-hanzi-w')?$('#e-hanzi-w').value:(this._customW||'');
      return `<div class="inline-setting"><div class="setting"><label>区号</label>
        <input type="text" id="e-hanzi-q" value="${esc(curQ)}" placeholder="1~94" maxlength="2" style="width:clamp(50px,8vw,65px)"></div></div>
        <div class="inline-setting"><div class="setting"><label>位号</label>
        <input type="text" id="e-hanzi-w" value="${esc(curW)}" placeholder="1~94" maxlength="2" style="width:clamp(50px,8vw,65px)"></div></div>`;
    }
    if(sub==='font'){
      const curS=$('#e-font-size')?$('#e-font-size').value:String(this._fontSize||'16');
      const curC=$('#e-font-count')?$('#e-font-count').value:String(this._fontCount||'1000');
      return `<div class="inline-setting"><div class="setting"><label>点阵规格</label>
        <select id="e-font-size">
          <option value="16" ${curS==='16'?'selected':''}>16 × 16</option>
          <option value="24" ${curS==='24'?'selected':''}>24 × 24</option>
          <option value="32" ${curS==='32'?'selected':''}>32 × 32</option>
        </select></div></div>
      <div class="inline-setting"><div class="setting"><label>汉字数量</label>
        <input type="text" id="e-font-count" value="${esc(curC)}" placeholder="如 1000" maxlength="5" style="width:clamp(60px,10vw,80px)"></div></div>`;
    }
    return '';
  },
  afterSettings(){
    const bind=id=>{ const el=$('#'+id); if(el) el.onchange=()=>Engine.reload(); };
    const bindInput=id=>{ const el=$('#'+id); if(el) el.onchange=()=>Engine.reload(); };
    ['e-ascii','e-font-size'].forEach(bind);
    ['e-ascii-custom','e-bcd-custom','e-hanzi-q','e-hanzi-w','e-font-count'].forEach(bindInput);
  },
  /* ---------- 内部辅助 ---------- */
  /* j < cur → done；j === cur → cur；j > cur → locked */
  _state(j,cur){ return j<cur?'done':j===cur?'cur':'locked'; },
  _badgeText(st){ return st==='done'?'已完成':st==='cur'?'当前计算':'待计算'; },
  /* 链路节点（顶部流程条） */
  _flowNode(sub,value,st){
    return `<div class="enc-node ${st}"><span class="enc-nl">${sub}</span><span class="enc-nv">${value||'—'}</span></div>`;
  },
  /* 步骤卡片外壳；locked 时渲染占位内容。
     注意：入场动画只挂在徽章等小装饰上，整卡不依赖动画透明度，避免后台冻结时整卡不可见。 */
  _card(st,title,sub,inner){
    return `<div class="enc-card ${st}">
      <div class="enc-card-hd">
        <div><div class="enc-card-title">${title}</div>${sub?`<div class="enc-card-sub">${sub}</div>`:''}</div>
        <span class="enc-badge ${st}">${this._badgeText(st)}</span>
      </div>
      ${inner}
    </div>`;
  },
  _lockedInner(text){
    return `<div class="enc-placeholder">${text||'等待本步骤计算…'}</div>`;
  },
  /* ========== ① ASCII：字符 → 十进制 → 8位二进制（计算机存储）→ 十六进制 ========== */
  genAscii(){
    const customEl=$('#e-ascii-custom');
    const custom=customEl?customEl.value.trim():'';
    let ch;
    if(custom && custom.length>0){
      ch=custom[0]; this._customChar=ch;
    }else{
      const el=$('#e-ascii');
      const group=el?el.value:'upper';
      const pools={upper:'ABCDEFGHIJKLMNOPQRSTUVWXYZ',lower:'abcdefghijklmnopqrstuvwxyz',digit:'0123456789',symbol:'!?#@%&*()/-'};
      ch=pick([...(pools[group]||pools.upper)]);
    }
    let code=ch.charCodeAt(0);
    /* ASCII 范围 0~127：超出范围（如中文等多字节字符）回退到所选随机分组，保证 8 位网格正确 */
    if(code>127){
      const el2=$('#e-ascii');
      const grp=el2?el2.value:'upper';
      const pools={upper:'ABCDEFGHIJKLMNOPQRSTUVWXYZ',lower:'abcdefghijklmnopqrstuvwxyz',digit:'0123456789',symbol:'!?#@%&*()/-'};
      ch=pick([...(pools[grp]||pools.upper)]);
      this._customChar='';
      code=ch.charCodeAt(0);
    }
    const bin=code.toString(2).padStart(8,'0');
    const hex=code.toString(16).toUpperCase().padStart(2,'0');
    const highNib=bin.slice(0,4), lowNib=bin.slice(4);
    const weights=[128,64,32,16,8,4,2,1];
    /* 除 2 取余竖式 */
    const rows=[]; let temp=code;
    if(temp===0) rows.push({n:0,q:0,r:'0'});
    while(temp>0){ const q=Math.floor(temp/2),r=temp%2; rows.push({n:temp,q,r}); temp=q; }
    /* 码表片段（前后各 3 个） */
    const tableStart=Math.max(32,code-3), tableEnd=Math.min(126,code+3);
    const tableRows=[];
    for(let c=tableStart;c<=tableEnd;c++) tableRows.push({ch:String.fromCharCode(c),code:c,isCur:c===code});

    const stateLabels=['输入字符','查表 · 十进制','计算机存储 · 二进制','十六进制 · 完成'];
    const steps=[];

    /* 卡片渲染器：cur 为当前步骤序号（0..3），已完成卡片保留同屏 */
    const renderCards=(cur)=>{
      /* 卡片 0：字符 */
      const c0=this._card(this._state(0,cur),'① 输入字符','待查询的字符',
        `<div class="enc-big" style="font-size:clamp(28px,5vw,42px)">“${esc(ch)}”</div>
         <div class="enc-sub">沿 字符 → 十进制 → 二进制 → 十六进制 链路逐步推导</div>`);
      /* 卡片 1：十进制码值 */
      let c1;
      if(cur>=1){
        const tableHtml=`<table class="exp-table enc-table">
          <tr><th>字符</th>${tableRows.map(r=>`<th class="${r.isCur?'cur':''}">${r.ch===' '?'␣':esc(r.ch)}</th>`).join('')}</tr>
          <tr><th>十进制</th>${tableRows.map(r=>`<td class="${r.isCur?'cur':r.code<code?'done':''}">${r.code}</td>`).join('')}</tr>
        </table>`;
        c1=this._card(this._state(1,cur),'② 查 ASCII 码表 · 十进制','字符对应的十进制码值',
          `<div class="enc-big">${code}<sub style="font-size:.45em;color:var(--text-3)">十进制</sub></div>
           ${tableHtml}
           <div style="text-align:center"><span class="enc-note">常见对照：A=65，a=97，0=48；本字符查表得 <b style="color:var(--accent)">${code}</b></span></div>`);
      }else{
        c1=this._card('locked','② 查 ASCII 码表 · 十进制','字符对应的十进制码值',this._lockedInner());
      }
      /* 卡片 2：8 位二进制（计算机存储） */
      let c2;
      if(cur>=2){
        const weightRow=weights.map(w=>`<div class="enc-bit">${w}</div>`).join('');
        const bitRow=bin.split('').map(b=>`<div class="enc-bit ${b==='1'?'one':'zero'}">${b}</div>`).join('');
        const divRows=rows.map(rd=>`<div>${rd.n} ÷ 2 = ${rd.q} <span style="color:var(--orange);font-weight:700">余 ${rd.r}</span></div>`).join('');
        c2=this._card(this._state(2,cur),'③ 8 位二进制 · 计算机存储','1 字节 = 8 位二进制',
          `<div class="enc-bitgrid weight">${weightRow}</div>
           <div class="enc-bitgrid">${bitRow}</div>
           <div class="enc-mini"><div class="enc-h">除 2 取余（余数从下往上逆序读出）</div>${divRows}</div>
           <div class="enc-big">${bin}<sub style="font-size:.45em;color:var(--text-3)">二进制 · 1 字节</sub></div>
           <div class="enc-sub">计算机内部以这 8 位二进制存储字符 “${esc(ch)}”</div>`);
      }else{
        c2=this._card('locked','③ 8 位二进制 · 计算机存储','1 字节 = 8 位二进制',this._lockedInner());
      }
      /* 卡片 3：十六进制 */
      let c3;
      if(cur>=3){
        c3=this._card('done','④ 转换为十六进制','4 位二进制 = 1 位十六进制',
          `<div class="enc-bitgroup">
             <div><div class="enc-b4-label">高 4 位</div><div class="enc-b4">${highNib.split('').map(b=>`<div class="enc-bit ${b==='1'?'one':'zero'}">${b}</div>`).join('')}</div></div>
             <div><div class="enc-b4-label">低 4 位</div><div class="enc-b4">${lowNib.split('').map(b=>`<div class="enc-bit ${b==='1'?'one':'zero'}">${b}</div>`).join('')}</div></div>
           </div>
           <div class="enc-calc">${highNib} 组 → <span class="enc-res">${hex[0]}</span> <span class="enc-op">·</span> ${lowNib} 组 → <span class="enc-res">${hex[1]}</span></div>
           <div class="enc-big">${hex}H<sub style="font-size:.45em;color:var(--text-3)">十六进制</sub></div>`);
      }else{
        c3=this._card('locked','④ 转换为十六进制','4 位二进制 = 1 位十六进制',this._lockedInner());
      }
      return c0+c1+c2+c3;
    };

    /* 步骤 0：字符 */
    steps.push({desc:`写出待查询字符 <b class="c-op">“${esc(ch)}”</b>，沿 <b>字符 → 十进制 → 二进制 → 十六进制</b> 链路逐步推导。`,
      draw:()=>{
        const cur=0;
        const flow=[this._flowNode('字符',`“${esc(ch)}”`,this._state(0,cur)),
          `<span class="enc-arrow">→</span>`,
          this._flowNode('十进制','',this._state(1,cur)),`<span class="enc-arrow">→</span>`,
          this._flowNode('8位二进制','',this._state(2,cur)),`<span class="enc-arrow">→</span>`,
          this._flowNode('十六进制','',this._state(3,cur))].join('');
        $('#stage').innerHTML=`<div class="stage-hd"><span>字符 → ASCII 码值</span><span class="state">${stateLabels[cur]}</span></div>
          <div class="enc-flow">${flow}</div>
          ${renderCards(cur)}
          <div class="enc-result-strip locked">完整转换链路将在最后一步汇总显示</div>`;
      }});
    /* 步骤 1：十进制 */
    steps.push({desc:`查 ASCII 码表：字符 <b class="c-op">${esc(ch)}</b> 对应的十进制码值 = <b class="c-op">${code}</b>。`,
      draw:()=>{
        const cur=1;
        const flow=[this._flowNode('字符',`“${esc(ch)}”`,this._state(0,cur)),
          `<span class="enc-arrow">→</span>`,
          this._flowNode('十进制',code,this._state(1,cur)),`<span class="enc-arrow">→</span>`,
          this._flowNode('8位二进制','',this._state(2,cur)),`<span class="enc-arrow">→</span>`,
          this._flowNode('十六进制','',this._state(3,cur))].join('');
        $('#stage').innerHTML=`<div class="stage-hd"><span>字符 → ASCII 码值</span><span class="state">${stateLabels[cur]}</span></div>
          <div class="enc-flow">${flow}</div>
          ${renderCards(cur)}
          <div class="enc-result-strip locked">完整转换链路将在最后一步汇总显示</div>`;
      }});
    /* 步骤 2：8 位二进制（计算机存储） */
    steps.push({desc:`十进制 <b class="c-op">${code}</b> 连续除 2 取余、余数<b>从下往上逆序</b>读出，补足 8 位得 <b class="c-op">${bin}</b>；这就是计算机中存储该字符的二进制形式（1 字节）。`,
      draw:()=>{
        const cur=2;
        const flow=[this._flowNode('字符',`“${esc(ch)}”`,this._state(0,cur)),
          `<span class="enc-arrow">→</span>`,
          this._flowNode('十进制',code,this._state(1,cur)),`<span class="enc-arrow">→</span>`,
          this._flowNode('8位二进制',bin,this._state(2,cur)),`<span class="enc-arrow">→</span>`,
          this._flowNode('十六进制','',this._state(3,cur))].join('');
        $('#stage').innerHTML=`<div class="stage-hd"><span>字符 → ASCII 码值</span><span class="state">${stateLabels[cur]}</span></div>
          <div class="enc-flow">${flow}</div>
          ${renderCards(cur)}
          <div class="enc-result-strip locked">完整转换链路将在最后一步汇总显示</div>`;
      }});
    /* 步骤 3：十六进制 + 完整链路 */
    steps.push({desc:`二进制按 4 位一组转换为十六进制：${highNib}→<b class="c-op">${hex[0]}</b>，${lowNib}→<b class="c-op">${hex[1]}</b>，得 <b class="c-op">${hex}H</b>。完整链路：“${esc(ch)}” = ${code}<sub>10</sub> = ${bin}<sub>2</sub> = ${hex}H。`,
      draw:()=>{
        const cur=3;
        const flow=[this._flowNode('字符',`“${esc(ch)}”`,this._state(0,cur)),
          `<span class="enc-arrow">→</span>`,
          this._flowNode('十进制',code,this._state(1,cur)),`<span class="enc-arrow">→</span>`,
          this._flowNode('8位二进制',bin,this._state(2,cur)),`<span class="enc-arrow">→</span>`,
          this._flowNode('十六进制',`${hex}H`,this._state(3,cur))].join('');
        $('#stage').innerHTML=`<div class="stage-hd"><span>字符 → ASCII 码值</span><span class="state">${stateLabels[cur]}</span></div>
          <div class="enc-flow">${flow}</div>
          ${renderCards(cur)}
          <div class="enc-result-strip">“${esc(ch)}” <span class="enc-op">=</span> ${code}<sub>10</sub> <span class="enc-op">=</span> ${bin}<sub>2</sub> <span class="enc-op">=</span> ${hex}H <span class="enc-op">（1 字节）</span></div>`;
      }});
    return steps;
  },

  /* ========== ② BCD：十进制数逐位转 4 位二进制（累积式渲染，骨架固定防抖动） ========== */
  genBcd(){
    const customEl=$('#e-bcd-custom');
    const custom=customEl?customEl.value.trim():'';
    let num;
    if(custom && /^\d{1,4}$/.test(custom)){
      num=parseInt(custom,10); this._customNum=custom;
    }else{
      num=randInt(10,999);
    }
    const digits=String(num).split('');
    const bits=digits.map(d=>parseInt(d,10).toString(2).padStart(4,'0'));
    const bcd=bits.join('');
    const weights=[8,4,2,1];
    const active=digits.map(d=>{ const v=parseInt(d,10); return weights.filter(w=>(v&w)!==0); });
    const N=digits.length;
    const label='BCD 编码';
    const steps=[];

    /* 固定骨架渲染：intro=题目步；digitIdx=-1 表示尚未开始；final=完成步 */
    const render=(intro,digitIdx,finalDone)=>{
      /* 源数字行（始终存在，稳定布局） */
      const sourceRow=`<div class="enc-chip-row">
        <span class="enc-sub" style="margin-right:6px">十进制数</span>
        ${digits.map(d=>`<span class="enc-chip src">${d}</span>`).join('')}
      </div>`;
      /* 对照表（骨架固定，未来位显示 ?） */
      const table=`<table class="exp-table enc-table" style="max-width:520px">
        <tr><th>十进制位</th>${digits.map(d=>`<th>${d}</th>`).join('')}</tr>
        <tr><th>BCD（4位）</th>${digits.map((d,j)=>{
          const st=digitIdx<0?'empty':j<digitIdx?'done':j===digitIdx?'cur':'empty';
          return `<td class="${st}">${digitIdx>=0&&j<=digitIdx?bits[j]:'?'}</td>`;
        }).join('')}</tr>
      </table>`;
      /* 累积展开区：已完成位保留同屏，当前位橙色高亮，未来位占位 */
      let expList;
      if(digitIdx<0){
        expList=`<div class="enc-placeholder">逐位转换中…（每位十进制数按 8421 权展开为 4 位二进制）</div>`;
      }else{
        expList=digits.map((d,j)=>{
          const st=j<digitIdx?'done':j===digitIdx?'cur':'locked';
          if(st==='locked') return '';
          const aw=active[j];
          return `<div class="enc-mini">
            <div class="enc-h">第 ${j+1} 位「${d}」· 8421 权展开</div>
            <div class="enc-calc-row">${weights.map(w=>`<span class="enc-chip ${aw.includes(w)?'cur':'empty'}" style="font-size:clamp(13px,1.8vw,17px);padding:4px 12px">${w}</span>`).join('')}</div>
            <div class="enc-calc-row">${bits[j].split('').map(b=>`<span class="enc-chip ${b==='1'?'cur':'empty'}" style="font-size:clamp(14px,2vw,19px);padding:4px 12px">${b}</span>`).join('')}</div>
            <div class="enc-calc">${aw.length?aw.join(' + '):'0'} <span class="enc-eq">=</span> <span class="enc-res">${d}</span> <span class="enc-op">→</span> <span class="enc-res">${bits[j]}</span></div>
          </div>`;
        }).join('');
      }
      /* 累积拼接行：已完成分组保留，未来分组占位 */
      const snapshot=`<div class="enc-chip-row" style="margin-top:10px">
        <span class="enc-sub" style="margin-right:6px">BCD 拼接</span>
        ${digits.map((d,j)=>{
          const st=digitIdx<0?'empty':j<digitIdx?'done':j===digitIdx?'cur':'empty';
          return `<span class="enc-chip ${st}">${digitIdx>=0&&j<=digitIdx?bits[j]:'????'}</span>`;
        }).join('')}
      </div>`;
      /* 结果条（预留高度，未完成时占位） */
      let strip;
      if(finalDone){
        strip=`<div class="enc-result-strip">${num}<sub>十进制</sub> <span class="enc-op">=</span> ${bcd}<sub>BCD</sub> <span class="enc-op">（${N}×4 = ${N*4} 位 = ${Math.ceil(N*4/8)} 字节）</span></div>`;
      }else{
        strip=`<div class="enc-result-strip locked">拼接结果将在最后一步汇总显示</div>`;
      }
      return sourceRow+table+expList+snapshot+strip;
    };

    /* 步骤 0：输入十进制数 */
    steps.push({desc:`写出十进制数 <b class="c-op">${num}</b>；BCD 编码规定：每位十进制数用 <b>4 位二进制</b>表示（8421 权：8、4、2、1）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">输入十进制数</span></div>
          <div class="step-title"><span class="st-dot"></span>待编码的十进制数 · 逐位转换</div>
          ${render(true,-1,false)}`;
      }});
    /* 步骤 1..N：逐位转换（累积保留） */
    digits.forEach((d,i)=>{
      const snap=bits.slice(0,i+1).join('');
      const aw=active[i];
      steps.push({desc:`第 ${i+1} 位 <b class="c-op">${d}</b>：${aw.length?aw.join(' + '):'0'} = ${d} → 4 位二进制 <b class="c-op">${bits[i]}</b>；已拼接 = <b class="c-op">${snap}</b>。`,
        draw:()=>{
          $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">第 ${i+1}/${N} 位编码</span></div>
            <div class="step-title"><span class="st-dot"></span>逐位编码（8421 权展开，已完成位保留显示）</div>
            ${render(false,i,false)}`;
        }});
    });
    /* 最后：拼接完成 */
    steps.push({desc:`拼接完成：<b class="c-op">${num}</b><sub>十进制</sub> = <b class="c-op">${bcd}</b><sub>BCD</sub>，共 ${N}×4 = <b>${N*4}</b> 位二进制，占 <b>${Math.ceil(N*4/8)}</b> 字节存储。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">拼接完成</span></div>
          <div class="step-title"><span class="st-dot"></span>编码结果（全部位数保留同屏）</div>
          ${render(false,N-1,true)}`;
      }});
    return steps;
  },

  /* ========== ③ 区位码 → 国标码 → 机内码（三码卡片同屏，步骤累加） ========== */
  genHanzi(){
    const qEl=$('#e-hanzi-q'), wEl=$('#e-hanzi-w');
    const qCustom=qEl?qEl.value.trim():'';
    const wCustom=wEl?wEl.value.trim():'';
    let qh,wh;
    if(qCustom && /^\d+$/.test(qCustom) && +qCustom>=1 && +qCustom<=94){ qh=+qCustom; this._customQ=qCustom; }
    else{ qh=randInt(16,55); }
    if(wCustom && /^\d+$/.test(wCustom) && +wCustom>=1 && +wCustom<=94){ wh=+wCustom; this._customW=wCustom; }
    else{ wh=randInt(1,94); }
    const qhHex=qh.toString(16).toUpperCase().padStart(2,'0');
    const whHex=wh.toString(16).toUpperCase().padStart(2,'0');
    const gbQ=(qh+0x20).toString(16).toUpperCase().padStart(2,'0');
    const gbW=(wh+0x20).toString(16).toUpperCase().padStart(2,'0');
    const neiQ=(qh+0xA0).toString(16).toUpperCase().padStart(2,'0');
    const neiW=(wh+0xA0).toString(16).toUpperCase().padStart(2,'0');
    /* 区位码 = 区号(两位) + 位号(两位)，如 区号16 位号01 → 1601 */
    const weiCode=String(qh).padStart(2,'0')+String(wh).padStart(2,'0');
    const label='区位码 → 国标码 → 机内码';
    /* 除16取余竖式（区位码十进制 → 十六进制） */
    const div16=(n)=>{ const rows=[]; let t=n;
      if(t===0) rows.push({n:0,q:0,r:'0'});
      while(t>0){ const q=Math.floor(t/16),r=t%16; rows.push({n:t,q,r:r.toString(16).toUpperCase()}); t=q; }
      return rows; };
    const qRows=div16(qh), wRows=div16(wh);
    const divHtml=rows=>rows.map(rd=>`<div>${rd.n} ÷ 16 = ${rd.q} <span style="color:var(--orange);font-weight:700">余 ${rd.r}</span></div>`).join('');
    const steps=[];

    /* 三码卡片渲染：cur=0 区位 / 1 国标 / 2 机内；未到达卡片锁定占位 */
    const renderCards=(cur)=>{
      /* 卡片 A：区位码 */
      let cA;
      if(cur>=0){
        cA=this._card(this._state(0,cur),'区位码','GB2312 中的位置 · 十进制',
          `<div class="enc-calc-row">
             <span class="enc-chip src">区号 ${qh}</span>
             <span class="enc-chip src">位号 ${wh}</span>
           </div>
           <div class="enc-mini">
             <div class="enc-h">区号 ${qh}₁₀ ÷ 16 取余（余数逆序）</div>${divHtml(qRows)}
             <div class="enc-h" style="margin-top:6px">位号 ${wh}₁₀ ÷ 16 取余（余数逆序）</div>${divHtml(wRows)}
           </div>
           <div class="enc-calc">区 ${qh}₁₀ <span class="enc-op">=</span> <span class="enc-res">${qhHex}H</span></div>
           <div class="enc-calc">位 ${wh}₁₀ <span class="enc-op">=</span> <span class="enc-res">${whHex}H</span></div>
           <div class="enc-big">区位码 ${weiCode}<sub style="font-size:.45em;color:var(--text-3)">十进制</sub></div>`);
      }else{
        cA=this._card('locked','区位码','GB2312 中的位置 · 十进制',this._lockedInner());
      }
      /* 卡片 B：国标码 */
      let cB;
      if(cur>=1){
        cB=this._card(this._state(1,cur),'国标码','区位码 + 2020H（区、位分别 +20H）',
          `<div class="enc-calc">区 ${qhHex}H <span class="enc-op">+ 20H</span> <span class="enc-eq">=</span> <span class="enc-res">${gbQ}H</span></div>
           <div class="enc-calc">位 ${whHex}H <span class="enc-op">+ 20H</span> <span class="enc-eq">=</span> <span class="enc-res">${gbW}H</span></div>
           <div class="enc-big">国标码 ${gbQ}${gbW}H</div>
           <div class="enc-sub">国标码 = 区位码 + 2020H</div>`);
      }else{
        cB=this._card('locked','国标码','区位码 + 2020H',this._lockedInner('等待“区位码 + 2020H”计算…'));
      }
      /* 卡片 C：机内码 */
      let cC;
      if(cur>=2){
        cC=this._card('done','机内码','国标码 + 8080H（区、位分别 +80H）',
          `<div class="enc-calc">区 ${gbQ}H <span class="enc-op">+ 80H</span> <span class="enc-eq">=</span> <span class="enc-res">${neiQ}H</span></div>
           <div class="enc-calc">位 ${gbW}H <span class="enc-op">+ 80H</span> <span class="enc-eq">=</span> <span class="enc-res">${neiW}H</span></div>
           <div class="enc-big">机内码 ${neiQ}${neiW}H</div>
           <div class="enc-sub">计算机内部存储汉字使用机内码 · 占 2 字节</div>`);
      }else{
        cC=this._card('locked','机内码','国标码 + 8080H',this._lockedInner('等待“国标码 + 8080H”计算…'));
      }
      return `<div class="enc-tri">${cA}${cB}${cC}</div>`;
    };

    /* 步骤 0：区位码 */
    steps.push({desc:`已知汉字区位码：区号 <b class="c-op">${qh}</b>，位号 <b class="c-op">${wh}</b>（十进制）；区、位分别除 16 取余转十六进制：${qh}→${qhHex}H，${wh}→${whHex}H。`,
      draw:()=>{
        const cur=0;
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">区位码 ${qh} ${wh}</span></div>
          <div class="step-title"><span class="st-dot"></span>已知区位码（十进制 → 十六进制）· 三码同屏对比</div>
          ${renderCards(cur)}
          <div class="enc-result-strip locked">国标码、机内码将在后续步骤依次累加显示</div>`;
      }});
    /* 步骤 1：+2020H 得国标码（保留区位码卡片） */
    steps.push({desc:`国标码 = 区位码 + 2020H：区 ${qhHex}H + 20H = <b class="c-op">${gbQ}H</b>，位 ${whHex}H + 20H = <b class="c-op">${gbW}H</b>，国标码为 <b class="c-op">${gbQ}${gbW}H</b>。`,
      draw:()=>{
        const cur=1;
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">+2020H · 国标码</span></div>
          <div class="step-title"><span class="st-dot"></span>求国标码（区位码卡片保留同屏）</div>
          ${renderCards(cur)}
          <div class="enc-result-strip">区位码 ${weiCode} <span class="enc-op">→</span> 国标码 ${gbQ}${gbW}H <span class="enc-op">→</span> <span class="enc-chip empty" style="font-size:inherit;padding:2px 12px">机内码 待计算</span></div>`;
      }});
    /* 步骤 2：+8080H 得机内码（前两码卡片保留同屏） */
    steps.push({desc:`机内码 = 国标码 + 8080H：区 ${gbQ}H + 80H = <b class="c-op">${neiQ}H</b>，位 ${gbW}H + 80H = <b class="c-op">${neiW}H</b>；汉字在机内占 <b>2 字节</b>。完整链路：区位码 ${weiCode} → 国标码 ${gbQ}${gbW}H → 机内码 ${neiQ}${neiW}H。`,
      draw:()=>{
        const cur=2;
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">+8080H · 机内码</span></div>
          <div class="step-title"><span class="st-dot"></span>求机内码（区位码、国标码卡片均保留同屏）</div>
          ${renderCards(cur)}
          <div class="enc-result-strip">区位码 ${weiCode} <span class="enc-op">→</span> 国标码 ${gbQ}${gbW}H <span class="enc-op">→</span> 机内码 ${neiQ}${neiW}H <span class="enc-op">（2 字节）</span></div>`;
      }});
    return steps;
  },

  /* ========== ④ 字形码容量：规格 → 总点数 → 字节数 → 总容量（步骤累加） ========== */
  genFont(){
    const elSize=$('#e-font-size'), elCount=$('#e-font-count');
    let size=elSize?+elSize.value:16;
    if(![16,24,32].includes(size)) size=16;
    this._fontSize=String(size);
    let count=elCount?parseInt(elCount.value.trim(),10):1000;
    if(!count||count<1) count=1000;
    this._fontCount=String(count);
    const w=size,h=size;
    const bits=w*h;
    const bytes=bits/8;
    const total=bytes*count;
    const kb=total/1024;
    const label='字形码容量';
    const steps=[];

    const renderCards=(cur)=>{
      /* 卡片 0：已知参数 */
      const c0=this._card(this._state(0,cur),'① 已知参数','点阵规格 · 汉字数量',
        cur>=0
          ?`<div class="enc-chip-row">
              <span class="enc-chip src">${w} × ${h} 点阵</span>
              <span class="enc-chip src">${count} 个汉字</span>
            </div>
            <div class="enc-sub">点阵中每个点用 1 位二进制表示</div>`
          :this._lockedInner());
      /* 卡片 1：总点数 */
      let c1;
      if(cur>=1){
        c1=this._card(this._state(1,cur),'② 计算每字总点数','宽 × 高 = 点数',
          `<div class="enc-calc">${w} <span class="enc-op">×</span> ${h} <span class="enc-eq">=</span> <span class="enc-res">${bits} 点</span></div>
           <div class="enc-sub">每个汉字的点阵总点数 = ${w} × ${h} = ${bits} 点</div>`);
      }else{
        c1=this._card('locked','② 计算每字总点数','宽 × 高 = 点数',this._lockedInner());
      }
      /* 卡片 2：字节数 */
      let c2;
      if(cur>=2){
        c2=this._card(this._state(2,cur),'③ 换算每字字节数','点数 ÷ 8（1 字节 = 8 位）',
          `<div class="enc-calc">${bits} <span class="enc-op">÷ 8</span> <span class="enc-eq">=</span> <span class="enc-res">${bytes} 字节/字</span></div>
           <div class="enc-sub">每字占用 ${bytes} 字节存储</div>`);
      }else{
        c2=this._card('locked','③ 换算每字字节数','点数 ÷ 8',this._lockedInner());
      }
      /* 卡片 3：总容量 */
      let c3;
      if(cur>=3){
        c3=this._card('done','④ 字库总存储容量','每字字节数 × 汉字数量',
          `<div class="enc-calc">${bytes} <span class="enc-op">×</span> ${count} <span class="enc-eq">=</span> <span class="enc-res">${total.toLocaleString()} B</span></div>
           <div class="enc-calc">${total.toLocaleString()} B <span class="enc-op">÷ 1024 ≈</span> <span class="enc-res">${kb.toFixed(1)} KB</span></div>
           <div class="enc-sub">1 KB = 1024 B</div>`);
      }else{
        c3=this._card('locked','④ 字库总存储容量','每字字节数 × 汉字数量',this._lockedInner());
      }
      return c0+c1+c2+c3;
    };

    /* 步骤 0：已知参数 */
    steps.push({desc:`题目：<b class="c-op">${w}×${h}</b> 点阵字形码，共 <b class="c-op">${count}</b> 个汉字，分步计算该字库的总存储容量。`,
      draw:()=>{
        const cur=0;
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">${w}×${h} 点阵 · ${count} 字</span></div>
          <div class="step-title"><span class="st-dot"></span>已知参数（后续步骤逐步累加保留）</div>
          ${renderCards(cur)}
          <div class="enc-result-strip locked">总容量结果将在最后一步汇总显示</div>`;
      }});
    /* 步骤 1：总点数 */
    steps.push({desc:`每个汉字的点阵总点数 = 宽 × 高 = ${w} × ${h} = <b class="c-op">${bits}</b> 点。`,
      draw:()=>{
        const cur=1;
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">每字总点数</span></div>
          <div class="step-title"><span class="st-dot"></span>计算每字占用点数</div>
          ${renderCards(cur)}
          <div class="enc-result-strip locked">总容量结果将在最后一步汇总显示</div>`;
      }});
    /* 步骤 2：每字字节数 */
    steps.push({desc:`每字字节数 = 点阵点数 ÷ 8 = ${bits} ÷ 8 = <b class="c-op">${bytes}</b> 字节/字（1 字节 = 8 位）。`,
      draw:()=>{
        const cur=2;
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">每字字节数</span></div>
          <div class="step-title"><span class="st-dot"></span>点数 ÷ 8 换算为字节</div>
          ${renderCards(cur)}
          <div class="enc-result-strip locked">总容量结果将在最后一步汇总显示</div>`;
      }});
    /* 步骤 3：总容量 */
    steps.push({desc:`总容量 = 每字字节数 × 汉字数 = ${bytes} × ${count} = <b class="c-op">${total.toLocaleString()}</b> 字节 ≈ <b class="c-op">${kb.toFixed(1)} KB</b>（1 KB = 1024 B）。`,
      draw:()=>{
        const cur=3;
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>字库总容量（前三步结果均保留同屏）</div>
          ${renderCards(cur)}
          <div class="enc-result-strip">${w}×${h} 点阵 × ${count} 字 = ${total.toLocaleString()} B ≈ ${kb.toFixed(1)} KB</div>`;
      }});
    return steps;
  },

  /* 初始化演示画布（切换子页面时清空旧内容） */
  initStage(){
    const sub=this.currentSub;
    const info=this.subTabs.find(t=>t.id===sub)||{};
    const labelMap={
      ascii:'字符 → ASCII 码值',
      bcd:'BCD 编码',
      hanzi:'区位码 → 国标码 → 机内码',
      font:'字形码容量计算'
    };
    $('#stage').innerHTML=`<div class="stage-hd"><span>${labelMap[sub]||'字符编码'}</span><span class="state">准备就绪</span></div>
      <div style="text-align:center;padding:clamp(16px,3vw,28px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.8vw,16px);margin-bottom:8px">使用 <b style="color:${this.color}">${info.method||''}</b></div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">点击「自动演示」或「下一步」开始逐步推导</div>
      </div>`;
  },
  generate(){
    const sub=this.currentSub;
    switch(sub){
      case 'ascii': this.steps=this.genAscii(); break;
      case 'bcd': this.steps=this.genBcd(); break;
      case 'hanzi': this.steps=this.genHanzi(); break;
      case 'font': this.steps=this.genFont(); break;
      default: this.steps=this.genAscii();
    }
  }
};
