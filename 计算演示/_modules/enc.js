/* ================================================================
 * 字符编码（ASCII / BCD / 区位码→国标码→机内码 / 字形码容量）
 * 子页面：ascii · bcd · hanzi · font
 * ================================================================ */
MOD.enc = {
  name:'字符编码', sub:'ASCII · BCD · 区位码国标码机内码 · 字形码',
  badge:'编码 · 转换', tag:'编码', cat:'bit',
  color:'#8B5CF6', soft:'#F3EEFF', type:'ASCII · BCD · 区位码 · 国标码 · 机内码 · 字形码',
  desc:'涵盖 ASCII 码值查询、BCD 编码、汉字区位码→国标码→机内码转换、字形码容量计算，每个转换方向独立子页面，用手写动画逐步演示推导过程。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6h16M4 12h10M4 18h7M16 15l5 5M16 20l5-5"/></svg>',
  setDesc:'',
  stageDesc:'橙色=当前计算项，绿色=已得出结果，逐步骤展示编码转换过程。',
  initDesc:'点击「自动演示」或「下一步」，观看字符编码的手写推导过程。',
  currentSub:'ascii',
  subTabs:[
    {id:'ascii',label:'ASCII码值',method:'字符→十进制→二进制→十六进制'},
    {id:'bcd',label:'BCD编码',method:'逐位4位二进制'},
    {id:'hanzi',label:'区位→国标→机内',method:'+2020H · +8080H'},
    {id:'font',label:'字形码容量',method:'位数→字节→总容量'}
  ],
  theory(){
    return `<div class="rule"><b>ASCII</b>：美国信息交换标准代码，用 7 位（1 字节）表示 128 个字符，含英文、数字、符号和控制符。</div>
    <div class="rule"><b>BCD 码</b>：二进制编码的十进制，用 4 位二进制表示 1 位十进制数（0~9），常用于数字显示。</div>
    <div class="rule"><b>区位码</b>：汉字在 GB2312 中的位置，由区号（1~94）和位号（1~94）组成，是十进制表示。</div>
    <div class="rule"><b>国标码</b> = 区位码（区、位分别）+ 2020H。即区号转十六进制后 +20H，位号 +20H。</div>
    <div class="rule"><b>机内码</b> = 国标码 + 8080H（即区、位分别 +80H），计算机内部存储和处理汉字时使用。</div>
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
        <input type="text" id="e-hanzi-q" value="${esc(curQ)}" placeholder="16~87" maxlength="2" style="width:clamp(50px,8vw,65px)"></div></div>
        <div class="inline-setting"><div class="setting"><label>位号</label>
        <input type="text" id="e-hanzi-w" value="${esc(curW)}" placeholder="1~94" maxlength="2" style="width:clamp(50px,8vw,65px)"></div></div>`;
    }
    if(sub==='font'){
      const curS=$('#e-font-size')?$('#e-font-size').value:'16';
      const curC=$('#e-font-count')?$('#e-font-count').value:'1000';
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
  /* 逐字出现动画辅助 */
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.06}s">${esc(ch)}</span>`).join('');
  },
  /* ========== ASCII 码值：字符 → 十进制 → 二进制 → 十六进制 ========== */
  genAscii(){
    const customEl=$('#e-ascii-custom');
    const custom=customEl?customEl.value.trim():'';
    let ch;
    if(custom && custom.length>0){
      ch=custom[0];
      this._customChar=ch;
    }else{
      const el=$('#e-ascii');
      const group=el?el.value:'upper';
      const pools={upper:'ABCDEFGHIJKLMNOPQRSTUVWXYZ',lower:'abcdefghijklmnopqrstuvwxyz',digit:'0123456789',symbol:'!?#@%&*()/-'};
      const pool=pools[group]||pools.upper;
      ch=pick([...pool]);
    }
    const code=ch.charCodeAt(0);
    const bin=code.toString(2).padStart(8,'0');
    const hex=code.toString(16).toUpperCase().padStart(2,'0');
    const label='字符 → ASCII 码值';
    /* 预计算除 2 取余竖式 */
    const rows=[];
    let temp=code;
    if(temp===0) rows.push({n:0,q:0,r:'0'});
    while(temp>0){
      const q=Math.floor(temp/2), r=temp%2;
      rows.push({n:temp,q,r}); temp=q;
    }
    /* 生成 ASCII 码表片段（当前字符前后各3个） */
    const tableStart=Math.max(32, code-3);
    const tableEnd=Math.min(126, code+3);
    const tableRows=[];
    for(let c=tableStart;c<=tableEnd;c++){
      tableRows.push({ch:String.fromCharCode(c), code:c, isCur:c===code});
    }
    const steps=[];
    /* 步骤1：写出字符 */
    steps.push({desc:`写出要查询的字符 <b class="c-op">"${ch}"</b>，准备对照 ASCII 码表求码值。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">准备</span></div>
          <div style="text-align:center;padding:clamp(18px,3.5vw,30px) 0">
            <span style="font-family:var(--mono);font-size:clamp(26px,5vw,38px);font-weight:700;color:var(--text)">
              "${this.typeChars(ch)}"
            </span>
            <span class="caret pen"></span>
          </div>`;
      }});
    /* 步骤2：查表得十进制码值（带码表片段） */
    steps.push({desc:`查 ASCII 码表：字符 <b class="c-op">${ch}</b> 对应的十进制码值 = <b class="c-op">${code}</b>。`,
      draw:()=>{
        const tableHtml=`<table class="exp-table" style="margin:10px auto;max-width:500px">
          <tr><th>字符</th>${tableRows.map(r=>`<th class="${r.isCur?'cur':''}">${r.ch===' '?'␣':esc(r.ch)}</th>`).join('')}</tr>
          <tr><th>十进制</th>${tableRows.map(r=>`<td class="${r.isCur?'cur':''}">${r.code}</td>`).join('')}</tr>
        </table>`;
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">查码表</span></div>
          <div class="step-title"><span class="st-dot"></span>对照 ASCII 码表（当前字符高亮）</div>
          ${tableHtml}
          <div class="annot" style="margin-top:6px;padding:6px 10px;background:#F8FAFC;border-radius:8px;display:inline-block">
            常见对照：A=65，a=97，0=48；本字符查表得 <b>${code}</b>。
          </div>`;
      }});
    /* 步骤3：除 2 取余竖式求二进制 */
    steps.push({desc:`十进制 <b class="c-op">${code}</b> 连续除以 2 并记录余数，余数<b>从下往上逆序</b>读出，得到 8 位二进制 <b class="c-op">${bin}</b>。`,
      draw:()=>{
        const rowsHtml=rows.map((rd,i)=>`<div class="dv-row write-in ${i===rows.length-1?'cur':''}" style="animation-delay:${i*0.09}s">
          <span class="dv-num">${rd.n}</span><span class="dv-op">÷</span><span class="dv-num">2</span><span class="dv-eq">=</span><span class="dv-q">${rd.q}</span><span class="dv-rem">${rd.r}</span>
        </div>`).join('');
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">除 2 取余</span></div>
          <div class="step-title"><span class="st-dot"></span>竖式除法（基数 = 2）</div>
          <div class="div-vertical" style="margin-bottom:10px">${rowsHtml}</div>
          <div class="calc-process">
            <div class="cp-label">余数逆序收集</div>
            <div class="cp-sum-line">
              余数从下往上：${rows.slice().reverse().map(rd=>`<span style="color:var(--accent);font-weight:700">${rd.r}</span>`).join(' ')}
              <span class="eq">=</span> <span class="sum">${bin}</span>
            </div>
          </div>`;
      }});
    /* 步骤4：十六进制 + 最终结果 */
    steps.push({desc:`十进制 <b>${code}</b> 对应的十六进制为 <b class="c-op">${hex}H</b>；ASCII 码在计算机中占 <b>1 字节</b>（8 位二进制）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>三种进制对照</div>
          <div class="exp-inline" style="display:flex;flex-wrap:wrap;align-items:center;gap:6px;justify-content:center;margin:10px 0">
            <span class="term done">"${esc(ch)}"</span><span class="op">=</span>
            <span class="term">${code}<sub style="font-size:.55em">10</sub></span><span class="op">=</span>
            <span class="term">${bin}<sub style="font-size:.55em">2</sub></span><span class="op">=</span>
            <span class="term cur">${hex}<sub style="font-size:.55em">16</sub></span>
          </div>
          <div class="write-result">"${esc(ch)}" → 十进制 ${code} → 二进制 ${bin} → 十六进制 ${hex}H（1 字节）</div>`;
      }});
    return steps;
  },
  /* ========== BCD 编码：十进制数逐位转 4 位二进制 ========== */
  genBcd(){
    const customEl=$('#e-bcd-custom');
    const custom=customEl?customEl.value.trim():'';
    let num;
    if(custom && /^\d{1,4}$/.test(custom)){
      num=parseInt(custom);
      this._customNum=custom;
    }else{
      num=randInt(10,999);
    }
    const digits=String(num).split('');
    const bits=digits.map(d=>parseInt(d).toString(2).padStart(4,'0'));
    const bcd=bits.join('');
    const label='BCD 编码';
    const weights=[8,4,2,1];
    const steps=[];
    /* 步骤1：写出十进制数 */
    steps.push({desc:`写出十进制数 <b class="c-op">${num}</b>；BCD 编码规定：每位十进制数用 <b>4 位二进制</b>表示（8421 码）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">${num}</span></div>
          <div class="step-title"><span class="st-dot"></span>待编码的十进制数</div>
          <div style="text-align:center;padding:clamp(12px,2vw,18px) 0">
            <div class="group-box">
              ${digits.map((d,i)=>`<span class="grp write-in" style="animation-delay:${i*0.1}s">${d}</span>`).join('')}
            </div>
            <div style="margin-top:10px;font-size:clamp(11px,1.4vw,12.5px);color:#94A3B8">每位 → 4 位二进制（8421 权）</div>
          </div>`;
      }});
    /* 步骤2：列出数位对照表 */
    steps.push({desc:`列出每一位十进制数码，逐位按 8421 权展开转换为 4 位二进制（0→0000，9→1001）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">列出数位</span></div>
          <div class="step-title"><span class="st-dot"></span>数位对照（8421 BCD 码）</div>
          <table class="exp-table">
            <tr><th>十进制位</th>${digits.map((d,i)=>`<th class="write-in" style="animation-delay:${i*0.08}s">${d}</th>`).join('')}</tr>
            <tr><th>BCD(4位)</th>${digits.map(()=>`<td class="empty">?</td>`).join('')}</tr>
          </table>
          <div class="annot" style="margin-top:8px;padding:6px 10px;background:#F8FAFC;border-radius:8px;display:inline-block">
            8421 权：8+4+2+1，每位十进制数 = 各位权值之和
          </div>`;
      }});
    /* 步骤3~N：逐位转换（带8421权展开） */
    let bcdSnap='';
    digits.forEach((d,i)=>{
      bcdSnap+=bits[i];
      const snap=bcdSnap, idx=i;
      const dv=parseInt(d);
      const activeWeights=weights.filter(w=>(dv & w)!==0);
      const calcHTML=`<div class="calc-process">
        <div class="cp-label">第 ${idx+1} 位「${d}」的 8421 权展开</div>
        <div class="cp-sum-line" style="margin-bottom:5px">
          位权：<span style="color:var(--accent)">8</span> <span style="color:var(--accent)">4</span> <span style="color:var(--accent)">2</span> <span style="color:var(--accent)">1</span>
        </div>
        <div class="cp-sum-line" style="margin-bottom:5px">
          二进制：${bits[idx].split('').map((b,k)=>`<span style="color:${b==='1'?'var(--orange)':'#CBD5E1'};font-weight:700">${b}</span>`).join(' ')}
        </div>
        <div class="cp-sum-line" style="background:#FFF4DB">
          ${activeWeights.length?activeWeights.join(' + ')+' = ':'0 = '}<span class="sum">${d}</span>
        </div>
      </div>`;
      steps.push({desc:`第 ${i+1} 位 <b class="c-op">${d}</b>：8421 权展开 ${activeWeights.join('+')||'0'} = ${d} → 4 位二进制 <b class="c-op">${bits[i]}</b>；已拼接 = <b class="c-op">${snap}</b>。`,
        draw:()=>{
          $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">第 ${idx+1}/${digits.length} 位</span></div>
            <div class="step-title"><span class="st-dot"></span>逐位编码（8421 权展开）</div>
            <table class="exp-table">
              <tr><th>十进制位</th>${digits.map(x=>`<th>${x}</th>`).join('')}</tr>
              <tr><th>BCD(4位)</th>${digits.map((x,j)=>`<td class="${j<idx?'done':j===idx?'cur':'empty'}">${j<=idx?bits[j]:'?'}</td>`).join('')}</tr>
            </table>
            ${calcHTML}
            <div class="exp-inline" style="display:inline-block;margin-top:8px">
              <span class="term cur">${d}</span><span class="op">=</span><span class="term cur">${bits[idx]}</span>
              <span class="op">→</span><span class="term">BCD = <b style="color:#4F6BFF">${snap}</b>${idx<digits.length-1?' …':''}</span>
            </div>`;
        }});
    });
    /* 最后：拼接完成 */
    steps.push({desc:`拼接完成：<b class="c-op">${num}</b><sub>十进制</sub> = <b class="c-op">${bcd}</b><sub>BCD</sub>，共 ${digits.length}×4 = <b>${digits.length*4}</b> 位二进制，占 <b>${Math.ceil(digits.length*4/8)}</b> 字节存储。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>编码结果</div>
          <div class="group-box" style="margin:10px 0;justify-content:center">
            ${digits.map(d=>`<span class="grp done">${d}</span>`).join('')}
          </div>
          <div style="text-align:center;font-size:clamp(12px,1.6vw,14px);color:#22C55E;margin:8px 0">↓ 逐位转 4 位二进制 ↓</div>
          <div class="group-box" style="justify-content:center">
            ${bits.map(b=>`<span class="grp done" style="border-color:#22C55E;background:#E8FBF0">${b}</span>`).join('')}
          </div>
          <div class="write-result">${num}<sub>十进制</sub> = ${bcd}<sub>BCD</sub>（${digits.length*4} 位 = ${Math.ceil(digits.length*4/8)} 字节）</div>`;
      }});
    return steps;
  },
  /* ========== 区位码 → 国标码 → 机内码 ========== */
  genHanzi(){
    const qEl=$('#e-hanzi-q'), wEl=$('#e-hanzi-w');
    const qCustom=qEl?qEl.value.trim():'';
    const wCustom=wEl?wEl.value.trim():'';
    let qh, wh;
    if(qCustom && /^\d+$/.test(qCustom) && +qCustom>=1 && +qCustom<=94){
      qh=+qCustom; this._customQ=qCustom;
    }else{
      qh=randInt(16,55);
    }
    if(wCustom && /^\d+$/.test(wCustom) && +wCustom>=1 && +wCustom<=94){
      wh=+wCustom; this._customW=wCustom;
    }else{
      wh=randInt(1,94);
    }
    const qhHex=qh.toString(16).toUpperCase().padStart(2,'0');
    const whHex=wh.toString(16).toUpperCase().padStart(2,'0');
    const gbQ=(qh+0x20).toString(16).toUpperCase().padStart(2,'0');
    const gbW=(wh+0x20).toString(16).toUpperCase().padStart(2,'0');
    const neiQ=(qh+0xA0).toString(16).toUpperCase().padStart(2,'0');
    const neiW=(wh+0xA0).toString(16).toUpperCase().padStart(2,'0');
    const label='区位码 → 国标码 → 机内码';
    /* 除16取余竖式 */
    const div16=(n)=>{
      const rows=[];
      let t=n;
      if(t===0) rows.push({n:0,q:0,r:'0'});
      while(t>0){ const q=Math.floor(t/16),r=t%16; rows.push({n:t,q,r:r.toString(16).toUpperCase()}); t=q; }
      return rows;
    };
    const qRows=div16(qh), wRows=div16(wh);
    const steps=[];
    /* 步骤1：写出区位号 */
    steps.push({desc:`某汉字的区位码：区号 <b class="c-op">${qh}</b>，位号 <b class="c-op">${wh}</b>（均为十进制）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">区位 ${qh} ${wh}</span></div>
          <div class="step-title"><span class="st-dot"></span>已知区位码（十进制）</div>
          <div style="text-align:center;padding:clamp(12px,2vw,18px) 0">
            <div class="group-box">
              <span class="grp write-in" style="animation-delay:.1s">区号 ${qh}</span>
              <span class="grp write-in" style="animation-delay:.25s">位号 ${wh}</span>
            </div>
          </div>`;
      }});
    /* 步骤2：区、位分别转十六进制（除16取余） */
    steps.push({desc:`区号、位号分别<b>除以 16 取余</b>换算为十六进制：<b class="c-op">${qh}</b> → ${qhHex}H，<b class="c-op">${wh}</b> → ${whHex}H。`,
      draw:()=>{
        const renderDiv=(rows,title)=>`<div style="flex:1;min-width:0">
          <div style="font-size:clamp(10px,1.2vw,12px);font-weight:700;color:var(--text-3);margin-bottom:4px;text-align:center">${title}</div>
          <div class="div-vertical" style="margin-bottom:6px">
            ${rows.map((rd,i)=>`<div class="dv-row write-in ${i===rows.length-1?'cur':''}" style="animation-delay:${i*0.08}s">
              <span class="dv-num">${rd.n}</span><span class="dv-op">÷</span><span class="dv-num">16</span><span class="dv-eq">=</span><span class="dv-q">${rd.q}</span><span class="dv-rem">${rd.r}</span>
            </div>`).join('')}
          </div>
          <div style="text-align:center;font-size:clamp(11px,1.4vw,12.5px);color:#22C55E;font-weight:700">余数逆序 = ${rows.slice().reverse().map(r=>r.r).join('')}H</div>
        </div>`;
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">转十六进制</span></div>
          <div class="step-title"><span class="st-dot"></span>区位码（十进制 → 十六进制）· 除16取余</div>
          <div style="display:flex;gap:12px;flex-wrap:wrap">
            ${renderDiv(qRows,`区号 ${qh} ÷16`)}
            ${renderDiv(wRows,`位号 ${wh} ÷16`)}
          </div>`;
      }});
    /* 步骤3：+2020H 得国标码 */
    steps.push({desc:`国标码 = 区位码 + 2020H：区 ${qhHex}H + 20H = <b class="c-op">${gbQ}H</b>，位 ${whHex}H + 20H = <b class="c-op">${gbW}H</b>，国标码为 <b class="c-op">${gbQ} ${gbW}H</b>。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">+2020H</span></div>
          <div class="step-title"><span class="st-dot"></span>求国标码（区位码 + 2020H）</div>
          <div class="calc-process">
            <div class="cp-label">十六进制加法（区、位分别 +20H）</div>
            <div class="cp-sum-line" style="margin-bottom:5px">${qhHex}H <span class="op">+ 20H</span> <span class="eq">=</span> <span class="sum">${gbQ}H</span></div>
            <div class="cp-sum-line">${whHex}H <span class="op">+ 20H</span> <span class="eq">=</span> <span class="sum">${gbW}H</span></div>
          </div>
          <div class="write-result">国标码 = ${gbQ} ${gbW}H</div>`;
      }});
    /* 步骤4：+8080H 得机内码 */
    steps.push({desc:`机内码 = 国标码 + 8080H：区 ${gbQ}H + 80H = <b class="c-op">${neiQ}H</b>，位 ${gbW}H + 80H = <b class="c-op">${neiW}H</b>；汉字在机内占 <b>2 字节</b>。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">+8080H</span></div>
          <div class="step-title"><span class="st-dot"></span>求机内码（国标码 + 8080H）</div>
          <div class="calc-process">
            <div class="cp-label">十六进制加法（区、位分别 +80H）</div>
            <div class="cp-sum-line" style="margin-bottom:5px">${gbQ}H <span class="op">+ 80H</span> <span class="eq">=</span> <span class="sum">${neiQ}H</span></div>
            <div class="cp-sum-line">${gbW}H <span class="op">+ 80H</span> <span class="eq">=</span> <span class="sum">${neiW}H</span></div>
          </div>
          <div class="write-result">区位 ${qh} ${wh} → 国标 ${gbQ}${gbW}H → 机内 ${neiQ}${neiW}H（2 字节）</div>`;
      }});
    return steps;
  },
  /* ========== 字形码容量：点阵位数 → 字节数 → 总容量 ========== */
  genFont(){
    const elSize=$('#e-font-size'), elCount=$('#e-font-count');
    const size=elSize?+elSize.value:16;
    const count=elCount?+elCount.value:1000;
    const w=size, h=size;
    const bits=w*h;
    const bytes=bits/8;
    const total=bytes*count;
    const kb=total/1024;
    const label='字形码容量';
    const steps=[];
    /* 步骤1：写出点阵参数与汉字数量 */
    steps.push({desc:`题目：${w}×${h} 点阵字形码，共 <b class="c-op">${count}</b> 个汉字，计算该字库的总存储容量。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">${w}×${h} 点阵</span></div>
          <div class="step-title"><span class="st-dot"></span>已知参数</div>
          <div style="text-align:center;padding:clamp(12px,2vw,18px) 0">
            <div class="group-box">
              <span class="grp write-in" style="animation-delay:.1s">${w} × ${h} 点阵</span>
              <span class="grp write-in" style="animation-delay:.25s">${count} 个汉字</span>
            </div>
          </div>`;
      }});
    /* 步骤2：每字点阵位数 */
    steps.push({desc:`每个汉字的点阵位数 = 宽 × 高 = ${w} × ${h} = <b class="c-op">${bits}</b> 位。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">每字位数</span></div>
          <div class="step-title"><span class="st-dot"></span>计算每字占用位数</div>
          <div class="exp-inline" style="display:inline-block;margin:10px 0">
            <span class="term cur">${w}</span><span class="op">×</span><span class="term cur">${h}</span>
            <span class="op">=</span><span class="term done write-in" style="animation-delay:.25s">${bits} 位</span>
          </div>
          <div class="annot" style="padding:6px 10px;background:#F8FAFC;border-radius:8px;display:inline-block">点阵中每个点用 1 位二进制表示。</div>`;
      }});
    /* 步骤3：每字字节数 */
    steps.push({desc:`每字字节数 = 点阵位数 ÷ 8 = ${bits} ÷ 8 = <b class="c-op">${bytes}</b> 字节/字。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">每字字节</span></div>
          <div class="step-title"><span class="st-dot"></span>位数 ÷ 8 换算为字节</div>
          <div class="exp-inline" style="display:inline-block;margin:10px 0">
            <span class="term cur">${bits}</span><span class="op">÷ 8 =</span>
            <span class="term done write-in" style="animation-delay:.25s">${bytes} 字节/字</span>
          </div>`;
      }});
    /* 步骤4：总容量与单位换算 */
    steps.push({desc:`总容量 = 每字字节数 × 汉字数 = ${bytes} × ${count} = <b class="c-op">${total.toLocaleString()}</b> 字节 ≈ <b class="c-op">${kb.toFixed(1)} KB</b>（1 KB = 1024 B）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>字库总容量</div>
          <div class="exp-inline" style="display:block;margin:6px 0">
            <span class="term cur">${bytes}</span><span class="op">×</span><span class="term cur">${count}</span>
            <span class="op">=</span><span class="term done">${total.toLocaleString()} B</span>
          </div>
          <div class="exp-inline" style="display:block;margin:6px 0">
            <span class="term done">${total.toLocaleString()} B</span><span class="op">÷ 1024 ≈</span>
            <span class="term cur write-in" style="animation-delay:.25s">${kb.toFixed(1)} KB</span>
          </div>
          <div class="write-result">${w}×${h} 点阵 × ${count} 字 = ${total.toLocaleString()} B ≈ ${kb.toFixed(1)} KB</div>`;
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
        <div style="font-size:clamp(11px,1.4vw,12.5px)">点击「自动演示」或「下一步」开始手写推导</div>
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
