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
      const cur=$('#e-ascii')?$('#e-ascii').value:'upper';
      return `<div class="inline-setting"><div class="setting"><label>字符类型</label>
        <select id="e-ascii">
          <option value="upper" ${cur==='upper'?'selected':''}>大写字母</option>
          <option value="lower" ${cur==='lower'?'selected':''}>小写字母</option>
          <option value="digit" ${cur==='digit'?'selected':''}>数字</option>
          <option value="symbol" ${cur==='symbol'?'selected':''}>符号</option>
        </select></div></div>`;
    }
    if(sub==='bcd'){
      const cur=$('#e-bcd-len')?$('#e-bcd-len').value:'3';
      return `<div class="inline-setting"><div class="setting"><label>数字位数</label>
        <select id="e-bcd-len">
          <option value="2" ${cur==='2'?'selected':''}>2 位</option>
          <option value="3" ${cur==='3'?'selected':''}>3 位</option>
        </select></div></div>`;
    }
    if(sub==='hanzi'){
      const cur=$('#e-hanzi-level')?$('#e-hanzi-level').value:'common';
      return `<div class="inline-setting"><div class="setting"><label>区位范围</label>
        <select id="e-hanzi-level">
          <option value="common" ${cur==='common'?'selected':''}>常用字（区号 16~55）</option>
          <option value="rare" ${cur==='rare'?'selected':''}>次常用字（区号 56~87）</option>
        </select></div></div>`;
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
        <select id="e-font-count">
          <option value="500" ${curC==='500'?'selected':''}>500 字</option>
          <option value="1000" ${curC==='1000'?'selected':''}>1000 字</option>
          <option value="2000" ${curC==='2000'?'selected':''}>2000 字</option>
        </select></div></div>`;
    }
    return '';
  },
  afterSettings(){
    const bind=id=>{ const el=$('#'+id); if(el) el.onchange=()=>Engine.reload(); };
    ['e-ascii','e-bcd-len','e-hanzi-level','e-font-size','e-font-count'].forEach(bind);
  },
  /* 逐字出现动画辅助 */
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.06}s">${esc(ch)}</span>`).join('');
  },
  /* ========== ASCII 码值：字符 → 十进制 → 二进制 → 十六进制 ========== */
  genAscii(){
    const el=$('#e-ascii');
    const group=el?el.value:'upper';
    const pools={
      upper:'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
      lower:'abcdefghijklmnopqrstuvwxyz',
      digit:'0123456789',
      symbol:'!?#@%&*()/-'
    };
    const pool=pools[group]||pools.upper;
    const ch=pick([...pool]);
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
    /* 步骤2：查表得十进制码值 */
    steps.push({desc:`查 ASCII 码表：字符 <b class="c-op">${ch}</b> 对应的十进制码值 = <b class="c-op">${code}</b>。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">查码表</span></div>
          <div class="step-title"><span class="st-dot"></span>对照 ASCII 码表</div>
          <div class="exp-inline" style="display:inline-flex;align-items:center;gap:8px;margin-top:8px">
            <span class="term cur" style="font-size:clamp(15px,2.2vw,19px)">"${esc(ch)}"</span>
            <span class="op">→</span>
            <span class="term write-in" style="animation-delay:.25s;color:#22C55E;font-weight:700;font-size:clamp(15px,2.2vw,19px)">${code}<sub style="font-size:.55em;color:#94A3B8">十进制</sub></span>
            <span class="caret pen" style="margin-left:6px"></span>
          </div>
          <div class="annot" style="margin-top:10px;padding:6px 10px;background:#F8FAFC;border-radius:8px;display:inline-block">
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
          <div class="annot" style="padding:6px 10px;background:#F8FAFC;border-radius:8px;display:inline-block">
            <b>余数收集</b>（从下往上读）：
            ${rows.slice().reverse().map(rd=>`<span class="chip g" style="margin:0 2px;font-size:clamp(12px,1.6vw,15px)">${rd.r}</span>`).join('')}
            <span class="up-arrow" style="margin-left:8px">↑ 逆序 = ${bin}</span>
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
    const el=$('#e-bcd-len');
    const len=el?+el.value:3;
    const num=randInt(Math.pow(10,len-1),Math.pow(10,len)-1);
    const digits=String(num).split('');
    const bits=digits.map(d=>parseInt(d).toString(2).padStart(4,'0'));
    const bcd=bits.join('');
    const label='BCD 编码';
    const steps=[];
    /* 步骤1：写出十进制数 */
    steps.push({desc:`写出十进制数 <b class="c-op">${num}</b>；BCD 编码规定：每位十进制数用 <b>4 位二进制</b>表示。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">${num}</span></div>
          <div class="step-title"><span class="st-dot"></span>待编码的十进制数</div>
          <div style="text-align:center;padding:clamp(12px,2vw,18px) 0">
            <div class="group-box">
              ${digits.map((d,i)=>`<span class="grp write-in" style="animation-delay:${i*0.1}s">${d}</span>`).join('')}
            </div>
            <div style="margin-top:10px;font-size:clamp(11px,1.4vw,12.5px);color:#94A3B8">每位 → 4 位二进制</div>
          </div>`;
      }});
    /* 步骤2：列出数位对照表 */
    steps.push({desc:`列出每一位十进制数码，逐位转换为 4 位二进制（0→0000，9→1001）。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">列出数位</span></div>
          <div class="step-title"><span class="st-dot"></span>数位对照</div>
          <table class="exp-table">
            <tr><th>十进制位</th>${digits.map((d,i)=>`<th class="write-in" style="animation-delay:${i*0.08}s">${d}</th>`).join('')}</tr>
            <tr><th>BCD(4位)</th>${digits.map(()=>`<td class="empty">?</td>`).join('')}</tr>
          </table>
          <div class="annot" style="margin-top:8px;padding:6px 10px;background:#F8FAFC;border-radius:8px;display:inline-block">
            参考对照：0→0000，1→0001，…，9→1001
          </div>`;
      }});
    /* 步骤3~N：逐位转换 */
    let bcdSnap='';
    digits.forEach((d,i)=>{
      bcdSnap+=bits[i];
      const snap=bcdSnap, idx=i;
      steps.push({desc:`第 ${i+1} 位 <b class="c-op">${d}</b> → 4 位二进制 <b class="c-op">${bits[i]}</b>；当前已拼接 BCD 码 = <b class="c-op">${snap}</b>${i<digits.length-1?' …':''}。`,
        draw:()=>{
          $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">第 ${idx+1}/${digits.length} 位</span></div>
            <div class="step-title"><span class="st-dot"></span>逐位编码</div>
            <table class="exp-table">
              <tr><th>十进制位</th>${digits.map(x=>`<th>${x}</th>`).join('')}</tr>
              <tr><th>BCD(4位)</th>${digits.map((x,j)=>`<td class="${j<idx?'done':j===idx?'cur':'empty'}">${j<=idx?bits[j]:'?'}</td>`).join('')}</tr>
            </table>
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
    const el=$('#e-hanzi-level');
    const lv=el?el.value:'common';
    const qh=lv==='rare'?randInt(56,87):randInt(16,55);
    const wh=randInt(1,94);
    const qhHex=qh.toString(16).toUpperCase().padStart(2,'0');
    const whHex=wh.toString(16).toUpperCase().padStart(2,'0');
    const gbQ=(qh+0x20).toString(16).toUpperCase().padStart(2,'0');
    const gbW=(wh+0x20).toString(16).toUpperCase().padStart(2,'0');
    const neiQ=(qh+0xA0).toString(16).toUpperCase().padStart(2,'0');
    const neiW=(wh+0xA0).toString(16).toUpperCase().padStart(2,'0');
    const label='区位码 → 国标码 → 机内码';
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
    /* 步骤2：区、位分别转十六进制 */
    steps.push({desc:`区号、位号分别换算为十六进制：<b class="c-op">${qh}</b> → ${qhHex}H，<b class="c-op">${wh}</b> → ${whHex}H。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">转十六进制</span></div>
          <div class="step-title"><span class="st-dot"></span>区位码（十进制 → 十六进制）</div>
          <table class="exp-table" style="max-width:420px;margin:0 auto">
            <tr><th>项</th><th>十进制</th><th>十六进制</th></tr>
            <tr><th>区号</th><td>${qh}</td><td class="cur write-in" style="animation-delay:.2s">${qhHex}H</td></tr>
            <tr><th>位号</th><td>${wh}</td><td class="cur write-in" style="animation-delay:.35s">${whHex}H</td></tr>
          </table>
          <div class="exp-inline" style="display:inline-block;margin-top:10px">
            <span class="term done">${qh} → ${qhHex}H</span><span class="op">，</span>
            <span class="term done">${wh} → ${whHex}H</span>
          </div>`;
      }});
    /* 步骤3：+2020H 得国标码 */
    steps.push({desc:`国标码 = 区位码 + 2020H：区 ${qhHex}H + 20H = <b class="c-op">${gbQ}H</b>，位 ${whHex}H + 20H = <b class="c-op">${gbW}H</b>，国标码为 <b class="c-op">${gbQ} ${gbW}H</b>。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">+2020H</span></div>
          <div class="step-title"><span class="st-dot"></span>求国标码（区位码 + 2020H）</div>
          <div class="exp-inline" style="display:block;margin:6px 0">
            <span class="term done">${qhHex}H</span><span class="op">+ 20H =</span>
            <span class="term cur write-in" style="animation-delay:.2s">${gbQ}H</span>
          </div>
          <div class="exp-inline" style="display:block;margin:6px 0">
            <span class="term done">${whHex}H</span><span class="op">+ 20H =</span>
            <span class="term cur write-in" style="animation-delay:.4s">${gbW}H</span>
          </div>
          <div class="write-result">国标码 = ${gbQ} ${gbW}H</div>`;
      }});
    /* 步骤4：+8080H 得机内码 */
    steps.push({desc:`机内码 = 国标码 + 8080H：区 ${gbQ}H + 80H = <b class="c-op">${neiQ}H</b>，位 ${gbW}H + 80H = <b class="c-op">${neiW}H</b>；汉字在机内占 <b>2 字节</b>。`,
      draw:()=>{
        $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">+8080H</span></div>
          <div class="step-title"><span class="st-dot"></span>求机内码（国标码 + 8080H）</div>
          <div class="exp-inline" style="display:block;margin:6px 0">
            <span class="term done">${gbQ}H</span><span class="op">+ 80H =</span>
            <span class="term cur write-in" style="animation-delay:.2s">${neiQ}H</span>
          </div>
          <div class="exp-inline" style="display:block;margin:6px 0">
            <span class="term done">${gbW}H</span><span class="op">+ 80H =</span>
            <span class="term cur write-in" style="animation-delay:.4s">${neiW}H</span>
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
