/* ================================================================
 * IP 地址与子网掩码
 * 子页面：classify 地址分类 / mask 子网掩码 / network 网络地址 / hostrange 主机范围
 * 复用全局辅助函数：ipToArr, arrToIp, ipAnd, ipOr, ipNot, ipAdd, prefixToMask, ipClass
 * ================================================================ */
MOD.ip = {
  name:'IP与子网掩码', sub:'IPv4分类 · 子网掩码 · 网络地址 · 主机范围',
  badge:'网络 · 子网', tag:'网络', cat:'net',
  color:'#06B6D4', soft:'#E6F9FE', type:'IPv4分类 · 前缀掩码 · 按位与 · 广播与主机范围',
  desc:'IPv4 地址分类、子网掩码二进制推导、网络地址按位与计算、广播地址与可用主机范围及主机数，逐步手写演示。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M2 12h20M12 2a15 15 0 0 1 0 20M12 2a15 15 0 0 0 0 20"/></svg>',
  setDesc:'输入 IP 地址并选择子网前缀；各子页面独立演示对应计算。',
  stageDesc:'橙色=当前计算段，绿色=已得出结果，逐段展示按位与与地址范围。',
  initDesc:'点击「自动演示」或「下一步」，观看 IP 分类与子网计算的手写推导。',
  currentSub:'classify',
  subTabs:[
    {id:'classify', label:'地址分类', method:'按第一段判定'},
    {id:'mask',     label:'子网掩码', method:'前缀转二进制'},
    {id:'network',  label:'网络地址', method:'按位与'},
    {id:'hostrange',label:'主机范围', method:'广播与主机数'}
  ],
  theory(){
    return `<div class="rule"><b>IP 地址组成</b>：32 位二进制，分 4 段（每段 8 位），用点分十进制表示，如 192.168.1.10。</div>
    <div class="rule"><b>IP 分类</b>：A 类（1~126，默认掩码 255.0.0.0）、B 类（128~191，255.255.0.0）、C 类（192~223，255.255.255.0）、D 类组播（224~239）、E 类保留（240~255）。</div>
    <div class="rule"><b>子网掩码</b>：连续的 1 后跟连续的 0，1 对应网络位，0 对应主机位，用 /前缀长度 表示，如 /24。</div>
    <div class="rule"><b>网络地址</b> = IP 地址 <b>按位与</b> 子网掩码（主机位全 0）。</div>
    <div class="rule"><b>广播地址</b> = 网络地址的主机位全 1（网络地址 | 反掩码）。</div>
    <div class="rule"><b>可用主机范围</b>：网络地址 +1 ~ 广播地址 −1；可用主机数 = 2<sup>主机位数</sup> − 2（减去网络地址和广播地址）。</div>`;
  },
  /* ---------- 参数区 ---------- */
  settingsHTML(){
    const sub = this.currentSub;
    let html = '';
    if(sub !== 'mask'){
      const oldIp = $('#ip-addr');
      const cur = oldIp ? oldIp.value : (this._ip || '');
      html += `<div class="inline-setting"><div class="setting"><label>IP 地址</label>
        <input type="text" id="ip-addr" value="${esc(cur)}" placeholder="如 192.168.1.10" style="width:clamp(130px,20vw,170px)"></div></div>`;
    }
    if(sub !== 'classify'){
      const oldPf = $('#ip-prefix');
      const cur = oldPf ? oldPf.value : String(this._prefix || 24);
      const opts = [8,16,24,25,26,27,28,29,30];
      html += `<div class="inline-setting"><div class="setting"><label>子网前缀</label>
        <select id="ip-prefix">${opts.map(p=>`<option value="${p}" ${String(p)===cur?'selected':''}>/${p}</option>`).join('')}</select></div></div>`;
    }
    return html;
  },
  afterSettings(){
    const ipEl = $('#ip-addr');
    if(ipEl && !ipEl.value.trim()){
      this._ip = this._ip || this._randIp();
      ipEl.value = this._ip;
    }
    if(ipEl) ipEl.onchange = ()=>{ this._ip = ipEl.value.trim(); Engine.reload(); };
    const pfEl = $('#ip-prefix');
    if(pfEl) pfEl.onchange = ()=>{ this._prefix = +pfEl.value; Engine.reload(); };
  },
  /* ---------- 内部辅助 ---------- */
  segBin(v){ return v.toString(2).padStart(8,'0'); },
  /* 位权展开：返回位权数组、二进制位、求和结果、被激活的位权 */
  weightExpand(binStr){
    const weights = [128,64,32,16,8,4,2,1];
    const bits = binStr.split('');
    let sum = 0;
    const active = [];
    bits.forEach((b,i)=>{ if(b==='1'){ sum += weights[i]; active.push(weights[i]); } });
    return { weights, bits, sum, active };
  },
  /* 生成位权展开计算卡片 HTML */
  weightCard(binStr){
    const { weights, bits, sum, active } = this.weightExpand(binStr);
    const weightRow = weights.map(w=>`<div class="bd weight">${w}</div>`).join('');
    const bitRow = bits.map((b,i)=>`<div class="bd ${b==='1'?'one':'zero'} write-in" style="animation-delay:${i*0.06}s">${b}</div>`).join('');
    const valRow = bits.map((b,i)=>`<div class="bd ${b==='1'?'cur':'empty'}" style="animation-delay:${0.5+i*0.05}s">${b==='1'?weights[i]:'0'}</div>`).join('');
    const sumExpr = active.length ? active.join(' + ') : '0';
    return `<div class="calc-process">
      <div class="cp-label">位权展开求和（从左到右位权：128 → 1）</div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">
        <span style="width:34px;font-size:clamp(9px,1.1vw,10.5px);color:var(--text-3);font-weight:700;text-align:right;flex:none">位权</span>
        <div class="bin-grid" style="flex:1">${weightRow}</div>
      </div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">
        <span style="width:34px;font-size:clamp(9px,1.1vw,10.5px);color:var(--text-3);font-weight:700;text-align:right;flex:none">二进制</span>
        <div class="bin-grid" style="flex:1">${bitRow}</div>
      </div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:8px">
        <span style="width:34px;font-size:clamp(9px,1.1vw,10.5px);color:var(--text-3);font-weight:700;text-align:right;flex:none">取值</span>
        <div class="bin-grid" style="flex:1">${valRow}</div>
      </div>
      <div class="cp-sum-line">${sumExpr} <span class="eq">=</span> <span class="sum">${sum}</span></div>
    </div>`;
  },
  /* 生成逐位 AND 计算卡片 HTML */
  andCard(ipBin, maskBin, resBin){
    const ipBits = ipBin.split('');
    const mkBits = maskBin.split('');
    const rsBits = resBin.split('');
    const row = (label,cls,bits,delayBase)=>`<div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">
      <span style="width:34px;font-size:clamp(9px,1.1vw,10.5px);color:var(--text-3);font-weight:700;text-align:right;flex:none">${label}</span>
      <div class="bin-grid" style="flex:1">${bits.map((b,i)=>`<div class="bd ${cls} ${cls==='and-result'?'write-in':''}" style="${cls==='and-result'?`animation-delay:${delayBase+i*0.06}s`:''}">${b}</div>`).join('')}</div>
    </div>`;
    return `<div class="calc-process">
      <div class="cp-label">逐位按位与（AND）：只有 1 AND 1 = 1，其余均为 0</div>
      ${row('IP','one',ipBits,0)}
      ${row('掩码','mask',mkBits,0)}
      ${row('结果','and-result',rsBits,0.3)}
      <div style="text-align:center;font-size:clamp(10px,1.2vw,11.5px);color:var(--text-3);margin-top:2px">蓝=IP位 &nbsp; 青=掩码位 &nbsp; 橙=AND结果</div>
    </div>`;
  },
  _randIp(){
    return [randInt(1,254), randInt(0,255), randInt(0,255), randInt(1,254)].join('.');
  },
  _parseIp(){
    const el = $('#ip-addr');
    let s = el ? el.value.trim() : '';
    if(!s){ s = this._ip || this._randIp(); if(el) el.value = s; }
    const parts = s.split('.');
    if(parts.length !== 4) throw new Error('IP 必须为 4 段点分十进制，如 192.168.1.10');
    parts.forEach(p=>{
      if(!/^\d{1,3}$/.test(p)) throw new Error(`段「${p}」不是数字`);
      if(+p < 0 || +p > 255) throw new Error(`段「${p}」超出 0~255`);
    });
    const arr = ipToArr(s);
    this._ip = arrToIp(arr);
    return { str:this._ip, arr };
  },
  _readPrefix(){
    const el = $('#ip-prefix');
    let v = el ? el.value : '24';
    if(v === 'auto' || v === '' || v == null) v = 24;
    this._prefix = +v;
    return this._prefix;
  },
  typeChars(text){
    return [...String(text)].map((ch,i)=>`<span class="char-in" style="animation-delay:${i*0.06}s">${esc(ch)}</span>`).join('');
  },
  classRows(){
    return [
      {key:'A', range:'1 ~ 126', mask:'255.0.0.0', prefix:'/8',  note:''},
      {key:'B', range:'128 ~ 191', mask:'255.255.0.0', prefix:'/16', note:''},
      {key:'C', range:'192 ~ 223', mask:'255.255.255.0', prefix:'/24', note:''},
      {key:'D', range:'224 ~ 239', mask:'—', prefix:'—', note:'组播地址，无默认主机掩码'},
      {key:'E', range:'240 ~ 255', mask:'—', prefix:'—', note:'保留地址，无默认主机掩码'}
    ];
  },
  classifyOf(arr){
    const cls = ipClass(arr);              // ['A','255.0.0.0','/8'] …
    const key = cls[0][0];                // A/B/C/D/E（含括号标注时取首字母）
    const rows = this.classRows();
    const row = rows.find(r=>r.key===key) || rows[0];
    const note = cls[0].includes('(') ? '127.x 为回环地址（仍属 A 类保留段）' : row.note;
    return Object.assign({}, row, { mask:cls[1], prefix:cls[2], note });
  },
  ipRow(arr, curSeg, label, doneAll){
    return `<div class="ip-row">${arr.map((v,i)=>`<div class="seg ${i===curSeg?'cur':doneAll?'done':''} ${i===curSeg?'write-in':''}"><div class="seg-hd">${label||('第 '+(i+1)+' 段')}</div><div class="val">${v}</div></div>`).join('')}</div>`;
  },

  /* ========== ① classify：IPv4 地址分类 ========== */
  genClassify(){
    const { str, arr } = this._parseIp();
    const f = arr[0];
    const cat = this.classifyOf(arr);
    const rows = this.classRows();
    const label = 'IPv4 地址分类';
    const steps = [];
    // 1. 逐段写出 IP 地址
    steps.push({ desc:`写出待分类的 IPv4 地址：<b class="c-op">${str}</b>`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">准备</span></div>
          <div style="text-align:center;padding:clamp(18px,3vw,28px) 0">
            <span style="font-family:var(--mono);font-size:clamp(22px,4vw,32px);font-weight:700;color:var(--text)">${this.typeChars(str)}</span>
            <span class="caret pen"></span>
          </div>`;
      }});
    // 2. 高亮第一段
    steps.push({ desc:`IPv4 只依据<b>第一段</b>判断类别，当前第一段 = <b class="c-op">${f}</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">观察第一段</span></div>
          <div class="ip-row">${arr.map((v,i)=>`<div class="seg ${i===0?'cur':''} write-in" style="animation-delay:${i*0.08}s"><div class="seg-hd">第 ${i+1} 段</div><div class="val">${v}</div></div>`).join('')}</div>
          <div class="annot" style="text-align:center">分类只看<b>第一段</b>的取值范围，后三段不影响类别</div>`;
      }});
    // 3. 对照表判定类别
    steps.push({ desc:`对照分类范围：第一段 ${f} 落在 <b>${cat.range}</b>，属于 <b class="c-op">${cat.key} 类</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">对照分类表</span></div>
          <div class="step-title"><span class="st-dot"></span>IPv4 分类对照表（看第一段）</div>
          <table class="exp-table">
            <tr><th>类别</th><th>第一段范围</th><th>默认掩码</th><th>前缀</th></tr>
            ${rows.map((r,i)=>`<tr class="write-in" style="animation-delay:${i*0.07}s">
              <td class="${r.key===cat.key?'cur':'done'}" style="font-weight:700">${r.key} 类</td>
              <td class="${r.key===cat.key?'cur':''}">${r.range}</td>
              <td class="${r.key===cat.key?'cur':''}">${r.mask}</td>
              <td class="${r.key===cat.key?'cur':''}">${r.prefix}</td>
            </tr>`).join('')}
          </table>`;
      }});
    // 4. 写出默认掩码结果
    steps.push({ desc:`${cat.key} 类默认子网掩码 = <b class="c-op">${cat.mask}</b>${cat.prefix!=='—'?`（${cat.prefix}）`:''}${cat.note?`，${cat.note}`:''}。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          ${this.ipRow(arr,-1,'',true)}
          <div class="write-result">${str} → ${cat.key} 类 · 默认掩码 ${cat.mask}${cat.prefix!=='—'?` (${cat.prefix})`:''}</div>`;
      }});
    return steps;
  },

  /* ========== ② mask：子网掩码计算 ========== */
  genMask(){
    const prefix = this._readPrefix();
    const mask = prefixToMask(prefix);
    const hostBits = 32 - prefix;
    const bin = mask.map(v=>this.segBin(v));
    const label = '子网掩码计算';
    const steps = [];
    // 1. 写出前缀
    steps.push({ desc:`子网前缀 <b class="c-op">/${prefix}</b>：前 ${prefix} 位为网络位（1），后 ${hostBits} 位为主机位（0）。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">准备</span></div>
          <div style="text-align:center;padding:clamp(16px,3vw,24px) 0">
            <span style="font-family:var(--mono);font-size:clamp(22px,4vw,30px);font-weight:700;color:var(--text)">/${this.typeChars(String(prefix))}</span>
            <span class="caret pen"></span>
          </div>
          <div class="flow-line" style="justify-content:center">
            <span class="part">网络位 <b class="res">${prefix}</b> 位</span><span class="part">主机位 <b class="res">${hostBits}</b> 位</span>
          </div>`;
      }});
    // 2. 32 位二进制逐位填写（前 N 位 1，后 M 位 0）
    steps.push({ desc:`写出掩码二进制：前 <b>${prefix}</b> 位填 1（网络位），后 <b>${hostBits}</b> 位填 0（主机位）。`,
      draw:()=>{
        let bits = '';
        for(let oct=0; oct<4; oct++){
          for(let b=0; b<8; b++){
            const pos = oct*8+b;
            const isNet = pos < prefix;
            bits += `<div class="bd ${isNet?'mask':''} write-in" style="animation-delay:${pos*0.018}s">${isNet?'1':'0'}</div>`;
          }
        }
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">32 位二进制</span></div>
          <div class="step-title"><span class="st-dot"></span>二进制掩码（蓝=网络位 1，灰=主机位 0）</div>
          <div class="bin-grid">${bits}</div>`;
      }});
    // 3-6. 逐段二进制 → 十进制（位权展开求和）
    for(let i=0; i<4; i++){
      const oct = i;
      const { active, sum } = this.weightExpand(bin[oct]);
      const sumExpr = active.length ? active.join(' + ') : '0';
      steps.push({ desc:`第 ${oct+1} 段：二进制 <b class="c-op">${bin[oct]}</b>，按位权展开 ${sumExpr} = <b class="c-op">${mask[oct]}</b>。`,
        draw:()=>{
          const gridBits = mask.map((v,j)=>this.segBin(v).split('').map((b,k)=>
            `<div class="bd ${j<oct?'net':j===oct?'cur':'empty'} ${j===oct?'write-in':''}" style="${j===oct?`animation-delay:${k*0.05}s`:''}">${j<=oct?b:'?'}</div>`).join('')).join('');
          $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">第 ${oct+1}/4 段</span></div>
            <div class="step-title"><span class="st-dot"></span>二进制 → 点分十进制（位权展开）</div>
            <div class="bin-grid">${gridBits}</div>
            ${this.weightCard(bin[oct])}
            <div class="ip-row">${mask.map((v,j)=>`<div class="seg ${j<oct?'done':j===oct?'cur write-in':''}"><div class="seg-hd">第 ${j+1} 段</div><div class="val">${j<=oct?v:'?'}</div></div>`).join('')}</div>`;
        }});
    }
    // 7. 拼接点分十进制结果
    steps.push({ desc:`四段拼接得子网掩码：<b class="c-op">${arrToIp(mask)}</b>（/${prefix}），网络位 ${prefix} 位、主机位 ${hostBits} 位。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>子网掩码</div>
          ${this.ipRow(mask,-1,'掩码',true)}
          <div class="write-result">/${prefix} → 掩码 ${arrToIp(mask)} · 网络位 ${prefix} 位 / 主机位 ${hostBits} 位</div>`;
      }});
    return steps;
  },

  /* ========== ③ network：网络地址计算 ========== */
  genNetwork(){
    const { str, arr:ip } = this._parseIp();
    const prefix = this._readPrefix();
    const mask = prefixToMask(prefix);
    const net = ipAnd(ip, mask);
    const label = '网络地址计算';
    const steps = [];
    // 1. 写出 IP 与前缀
    steps.push({ desc:`IP 地址 <b class="c-op">${str}</b>，前缀 /${prefix}，对应掩码 ${arrToIp(mask)}。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">准备</span></div>
          <div style="text-align:center;padding:clamp(14px,2.5vw,22px) 0">
            <span style="font-family:var(--mono);font-size:clamp(20px,3.5vw,28px);font-weight:700;color:var(--text)">${this.typeChars(str)}<sub style="font-size:.45em;color:var(--text-3)"> /${prefix}</sub></span>
            <span class="caret pen"></span>
          </div>`;
      }});
    // 2. 方法说明，IP 与掩码并排
    steps.push({ desc:`<b>网络地址 = IP 按位与 子网掩码</b>：掩码为 1 的位保留原值，为 0 的位清 0（主机位全 0）。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">按位与</span></div>
          <div class="step-title"><span class="st-dot"></span>IP 与掩码并排</div>
          ${this.ipRow(ip,-1,'IP')}${this.ipRow(mask,-1,'掩码')}`;
      }});
    // 3-6. 逐段按位与（逐位 AND 详细过程）
    for(let i=0; i<4; i++){
      const r = ip[i] & mask[i];
      const oct = i;
      const ipBin = this.segBin(ip[oct]);
      const mkBin = this.segBin(mask[oct]);
      const rsBin = this.segBin(r);
      steps.push({ desc:`第 ${oct+1} 段按位与：${ipBin} AND ${mkBin} = ${rsBin}，即 ${ip[oct]} AND ${mask[oct]} = <b class="c-op">${r}</b>，网络地址已写出 ${oct+1} 段。`,
        draw:()=>{
          const table = `<table class="exp-table" style="margin-bottom:8px">
            <tr><th>项目</th>${[0,1,2,3].map(k=>`<th>第 ${k+1} 段</th>`).join('')}</tr>
            <tr><th>IP</th>${ip.map((v,k)=>`<td class="${k<oct?'done':k===oct?'cur':''} ${k===oct?'write-in':''}">${k<=oct?v:'?'}</td>`).join('')}</tr>
            <tr><th>掩码</th>${mask.map((v,k)=>`<td class="${k<oct?'done':k===oct?'cur':''} ${k===oct?'write-in':''}">${k<=oct?v:'?'}</td>`).join('')}</tr>
            <tr><th>AND 结果</th>${net.map((v,k)=>`<td class="${k<oct?'done':k===oct?'cur':'empty'} ${k===oct?'write-in':''}">${k<=oct?v:'?'}</td>`).join('')}</tr>
          </table>`;
          $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">第 ${oct+1}/4 段</span></div>
            <div class="step-title"><span class="st-dot"></span>逐段按位与（逐位 AND）</div>${table}
            ${this.andCard(ipBin, mkBin, rsBin)}
            <div class="exp-inline"><span class="term">${ipBin}</span><span class="op">AND</span><span class="term">${mkBin}</span><span class="op">=</span><span class="term cur">${rsBin}</span><span class="op">=</span><span class="term cur">${r}</span></div>`;
        }});
    }
    // 7. 网络地址结果
    steps.push({ desc:`网络地址 = <b class="c-op">${arrToIp(net)}</b>（主机位全 0），/${prefix}。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">完成</span></div>
          <div class="step-title"><span class="st-dot"></span>网络地址</div>
          ${this.ipRow(net,-1,'网络地址',true)}
          <div class="write-result">网络地址 = ${arrToIp(net)} /${prefix}</div>`;
      }});
    return steps;
  },

  /* ========== ④ hostrange：主机范围与数量 ========== */
  genHostrange(){
    const { str, arr:ip } = this._parseIp();
    const prefix = this._readPrefix();
    const mask = prefixToMask(prefix);
    const net = ipAnd(ip, mask);
    const notMask = ipNot(mask);
    const broad = ipOr(net, notMask);
    const first = ipAdd(net, 1);
    const last = ipAdd(broad, -1);
    const hostBits = 32 - prefix;
    const count = Math.pow(2, hostBits) - 2;
    const label = '主机范围与数量';
    const steps = [];
    // 1. 网络地址
    steps.push({ desc:`先求网络地址：${str} AND ${arrToIp(mask)} = <b class="c-op">${arrToIp(net)}</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">网络地址</span></div>
          <div class="step-title"><span class="st-dot"></span>网络地址（主机位全 0）</div>
          <div class="exp-inline" style="margin-bottom:6px"><span class="term">${str}</span><span class="op">AND</span><span class="term">${arrToIp(mask)}</span><span class="op">=</span><span class="term cur">${arrToIp(net)}</span></div>
          ${this.ipRow(net,-1,'网络地址',true)}
          <div class="write-result">网络地址 = ${arrToIp(net)} /${prefix}</div>`;
      }});
    // 2. 广播地址：主机位全 1（逐位 OR 过程）
    steps.push({ desc:`广播地址 = 网络地址的<b>主机位全 1</b>：网络地址 OR 反掩码，主机位 0→1，网络位保持不变 = <b class="c-op">${arrToIp(broad)}</b>。`,
      draw:()=>{
        /* 找出主机位所在的段范围 */
        const netOct = Math.floor(prefix / 8);
        const netBitInOct = prefix % 8;
        /* 构建32位二进制三行展示 */
        const bitRow32 = (arr, cls, highlightHost, delayBase)=>{
          let html = '';
          for(let oct=0; oct<4; oct++){
            const bits = this.segBin(arr[oct]).split('');
            bits.forEach((b,k)=>{
              const pos = oct*8+k;
              const isHost = pos >= prefix;
              const hl = highlightHost && isHost;
              html += `<div class="bd ${cls} ${hl?'cur':''} ${cls==='and-result'?'write-in':''}" style="${cls==='and-result'?`animation-delay:${delayBase+pos*0.025}s`:''}">${b}</div>`;
            });
          }
          return html;
        };
        const row32 = (label, arr, cls, hl, delay)=>`<div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">
          <span style="width:42px;font-size:clamp(9px,1.1vw,10.5px);color:var(--text-3);font-weight:700;text-align:right;flex:none">${label}</span>
          <div class="bin-grid" style="flex:1;grid-template-columns:repeat(32,1fr);gap:2px">${bitRow32(arr,cls,hl,delay)}</div>
        </div>`;
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">广播地址</span></div>
          <div class="step-title"><span class="st-dot"></span>主机位全 1 → 广播地址（32 位逐位 OR）</div>
          <div class="calc-process">
            <div class="cp-label">网络位（前 ${prefix} 位）保持不变，主机位（后 ${hostBits} 位）全部置 1</div>
            ${row32('网络地址', net, 'one', false, 0)}
            ${row32('反掩码', notMask, 'mask', false, 0)}
            ${row32('广播地址', broad, 'and-result', true, 0.3)}
            <div style="text-align:center;font-size:clamp(10px,1.2vw,11.5px);color:var(--text-3);margin-top:4px">橙色高亮 = 主机位，由 0 全部变为 1</div>
          </div>
          <table class="exp-table" style="margin-bottom:8px">
            <tr><th>项目</th>${[0,1,2,3].map(k=>`<th>段 ${k+1}</th>`).join('')}</tr>
            <tr><th>网络地址</th>${net.map(v=>`<td class="done">${v}</td>`).join('')}</tr>
            <tr><th>反掩码</th>${notMask.map(v=>`<td>${v}</td>`).join('')}</tr>
            <tr><th>广播地址</th>${broad.map(v=>`<td class="cur">${v}</td>`).join('')}</tr>
          </table>
          <div class="exp-inline"><span class="term done">${arrToIp(net)}</span><span class="op">OR</span><span class="term">${arrToIp(notMask)}</span><span class="op">=</span><span class="term cur">${arrToIp(broad)}</span></div>`;
      }});
    // 3. 可用主机范围
    steps.push({ desc:`第一个可用主机 = 网络地址 + 1 = <b class="c-op">${arrToIp(first)}</b>；最后一个 = 广播地址 − 1 = <b class="c-op">${arrToIp(last)}</b>。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">可用范围</span></div>
          <div class="step-title"><span class="st-dot"></span>排除网络地址与广播地址</div>
          <div class="exp-inline" style="margin-bottom:6px"><span class="term done">${arrToIp(net)}</span><span class="op">+1 =</span><span class="term cur">${arrToIp(first)}</span></div>
          <div class="exp-inline" style="margin-bottom:6px"><span class="term done">${arrToIp(broad)}</span><span class="op">−1 =</span><span class="term cur">${arrToIp(last)}</span></div>
          <div class="write-result">可用主机范围：${arrToIp(first)} ~ ${arrToIp(last)}</div>`;
      }});
    // 4. 可用主机数 = 2^主机位 − 2（幂运算展开）
    steps.push({ desc:`主机位 = ${hostBits} 位，2<sup>${hostBits}</sup> = ${Math.pow(2,hostBits).toLocaleString()}，减去网络地址和广播地址各 1 个，可用主机数 = <b class="c-op">${count.toLocaleString()}</b>。`,
      draw:()=>{
        const total = Math.pow(2, hostBits);
        /* 2^n 连乘展开：n<=12 时展示完整连乘，否则展示分组 */
        let expandExpr;
        if(hostBits <= 12){
          expandExpr = Array(hostBits).fill('2').join(' × ');
        } else {
          const half = Math.ceil(hostBits/2);
          expandExpr = `2 × 2 × … × 2（${hostBits} 个 2 相乘）`;
        }
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">可用主机数</span></div>
          <div class="step-title"><span class="st-dot"></span>2<sup>主机位</sup> − 2（幂运算展开）</div>
          <div class="calc-process">
            <div class="cp-label">每一位主机位有 0/1 两种可能，${hostBits} 位共有 2<sup>${hostBits}</sup> 种组合</div>
            <div class="cp-sum-line" style="margin-bottom:8px">
              2<sup>${hostBits}</sup> = ${expandExpr} = <span class="sum">${total.toLocaleString()}</span>
            </div>
            <div class="cp-sum-line">
              ${total.toLocaleString()} <span class="eq">−</span> 2（网络地址+广播地址） = <span class="sum">${count.toLocaleString()} 台</span>
            </div>
          </div>
          <table class="exp-table" style="margin-bottom:8px">
            <tr><th>主机位数</th><th>2<sup>主机位</sup>（总地址数）</th><th>减网络/广播</th><th>可用主机数</th></tr>
            <tr><td class="cur">${hostBits}</td><td>${total.toLocaleString()}</td><td>−2</td><td class="done">${count.toLocaleString()}</td></tr>
          </table>`;
      }});
    // 5. 汇总
    steps.push({ desc:`汇总：${str}/${prefix} 的网络地址、广播地址、可用范围与主机数量。`,
      draw:()=>{
        $('#stage').innerHTML = `<div class="stage-hd"><span>${label}</span><span class="state">汇总</span></div>
          <div class="ip-info">
            <div class="ii"><div class="ii-hd">网络地址</div><div class="ii-val">${arrToIp(net)}</div></div>
            <div class="ii"><div class="ii-hd">广播地址</div><div class="ii-val">${arrToIp(broad)}</div></div>
            <div class="ii"><div class="ii-hd">可用主机范围</div><div class="ii-val">${arrToIp(first)} ~ ${arrToIp(last)}</div></div>
            <div class="ii"><div class="ii-hd">可用主机数</div><div class="ii-val">${count.toLocaleString()} 台</div></div>
          </div>
          <div class="write-result">${arrToIp(net)} ~ ${arrToIp(broad)} · 可用 ${count.toLocaleString()} 台（${arrToIp(first)} ~ ${arrToIp(last)}）</div>`;
      }});
    return steps;
  },

  /* ---------- 初始化画布 ---------- */
  initStage(){
    const sub = this.currentSub;
    const info = this.subTabs.find(t=>t.id===sub) || {};
    const labelMap = { classify:'IPv4 地址分类', mask:'子网掩码计算', network:'网络地址计算', hostrange:'主机范围与数量' };
    $('#stage').innerHTML = `<div class="stage-hd"><span>${labelMap[sub]||'IP 与子网'}</span><span class="state">准备就绪</span></div>
      <div style="text-align:center;padding:clamp(16px,3vw,28px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.8vw,16px);margin-bottom:8px">使用 <b style="color:#06B6D4">${info.method||''}</b></div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">点击「自动演示」或「下一步」开始手写推导</div>
      </div>`;
  },
  generate(){
    switch(this.currentSub){
      case 'classify': this.steps = this.genClassify(); break;
      case 'mask':     this.steps = this.genMask(); break;
      case 'network':  this.steps = this.genNetwork(); break;
      case 'hostrange':this.steps = this.genHostrange(); break;
      default: this.steps = this.genClassify();
    }
  }
};
