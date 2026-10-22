

;;;

/* ================================================================
 * ③ 音视频容量
 * ================================================================ */
/* ================================================================
 * 音视频容量（未压缩文件大小计算）
 * 子页面：audio · video
 * ================================================================ */
/* ================================================================
 * 音视频容量（未压缩文件大小计算）
 * 子页面：audio · video
 * ================================================================ */
/* ================================================================
 * 音视频容量（未压缩文件大小计算）
 * 子页面：audio · video
 * ================================================================ */
/* ================================================================
 * 音视频容量（未压缩文件大小计算）
 * 子页面：audio 音频 / video 视频 / image 图像
 * 步骤覆盖展示：每一步只渲染当前步骤内容（橙色高亮），历史步骤由「上一步」回退重绘
 * 采样率支持 Hz / kHz / MHz 单位切换，内部统一归一化为 Hz 计算
 * ================================================================ */
if(!document.getElementById('mod-media-css')){
  document.head.insertAdjacentHTML('beforeend', `<style id="mod-media-css">
    .mm-wrap{display:flex;flex-direction:column;gap:clamp(10px,2vw,16px)}
    .mm-step{border:1px solid var(--border);border-radius:var(--radius);background:#fff;padding:clamp(6px,1vw,10px)}
    .mm-step.done{opacity:.72;background:var(--surface-2);border-left:4px solid var(--green)}
    .mm-step.cur{border:2px solid var(--orange);background:#FFFCF5;box-shadow:0 4px 18px rgba(245,158,11,.12)}
    .mm-step-hd{display:flex;align-items:center;gap:8px;margin-bottom:clamp(4px,.8vw,6px);font-weight:700;color:var(--text);font-size:clamp(12px,1.5vw,15px)}
    .mm-step-hd .mm-idx{display:inline-flex;align-items:center;justify-content:center;width:clamp(20px,2.4vw,25px);height:clamp(20px,2.4vw,25px);border-radius:50%;font-family:var(--mono);font-size:clamp(11px,1.4vw,13px);font-weight:700;background:var(--surface-3);color:var(--text-3);flex:none}
    .mm-step.cur .mm-step-hd .mm-idx{background:var(--orange);color:#fff}
    .mm-step.done .mm-step-hd .mm-idx{background:var(--green);color:#fff}
    .mm-step-hd .mm-status{margin-left:auto;font-size:clamp(10px,1.2vw,11.5px);font-weight:600}
    .mm-step.done .mm-status{color:var(--green)}
    .mm-step.cur .mm-status{color:var(--orange)}
    .mm-body{font-family:var(--mono);font-size:clamp(14px,1.9vw,18px);line-height:1.6;color:var(--text-2)}
    .mm-num{font-size:clamp(16px,2.4vw,22px);font-weight:800;color:var(--accent)}
    .mm-num.hot{color:var(--orange)}
    .mm-num.g{color:var(--green)}
    .mm-op{color:var(--text-3);margin:0 .35em;font-weight:600}
    .mm-chips{display:flex;flex-wrap:wrap;gap:clamp(5px,.9vw,8px);margin:clamp(4px,.8vw,7px) 0}
    .mm-chip{background:var(--surface-3);border:1px solid var(--border);border-radius:10px;padding:clamp(4px,.8vw,7px) clamp(8px,1.3vw,12px);font-family:var(--mono);font-size:clamp(13px,1.7vw,16px);font-weight:700;color:var(--text)}
    .mm-chip .mm-chip-l{display:block;font-family:var(--font);font-size:clamp(10px,1.2vw,11.5px);color:var(--text-3);font-weight:600;margin-bottom:1px}
    .mm-chip.done{background:var(--green-soft);border-color:var(--green);color:#15803D}
    .mm-formula{background:var(--surface-2);border:1px dashed var(--border);border-radius:10px;padding:clamp(5px,1vw,8px);margin-top:clamp(5px,.9vw,8px);font-family:var(--font);font-size:clamp(12px,1.6vw,15px);color:var(--text-3);line-height:1.6}
    .mm-result{margin-top:clamp(6px,1vw,10px);font-family:var(--mono);font-size:clamp(15px,2.1vw,20px);font-weight:800;color:var(--accent);background:linear-gradient(135deg,#EEF3FF,#E0F2FE);border:1px solid #C7D2FE;border-radius:12px;padding:clamp(6px,1.2vw,10px);text-align:center;line-height:1.5}
    .mm-note{font-family:var(--font);font-size:clamp(11.5px,1.5vw,13.5px);color:var(--text-3);line-height:1.5;margin-top:3px}
    .mm-conv{display:flex;flex-direction:column;gap:clamp(4px,.8vw,7px);margin:clamp(5px,.9vw,8px) 0}
    .mm-conv div{background:#fff;border:1px solid var(--border);border-radius:10px;padding:clamp(4px,.9vw,8px) clamp(8px,1.3vw,12px);font-family:var(--mono);font-size:clamp(14px,1.9vw,18px);font-weight:700;color:var(--text-2);line-height:1.6}
  </style>`);
}

MOD.media = {
  name:'音视频容量', sub:'音频 / 视频 / 图像 文件大小（未压缩）计算',
  badge:'存储 · 公式', tag:'容量', cat:'store',
  color:'#38BDF8', soft:'#EAF7FF', type:'音频采样 · 视频帧率 · 像素位深 · 字节换算',
  desc:'根据时长、采样率、位深、声道、帧率与分辨率等参数，逐步代入公式计算未压缩音视频文件的容量并换算单位；采样率支持 Hz/kHz/MHz 自定义单位。',
  icon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M10 8.5v7l6-3.5z"/></svg>',
  setDesc:'',
  stageDesc:'每步只显示当前计算内容，点击「上一步」可回看之前步骤；关键数字放大显示。',
  initDesc:'点击「自动演示」或「下一步」，一步步代入公式算出文件容量。',
  currentSub:'audio',
  subTabs:[
    {id:'audio',label:'音频容量',method:'采样率×位深×声道×时长÷8'},
    {id:'video',label:'视频容量',method:'帧率×宽×高×位深×时长÷8'},
    {id:'image',label:'图像编码',method:'采样·量化·宽×高×位深÷8'}
  ],
  theory(){
    return `<div class="rule"><b>音频容量</b> = 采样率(Hz) × 采样位数(bit) × 声道数 × 时长(s) ÷ 8（字节）。</div>
    <div class="rule"><b>采样率单位换算</b>：1 kHz = 1000 Hz，1 MHz = 1,000,000 Hz；计算前先统一换算成 Hz。常见采样率：电话 8 kHz、广播 22.05 kHz、CD 44.1 kHz、专业 48 kHz。</div>
    <div class="rule"><b>视频容量</b> = 时长(s) × 帧率(fps) × 宽 × 高 × 位深(bit) ÷ 8（字节）。</div>
    <div class="rule"><b>图像容量</b> = 宽 × 高 × 位深(bit) ÷ 8（字节）；位深=每个像素的颜色位数，24 位=真彩色。</div>
    <div class="rule"><b>图像编码三步</b>：①采样（把连续图像分成像素网格）②量化（把颜色值离散化为有限级别）③编码（按位存储像素数据）。</div>
    <div class="rule"><b>单位换算</b>：1 KB = 1024 B，1 MB = 1024 KB，1 GB = 1024 MB；除以 8 是把 bit 换成字节。</div>`;
  },

  /* ---------- 参数区 ---------- */
  settingsHTML(){
    const sub=this.currentSub;
    if(sub==='audio'){
      const v=(id,d)=>{const el=$(id);return el?el.value:String(d);};
      const curT=v('#md-a-time',60), curSR=v('#md-a-sr-val',44100), curSU=v('#md-a-sr-unit','Hz'),
            curB=v('#md-a-bits',16), curC=v('#md-a-ch',2);
      return `<div class="inline-setting"><div class="setting"><label>时长（秒）</label>
        <input type="number" id="md-a-time" value="${esc(curT)}" min="0.001" step="any" placeholder="如 60" style="width:clamp(80px,12vw,110px)"></div></div>
      <div class="inline-setting"><div class="setting"><label>采样率数值</label>
        <div style="display:flex;gap:5px;align-items:center">
          <input type="number" id="md-a-sr-val" value="${esc(curSR)}" min="0" step="any" placeholder="如 44100" style="width:clamp(86px,13vw,118px)">
          <select id="md-a-sr-unit" style="min-width:64px">
            <option value="Hz" ${curSU==='Hz'?'selected':''}>Hz</option>
            <option value="kHz" ${curSU==='kHz'?'selected':''}>kHz</option>
            <option value="MHz" ${curSU==='MHz'?'selected':''}>MHz</option>
          </select>
        </div></div></div>
      <div class="inline-setting"><div class="setting"><label>采样位数（bit）</label>
        <input type="number" id="md-a-bits" value="${esc(curB)}" min="1" step="1" placeholder="如 16" style="width:clamp(70px,10vw,95px)"></div></div>
      <div class="inline-setting"><div class="setting"><label>声道数</label>
        <input type="number" id="md-a-ch" value="${esc(curC)}" min="1" step="1" placeholder="如 2" style="width:clamp(64px,9vw,88px)"></div></div>`;
    }
    if(sub==='video'){
      const v=(id,d)=>{const el=$(id);return el?el.value:String(d);};
      const curT=v('#md-v-time',10), curF=v('#md-v-fps',24), curW=v('#md-v-w',1280),
            curH=v('#md-v-h',720), curB=v('#md-v-bits',24);
      return `<div class="inline-setting"><div class="setting"><label>时长（秒）</label>
        <input type="number" id="md-v-time" value="${esc(curT)}" min="0.001" step="any" placeholder="如 10" style="width:clamp(70px,10vw,95px)"></div></div>
      <div class="inline-setting"><div class="setting"><label>帧率（fps）</label>
        <input type="number" id="md-v-fps" value="${esc(curF)}" min="1" step="1" placeholder="如 24" style="width:clamp(70px,10vw,95px)"></div></div>
      <div class="inline-setting"><div class="setting"><label>宽度（px）</label>
        <input type="number" id="md-v-w" value="${esc(curW)}" min="1" step="1" placeholder="如 1280" style="width:clamp(76px,11vw,104px)"></div></div>
      <div class="inline-setting"><div class="setting"><label>高度（px）</label>
        <input type="number" id="md-v-h" value="${esc(curH)}" min="1" step="1" placeholder="如 720" style="width:clamp(76px,11vw,104px)"></div></div>
      <div class="inline-setting"><div class="setting"><label>位深（bit）</label>
        <input type="number" id="md-v-bits" value="${esc(curB)}" min="1" step="1" placeholder="如 24" style="width:clamp(70px,10vw,95px)"></div></div>`;
    }
    if(sub==='image'){
      const v=(id,d)=>{const el=$(id);return el?el.value:String(d);};
      const curW=v('#md-i-w',800), curH=v('#md-i-h',600), curB=v('#md-i-bits',24);
      return `<div class="inline-setting"><div class="setting"><label>宽度（px）</label>
        <input type="number" id="md-i-w" value="${esc(curW)}" min="1" step="1" placeholder="如 800" style="width:clamp(70px,10vw,95px)"></div></div>
      <div class="inline-setting"><div class="setting"><label>高度（px）</label>
        <input type="number" id="md-i-h" value="${esc(curH)}" min="1" step="1" placeholder="如 600" style="width:clamp(70px,10vw,95px)"></div></div>
      <div class="inline-setting"><div class="setting"><label>位深（bit）</label>
        <select id="md-i-bits">
          <option value="1" ${curB==='1'?'selected':''}>1 bit（黑白）</option>
          <option value="8" ${curB==='8'?'selected':''}>8 bit（256 色）</option>
          <option value="24" ${curB==='24'?'selected':''}>24 bit（真彩）</option>
        </select></div></div>`;
    }
    return '';
  },
  afterSettings(){
    const bind=id=>{ const el=$('#'+id); if(el) el.onchange=()=>Engine.reload(); };
    ['md-a-time','md-a-sr-val','md-a-sr-unit','md-a-bits','md-a-ch',
     'md-v-time','md-v-fps','md-v-w','md-v-h','md-v-bits',
     'md-i-w','md-i-h','md-i-bits'].forEach(bind);
  },

  /* ---------- 内部辅助 ---------- */
  _fmt(n){
    n=Number(n);
    if(!isFinite(n)) return '—';
    return Number.isInteger(n) ? n.toLocaleString('en-US') : n.toLocaleString('en-US',{maximumFractionDigits:2});
  },
  _num(id, def, label, min, integer){
    const el=$(id); let raw=el?el.value.trim():'';
    if(raw==='') raw=String(def);
    if(!/^\d+(?:\.\d+)?$/.test(raw)) throw new Error(`${label}需为非负数字，当前输入「${raw}」`);
    const n=Number(raw);
    if(!isFinite(n)||n<min) throw new Error(`${label}需 ≥ ${min}`);
    if(integer && !Number.isInteger(n)) throw new Error(`${label}需为整数`);
    return n;
  },
  /* 字节 → 易读单位 */
  humanSize(bytes){
    if(bytes>=1073741824) return [bytes/1073741824,'GB'];
    if(bytes>=1048576) return [bytes/1048576,'MB'];
    if(bytes>=1024) return [bytes/1024,'KB'];
    return [bytes,'B'];
  },
  /* 单位换算过程行（B → KB → MB → GB，逐级展示） */
  _convLines(bytes){
    const out=[];
    if(bytes>=1024){
      const kb=bytes/1024;
      out.push(`<div>${this._fmt(bytes)} B <span class="mm-op">÷ 1024 =</span> <span class="mm-num g">${this._fmt(kb)} KB</span></div>`);
      if(bytes>=1048576){
        const mb=kb/1024;
        out.push(`<div>${this._fmt(kb)} KB <span class="mm-op">÷ 1024 =</span> <span class="mm-num hot">${this._fmt(mb)} MB</span></div>`);
        if(bytes>=1073741824){
          const gb=mb/1024;
          out.push(`<div>${this._fmt(mb)} MB <span class="mm-op">÷ 1024 =</span> <span class="mm-num hot">${this._fmt(gb)} GB</span></div>`);
        }
      }
    }
    return out.join('');
  },
  /* 覆盖式步骤画布：blocks 为 {title, desc, body} 数组；第 i 步只渲染当前步骤，历史步骤由 Engine 回退重绘 */
  _renderStep(label, statePrefix, blocks, i){
    const bk=blocks[i];
    $('#stage').innerHTML=`<div class="stage-hd"><span>${label}</span><span class="state">${statePrefix} 第 ${i+1}/${blocks.length} 步</span></div>
      <div class="mm-wrap">
        <div class="mm-step cur write-in">
          <div class="mm-step-hd"><span class="mm-idx">${i+1}</span>${bk.title}
            <span class="mm-status">进行中</span></div>
          <div class="mm-note" style="margin-bottom:clamp(6px,1vw,10px)">${bk.desc}</div>
          <div class="mm-body">${bk.body()}</div>
        </div>
      </div>`;
  },
  _stepsFromBlocks(label, statePrefix, blocks){
    return blocks.map((bk,i)=>({
      desc:bk.desc,
      draw:()=>this._renderStep(label, statePrefix, blocks, i)
    }));
  },

  /* ========== ① 音频容量 ========== */
  genAudio(){
    const t=this._num('#md-a-time',60,'时长',0.001);
    const srVal=this._num('#md-a-sr-val',44100,'采样率数值',0.000001);
    const unitEl=$('#md-a-sr-unit');
    const srUnit=unitEl?unitEl.value:(this._srUnit||'Hz');
    this._srUnit=srUnit;
    const factor=srUnit==='MHz'?1000000:srUnit==='kHz'?1000:1;
    const srHz=srVal*factor;
    const bits=this._num('#md-a-bits',16,'采样位数',1,true);
    const ch=this._num('#md-a-ch',2,'声道数',1,true);

    const bitPerSec=srHz*bits;            // 每秒 bit（单声道）
    const bitPerSecCh=bitPerSec*ch;       // 每秒 bit（全部声道）
    const bitsTotal=bitPerSecCh*t;        // 总 bit
    const bytes=bitsTotal/8;
    const h=this.humanSize(bytes);
    const label='音频文件（未压缩）';
    const srHzTxt=this._fmt(srHz);

    const blocks=[
      { title:'已知参数与公式',
        desc:`已知参数：时长 <b class="c-op">${this._fmt(t)} 秒</b>、采样率 <b class="c-op">${this._fmt(srVal)} ${srUnit}</b>${srUnit!=='Hz'?`（= ${srHzTxt} Hz）`:''}、采样位数 <b class="c-op">${bits} bit</b>、声道数 <b class="c-op">${ch}</b>；音频容量 = 采样率(Hz) × 采样位数(bit) × 声道数 × 时长(s) ÷ 8（字节）。`,
        body:()=>`
          <div class="mm-chips">
            <span class="mm-chip done"><span class="mm-chip-l">时长</span>${this._fmt(t)} s</span>
            <span class="mm-chip done"><span class="mm-chip-l">采样率</span>${this._fmt(srVal)} ${srUnit}${srUnit!=='Hz'?`<div class="mm-note" style="margin:0">= ${srHzTxt} Hz</div>`:''}</span>
            <span class="mm-chip done"><span class="mm-chip-l">采样位数</span>${bits} bit</span>
            <span class="mm-chip done"><span class="mm-chip-l">声道数</span>${ch}</span>
          </div>
          <div class="mm-formula">采样率(Hz) <span class="mm-op">×</span> 采样位数(bit) <span class="mm-op">×</span> 声道数 <span class="mm-op">×</span> 时长(s) <span class="mm-op">=</span> 总 bit 数；总 bit <span class="mm-op">÷ 8</span> = 字节数</div>
          <div class="mm-note">单位换算：1 kHz = 1000 Hz，1 MHz = 1,000,000 Hz，内部统一换算为 Hz 后再参与计算。</div>`
      },
      { title:'采样率 × 采样位数',
        desc:`采样率 <b class="c-op">${srHzTxt} Hz</b>（每秒采样 ${srHzTxt} 次）× 采样位数 <b class="c-op">${bits} bit</b>/采样点 = <b class="c-op">${this._fmt(bitPerSec)}</b> bit/s（单声道每秒所需 bit）。`,
        body:()=>`
          <div><span class="mm-num g">${srHzTxt} Hz</span><span class="mm-op">×</span><span class="mm-num hot">${bits} bit</span><span class="mm-op">=</span><span class="mm-num">${this._fmt(bitPerSec)}</span> <span style="font-size:clamp(13px,1.8vw,17px);color:var(--text-3)">bit/s</span></div>
          <div class="mm-note">每个采样点用 ${bits} bit 表示，每秒采样 ${srHzTxt} 次。</div>`
      },
      { title:'× 声道数',
        desc:`再乘声道数 <b class="c-op">${ch}</b>：${this._fmt(bitPerSec)} × ${ch} = <b class="c-op">${this._fmt(bitPerSecCh)}</b> bit/s（全部声道每秒总数据量）。`,
        body:()=>`
          <div><span class="mm-num g">${this._fmt(bitPerSec)}</span><span class="mm-op">×</span><span class="mm-num hot">${ch}</span><span class="mm-op">=</span><span class="mm-num">${this._fmt(bitPerSecCh)}</span> <span style="font-size:clamp(13px,1.8vw,17px);color:var(--text-3)">bit/s</span></div>
          <div class="mm-note">${ch} 个声道同时采样，每秒数据量为单声道的 ${ch} 倍。</div>`
      },
      { title:'× 时长',
        desc:`再乘时长 <b class="c-op">${this._fmt(t)} s</b>：${this._fmt(bitPerSecCh)} × ${this._fmt(t)} = <b class="c-op">${this._fmt(bitsTotal)}</b> bit（文件总数据量）。`,
        body:()=>`
          <div><span class="mm-num g">${this._fmt(bitPerSecCh)}</span><span class="mm-op">×</span><span class="mm-num hot">${this._fmt(t)} s</span><span class="mm-op">=</span><span class="mm-num">${this._fmt(bitsTotal)}</span> <span style="font-size:clamp(13px,1.8vw,17px);color:var(--text-3)">bit</span></div>
          <div class="mm-note">每秒 bit 数 × 时长（秒）= 文件总 bit 数。</div>`
      },
      { title:'bit ÷ 8 = 字节',
        desc:`1 字节 = 8 bit，${this._fmt(bitsTotal)} ÷ 8 = <b class="c-op">${this._fmt(bytes)}</b> 字节（B）。`,
        body:()=>`
          <div><span class="mm-num g">${this._fmt(bitsTotal)} bit</span><span class="mm-op">÷ 8 =</span><span class="mm-num hot">${this._fmt(bytes)} B</span></div>
          <div class="mm-note">除以 8 是把 bit 换算为字节。</div>`
      },
      { title:'容量单位换算与结果',
        desc:`总容量 = <b class="c-op">${this._fmt(bytes)}</b> 字节 ≈ <b class="c-op">${h[0].toFixed(2)} ${h[1]}</b>（1 KB = 1024 B）。`,
        body:()=>`
          <div class="mm-conv">${this._convLines(bytes)}</div>
          <div class="mm-result">${this._fmt(t)} s × ${srHzTxt} Hz × ${bits} bit × ${ch} 声道 ÷ 8<br>= ${this._fmt(bytes)} B ≈ ${h[0].toFixed(2)} ${h[1]}</div>`
      }
    ];
    return this._stepsFromBlocks(label,'音频计算',blocks);
  },

  /* ========== ② 视频容量 ========== */
  genVideo(){
    const t=this._num('#md-v-time',10,'时长',0.001);
    const fps=this._num('#md-v-fps',24,'帧率',1,true);
    const w=this._num('#md-v-w',1280,'宽度',1,true);
    const h=this._num('#md-v-h',720,'高度',1,true);
    const bits=this._num('#md-v-bits',24,'位深',1,true);

    const pxPerFrame=w*h;
    const bitsPerFrame=pxPerFrame*bits;
    const bitsPerSec=bitsPerFrame*fps;
    const bitsTotal=bitsPerSec*t;
    const bytes=bitsTotal/8;
    const hh=this.humanSize(bytes);
    const label='视频文件（未压缩）';

    const blocks=[
      { title:'已知参数与公式',
        desc:`已知参数：时长 <b class="c-op">${this._fmt(t)} 秒</b>、帧率 <b class="c-op">${fps} fps</b>、分辨率 <b class="c-op">${w}×${h}</b>、位深 <b class="c-op">${bits} bit</b>；视频容量 = 时长(s) × 帧率(fps) × 宽 × 高 × 位深(bit) ÷ 8（字节）。`,
        body:()=>`
          <div class="mm-chips">
            <span class="mm-chip done"><span class="mm-chip-l">时长</span>${this._fmt(t)} s</span>
            <span class="mm-chip done"><span class="mm-chip-l">帧率</span>${fps} fps</span>
            <span class="mm-chip done"><span class="mm-chip-l">分辨率</span>${w} × ${h} px</span>
            <span class="mm-chip done"><span class="mm-chip-l">位深</span>${bits} bit</span>
          </div>
          <div class="mm-formula">时长(s) <span class="mm-op">×</span> 帧率(fps) <span class="mm-op">×</span> 宽 <span class="mm-op">×</span> 高 <span class="mm-op">×</span> 位深(bit) <span class="mm-op">=</span> 总 bit 数；总 bit <span class="mm-op">÷ 8</span> = 字节数</div>`
      },
      { title:'宽 × 高 = 每帧像素',
        desc:`宽 ${w} px × 高 ${h} px = <b class="c-op">${this._fmt(pxPerFrame)}</b> 像素/帧（每帧画面的像素总数）。`,
        body:()=>`
          <div><span class="mm-num g">${w} px</span><span class="mm-op">×</span><span class="mm-num hot">${h} px</span><span class="mm-op">=</span><span class="mm-num">${this._fmt(pxPerFrame)}</span> <span style="font-size:clamp(13px,1.8vw,17px);color:var(--text-3)">像素/帧</span></div>`
      },
      { title:'× 位深 = 每帧 bit',
        desc:`每帧像素 × 位深 ${bits} bit：${this._fmt(pxPerFrame)} × ${bits} = <b class="c-op">${this._fmt(bitsPerFrame)}</b> bit/帧。`,
        body:()=>`
          <div><span class="mm-num g">${this._fmt(pxPerFrame)}</span><span class="mm-op">×</span><span class="mm-num hot">${bits} bit</span><span class="mm-op">=</span><span class="mm-num">${this._fmt(bitsPerFrame)}</span> <span style="font-size:clamp(13px,1.8vw,17px);color:var(--text-3)">bit/帧</span></div>
          <div class="mm-note">每个像素用 ${bits} bit 表示颜色。</div>`
      },
      { title:'× 帧率 = 每秒 bit',
        desc:`再乘帧率 ${fps} 帧/秒：${this._fmt(bitsPerFrame)} × ${fps} = <b class="c-op">${this._fmt(bitsPerSec)}</b> bit/s。`,
        body:()=>`
          <div><span class="mm-num g">${this._fmt(bitsPerFrame)}</span><span class="mm-op">×</span><span class="mm-num hot">${fps} fps</span><span class="mm-op">=</span><span class="mm-num">${this._fmt(bitsPerSec)}</span> <span style="font-size:clamp(13px,1.8vw,17px);color:var(--text-3)">bit/s</span></div>`
      },
      { title:'× 时长 = 总 bit',
        desc:`再乘时长 ${this._fmt(t)} s：${this._fmt(bitsPerSec)} × ${this._fmt(t)} = <b class="c-op">${this._fmt(bitsTotal)}</b> bit（文件总数据量）。`,
        body:()=>`
          <div><span class="mm-num g">${this._fmt(bitsPerSec)}</span><span class="mm-op">×</span><span class="mm-num hot">${this._fmt(t)} s</span><span class="mm-op">=</span><span class="mm-num">${this._fmt(bitsTotal)}</span> <span style="font-size:clamp(13px,1.8vw,17px);color:var(--text-3)">bit</span></div>`
      },
      { title:'bit ÷ 8 = 字节',
        desc:`1 字节 = 8 bit，${this._fmt(bitsTotal)} ÷ 8 = <b class="c-op">${this._fmt(bytes)}</b> 字节（B）。`,
        body:()=>`
          <div><span class="mm-num g">${this._fmt(bitsTotal)} bit</span><span class="mm-op">÷ 8 =</span><span class="mm-num hot">${this._fmt(bytes)} B</span></div>`
      },
      { title:'容量单位换算与结果',
        desc:`总容量 = <b class="c-op">${this._fmt(bytes)}</b> 字节 ≈ <b class="c-op">${hh[0].toFixed(2)} ${hh[1]}</b>（未压缩画面，1 KB = 1024 B）。`,
        body:()=>`
          <div class="mm-conv">${this._convLines(bytes)}</div>
          <div class="mm-result">${this._fmt(t)} s × ${fps} fps × ${w}×${h} × ${bits} bit ÷ 8<br>= ${this._fmt(bytes)} B ≈ ${hh[0].toFixed(2)} ${hh[1]}</div>`
      }
    ];
    return this._stepsFromBlocks(label,'视频计算',blocks);
  },

  /* ========== ③ 图像编码：采样 → 量化 → 容量计算 ========== */
  genImage(){
    const w=this._num('#md-i-w',800,'宽度',1,true);
    const h=this._num('#md-i-h',600,'高度',1,true);
    const bits=this._num('#md-i-bits',24,'位深',1,true);
    const pixels=w*h;
    const totalBits=pixels*bits;
    const bytes=totalBits/8;
    const hh=this.humanSize(bytes);
    const label='图像编码';

    const gridW=Math.min(w,20), gridH=Math.min(h,15);
    let cells='';
    for(let r=0;r<gridH;r++){
      for(let c=0;c<gridW;c++){
        const shade=180+Math.floor((r+c)*3)%60;
        cells+=`<div style="width:100%;aspect-ratio:1;background:rgb(${shade},${shade},${shade+10})"></div>`;
      }
    }
    const colorLevels=Math.pow(2,bits);
    const colorDesc=bits===1?'黑白二值':bits===8?'256级灰度':bits===24?'真彩色(1677万色)':`${colorLevels}色`;
    const bars=[];
    const n=Math.min(8,bits===1?2:bits===8?8:6);
    for(let i=0;i<n;i++){
      const v=Math.floor(255*i/(n-1||1));
      bars.push(`<div style="flex:1;height:24px;background:rgb(${v},${v},${v})" title="灰度 ${v}"></div>`);
    }

    const blocks=[
      { title:'第一步：采样（像素网格）',
        desc:`将连续图像划分为 ${w} × ${h} = <b class="c-op">${this._fmt(pixels)}</b> 个像素（采样点），每个像素用固定位数表示颜色。`,
        body:()=>`
          <div style="display:flex;gap:clamp(12px,2vw,18px);flex-wrap:wrap;align-items:center">
            <div style="flex:1;min-width:120px;max-width:220px;margin:0 auto">
              <div style="display:grid;grid-template-columns:repeat(${gridW},1fr);gap:1px;border:1px solid var(--border);border-radius:6px;overflow:hidden">${cells}</div>
              <div class="mm-note" style="text-align:center">${gridW}×${gridH} 示意（实际 ${w}×${h}）</div>
            </div>
            <div style="flex:1;min-width:150px">
              <div><span class="mm-num g">${w}</span><span class="mm-op">×</span><span class="mm-num hot">${h}</span><span class="mm-op">=</span><span class="mm-num">${this._fmt(pixels)}</span> 像素</div>
            </div>
          </div>`
      },
      { title:'第二步：量化（颜色级别）',
        desc:`每个像素用 <b class="c-op">${bits} bit</b> 表示颜色，可表示 ${colorLevels.toLocaleString()} 种颜色（${colorDesc}）；位深越大，颜色越丰富。`,
        body:()=>`
          <div style="display:flex;gap:2px;margin:10px 0;border-radius:6px;overflow:hidden">${bars.join('')}</div>
          <div><span class="mm-num g">2</span><span class="mm-op">^</span><span class="mm-num hot">${bits}</span><span class="mm-op">=</span><span class="mm-num">${colorLevels.toLocaleString()}</span> 色（${colorDesc}）</div>`
      },
      { title:'第三步：编码（容量计算）',
        desc:`总数据量 = 像素数 × 位深 = ${this._fmt(pixels)} × ${bits} = <b class="c-op">${this._fmt(totalBits)}</b> bit，除以 8 得 <b class="c-op">${this._fmt(bytes)}</b> 字节。`,
        body:()=>`
          <div class="mm-conv">
            <div>${this._fmt(pixels)} 像素 <span class="mm-op">×</span> ${bits} bit <span class="mm-op">=</span> <span class="mm-num">${this._fmt(totalBits)} bit</span></div>
            <div>${this._fmt(totalBits)} bit <span class="mm-op">÷ 8 =</span> <span class="mm-num hot">${this._fmt(bytes)} B</span></div>
          </div>`
      },
      { title:'图像编码结果',
        desc:`${w}×${h} 像素、${bits} bit 位深的未压缩图像，容量 = <b class="c-op">${this._fmt(bytes)}</b> 字节 ≈ <b class="c-op">${hh[0].toFixed(2)} ${hh[1]}</b>。`,
        body:()=>`
          <div class="mm-result">${w}×${h} 像素 × ${bits} bit ÷ 8 = ${this._fmt(bytes)} B ≈ ${hh[0].toFixed(2)} ${hh[1]}（未压缩 BMP）</div>
          <div class="mm-note" style="text-align:center">采样 → 量化 → 编码，三步完成图像数字化</div>`
      }
    ];
    return this._stepsFromBlocks(label,'图像编码',blocks);
  },

  /* ---------- 初始化画布（切换子页面时清空旧内容） ---------- */
  initStage(){
    const sub=this.currentSub;
    const info=this.subTabs.find(t=>t.id===sub)||{};
    const labelMap={
      audio:'音频文件容量（未压缩）',
      video:'视频文件容量（未压缩）',
      image:'图像编码（采样 · 量化 · 编码）'
    };
    $('#stage').innerHTML=`<div class="stage-hd"><span>${labelMap[sub]||'音视频容量'}</span><span class="state">准备就绪</span></div>
      <div style="text-align:center;padding:clamp(16px,3vw,28px) 0;color:#94A3B8">
        <div style="font-size:clamp(13px,1.8vw,16px);margin-bottom:8px">使用 <b style="color:${this.color}">${info.method||''}</b></div>
        <div style="font-size:clamp(11px,1.4vw,12.5px)">点击「自动演示」或「下一步」开始手写推导</div>
      </div>`;
  },
  generate(){
    switch(this.currentSub){
      case 'audio': this.steps=this.genAudio(); break;
      case 'video': this.steps=this.genVideo(); break;
      case 'image': this.steps=this.genImage(); break;
      default: this.steps=this.genAudio();
    }
  }
};
