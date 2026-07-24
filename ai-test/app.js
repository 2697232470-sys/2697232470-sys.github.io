const $ = (s) => document.querySelector(s);
const state = { files: [], records: [], ocr: null, running: false };
const FIELD_LABELS = {
  invoiceNumber:'发票号码', invoiceDate:'开票日期', buyerName:'购买方名称', buyerTaxId:'购买方税号',
  sellerName:'销售方名称', sellerTaxId:'销售方税号', totalAmount:'不含税金额', totalTax:'税额',
  totalAmountWithTax:'价税合计', drawer:'开票人'
};
const CRITICAL_FIELDS = ['invoiceNumber','invoiceDate','buyerName','sellerName','totalAmount','totalTax','totalAmountWithTax'];

window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

function safe(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function norm(v){return String(v??'').replace(/\u3000/g,' ').replace(/[ \t]+/g,' ').replace(/\s*([：:,，])\s*/g,'$1').trim();}
function compact(v){return norm(v).replace(/\s+/g,'').toUpperCase();}
function num(v){if(v===null||v===undefined||v==='')return null;const n=Number(String(v).replace(/[¥￥,\s]/g,''));return Number.isFinite(n)?n:null;}
function money(v){const n=num(v);return n===null?'':n.toFixed(2);}
function unique(a){return [...new Set(a.filter(Boolean))];}
function sizeText(b){return b<1024?b+' B':b<1048576?(b/1024).toFixed(1)+' KB':(b/1048576).toFixed(2)+' MB';}
function showToast(text){const t=$('#toast');t.textContent=text;t.classList.remove('hidden');clearTimeout(showToast.timer);showToast.timer=setTimeout(()=>t.classList.add('hidden'),4200);}
function setStep(id,status){const e=$(id);e.classList.remove('active','done');if(status)e.classList.add(status);}
function progress(p,title,detail){$('#progressCard').classList.remove('hidden');$('#progressBar').style.width=Math.max(0,Math.min(100,p))+'%';$('#progressPercent').textContent=Math.round(p)+'%';if(title)$('#progressTitle').textContent=title;if(detail!==undefined)$('#progressDetail').textContent=detail;}

function cleanCompany(v){
  let s=norm(v).replace(/[|｜]/g,' ').replace(/\s+/g,' ').trim();
  for(let i=0;i<4;i++)s=s.replace(/^(?:购\s*买\s*方|销\s*售\s*方|购买方信息|销售方信息|购买方|销售方|购|销|信息)\s*/,'').replace(/^(?:名\s*称|名称)\s*[:：]?\s*/,'');
  s=s.replace(/\s*(?:购买方信息|销售方信息|统一社会信用代码.*|纳税人识别号.*)$/,'').replace(/^[：:]+|[：:]+$/g,'').trim();
  const m=s.match(/[\u4e00-\u9fa5A-Za-z0-9（）()·&\-]{2,90}?(?:有限责任公司|股份有限公司|有限公司|公司|经营部|事务所|合作社|商行|中心|厂|个体工商户)/);
  return (m?m[0]:s).trim();
}
function companyKey(v){return cleanCompany(v).replace(/[（）()\s]/g,'').replace(/有限责任公司$/,'有限公司').toUpperCase();}
function parseDate(v){const m=String(v).match(/(20\d{2})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})\s*日?/);return m?`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`:'';}
function bboxFromPoly(poly){
  const pts=(poly||[]).flatMap(p=>Array.isArray(p)&&p.length>=2?[p]:[]);if(!pts.length)return{x:0,y:0,w:0,h:0,cx:0,cy:0};
  const xs=pts.map(p=>Number(p[0])),ys=pts.map(p=>Number(p[1]));const x=Math.min(...xs),y=Math.min(...ys),x2=Math.max(...xs),y2=Math.max(...ys);return{x,y,w:x2-x,h:y2-y,cx:(x+x2)/2,cy:(y+y2)/2};
}
function rowsFromItems(items,tolerance=12){
  const sorted=[...items].sort((a,b)=>a.cy-b.cy||a.x-b.x);const rows=[];
  for(const item of sorted){let row=rows.find(r=>Math.abs(r.cy-item.cy)<=Math.max(tolerance,(item.h||12)*.55));if(!row){row={cy:item.cy,items:[]};rows.push(row);}row.items.push(item);row.cy=row.items.reduce((s,x)=>s+x.cy,0)/row.items.length;}
  rows.forEach(r=>r.items.sort((a,b)=>a.x-b.x));return rows.sort((a,b)=>a.cy-b.cy);
}
function rowText(row){return row.items.map(i=>i.text).join(' ').trim();}
function extractTextLayer(page,viewport){return page.getTextContent().then(tc=>tc.items.filter(x=>x.str&&x.str.trim()).map(x=>{const x=x.transform[4]*viewport.scale,y=viewport.height-x.transform[5]*viewport.scale,w=(x.width||0)*viewport.scale,h=Math.abs(x.transform[3]||10)*viewport.scale;return{text:x.str,score:1,x,y,w,h,cx:x+w/2,cy:y+h/2,source:'pdf'};}));}
function ocrItems(result){return (result?.items||[]).filter(x=>x.text&&String(x.text).trim()).map(x=>{const b=bboxFromPoly(x.poly);return{text:String(x.text),score:Number(x.score??0),...b,source:'ai'};});}

function parseSource(items,width,height,source){
  const rows=rowsFromItems(items,source==='ai'?14:8);const texts=rows.map(rowText);const full=texts.join('\n');
  const topItems=items.filter(i=>i.cy<height*.38);
  const pure20=topItems.map(i=>compact(i.text)).filter(x=>/^\d{20}$/.test(x));
  let invoiceNumber=pure20[0]||'';
  if(!invoiceNumber){const m=full.match(/发票号码[:：]?\s*(\d{8,24})/);invoiceNumber=m?.[1]||'';}
  if(!invoiceNumber){invoiceNumber=(unique(full.match(/(?<!\d)\d{20}(?!\d)/g)||[]))[0]||'';}
  const invoiceDate=parseDate(full);
  const companyItems=items.filter(i=>i.cy<height*.58&&/(有限责任公司|股份有限公司|有限公司|公司|经营部|商行|中心|厂|事务所|合作社|个体工商户)/.test(i.text)&&!/(发票|购买方|销售方|开票人)/.test(i.text));
  const buyerCandidates=companyItems.filter(i=>i.cx<width*.5).map(i=>cleanCompany(i.text)).filter(Boolean).sort((a,b)=>b.length-a.length);
  const sellerCandidates=companyItems.filter(i=>i.cx>=width*.5).map(i=>cleanCompany(i.text)).filter(Boolean).sort((a,b)=>b.length-a.length);
  const allCompanies=unique(companyItems.map(i=>cleanCompany(i.text))).sort((a,b)=>b.length-a.length);
  let buyerName=buyerCandidates[0]||allCompanies[0]||'';
  let sellerName=sellerCandidates[0]||allCompanies.find(x=>companyKey(x)!==companyKey(buyerName))||'';
  const taxItems=items.map(i=>({...i,v:compact(i.text)})).filter(i=>/^[0-9A-Z]{15,20}$/.test(i.v)&&/[A-Z]/.test(i.v)&&!/^\d{20}$/.test(i.v));
  const buyerTaxId=(taxItems.filter(i=>i.cx<width*.5).sort((a,b)=>a.cy-b.cy)[0]?.v)||'';
  const sellerTaxId=(taxItems.filter(i=>i.cx>=width*.5).sort((a,b)=>a.cy-b.cy)[0]?.v)||taxItems.find(i=>i.v!==buyerTaxId)?.v||'';

  const itemsDetail=parseItemRows(rows,width,height);
  const totals=parseTotals(rows,items,width,height,itemsDetail);
  const drawerMatch=full.match(/开\s*票\s*人[:：]?\s*([^\s,，;；]{2,16})/);
  const drawer=drawerMatch?.[1]?.replace(/^[：:]/,'')||'';
  const scores=items.map(i=>i.score).filter(Number.isFinite);const avgScore=scores.length?scores.reduce((a,b)=>a+b,0)/scores.length:1;
  return {source,rawText:full,invoiceNumber,invoiceDate,buyerName,sellerName,buyerTaxId,sellerTaxId,items:itemsDetail,...totals,drawer,avgScore};
}

function parseItemRows(rows,width,height){
  const result=[];
  for(const row of rows){
    const text=rowText(row);if(row.cy<height*.25||row.cy>height*.82)continue;if(!/(\d+(?:\.\d+)?%|免税|不征税)/.test(text))continue;if(/税率|征收率|合计|价税/.test(text))continue;
    const zone=(a,b)=>row.items.filter(i=>i.cx>=width*a&&i.cx<width*b).map(i=>i.text).join(' ').trim();
    const firstNumber=(a,b)=>{const m=zone(a,b).replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);return m?Number(m[0]):null;};
    let project=zone(0,.25),spec=zone(.18,.36),unit=zone(.32,.43),quantity=firstNumber(.40,.52),unitPrice=firstNumber(.48,.65),amount=firstNumber(.62,.78),taxAmount=firstNumber(.86,1.01);
    const taxRate=(zone(.72,.91).match(/\d+(?:\.\d+)?%|免税|不征税/)||[])[0]||'';
    project=project.replace(/项目名称|货物或应税劳务、服务名称/g,'').trim();
    if(project||amount!==null||taxAmount!==null)result.push({projectName:project,specification:spec,unit,quantity,unitPrice,amount,taxRate,taxAmount,raw:text});
  }
  return result;
}
function parseTotals(rows,items,width,height,details){
  let totalAmount=null,totalTax=null,totalAmountWithTax=null,totalAmountInWords='';
  const currency=[];
  items.forEach(i=>{const s=compact(i.text);if(/^[¥￥]-?[\d,]+(?:\.\d{1,2})?$/.test(s))currency.push({n:num(s),x:i.cx,y:i.cy});});
  const priceLabel=rows.find(r=>/价税合计/.test(rowText(r)));
  if(priceLabel){
    const near=items.filter(i=>Math.abs(i.cy-priceLabel.cy)<50&&i.cy>=priceLabel.cy-15).map(i=>num(i.text)).filter(n=>n!==null&&n>=0&&n<1e9);
    const curr=currency.filter(x=>Math.abs(x.y-priceLabel.cy)<60).map(x=>x.n);const vals=[...curr,...near];if(vals.length)totalAmountWithTax=Math.max(...vals);
    const scope=rows.filter(r=>Math.abs(r.cy-priceLabel.cy)<70).map(rowText).join(' ');const w=scope.match(/[零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整]+/);if(w)totalAmountInWords=w[0];
  }
  const sumLabel=rows.find(r=>/合\s*计/.test(rowText(r))&&!/价税/.test(rowText(r)));
  if(sumLabel){const near=currency.filter(x=>Math.abs(x.y-sumLabel.cy)<45).sort((a,b)=>a.x-b.x).map(x=>x.n);if(near.length>=2){totalAmount=near[near.length-2];totalTax=near[near.length-1];}}
  if(details.length){const am=details.map(x=>x.amount).filter(x=>x!==null),tx=details.map(x=>x.taxAmount).filter(x=>x!==null);if(totalAmount===null&&am.length)totalAmount=Number(am.reduce((a,b)=>a+b,0).toFixed(2));if(totalTax===null&&tx.length)totalTax=Number(tx.reduce((a,b)=>a+b,0).toFixed(2));}
  if(totalAmountWithTax===null&&currency.length)totalAmountWithTax=Math.max(...currency.map(x=>x.n));
  if(totalAmountWithTax===null&&totalAmount!==null&&totalTax!==null)totalAmountWithTax=Number((totalAmount+totalTax).toFixed(2));
  if(totalAmount===null&&totalAmountWithTax!==null&&totalTax!==null)totalAmount=Number((totalAmountWithTax-totalTax).toFixed(2));
  if(totalTax===null&&totalAmountWithTax!==null&&totalAmount!==null)totalTax=Number((totalAmountWithTax-totalAmount).toFixed(2));
  return {totalAmount,totalTax,totalAmountWithTax,totalAmountInWords};
}

function equivalent(field,a,b){
  if(a===null||a===undefined||a===''||b===null||b===undefined||b==='')return false;
  if(['totalAmount','totalTax','totalAmountWithTax'].includes(field))return Math.abs(num(a)-num(b))<=.02;
  if(field==='invoiceDate')return parseDate(a)===parseDate(b);
  if(field==='buyerName'||field==='sellerName'){const x=companyKey(a),y=companyKey(b);return x===y||x.includes(y)||y.includes(x);}
  return compact(a)===compact(b);
}
function chooseField(field,pdf,ai){
  const p=pdf?.[field],a=ai?.[field];
  if(p!==null&&p!==undefined&&p!==''&&a!==null&&a!==undefined&&a!==''){
    if(equivalent(field,p,a)){
      if(field==='buyerName'||field==='sellerName')return{value:cleanCompany(p).length<=cleanCompany(a).length?cleanCompany(p):cleanCompany(a),state:'agree'};
      return{value:p,state:'agree'};
    }
    return{value:p,state:'conflict',pdf:p,ai:a};
  }
  if(p!==null&&p!==undefined&&p!=='')return{value:p,state:'pdf_only'};
  if(a!==null&&a!==undefined&&a!=='')return{value:a,state:'ai_only'};
  return{value:'',state:'missing'};
}
function mergeSources(fileName,page,pdf,ai){
  const evidence={},final={};let agree=0;const conflicts=[];const fields=Object.keys(FIELD_LABELS);
  for(const field of fields){const r=chooseField(field,pdf,ai);evidence[field]=r;final[field]=r.value;if(r.state==='agree')agree++;if(r.state==='conflict')conflicts.push(`${FIELD_LABELS[field]}两路结果不一致`);}
  const pdfItems=pdf?.items||[],aiItems=ai?.items||[];const finalItems=aiItems.length>=pdfItems.length?aiItems:pdfItems;
  const amountEquation=final.totalAmount!==''&&final.totalTax!==''&&final.totalAmountWithTax!==''?Math.abs(num(final.totalAmount)+num(final.totalTax)-num(final.totalAmountWithTax))<=.02:false;
  if(!amountEquation)conflicts.push('金额关系未通过');
  if(finalItems.length){const itemAmount=finalItems.map(x=>num(x.amount)).filter(x=>x!==null).reduce((a,b)=>a+b,0);const itemTax=finalItems.map(x=>num(x.taxAmount)).filter(x=>x!==null).reduce((a,b)=>a+b,0);if(itemAmount&&final.totalAmount!==''&&Math.abs(itemAmount-num(final.totalAmount))>.02)conflicts.push('商品明细金额合计与票面不一致');if(itemTax&&final.totalTax!==''&&Math.abs(itemTax-num(final.totalTax))>.02)conflicts.push('商品明细税额合计与票面不一致');}
  const missing=CRITICAL_FIELDS.filter(k=>final[k]===null||final[k]===undefined||final[k]==='');if(missing.length)conflicts.push('缺少：'+missing.map(k=>FIELD_LABELS[k]).join('、'));
  let status='verified';if(!pdf&&!ai)status='failed';else if(conflicts.length||agree<4)status='conflict';
  return {id:`${Date.now()}-${Math.random()}`,fileName,page,status,final,evidence,conflicts:[...new Set(conflicts)],agreementCount:agree,pdf,ai,items:finalItems,processedAt:new Date().toLocaleString('zh-CN',{hour12:false})};
}

async function loadPaddle(){
  if(state.ocr)return state.ocr;
  setStep('#stepModel','active');progress(3,'正在加载免费 AI 模型','首次使用需要下载 PP-OCRv5 模型和浏览器推理组件');
  const urls=['https://cdn.jsdelivr.net/npm/@paddleocr/paddleocr-js/+esm','https://esm.sh/@paddleocr/paddleocr-js?bundle'];let mod,last;
  for(const url of urls){try{mod=await import(url);if(mod?.PaddleOCR)break;}catch(e){last=e;}}
  if(!mod?.PaddleOCR)throw new Error('PaddleOCR.js 加载失败：'+(last?.message||'未知错误'));
  state.ocr=await mod.PaddleOCR.create({lang:'ch',ocrVersion:'PP-OCRv5',worker:false,ortOptions:{backend:'wasm',wasmPaths:'https://cdn.jsdelivr.net/npm/onnxruntime-web/dist/',numThreads:1,simd:true}});
  setStep('#stepModel','done');return state.ocr;
}
async function canvasFromImage(file){const url=URL.createObjectURL(file);try{const img=await new Promise((res,rej)=>{const i=new Image();i.onload=()=>res(i);i.onerror=()=>rej(new Error('图片无法打开'));i.src=url;});const max=2600,scale=Math.min(1,max/Math.max(img.naturalWidth,img.naturalHeight));const c=document.createElement('canvas');c.width=Math.round(img.naturalWidth*scale);c.height=Math.round(img.naturalHeight*scale);const x=c.getContext('2d',{willReadFrequently:true});x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(img,0,0,c.width,c.height);return c;}finally{URL.revokeObjectURL(url);}}
async function runOcr(canvas,label){const ocr=await loadPaddle();setStep('#stepOcr','active');progress(35,'正在进行 PP-OCRv5 AI 重识别',label);const [r]=await ocr.predict(canvas,{textDetLimitSideLen:1800,textDetLimitType:'max',textDetMaxSideLimit:2800,textDetThresh:.28,textDetBoxThresh:.48,textDetUnclipRatio:1.6,textRecScoreThresh:.42});return r;}
async function processPdf(file,fileIndex,totalFiles){
  const data=await file.arrayBuffer();const doc=await window.pdfjsLib.getDocument({data}).promise;const out=[];
  for(let p=1;p<=doc.numPages;p++){
    const page=await doc.getPage(p);const raw=page.getViewport({scale:1});const targetWidth=Math.min(2600,Math.max(1800,raw.width*3));const scale=targetWidth/raw.width;const vp=page.getViewport({scale});const canvas=document.createElement('canvas');canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);const ctx=canvas.getContext('2d',{willReadFrequently:true});ctx.fillStyle='#fff';ctx.fillRect(0,0,canvas.width,canvas.height);await page.render({canvasContext:ctx,viewport:vp}).promise;
    setStep('#stepPdf','active');progress(12+fileIndex/totalFiles*10,'正在读取 PDF 原始文字层',`${file.name} · 第 ${p}/${doc.numPages} 页`);const pdfItems=await extractTextLayer(page,vp);const pdfParsed=pdfItems.length?parseSource(pdfItems,canvas.width,canvas.height,'pdf'):null;setStep('#stepPdf','done');
    const aiResult=await runOcr(canvas,`${file.name} · 第 ${p}/${doc.numPages} 页`);const aiParsed=parseSource(ocrItems(aiResult),canvas.width,canvas.height,'ai');
    setStep('#stepCompare','active');progress(72,'正在交叉校验','比较 PDF 文字层、AI 识别和金额关系');out.push(mergeSources(file.name,p,pdfParsed,aiParsed));setStep('#stepCompare','done');
  }return out;
}
async function processImage(file){const canvas=await canvasFromImage(file);const aiResult=await runOcr(canvas,file.name);const aiParsed=parseSource(ocrItems(aiResult),canvas.width,canvas.height,'ai');return[mergeSources(file.name,1,null,aiParsed)];}

function renderFiles(){
  $('#fileList').innerHTML=state.files.map((x,i)=>`<div class="file-row"><div><div class="file-name">${safe(x.name)}</div><div class="file-meta">${sizeText(x.size)} · ${safe(x.type||'文件')}</div></div><div class="file-status">${safe(x.status||'等待识别')}</div><button class="remove" data-index="${i}">×</button></div>`).join('');
  $('#runBtn').disabled=!state.files.length||state.running;$('#clearBtn').disabled=!state.files.length||state.running;
}
function addFiles(list){for(const f of list){const ext=f.name.split('.').pop().toLowerCase();if(!['pdf','jpg','jpeg','png'].includes(ext)){showToast('已跳过不支持的文件：'+f.name);continue;}if(!state.files.some(x=>x.name===f.name&&x.size===f.size))state.files.push(Object.assign(f,{status:'等待识别'}));}renderFiles();}
function renderResults(){
  const counts={verified:0,conflict:0,failed:0};state.records.forEach(r=>counts[r.status]++);$('#mTotal').textContent=state.records.length;$('#mVerified').textContent=counts.verified;$('#mConflict').textContent=counts.conflict;$('#mFailed').textContent=counts.failed;
  $('#resultBody').innerHTML=state.records.map(r=>{const f=r.final;const text=r.status==='verified'?'双路一致':r.status==='failed'?'识别失败':'存在冲突';return`<tr><td><span class="badge ${r.status}">${text}</span></td><td title="${safe(r.fileName)}">${safe(r.fileName)}</td><td>${safe(f.invoiceNumber)}</td><td>${safe(f.invoiceDate)}</td><td title="${safe(f.buyerName)}">${safe(f.buyerName)}</td><td title="${safe(f.sellerName)}">${safe(f.sellerName)}</td><td>${money(f.totalAmount)}</td><td>${money(f.totalTax)}</td><td>${money(f.totalAmountWithTax)}</td><td title="${safe(r.conflicts.join('；'))}">${safe(r.conflicts.join('；')||`有 ${r.agreementCount} 个字段双路一致`)}</td><td><button class="evidence-btn" data-id="${r.id}">查看证据</button></td></tr>`;}).join('');$('#resultSection').classList.remove('hidden');
}
function openEvidence(id){const r=state.records.find(x=>x.id===id);if(!r)return;$('#dialogFile').textContent=`${r.fileName} · 第 ${r.page} 页`;let rows='<div class="head">字段</div><div class="head">PDF文字层</div><div class="head">AI识别</div><div class="head">最终采用</div>';for(const [k,label] of Object.entries(FIELD_LABELS)){rows+=`<div>${label}</div><div>${safe(r.pdf?.[k]??'')}</div><div>${safe(r.ai?.[k]??'')}</div><div>${safe(r.final[k]??'')}<br><small>${safe(r.evidence[k]?.state||'')}</small></div>`;}$('#evidenceBody').innerHTML=`<div class="evidence-grid">${rows}</div><div class="raw-box"><div><h4>PDF文字层原文</h4><pre>${safe(r.pdf?.rawText||'无')}</pre></div><div><h4>AI OCR原文</h4><pre>${safe(r.ai?.rawText||'无')}</pre></div></div>`;$('#evidenceDialog').showModal();}

async function run(){
  if(!state.files.length||state.running)return;state.running=true;state.records=[];renderFiles();$('#resultSection').classList.add('hidden');['#stepModel','#stepPdf','#stepOcr','#stepCompare','#stepExport'].forEach(id=>setStep(id,''));
  try{await loadPaddle();for(let i=0;i<state.files.length;i++){const f=state.files[i];f.status='识别中';renderFiles();progress(8+i/state.files.length*75,'正在处理发票',`${i+1}/${state.files.length} · ${f.name}`);try{const ext=f.name.split('.').pop().toLowerCase();const records=ext==='pdf'?await processPdf(f,i,state.files.length):await processImage(f);state.records.push(...records);f.status='完成';}catch(e){console.error(e);f.status='失败';state.records.push({id:`fail-${Date.now()}-${i}`,fileName:f.name,page:1,status:'failed',final:{},evidence:{},conflicts:[e.message||String(e)],agreementCount:0,pdf:null,ai:null,items:[],processedAt:new Date().toLocaleString('zh-CN',{hour12:false})});}renderFiles();}
    setStep('#stepOcr','done');setStep('#stepExport','active');progress(95,'正在整理结果','生成结构化结果和证据记录');renderResults();setStep('#stepExport','done');progress(100,'处理完成',`共生成 ${state.records.length} 条记录`);showToast('AI 测试识别完成，请先查看冲突和证据');
  }catch(e){console.error(e);showToast(e.message||String(e));progress(0,'初始化失败',e.message||String(e));}finally{state.running=false;renderFiles();}
}
function exportExcel(){
  if(!state.records.length)return;const summary=state.records.map(r=>{const f=r.final;return{'状态':r.status==='verified'?'双路一致':r.status==='failed'?'识别失败':'存在冲突','原始文件名':r.fileName,'页码':r.page,'发票号码':f.invoiceNumber||'','开票日期':f.invoiceDate||'','购买方名称':f.buyerName||'','购买方税号':f.buyerTaxId||'','销售方名称':f.sellerName||'','销售方税号':f.sellerTaxId||'','不含税金额':num(f.totalAmount),'税额':num(f.totalTax),'价税合计':num(f.totalAmountWithTax),'开票人':f.drawer||'','双路一致字段数':r.agreementCount,'异常说明':r.conflicts.join('；'),'处理时间':r.processedAt};});
  const details=[];state.records.forEach(r=>(r.items||[]).forEach((x,i)=>details.push({'原始文件名':r.fileName,'发票号码':r.final.invoiceNumber||'','明细序号':i+1,'项目名称':x.projectName||'','规格型号':x.specification||'','单位':x.unit||'','数量':x.quantity,'单价':x.unitPrice,'不含税金额':x.amount,'税率':x.taxRate||'','税额':x.taxAmount,'原始行':x.raw||''})));
  const conflicts=[];state.records.forEach(r=>Object.entries(r.evidence||{}).forEach(([k,v])=>{if(v.state==='conflict'||v.state==='missing')conflicts.push({'原始文件名':r.fileName,'页码':r.page,'字段':FIELD_LABELS[k]||k,'PDF文字层':v.pdf??r.pdf?.[k]??'','AI识别':v.ai??r.ai?.[k]??'','最终采用':r.final[k]??'','判断':v.state});}));
  const raw=state.records.map(r=>({'原始文件名':r.fileName,'页码':r.page,'PDF文字层原文':r.pdf?.rawText||'','AI OCR原文':r.ai?.rawText||''}));
  const wb=XLSX.utils.book_new();[['AI校验汇总',summary],['发票明细',details],['字段冲突',conflicts],['识别原文',raw]].forEach(([name,data])=>{const ws=XLSX.utils.json_to_sheet(data.length?data:[{'暂无数据':''}]);ws['!autofilter']={ref:ws['!ref']};ws['!cols']=Object.keys(data[0]||{'暂无数据':''}).map(k=>({wch:Math.min(45,Math.max(12,k.length*2+2,...data.slice(0,100).map(x=>String(x[k]??'').length+2)))}));XLSX.utils.book_append_sheet(wb,ws,name);});const d=new Date(),p=n=>String(n).padStart(2,'0');XLSX.writeFile(wb,`免费AI发票识别测试_${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}.xlsx`);
}

$('#chooseBtn').onclick=()=>$('#fileInput').click();$('#fileInput').onchange=e=>{addFiles(e.target.files);e.target.value='';};$('#dropZone').onclick=e=>{if(e.target.id!=='chooseBtn')$('#fileInput').click();};['dragenter','dragover'].forEach(t=>$('#dropZone').addEventListener(t,e=>{e.preventDefault();$('#dropZone').classList.add('drag');}));['dragleave','drop'].forEach(t=>$('#dropZone').addEventListener(t,e=>{e.preventDefault();$('#dropZone').classList.remove('drag');}));$('#dropZone').addEventListener('drop',e=>addFiles(e.dataTransfer.files));$('#fileList').onclick=e=>{const b=e.target.closest('[data-index]');if(b){state.files.splice(Number(b.dataset.index),1);renderFiles();}};$('#clearBtn').onclick=()=>{state.files=[];state.records=[];renderFiles();$('#resultSection').classList.add('hidden');};$('#runBtn').onclick=run;$('#resultBody').onclick=e=>{const b=e.target.closest('[data-id]');if(b)openEvidence(b.dataset.id);};$('#closeDialog').onclick=()=>$('#evidenceDialog').close();$('#exportBtn').onclick=exportExcel;renderFiles();