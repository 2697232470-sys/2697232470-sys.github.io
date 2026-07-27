const $ = (s) => document.querySelector(s);
const state = {
  recognitionFile: null,
  taxFile: null,
  recognitionRows: [],
  recognitionDetails: [],
  taxRows: [],
  taxDetails: [],
  results: [],
  differenceRows: [],
  abnormalRows: [],
  pdfOnlyRows: [],
  taxOnlyRows: [],
  detailRows: [],
  currentView: 'summary'
};

const STATUS_TEXT = {
  matched: '完全一致',
  corrected: '已自动修正',
  difference: '字段差异',
  tax_abnormal: '税务状态异常',
  pdf_only: '税务局未找到',
  suspected: '疑似匹配'
};

function safe(v){return String(v ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function text(v){return v===null||v===undefined?'':String(v).trim();}
function idText(v){return text(v).replace(/\.0$/,'').replace(/\s+/g,'');}
function num(v){if(v===null||v===undefined||v==='')return null;const n=Number(String(v).replace(/[¥￥,\s]/g,''));return Number.isFinite(n)?n:null;}
function money(v){const n=num(v);return n===null?'':n.toFixed(2);}
function dateOnly(v){const s=text(v);const m=s.match(/(20\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})/);return m?`${m[1]}-${String(Number(m[2])).padStart(2,'0')}-${String(Number(m[3])).padStart(2,'0')}`:s.slice(0,10);}
function normalizeName(v){return text(v).replace(/购买方信息|销售方信息|名称[:：]?/g,'').replace(/[（）()\s,，。:：;；·_\-\/\\]/g,'').replace(/个人$/,'').toUpperCase();}
function amountsEqual(a,b){const x=num(a),y=num(b);return x!==null&&y!==null&&Math.abs(x-y)<=0.02;}
function stringsEqual(a,b){const x=text(a).toUpperCase(),y=text(b).toUpperCase();return x!==''&&x===y;}
function namesEqual(a,b){const x=normalizeName(a),y=normalizeName(b);return x!==''&&y!==''&&(x===y||x.includes(y)||y.includes(x));}
function showToast(msg){const el=$('#toast');el.textContent=msg;el.classList.remove('hidden');clearTimeout(showToast.timer);showToast.timer=setTimeout(()=>el.classList.add('hidden'),4500);}
function setProgress(p,title,detail){$('#progressCard').classList.remove('hidden');$('#progressBar').style.width=Math.max(0,Math.min(100,p))+'%';$('#progressPercent').textContent=Math.round(p)+'%';if(title)$('#progressTitle').textContent=title;if(detail!==undefined)$('#progressDetail').textContent=detail;}
function updateRunState(){$('#runBtn').disabled=!(state.recognitionFile&&state.taxFile);}
function fileAmount(name){const m=text(name).match(/-(\d+(?:\.\d+)?)(?=（|\.pdf$|$)/i);return m?Number(m[1]):null;}
function firstValue(row,names){for(const n of names){if(Object.prototype.hasOwnProperty.call(row,n)&&row[n]!==null&&row[n]!==undefined&&text(row[n])!=='')return row[n];}return null;}
function sheetByNames(wb,names){for(const n of names){if(wb.Sheets[n])return {name:n,sheet:wb.Sheets[n]};}const lower=names.map(n=>n.toLowerCase());for(const n of wb.SheetNames){if(lower.some(x=>n.toLowerCase().includes(x)))return {name:n,sheet:wb.Sheets[n]};}return null;}
function sheetRows(sheet){return XLSX.utils.sheet_to_json(sheet,{defval:null,raw:true,blankrows:false});}
function readWorkbook(file){return new Promise((resolve,reject)=>{const fr=new FileReader();fr.onload=()=>{try{resolve(XLSX.read(fr.result,{type:'array',cellDates:false}));}catch(e){reject(e);}};fr.onerror=()=>reject(fr.error||new Error('文件读取失败'));fr.readAsArrayBuffer(file);});}

function parseRecognitionWorkbook(wb){
  const summarySheet=sheetByNames(wb,['发票汇总','最终确认结果','AI校验汇总','识别汇总','对账汇总'])||{name:wb.SheetNames[0],sheet:wb.Sheets[wb.SheetNames[0]]};
  const rows=sheetRows(summarySheet.sheet).map((r,i)=>({
    rowNo:i+2,
    sourceFileName:text(firstValue(r,['原始文件名','文件名','PDF文件名','来源文件'])),
    invoiceCode:idText(firstValue(r,['发票代码'])),
    invoiceNumber:idText(firstValue(r,['数电发票号码','发票号码','票据号码','发票No.'])),
    invoiceDate:dateOnly(firstValue(r,['开票日期','发票日期'])),
    buyerName:text(firstValue(r,['购买方名称','购方名称','购买方'])),
    buyerTaxId:idText(firstValue(r,['购买方税号','购方识别号','购买方统一社会信用代码','购买方纳税人识别号'])),
    sellerName:text(firstValue(r,['销售方名称','销方名称','销售方'])),
    sellerTaxId:idText(firstValue(r,['销售方税号','销方识别号','销售方统一社会信用代码','销售方纳税人识别号'])),
    amount:num(firstValue(r,['不含税金额','金额','合计金额'])),
    tax:num(firstValue(r,['税额','合计税额'])),
    total:num(firstValue(r,['价税合计','含税金额','合计价税'])),
    invoiceType:text(firstValue(r,['发票类型','发票票种','票种'])),
    recognitionStatus:text(firstValue(r,['识别状态','状态','对账状态'])),
    drawer:text(firstValue(r,['开票人'])),
    raw:r
  })).filter(r=>r.invoiceNumber||r.sourceFileName||r.sellerTaxId);

  const detailSheet=sheetByNames(wb,['发票明细','PDF商品明细','商品明细','明细']);
  const details=detailSheet?sheetRows(detailSheet.sheet).map((r,i)=>({
    rowNo:i+2,
    sourceFileName:text(firstValue(r,['原始文件名','文件名'])),
    invoiceNumber:idText(firstValue(r,['数电发票号码','发票号码'])),
    lineNo:Number(firstValue(r,['明细序号','序号']))||i+1,
    projectName:text(firstValue(r,['项目名称','货物或应税劳务名称','商品名称'])),
    specification:text(firstValue(r,['规格型号'])),
    unit:text(firstValue(r,['单位'])),
    quantity:num(firstValue(r,['数量'])),
    unitPrice:num(firstValue(r,['单价'])),
    amount:num(firstValue(r,['不含税金额','金额'])),
    taxRate:text(firstValue(r,['税率','税率/征收率'])),
    tax:num(firstValue(r,['税额'])),
    raw:r
  })).filter(r=>r.invoiceNumber||r.sourceFileName):[];
  return {rows,details,summarySheetName:summarySheet.name,detailSheetName:detailSheet?.name||''};
}

function parseTaxWorkbook(wb){
  const baseSheet=sheetByNames(wb,['发票基础信息']);
  const detailSheet=sheetByNames(wb,['信息汇总表']);
  if(!baseSheet)throw new Error('税务局文件中没有找到“发票基础信息”工作表');
  const rows=sheetRows(baseSheet.sheet).map((r,i)=>({
    rowNo:i+2,
    invoiceCode:idText(firstValue(r,['发票代码'])),
    invoiceNumber:idText(firstValue(r,['数电发票号码','发票号码'])),
    sellerTaxId:idText(firstValue(r,['销方识别号','销售方税号'])),
    sellerName:text(firstValue(r,['销方名称','销售方名称'])),
    buyerTaxId:idText(firstValue(r,['购方识别号','购买方税号'])),
    buyerName:text(firstValue(r,['购买方名称','购方名称'])),
    invoiceDate:dateOnly(firstValue(r,['开票日期'])),
    amount:num(firstValue(r,['金额','不含税金额'])),
    tax:num(firstValue(r,['税额'])),
    total:num(firstValue(r,['价税合计'])),
    source:text(firstValue(r,['发票来源'])),
    invoiceType:text(firstValue(r,['发票票种'])),
    taxStatus:text(firstValue(r,['发票状态'])),
    isPositive:text(firstValue(r,['是否正数发票'])),
    riskLevel:text(firstValue(r,['发票风险等级'])),
    drawer:text(firstValue(r,['开票人'])),
    remark:text(firstValue(r,['备注'])),
    raw:r
  })).filter(r=>r.invoiceNumber);

  const details=detailSheet?sheetRows(detailSheet.sheet).map((r,i)=>({
    rowNo:i+2,
    invoiceNumber:idText(firstValue(r,['数电发票号码','发票号码'])),
    classificationCode:idText(firstValue(r,['税收分类编码'])),
    businessType:text(firstValue(r,['特定业务类型'])),
    projectName:text(firstValue(r,['货物或应税劳务名称','项目名称'])),
    specification:text(firstValue(r,['规格型号'])),
    unit:text(firstValue(r,['单位'])),
    quantity:num(firstValue(r,['数量'])),
    unitPrice:num(firstValue(r,['单价'])),
    amount:num(firstValue(r,['金额'])),
    taxRate:text(firstValue(r,['税率'])),
    tax:num(firstValue(r,['税额'])),
    lineTotal:num(firstValue(r,['价税合计'])),
    raw:r
  })).filter(r=>r.invoiceNumber):[];
  return {rows,details,baseSheetName:baseSheet.name,detailSheetName:detailSheet?.name||''};
}

function groupBy(rows,key){const m=new Map();for(const r of rows){const k=r[key]||'';if(!m.has(k))m.set(k,[]);m.get(k).push(r);}return m;}
function fallbackCandidates(rec,taxRows){
  let list=taxRows;
  if(rec.buyerTaxId)list=list.filter(t=>t.buyerTaxId===rec.buyerTaxId);
  if(rec.sellerTaxId)list=list.filter(t=>t.sellerTaxId===rec.sellerTaxId);
  const expectedTotal=rec.total ?? fileAmount(rec.sourceFileName);
  if(expectedTotal!==null)list=list.filter(t=>amountsEqual(t.total,expectedTotal));
  if(rec.invoiceDate)list=list.filter(t=>t.invoiceDate===rec.invoiceDate);
  return list;
}

function compareField(field,label,recVal,taxVal,kind,context){
  const recMissing=recVal===null||recVal===undefined||text(recVal)==='';
  const taxMissing=taxVal===null||taxVal===undefined||text(taxVal)==='';
  if(recMissing&&taxMissing)return null;
  if(recMissing&&!taxMissing)return {field,label,pdfValue:'',taxValue:taxVal,action:'按税务局补齐',severity:'corrected'};
  if(!recMissing&&taxMissing)return {field,label,pdfValue:recVal,taxValue:'',action:'税务局字段为空，保留PDF值',severity:'difference'};
  let equal=false;
  if(kind==='amount')equal=amountsEqual(recVal,taxVal);
  else if(kind==='name')equal=namesEqual(recVal,taxVal);
  else if(kind==='date')equal=dateOnly(recVal)===dateOnly(taxVal);
  else equal=stringsEqual(recVal,taxVal);
  if(equal)return null;
  if(kind==='name'&&context.taxIdEqual)return {field,label,pdfValue:recVal,taxValue:taxVal,action:'税号一致，按税务局名称修正',severity:'corrected'};
  return {field,label,pdfValue:recVal,taxValue:taxVal,action:'需要确认',severity:'difference'};
}

function reconcile(recognition,tax,includeTaxOnly,useTaxTruth){
  const taxIndex=new Map(tax.rows.map(r=>[r.invoiceNumber,r]));
  const recDetailMap=groupBy(recognition.details,'invoiceNumber');
  const taxDetailMap=groupBy(tax.details,'invoiceNumber');
  const results=[],differenceRows=[],abnormalRows=[],pdfOnlyRows=[],detailRows=[];
  const matchedTaxNumbers=new Set();

  for(const rec of recognition.rows){
    let taxRow=rec.invoiceNumber?taxIndex.get(rec.invoiceNumber):null;
    let matchMethod=taxRow?'发票号码精确匹配':'';
    if(!taxRow){const candidates=fallbackCandidates(rec,tax.rows);if(candidates.length===1){taxRow=candidates[0];matchMethod='税号/日期/金额组合匹配';}}
    if(!taxRow){const row={status:'pdf_only',matchMethod:'未匹配',sourceFileName:rec.sourceFileName,invoiceNumber:rec.invoiceNumber,invoiceDate:rec.invoiceDate,buyerName:rec.buyerName,sellerName:rec.sellerName,amount:rec.amount,tax:rec.tax,total:rec.total,note:'税务局报表中没有找到对应记录'};results.push(row);pdfOnlyRows.push(row);continue;}
    matchedTaxNumbers.add(taxRow.invoiceNumber);
    const buyerTaxEqual=rec.buyerTaxId&&taxRow.buyerTaxId&&rec.buyerTaxId===taxRow.buyerTaxId;
    const sellerTaxEqual=rec.sellerTaxId&&taxRow.sellerTaxId&&rec.sellerTaxId===taxRow.sellerTaxId;
    const checks=[
      compareField('invoiceNumber','发票号码',rec.invoiceNumber,taxRow.invoiceNumber,'text',{}),
      compareField('invoiceDate','开票日期',rec.invoiceDate,taxRow.invoiceDate,'date',{}),
      compareField('buyerName','购买方名称',rec.buyerName,taxRow.buyerName,'name',{taxIdEqual:buyerTaxEqual}),
      compareField('buyerTaxId','购买方税号',rec.buyerTaxId,taxRow.buyerTaxId,'text',{}),
      compareField('sellerName','销售方名称',rec.sellerName,taxRow.sellerName,'name',{taxIdEqual:sellerTaxEqual}),
      compareField('sellerTaxId','销售方税号',rec.sellerTaxId,taxRow.sellerTaxId,'text',{}),
      compareField('amount','不含税金额',rec.amount,taxRow.amount,'amount',{}),
      compareField('tax','税额',rec.tax,taxRow.tax,'amount',{}),
      compareField('total','价税合计',rec.total,taxRow.total,'amount',{})
    ].filter(Boolean);
    checks.forEach(d=>differenceRows.push({sourceFileName:rec.sourceFileName,invoiceNumber:taxRow.invoiceNumber,field:d.label,pdfValue:d.pdfValue,taxValue:d.taxValue,action:d.action,severity:d.severity,matchMethod}));

    const taxAbnormal=taxRow.taxStatus!=='正常'||taxRow.isPositive!=='是'||taxRow.riskLevel!=='正常';
    const hardDiff=checks.some(d=>d.severity==='difference');
    const corrected=checks.some(d=>d.severity==='corrected')||matchMethod!=='发票号码精确匹配';
    let status=taxAbnormal?'tax_abnormal':hardDiff?'difference':corrected?'corrected':'matched';
    const finalValue=(r,t)=>useTaxTruth?(t??r):(r??t);
    const row={
      status,matchMethod,sourceFileName:rec.sourceFileName,invoiceNumber:taxRow.invoiceNumber,invoiceDate:finalValue(rec.invoiceDate,taxRow.invoiceDate),
      buyerName:finalValue(rec.buyerName,taxRow.buyerName),buyerTaxId:finalValue(rec.buyerTaxId,taxRow.buyerTaxId),
      sellerName:finalValue(rec.sellerName,taxRow.sellerName),sellerTaxId:finalValue(rec.sellerTaxId,taxRow.sellerTaxId),
      amount:finalValue(rec.amount,taxRow.amount),tax:finalValue(rec.tax,taxRow.tax),total:finalValue(rec.total,taxRow.total),
      invoiceType:taxRow.invoiceType||rec.invoiceType,taxStatus:taxRow.taxStatus,isPositive:taxRow.isPositive,riskLevel:taxRow.riskLevel,
      drawer:taxRow.drawer||rec.drawer,remark:taxRow.remark,differenceCount:checks.length,
      note:taxAbnormal?`税务状态=${taxRow.taxStatus||'空'}；是否正数=${taxRow.isPositive||'空'}；风险等级=${taxRow.riskLevel||'空'}`:checks.map(d=>`${d.label}:${d.action}`).join('；')
    };
    results.push(row);if(taxAbnormal)abnormalRows.push(row);

    const recDetails=recDetailMap.get(rec.invoiceNumber)||[];
    const taxDetails=taxDetailMap.get(taxRow.invoiceNumber)||[];
    const recAmount=recDetails.reduce((s,x)=>s+(num(x.amount)||0),0),recTax=recDetails.reduce((s,x)=>s+(num(x.tax)||0),0);
    const taxAmount=taxDetails.reduce((s,x)=>s+(num(x.amount)||0),0),taxTax=taxDetails.reduce((s,x)=>s+(num(x.tax)||0),0);
    let detailStatus='未提供PDF明细';
    if(recDetails.length){detailStatus=amountsEqual(recAmount,taxAmount)&&amountsEqual(recTax,taxTax)?'明细金额一致':'明细金额存在差异';}
    detailRows.push({sourceFileName:rec.sourceFileName,invoiceNumber:taxRow.invoiceNumber,pdfLineCount:recDetails.length,taxLineCount:taxDetails.length,pdfAmount:recDetails.length?recAmount:null,taxAmount:taxDetails.length?taxAmount:null,pdfTax:recDetails.length?recTax:null,taxTax:taxDetails.length?taxTax:null,status:detailStatus});
  }

  const taxOnlyRows=[];
  if(includeTaxOnly&&recognition.rows.length){
    const dates=recognition.rows.map(r=>r.invoiceDate).filter(Boolean).sort();const minDate=dates[0],maxDate=dates[dates.length-1];
    const buyers=new Set(recognition.rows.map(r=>r.buyerTaxId).filter(Boolean));
    for(const t of tax.rows){
      if(matchedTaxNumbers.has(t.invoiceNumber))continue;
      if(minDate&&t.invoiceDate<minDate)continue;if(maxDate&&t.invoiceDate>maxDate)continue;
      if(buyers.size&&t.buyerTaxId&&!buyers.has(t.buyerTaxId))continue;
      taxOnlyRows.push({invoiceNumber:t.invoiceNumber,invoiceDate:t.invoiceDate,buyerName:t.buyerName,sellerName:t.sellerName,amount:t.amount,tax:t.tax,total:t.total,taxStatus:t.taxStatus,isPositive:t.isPositive,riskLevel:t.riskLevel,note:'税务局有记录，但本批识别结果中没有对应PDF'});
    }
  }
  return {results,differenceRows,abnormalRows,pdfOnlyRows,taxOnlyRows,detailRows};
}

function badge(status){return `<span class="badge ${status}">${safe(STATUS_TEXT[status]||status)}</span>`;}
function filterBySearch(rows){const q=$('#searchInput').value.trim().toLowerCase();if(!q)return rows;return rows.filter(r=>Object.values(r).join(' ').toLowerCase().includes(q));}
function table(headers,rows){if(!rows.length)return '<div class="empty">当前没有记录</div>';return `<table class="data-table"><thead><tr>${headers.map(h=>`<th>${safe(h.label)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${headers.map(h=>`<td class="${h.money?'money':''}" title="${safe(h.get?r[h.key]??h.get(r):r[h.key])}">${h.html?(h.get?h.get(r):r[h.key]):safe(h.money?money(r[h.key]):(h.get?h.get(r):r[h.key]))}</td>`).join('')}</tr>`).join('')}</tbody></table>`;}

function render(){
  const view=state.currentView;let html='';
  if(view==='summary')html=table([
    {label:'状态',key:'status',html:true,get:r=>badge(r.status)},{label:'匹配方式',key:'matchMethod'},{label:'原始文件名',key:'sourceFileName'},{label:'发票号码',key:'invoiceNumber'},{label:'开票日期',key:'invoiceDate'},{label:'购买方',key:'buyerName'},{label:'销售方',key:'sellerName'},{label:'不含税金额',key:'amount',money:true},{label:'税额',key:'tax',money:true},{label:'价税合计',key:'total',money:true},{label:'税务状态',key:'taxStatus'},{label:'说明',key:'note'}
  ],filterBySearch(state.results));
  else if(view==='differences')html=table([
    {label:'文件名',key:'sourceFileName'},{label:'发票号码',key:'invoiceNumber'},{label:'字段',key:'field'},{label:'PDF识别值',key:'pdfValue'},{label:'税务局值',key:'taxValue'},{label:'处理',key:'action'},{label:'匹配方式',key:'matchMethod'}
  ],filterBySearch(state.differenceRows));
  else if(view==='abnormal')html=table([
    {label:'发票号码',key:'invoiceNumber'},{label:'开票日期',key:'invoiceDate'},{label:'销售方',key:'sellerName'},{label:'价税合计',key:'total',money:true},{label:'发票状态',key:'taxStatus'},{label:'是否正数',key:'isPositive'},{label:'风险等级',key:'riskLevel'},{label:'说明',key:'note'}
  ],filterBySearch(state.abnormalRows));
  else if(view==='pdfOnly')html=table([
    {label:'文件名',key:'sourceFileName'},{label:'发票号码',key:'invoiceNumber'},{label:'开票日期',key:'invoiceDate'},{label:'购买方',key:'buyerName'},{label:'销售方',key:'sellerName'},{label:'价税合计',key:'total',money:true},{label:'说明',key:'note'}
  ],filterBySearch(state.pdfOnlyRows));
  else if(view==='taxOnly')html=table([
    {label:'发票号码',key:'invoiceNumber'},{label:'开票日期',key:'invoiceDate'},{label:'购买方',key:'buyerName'},{label:'销售方',key:'sellerName'},{label:'价税合计',key:'total',money:true},{label:'发票状态',key:'taxStatus'},{label:'说明',key:'note'}
  ],filterBySearch(state.taxOnlyRows));
  else html=table([
    {label:'文件名',key:'sourceFileName'},{label:'发票号码',key:'invoiceNumber'},{label:'PDF明细行数',key:'pdfLineCount'},{label:'税局明细行数',key:'taxLineCount'},{label:'PDF明细金额',key:'pdfAmount',money:true},{label:'税局明细金额',key:'taxAmount',money:true},{label:'PDF明细税额',key:'pdfTax',money:true},{label:'税局明细税额',key:'taxTax',money:true},{label:'结论',key:'status'}
  ],filterBySearch(state.detailRows));
  $('#tableWrap').innerHTML=html;
}

function updateMetrics(){
  const count=s=>state.results.filter(r=>r.status===s).length;
  $('#mTotal').textContent=state.results.length;$('#mMatched').textContent=count('matched');$('#mCorrected').textContent=count('corrected');$('#mDifference').textContent=count('difference');$('#mAbnormal').textContent=count('tax_abnormal');$('#mPdfOnly').textContent=count('pdf_only');$('#mTaxOnly').textContent=state.taxOnlyRows.length;
}

async function run(){
  try{
    $('#runBtn').disabled=true;setProgress(8,'正在读取网站识别结果',state.recognitionFile.name);
    const recWb=await readWorkbook(state.recognitionFile);const recognition=parseRecognitionWorkbook(recWb);
    setProgress(34,'正在读取税务局全量报表',state.taxFile.name);
    const taxWb=await readWorkbook(state.taxFile);const tax=parseTaxWorkbook(taxWb);
    setProgress(66,'正在逐张匹配与字段校验',`识别结果 ${recognition.rows.length} 张；税务局基础记录 ${tax.rows.length} 张`);
    const out=reconcile(recognition,tax,$('#includeTaxOnly').checked,$('#useTaxTruth').checked);
    Object.assign(state,{recognitionRows:recognition.rows,recognitionDetails:recognition.details,taxRows:tax.rows,taxDetails:tax.details,...out});
    setProgress(100,'对账完成',`完全一致 ${out.results.filter(r=>r.status==='matched').length} 张；自动修正 ${out.results.filter(r=>r.status==='corrected').length} 张；差异/异常 ${out.results.filter(r=>['difference','tax_abnormal','pdf_only'].includes(r.status)).length} 张`);
    updateMetrics();render();$('#results').classList.remove('hidden');showToast('税务局对账完成，可查看差异或导出结果');
  }catch(e){console.error(e);showToast('对账失败：'+(e.message||e));setProgress(0,'对账失败',e.message||String(e));}
  finally{updateRunState();}
}

function exportWorkbook(){
  if(!state.results.length)return;
  const summary=state.results.map(r=>({'对账状态':STATUS_TEXT[r.status]||r.status,'匹配方式':r.matchMethod,'原始文件名':r.sourceFileName,'发票号码':r.invoiceNumber,'开票日期':r.invoiceDate,'购买方名称':r.buyerName,'购买方税号':r.buyerTaxId,'销售方名称':r.sellerName,'销售方税号':r.sellerTaxId,'不含税金额':num(r.amount),'税额':num(r.tax),'价税合计':num(r.total),'发票票种':r.invoiceType,'发票状态':r.taxStatus,'是否正数发票':r.isPositive,'风险等级':r.riskLevel,'开票人':r.drawer,'差异数量':r.differenceCount||0,'说明':r.note}));
  const differences=state.differenceRows.map(r=>({'原始文件名':r.sourceFileName,'发票号码':r.invoiceNumber,'字段':r.field,'PDF识别值':r.pdfValue,'税务局值':r.taxValue,'处理方式':r.action,'严重程度':r.severity,'匹配方式':r.matchMethod}));
  const abnormal=state.abnormalRows.map(r=>({'发票号码':r.invoiceNumber,'开票日期':r.invoiceDate,'销售方':r.sellerName,'价税合计':num(r.total),'发票状态':r.taxStatus,'是否正数发票':r.isPositive,'风险等级':r.riskLevel,'说明':r.note}));
  const pdfOnly=state.pdfOnlyRows.map(r=>({'原始文件名':r.sourceFileName,'发票号码':r.invoiceNumber,'开票日期':r.invoiceDate,'购买方':r.buyerName,'销售方':r.sellerName,'价税合计':num(r.total),'说明':r.note}));
  const taxOnly=state.taxOnlyRows.map(r=>({'发票号码':r.invoiceNumber,'开票日期':r.invoiceDate,'购买方':r.buyerName,'销售方':r.sellerName,'不含税金额':num(r.amount),'税额':num(r.tax),'价税合计':num(r.total),'发票状态':r.taxStatus,'是否正数发票':r.isPositive,'风险等级':r.riskLevel,'说明':r.note}));
  const details=state.detailRows.map(r=>({'原始文件名':r.sourceFileName,'发票号码':r.invoiceNumber,'PDF明细行数':r.pdfLineCount,'税务局明细行数':r.taxLineCount,'PDF明细金额':num(r.pdfAmount),'税务局明细金额':num(r.taxAmount),'PDF明细税额':num(r.pdfTax),'税务局明细税额':num(r.taxTax),'核对结论':r.status}));
  const log=[{'处理时间':new Date().toLocaleString('zh-CN',{hour12:false}),'识别结果文件':state.recognitionFile.name,'税务局文件':state.taxFile.name,'识别结果条数':state.recognitionRows.length,'税务局基础记录数':state.taxRows.length,'税务局明细记录数':state.taxDetails.length,'是否检查税局有PDF无':$('#includeTaxOnly').checked?'是':'否','是否以税务局为准':$('#useTaxTruth').checked?'是':'否'}];
  const wb=XLSX.utils.book_new();
  [['对账汇总',summary],['字段差异',differences],['税务状态异常',abnormal],['PDF有税局无',pdfOnly],['税局有PDF无',taxOnly],['明细核对',details],['处理日志',log]].forEach(([name,data])=>{const ws=XLSX.utils.json_to_sheet(data.length?data:[{'暂无数据':''}]);ws['!autofilter']={ref:ws['!ref']};ws['!cols']=Object.keys(data[0]||{'暂无数据':''}).map(k=>({wch:Math.min(42,Math.max(12,k.length*2+2,...data.slice(0,100).map(x=>String(x[k]??'').length+2)))}));XLSX.utils.book_append_sheet(wb,ws,name);});
  const d=new Date(),p=n=>String(n).padStart(2,'0');XLSX.writeFile(wb,`发票识别与税务局对账结果_${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}.xlsx`);
}

function bindFile(button,input,stateKey,card,stateEl){$(button).onclick=()=>$(input).click();$(input).onchange=e=>{const f=e.target.files?.[0];if(!f)return;state[stateKey]=f;$(card).classList.add('ready');$(stateEl).textContent=f.name;updateRunState();};}
bindFile('#chooseRecognition','#recognitionInput','recognitionFile','#recognitionCard','#recognitionState');
bindFile('#chooseTax','#taxInput','taxFile','#taxCard','#taxState');
$('#runBtn').onclick=run;$('#exportBtn').onclick=exportWorkbook;$('#searchInput').oninput=render;
$('#tabs').onclick=e=>{const b=e.target.closest('[data-view]');if(!b)return;document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));b.classList.add('active');state.currentView=b.dataset.view;render();};
updateRunState();
