(async()=>{
  try{
    const version='20260727-dedupe1';
    const r=await fetch('final-engine.gz.b64?v='+version,{cache:'no-store'});
    if(!r.ok)throw new Error('最终识别引擎下载失败：'+r.status);
    const b64=(await r.text()).trim();
    const raw=Uint8Array.from(atob(b64),c=>c.charCodeAt(0));
    let code;
    if(typeof DecompressionStream!=='undefined'){
      const ds=new DecompressionStream('gzip');
      code=await new Response(new Blob([raw]).stream().pipeThrough(ds)).text();
    }else{
      await new Promise((res,rej)=>{
        const s=document.createElement('script');
        s.src='https://cdn.jsdelivr.net/npm/pako@2.1.0/dist/pako.min.js';
        s.onload=res;s.onerror=rej;document.head.appendChild(s);
      });
      code=new TextDecoder().decode(window.pako.ungzip(raw));
    }

    const totalsDeclaration="const details=parseDetails(items,width,height,rows);const sumAmount=Number(details.reduce((s,x)=>s+(numberValue(x.amount)||0),0).toFixed(2)),sumTax=Number(details.reduce((s,x)=>s+(numberValue(x.taxAmount)||0),0).toFixed(2));";
    const patchedDeclaration="let details=parseDetails(items,width,height,rows);let sumAmount=Number(details.reduce((s,x)=>s+(numberValue(x.amount)||0),0).toFixed(2)),sumTax=Number(details.reduce((s,x)=>s+(numberValue(x.taxAmount)||0),0).toFixed(2));";
    if(!code.includes(totalsDeclaration))throw new Error('未找到金额汇总逻辑，已停止加载以避免使用旧算法');
    code=code.replace(totalsDeclaration,patchedDeclaration);

    const totalsDecision="if(details.length&&totalAmountWithTax!==null&&Math.abs(sumAmount+sumTax-totalAmountWithTax)<=.02){totalAmount=sumAmount;totalTax=sumTax;}else{if(totalAmount===null&&details.length)totalAmount=sumAmount;if(totalTax===null&&details.length)totalTax=sumTax;if(totalAmountWithTax===null&&totalAmount!==null&&totalTax!==null)totalAmountWithTax=Number((totalAmount+totalTax).toFixed(2));}";
    const patchedDecision=`if(details.length>1&&totalAmountWithTax!==null&&Math.abs(sumAmount+sumTax-totalAmountWithTax*2)<=.03){
      const detailKey=x=>[compact(x.projectName),compact(x.specification),compact(x.unit),x.quantity??'',x.unitPrice??'',x.amount??'',compact(x.taxRate),x.taxAmount??''].join('|');
      const buckets=new Map();
      for(const x of details){const k=detailKey(x);if(!buckets.has(k))buckets.set(k,[]);buckets.get(k).push(x);}
      const groups=[...buckets.values()];
      if(groups.length&&groups.every(g=>g.length%2===0)){
        details=groups.flatMap(g=>g.slice(0,g.length/2));
        sumAmount=Number(details.reduce((s,x)=>s+(numberValue(x.amount)||0),0).toFixed(2));
        sumTax=Number(details.reduce((s,x)=>s+(numberValue(x.taxAmount)||0),0).toFixed(2));
      }
    }
    if(details.length&&totalAmountWithTax!==null&&Math.abs(sumAmount+sumTax-totalAmountWithTax)<=.02){totalAmount=sumAmount;totalTax=sumTax;}else{if(totalAmount===null&&details.length)totalAmount=sumAmount;if(totalTax===null&&details.length)totalTax=sumTax;if(totalAmountWithTax===null&&totalAmount!==null&&totalTax!==null)totalAmountWithTax=Number((totalAmount+totalTax).toFixed(2));}`;
    if(!code.includes(totalsDecision))throw new Error('未找到金额校验逻辑，已停止加载以避免使用旧算法');
    code=code.replace(totalsDecision,patchedDecision);

    new Function(code+'\n//# sourceURL=final-engine-runtime-dedupe1.js')();
  }catch(e){
    console.error(e);
    document.body.insertAdjacentHTML('afterbegin','<div style="position:fixed;z-index:9999;left:18px;right:18px;top:18px;padding:16px;border-radius:12px;background:#4a2028;color:#fff;border:1px solid #ff8a96;font-family:system-ui">网站最终识别引擎加载失败：'+String(e.message||e).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))+'。请按 Ctrl+F5 刷新页面。</div>');
  }
})();
