(async()=>{
  try{
    const response=await fetch('./app.js?v=20260725-ai1',{cache:'no-store'});
    if(!response.ok)throw new Error('主程序加载失败：'+response.status);
    let source=await response.text();
    const broken="function extractTextLayer(page,viewport){return page.getTextContent().then(tc=>tc.items.filter(x=>x.str&&x.str.trim()).map(x=>{const x=x.transform[4]*viewport.scale,y=viewport.height-x.transform[5]*viewport.scale,w=(x.width||0)*viewport.scale,h=Math.abs(x.transform[3]||10)*viewport.scale;return{text:x.str,score:1,x,y,w,h,cx:x+w/2,cy:y+h/2,source:'pdf'};}));}";
    const fixed="function extractTextLayer(page,viewport){return page.getTextContent().then(tc=>tc.items.filter(item=>item.str&&item.str.trim()).map(item=>{const px=item.transform[4]*viewport.scale,py=viewport.height-item.transform[5]*viewport.scale,pw=(item.width||0)*viewport.scale,ph=Math.abs(item.transform[3]||10)*viewport.scale;return{text:item.str,score:1,x:px,y:py,w:pw,h:ph,cx:px+pw/2,cy:py+ph/2,source:'pdf'};}));}";
    if(!source.includes(broken))throw new Error('测试程序版本不匹配，未找到需要修复的代码段');
    source=source.replace(broken,fixed).replaceAll('XLSX.utils.','window.XLSX.utils.').replaceAll('XLSX.writeFile(','window.XLSX.writeFile(');
    const blob=new Blob([source+'\n//# sourceURL=ai-test/app-runtime.js'],{type:'text/javascript'});
    const url=URL.createObjectURL(blob);
    await import(url);
    setTimeout(()=>URL.revokeObjectURL(url),10000);
  }catch(error){
    console.error(error);
    document.body.insertAdjacentHTML('afterbegin',`<div style="position:fixed;z-index:9999;left:20px;right:20px;top:20px;padding:18px;border-radius:12px;background:#4a2028;color:#fff;border:1px solid #ff8995;font-family:system-ui">免费AI测试页启动失败：${String(error.message||error).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}</div>`);
  }
})();
