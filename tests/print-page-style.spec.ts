import{test,expect}from'@playwright/test';
import{writeFile}from'node:fs/promises';
test('many printed pages preserve section paper sizes without invalidating document styles per page',async({page})=>{
 await page.goto('/');
 const result=await page.evaluate(async()=>{
  const{createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts');
  const first={...defaultPage,size:'Custom' as const,width:120,height:180,margin:10,pageNumbers:false};
  const second={...defaultPage,size:'Custom' as const,width:180,height:120,margin:10,pageNumbers:false,breakType:'nextPage' as const};
  const block=(label:string)=>Array.from({length:6},(_,i)=>`<p style="font-family:Arial;font-size:11pt;line-height:20px;margin:0">${label}${i}</p>${i<5?'<div data-page-break="true"></div>':''}`).join('');
  const html=block('FIRST')+`<div data-section-page='${JSON.stringify(second)}'></div>`+block('SECOND');
  const mutations:MutationRecord[]=[];const observer=new MutationObserver(records=>mutations.push(...records));observer.observe(document.body,{subtree:true,childList:true});
  const layout=await createPrintLayout(html,first);const invalidations=[...mutations,...observer.takeRecords()].filter(r=>r.target instanceof HTMLStyleElement&&r.target.closest('.print-layout')).length;observer.disconnect();
  layout.element.classList.replace('measuring','preview');document.body.classList.add('studio-printing');
  const papers=[...layout.element.querySelectorAll<HTMLElement>('.print-page')].map(p=>({width:p.getBoundingClientRect().width,height:p.getBoundingClientRect().height,text:p.querySelector('.print-content')?.textContent}));
  return{invalidations,papers,pages:layout.pages};
 });
 expect(result.invalidations).toBeLessThanOrEqual(2);expect(result.pages).toBe(12);
 const pdf=await page.pdf({preferCSSPageSize:true,printBackground:true});await writeFile(test.info().outputPath('named-paper-sections.pdf'),pdf);await writeFile(test.info().outputPath('named-paper-styles.json'),JSON.stringify(result,null,2));
 // Chromium's controlled output writes the effective media box on every page.
 const boxes=[...pdf.toString('latin1').matchAll(/\/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]/g)].map(m=>[Number(m[1]),Number(m[2])]);
 expect(boxes).toHaveLength(12);
 for(let i=0;i<12;i++){
  expect(result.papers[i].text).toBe((i<6?'FIRST':'SECOND')+(i%6));const width=i<6?120:180,height=i<6?180:120;
  expect(Math.abs(boxes[i][0]-width*72/25.4)).toBeLessThan(1);expect(Math.abs(boxes[i][1]-height*72/25.4)).toBeLessThan(1);
 }
});
