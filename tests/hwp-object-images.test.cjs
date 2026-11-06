const{test}=require('node:test'),assert=require('node:assert/strict'),fixture=require('./hwp-object-image-fixture.cjs');
let pending;const samples=()=>pending??=fixture();
const decode=s=>s.replace(/&quot;/g,'"').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>');
for(const format of ['hwp','hwpx'])test(`${format}: an isolated group retains its own image bytes, crop and dotted tab leaders`,async()=>{
 const sample=(await samples())[format],images=sample.ops.filter(op=>op.type==='image'),tabs=sample.ops.filter(op=>op.type==='tabLeader');
 assert.equal(images.length,1,'the source contains an actual picture');assert.equal(tabs.length,1,'the source contains a native leader');assert.equal(tabs[0].leaders[0].fillType,3);
 const raw=sample.editable.html.match(/data-hwp-object="([^"]+)"/)?.[1];assert.ok(raw,'the complete group is not replaced by a placeholder');
 const root=JSON.parse(decode(raw)),ops=[];function walk(n){ops.push(...n.ops??[]);n.children?.forEach(walk);if(n.child)walk(n.child);}walk(root);
 assert.equal(ops.filter(o=>o.type==='image').length,1);assert.equal(ops.filter(o=>o.type==='tabLeader').length,1);
 const img=ops.find(o=>o.type==='image');assert.match(img.base64,/^[a-zA-Z0-9+/]+=*$/);assert.equal(img.mime,'image/png');assert.ok(img.originalSizeHu.every(n=>n>0),'crop endpoints have an explicit original size');
 assert.equal(img.crop.left,0);assert.ok(img.crop.right<img.originalSizeHu[0],'the source crop is not silently ignored');
 assert.ok(ops.every(o=>o.type!=='pageBackground'),'no page background is copied');assert.doesNotMatch(raw,/BODY OUTSIDE GROUP/);assert.match(decode(sample.editable.html),/TITLE 한글/);
 assert.doesNotMatch(sample.editable.html,/원본 보기에서 확인/);
});

let pure;
for(const format of ['hwp','hwpx'])test(`${format}: an inline image-only group retains its declared height without extra preview padding`,async()=>{
 const sample=(await(pure??=fixture({pureImage:true})))[format],raw=sample.editable.html.match(/data-hwp-object="([^"]+)"/)?.[1];assert.ok(raw);
 const root=JSON.parse(decode(raw)),tag=sample.editable.html.match(/<img\b[^>]*data-hwp-object=[^>]*>/)?.[0];assert.ok(tag);
 assert.equal(root.previewPadding,0,'pictures and an unbordered fill have no ink beyond the declared bounds');
 assert.match(sample.editable.html,/data-word-text-alignment="top"/,'an object-only native line retains its measured top alignment');
 assert.equal(Number(tag.match(/\bheight="([^"]+)"/)[1]),root.bounds.height,'extra preview pixels must not increase the source line budget');
});

let hidden;
for(const format of ['hwp','hwpx'])test(`${format}: a nonvisual PageHide on the same line does not add a font descender to the only graphic`,async()=>{
 const sample=(await(hidden??=fixture({pureImage:true,pageHide:true})))[format];
 assert.match(sample.editable.html,/data-hwp-object=/,'the actual source group survives');
 assert.match(sample.editable.html,/data-word-text-alignment="top"/,'PageHide keeps its story policy without blocking the only graphic line alignment');
});
