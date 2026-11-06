const {test}=require('node:test'),assert=require('node:assert/strict'),fixture=require('./hwp-cell-break-policy-fixture.cjs');
for(const inline of [false,true])for(const nested of [false,true])for(const format of ['hwp','hwpx'])for(const policy of inline?[2]:[0,1,2])test(`${format} ${nested?'nested':'ordinary'} ${['NONE','TABLE','CELL'][policy]} ${inline?'inline':'floating'}: preserve declared cell splitting even when the original row fits one page`,async()=>{
 const sample=(await fixture())[`${format}-${policy}-${nested?'nested':'ordinary'}${inline?'-inline':''}`];assert.equal(sample.sourcePages,1);assert.equal(sample.sourceCellPages,1);assert.equal(sample.policy,nested?2:policy);assert.equal(sample.xmlPolicy,['NONE','TABLE','CELL'][policy]);
 const rows=sample.editable.html.match(/<tr\b[^>]*>/g);assert.equal(rows.length,nested?2:1);assert.equal(rows[nested?1:0].includes('data-row-split="allow"'),policy===2&&!inline);
 assert.equal((sample.editable.html.match(/LEFT \d+/g)||[]).length,24);assert.equal((sample.editable.html.match(/RIGHT \d+/g)||[]).length,24);
});
