const test=require('node:test'),assert=require('node:assert/strict'),fixture=require('./hwp-range-bookmark-fixture.cjs');
for(const format of ['hwp','hwpx'])for(const kind of ['ordinary','nested'])test(`${format} ${kind}: range bookmark starts retain names without adding visible fallback text`,async()=>{
 const source=(await fixture())[format+'-'+kind];assert.equal(source.sourcePages,1);
 assert.equal((source.editable.html.match(/data-docx-bookmark=/g)||[]).length,2);
 assert.ok(source.editable.html.includes('정부 범위 BODY'));assert.ok(source.editable.html.includes('정부 범위 CELL'));
 assert.ok(!source.editable.html.includes('원본 보기에서 확인'));assert.ok(source.editable.html.includes('BOOKMARK_BODY'));assert.ok(source.editable.html.includes('BOOKMARK_CELL'));
});
for(const kind of ['unknown','invalid-name'])test(`${kind} fields remain explicit and retain their authored text`,async()=>{
 const source=(await fixture())[kind];assert.ok(!source.editable.html.includes('data-docx-bookmark='));assert.ok(source.editable.html.includes('원본 보기에서 확인'));assert.ok(source.editable.html.includes('BOOKMARK_BODY'));assert.ok(source.editable.html.includes('BOOKMARK_CELL'));
});
