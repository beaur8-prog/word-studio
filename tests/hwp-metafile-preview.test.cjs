const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),ts=require('typescript');
const source=path.resolve(__dirname,'../src/hwp/metafile-preview.ts'),mod={exports:{}};
new Function('module','exports',ts.transpileModule(fs.readFileSync(source,'utf8'),{compilerOptions:{target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.CommonJS}}).outputText)(mod,mod.exports);
const{metafilePreview}=mod.exports,{emfBytes}=require('./hwp-metafile-fixture.cjs');
test('unknown octet-stream data is not guessed to be an EMF or rewritten',()=>{
 const html='<p><img src="data:application/octet-stream;base64,'+Buffer.alloc(96).toString('base64')+'" width="100" height="80"></p>';
 assert.equal(metafilePreview(html,{},new Map()),html);
});
test('an identified EMF with a contradictory byte count fails before rendering instead of losing the picture',()=>{
 const bytes=emfBytes();bytes.writeUInt32LE(bytes.length+4,48);
 assert.throws(()=>metafilePreview('<img src="data:application/octet-stream;base64,'+bytes.toString('base64')+'">',{},new Map()),/파일 구조/);
});
