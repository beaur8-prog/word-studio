const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path'),ts=require('typescript');
function load(file){const mod={exports:{}};new Function('require','module','exports',ts.transpileModule(fs.readFileSync(file,'utf8'),{compilerOptions:{target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.CommonJS}}).outputText)(name=>name.startsWith('.')?load(path.resolve(path.dirname(file),name+'.ts')):require(name),mod,mod.exports);return mod.exports;}
const source=path.resolve(__dirname,'../src/hwp/metafile-preview.ts'),mod={exports:load(source)};
const{metafilePreview}=mod.exports,{emfBytes}=require('./hwp-metafile-fixture.cjs');
test('unknown octet-stream data is not guessed to be an EMF or rewritten',()=>{
 const html='<p><img src="data:application/octet-stream;base64,'+Buffer.alloc(96).toString('base64')+'" width="100" height="80"></p>';
 assert.equal(metafilePreview(html,{},new Map()),html);
});
test('an identified EMF with a contradictory byte count fails before rendering instead of losing the picture',()=>{
 const bytes=emfBytes();bytes.writeUInt32LE(bytes.length+4,48);
 assert.throws(()=>metafilePreview('<img src="data:application/octet-stream;base64,'+bytes.toString('base64')+'">',{},new Map()),/파일 구조/);
});

const{restoreMetafileGdi}=load(path.resolve(__dirname,'../src/hwp/metafile-gdi.ts'));
const{formEmfBytes}=require('./hwp-metafile-fixture.cjs');
const formSVG='<svg xmlns="http://www.w3.org/2000/svg"><g transform="matrix(1 0 0 1 0 0)"><g><text x="5" y="5">TOPTEXT</text></g><g><line x1="0" y1="40" x2="200" y2="40"/></g></g></svg>';
test('EMF GDI restoration retains PATCOPY paint order and TA_TOP instead of an alphabetic baseline',()=>{
 const out=restoreMetafileGdi(formSVG,formEmfBytes());assert.match(out,/<rect x="0" y="0" width="200" height="40" fill="rgb\(192,192,192\)"/);assert.match(out,/dominant-baseline="text-before-edge"/);assert.ok(out.indexOf('<rect')<out.indexOf('<text'));assert.ok(out.indexOf('<text')<out.indexOf('<line'));assert.match(out,/matrix\(1 0 0 1 0 0\)/);
});
test('mismatched native text and unknown painting records never receive guessed EMF corrections',()=>{
 assert.equal(restoreMetafileGdi(formSVG.replace('TOPTEXT','OTHER'),formEmfBytes()),formSVG.replace('TOPTEXT','OTHER'));
 const bytes=formEmfBytes(),offset=88;bytes.writeUInt32LE(999,offset);assert.equal(restoreMetafileGdi(formSVG,bytes),formSVG);
 assert.equal(restoreMetafileGdi(formSVG,formEmfBytes().subarray(0,-3)),formSVG);
});
test('an explicit TA_BASELINE reference point stays unchanged while form background is restored',()=>{
 const bytes=formEmfBytes();for(let o=88;o<bytes.length;o+=bytes.readUInt32LE(o+4))if(bytes.readUInt32LE(o)===22)bytes.writeUInt32LE(24,o+8);
 const out=restoreMetafileGdi(formSVG,bytes);assert.doesNotMatch(out,/dominant-baseline/);assert.match(out,/rgb\(192,192,192\)/);
});
test('saved GDI brush and cell clip restore before the next PATCOPY operation',()=>{
 const rec=(t,a)=>{const b=Buffer.alloc(8+a.length*4);b.writeUInt32LE(t);b.writeUInt32LE(b.length,4);a.forEach((x,i)=>b.writeUInt32LE(x>>>0,8+i*4));return b;};
 const base=formEmfBytes(),parts=[];let fill;
 for(let o=88;o<base.length;o+=base.readUInt32LE(o+4))if(base.readUInt32LE(o)===76)fill=base.subarray(o,o+100);
 const extra=[rec(33,[]),rec(30,[10,10,40,30]),rec(39,[3,0,255,0]),rec(37,[3]),fill,rec(34,[-1]),fill];
 const bytes=Buffer.concat([base.subarray(0,-20),...extra,base.subarray(-20)]);bytes.writeUInt32LE(bytes.length,48);
 const out=restoreMetafileGdi(formSVG,bytes);assert.match(out,/<rect x="10" y="10" width="30" height="20" fill="rgb\(255,0,0\)"/);assert.equal((out.match(/fill="rgb\(192,192,192\)"/g)||[]).length,2);assert.ok(out.lastIndexOf('rgb(192,192,192)')>out.indexOf('rgb(255,0,0)'));
});
