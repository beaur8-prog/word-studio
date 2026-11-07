const fs=require('node:fs/promises'),path=require('node:path'),{pathToFileURL}=require('node:url');
// Independently authored MS-EMF fixture: red left half, blue right half.
// EMR_HEADER / CREATEBRUSHINDIRECT / SELECTOBJECT / RECTANGLE / EOF.
function emfBytes(){
 const rec=(type,values)=>{const b=Buffer.alloc(8+values.length*4);b.writeUInt32LE(type);b.writeUInt32LE(b.length,4);values.forEach((v,i)=>b.writeUInt32LE(v>>>0,8+i*4));return b;};
 const records=[rec(39,[1,0,0x000000ff,0]),rec(39,[2,0,0x00ff0000,0]),rec(37,[0x80000008]),rec(37,[1]),rec(43,[0,0,100,80]),rec(37,[2]),rec(43,[100,0,200,80]),rec(14,[0,0,20])];
 const h=Buffer.alloc(88);h.writeUInt32LE(1);h.writeUInt32LE(88,4);h.writeInt32LE(200,16);h.writeInt32LE(80,20);h.writeInt32LE(5292,32);h.writeInt32LE(2117,36);h.writeUInt32LE(0x464d4520,40);h.writeUInt32LE(0x10000,44);h.writeUInt32LE(88+records.reduce((s,b)=>s+b.length,0),48);h.writeUInt32LE(records.length+1,52);h.writeUInt16LE(3,56);h.writeInt32LE(960,72);h.writeInt32LE(384,76);h.writeInt32LE(254,80);h.writeInt32LE(102,84);return Buffer.concat([h,...records]);
}
async function fixture(directory){
 const {unzipSync,zipSync,strFromU8,strToU8}=require('fflate');
 const original=await require('./hwp-crop-fixture.cjs')(directory),entries=unzipSync(await fs.readFile(original.hwpx)),emf=emfBytes(),updated={};
 // Retain the independent native crop fixture's body/cell structure and
 // replace only its known two-colour asset with the equivalent EMF asset.
 for(const[name,bytes]of Object.entries(entries)){
  if(/^BinData\/.*\.png$/.test(name))updated[name.replace(/\.png$/,'.emf')]=emf;
  else if(/\.(?:xml|hpf)$/.test(name))updated[name]=strToU8(strFromU8(bytes).replace(/\.png(?=["'])/g,'.emf').replace(/image\/png/g,'image/x-emf').replace(/format="png"/g,'format="emf"'));
  else updated[name]=bytes;
 }
 const core=path.dirname(require.resolve('@rhwp/core')),{default:init,HwpDocument:C}=await import(pathToFileURL(core+'/rhwp.js'));await init({module_or_path:await fs.readFile(core+'/rhwp_bg.wasm')});
 const d=new C(zipSync(updated));try{
  const files={};for(const f of ['hwp','hwpx']){files[f]=path.join(directory,'metafile.'+f);await fs.writeFile(files[f],f==='hwp'?d.exportHwp():d.exportHwpx());}return files;
 }finally{d.free();}
}
module.exports={fixture,emfBytes};
