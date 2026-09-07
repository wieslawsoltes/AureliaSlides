import {makeDocument,makeElement,makeSlide,validateDocument,color,clamp,escapeHTML} from './core.js';
import {expandElement,paintSlide2D,ImagePool} from './renderer.js';
/** ZIP + PresentationML interchange. No network, macros, executable content or external relationships. */
const UTF8=new TextEncoder();
const CRC_TABLE=Uint32Array.from({length:256},(_,i)=>{let c=i;for(let k=0;k<8;k++)c=(c&1)?0xedb88320^(c>>>1):c>>>1;return c>>>0;});
export function crc32(bytes){let c=0xffffffff;for(const b of bytes)c=CRC_TABLE[(c^b)&255]^(c>>>8);return(c^0xffffffff)>>>0;}
export function zipStore(files){
 const chunks=[],central=[];let offset=0;
 for(const [name,input]of Object.entries(files)){
  const path=UTF8.encode(name),data=typeof input==='string'?UTF8.encode(input):new Uint8Array(input),crc=crc32(data),local=new Uint8Array(30+path.length),v=new DataView(local.buffer);v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,0x800,true);v.setUint32(14,crc,true);v.setUint32(18,data.length,true);v.setUint32(22,data.length,true);v.setUint16(26,path.length,true);local.set(path,30);chunks.push(local,data);
  const head=new Uint8Array(46+path.length),h=new DataView(head.buffer);h.setUint32(0,0x02014b50,true);h.setUint16(4,20,true);h.setUint16(6,20,true);h.setUint16(8,0x800,true);h.setUint32(16,crc,true);h.setUint32(20,data.length,true);h.setUint32(24,data.length,true);h.setUint16(28,path.length,true);h.setUint32(42,offset,true);head.set(path,46);central.push(head);offset+=local.length+data.length;
 }
 const size=central.reduce((s,c)=>s+c.length,0),end=new Uint8Array(22),v=new DataView(end.buffer);v.setUint32(0,0x06054b50,true);v.setUint16(8,central.length,true);v.setUint16(10,central.length,true);v.setUint32(12,size,true);v.setUint32(16,offset,true);return new Blob([...chunks,...central,end],{type:'application/zip'});
}
export async function unzipSafe(buffer){
 const bytes=new Uint8Array(buffer),view=new DataView(buffer),decoder=new TextDecoder();if(bytes.length>100*1024*1024)throw new Error('Presentation files are limited to 100 MB.');let end=-1;
 for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(view.getUint32(i,true)===0x06054b50){end=i;break;}if(end<0)throw new Error('Invalid ZIP / PowerPoint file.');
 const count=view.getUint16(end+10,true);if(count>10000)throw new Error('Too many files in the presentation.');let at=view.getUint32(end+16,true),total=0;const files=new Map();
 for(let n=0;n<count;n++){
  if(at+46>bytes.length||view.getUint32(at,true)!==0x02014b50)throw new Error('Corrupt ZIP directory.');
  const flags=view.getUint16(at+8,true),method=view.getUint16(at+10,true),crc=view.getUint32(at+16,true),size=view.getUint32(at+20,true),rawSize=view.getUint32(at+24,true),nameLen=view.getUint16(at+28,true),extra=view.getUint16(at+30,true),comment=view.getUint16(at+32,true),pos=view.getUint32(at+42,true),name=decoder.decode(bytes.subarray(at+46,at+46+nameLen));at+=46+nameLen+extra+comment;
  if(flags&1)throw new Error('Encrypted PowerPoint files are not supported.');if(name.includes('..')||name.startsWith('/'))throw new Error('Unsafe ZIP path.');total+=rawSize;if(total>160*1024*1024||rawSize>40*1024*1024)throw new Error('Decompressed presentation exceeds the safety limit.');
  if(pos+30>bytes.length||view.getUint32(pos,true)!==0x04034b50)throw new Error('Corrupt ZIP entry.');const begin=pos+30+view.getUint16(pos+26,true)+view.getUint16(pos+28,true);if(begin+size>bytes.length)throw new Error('Truncated ZIP entry.');let data=bytes.subarray(begin,begin+size);
  if(method===8){if(typeof DecompressionStream==='undefined')throw new Error('This browser cannot decompress PPTX files. Use a current browser.');const stream=new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));const reader=stream.getReader(),parts=[];let expanded=0;while(true){const {done,value}=await reader.read();if(done)break;expanded+=value.byteLength;if(expanded>rawSize||expanded>40*1024*1024){await reader.cancel();throw new Error('ZIP expansion limit exceeded.');}parts.push(value);}data=new Uint8Array(expanded);let k=0;for(const part of parts){data.set(part,k);k+=part.byteLength;}}
  else if(method!==0)throw new Error(`Unsupported ZIP compression method ${method}.`);
  if(data.length!==rawSize||crc32(data)!==crc)throw new Error(`ZIP checksum failed: ${name}`);files.set(name,data);
 }return files;
}
const A_NS='http://schemas.openxmlformats.org/drawingml/2006/main',P_NS='http://schemas.openxmlformats.org/presentationml/2006/main',R_NS='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_HEAD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const xmlEsc=s=>escapeHTML(String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,''));
const emu=n=>Math.round(n*9525);
const ns=`xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"`;
const groupStart='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const xmlRels=items=>XML_HEAD+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id,type,target])=>`<Relationship Id="${id}" Type="${type.startsWith('http')?type:R_NS+'/'+type}" Target="${xmlEsc(target)}"/>`).join('')}</Relationships>`;
function solidXML(fill,theme,opacity=1){if(fill==='none')return'<a:noFill/>';const c=color(fill,theme).slice(1),alpha=opacity*(c.length===8?parseInt(c.slice(6),16)/255:1);return`<a:solidFill><a:srgbClr val="${c.slice(0,6)}">${alpha<1?`<a:alpha val="${Math.round(alpha*100000)}"/>`:''}</a:srgbClr></a:solidFill>`;}
function transformXML(e){return`<a:xfrm rot="${Math.round((e.rotation||0)*60000)}"><a:off x="${emu(e.x)}" y="${emu(e.y)}"/><a:ext cx="${emu(e.w)}" cy="${emu(e.h)}"/></a:xfrm>`;}
function textBodyXML(e,theme){
 const pad=emu(e.padding||0),rPr=`sz="${Math.round((e.fontSize||24)*75)}" b="${e.bold?1:0}" i="${e.italic?1:0}"${e.underline?' u="sng"':''}`;
 return`<p:txBody><a:bodyPr wrap="square" lIns="${pad}" tIns="${pad}" rIns="${pad}" bIns="${pad}" anchor="${({top:'t',middle:'ctr',bottom:'b'})[e.valign]||'t'}"><a:noAutofit/></a:bodyPr><a:lstStyle/>${String(e.text||'').split('\n').map(t=>`<a:p><a:pPr algn="${({left:'l',center:'ctr',right:'r'})[e.align]||'l'}"><a:lnSpc><a:spcPct val="${Math.round((e.lineHeight||1.2)*100000)}"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr><a:r><a:rPr lang="en-US" ${rPr}>${solidXML(e.fill,theme,e.opacity??1)}<a:latin typeface="${xmlEsc(e.fontFamily||'Arial')}"/><a:ea typeface="${xmlEsc(e.fontFamily||'Arial')}"/><a:cs typeface="${xmlEsc(e.fontFamily||'Arial')}"/></a:rPr><a:t xml:space="preserve">${xmlEsc(t)}</a:t></a:r><a:endParaRPr lang="en-US" ${rPr}/></a:p>`).join('')}</p:txBody>`;
}
function shapeXML(e,id,theme){
 const preset={rect:'rect',roundRect:'roundRect',ellipse:'ellipse',triangle:'triangle',diamond:'diamond',arrow:'rightArrow',star:'star5',line:'rect'}[e.type]||'rect';
 const adj=e.type==='roundRect'?`<a:gd name="adj" fmla="val ${Math.round(Math.min(.5,(e.radius||0)/Math.min(e.w,e.h))*100000)}"/>`:'';
 return`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xmlEsc(e.name||e.type)}"/><p:cNvSpPr${e.type==='text'?' txBox="1"':''}/><p:nvPr/></p:nvSpPr><p:spPr>${transformXML(e)}<a:prstGeom prst="${preset}"><a:avLst>${adj}</a:avLst></a:prstGeom>${e.type==='text'?'<a:noFill/>':solidXML(e.fill,theme,e.opacity??1)}<a:ln w="${emu(e.strokeWidth||0)}">${e.strokeWidth?solidXML(e.stroke,theme,e.opacity??1):'<a:noFill/>'}</a:ln></p:spPr>${e.type==='text'?textBodyXML(e,theme):''}</p:sp>`;
}
function dataBytes(src){const str=atob(src.split(',')[1]);return Uint8Array.from(str,c=>c.charCodeAt(0));}
function bytesData(bytes,type){let s='';for(let i=0;i<bytes.length;i+=32768)s+=String.fromCharCode(...bytes.subarray(i,i+32768));return`data:${type};base64,${btoa(s)}`;}
export async function exportPPTX(doc){
 const files={},parts=[],presentationRels=[['rId1','slideMaster','slideMasters/slideMaster1.xml'],['rId2','notesMaster','notesMasters/notesMaster1.xml']];let mediaCount=0;const imagePool=new ImagePool();await imagePool.ready(doc.slides.flatMap(s=>s.elements));
 const part=(name,content,type)=>{files[name]=XML_HEAD+content;parts.push([name,type]);};
 const type=s=>'application/vnd.openxmlformats-officedocument.presentationml.'+s+'+xml';
 part('ppt/presentation.xml',`<p:presentation ${ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst><p:sldIdLst>${doc.slides.map((_,i)=>`<p:sldId id="${256+i}" r:id="rId${i+3}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${emu(doc.width)}" cy="${emu(doc.height)}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle/></p:presentation>`,type('presentation.main'));
 const clrMap='<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
 part('ppt/slideMasters/slideMaster1.xml',`<p:sldMaster ${ns}><p:cSld><p:spTree>${groupStart}</p:spTree></p:cSld>${clrMap}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,type('slideMaster'));
 files['ppt/slideMasters/_rels/slideMaster1.xml.rels']=xmlRels([['rId1','slideLayout','../slideLayouts/slideLayout1.xml'],['rId2','theme','../theme/theme1.xml']]);
 part('ppt/slideLayouts/slideLayout1.xml',`<p:sldLayout ${ns} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${groupStart}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,type('slideLayout'));
 files['ppt/slideLayouts/_rels/slideLayout1.xml.rels']=xmlRels([['rId1','slideMaster','../slideMasters/slideMaster1.xml']]);
 const colors=['17313A','F5F2EC','71827F','DDDCD4',color('@accent',doc.theme).slice(1),color('@secondary',doc.theme).slice(1),'4472C4','70AD47','FFC000','5B9BD5','0563C1','954F72'];
 part('ppt/theme/theme1.xml',`<a:theme xmlns:a="${A_NS}" name="Aurelia"><a:themeElements><a:clrScheme name="Aurelia">${['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'].map((k,i)=>`<a:${k}><a:srgbClr val="${colors[i]}"/></a:${k}>`).join('')}</a:clrScheme><a:fontScheme name="Aurelia"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Aurelia"><a:fillStyleLst>${Array(3).fill('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>').join('')}</a:fillStyleLst><a:lnStyleLst>${[9525,19050,28575].map(w=>`<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`).join('')}</a:lnStyleLst><a:effectStyleLst>${Array(3).fill('<a:effectStyle><a:effectLst/></a:effectStyle>').join('')}</a:effectStyleLst><a:bgFillStyleLst>${Array(3).fill('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>').join('')}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,'application/vnd.openxmlformats-officedocument.theme+xml');
 part('ppt/notesMasters/notesMaster1.xml',`<p:notesMaster ${ns}><p:cSld><p:spTree>${groupStart}</p:spTree></p:cSld>${clrMap}<p:hf hdr="0" ftr="0" dt="0" sldNum="0"/><p:notesStyle/></p:notesMaster>`,type('notesMaster'));files['ppt/notesMasters/_rels/notesMaster1.xml.rels']=xmlRels([['rId1','theme','../theme/theme1.xml']]);
 for(let i=0;i<doc.slides.length;i++){
  const s=doc.slides[i],rels=[['rId1','slideLayout','../slideLayouts/slideLayout1.xml'],['rId2','notesSlide',`../notesSlides/notesSlide${i+1}.xml`]];let content='',id=2;
  for(const parent of s.elements.filter(e=>!e.hidden))for(let e of expandElement(parent)){
   if(e.type==='image'||e.type==='chart'){
    let src=e.src;
    if(e.type==='chart'){const c=document.createElement('canvas');c.width=Math.ceil(e.w*2);c.height=Math.ceil(e.h*2);const temp={...e,x:0,y:0,rotation:0,opacity:1};paintSlide2D(c,{bg:s.bg,elements:[temp]},{...doc,width:e.w,height:e.h},imagePool,{scale:2});src=c.toDataURL('image/png');}
    let extension=/^data:image\/jpeg/.test(src)?'jpg':'png';
    if(/^data:image\/webp/.test(src)){const im=imagePool.get(src);await im.decode();const c=document.createElement('canvas');c.width=im.naturalWidth;c.height=im.naturalHeight;c.getContext('2d').drawImage(im,0,0);src=c.toDataURL('image/png');}
    const name=`image${++mediaCount}.${extension}`,rid=`rId${rels.length+1}`;files[`ppt/media/${name}`]=dataBytes(src);rels.push([rid,'image',`../media/${name}`]);let crop='';
    if(e.type==='image'){const im=imagePool.get(e.src),iw=im.naturalWidth,ih=im.naturalHeight;if(iw&&ih){if(e.fit==='contain'){const sc=Math.min(e.w/iw,e.h/ih),ww=iw*sc,hh=ih*sc;e={...e,x:e.x+(e.w-ww)/2,y:e.y+(e.h-hh)/2,w:ww,h:hh};}else{const sc=Math.max(e.w/iw,e.h/ih),l=(1-e.w/(iw*sc))/2,t=(1-e.h/(ih*sc))/2;crop=`<a:srcRect l="${Math.round(l*100000)}" r="${Math.round(l*100000)}" t="${Math.round(t*100000)}" b="${Math.round(t*100000)}"/>`;}}}
    content+=`<p:pic><p:nvPicPr><p:cNvPr id="${id++}" name="${xmlEsc(e.name||'Picture')}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}">${(e.opacity??1)<1?`<a:alphaModFix amt="${Math.round(e.opacity*100000)}"/>`:''}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${transformXML(e)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
   }else content+=shapeXML(e,id++,doc.theme);
  }
  const transition=s.transition==='fade'?'<p:transition><p:fade/></p:transition>':s.transition==='push'?'<p:transition><p:push dir="l"/></p:transition>':'';
  part(`ppt/slides/slide${i+1}.xml`,`<p:sld ${ns}><p:cSld name="${xmlEsc(s.name)}"><p:bg><p:bgPr>${solidXML(s.bg,doc.theme)}<a:effectLst/></p:bgPr></p:bg><p:spTree>${groupStart}${content}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transition}</p:sld>`,type('slide'));
  files[`ppt/slides/_rels/slide${i+1}.xml.rels`]=xmlRels(rels);presentationRels.push([`rId${i+3}`,'slide',`slides/slide${i+1}.xml`]);
  const noteElement=makeElement('text',{text:s.notes,x:100,y:100,w:600,h:600,fontSize:16,fill:'@ink'});
  part(`ppt/notesSlides/notesSlide${i+1}.xml`,`<p:notes ${ns}><p:cSld><p:spTree>${groupStart}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>${textBodyXML(noteElement,doc.theme)}</p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,type('notesSlide'));
  files[`ppt/notesSlides/_rels/notesSlide${i+1}.xml.rels`]=xmlRels([['rId1','notesMaster','../notesMasters/notesMaster1.xml'],['rId2','slide',`../slides/slide${i+1}.xml`]]);
 }
 files['ppt/_rels/presentation.xml.rels']=xmlRels(presentationRels);files['_rels/.rels']=xmlRels([['rId1','officeDocument','ppt/presentation.xml'],['rId2','http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties','docProps/core.xml'],['rId3','https://aurelia.slides/document','aurelia/document.json']]);
 part('docProps/core.xml',`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${xmlEsc(doc.title)}</dc:title><dc:creator>Aurelia Slides</dc:creator><dc:description>Created with Aurelia Slides. Editable vector shapes and text.</dc:description></cp:coreProperties>`,'application/vnd.openxmlformats-package.core-properties+xml');
 files['aurelia/document.json']=JSON.stringify(doc);
 files['[Content_Types].xml']=XML_HEAD+`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="json" ContentType="application/json"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/>${parts.map(([p,t])=>`<Override PartName="/${p}" ContentType="${t}"/>`).join('')}</Types>`;
 return new Blob([await zipStore(files).arrayBuffer()],{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'});
}
const childrenNamed=(n,name)=>Array.from(n?.children||[]).filter(c=>c.localName===name);
const firstNamed=(n,name)=>n?.getElementsByTagNameNS('*',name)[0]||null;
const allNamed=(n,name)=>Array.from(n?.getElementsByTagNameNS('*',name)||[]);
function parseXML(bytes){const x=new DOMParser().parseFromString(new TextDecoder().decode(bytes),'application/xml');if(x.getElementsByTagName('parsererror').length)throw new Error('Malformed PowerPoint XML.');return x;}
function resolvePart(base,target){const segments=(base.slice(0,base.lastIndexOf('/')+1)+target).split('/'),result=[];for(const s of segments){if(s==='..')result.pop();else if(s!=='.'&&s)result.push(s);}return result.join('/');}
export async function importPPTX(buffer){
 const files=await unzipSafe(buffer),warnings=new Set();
 if(files.has('aurelia/document.json'))return{doc:validateDocument(JSON.parse(new TextDecoder().decode(files.get('aurelia/document.json')))),warnings:[]};
 if(!files.has('ppt/presentation.xml'))throw new Error('This ZIP does not contain a PowerPoint presentation.');
 const doc=makeDocument(),root=parseXML(files.get('ppt/presentation.xml')),sz=firstNamed(root,'sldSz');doc.width=(+sz?.getAttribute('cx')||12192000)/9525;doc.height=(+sz?.getAttribute('cy')||6858000)/9525;doc.slides=[];
 const relsFor=path=>{const slash=path.lastIndexOf('/'),relPath=path.slice(0,slash+1)+'_rels/'+path.slice(slash+1)+'.rels';if(!files.has(relPath))return new Map();return new Map(allNamed(parseXML(files.get(relPath)),'Relationship').filter(r=>r.getAttribute('TargetMode')!=='External').map(r=>[r.getAttribute('Id'),{path:resolvePart(path,r.getAttribute('Target')),type:r.getAttribute('Type')}]));};
 const presRels=relsFor('ppt/presentation.xml');let theme={dk1:'000000',lt1:'FFFFFF',dk2:'17313A',lt2:'F5F2EC',accent1:'DA735B',accent2:'A5B8AF',tx1:'000000',bg1:'FFFFFF'};
 const themeFile=[...files.keys()].find(k=>/^ppt\/theme\/theme\d+\.xml$/.test(k));if(themeFile){const scheme=firstNamed(parseXML(files.get(themeFile)),'clrScheme');for(const el of scheme?.children||[]){const c=el.children[0];theme[el.localName]=c?.getAttribute('val')?.match(/^[0-9a-f]{6}$/i)?.[0]||c?.getAttribute('lastClr')||'000000';}theme.tx1=theme.dk1;theme.bg1=theme.lt1;theme.tx2=theme.dk2;theme.bg2=theme.lt2;}
 const getColor=(node,defaultValue='#17313A')=>{if(!node)return defaultValue;const fill=childrenNamed(node,'solidFill')[0];if(!fill)return childrenNamed(node,'noFill').length?'none':defaultValue;const c=fill.children[0];if(!c)return defaultValue;return'#'+(c.localName==='srgbClr'?c.getAttribute('val'):theme[c.getAttribute('val')]||'17313A');};
 const transform=node=>{const xf=firstNamed(node,'xfrm'),off=childrenNamed(xf,'off')[0],ext=childrenNamed(xf,'ext')[0];return{x:(+off?.getAttribute('x')||0)/9525,y:(+off?.getAttribute('y')||0)/9525,w:(+ext?.getAttribute('cx')||3048000)/9525,h:(+ext?.getAttribute('cy')||1524000)/9525,rotation:(+xf?.getAttribute('rot')||0)/60000};};
 if(files.has('docProps/core.xml'))doc.title=firstNamed(parseXML(files.get('docProps/core.xml')),'title')?.textContent||'Imported presentation';
 for(const [i,sid]of allNamed(root,'sldId').entries()){
  const rid=sid.getAttributeNS(R_NS,'id'),path=presRels.get(rid)?.path;if(!path||!files.has(path))continue;const xml=parseXML(files.get(path)),rels=relsFor(path),s=makeSlide();s.name=firstNamed(xml,'cSld')?.getAttribute('name')||`Slide ${i+1}`;s.bg=getColor(firstNamed(xml,'bgPr'),'#FFFFFF');
  const readTree=(tree,offset={x:0,y:0,sx:1,sy:1})=>{for(const node of tree?.children||[]){
   const kind=node.localName;if(!['sp','pic','cxnSp','grpSp','graphicFrame'].includes(kind))continue;
   if(kind==='grpSp'){const gp=childrenNamed(node,'grpSpPr')[0],xf=firstNamed(gp,'xfrm'),o=childrenNamed(xf,'off')[0],ex=childrenNamed(xf,'ext')[0],co=childrenNamed(xf,'chOff')[0],ce=childrenNamed(xf,'chExt')[0],sx=(+ex?.getAttribute('cx')||1)/(+ce?.getAttribute('cx')||1),sy=(+ex?.getAttribute('cy')||1)/(+ce?.getAttribute('cy')||1);readTree(node,{x:offset.x+((+o?.getAttribute('x')||0)-(+co?.getAttribute('x')||0)*sx)/9525*offset.sx,y:offset.y+((+o?.getAttribute('y')||0)-(+co?.getAttribute('y')||0)*sy)/9525*offset.sy,sx:offset.sx*sx,sy:offset.sy*sy});if(+xf?.getAttribute('rot'))warnings.add('Rotated PowerPoint groups are imported without their group rotation.');continue;}
   if(kind==='graphicFrame'){warnings.add('Native PowerPoint charts, SmartArt, tables, and embedded objects are not imported.');continue;}
   const sp=childrenNamed(node,'spPr')[0],t=transform(sp),tr={x:offset.x+t.x*offset.sx,y:offset.y+t.y*offset.sy,w:Math.max(1,t.w*offset.sx),h:Math.max(1,t.h*offset.sy),rotation:t.rotation},name=firstNamed(node,'cNvPr')?.getAttribute('name')||'Imported object';
   if(kind==='pic'){const embed=firstNamed(node,'blip')?.getAttributeNS(R_NS,'embed'),target=rels.get(embed)?.path,bytes=files.get(target);if(bytes){const ext=target.split('.').pop().toLowerCase(),mime={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',webp:'image/webp'}[ext];if(mime)s.elements.push(makeElement('image',{...tr,name,src:bytesData(bytes,mime),fit:'cover'}));else warnings.add('Unsupported picture formats (including SVG/EMF/WMF) are skipped.');}if(firstNamed(node,'srcRect'))warnings.add('Picture crop windows are approximated by center-cover.');continue;}
   const geometry=firstNamed(sp,'prstGeom')?.getAttribute('prst')||'rect',type={rect:'rect',roundRect:'roundRect',ellipse:'ellipse',triangle:'triangle',diamond:'diamond',rightArrow:'arrow',star5:'star',line:'line'}[geometry]||'rect',tx=childrenNamed(node,'txBody')[0],fill=getColor(sp,'none'),ln=childrenNamed(sp,'ln')[0];
   if(fill!=='none'||!tx)s.elements.push(makeElement(type,{...tr,name,fill,stroke:getColor(ln,'none'),strokeWidth:(+ln?.getAttribute('w')||0)/9525}));
   if(tx){const paragraphs=childrenNamed(tx,'p'),text=paragraphs.map(p=>Array.from(p.children).map(child=>child.localName==='br'?'\n':allNamed(child,'t').map(t=>t.textContent).join('')).join('')).join('\n');const style=firstNamed(tx,'rPr')||firstNamed(tx,'defRPr')||firstNamed(tx,'endParaRPr'),body=firstNamed(tx,'bodyPr'),ppr=firstNamed(tx,'pPr'),font=firstNamed(style,'latin')?.getAttribute('typeface');s.elements.push(makeElement('text',{...tr,name,text,fontSize:(+style?.getAttribute('sz')||2400)/75*offset.sy,fontFamily:font?.startsWith('+')?'Arial':font||'Arial',bold:style?.getAttribute('b')==='1',italic:style?.getAttribute('i')==='1',underline:style?.getAttribute('u')==='sng',fill:getColor(style,'#17313A'),align:({l:'left',ctr:'center',r:'right'})[ppr?.getAttribute('algn')]||'left',padding:(+body?.getAttribute('lIns')||0)/9525,valign:({t:'top',ctr:'middle',b:'bottom'})[body?.getAttribute('anchor')]||'top'}));if(allNamed(tx,'rPr').length>1)warnings.add('Mixed text runs are normalized to one style per text box.');}
   if(!childrenNamed(sp,'xfrm').length)warnings.add('Layout/master placeholder inheritance is approximated; some imported positions need adjustment.');
  }};readTree(firstNamed(xml,'spTree'));
  const notes=[...rels.values()].find(r=>r.type.endsWith('/notesSlide'));if(notes&&files.has(notes.path)){const n=parseXML(files.get(notes.path));s.notes=allNamed(n,'sp').filter(sp=>firstNamed(sp,'ph')?.getAttribute('type')==='body').map(sp=>allNamed(sp,'p').map(p=>allNamed(p,'t').map(t=>t.textContent).join('')).join('\n')).join('\n');}
  s.transition=firstNamed(xml,'fade')?'fade':firstNamed(xml,'push')?'push':'none';doc.slides.push(s);
 }
 if(!doc.slides.length)throw new Error('No readable slides found.');warnings.add('PPTX import is best-effort: masters, effects, media, animations, and advanced typography may differ.');return{doc:validateDocument(doc),warnings:[...warnings]};
}
