/** Aurelia document/geometry/command kernel. No browser or renderer dependencies. */
export const VERSION = 1;
export const THEMES = {
  studio: { name:'Terracotta', bg:'#F5F2EC', ink:'#17313A', dark:'#17313A', light:'#F5F2EC', accent:'#DA735B', secondary:'#A5B8AF', muted:'#71827F', line:'#DDDCD4' },
  midnight: { name:'Midnight', bg:'#EDF0F7', ink:'#202941', dark:'#202941', light:'#EDF0F7', accent:'#8992DF', secondary:'#B4C9E2', muted:'#7B829A', line:'#D3D7E4' },
  forest: { name:'Botanical', bg:'#F2F3EA', ink:'#203B32', dark:'#203B32', light:'#F2F3EA', accent:'#99AA69', secondary:'#C7D4B8', muted:'#6C8271', line:'#D5DCCD' },
  cobalt: { name:'Blueprint', bg:'#F3F6FF', ink:'#172B51', dark:'#172B51', light:'#F3F6FF', accent:'#446CE9', secondary:'#ADC8EF', muted:'#7183A5', line:'#D7DFF0' },
  rose: { name:'Atelier', bg:'#FBF3F1', ink:'#4C313C', dark:'#4C313C', light:'#FBF3F1', accent:'#CC788D', secondary:'#D9BEA6', muted:'#997B85', line:'#ECD7DD' }
};
export const clone = value => structuredClone(value);
export const uid = (prefix='e') => `${prefix}_${globalThis.crypto?.randomUUID?.() || Math.random().toString(36).slice(2)+Date.now().toString(36)}`;
export const clamp = (n,a,b) => Math.max(a,Math.min(b,n));
export const color = (value,theme='studio') => value?.startsWith('@') ? (THEMES[theme]?.[value.slice(1)] || '#17313A') : (value || '#000000');
export const escapeHTML = s => String(s ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function makeElement(type='rect', props={}) {
  return { id:uid(), type, name:({text:'Text box',rect:'Rectangle',roundRect:'Rounded rectangle',ellipse:'Ellipse',line:'Line',triangle:'Triangle',diamond:'Diamond',arrow:'Arrow',star:'Star',image:'Picture',chart:'Chart',table:'Table'})[type]||type,
    x:100,y:100,w:320,h:160,rotation:0,opacity:1,fill:'@accent',stroke:'none',strokeWidth:0,radius:22,locked:false,hidden:false,groupId:null,animation:'none',
    ...(type==='text'?{text:'Your next great idea',fill:'@ink',fontFamily:'Arial',fontSize:40,bold:false,italic:false,underline:false,align:'left',valign:'top',lineHeight:1.2,padding:4}:{}),...props };
}
export function makeSlide(layout='blank', theme='studio') {
  const s={id:uid('s'),name:layout==='blank'?'Blank slide':'Untitled slide',bg:'@bg',notes:'',transition:'fade',duration:0.45,elements:[]};
  if(layout==='title') s.elements=[makeElement('text',{name:'Title',x:88,y:215,w:1100,h:160,text:'A new perspective.',fontSize:92,bold:true}),makeElement('text',{name:'Subtitle',x:94,y:415,w:1000,h:80,text:'An idea worth sharing.',fontSize:28,fill:'@muted'})];
  if(layout==='content'||layout==='split') {
    s.elements.push(makeElement('text',{name:'Title',x:78,y:65,w:1110,h:100,text:'Make your point.',fontSize:62,bold:true}));
    s.elements.push(makeElement('text',{name:'Content',x:82,y:225,w:layout==='split'?515:1100,h:385,text:'Start with what matters.\n\nConnect the dots.\n\nGive your audience a reason to care.',fontSize:31,fill:'@muted'}));
    if(layout==='split')s.elements.push(makeElement('roundRect',{name:'Visual',x:660,y:222,w:530,h:380,fill:'@secondary',radius:24}));
  }
  if(layout==='quote') s.elements=[makeElement('text',{x:90,y:110,w:180,h:180,text:'“',fontSize:180,fill:'@accent',fontFamily:'Georgia'}),makeElement('text',{name:'Quote',x:120,y:270,w:1000,h:240,text:'Great stories make\ncomplex ideas feel simple.',fontSize:64,fontFamily:'Georgia'}),makeElement('text',{x:125,y:570,w:700,h:50,text:'YOUR NAME  /  YOUR PERSPECTIVE',fontSize:18,fill:'@muted'})];
  return s;
}
export function makeDocument() { return {format:'aurelia',version:VERSION,id:uid('deck'),title:'Untitled presentation',width:1280,height:720,theme:'studio',slides:[makeSlide('title')]}; }
const VALID_TYPES=new Set(['rect','roundRect','ellipse','triangle','diamond','arrow','star','line','text','image','chart','table']);
const VALID_COLORS=/^(?:@[a-z]+|#[0-9a-fA-F]{6}|#[0-9a-fA-F]{8}|none)$/;
export function validateDocument(raw) {
  if(!raw||raw.format!=='aurelia'||raw.version!==VERSION)throw new Error('This is not a supported Aurelia presentation.');
  if(!Array.isArray(raw.slides)||raw.slides.length<1||raw.slides.length>500)throw new Error('A presentation must contain 1–500 slides.');
  const d=makeDocument(); d.id=String(raw.id||d.id);d.title=String(raw.title||'Untitled presentation').slice(0,200);d.width=clamp(Number(raw.width)||1280,320,4096);d.height=clamp(Number(raw.height)||720,240,4096);d.theme=THEMES[raw.theme]?raw.theme:'studio';
  const seen=new Set();
  function safeId(id){if(typeof id!=='string'||seen.has(id))id=uid();seen.add(id);return id;}
  const safeColor=(v,fallback)=>typeof v==='string'&&VALID_COLORS.test(v)?v:fallback;
  d.slides=raw.slides.map((s,index)=>{
    if(!Array.isArray(s.elements)||s.elements.length>5000)throw new Error(`Slide ${index+1} exceeds the 5,000-object limit.`);
    const out={id:safeId(s.id),name:String(s.name||`Slide ${index+1}`).slice(0,200),bg:safeColor(s.bg,'@bg'),notes:String(s.notes||'').slice(0,100000),transition:['none','fade','push','zoom'].includes(s.transition)?s.transition:'none',duration:clamp(+s.duration||0.45,0.1,5),elements:[]};
    out.elements=s.elements.map(v=>{
      if(!VALID_TYPES.has(v.type))throw new Error(`Unsupported object type: ${v.type}`);
      const e=makeElement(v.type);e.id=safeId(v.id);e.name=String(v.name||e.name).slice(0,200);
      for(const p of ['x','y','w','h','rotation','opacity','strokeWidth','radius'])if(Number.isFinite(v[p]))e[p]=clamp(v[p],-100000,100000);
      e.w=clamp(e.w,1,16384);e.h=clamp(e.h,1,16384);e.opacity=clamp(e.opacity,0,1);e.strokeWidth=clamp(e.strokeWidth,0,100);
      e.fill=safeColor(v.fill,e.fill);e.stroke=safeColor(v.stroke,'none');e.locked=!!v.locked;e.hidden=!!v.hidden;e.groupId=typeof v.groupId==='string'?v.groupId.slice(0,100):null;e.animation=['none','fade','rise','zoom'].includes(v.animation)?v.animation:'none';
      if(e.type==='text') {e.text=String(v.text||'').slice(0,100000);e.fontFamily=String(v.fontFamily||'Arial').replace(/[^a-zA-Z0-9 ,\-]/g,'').slice(0,100);e.fontSize=clamp(+v.fontSize||40,4,600);e.bold=!!v.bold;e.italic=!!v.italic;e.underline=!!v.underline;e.align=['left','center','right'].includes(v.align)?v.align:'left';e.valign=['top','middle','bottom'].includes(v.valign)?v.valign:'top';e.lineHeight=clamp(+v.lineHeight||1.2,0.8,3);e.padding=clamp(+v.padding||0,0,200);}
      if(e.type==='image'){if(typeof v.src!=='string'||!/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/=]+$/.test(v.src)||v.src.length>30000000)throw new Error('Pictures must be embedded PNG, JPEG or WebP files, at most 22 MB each.');e.src=v.src;e.fit=v.fit==='contain'?'contain':'cover';}
      if(e.type==='chart'){e.chartType=['bar','line','donut'].includes(v.chartType)?v.chartType:'bar';e.labels=(v.labels||['A','B','C']).slice(0,30).map(x=>String(x).slice(0,80));e.values=(v.values||[20,40,70]).slice(0,e.labels.length).map(x=>clamp(Number(x)||0,0,1e12));while(e.values.length<e.labels.length)e.values.push(0);e.showValues=v.showValues!==false;}
      if(e.type==='table'){if(!Array.isArray(v.cells)||!v.cells.length)throw new Error('Invalid table.');e.cells=v.cells.slice(0,30).map(r=>Array.isArray(r)?r.slice(0,12).map(c=>String(c).slice(0,1000)):['']);e.fontSize=clamp(+v.fontSize||22,8,100);}
      return e;
    });return out;
  });return d;
}
export function localPoint(e,p) {
  const a=-(e.rotation||0)*Math.PI/180,c=Math.cos(a),s=Math.sin(a),dx=p.x-e.x-e.w/2,dy=p.y-e.y-e.h/2;
  return {x:dx*c-dy*s+e.w/2,y:dx*s+dy*c+e.h/2};
}
export function hitElement(e,p,tolerance=4) {
  if(e.hidden)return false;
  const q=localPoint(e,p),cx=q.x-e.w/2,cy=q.y-e.h/2;
  if(e.type==='ellipse')return (cx/(e.w/2+tolerance))**2+(cy/(e.h/2+tolerance))**2<=1;
  if(e.type==='diamond')return Math.abs(cx)/(e.w/2+tolerance)+Math.abs(cy)/(e.h/2+tolerance)<=1;
  if(e.type==='line')return Math.abs(q.y-e.h/2)<=Math.max(tolerance,e.h/2)&&q.x>=-tolerance&&q.x<=e.w+tolerance;
  if(e.type==='triangle'){const allowance=(q.y/e.h)*(e.w/2)+tolerance;return q.y>=-tolerance&&q.y<=e.h+tolerance&&Math.abs(cx)<=allowance;}
  return q.x>=-tolerance&&q.y>=-tolerance&&q.x<=e.w+tolerance&&q.y<=e.h+tolerance;
}
export function elementBounds(e) {
  const a=(e.rotation||0)*Math.PI/180,c=Math.abs(Math.cos(a)),s=Math.abs(Math.sin(a)),w=e.w*c+e.h*s,h=e.w*s+e.h*c;
  return {x:e.x+(e.w-w)/2,y:e.y+(e.h-h)/2,w,h};
}
export function unionBounds(elements) {
  if(!elements.length)return null; const bs=elements.map(elementBounds),x=Math.min(...bs.map(b=>b.x)),y=Math.min(...bs.map(b=>b.y));
  return {x,y,w:Math.max(...bs.map(b=>b.x+b.w))-x,h:Math.max(...bs.map(b=>b.y+b.h))-y};
}
export function rectIntersects(a,b) {return a.x<=b.x+b.w&&a.x+a.w>=b.x&&a.y<=b.y+b.h&&a.y+a.h>=b.y;}
export function snapMove(bounds,others,width,height,threshold=7) {
  const xs=[0,width/2,width],ys=[0,height/2,height];for(const e of others){const b=elementBounds(e);xs.push(b.x,b.x+b.w/2,b.x+b.w);ys.push(b.y,b.y+b.h/2,b.y+b.h);}
  function match(points,targets){let best={delta:0,line:null,d:threshold};for(const p of points)for(const t of targets){const d=Math.abs(t-p);if(d<best.d)best={delta:t-p,line:t,d};}return best;}
  const x=match([bounds.x,bounds.x+bounds.w/2,bounds.x+bounds.w],xs),y=match([bounds.y,bounds.y+bounds.h/2,bounds.y+bounds.h],ys);
  return {dx:x.delta,dy:y.delta,lines:[...(x.line===null?[]:[{axis:'x',value:x.line}]),...(y.line===null?[]:[{axis:'y',value:y.line}])]};
}
export class Store {
  constructor(doc=makeDocument()){this.doc=doc;this.active=doc.slides[0].id;this.selection=new Set();this.past=[];this.future=[];this.listeners=new Set();this.pending=null;this.revision=0;this.maxHistory=100;this.maxHistoryBytes=40*1024*1024;}
  get slide(){return this.doc.slides.find(s=>s.id===this.active)||this.doc.slides[0];}
  get selected(){return this.slide.elements.filter(e=>this.selection.has(e.id));}
  onChange(fn){this.listeners.add(fn);return()=>this.listeners.delete(fn);}
  emit(kind='change'){for(const fn of this.listeners)fn(kind);}
  select(ids){this.selection=new Set(ids);this.emit('selection');}
  activate(id){if(!this.doc.slides.some(s=>s.id===id))return;this.active=id;this.selection.clear();this.emit('active');}
  snapshot(){return JSON.stringify({doc:this.doc,active:this.active});}
  begin(label='Edit'){if(!this.pending)this.pending={label,before:this.snapshot()};}
  commit(){if(!this.pending)return false;const {label,before}=this.pending;this.pending=null;const after=this.snapshot();if(before===after)return false;this.past.push({label,before,after});this.future=[];let size=this.past.reduce((n,x)=>n+x.before.length+x.after.length,0);while(this.past.length>1&&(this.past.length>this.maxHistory||size>this.maxHistoryBytes)){const x=this.past.shift();size-=x.before.length+x.after.length;}this.revision++;this.emit('commit');return true;}
  cancel(){if(!this.pending)return;const snap=this.pending.before;this.pending=null;this.restore(snap);}
  transaction(label,fn){this.begin(label);try{fn(this.doc,this.slide);this.commit();}catch(e){this.cancel();throw e;}}
  restore(snapshot){const state=JSON.parse(snapshot);this.doc=state.doc;this.active=state.active;this.selection=new Set([...this.selection].filter(id=>this.slide.elements.some(e=>e.id===id)));this.revision++;this.emit('restore');}
  undo(){if(this.pending)this.commit();const cmd=this.past.pop();if(!cmd)return;this.future.push(cmd);this.restore(cmd.before);}
  redo(){const cmd=this.future.pop();if(!cmd)return;this.past.push(cmd);this.restore(cmd.after);}
  replace(doc){this.transaction('Open presentation',()=>{this.doc=doc;this.active=doc.slides[0].id;this.selection.clear();});}
  add(element){this.transaction(`Insert ${element.type}`,()=>this.slide.elements.push(element));this.select([element.id]);return element;}
  updateSelected(props,label='Format objects'){this.transaction(label,()=>{for(const e of this.selected)if(!e.locked)Object.assign(e,props);});}
  deleteSelected(){this.transaction('Delete objects',()=>{this.slide.elements=this.slide.elements.filter(e=>!this.selection.has(e.id)||e.locked);});this.select([]);}
  duplicateSelected(){const groups=new Map(),els=clone(this.selected).filter(e=>!e.locked).map(e=>{e.id=uid();e.x+=24;e.y+=24;if(e.groupId){if(!groups.has(e.groupId))groups.set(e.groupId,uid('g'));e.groupId=groups.get(e.groupId);}return e;});this.transaction('Duplicate objects',()=>this.slide.elements.push(...els));this.select(els.map(e=>e.id));}
  addSlide(layout='blank'){const s=makeSlide(layout,this.doc.theme);this.transaction('New slide',()=>{const i=this.doc.slides.findIndex(s=>s.id===this.active);this.doc.slides.splice(i+1,0,s);this.active=s.id;});this.select([]);return s;}
  duplicateSlide(){const s=clone(this.slide);s.id=uid('s');s.name+=' copy';const groups=new Map();for(const e of s.elements){e.id=uid();if(e.groupId){if(!groups.has(e.groupId))groups.set(e.groupId,uid('g'));e.groupId=groups.get(e.groupId);}}this.transaction('Duplicate slide',()=>{const i=this.doc.slides.findIndex(x=>x.id===this.active);this.doc.slides.splice(i+1,0,s);this.active=s.id;});this.select([]);}
  deleteSlide(){if(this.doc.slides.length===1)return false;this.transaction('Delete slide',()=>{const i=this.doc.slides.findIndex(s=>s.id===this.active);this.doc.slides.splice(i,1);this.active=this.doc.slides[Math.min(i,this.doc.slides.length-1)].id;});this.select([]);return true;}
  reorderSlide(from,to){this.transaction('Reorder slides',()=>{const [slide]=this.doc.slides.splice(from,1);this.doc.slides.splice(to,0,slide);});}
  align(mode){const els=this.selected.filter(e=>!e.locked);if(!els.length)return;const b=els.length===1?{x:0,y:0,w:this.doc.width,h:this.doc.height}:unionBounds(els);this.transaction('Align objects',()=>{for(const e of els){const r=elementBounds(e);if(mode==='left')e.x+=b.x-r.x;if(mode==='center')e.x+=b.x+b.w/2-r.x-r.w/2;if(mode==='right')e.x+=b.x+b.w-r.x-r.w;if(mode==='top')e.y+=b.y-r.y;if(mode==='middle')e.y+=b.y+b.h/2-r.y-r.h/2;if(mode==='bottom')e.y+=b.y+b.h-r.y-r.h;}});}
  distribute(axis){const es=this.selected.filter(e=>!e.locked).sort((a,b)=>elementBounds(a)[axis]-elementBounds(b)[axis]);if(es.length<3)return;const k=axis==='x'?'w':'h',bs=es.map(elementBounds),total=bs.reduce((n,b)=>n+b[k],0),gap=(bs.at(-1)[axis]+bs.at(-1)[k]-bs[0][axis]-total)/(es.length-1);this.transaction('Distribute objects',()=>{let p=bs[0][axis];es.forEach((e,i)=>{e[axis]+=p-bs[i][axis];p+=bs[i][k]+gap;});});}
  arrange(mode){this.transaction('Arrange objects',()=>{const list=this.slide.elements,chosen=list.filter(e=>this.selection.has(e.id));if(mode==='front')this.slide.elements=[...list.filter(e=>!this.selection.has(e.id)),...chosen];else if(mode==='back')this.slide.elements=[...chosen,...list.filter(e=>!this.selection.has(e.id))];else if(mode==='forward'){for(let i=list.length-2;i>=0;i--)if(this.selection.has(list[i].id)&&!this.selection.has(list[i+1].id))[list[i],list[i+1]]=[list[i+1],list[i]];}else{for(let i=1;i<list.length;i++)if(this.selection.has(list[i].id)&&!this.selection.has(list[i-1].id))[list[i],list[i-1]]=[list[i-1],list[i]];}});}
  group(){if(this.selected.length<2)return;this.updateSelected({groupId:uid('g')},'Group objects');}
  ungroup(){this.updateSelected({groupId:null},'Ungroup objects');}
}
