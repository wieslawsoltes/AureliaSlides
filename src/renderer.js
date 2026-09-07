import {color,clamp} from './core.js';
/** Ordered retained-scene renderer: analytic GPU shapes + cached, shaped text/image tiles. */
export const SHAPE_IDS={rect:0,roundRect:1,ellipse:2,triangle:3,diamond:4,arrow:5,star:6,line:7};
export const GPU_SHADER=`
struct View { size: vec4f, }
struct Instance { rect:vec4f, transform:vec4f, fill:vec4f, stroke:vec4f, params:vec4f, }
@group(0) @binding(0) var<uniform> view:View;
@group(0) @binding(1) var<storage,read> items:array<Instance>;
struct Vertex { @builtin(position) position:vec4f, @location(0) local:vec2f, @location(1) @interpolate(flat) index:u32, }
fn corner(i:u32)->vec2f { let p=array<vec2f,6>(vec2f(0,0),vec2f(1,0),vec2f(0,1),vec2f(0,1),vec2f(1,0),vec2f(1,1));return p[i]; }
fn project(p:vec2f)->vec4f{return vec4f(p.x/view.size.x*2.-1.,1.-p.y/view.size.y*2.,0.,1.);}
@vertex fn shapeVertex(@builtin(vertex_index) vi:u32,@builtin(instance_index) ii:u32)->Vertex {
 let e=items[ii];let p=(corner(vi)-.5)*(e.rect.zw+vec2f(4));let r=vec2f(p.x*e.transform.x-p.y*e.transform.y,p.x*e.transform.y+p.y*e.transform.x);var o:Vertex;o.position=project(e.rect.xy+e.rect.zw*.5+r);o.local=p;o.index=ii;return o;
}
fn polygon(p:vec2f,points:array<vec2f,10>,count:u32)->f32 {
 var d=dot(p-points[0],p-points[0]);var sign=1.;var j=count-1u;
 for(var i=0u;i<count;i++) {let a=points[i];let b=points[j];let e=b-a;let v=p-a;let q=v-e*clamp(dot(v,e)/max(dot(e,e),.0001),0.,1.);d=min(d,dot(q,q));let c=vec3<bool>(p.y>=a.y,p.y<b.y,e.x*v.y>e.y*v.x);if(all(c)||all(!c)){sign=-sign;}j=i;}
 return sign*sqrt(d);
}
fn distance(p:vec2f,h:vec2f,kind:u32,r:f32)->f32 {
 if(kind==2u){return (length(p/max(h,vec2f(.01)))-1.)*min(h.x,h.y);}
 var pts:array<vec2f,10>;
 if(kind==3u){pts[0]=vec2f(0,-h.y);pts[1]=h;pts[2]=vec2f(-h.x,h.y);return polygon(p,pts,3u);}
 if(kind==4u){pts[0]=vec2f(0,-h.y);pts[1]=vec2f(h.x,0);pts[2]=vec2f(0,h.y);pts[3]=vec2f(-h.x,0);return polygon(p,pts,4u);}
 if(kind==5u){pts[0]=vec2f(-h.x,-h.y*.36);pts[1]=vec2f(h.x*.2,-h.y*.36);pts[2]=vec2f(h.x*.2,-h.y);pts[3]=vec2f(h.x,0);pts[4]=vec2f(h.x*.2,h.y);pts[5]=vec2f(h.x*.2,h.y*.36);pts[6]=vec2f(-h.x,h.y*.36);return polygon(p,pts,7u);}
 if(kind==6u){for(var i=0u;i<10u;i++){let a=f32(i)*.62831853-1.57079633;let rad=select(.44,1.,i%2u==0u);pts[i]=vec2f(cos(a),sin(a))*h*rad;}return polygon(p,pts,10u);}
 let radius=select(0.,min(r,min(h.x,h.y)),kind==1u);let q=abs(p)-h+radius;return length(max(q,vec2f(0)))+min(max(q.x,q.y),0.)-radius;
}
@fragment fn shapeFragment(v:Vertex)->@location(0) vec4f {
 let e=items[v.index];let d=distance(v.local,e.rect.zw*.5,u32(e.transform.w),e.params.x);let aa=max(fwidth(d),.5);let coverage=1.-smoothstep(-aa*.5,aa*.5,d);var c=e.fill;
 if(e.params.y>0.){let border=smoothstep(-e.params.y-aa*.5,-e.params.y+aa*.5,d);c=mix(c,e.stroke,border);}
 let alpha=c.a*coverage*e.transform.z;return vec4f(c.rgb*alpha,alpha);
}
@group(1) @binding(0) var textureSampler:sampler;
@group(1) @binding(1) var tile:texture_2d<f32>;
@vertex fn textureVertex(@builtin(vertex_index) vi:u32,@builtin(instance_index) ii:u32)->Vertex {
 let e=items[ii];let uv=corner(vi);let p=(uv-.5)*e.rect.zw;let r=vec2f(p.x*e.transform.x-p.y*e.transform.y,p.x*e.transform.y+p.y*e.transform.x);var o:Vertex;o.position=project(e.rect.xy+e.rect.zw*.5+r);o.local=uv;o.index=ii;return o;
}
@fragment fn textureFragment(v:Vertex)->@location(0) vec4f {return textureSample(tile,textureSampler,v.local)*items[v.index].transform.z;}
`;
export function rgba(value,theme){const c=color(value,theme);if(c==='none')return[0,0,0,0];const h=c.slice(1);return[parseInt(h.slice(0,2),16)/255,parseInt(h.slice(2,4),16)/255,parseInt(h.slice(4,6),16)/255,h.length===8?parseInt(h.slice(6,8),16)/255:1];}
export function fontCSS(e){return `${e.italic?'italic ':''}${e.bold?'700':'400'} ${e.fontSize||24}px "${e.fontFamily||'Arial'}", sans-serif`;}
export function wrapText(ctx,text,maxWidth){
  const lines=[];maxWidth=Math.max(1,maxWidth);
  for(const para of String(text).replace(/\r\n?/g,'\n').split('\n')){
    if(!para){lines.push('');continue;}let line='';
    for(const token of para.match(/\S+\s*|\s+/gu)||[]){
      if(ctx.measureText(line+token.trimEnd()).width<=maxWidth){line+=token;continue;}
      if(line.trim()){lines.push(line.trimEnd());line='';}
      if(ctx.measureText(token.trimEnd()).width<=maxWidth){line=token.trimStart();continue;}
      const chars=globalThis.Intl?.Segmenter?[...new Intl.Segmenter(undefined,{granularity:'grapheme'}).segment(token)].map(x=>x.segment):Array.from(token);
      for(const ch of chars){if(line&&ctx.measureText(line+ch).width>maxWidth){lines.push(line);line='';}line+=ch;}
    }lines.push(line.trimEnd());
  }return lines;
}
export function drawText2D(ctx,e,theme){
  ctx.font=fontCSS(e);ctx.textBaseline='alphabetic';ctx.fillStyle=color(e.fill,theme);ctx.textAlign=e.align||'left';
  const pad=e.padding??4,fs=e.fontSize||24,lh=fs*(e.lineHeight||1.2),lines=wrapText(ctx,e.text,e.w-2*pad),height=lines.length*lh;
  let top=pad;if(e.valign==='middle')top=Math.max(pad,(e.h-height)/2);if(e.valign==='bottom')top=Math.max(pad,e.h-height-pad);
  const x=e.align==='center'?e.w/2:e.align==='right'?e.w-pad:pad;
  for(let i=0;i<lines.length;i++){const y=top+i*lh+fs*.85;if(y-fs>e.h)break;ctx.fillText(lines[i],x,y);if(e.underline){const w=ctx.measureText(lines[i]).width,xx=x-(e.align==='center'?w/2:e.align==='right'?w:0);ctx.fillRect(xx,y+fs*.12,w,Math.max(1,fs/20));}}
}
export function primitivePath(ctx,e){
  const w=e.w,h=e.h;ctx.beginPath();
  if(e.type==='ellipse'){ctx.ellipse(w/2,h/2,w/2,h/2,0,0,Math.PI*2);return;}
  if(e.type==='roundRect'){ctx.roundRect(0,0,w,h,Math.min(e.radius||0,w/2,h/2));return;}
  let p;
  if(e.type==='triangle')p=[[w/2,0],[w,h],[0,h]];
  else if(e.type==='diamond')p=[[w/2,0],[w,h/2],[w/2,h],[0,h/2]];
  else if(e.type==='arrow')p=[[0,h*.32],[w*.6,h*.32],[w*.6,0],[w,h/2],[w*.6,h],[w*.6,h*.68],[0,h*.68]];
  else if(e.type==='star')p=Array.from({length:10},(_,i)=>{const a=i*Math.PI/5-Math.PI/2,r=i%2?.44:1;return[w/2+Math.cos(a)*w/2*r,h/2+Math.sin(a)*h/2*r];});
  else{ctx.rect(0,0,w,h);return;}
  p.forEach(([x,y],i)=>i?ctx.lineTo(x,y):ctx.moveTo(x,y));ctx.closePath();
}
function childElement(parent,local,id){
  const a=(parent.rotation||0)*Math.PI/180,c=Math.cos(a),s=Math.sin(a),dx=local.x+local.w/2-parent.w/2,dy=local.y+local.h/2-parent.h/2;
  return{opacity:1,stroke:'none',strokeWidth:0,rotation:0,...local,id:parent.id+':'+id,x:parent.x+parent.w/2+dx*c-dy*s-local.w/2,y:parent.y+parent.h/2+dx*s+dy*c-local.h/2,rotation:(parent.rotation||0)+(local.rotation||0),opacity:(parent.opacity??1)*(local.opacity??1)};
}
export function expandElement(e){
  const out=[];const add=(v)=>out.push(childElement(e,v,out.length));
  const txt=(text,x,y,w,h,extra={})=>add({type:'text',text:String(text),x,y,w,h,fontFamily:'Arial',fontSize:17,fill:'@muted',padding:0,align:'left',lineHeight:1.2,...extra});
  if(e.type==='table'){
    const rows=e.cells||[['A','B'],['C','D']],cols=Math.max(1,...rows.map(r=>r.length)),cw=e.w/cols,rh=e.h/rows.length;
    rows.forEach((row,r)=>{for(let c=0;c<cols;c++){add({type:'rect',x:c*cw,y:r*rh,w:cw,h:rh,fill:r===0?(e.fill||'@dark'):(r%2?'@bg':'@line'),stroke:'@light',strokeWidth:1});txt(row[c]??'',c*cw+14,r*rh+8,cw-28,rh-16,{fontSize:e.fontSize||22,fill:r===0?'@light':'@ink',bold:r===0,valign:'middle'});}});return out;
  }
  if(e.type!=='chart'||e.chartType==='donut')return[e];
  const labels=e.labels||['A','B','C'],values=e.values||[30,50,70],mx=Math.max(1,...values),max=Math.ceil(mx/5)*5,left=54,top=24,right=20,bottom=48,W=Math.max(1,e.w-left-right),H=Math.max(1,e.h-top-bottom),bw=W/Math.max(1,labels.length);
  for(let i=0;i<=4;i++){const y=top+H-H*i/4;add({type:'line',x:left,y,w:W,h:1,fill:'@line'});txt(Math.round(max*i/4),0,y-10,43,22,{fontSize:14,align:'right'});}
  values.forEach((v,i)=>{
    const x=left+i*bw,y=top+H-H*v/max;
    if(e.chartType==='bar')add({type:'roundRect',x:x+bw*.18,y,w:bw*.64,h:Math.max(1,top+H-y),fill:i===values.length-1?(e.fill||'@accent'):'@secondary',radius:5});
    else{const cx=x+bw/2;if(i>0){const px=cx-bw,py=top+H-H*values[i-1]/max,dx=cx-px,dy=y-py,len=Math.hypot(dx,dy);add({type:'line',x:(px+cx)/2-len/2,y:(py+y)/2-2,w:len,h:4,rotation:Math.atan2(dy,dx)*180/Math.PI,fill:e.fill||'@accent'});}add({type:'ellipse',x:cx-5,y:y-5,w:10,h:10,fill:e.fill||'@accent'});}
    if(e.showValues!==false)txt(v,x,y-26,bw,23,{fontSize:16,align:'center',fill:'@ink',bold:true});
    txt(labels[i]||'',x,top+H+15,bw,30,{fontSize:15,align:'center'});
  });return out;
}
export function flattenScene(elements){return elements.filter(e=>!e.hidden).flatMap(expandElement);}
export class ImagePool{
  constructor(onLoad=()=>{}){this.items=new Map();this.onLoad=onLoad;}
  get(src){if(this.items.has(src))return this.items.get(src);const image=new Image();image.onload=()=>this.onLoad();image.onerror=()=>this.onLoad();image.src=src;this.items.set(src,image);return image;}
  async ready(elements){await Promise.all(elements.filter(e=>e.type==='image').map(e=>{const im=this.get(e.src);return im.decode().catch(()=>{});}));}
  retain(sources){for(const src of this.items.keys())if(!sources.has(src))this.items.delete(src);}
}
export function drawDonut(ctx,e,theme){
  const values=e.values||[],total=values.reduce((a,b)=>a+b,0)||1,cx=e.w*.35,cy=e.h/2,r=Math.min(e.w*.28,e.h*.4),colors=[e.fill||'@accent','@secondary','@dark','@muted','@line'];let a=-Math.PI/2;
  values.forEach((v,i)=>{const end=a+v/total*Math.PI*2;ctx.beginPath();ctx.arc(cx,cy,r,a,end);ctx.arc(cx,cy,r*.63,end,a,true);ctx.closePath();ctx.fillStyle=color(colors[i%colors.length],theme);ctx.fill();a=end;ctx.fillRect(e.w*.72,35+i*37,12,12);ctx.font='16px Arial';ctx.fillStyle=color('@ink',theme);ctx.fillText(`${e.labels?.[i]||''}  ${v}`,e.w*.72+22,46+i*37);});ctx.textAlign='center';ctx.font='bold 42px Arial';ctx.fillStyle=color('@ink',theme);ctx.fillText(values.reduce((a,b)=>a+b,0),cx,cy+10);ctx.font='14px Arial';ctx.fillStyle=color('@muted',theme);ctx.fillText('TOTAL',cx,cy+36);
}
export function drawLeaf2D(ctx,e,theme,images){
  if(e.hidden)return;ctx.save();ctx.globalAlpha=e.opacity??1;ctx.translate(e.x+e.w/2,e.y+e.h/2);ctx.rotate((e.rotation||0)*Math.PI/180);ctx.translate(-e.w/2,-e.h/2);
  if(e.type==='text'||e.type==='image'||e.type==='chart'){
    ctx.beginPath();ctx.rect(0,0,e.w,e.h);ctx.clip();
    if(e.type==='text')drawText2D(ctx,e,theme);
    if(e.type==='chart')drawDonut(ctx,e,theme);
    if(e.type==='image'){
      const image=images.get(e.src);if(image.complete&&image.naturalWidth){const s=e.fit==='contain'?Math.min(e.w/image.naturalWidth,e.h/image.naturalHeight):Math.max(e.w/image.naturalWidth,e.h/image.naturalHeight),w=image.naturalWidth*s,h=image.naturalHeight*s;ctx.drawImage(image,(e.w-w)/2,(e.h-h)/2,w,h);}else{ctx.fillStyle='#e3e6e7';ctx.fillRect(0,0,e.w,e.h);}
    }
  }else{primitivePath(ctx,e);if(e.fill!=='none'){ctx.fillStyle=color(e.fill,theme);ctx.fill();}if(e.strokeWidth>0&&e.stroke!=='none'){ctx.clip();ctx.lineWidth=e.strokeWidth*2;ctx.strokeStyle=color(e.stroke,theme);ctx.stroke();}}
  ctx.restore();
}
export function paintSlide2D(canvas,slide,doc,images,{scale=1,ignoreId=null}={}){
  const ctx=canvas.getContext('2d');ctx.save();ctx.setTransform(scale,0,0,scale,0,0);ctx.clearRect(0,0,doc.width,doc.height);ctx.fillStyle=color(slide.bg,doc.theme);ctx.fillRect(0,0,doc.width,doc.height);
  for(const e of flattenScene(slide.elements))if(e.id!==ignoreId)drawLeaf2D(ctx,e,doc.theme,images);ctx.restore();
}
function tileSignature(e,theme,resolution){
  if(e.type==='image')return `${e.id}|${e.w}|${e.h}|${e.fit}|${e.src.length}|${resolution}`;
  return JSON.stringify([e.type,e.text,e.w,e.h,e.fill,e.fontSize,e.fontFamily,e.bold,e.italic,e.underline,e.align,e.valign,e.lineHeight,e.padding,e.chartType,e.labels,e.values,theme,resolution]);
}
export class SceneRenderer{
  constructor(canvas,onStatus=()=>{},onInvalidate=()=>{}){
    this.canvas=canvas;this.onStatus=onStatus;this.onInvalidate=onInvalidate;this.mode='initializing';this.device=null;this.cache=new Map();this.cacheBytes=0;this.budget=96*1024*1024;this.frame=0;this.capacity=0;this.images=new ImagePool(()=>{this.invalidateTiles();this.onInvalidate();});this.lastStats={objects:0,draws:0,ms:0,uploads:0};this.disposed=false;
  }
  async init(){
    try{
      if(new URLSearchParams(location.search).get('renderer')==='canvas')throw new Error('Canvas 2D selected');
      if(!navigator.gpu)throw new Error('WebGPU is unavailable');const adapter=await navigator.gpu.requestAdapter({powerPreference:'high-performance'});if(!adapter)throw new Error('No WebGPU adapter');
      const device=await adapter.requestDevice();this.device=device;this.context=this.canvas.getContext('webgpu');if(!this.context)throw new Error('Cannot initialize GPU canvas');
      this.format=navigator.gpu.getPreferredCanvasFormat();this.context.configure({device,format:this.format,alphaMode:'premultiplied'});
      const shader=device.createShaderModule({label:'Aurelia analytic geometry + texture compositor',code:GPU_SHADER});const info=await shader.getCompilationInfo();const errs=info.messages.filter(m=>m.type==='error');if(errs.length)throw new Error(errs.map(m=>m.message).join('\n'));
      this.sceneLayout=device.createBindGroupLayout({entries:[{binding:0,visibility:GPUShaderStage.VERTEX,buffer:{type:'uniform'}},{binding:1,visibility:GPUShaderStage.VERTEX|GPUShaderStage.FRAGMENT,buffer:{type:'read-only-storage'}}]});
      this.tileLayout=device.createBindGroupLayout({entries:[{binding:0,visibility:GPUShaderStage.FRAGMENT,sampler:{type:'filtering'}},{binding:1,visibility:GPUShaderStage.FRAGMENT,texture:{sampleType:'float'}}]});
      const blend={color:{srcFactor:'one',dstFactor:'one-minus-src-alpha',operation:'add'},alpha:{srcFactor:'one',dstFactor:'one-minus-src-alpha',operation:'add'}};
      this.shapePipeline=await device.createRenderPipelineAsync({label:'Analytic instanced shapes',layout:device.createPipelineLayout({bindGroupLayouts:[this.sceneLayout]}),vertex:{module:shader,entryPoint:'shapeVertex'},fragment:{module:shader,entryPoint:'shapeFragment',targets:[{format:this.format,blend}]},primitive:{topology:'triangle-list'}});
      this.texturePipeline=await device.createRenderPipelineAsync({label:'Cached text and image quads',layout:device.createPipelineLayout({bindGroupLayouts:[this.sceneLayout,this.tileLayout]}),vertex:{module:shader,entryPoint:'textureVertex'},fragment:{module:shader,entryPoint:'textureFragment',targets:[{format:this.format,blend}]},primitive:{topology:'triangle-list'}});
      this.uniform=device.createBuffer({size:16,usage:GPUBufferUsage.UNIFORM|GPUBufferUsage.COPY_DST});this.sampler=device.createSampler({magFilter:'linear',minFilter:'linear'});this.ensureCapacity(128);
      device.lost.then(info=>{if(!this.disposed){console.warn('Aurelia GPU device lost:',info.message);this.fallback('GPU device lost — safe rendering enabled');}});device.addEventListener('uncapturederror',event=>{console.error('WebGPU validation:',event.error.message);if(!this.disposed)this.fallback(event.error.message);});
      this.mode='webgpu';this.adapterInfo=adapter.info?.description||adapter.info?.architecture||'WebGPU device';this.onStatus('WebGPU',this.adapterInfo);
    }catch(error){this.fallback(error.message);}
    return this.mode;
  }
  fallback(reason){
    if(this.disposed)return;this.invalidateTiles();
    if(this.context){this.context.unconfigure();const old=this.canvas,replacement=old.cloneNode(false);old.replaceWith(replacement);this.canvas=replacement;this.context=null;}
    this.mode='canvas';this.onStatus('Canvas 2D',reason);this.onInvalidate();
  }
  ensureCapacity(count){
    if(count<=this.capacity)return;this.capacity=2**Math.ceil(Math.log2(Math.max(128,count)));this.instances?.destroy();this.instances=this.device.createBuffer({label:'Growable retained scene instances',size:this.capacity*80,usage:GPUBufferUsage.STORAGE|GPUBufferUsage.COPY_DST});this.data=new Float32Array(this.capacity*20);
    this.sceneGroup=this.device.createBindGroup({layout:this.sceneLayout,entries:[{binding:0,resource:{buffer:this.uniform}},{binding:1,resource:{buffer:this.instances}}]});
  }
  invalidateTiles(){for(const v of this.cache.values())v.texture.destroy();this.cache.clear();this.cacheBytes=0;}
  tile(e,theme,resolution){
    const maxSize=Math.min(4096,this.device.limits.maxTextureDimension2D);resolution=Math.min(resolution,maxSize/e.w,maxSize/e.h);resolution=Math.max(.03,resolution);const key=e.id,signature=tileSignature(e,theme,resolution);let v=this.cache.get(key);
    if(v&&v.signature===signature){v.used=this.frame;return v;}
    if(v){v.texture.destroy();this.cacheBytes-=v.bytes;this.cache.delete(key);}
    const c=document.createElement('canvas');c.width=Math.max(1,Math.ceil(e.w*resolution));c.height=Math.max(1,Math.ceil(e.h*resolution));const ctx=c.getContext('2d');ctx.scale(c.width/e.w,c.height/e.h);drawLeaf2D(ctx,{...e,x:0,y:0,rotation:0,opacity:1},theme,this.images);
    const texture=this.device.createTexture({label:`Tile: ${e.type}`,size:[c.width,c.height],format:'rgba8unorm',usage:GPUTextureUsage.TEXTURE_BINDING|GPUTextureUsage.COPY_DST|GPUTextureUsage.RENDER_ATTACHMENT});this.device.queue.copyExternalImageToTexture({source:c},{texture,premultipliedAlpha:true},[c.width,c.height]);
    const bindGroup=this.device.createBindGroup({layout:this.tileLayout,entries:[{binding:0,resource:this.sampler},{binding:1,resource:texture.createView()}]});v={texture,bindGroup,signature,used:this.frame,bytes:c.width*c.height*4};this.cache.set(key,v);this.cacheBytes+=v.bytes;this.lastStats.uploads++;return v;
  }
  render(slide,doc,{cssWidth=1280,cssHeight=720,dpr=window.devicePixelRatio||1,ignoreId=null}={}){
    if(this.mode==='initializing'||this.disposed)return;const start=performance.now();this.frame++;this.lastStats.uploads=0;
    const width=Math.max(1,Math.round(cssWidth*dpr)),height=Math.max(1,Math.round(cssHeight*dpr));if(this.canvas.width!==width||this.canvas.height!==height){this.canvas.width=width;this.canvas.height=height;}
    if(this.mode==='canvas'){paintSlide2D(this.canvas,slide,doc,this.images,{scale:width/doc.width,ignoreId});this.lastStats={...this.lastStats,ms:performance.now()-start,objects:slide.elements.length,draws:slide.elements.length};return;}
    const elements=flattenScene(slide.elements).filter(e=>e.id!==ignoreId);this.ensureCapacity(elements.length||1);const data=this.data;
    elements.forEach((e,i)=>{const n=i*20,a=(e.rotation||0)*Math.PI/180;data.set([e.x,e.y,e.w,e.h,Math.cos(a),Math.sin(a),e.opacity??1,SHAPE_IDS[e.type]??0,...rgba(e.fill,doc.theme),...rgba(e.stroke||'none',doc.theme),e.radius||0,e.strokeWidth||0,0,0],n);});
    this.device.queue.writeBuffer(this.uniform,0,new Float32Array([doc.width,doc.height,width,height]));if(elements.length)this.device.queue.writeBuffer(this.instances,0,data.buffer,0,elements.length*80);
    const encoder=this.device.createCommandEncoder({label:'Aurelia frame'}),clear=rgba(slide.bg,doc.theme),pass=encoder.beginRenderPass({colorAttachments:[{view:this.context.getCurrentTexture().createView(),clearValue:{r:clear[0],g:clear[1],b:clear[2],a:1},loadOp:'clear',storeOp:'store'}]});let draws=0;
    const resolution=clamp(Math.ceil(width/doc.width*2)/2,.5,3);
    for(let i=0;i<elements.length;){const e=elements[i];if(e.type in SHAPE_IDS){let j=i+1;while(j<elements.length&&elements[j].type in SHAPE_IDS)j++;pass.setPipeline(this.shapePipeline);pass.setBindGroup(0,this.sceneGroup);pass.draw(6,j-i,0,i);draws++;i=j;}else{const t=this.tile(e,doc.theme,resolution);pass.setPipeline(this.texturePipeline);pass.setBindGroup(0,this.sceneGroup);pass.setBindGroup(1,t.bindGroup);pass.draw(6,1,0,i);draws++;i++;}}
    pass.end();this.device.queue.submit([encoder.finish()]);
    if(this.cacheBytes>this.budget){const entries=[...this.cache].sort((a,b)=>a[1].used-b[1].used);for(const [key,v]of entries){if(this.cacheBytes<=this.budget||v.used===this.frame)break;v.texture.destroy();this.cache.delete(key);this.cacheBytes-=v.bytes;}}
    this.lastStats={...this.lastStats,objects:elements.length,draws,ms:performance.now()-start};
  }
  dispose(){this.disposed=true;this.invalidateTiles();this.instances?.destroy();this.uniform?.destroy();this.device?.destroy();}
}
