import {Store,makeDocument,makeSlide,makeElement,validateDocument,uid,clone,clamp,color,THEMES,escapeHTML,localPoint,hitElement,elementBounds,unionBounds,rectIntersects,snapMove} from './core.js';
import {SceneRenderer,ImagePool,paintSlide2D,flattenScene,expandElement,primitivePath,fontCSS,wrapText} from './renderer.js';
import {createDemo} from './demo.js';
import {exportPPTX,importPPTX,zipStore} from './pptx.js';
import {icon} from './icons.js';
const $=(selector,root=document)=>root.querySelector(selector), $$=(selector,root=document)=>[...root.querySelectorAll(selector)];
const store=new Store(createDemo());
const stage=$('#stage'),viewport=$('#viewport'),interaction=$('#interaction'),textEditor=$('#textEditor'),menu=$('#menu'),modal=$('#modal');
const state={tab:'Home',panel:'format',zoom:'fit',scale:1,showGrid:false,snap:true,showNotes:true,showInspector:true,tool:null,drag:null,space:false,editing:null,clipboard:null,guides:[],marquee:null,drawPreview:null,rendererReason:'',presenting:false};
let framePending=0,thumbPending=0,savePending=0,lastPersistedRevision=-1,renderCount=0,storageDB=null,storageAvailable=true;
const thumbnailsImages=new ImagePool(()=>scheduleThumbnails(true));
const renderer=new SceneRenderer($('#surface'),(mode,reason)=>{state.rendererReason=reason;$('#rendererStatus').innerHTML=`<i${mode==='Canvas 2D'?' style="background:#b89a5d"':''}></i>${mode}${mode==='WebGPU'?' accelerated':''}`;$('#rendererStatus').title=reason;},invalidate);
const thumbSignatures=new Map();
const fonts=['Arial','Segoe UI','Georgia','Verdana','Trebuchet MS','Times New Roman','Courier New'];
const RIBBON_TABS=['Home','Insert','Design','Transitions','Animations','Slide Show','View'];
function hydrateIcons(root=document){$$('[data-icon]',root).forEach(el=>el.innerHTML=icon(el.dataset.icon));}
function toast(message,error=false){const el=document.createElement('div');el.className='toast'+(error?' error':'');el.textContent=message;$('#toastRegion').append(el);setTimeout(()=>el.remove(),error?6500:3500);}
function safeRun(fn){return async(...args)=>{try{return await fn(...args);}catch(error){console.error(error);toast(error.message||'The operation could not be completed.',true);}};}
function invalidate(){if(!framePending)framePending=requestAnimationFrame(renderFrame);}
function renderFrame(){framePending=0;renderCount++;const b=stage.getBoundingClientRect();if(b.width>0)renderer.render(store.slide,store.doc,{cssWidth:b.width,cssHeight:b.height,ignoreId:state.editing?.id});drawOverlay();syncTextEditorPosition();}
function resizeStage(){
 const {width,height}=store.doc,w=viewport.clientWidth,h=viewport.clientHeight;
 const fit=Math.min((w-84)/width,(h-73)/height),scale=state.zoom==='fit'?Math.max(.12,fit):state.zoom/100;state.scale=scale;
 const sw=width*scale,sh=height*scale,left=Math.max(42,(w-sw)/2),top=Math.max(18,(h-sh-17)/2);
 Object.assign(stage.style,{width:sw+'px',height:sh+'px',left:left+'px',top:top+'px'});Object.assign($('#stageSizer').style,{width:Math.max(w,sw+84)+'px',height:Math.max(h,sh+70)+'px'});interaction.setAttribute('viewBox',`0 0 ${width} ${height}`);
 $('#zoomSlider').value=clamp(Math.round(scale*100),25,200);$('#zoomLabel').textContent=Math.round(scale*100)+'%';invalidate();
}
function screenToSlide(event){const r=stage.getBoundingClientRect();return{x:(event.clientX-r.left)/state.scale,y:(event.clientY-r.top)/state.scale};}
function zoomTo(value,anchor=null){
 const old=stage.getBoundingClientRect(),p=anchor?{x:(anchor.x-old.left)/state.scale,y:(anchor.y-old.top)/state.scale}:null;
 state.zoom=value==='fit'?'fit':clamp(Math.round(value),25,200);resizeStage();
 if(p){const r=stage.getBoundingClientRect();viewport.scrollLeft+=r.left+p.x*state.scale-anchor.x;viewport.scrollTop+=r.top+p.y*state.scale-anchor.y;}
}
function drawOverlay(){
 const scale=state.scale,k=1/scale,selected=store.selected,b=unionBounds(selected),parts=[];
 if(state.showGrid){parts.push(`<defs><pattern id="editor-grid" width="20" height="20" patternUnits="userSpaceOnUse"><circle cx="0" cy="0" r="${.7*k}" fill="#6b8d80" opacity=".4"/></pattern></defs><rect x="0" y="0" width="${store.doc.width}" height="${store.doc.height}" fill="url(#editor-grid)" pointer-events="none"/>`);}
 for(const guide of state.guides){parts.push(guide.axis==='x'?`<line x1="${guide.value}" y1="0" x2="${guide.value}" y2="${store.doc.height}" stroke="#d071c3" stroke-width="${k}" stroke-dasharray="${4*k} ${3*k}" pointer-events="none"/>`:`<line x1="0" y1="${guide.value}" x2="${store.doc.width}" y2="${guide.value}" stroke="#d071c3" stroke-width="${k}" stroke-dasharray="${4*k} ${3*k}" pointer-events="none"/>`);}
 if(selected.length>1)for(const e of selected)parts.push(`<g transform="translate(${e.x+e.w/2} ${e.y+e.h/2}) rotate(${e.rotation||0})"><rect x="${-e.w/2}" y="${-e.h/2}" width="${e.w}" height="${e.h}" fill="none" stroke="#b9574188" stroke-width="${k}" pointer-events="none"/></g>`);
 if(b&&!state.editing){
  const e=selected.length===1?selected[0]:null,box=e?{x:-e.w/2,y:-e.h/2,w:e.w,h:e.h}:{x:-b.w/2,y:-b.h/2,w:b.w,h:b.h},cx=e?e.x+e.w/2:b.x+b.w/2,cy=e?e.y+e.h/2:b.y+b.h/2,rot=e?.rotation||0;
  const locked=selected.every(e=>e.locked),sc=locked?'#8d9991':'#b95741';parts.push(`<g transform="translate(${cx} ${cy}) rotate(${rot})"><rect x="${box.x}" y="${box.y}" width="${box.w}" height="${box.h}" fill="none" stroke="${sc}" stroke-width="${1.1*k}" pointer-events="none"/>`);
  if(!locked){const handles=[['nw',0,0,'nwse-resize'],['n',.5,0,'ns-resize'],['ne',1,0,'nesw-resize'],['e',1,.5,'ew-resize'],['se',1,1,'nwse-resize'],['s',.5,1,'ns-resize'],['sw',0,1,'nesw-resize'],['w',0,.5,'ew-resize']];
   for(const [key,x,y,cursor]of handles)parts.push(`<rect data-handle="${key}" x="${box.x+x*box.w-3.5*k}" y="${box.y+y*box.h-3.5*k}" width="${7*k}" height="${7*k}" rx="${1*k}" fill="white" stroke="${sc}" stroke-width="${k}" style="cursor:${cursor}"/>`);
   parts.push(`<line x1="0" y1="${box.y}" x2="0" y2="${box.y-24*k}" stroke="${sc}" stroke-width="${k}" pointer-events="none"/><circle data-handle="rotate" cx="0" cy="${box.y-27*k}" r="${4*k}" fill="white" stroke="${sc}" stroke-width="${k}" style="cursor:grab"/>`);
  }parts.push('</g>');
 }
 if(state.marquee){const m=state.marquee;parts.push(`<rect x="${m.x}" y="${m.y}" width="${m.w}" height="${m.h}" fill="#b9574111" stroke="#b95741" stroke-width="${k}" pointer-events="none"/>`);}
 if(state.drawPreview){const d=state.drawPreview;parts.push(`<rect x="${d.x}" y="${d.y}" width="${d.w}" height="${d.h}" fill="#b9574122" stroke="#b95741" stroke-width="${k}" stroke-dasharray="${4*k}" pointer-events="none"/>`);}
 interaction.innerHTML=parts.join('');
}
function bigButton(label,cmd,ic,extra=''){return`<button class="ribbon-large ${extra}" data-cmd="${cmd}" title="${escapeHTML(label)}">${icon(ic)}<span>${label}</span></button>`;}
function smallButton(label,cmd,ic){return`<button data-cmd="${cmd}" title="${escapeHTML(label)}">${icon(ic,16)}${label}</button>`;}
function iconButton(label,cmd,ic,active=false){return`<button data-cmd="${cmd}" aria-label="${escapeHTML(label)}" title="${escapeHTML(label)}" class="${active?'active':''}">${icon(ic,16)}</button>`;}
function ribbonGroup(title,body,extra=''){return`<div class="ribbon-group ${extra}">${body}<span class="group-title">${title}</span></div>`;}
function fontOptions(current){return fonts.map(f=>`<option${f===current?' selected':''}>${escapeHTML(f)}</option>`).join('')+(current&&!fonts.includes(current)?`<option selected>${escapeHTML(current)}</option>`:'');}
function renderRibbon(){
 $('#ribbonTabs').innerHTML=RIBBON_TABS.map(t=>`<button role="tab" aria-selected="${state.tab===t}" class="${state.tab===t?'active':''}" data-cmd="tab:${t}">${t}</button>`).join('');const e=store.selected.find(e=>e.type==='text')||{};let html='';
 if(state.tab==='Home'){
  html+=ribbonGroup('Clipboard',bigButton('Paste','paste','paste')+`<div class="ribbon-stack">${smallButton('Copy','copy','copy')}${smallButton('Duplicate','duplicate','newslide')}</div>`);
  html+=ribbonGroup('Slides',bigButton('New slide','new-slide','newslide')+`<div class="ribbon-stack">${smallButton('Layout','layout-menu','layout')}${smallButton('Delete','delete-slide','trash')}</div>`);
  html+=ribbonGroup('Font',`<div class="ribbon-stack"><div class="ribbon-row"><select data-prop="fontFamily" class="ribbon-font" aria-label="Font family">${fontOptions(e.fontFamily||'Arial')}</select><input data-prop="fontSize" aria-label="Font size" class="ribbon-size" type="number" min="4" max="600" value="${e.fontSize||40}"></div><div class="ribbon-row">${iconButton('Bold (Ctrl B)','bold','bold',e.bold)}${iconButton('Italic (Ctrl I)','italic','italic',e.italic)}${iconButton('Underline (Ctrl U)','underline','underline',e.underline)}<span class="ribbon-color">${icon('fontcolor',17)}<input aria-label="Text color" type="color" data-prop="fill" value="${color(e.fill||'@ink',store.doc.theme).slice(0,7)}"></span>${iconButton('Increase font size','font-grow','plus')}${iconButton('Decrease font size','font-shrink','minus')}</div></div>`,'ribbon-font-group');
  html+=ribbonGroup('Paragraph',`<div class="ribbon-stack"><div class="ribbon-row">${iconButton('Align text left','text-align:left','alignleft',e.align==='left')}${iconButton('Center text','text-align:center','aligncenter',e.align==='center')}${iconButton('Align text right','text-align:right','alignright',e.align==='right')}${iconButton('Toggle bullets','bullets','bullets')}</div><div class="ribbon-row"><select class="ribbon-select" data-prop="valign" aria-label="Vertical text alignment"><option value="top"${e.valign==='top'?' selected':''}>Align top</option><option value="middle"${e.valign==='middle'?' selected':''}>Align middle</option><option value="bottom"${e.valign==='bottom'?' selected':''}>Align bottom</option></select></div></div>`);
  html+=ribbonGroup('Drawing',`<div class="ribbon-mini-shapes">${['rect','ellipse','line','arrow','roundRect','triangle','diamond','star'].map(t=>iconButton(`Draw ${t}`,'draw:'+t,t)).join('')}</div>`+bigButton('Text box','insert-text','text')+`<div class="ribbon-stack"><label class="ribbon-color" title="Shape fill">${icon('fill',17)}<input type="color" aria-label="Shape fill" data-prop="fill" value="${color(store.selected[0]?.fill||'@accent',store.doc.theme).slice(0,7)==='none'?'#ffffff':color(store.selected[0]?.fill||'@accent',store.doc.theme).slice(0,7)}"></label><label class="ribbon-color" title="Shape outline">${icon('outline',17)}<input type="color" aria-label="Shape outline" data-prop="stroke" value="${color(store.selected[0]?.stroke==='none'?'@ink':store.selected[0]?.stroke||'@ink',store.doc.theme).slice(0,7)}"></label></div>`);
  html+=ribbonGroup('Arrange',bigButton('Arrange','arrange-menu','layers')+bigButton('Select','panel:layers','pointer'),'optional-group');
 }
 if(state.tab==='Insert'){
  html+=ribbonGroup('Slides',bigButton('New slide','new-slide','newslide')+bigButton('Layouts','layout-menu','layout'));
  html+=ribbonGroup('Text & media',bigButton('Text box','insert-text','text')+bigButton('Pictures','insert-image','image'));
  html+=ribbonGroup('Illustrations',bigButton('Shapes','shapes-menu','shapes')+`<div class="ribbon-mini-shapes">${['rect','ellipse','line','arrow','roundRect','triangle','diamond','star'].map(t=>iconButton(`Draw ${t}`,'draw:'+t,t)).join('')}</div>`+bigButton('Diagram','insert-diagram','group'));
  html+=ribbonGroup('Data',bigButton('Chart','insert-chart','chart')+bigButton('Table','insert-table','table'));
  html+=ribbonGroup('Slide information',bigButton('Slide number','insert-number','slide')+bigButton('Date','insert-date','timer'));
 }
 if(state.tab==='Design'){
  html+=ribbonGroup('Themes',Object.entries(THEMES).map(([key,t])=>`<button class="ribbon-preview ${store.doc.theme===key?'selected':''}" data-cmd="theme:${key}" title="Apply ${t.name} theme" style="background:${t.bg};color:${t.ink}"><span class="preview-big">Aa</span><span class="preview-line"></span><small>${t.name}</small><i style="background:${t.accent}"></i></button>`).join(''));
  html+=ribbonGroup('Customize',bigButton('Slide size','slide-size','fit')+bigButton('Background','background','fill')+bigButton('Layouts','layout-menu','layout'));
 }
 if(state.tab==='Transitions'){
  html+=ribbonGroup('Transition to this slide',['none','fade','push','zoom'].map(t=>`<button class="effect-button ${store.slide.transition===t?'active':''}" data-cmd="transition:${t}">${icon(t==='none'?'slide':t==='zoom'?'zoomin':'transition')}<span>${t[0].toUpperCase()+t.slice(1)}</span></button>`).join(''));
  html+=ribbonGroup('Timing',`<div class="ribbon-stack"><label style="font-size:10px">Duration <input data-slide-prop="duration" type="number" min="0.1" max="5" step="0.1" value="${store.slide.duration}" style="width:60px;height:27px"> s</label>${smallButton('Apply to all','apply-transition','check')}</div>`);
  html+=ribbonGroup('Preview',bigButton('Preview','preview-transition','play'));
 }
 if(state.tab==='Animations'){
  const anim=store.selected[0]?.animation||'none';html+=ribbonGroup('Entrance animation',['none','fade','rise','zoom'].map(t=>`<button class="effect-button ${anim===t?'active':''}" data-cmd="animation:${t}">${icon(t==='none'?'rect':t==='rise'?'upload':t==='zoom'?'zoomin':'spark')}<span>${t[0].toUpperCase()+t.slice(1)}</span></button>`).join(''));
  html+=ribbonGroup('Playback',`<div style="font-size:11px;color:var(--muted);line-height:1.7;padding:0 10px">Select an object to animate.<br>Animations play on click, in layer order.</div>`+bigButton('Preview','present','play'));
 }
 if(state.tab==='Slide Show'){
  html+=ribbonGroup('Start slide show',bigButton('From beginning','present-first','play')+bigButton('From current','present','screen')+bigButton('Presenter view','present-notes','notes'));
  html+=ribbonGroup('Practice',bigButton('Rehearse','present-rehearse','timer')+bigButton('Speaker notes','toggle-notes','notes'));
  html+=ribbonGroup('Navigate',`<div style="font-size:11px;color:var(--muted);line-height:1.8;padding:0 10px">← → Previous / next &nbsp; · &nbsp; Esc Exit<br>B Black screen &nbsp; · &nbsp; L Laser pointer</div>`);
 }
 if(state.tab==='View'){
  html+=ribbonGroup('Views',bigButton('Normal','normal','slide')+bigButton('Slide sorter','sorter','grid')+bigButton('Slide show','present','screen'));
  html+=ribbonGroup('Show',bigButton('Grid','toggle-grid','grid',state.showGrid?'active':'')+bigButton('Snap guides','toggle-snap','align',state.snap?'active':'')+bigButton('Notes','toggle-notes','notes',state.showNotes?'active':'')+bigButton('Selection pane','panel:layers','layers'));
  html+=ribbonGroup('Zoom',bigButton('Fit to window','zoom-fit','fit')+bigButton('100%','zoom-100','zoomin'));
  html+=ribbonGroup('Workspace',bigButton('Appearance','toggle-dark','sun')+bigButton('Engine','diagnostics','code'));
 }
 $('#ribbon').innerHTML=html;
}
function propertyField(label,name,value,type='number',extra=''){return`<label class="property-field">${label}<input data-prop="${name}" type="${type}" value="${escapeHTML(typeof value==='number'?Math.round(value*10)/10:value||'')}" ${extra}></label>`;}
function paletteHTML(prop,selected){const values=['@ink','@accent','@secondary','@muted','@bg','#FFFFFF','#000000','#DCC497'];return`<div class="color-palette">${values.map(v=>`<button data-color-prop="${prop}" data-value="${v}" title="${color(v,store.doc.theme)}" aria-label="Set ${prop} ${color(v,store.doc.theme)}" class="${selected===v?'selected':''}" style="background:${color(v,store.doc.theme)}"></button>`).join('')}</div>`;}
function colorField(label,prop,value,slideProp=false){const c=value==='none'?'#FFFFFF':color(value,store.doc.theme);return`<div class="color-row"><span>${label}</span><span class="color-field"><input type="color" ${slideProp?'data-slide-prop':'data-prop'}="${prop}" value="${c.slice(0,7)}" aria-label="${label}"><span>${c.toUpperCase()}</span></span></div>`;}
function renderInspector(){
 $('#inspector').hidden=!state.showInspector;$('#propertiesTab').classList.toggle('active',state.panel==='format');$('#layersTab').classList.toggle('active',state.panel==='layers');const e=store.selected[0];
 $('#inspectorTitle').textContent=state.panel==='layers'?'Selection pane':!e?'Slide design':store.selected.length>1?`${store.selected.length} objects selected`:e.type==='text'?'Format text':e.type==='image'?'Format picture':e.type==='chart'?'Format chart':e.type==='table'?'Format table':'Format shape';
 if(state.panel==='layers'){
  $('#inspectorContent').innerHTML=`<div class="layer-tools"><span>${store.slide.elements.length} objects · top to bottom</span><div>${iconButton('Bring forward','forward','front')}${iconButton('Send backward','backward','back')}</div></div>`+store.slide.elements.toReversed().map(e=>`<div class="layer ${store.selection.has(e.id)?'active':''} ${e.hidden?'hidden-layer':''}" data-layer="${e.id}" role="button" tabindex="0" aria-label="Select ${escapeHTML(e.name)}"><span class="layer-icon">${icon(e.type==='text'?'text':e.type)}</span><span class="layer-name" title="${escapeHTML(e.name)}">${escapeHTML(e.type==='text'?(e.text.slice(0,40)||e.name):e.name)}</span><button data-layer-action="visibility" data-id="${e.id}" title="${e.hidden?'Show':'Hide'} object" aria-label="${e.hidden?'Show':'Hide'} object">${icon(e.hidden?'hidden':'eye')}</button><button data-layer-action="lock" data-id="${e.id}" title="${e.locked?'Unlock':'Lock'} object" aria-label="${e.locked?'Unlock':'Lock'} object">${icon(e.locked?'lock':'unlock')}</button></div>`).join('');return;
 }
 if(!e){const t=THEMES[store.doc.theme];
  $('#inspectorContent').innerHTML=`<section class="property-section"><h3>Your theme <span>Made for your story</span></h3><div class="design-preview" style="background:${t.bg};color:${t.ink}"><div class="tag">A FRESH PERSPECTIVE</div><div class="words">Ideas<br>in motion.</div><div class="arch"></div><div class="dot" style="background:${t.accent}"></div></div><div class="theme-caption"><span>${t.name}</span><span class="tag-pill">16:9 · ${store.doc.width} × ${store.doc.height}</span></div><div class="theme-swatches">${Object.entries(THEMES).map(([key,t])=>`<button class="theme-swatch ${key===store.doc.theme?'active':''}" data-cmd="theme:${key}" title="Apply ${t.name} theme" aria-label="Apply ${t.name} theme"><span style="background:${t.bg}"></span><span style="background:${t.accent}"></span><span style="background:${t.dark}"></span></button>`).join('')}</div></section>
  <section class="property-section"><h3>Background</h3>${colorField('Solid fill','bg',store.slide.bg,true)}${paletteHTML('bg',store.slide.bg)}</section>
  <section class="property-section"><h3>Add a slide <span>Choose a layout</span></h3><div class="layout-grid">${['title','content','split','blank'].map(t=>`<button class="layout-choice" data-cmd="new-layout:${t}"><span class="layout-mini ${t}"></span>${{title:'Title slide',content:'Title & content',split:'Two columns',blank:'Blank canvas'}[t]}</button>`).join('')}</div></section>
  <section class="property-section"><h3>Transition</h3><select data-slide-prop="transition" aria-label="Slide transition" style="width:100%;font-size:11px">${['none','fade','push','zoom'].map(t=>`<option value="${t}"${store.slide.transition===t?' selected':''}>${t[0].toUpperCase()+t.slice(1)}</option>`).join('')}</select><p>Select an object to format it.<br>Double-click text to start writing.</p></section>`;return;
 }
 const fill=e.fill==='none'?'#FFFFFF':e.fill;
 let html=`<section class="property-section"><h3>Transform <span>${e.locked?'Locked':e.groupId?'Grouped':'Slide coordinates'}</span></h3><div class="property-grid">${propertyField('X position','x',e.x)}${propertyField('Y position','y',e.y)}${propertyField('Width','w',e.w,'number','min="1"')}${propertyField('Height','h',e.h,'number','min="1"')}${propertyField('Rotation °','rotation',e.rotation)}${propertyField('Opacity 0–1','opacity',e.opacity,'number','min="0" max="1" step="0.05"')}</div><div class="inline-actions" style="margin-top:12px">${smallButton(e.locked?'Unlock':'Lock','lock',e.locked?'unlock':'lock')}${smallButton('Arrange','arrange-menu','layers')}</div></section>`;
 if(e.type==='text')html+=`<section class="property-section"><h3>Typography</h3><div class="property-grid"><label class="property-field full">Font family<select data-prop="fontFamily">${fontOptions(e.fontFamily)}</select></label>${propertyField('Font size','fontSize',e.fontSize,'number','min="4" max="600"')}${propertyField('Line height','lineHeight',e.lineHeight,'number','min="0.8" max="3" step="0.05"')}</div><div class="inline-actions" style="margin-top:12px">${iconButton('Bold','bold','bold',e.bold)}${iconButton('Italic','italic','italic',e.italic)}${iconButton('Underline','underline','underline',e.underline)}${iconButton('Left','text-align:left','alignleft',e.align==='left')}${iconButton('Center','text-align:center','aligncenter',e.align==='center')}${iconButton('Right','text-align:right','alignright',e.align==='right')}</div><button class="property-button" data-cmd="edit-text">${icon('text')}Edit text</button></section>`;
 if(e.type!=='image')html+=`<section class="property-section"><h3>${e.type==='text'?'Text color':'Appearance'}</h3>${colorField(e.type==='text'?'Color':'Fill','fill',fill)}${paletteHTML('fill',e.fill)}${e.type!=='text'?`<div style="margin-top:15px">${colorField('Outline','stroke',e.stroke==='none'?'@ink':e.stroke)}</div><div class="property-grid">${propertyField('Outline width','strokeWidth',e.strokeWidth,'number','min="0" max="100"')}${e.type==='roundRect'?propertyField('Corner radius','radius',e.radius,'number','min="0"'):''}</div><button class="property-button" data-cmd="no-fill">No fill</button>`:''}</section>`;
 if(e.type==='image')html+=`<section class="property-section"><h3>Picture</h3><label class="property-field">Image fit<select data-prop="fit"><option value="cover"${e.fit==='cover'?' selected':''}>Fill (center crop)</option><option value="contain"${e.fit==='contain'?' selected':''}>Fit entire image</option></select></label><button class="property-button" data-cmd="replace-image">${icon('image')}Replace picture</button></section>`;
 if(e.type==='chart')html+=`<section class="property-section"><h3>Chart data</h3><label class="property-field">Chart type<select data-prop="chartType">${['bar','line','donut'].map(t=>`<option value="${t}"${t===e.chartType?' selected':''}>${t[0].toUpperCase()+t.slice(1)}</option>`).join('')}</select></label><button class="property-button" data-cmd="edit-data">${icon('chart')}Edit chart data</button></section>`;
 if(e.type==='table')html+=`<section class="property-section"><h3>Table data</h3>${propertyField('Cell font size','fontSize',e.fontSize)}<button class="property-button" data-cmd="edit-data">${icon('table')}Edit table</button></section>`;
 html+=`<section class="property-section"><h3>Entrance animation</h3><select data-prop="animation" style="width:100%;font-size:11px" aria-label="Object animation">${['none','fade','rise','zoom'].map(t=>`<option value="${t}"${e.animation===t?' selected':''}>${t==='none'?'No animation':t[0].toUpperCase()+t.slice(1)+' · On click'}</option>`).join('')}</select></section>`;
 $('#inspectorContent').innerHTML=html;
}
function buildThumbnails(){
 const ids=store.doc.slides.map(s=>s.id).join('|');if($('#thumbnails').dataset.ids!==ids){$('#thumbnails').dataset.ids=ids;$('#thumbnails').innerHTML=store.doc.slides.map((s,i)=>`<button class="thumb" data-slide="${s.id}" draggable="true" role="option" aria-label="Slide ${i+1}: ${escapeHTML(s.name)}"><span class="thumb-number">${i+1}</span><span class="thumb-preview"><canvas width="320" height="180"></canvas><span class="thumb-label">${escapeHTML(s.name)}</span></span></button>`).join('');thumbSignatures.clear();}
 $$('.thumb').forEach(t=>{t.classList.toggle('active',t.dataset.slide===store.active);t.setAttribute('aria-selected',t.dataset.slide===store.active?'true':'false');});
}
function scheduleThumbnails(force=false){if(force)thumbSignatures.clear();if(thumbPending)return;thumbPending=requestAnimationFrame(()=>{thumbPending=0;let painted=0,more=false;for(const s of store.doc.slides){const button=$(`[data-slide="${s.id}"]`);if(!button)continue;const sig=JSON.stringify([s.bg,s.elements,store.doc.theme,store.doc.width,store.doc.height]);if(thumbSignatures.get(s.id)===sig)continue;if(painted>=4){more=true;break;}const c=$('canvas',button);c.width=320;c.height=Math.round(320*store.doc.height/store.doc.width);c.style.aspectRatio=store.doc.width+'/'+store.doc.height;paintSlide2D(c,s,store.doc,thumbnailsImages,{scale:320/store.doc.width});thumbSignatures.set(s.id,sig);painted++;}if(more)scheduleThumbnails();});}
function syncUI(kind='change'){
 if(document.activeElement!==$('#documentTitle'))$('#documentTitle').value=store.doc.title;
 document.title=`${store.doc.title} — Aurelia Slides`;
 if(document.activeElement!==$('#notes'))$('#notes').value=store.slide.notes;
 const i=store.doc.slides.findIndex(s=>s.id===store.active);$('#slideCount').textContent=store.doc.slides.length;$('#slideStatus').textContent=`Slide ${i+1} of ${store.doc.slides.length}`;$('#slideBreadcrumb').textContent=`SLIDE ${String(i+1).padStart(2,'0')}`;$('#slideName').textContent=store.slide.name;
 $('#selectionStatus').textContent=state.editing?'Editing text':store.selected.length?store.selected.length===1?store.selected[0].name:`${store.selected.length} objects selected`:'Ready';
 $$('[data-cmd="undo"]').forEach(b=>{b.disabled=!store.past.length;b.title=store.past.length?`Undo: ${store.past.at(-1).label}`:'Nothing to undo';});$$('[data-cmd="redo"]').forEach(b=>b.disabled=!store.future.length);
 buildThumbnails();if(!['live','selection'].includes(kind))scheduleThumbnails();
 if(!['live'].includes(kind)){renderRibbon();if(!$('#inspectorContent').contains(document.activeElement))renderInspector();}
 invalidate();
}
store.onChange(kind=>{syncUI(kind);if(kind==='commit'||kind==='restore'){schedulePersist();resizeStage();const sources=new Set(store.doc.slides.flatMap(s=>s.elements.filter(e=>e.type==='image').map(e=>e.src)));renderer.images.retain(sources);thumbnailsImages.retain(sources);}});
async function openStorage(){
 try{storageDB=await new Promise((resolve,reject)=>{const req=indexedDB.open('aurelia-slides-v1',1);req.onupgradeneeded=()=>req.result.createObjectStore('documents');req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);req.onblocked=()=>reject(new Error('Local storage is blocked by another tab.'));});return await new Promise((resolve,reject)=>{const tx=storageDB.transaction('documents','readonly'),r=tx.objectStore('documents').get('current');r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});}
 catch(error){storageAvailable=false;try{const s=localStorage.getItem('aurelia-backup');return s?JSON.parse(s):null;}catch{return null;}}
}
function schedulePersist(){clearTimeout(savePending);$('#saveText').textContent='Saving locally…';savePending=setTimeout(persist,450);}
async function persist(){
 clearTimeout(savePending);const revision=store.revision,doc=clone(store.doc);
 try{if(storageDB){await new Promise((resolve,reject)=>{const tx=storageDB.transaction('documents','readwrite');tx.objectStore('documents').put(doc,'current');tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error);});}else localStorage.setItem('aurelia-backup',JSON.stringify(doc));lastPersistedRevision=revision;$('#saveText').textContent='Saved on this device';$('#saveIndicator').title='Saved on this device. Export a file to keep an independent backup.';}
 catch(error){$('#saveText').textContent='Not saved';$('#saveIndicator').title=error.message;toast('Local storage is unavailable or full. Use File → Save .aurelia to keep your work.',true);}
}
function fileName(ext){return(store.doc.title.replace(/[<>:"/\\|?*\u0000-\u001f]/g,'').slice(0,120)||'Presentation')+'.'+ext;}
function downloadBlob(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),20000);}
function nativeSave(){finishTextEditing();persist();downloadBlob(new Blob([JSON.stringify(store.doc,null,2)],{type:'application/json'}),fileName('aurelia'));toast('Presentation saved as an editable .aurelia file.');}
function closeMenu(){menu.hidden=true;menu.innerHTML='';}
function showMenu(html,anchor){menu.innerHTML=html;menu.hidden=false;const rect=anchor?.getBoundingClientRect?.()||{left:anchor?.x||40,bottom:anchor?.y||120};menu.style.left=rect.left+'px';menu.style.top=(rect.bottom+5)+'px';const b=menu.getBoundingClientRect();menu.style.left=clamp(rect.left,8,window.innerWidth-b.width-8)+'px';menu.style.top=clamp(rect.bottom+5,8,window.innerHeight-b.height-8)+'px';}
function menuItem(label,cmd,ic,shortcut=''){return`<button class="menu-item" data-cmd="${cmd}">${icon(ic)}<span>${label}</span>${shortcut?`<kbd>${shortcut}</kbd>`:''}</button>`;}
function showFileMenu(anchor){showMenu(`<div class="menu-title">Presentation</div>${menuItem('New presentation','new-document','newslide','Ctrl N')}${menuItem('Open .aurelia / .pptx','open','folder','Ctrl O')}${menuItem('Save .aurelia','save','save','Ctrl S')}<div class="menu-separator"></div><div class="menu-title">Export</div>${menuItem('PowerPoint .pptx','export-pptx','export')}${menuItem('Current slide · PNG','export-png','image')}${menuItem('Current slide · SVG','export-svg','shapes')}${menuItem('All slides · PNG archive','export-images','download')}${menuItem('Print / Save as PDF','print','slide')}<div class="menu-separator"></div>${menuItem('Presentation settings','slide-size','fit')}${menuItem('Restore example deck','restore-demo','spark')}${menuItem('About Aurelia Slides','about','help')}`,anchor);}
function showShapeMenu(anchor){showMenu(`<div class="menu-title">Draw a shape</div><div class="menu-grid">${['rect','roundRect','ellipse','triangle','diamond','arrow','star','line'].map(t=>`<button data-cmd="draw:${t}" title="${t}">${icon(t)}<span>${{rect:'Rectangle',roundRect:'Rounded',ellipse:'Ellipse',triangle:'Triangle',diamond:'Diamond',arrow:'Arrow',star:'Star',line:'Line'}[t]}</span></button>`).join('')}</div><div class="menu-title">Click & drag on the slide</div>`,anchor);}
function showLayoutMenu(anchor){showMenu(`<div class="menu-title">New slide layout</div>${menuItem('Title slide','new-layout:title','text')}${menuItem('Title and content','new-layout:content','slide')}${menuItem('Two columns','new-layout:split','layout')}${menuItem('Quote','new-layout:quote','text')}${menuItem('Blank canvas','new-layout:blank','rect')}<div class="menu-separator"></div>${menuItem('Duplicate current slide','duplicate-slide','copy')}${menuItem('Reset current slide layout…','reset-layout','layout')}`,anchor);}
function showArrangeMenu(anchor){showMenu(`<div class="menu-title">Arrange</div>${menuItem('Bring to front','front','front')}${menuItem('Send to back','back','back')}${menuItem('Bring forward','forward','front')}${menuItem('Send backward','backward','back')}<div class="menu-separator"></div>${menuItem('Align left','align:left','alignleft')}${menuItem('Align center','align:center','aligncenter')}${menuItem('Align right','align:right','alignright')}${menuItem('Align top','align:top','align')}${menuItem('Align middle','align:middle','align')}${menuItem('Align bottom','align:bottom','align')}${menuItem('Distribute horizontally','distribute:x','layout')}${menuItem('Distribute vertically','distribute:y','layout')}<div class="menu-separator"></div>${menuItem('Group','group','group','Ctrl G')}${menuItem('Ungroup','ungroup','shapes','Ctrl Shift G')}${menuItem('Lock / unlock','lock','lock')}`,anchor);}
function openModal(title,body,actions=[],width=520){finishTextEditing();closeMenu();$('#modalTitle').textContent=title;$('#modalContent').innerHTML=body;$('#modalActions').innerHTML='';modal.style.width=width+'px';for(const a of actions){const button=document.createElement('button');button.textContent=a.label;button.className=a.primary?'primary':'secondary';button.addEventListener('click',safeRun(async()=>{const done=await a.action();if(done!==false)modal.close();}));$('#modalActions').append(button);}if(!modal.open)modal.showModal();}
function askReplace(title,callback){openModal(title,'<p class="modal-description">This replaces the current presentation. Your current work remains available through Undo. Save an independent .aurelia file first to keep a backup.</p>',[{label:'Cancel',action:()=>{}},{label:'Save current file',action:()=>{nativeSave();return false;}},{label:'Continue',primary:true,action:callback}]);}
function selectElement(element,additive=false){
 let ids=element.groupId?store.slide.elements.filter(e=>e.groupId===element.groupId&&!e.hidden).map(e=>e.id):[element.id];
 if(additive){const current=new Set(store.selection),remove=ids.every(id=>current.has(id));ids.forEach(id=>remove?current.delete(id):current.add(id));ids=[...current];}
 store.select(ids);
}
function setTool(type){finishTextEditing();state.tool=type;store.select([]);stage.classList.add('drawing');$('#canvasHint').textContent=`Draw ${type==='text'?'a text box':type} · Click and drag on the slide · Esc cancels`;closeMenu();}
function clearTool(){state.tool=null;stage.classList.remove('drawing');$('#canvasHint').textContent='Double-click text to edit · Drag to move · Shift-click to select multiple';}
function startTextEditing(e){
 if(!e||e.type!=='text'||e.locked)return;finishTextEditing();clearTool();store.select([e.id]);store.begin('Edit text');state.editing={id:e.id,original:e.text};textEditor.value=e.text;textEditor.hidden=false;syncTextEditorPosition();textEditor.focus();textEditor.select();syncUI('selection');
}
function syncTextEditorPosition(){
 if(!state.editing)return;const e=store.slide.elements.find(e=>e.id===state.editing.id);if(!e)return;const sc=state.scale;
 Object.assign(textEditor.style,{left:e.x*sc+'px',top:e.y*sc+'px',width:e.w*sc+'px',height:e.h*sc+'px',fontFamily:`"${e.fontFamily}", sans-serif`,fontSize:e.fontSize*sc+'px',fontWeight:e.bold?'700':'400',fontStyle:e.italic?'italic':'normal',textDecoration:e.underline?'underline':'none',lineHeight:e.lineHeight||1.2,color:color(e.fill,store.doc.theme),textAlign:e.align||'left',padding:(e.padding||0)*sc+'px',transform:`rotate(${e.rotation||0}deg)`});
 if(e.valign==='middle'||e.valign==='bottom'){const c=document.createElement('canvas'),ctx=c.getContext('2d');ctx.font=fontCSS(e);const lines=wrapText(ctx,e.text,e.w-2*(e.padding||0)),h=lines.length*e.fontSize*(e.lineHeight||1.2),p=e.valign==='middle'?Math.max(e.padding||0,(e.h-h)/2):Math.max(e.padding||0,e.h-h-(e.padding||0));textEditor.style.paddingTop=p*sc+'px';}
}
function finishTextEditing(cancel=false){
 if(!state.editing)return;const id=state.editing.id;state.editing=null;textEditor.hidden=true;
 if(cancel)store.cancel();else{const e=store.slide.elements.find(e=>e.id===id);if(e)e.text=textEditor.value;store.commit();}
 viewport.focus({preventScroll:true});syncUI('selection');scheduleThumbnails();
}
textEditor.addEventListener('input',()=>{const e=store.slide.elements.find(e=>e.id===state.editing?.id);if(e){e.text=textEditor.value;invalidate();}});
textEditor.addEventListener('pointerdown',event=>event.stopPropagation());
textEditor.addEventListener('keydown',event=>{event.stopPropagation();if(event.key==='Escape'){event.preventDefault();finishTextEditing(true);}if((event.metaKey||event.ctrlKey)&&event.key==='Enter'){event.preventDefault();finishTextEditing();}});
textEditor.addEventListener('blur',()=>{if(state.editing)setTimeout(()=>{if(state.editing&&document.activeElement!==textEditor)finishTextEditing();},0);});
stage.addEventListener('pointerdown',event=>{
 if(event.button!==0&&event.button!==1)return;
 if(state.space||event.button===1){beginPan(event);return;}
 if(state.editing)finishTextEditing();closeMenu();event.preventDefault();viewport.focus({preventScroll:true});const p=screenToSlide(event),handle=event.target.dataset?.handle;
 if(state.tool){state.drag={type:'draw',start:p,tool:state.tool,pointer:event.pointerId};stage.setPointerCapture(event.pointerId);return;}
 if(handle&&store.selected.length){const orig=clone(store.selected.filter(e=>!e.locked));if(!orig.length)return;store.begin(handle==='rotate'?'Rotate objects':'Resize objects');const b=unionBounds(orig);state.drag={type:handle==='rotate'?'rotate':'resize',handle,start:p,orig,bounds:b,pointer:event.pointerId,angle:Math.atan2(p.y-b.y-b.h/2,p.x-b.x-b.w/2)};stage.setPointerCapture(event.pointerId);return;}
 const hit=store.slide.elements.toReversed().find(e=>hitElement(e,p,3/state.scale));
 if(hit){
  if(event.shiftKey)selectElement(hit,true);else if(!store.selection.has(hit.id))selectElement(hit);
  if(!store.selection.has(hit.id))return;
  if(event.altKey&&!hit.locked)store.duplicateSelected();
  const orig=clone(store.selected.filter(e=>!e.locked));if(!orig.length)return;store.begin('Move objects');state.drag={type:'move',start:p,orig,bounds:unionBounds(orig),pointer:event.pointerId};
 }else{state.drag={type:'marquee',start:p,old:event.shiftKey?[...store.selection]:[],pointer:event.pointerId};if(!event.shiftKey)store.select([]);}
 stage.setPointerCapture(event.pointerId);
});
function beginPan(event){event.preventDefault();state.drag={type:'pan',client:{x:event.clientX,y:event.clientY},scroll:{x:viewport.scrollLeft,y:viewport.scrollTop},pointer:event.pointerId};viewport.setPointerCapture(event.pointerId);viewport.classList.add('dragging');}
viewport.addEventListener('pointerdown',event=>{if(event.target===viewport||event.target===$('#stageSizer')){if(state.space||event.button===1)beginPan(event);else if(event.button===0){finishTextEditing();store.select([]);}}});
function applyOriginalChange(orig,fn){for(const initial of orig){const e=store.slide.elements.find(x=>x.id===initial.id);if(e)fn(e,initial);}}
function resizeSingle(d,p,event){
 const o=d.orig[0],q=localPoint(o,p),start=localPoint(o,d.start);let l=0,t=0,r=o.w,b=o.h;const h=d.handle;
 if(h.includes('w'))l=Math.min(o.w-8,q.x-start.x);if(h.includes('e'))r=Math.max(8,o.w+q.x-start.x);if(h.includes('n'))t=Math.min(o.h-8,q.y-start.y);if(h.includes('s'))b=Math.max(8,o.h+q.y-start.y);
 if(event.shiftKey&&h.length===2){const ratio=o.w/o.h,nw=r-l,nh=b-t;if(nw/nh>ratio){const desired=nw/ratio;if(h.includes('n'))t=b-desired;else b=t+desired;}else{const desired=nh*ratio;if(h.includes('w'))l=r-desired;else r=l+desired;}}
 const w=r-l,hh=b-t,dx=(l+r)/2-o.w/2,dy=(t+b)/2-o.h/2,a=(o.rotation||0)*Math.PI/180,c=Math.cos(a),s=Math.sin(a);
 applyOriginalChange(d.orig,e=>{e.x=o.x+o.w/2+dx*c-dy*s-w/2;e.y=o.y+o.h/2+dx*s+dy*c-hh/2;e.w=w;e.h=hh;});
}
function resizeMultiple(d,p,event){
 const b=d.bounds,h=d.handle,dx=p.x-d.start.x,dy=p.y-d.start.y;let x=b.x,y=b.y,w=b.w,hh=b.h;
 if(h.includes('w')){x=Math.min(b.x+b.w-10,b.x+dx);w=b.x+b.w-x;}if(h.includes('e'))w=Math.max(10,b.w+dx);if(h.includes('n')){y=Math.min(b.y+b.h-10,b.y+dy);hh=b.y+b.h-y;}if(h.includes('s'))hh=Math.max(10,b.h+dy);
 if(event.shiftKey&&h.length===2){const s=Math.max(w/b.w,hh/b.h);w=b.w*s;hh=b.h*s;if(h.includes('w'))x=b.x+b.w-w;if(h.includes('n'))y=b.y+b.h-hh;}
 const sx=w/b.w,sy=hh/b.h;applyOriginalChange(d.orig,(e,o)=>{e.x=x+(o.x-b.x)*sx;e.y=y+(o.y-b.y)*sy;e.w=Math.max(1,o.w*sx);e.h=Math.max(1,o.h*sy);if(e.type==='text'||e.type==='table')e.fontSize=o.fontSize*Math.min(sx,sy);});
}
function pointerMove(event){
 if(!state.drag){if(!state.tool&&!state.editing){const p=screenToSlide(event);const e=store.slide.elements.toReversed().find(e=>hitElement(e,p,2/state.scale));interaction.style.cursor=e?(e.locked?'not-allowed':'move'):'default';}return;}
 const d=state.drag;if(event.pointerId!==d.pointer)return;
 if(d.type==='pan'){viewport.scrollLeft=d.scroll.x+d.client.x-event.clientX;viewport.scrollTop=d.scroll.y+d.client.y-event.clientY;return;}
 const p=screenToSlide(event),dx=p.x-d.start.x,dy=p.y-d.start.y;
 if(d.type==='draw'){state.drawPreview={x:Math.min(p.x,d.start.x),y:Math.min(p.y,d.start.y),w:Math.abs(dx),h:Math.abs(dy)};if(event.shiftKey)state.drawPreview.h=state.drawPreview.w;}
 if(d.type==='marquee'){state.marquee={x:Math.min(p.x,d.start.x),y:Math.min(p.y,d.start.y),w:Math.abs(dx),h:Math.abs(dy)};const selected=store.slide.elements.filter(e=>!e.hidden&&rectIntersects(elementBounds(e),state.marquee)).flatMap(e=>e.groupId?store.slide.elements.filter(o=>o.groupId===e.groupId).map(o=>o.id):[e.id]);store.selection=new Set([...d.old,...selected]);}
 if(d.type==='move'){
  let x=dx,y=dy;if(event.shiftKey){if(Math.abs(x)>Math.abs(y))y=0;else x=0;}state.guides=[];
  if(state.snap&&!event.altKey){const b={...d.bounds,x:d.bounds.x+x,y:d.bounds.y+y},others=store.slide.elements.filter(e=>!store.selection.has(e.id)&&!e.hidden),snap=snapMove(b,others,store.doc.width,store.doc.height,6/state.scale);x+=snap.dx;y+=snap.dy;state.guides=snap.lines;}
  applyOriginalChange(d.orig,(e,o)=>{e.x=o.x+x;e.y=o.y+y;});
 }
 if(d.type==='resize'){if(d.orig.length===1)resizeSingle(d,p,event);else resizeMultiple(d,p,event);}
 if(d.type==='rotate'){const b=d.bounds,cx=b.x+b.w/2,cy=b.y+b.h/2;let delta=(Math.atan2(p.y-cy,p.x-cx)-d.angle)*180/Math.PI;if(event.shiftKey)delta=Math.round(delta/15)*15;const a=delta*Math.PI/180,c=Math.cos(a),s=Math.sin(a);applyOriginalChange(d.orig,(e,o)=>{const x=o.x+o.w/2-cx,y=o.y+o.h/2-cy;e.x=cx+x*c-y*s-o.w/2;e.y=cy+x*s+y*c-o.h/2;e.rotation=(o.rotation+delta+360)%360;});}
 invalidate();
}
function pointerUp(event){
 const d=state.drag;if(!d||event.pointerId!==d.pointer)return;state.drag=null;state.guides=[];state.marquee=null;state.drawPreview=null;viewport.classList.remove('dragging');
 if(d.type==='draw'){
  const p=screenToSlide(event);let w=Math.abs(p.x-d.start.x),h=Math.abs(p.y-d.start.y),x=Math.min(p.x,d.start.x),y=Math.min(p.y,d.start.y),rotation=0;
  if(w<8&&h<8){x=d.start.x;y=d.start.y;w=d.tool==='text'?420:240;h=d.tool==='text'?90:d.tool==='line'?3:170;}
  else if(d.tool==='line'){w=Math.hypot(p.x-d.start.x,p.y-d.start.y);h=3;x=(p.x+d.start.x)/2-w/2;y=(p.y+d.start.y)/2-h/2;rotation=Math.atan2(p.y-d.start.y,p.x-d.start.x)*180/Math.PI;}
  else if(event.shiftKey)h=w;
  const type=d.tool;clearTool();const e=makeElement(type,{x,y,w:Math.max(8,w),h:Math.max(type==='line'?1:8,h),rotation});store.add(e);if(type==='text')startTextEditing(e);
 }else if(['move','resize','rotate'].includes(d.type)){store.commit();renderInspector();}
 else if(d.type==='marquee')store.emit('selection');
 invalidate();
}
stage.addEventListener('pointermove',pointerMove);stage.addEventListener('pointerup',pointerUp);viewport.addEventListener('pointermove',event=>{if(state.drag?.type==='pan')pointerMove(event);});viewport.addEventListener('pointerup',event=>{if(state.drag?.type==='pan')pointerUp(event);});
function cancelPointer(){if(state.drag&&store.pending)store.cancel();state.drag=null;state.guides=[];state.marquee=null;state.drawPreview=null;viewport.classList.remove('dragging');invalidate();}
stage.addEventListener('pointercancel',cancelPointer);viewport.addEventListener('pointercancel',cancelPointer);
stage.addEventListener('dblclick',event=>{if(state.tool)return;const p=screenToSlide(event),e=store.slide.elements.toReversed().find(e=>hitElement(e,p,2/state.scale));if(!e)return;store.select([e.id]);if(e.type==='text')startTextEditing(e);else if(e.type==='chart'||e.type==='table')editData(e);});
stage.addEventListener('contextmenu',event=>{event.preventDefault();finishTextEditing();const p=screenToSlide(event),hit=store.slide.elements.toReversed().find(e=>hitElement(e,p,3/state.scale));if(hit&&!store.selection.has(hit.id))selectElement(hit);let items=menuItem('Paste','paste','paste','Ctrl V');if(store.selected.length)items=menuItem('Copy','copy','copy','Ctrl C')+menuItem('Duplicate','duplicate','copy','Ctrl D')+items+menuItem('Delete','delete','trash','Del')+'<div class="menu-separator"></div>'+menuItem('Bring to front','front','front')+menuItem('Send to back','back','back')+menuItem('Group','group','group')+menuItem('Ungroup','ungroup','shapes')+menuItem('Lock / unlock','lock','lock');else items+=menuItem('New slide','new-slide','newslide')+menuItem('Background','background','fill');showMenu(items,{x:event.clientX,y:event.clientY});});
viewport.addEventListener('wheel',event=>{if(event.ctrlKey||event.metaKey){event.preventDefault();zoomTo(state.scale*100*Math.exp(-event.deltaY*.003),{x:event.clientX,y:event.clientY});}},{passive:false});
let draggingSlide=null;
$('#thumbnails').addEventListener('click',event=>{const t=event.target.closest('[data-slide]');if(t){finishTextEditing();clearTool();store.activate(t.dataset.slide);renderInspector();}});
$('#thumbnails').addEventListener('contextmenu',event=>{const t=event.target.closest('[data-slide]');if(!t)return;event.preventDefault();store.activate(t.dataset.slide);showMenu(menuItem('New slide','new-slide','newslide')+menuItem('Duplicate slide','duplicate-slide','copy')+menuItem('Delete slide','delete-slide','trash')+menuItem('Rename slide','rename-slide','text'),{x:event.clientX,y:event.clientY});});
$('#thumbnails').addEventListener('dragstart',event=>{const t=event.target.closest('[data-slide]');if(!t)return;draggingSlide=t.dataset.slide;event.dataTransfer.effectAllowed='move';event.dataTransfer.setData('text/plain',draggingSlide);t.classList.add('dragging');});
$('#thumbnails').addEventListener('dragover',event=>{const t=event.target.closest('[data-slide]');if(!t||!draggingSlide)return;event.preventDefault();$$('.thumb.dragover').forEach(e=>e.classList.remove('dragover'));t.classList.add('dragover');});
$('#thumbnails').addEventListener('drop',event=>{const t=event.target.closest('[data-slide]');event.preventDefault();if(t&&draggingSlide){const from=store.doc.slides.findIndex(s=>s.id===draggingSlide),to=store.doc.slides.findIndex(s=>s.id===t.dataset.slide);if(from>=0&&to>=0&&from!==to)store.reorderSlide(from,to);}draggingSlide=null;$$('.thumb').forEach(t=>t.classList.remove('dragover','dragging'));});
$('#thumbnails').addEventListener('dragend',()=>{draggingSlide=null;$$('.thumb').forEach(t=>t.classList.remove('dragover','dragging'));});
function selectedText(){return store.selected.filter(e=>e.type==='text'&&!e.locked);}
function formatText(props,label='Format text'){if(!selectedText().length){toast('Select a text box to change its formatting.');return;}store.transaction(label,()=>selectedText().forEach(e=>Object.assign(e,props)));}
function updateProperty(name,value){
 if(!store.selected.length){toast('Select an object first.');renderRibbon();return;}
 const numbers={x:[-100000,100000],y:[-100000,100000],w:[1,16384],h:[1,16384],rotation:[-36000,36000],opacity:[0,1],fontSize:[4,600],lineHeight:[.8,3],strokeWidth:[0,100],radius:[0,2000]};
 if(name in numbers){value=Number(value);if(!Number.isFinite(value))return;value=clamp(value,...numbers[name]);}
 if(['fontFamily','bold','italic','underline','align','valign','lineHeight'].includes(name)){formatText({[name]:value});return;}
 store.transaction('Change '+name,()=>{for(const e of store.selected)if(!e.locked){e[name]=value;if(name==='stroke'&&e.strokeWidth===0)e.strokeWidth=2;}});
}
document.addEventListener('change',event=>{
 const el=event.target;
 if(el.matches('[data-prop]'))updateProperty(el.dataset.prop,el.type==='checkbox'?el.checked:el.value);
 if(el.matches('[data-slide-prop]')){const key=el.dataset.slideProp,value=key==='duration'?clamp(+el.value||.45,.1,5):el.value;store.transaction('Change slide '+key,()=>store.slide[key]=value);}
});
let notesTimer=0;
$('#notes').addEventListener('focus',()=>store.begin('Edit speaker notes'));
$('#notes').addEventListener('input',()=>{if(!store.pending)store.begin('Edit speaker notes');store.slide.notes=$('#notes').value;clearTimeout(notesTimer);notesTimer=setTimeout(()=>{store.commit();},800);});
$('#notes').addEventListener('blur',()=>{clearTimeout(notesTimer);if(store.pending?.label==='Edit speaker notes')store.commit();});
$('#documentTitle').addEventListener('change',()=>{const value=$('#documentTitle').value.trim()||'Untitled presentation';store.transaction('Rename presentation',()=>store.doc.title=value.slice(0,200));});
$('#zoomSlider').addEventListener('input',()=>zoomTo(+$('#zoomSlider').value));
$('#modalClose').addEventListener('click',()=>modal.close());
modal.addEventListener('click',event=>{const r=modal.getBoundingClientRect();if(event.target===modal&&(event.clientX<r.left||event.clientX>r.right||event.clientY<r.top||event.clientY>r.bottom))modal.close();});
modal.addEventListener('close',()=>{viewport.focus({preventScroll:true});});
document.addEventListener('pointerdown',event=>{if(!menu.hidden&&!menu.contains(event.target)&&!event.target.closest('[data-cmd]'))closeMenu();});
document.addEventListener('click',safeRun(async event=>{
 const colorButton=event.target.closest('[data-color-prop]');if(colorButton){const key=colorButton.dataset.colorProp,val=colorButton.dataset.value;if(key==='bg')store.transaction('Background color',()=>store.slide.bg=val);else updateProperty(key,val);return;}
 const layerAction=event.target.closest('[data-layer-action]');if(layerAction){const e=store.slide.elements.find(e=>e.id===layerAction.dataset.id);if(e)store.transaction(layerAction.dataset.layerAction==='lock'?'Toggle object lock':'Toggle object visibility',()=>{if(layerAction.dataset.layerAction==='lock')e.locked=!e.locked;else e.hidden=!e.hidden;});renderInspector();return;}
 const layer=event.target.closest('[data-layer]');if(layer){const e=store.slide.elements.find(e=>e.id===layer.dataset.layer);if(e)selectElement(e,event.shiftKey);renderInspector();return;}
 const button=event.target.closest('[data-cmd]');if(button){const cmd=button.dataset.cmd;if(menu.contains(button))closeMenu();await command(cmd,button);}
}));
async function copySelection(){
 if(!store.selected.length){toast('Select objects to copy.');return;}state.clipboard=clone(store.selected);const value='aurelia-objects:'+JSON.stringify(state.clipboard);try{await navigator.clipboard.writeText(value);}catch{}toast(`${state.clipboard.length} object${state.clipboard.length===1?'':'s'} copied.`);
}
async function pasteSelection(){
 let values=state.clipboard;try{const text=await navigator.clipboard.readText();if(text.startsWith('aurelia-objects:')){const raw=JSON.parse(text.slice(16)),temp=makeDocument();temp.slides[0].elements=raw;values=validateDocument(temp).slides[0].elements;}else if(text&&!values){const e=makeElement('text',{text:text.slice(0,100000),x:160,y:180,w:900,h:320,fontSize:32});store.add(e);return;}}catch{}
 if(!values?.length){toast('Copy an object first, or paste text or an image directly into the canvas.');return;}
 const groups=new Map(),els=clone(values).map(e=>{e.id=uid();e.x+=28;e.y+=28;e.locked=false;if(e.groupId){if(!groups.has(e.groupId))groups.set(e.groupId,uid('g'));e.groupId=groups.get(e.groupId);}return e;});store.transaction('Paste objects',()=>store.slide.elements.push(...els));store.select(els.map(e=>e.id));
}
let replaceImageId=null;
async function insertImageFile(file){
 if(!/^image\/(png|jpeg|webp|gif)$/.test(file.type))throw new Error('Choose a PNG, JPEG, WebP or GIF picture.');if(file.size>22*1024*1024)throw new Error('Pictures are limited to 22 MB.');
 const src=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result);reader.onerror=()=>reject(reader.error);reader.readAsDataURL(file);});const image=new Image();image.src=src;await image.decode();
 const downscale=Math.min(1,4096/image.naturalWidth,4096/image.naturalHeight),canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(image.naturalWidth*downscale));canvas.height=Math.max(1,Math.round(image.naturalHeight*downscale));canvas.getContext('2d').drawImage(image,0,0,canvas.width,canvas.height);const data=canvas.toDataURL(file.type==='image/jpeg'?'image/jpeg':'image/png',.94);
 if(replaceImageId){const e=store.slide.elements.find(e=>e.id===replaceImageId);if(e)store.transaction('Replace picture',()=>{e.src=data;e.name=file.name;});replaceImageId=null;renderer.invalidateTiles();return;}
 const sc=Math.min(store.doc.width*.7/canvas.width,store.doc.height*.7/canvas.height,1),w=canvas.width*sc,h=canvas.height*sc;store.add(makeElement('image',{name:file.name,src:data,fit:'cover',x:(store.doc.width-w)/2,y:(store.doc.height-h)/2,w,h}));
}
$('#imageFile').addEventListener('change',safeRun(async event=>{for(const f of event.target.files)await insertImageFile(f);event.target.value='';}));
$('#openFile').addEventListener('change',safeRun(async event=>{const file=event.target.files[0];event.target.value='';if(!file)return;if(file.size>100*1024*1024)throw new Error('Presentation files are limited to 100 MB.');finishTextEditing();let doc,warnings=[];
 if(file.name.toLowerCase().endsWith('.pptx')){toast('Reading PowerPoint presentation…');({doc,warnings}=await importPPTX(await file.arrayBuffer()));}else doc=validateDocument(JSON.parse(await file.text()));store.replace(doc);resizeStage();renderInspector();toast(`Opened ${doc.slides.length} slides.`);if(warnings.length)openModal('PowerPoint import report',`<p class="modal-description">The supported content is now editable. Check the slides for these known import limitations:</p>${warnings.map(w=>`<p>${escapeHTML(w)}</p>`).join('')}`,[{label:'Continue editing',primary:true,action:()=>{}}]);}));
viewport.addEventListener('dragover',event=>{if(event.dataTransfer.types.includes('Files')){event.preventDefault();event.dataTransfer.dropEffect='copy';}});
viewport.addEventListener('drop',safeRun(async event=>{event.preventDefault();for(const file of event.dataTransfer.files)if(file.type.startsWith('image/'))await insertImageFile(file);}));
document.addEventListener('paste',safeRun(async event=>{if(['INPUT','TEXTAREA'].includes(document.activeElement?.tagName)||state.presenting||modal.open)return;const image=[...event.clipboardData.items].find(i=>i.kind==='file'&&i.type.startsWith('image/'));if(image){event.preventDefault();await insertImageFile(image.getAsFile());return;}const text=event.clipboardData.getData('text/plain');if(text.startsWith('aurelia-objects:')){event.preventDefault();const temp=makeDocument();temp.slides[0].elements=JSON.parse(text.slice(16));state.clipboard=validateDocument(temp).slides[0].elements;await pasteSelection();}else if(text){event.preventDefault();store.add(makeElement('text',{text:text.slice(0,100000),x:140,y:160,w:900,h:300,fontSize:32}));}}));
function insertDiagram(){
 const groupId=uid('g'),x=store.doc.width*.085,y=store.doc.height*.4,w=store.doc.width*.225,h=142,gap=store.doc.width*.078,els=[];
 ['Discover','Create','Deliver'].forEach((label,i)=>{const xx=x+i*(w+gap);els.push(makeElement('roundRect',{x:xx,y,w,h,fill:i===1?'@accent':'@dark',radius:16,groupId,name:label+' step'}));els.push(makeElement('text',{x:xx+18,y:y+48,w:w-36,h:60,text:label,fontSize:36,bold:true,align:'center',fill:'@light',groupId}));if(i<2)els.push(makeElement('arrow',{x:xx+w+15,y:y+h/2-12,w:gap-30,h:24,fill:'@secondary',groupId}));});store.transaction('Insert process diagram',()=>store.slide.elements.push(...els));store.select(els.map(e=>e.id));
}
function editData(e){
 if(e.locked){toast('Unlock this object before editing its data.');return;}
 if(e.type==='table'){
  openModal('Edit table',`<p class="modal-description">Each line is a row. Separate columns with a tab. You can paste a range from a spreadsheet. The first row is styled as the header.</p><textarea id="tableData" class="code-textarea" spellcheck="false" aria-label="Table data">${escapeHTML(e.cells.map(r=>r.join('\t')).join('\n'))}</textarea>`,[{label:'Cancel',action:()=>{}},{label:'Apply table',primary:true,action:()=>{const text=$('#tableData').value.trim(),cells=text.split('\n').map(r=>r.split('\t'));if(!text||cells.length>30||Math.max(...cells.map(r=>r.length))>12)throw new Error('Use 1–30 rows and 1–12 columns.');const columns=Math.max(...cells.map(r=>r.length));for(const r of cells)while(r.length<columns)r.push('');store.transaction('Edit table data',()=>e.cells=cells);}}],640);
  $('#tableData').addEventListener('keydown',event=>{if(event.key==='Tab'){event.preventDefault();const el=event.target,start=el.selectionStart;el.setRangeText('\t',start,el.selectionEnd,'end');}});return;
 }
 let labels=clone(e.labels),values=clone(e.values);
 const rows=()=>labels.map((label,i)=>`<tr><td><input data-label-index="${i}" aria-label="Category ${i+1}" value="${escapeHTML(label)}" maxlength="80"></td><td><input data-value-index="${i}" aria-label="Value ${i+1}" type="number" min="0" value="${values[i]}"></td><td><button data-remove-row="${i}" title="Remove row" aria-label="Remove row">${icon('trash',15)}</button></td></tr>`).join('');
 openModal('Edit chart data',`<p class="modal-description">Real data, live charts. Values must be non-negative. Choose the visual in the properties panel.</p><table class="data-table"><thead><tr><th>Category</th><th>Value</th><th></th></tr></thead><tbody id="chartRows">${rows()}</tbody></table><button id="addChartRow" class="secondary">${icon('plus',15)}Add category</button><label style="display:flex;gap:7px;align-items:center;margin-top:15px;font-size:11px"><input type="checkbox" id="chartShowValues"${e.showValues!==false?' checked':''}>Show value labels</label>`,[{label:'Cancel',action:()=>{}},{label:'Apply data',primary:true,action:()=>{readRows();if(!labels.length)throw new Error('A chart needs at least one category.');if(values.some(v=>!Number.isFinite(v)||v<0))throw new Error('Enter a non-negative number for every category.');const show=$('#chartShowValues').checked;store.transaction('Edit chart data',()=>{e.labels=labels;e.values=values;e.showValues=show;});}}]);
 function readRows(){labels=$$('[data-label-index]',modal).map(i=>i.value);values=$$('[data-value-index]',modal).map(i=>+i.value);}
 $('#chartRows').addEventListener('click',event=>{const b=event.target.closest('[data-remove-row]');if(b&&labels.length>1){readRows();const index=+b.dataset.removeRow;labels.splice(index,1);values.splice(index,1);$('#chartRows').innerHTML=rows();}});
 $('#addChartRow').addEventListener('click',()=>{readRows();if(labels.length>=30){toast('Charts support up to 30 categories.');return;}labels.push('New');values.push(50);$('#chartRows').innerHTML=rows();});
}
function showSlideSize(){
 openModal('Presentation size',`<p class="modal-description">Slide dimensions use a 96-pixel-per-inch design coordinate system. Rescale content to keep your layout proportional.</p><div class="modal-form"><label>Preset<select id="sizePreset"><option value="1280,720">Widescreen · 16:9</option><option value="1200,900">Standard · 4:3</option><option value="1080,1080">Square · 1:1</option><option value="1080,1920">Portrait · 9:16</option><option value="custom" selected>Custom</option></select></label><div class="property-grid"><label>Width<input id="deckWidth" type="number" min="320" max="4096" value="${store.doc.width}"></label><label>Height<input id="deckHeight" type="number" min="240" max="4096" value="${store.doc.height}"></label></div><label style="display:flex;flex-direction:row;align-items:center"><input id="scaleContent" type="checkbox" checked>Rescale all slide content</label></div>`,[{label:'Cancel',action:()=>{}},{label:'Apply size',primary:true,action:()=>{const w=clamp(+$('#deckWidth').value||1280,320,4096),h=clamp(+$('#deckHeight').value||720,240,4096),sx=w/store.doc.width,sy=h/store.doc.height,scale=$('#scaleContent').checked;store.transaction('Resize presentation',()=>{if(scale)for(const s of store.doc.slides)for(const e of s.elements){e.x*=sx;e.y*=sy;e.w*=sx;e.h*=sy;if(e.fontSize)e.fontSize*=Math.min(sx,sy);e.radius*=Math.min(sx,sy);}store.doc.width=w;store.doc.height=h;});zoomTo('fit');}}]);
 $('#sizePreset').addEventListener('change',event=>{if(event.target.value!=='custom'){const [w,h]=event.target.value.split(',');$('#deckWidth').value=w;$('#deckHeight').value=h;}});
}
function showSorter(){
 openModal('Slide sorter',`<p class="modal-description">Choose a slide to edit it. Reorder slides by dragging their thumbnails in the left panel.</p><div class="sorter-grid">${store.doc.slides.map((s,i)=>`<button class="sorter-card ${s.id===store.active?'active':''}" data-sorter-id="${s.id}"><canvas width="480" height="${Math.round(480*store.doc.height/store.doc.width)}"></canvas><span>${String(i+1).padStart(2,'0')} &nbsp; ${escapeHTML(s.name)}</span></button>`).join('')}</div>`,[],920);
 $$('[data-sorter-id]',modal).forEach(b=>{const s=store.doc.slides.find(s=>s.id===b.dataset.sorterId);paintSlide2D($('canvas',b),s,store.doc,thumbnailsImages,{scale:480/store.doc.width});b.addEventListener('click',()=>{store.activate(s.id);modal.close();renderInspector();});});
}
function showHelp(){
 const keys=[['Undo / redo','Ctrl/⌘ Z · Ctrl/⌘ Shift Z'],['Save / open presentation','Ctrl/⌘ S · Ctrl/⌘ O'],['Find commands','Ctrl/⌘ K'],['Select all objects','Ctrl/⌘ A'],['Copy / paste / duplicate','Ctrl/⌘ C · V · D'],['Group / ungroup','Ctrl/⌘ G · Ctrl/⌘ Shift G'],['New slide','Ctrl/⌘ M'],['Edit text','Double-click a text box'],['Finish / cancel text edit','Ctrl/⌘ Enter · Escape'],['Select multiple objects','Shift-click or drag a marquee'],['Move with precision','Arrow keys · Shift = 10 px'],['Constrain resize / rotation','Hold Shift'],['Duplicate while dragging','Hold Alt/Option before dragging'],['Disable snapping for one drag','Hold Alt/Option'],['Zoom around pointer','Ctrl/⌘ + mouse wheel'],['Pan the workspace','Space + drag · middle drag'],['Present / present current','F5 · Shift F5'],['Presentation navigation','← → · Space · Esc'],['Laser / blackout','L · B']];
 openModal('Make yourself at home',`<p class="modal-description">A familiar workspace for your next great story. All editing stays on this device. Export a file for an independent backup.</p><table class="help-table">${keys.map(([a,b])=>`<tr><td>${a}</td><td>${b}</td></tr>`).join('')}</table>`,[{label:'Start creating',primary:true,action:()=>{}}],630);
}
function showAbout(){openModal('Ideas in motion.',`<div class="about-logo"><span class="brand-mark">${icon('logo',36)}</span><div><strong>Aurelia Slides</strong><span class="version">Version 1.0 · A local-first creative workspace</span></div></div><p>A presentation editor written in plain JavaScript, with a real WebGPU scene renderer and a Canvas 2D compatibility path.</p><div class="feature-chips"><span>No framework</span><span>No account</span><span>No network dependencies</span><span>Editable PowerPoint export</span></div><p class="modal-description">An independent application inspired by familiar presentation workflows. It is not affiliated with Microsoft. This version supports editable text boxes, shapes, images, single-series charts, simple tables, slide organization, themes, transitions, and click-triggered entrance effects.</p><p class="modal-description">PowerPoint import is best-effort, not complete Office compatibility. No mixed-style text runs, media playback, comments, collaboration, SmartArt, macros, native Office chart embedding, or arbitrary master/layout inheritance. Charts and tables export as editable drawing objects; donut charts export as images. Aurelia-authored PPTX files also include a lossless native project.</p>`,[{label:'Keyboard shortcuts',action:()=>{showHelp();return false;}},{label:'Back to creating',primary:true,action:()=>{}}],620);}
function showDiagnostics(){
 const stats=renderer.lastStats;openModal('Rendering engine',`<p class="modal-description">${renderer.mode==='webgpu'?'WebGPU · direct instanced drawing with analytic shape coverage and cached text/image tiles.':'Canvas 2D compatibility renderer is active.'}</p><div class="diagnostic-grid"><div><strong>${renderer.mode==='webgpu'?'GPU':'2D'}</strong><span>Active renderer</span></div><div><strong>${stats.ms.toFixed(2)} ms</strong><span>Last CPU render / submission time</span></div><div><strong>${stats.objects}</strong><span>Flattened visible objects</span></div><div><strong>${stats.draws}</strong><span>Draw calls in last frame</span></div><div><strong>${(renderer.cacheBytes/1048576).toFixed(1)} MB</strong><span>GPU tile cache (96 MB soft limit)</span></div><div><strong>${store.past.length}</strong><span>Undo transactions retained</span></div></div><p class="modal-description" style="margin-top:16px">${escapeHTML(state.rendererReason)}<br>Render-on-demand; no continuous editing-frame loop. Timing above is CPU preparation and submission time, not measured GPU execution time or an FPS benchmark.</p>`,[{label:'Close',primary:true,action:()=>{}}],550);
}
const COMMANDS=[['New slide','new-slide','newslide','Ctrl M'],['Insert text box','insert-text','text'],['Insert picture','insert-image','image'],['Insert chart','insert-chart','chart'],['Insert table','insert-table','table'],['Insert process diagram','insert-diagram','group'],['Draw rectangle','draw:rect','rect'],['Draw ellipse','draw:ellipse','ellipse'],['Draw arrow','draw:arrow','arrow'],['Save Aurelia presentation','save','save','Ctrl S'],['Open presentation','open','folder','Ctrl O'],['Export PowerPoint','export-pptx','export'],['Export PNG image','export-png','image'],['Export vector SVG','export-svg','shapes'],['Print / Save as PDF','print','slide'],['Undo','undo','undo','Ctrl Z'],['Redo','redo','redo','Ctrl Shift Z'],['Duplicate selected objects','duplicate','copy','Ctrl D'],['Duplicate current slide','duplicate-slide','newslide'],['Delete current slide','delete-slide','trash'],['Select all objects','select-all','pointer','Ctrl A'],['Bring to front','front','front'],['Send to back','back','back'],['Align center','align:center','aligncenter'],['Align middle','align:middle','align'],['Group objects','group','group','Ctrl G'],['Ungroup objects','ungroup','shapes','Ctrl Shift G'],['Slide sorter','sorter','grid'],['Present from beginning','present-first','play','F5'],['Present from current slide','present','screen','Shift F5'],['Presenter notes view','present-notes','notes'],['Toggle grid','toggle-grid','grid'],['Toggle snap guides','toggle-snap','align'],['Toggle notes','toggle-notes','notes'],['Toggle appearance','toggle-dark','moon'],['Fit slide to window','zoom-fit','fit'],['Slide dimensions','slide-size','fit'],['Rendering diagnostics','diagnostics','code'],['Keyboard shortcuts','help','keyboard']];
function showCommands(){
 openModal('Find a command',`<input id="commandSearch" class="command-search" type="search" placeholder="Try “insert chart” or “export”…" aria-label="Search commands"><div id="commandResults"></div>`,[],550);
 const draw=()=>{const query=$('#commandSearch').value.trim().toLowerCase(),items=COMMANDS.filter(c=>c[0].toLowerCase().includes(query));$('#commandResults').innerHTML=items.map(c=>`<button class="command-result" data-command-result="${c[1]}">${icon(c[2])}<span>${c[0]}</span>${c[3]?`<kbd>${c[3]}</kbd>`:''}</button>`).join('')||'<p class="modal-description">No matching commands.</p>';};draw();$('#commandSearch').addEventListener('input',draw);$('#commandSearch').addEventListener('keydown',event=>{if(event.key==='Enter'){const b=$('[data-command-result]',modal);if(b){event.preventDefault();modal.close();safeRun(command)(b.dataset.commandResult);}}if(event.key==='ArrowDown'){event.preventDefault();$('[data-command-result]',modal)?.focus();}});$('#commandResults').addEventListener('click',safeRun(async event=>{const b=event.target.closest('[data-command-result]');if(b){modal.close();await command(b.dataset.commandResult);}}));setTimeout(()=>$('#commandSearch')?.focus(),40);
}
async function rasterizeSlide(slide,scale=2){const canvas=document.createElement('canvas');canvas.width=Math.round(store.doc.width*scale);canvas.height=Math.round(store.doc.height*scale);await renderer.images.ready(slide.elements);paintSlide2D(canvas,slide,store.doc,renderer.images,{scale});return canvas;}
async function exportPNG(){const canvas=await rasterizeSlide(store.slide);const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));downloadBlob(blob,fileName(`slide-${store.doc.slides.indexOf(store.slide)+1}.png`));toast('Slide exported at 2× resolution.');}
async function exportAllImages(){toast('Rendering slide images…');const files={};for(const [i,s]of store.doc.slides.entries()){const c=await rasterizeSlide(s);const b=await new Promise(resolve=>c.toBlob(resolve,'image/png'));files[`Slide-${String(i+1).padStart(3,'0')}.png`]=new Uint8Array(await b.arrayBuffer());}downloadBlob(zipStore(files),fileName('slides.zip'));toast(`Exported ${store.doc.slides.length} PNG images.`);}
async function exportSVG(){
 const doc=store.doc,s=store.slide,ctx=document.createElement('canvas').getContext('2d');let markup=`<rect width="${doc.width}" height="${doc.height}" fill="${color(s.bg,doc.theme)}"/>`,idx=0;
 for(const e of flattenScene(s.elements)){
  const clip='clip'+idx++,transform=`translate(${e.x+e.w/2} ${e.y+e.h/2}) rotate(${e.rotation||0}) translate(${-e.w/2} ${-e.h/2})`,style=`fill="${color(e.fill,doc.theme)}" stroke="${e.strokeWidth?color(e.stroke,doc.theme):'none'}" stroke-width="${e.strokeWidth||0}"`;let body='';
  if(e.type==='text'){
   ctx.font=fontCSS(e);const lines=wrapText(ctx,e.text,e.w-2*(e.padding||0)),lh=e.fontSize*(e.lineHeight||1.2),total=lines.length*lh,p=e.padding||0,x=e.align==='center'?e.w/2:e.align==='right'?e.w-p:p,top=e.valign==='middle'?Math.max(p,(e.h-total)/2):e.valign==='bottom'?Math.max(p,e.h-total-p):p;
   body=`<text fill="${color(e.fill,doc.theme)}" font-family="${escapeHTML(e.fontFamily)}" font-size="${e.fontSize}" font-weight="${e.bold?'700':'400'}" font-style="${e.italic?'italic':'normal'}" text-decoration="${e.underline?'underline':'none'}" text-anchor="${e.align==='center'?'middle':e.align==='right'?'end':'start'}">${lines.map((line,i)=>`<tspan x="${x}" y="${top+i*lh+e.fontSize*.85}" xml:space="preserve">${escapeHTML(line)}</tspan>`).join('')}</text>`;
  }else if(e.type==='image')body=`<image href="${e.src}" width="${e.w}" height="${e.h}" preserveAspectRatio="xMidYMid ${e.fit==='contain'?'meet':'slice'}"/>`;
  else if(e.type==='chart'){const c=document.createElement('canvas');c.width=Math.ceil(e.w*2);c.height=Math.ceil(e.h*2);paintSlide2D(c,{bg:s.bg,elements:[{...e,x:0,y:0,rotation:0,opacity:1}]},{...doc,width:e.w,height:e.h},renderer.images,{scale:2});body=`<image href="${c.toDataURL('image/png')}" width="${e.w}" height="${e.h}"/>`;}
  else if(e.type==='ellipse')body=`<ellipse cx="${e.w/2}" cy="${e.h/2}" rx="${e.w/2}" ry="${e.h/2}" ${style}/>`;
  else if(['rect','roundRect','line'].includes(e.type))body=`<rect width="${e.w}" height="${e.h}" rx="${e.type==='roundRect'?Math.min(e.radius,e.w/2,e.h/2):0}" ${style}/>`;
  else{let points;if(e.type==='triangle')points=[[e.w/2,0],[e.w,e.h],[0,e.h]];if(e.type==='diamond')points=[[e.w/2,0],[e.w,e.h/2],[e.w/2,e.h],[0,e.h/2]];if(e.type==='arrow')points=[[0,e.h*.32],[e.w*.6,e.h*.32],[e.w*.6,0],[e.w,e.h/2],[e.w*.6,e.h],[e.w*.6,e.h*.68],[0,e.h*.68]];if(e.type==='star')points=Array.from({length:10},(_,i)=>{const a=i*Math.PI/5-Math.PI/2,r=i%2?.44:1;return[e.w/2+Math.cos(a)*e.w/2*r,e.h/2+Math.sin(a)*e.h/2*r];});body=`<polygon points="${points.map(p=>p.join(',')).join(' ')}" ${style}/>`;}
  markup+=`<g transform="${transform}" opacity="${e.opacity??1}"><defs><clipPath id="${clip}"><rect width="${e.w}" height="${e.h}"/></clipPath></defs><g clip-path="url(#${clip})">${body}</g></g>`;
 }
 downloadBlob(new Blob([`<?xml version="1.0" encoding="utf-8"?><svg xmlns="http://www.w3.org/2000/svg" width="${doc.width}" height="${doc.height}" viewBox="0 0 ${doc.width} ${doc.height}"><title>${escapeHTML(doc.title)}</title>${markup}</svg>`],{type:'image/svg+xml'}),fileName('svg'));toast('Editable vector SVG exported.');
}
async function printDeck(){
 const win=window.open('','_blank');if(!win)throw new Error('Allow pop-up windows to print the presentation.');win.document.write('<!doctype html><title>Preparing presentation</title><body style="font:16px system-ui;padding:40px">Preparing slides for printing…</body>');let images='';for(const s of store.doc.slides){const c=await rasterizeSlide(s);images+=`<section><img src="${c.toDataURL('image/png')}" alt="${escapeHTML(s.name)}"></section>`;}
 win.document.open();win.document.write(`<!doctype html><html><head><title>${escapeHTML(store.doc.title)}</title><style>@page{size:${store.doc.width/96}in ${store.doc.height/96}in;margin:0}*{box-sizing:border-box}body{margin:0;background:#e8e8e5;font:14px system-ui}section{break-after:page;page-break-after:always;width:${store.doc.width/96}in;height:${store.doc.height/96}in;margin:20px auto;background:white}section:last-child{break-after:auto}img{display:block;width:100%;height:100%}header{padding:20px;text-align:center}button{padding:10px 20px;cursor:pointer}@media print{header{display:none}body{background:white}section{margin:0}}</style></head><body><header><button onclick="window.print()">Print presentation / Save as PDF</button><p>Choose landscape, no margins, and background graphics. Slides are rendered images.</p></header>${images}</body></html>`);win.document.close();
}
async function command(cmd,anchor){
 if(state.editing&&!['copy','paste'].includes(cmd))finishTextEditing();
 if(cmd.startsWith('tab:')){state.tab=cmd.slice(4);renderRibbon();return;}
 if(cmd.startsWith('panel:')){state.panel=cmd.slice(6);state.showInspector=true;renderInspector();resizeStage();return;}
 if(cmd.startsWith('draw:')){setTool(cmd.slice(5));return;}
 if(cmd.startsWith('theme:')){store.transaction('Apply theme',()=>store.doc.theme=cmd.slice(6));renderInspector();return;}
 if(cmd.startsWith('new-layout:')){store.addSlide(cmd.slice(11));renderInspector();return;}
 if(cmd.startsWith('text-align:')){formatText({align:cmd.slice(11)});return;}
 if(cmd.startsWith('align:')){store.align(cmd.slice(6));return;}
 if(cmd.startsWith('distribute:')){if(store.selected.length<3)toast('Select at least three objects to distribute.');else store.distribute(cmd.slice(11));return;}
 if(cmd.startsWith('transition:')){store.transaction('Set transition',()=>store.slide.transition=cmd.slice(11));return;}
 if(cmd.startsWith('animation:')){if(!store.selected.length)toast('Select one or more objects to animate.');else store.updateSelected({animation:cmd.slice(10)},'Set entrance animation');return;}
 switch(cmd){
  case'file':showFileMenu(anchor);break;
  case'new-document':askReplace('Start a new presentation?',()=>store.replace(makeDocument()));break;
  case'restore-demo':askReplace('Restore the example deck?',()=>store.replace(createDemo()));break;
  case'open':$('#openFile').click();break;
  case'save':nativeSave();break;
  case'undo':store.undo();renderInspector();break;
  case'redo':store.redo();renderInspector();break;
  case'copy':await copySelection();break;
  case'paste':await pasteSelection();break;
  case'duplicate':store.duplicateSelected();break;
  case'delete':store.deleteSelected();break;
  case'select-all':store.select(store.slide.elements.filter(e=>!e.hidden).map(e=>e.id));break;
  case'new-slide':store.addSlide('content');renderInspector();break;
  case'duplicate-slide':store.duplicateSlide();renderInspector();break;
  case'delete-slide':if(!store.deleteSlide())toast('Keep at least one slide in your presentation.');renderInspector();break;
  case'rename-slide':openModal('Rename slide',`<div class="modal-form"><label>Slide name<input id="slideNameInput" maxlength="200" value="${escapeHTML(store.slide.name)}"></label></div>`,[{label:'Cancel',action:()=>{}},{label:'Rename',primary:true,action:()=>store.transaction('Rename slide',()=>store.slide.name=$('#slideNameInput').value.trim()||'Untitled slide')}]);break;
  case'layout-menu':showLayoutMenu(anchor);break;
  case'reset-layout':openModal('Reset current slide layout',`<p class="modal-description">This replaces the objects on the current slide. The change can be undone.</p><div class="modal-form"><label>Layout<select id="resetLayout"><option value="title">Title slide</option><option value="content">Title and content</option><option value="split">Two columns</option><option value="quote">Quote</option><option value="blank">Blank</option></select></label></div>`,[{label:'Cancel',action:()=>{}},{label:'Reset layout',primary:true,action:()=>{store.transaction('Reset slide layout',()=>store.slide.elements=makeSlide($('#resetLayout').value).elements);store.select([]);}}]);break;
  case'insert-text':{const e=makeElement('text',{x:store.doc.width*.14,y:store.doc.height*.35,w:store.doc.width*.7,h:100,text:'Make your point.',fontSize:48});store.add(e);startTextEditing(e);break;}
  case'insert-image':replaceImageId=null;$('#imageFile').click();break;
  case'replace-image':if(store.selected[0]?.type==='image'){replaceImageId=store.selected[0].id;$('#imageFile').click();}break;
  case'insert-chart':{const e=makeElement('chart',{x:store.doc.width*.15,y:store.doc.height*.2,w:store.doc.width*.7,h:store.doc.height*.6,chartType:'bar',labels:['Q1','Q2','Q3','Q4'],values:[28,46,57,84],showValues:true,fill:'@accent'});store.add(e);break;}
  case'insert-table':{const e=makeElement('table',{x:store.doc.width*.12,y:store.doc.height*.24,w:store.doc.width*.76,h:store.doc.height*.46,fill:'@dark',fontSize:24,cells:[['Initiative','Owner','Status'],['Discover','Design team','In progress'],['Build','Engineering','Planned'],['Launch','Everyone','Up next']]});store.add(e);break;}
  case'insert-diagram':insertDiagram();break;
  case'insert-number':store.add(makeElement('text',{x:store.doc.width-110,y:store.doc.height-60,w:70,h:35,fontSize:18,align:'right',text:String(store.doc.slides.indexOf(store.slide)+1),fill:'@muted',name:'Slide number (static)'}));break;
  case'insert-date':store.add(makeElement('text',{x:75,y:store.doc.height-60,w:450,h:35,fontSize:18,text:new Date().toLocaleDateString(undefined,{dateStyle:'long'}),fill:'@muted',name:'Date (static)'}));break;
  case'edit-data':if(store.selected[0]&&['chart','table'].includes(store.selected[0].type))editData(store.selected[0]);break;
  case'edit-text':startTextEditing(store.selected.find(e=>e.type==='text'));break;
  case'shapes-menu':showShapeMenu(anchor);break;
  case'arrange-menu':showArrangeMenu(anchor);break;
  case'front':case'back':case'forward':case'backward':store.arrange(cmd);break;
  case'group':if(store.selected.length<2)toast('Select two or more objects to group.');else store.group();break;
  case'ungroup':store.ungroup();break;
  case'lock':{const lock=!store.selected.every(e=>e.locked);store.transaction(lock?'Lock objects':'Unlock objects',()=>store.selected.forEach(e=>e.locked=lock));renderInspector();break;}
  case'bold':case'italic':case'underline':formatText({[cmd]:!selectedText().every(e=>e[cmd])});break;
  case'font-grow':case'font-shrink':if(selectedText().length)store.transaction('Change font size',()=>selectedText().forEach(e=>e.fontSize=clamp(e.fontSize+(cmd==='font-grow'?2:-2),4,600)));break;
  case'bullets':if(selectedText().length){const remove=selectedText().every(e=>e.text.split('\n').every(l=>!l.trim()||l.startsWith('• ')));store.transaction('Toggle bullets',()=>selectedText().forEach(e=>e.text=e.text.split('\n').map(l=>!l.trim()?l:remove?l.replace(/^• /,''):l.startsWith('• ')?l:'• '+l).join('\n')));}break;
  case'no-fill':store.updateSelected({fill:'none'},'Remove fill');break;
  case'background':state.panel='format';state.showInspector=true;store.select([]);renderInspector();resizeStage();setTimeout(()=>$('[data-slide-prop="bg"]',$('#inspector'))?.click(),0);break;
  case'slide-size':showSlideSize();break;
  case'apply-transition':{const {transition,duration}=store.slide;store.transaction('Apply transition to all slides',()=>store.doc.slides.forEach(s=>Object.assign(s,{transition,duration})));toast('Transition applied to all slides.');break;}
  case'preview-transition':animateTransition(stage,store.slide);break;
  case'toggle-grid':state.showGrid=!state.showGrid;$('#gridButton').classList.toggle('active',state.showGrid);renderRibbon();invalidate();break;
  case'toggle-snap':state.snap=!state.snap;toast(`Smart alignment guides ${state.snap?'on':'off'}.`);renderRibbon();break;
  case'toggle-notes':state.showNotes=!state.showNotes;$('#notesPanel').hidden=!state.showNotes;renderRibbon();resizeStage();break;
  case'toggle-inspector':state.showInspector=!state.showInspector;renderInspector();resizeStage();break;
  case'toggle-dark':document.body.classList.toggle('dark');try{localStorage.setItem('aurelia-dark',document.body.classList.contains('dark')?'1':'0');}catch{}break;
  case'normal':if(modal.open)modal.close();state.showInspector=true;state.showNotes=true;$('#notesPanel').hidden=false;renderInspector();zoomTo('fit');break;
  case'sorter':showSorter();break;
  case'zoom-fit':zoomTo('fit');break;
  case'zoom-100':zoomTo(100);break;
  case'zoom-in':zoomTo(state.scale*100+10);break;
  case'zoom-out':zoomTo(state.scale*100-10);break;
  case'export-pptx':toast('Building editable PowerPoint file…');downloadBlob(await exportPPTX(store.doc),fileName('pptx'));toast('PowerPoint file exported. Text and shapes remain editable.');break;
  case'export-png':await exportPNG();break;
  case'export-svg':await exportSVG();break;
  case'export-images':await exportAllImages();break;
  case'print':await printDeck();break;
  case'present':await startPresentation(store.doc.slides.indexOf(store.slide));break;
  case'present-first':await startPresentation(0);break;
  case'present-notes':await startPresentation(store.doc.slides.indexOf(store.slide),true);break;
  case'present-rehearse':await startPresentation(0,true);break;
  case'commands':showCommands();break;
  case'help':showHelp();break;
  case'about':showAbout();break;
  case'diagnostics':showDiagnostics();break;
  default:console.warn('Unknown command:',cmd);
 }
}
let presentationRenderer=null,presentationIndex=0,presentationStep=0,presentationStart=0,presentationAnim=null,presentationRAF=0,presentationTimerId=0,laserEnabled=false,numberBuffer='';
function animateTransition(element,slide){
 if(window.matchMedia('(prefers-reduced-motion: reduce)').matches||slide.transition==='none')return;
 const first=slide.transition==='push'?{opacity:.2,transform:'translateX(8%)'}:slide.transition==='zoom'?{opacity:0,transform:'scale(.9)'}:{opacity:0};
 element.animate([first,{opacity:1,transform:'none'}],{duration:(slide.duration||.45)*1000,easing:'cubic-bezier(.2,.8,.2,1)'});
}
function animationEntries(slide){return slide.elements.filter(e=>!e.hidden&&e.animation!=='none');}
function presentationResize(){
 if(!state.presenting)return;const notes=!$('#presenterNotes').hidden,width=window.innerWidth-(notes?300:0),height=window.innerHeight,scale=Math.min(width/store.doc.width,height/store.doc.height);Object.assign($('#presentStage').style,{width:store.doc.width*scale+'px',height:store.doc.height*scale+'px'});drawPresentation();
}
function drawPresentation(now=performance.now()){
 if(!state.presenting||!presentationRenderer||presentationRenderer.mode==='initializing')return;
 const slide=store.doc.slides[presentationIndex],entries=animationEntries(slide),positions=new Map(entries.map((e,i)=>[e.id,i]));let running=false;
 const elements=slide.elements.map(e=>{
  const idx=positions.get(e.id);if(idx===undefined)return e;if(idx>=presentationStep)return{...e,hidden:true};
  if(presentationAnim&&idx===presentationStep-1){const p=clamp((now-presentationAnim.start)/500,0,1),t=1-(1-p)**3;running=p<1;if(p<1){if(e.animation==='rise')return{...e,y:e.y+(1-t)*45,opacity:(e.opacity??1)*t};if(e.animation==='zoom'){const scale=.75+.25*t;return{...e,x:e.x+e.w*(1-scale)/2,y:e.y+e.h*(1-scale)/2,w:e.w*scale,h:e.h*scale,opacity:(e.opacity??1)*t};}return{...e,opacity:(e.opacity??1)*t};}}
  return e;
 });
 const b=$('#presentStage').getBoundingClientRect();presentationRenderer.render({...slide,elements},store.doc,{cssWidth:b.width,cssHeight:b.height});
 if(running){cancelAnimationFrame(presentationRAF);presentationRAF=requestAnimationFrame(drawPresentation);}else presentationAnim=null;
}
function updatePresenterUI(){
 const s=store.doc.slides[presentationIndex];$('#presentCounter').textContent=`${presentationIndex+1} / ${store.doc.slides.length}`;$('#presentNotesText').textContent=s.notes||'No speaker notes for this slide.';const next=store.doc.slides[presentationIndex+1]||s,c=$('#nextSlidePreview');c.width=440;c.height=Math.round(440*store.doc.height/store.doc.width);paintSlide2D(c,next,store.doc,thumbnailsImages,{scale:440/store.doc.width});
}
async function startPresentation(index=0,notes=false){
 finishTextEditing();clearTool();closeMenu();if(modal.open)modal.close();presentationIndex=clamp(index,0,store.doc.slides.length-1);presentationStep=0;presentationAnim=null;state.presenting=true;$('#presentOverlay').hidden=false;$('#presenterNotes').hidden=!notes;$('#blackout').hidden=true;$('#laser').hidden=true;laserEnabled=false;presentationStart=performance.now();
 try{const promise=$('#presentOverlay').requestFullscreen?.();promise?.catch(()=>{});}catch{}
 if(!presentationRenderer){presentationRenderer=new SceneRenderer($('#presentCanvas'),()=>{},()=>{if(state.presenting)requestAnimationFrame(drawPresentation);});await presentationRenderer.init();}
 updatePresenterUI();presentationResize();animateTransition($('#presentStage'),store.doc.slides[presentationIndex]);clearInterval(presentationTimerId);presentationTimerId=setInterval(()=>{const seconds=Math.floor((performance.now()-presentationStart)/1000);$('#presentTimer').textContent=`${String(Math.floor(seconds/60)).padStart(2,'0')}:${String(seconds%60).padStart(2,'0')}`;},1000);
}
function advancePresentation(direction=1){
 $('#blackout').hidden=true;const entries=animationEntries(store.doc.slides[presentationIndex]);
 if(direction>0&&presentationStep<entries.length){presentationStep++;presentationAnim={start:performance.now()};drawPresentation();return;}
 if(direction<0&&presentationStep>0){presentationStep--;presentationAnim=null;drawPresentation();return;}
 const next=presentationIndex+direction;if(next<0)return;if(next>=store.doc.slides.length){toast('End of presentation. Press Esc to return to editing.');return;}
 presentationIndex=next;presentationStep=direction<0?animationEntries(store.doc.slides[next]).length:0;presentationAnim=null;updatePresenterUI();drawPresentation();animateTransition($('#presentStage'),store.doc.slides[presentationIndex]);
}
function exitPresentation(){if(!state.presenting)return;state.presenting=false;$('#presentOverlay').hidden=true;cancelAnimationFrame(presentationRAF);clearInterval(presentationTimerId);if(document.fullscreenElement)document.exitFullscreen().catch(()=>{});store.activate(store.doc.slides[presentationIndex].id);viewport.focus({preventScroll:true});resizeStage();}
$('#presentStage').addEventListener('click',()=>advancePresentation(1));
$('#presentStage').addEventListener('pointermove',event=>{if(!laserEnabled)return;const r=$('#presentStage').getBoundingClientRect();$('#laser').style.left=event.clientX-r.left+'px';$('#laser').style.top=event.clientY-r.top+'px';$('#laser').hidden=false;});
$('#presentOverlay').addEventListener('click',event=>{const button=event.target.closest('[data-present]');if(!button)return;switch(button.dataset.present){case'exit':exitPresentation();break;case'prev':advancePresentation(-1);break;case'next':advancePresentation(1);break;case'notes':$('#presenterNotes').hidden=!$('#presenterNotes').hidden;presentationResize();break;case'laser':laserEnabled=!laserEnabled;$('#laser').hidden=!laserEnabled;break;case'black':$('#blackout').hidden=!$('#blackout').hidden;break;}});
document.addEventListener('fullscreenchange',()=>{if(state.presenting&&!document.fullscreenElement)exitPresentation();else presentationResize();});
window.addEventListener('resize',()=>{resizeStage();presentationResize();closeMenu();});
window.addEventListener('blur',()=>{state.space=false;viewport.classList.remove('space-pan');});
function keyTargetIsEditing(event){return event.target.closest('input,textarea,select,[contenteditable="true"]');}
document.addEventListener('keydown',safeRun(async event=>{
 if(state.presenting){
  if(['ArrowRight','ArrowDown','PageDown',' ','Enter'].includes(event.key)){event.preventDefault();if(event.key==='Enter'&&numberBuffer){const index=+numberBuffer-1;numberBuffer='';if(index>=0&&index<store.doc.slides.length){presentationIndex=index;presentationStep=0;updatePresenterUI();drawPresentation();}}else advancePresentation(1);}
  else if(['ArrowLeft','ArrowUp','PageUp','Backspace'].includes(event.key)){event.preventDefault();advancePresentation(-1);}
  else if(event.key==='Escape'){event.preventDefault();exitPresentation();}
  else if(event.key.toLowerCase()==='b')$('#blackout').hidden=!$('#blackout').hidden;
  else if(event.key.toLowerCase()==='l'){laserEnabled=!laserEnabled;$('#laser').hidden=!laserEnabled;}
  else if(event.key==='Home'){presentationIndex=0;presentationStep=0;updatePresenterUI();drawPresentation();}
  else if(event.key==='End'){presentationIndex=store.doc.slides.length-1;presentationStep=0;updatePresenterUI();drawPresentation();}
  else if(/^\d$/.test(event.key))numberBuffer=(numberBuffer+event.key).slice(-3);return;
 }
 if(modal.open||keyTargetIsEditing(event))return;
 const ctrl=event.ctrlKey||event.metaKey,key=event.key.toLowerCase();
 if(ctrl){
  const map={z:event.shiftKey?'redo':'undo',y:'redo',s:'save',o:'open',d:'duplicate',a:'select-all',g:event.shiftKey?'ungroup':'group',m:'new-slide',k:'commands',b:'bold',i:'italic',u:'underline',c:'copy',n:'new-document'};
  if(key==='v'){/* Native paste event is the authoritative source for system clipboard data. */return;}
  if(key==='x'){event.preventDefault();await copySelection();store.deleteSelected();return;}
  if(map[key]){event.preventDefault();await command(map[key]);return;}
 }
 if(event.key==='F5'){event.preventDefault();await command(event.shiftKey?'present':'present-first');return;}
 if(event.key==='F1'){event.preventDefault();showHelp();return;}
 if(event.key==='Escape'){event.preventDefault();cancelPointer();clearTool();closeMenu();store.select([]);return;}
 if(event.key==='Delete'||event.key==='Backspace'){event.preventDefault();if(store.selected.length)store.deleteSelected();else if(document.activeElement?.closest('#thumbnails'))store.deleteSlide();return;}
 if(event.key===' '){event.preventDefault();state.space=true;viewport.classList.add('space-pan');return;}
 if(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(event.key)){
  event.preventDefault();if(store.selected.length){const amount=event.shiftKey?10:1,dx=event.key==='ArrowLeft'?-amount:event.key==='ArrowRight'?amount:0,dy=event.key==='ArrowUp'?-amount:event.key==='ArrowDown'?amount:0;store.transaction('Nudge objects',()=>store.selected.forEach(e=>{if(!e.locked){e.x+=dx;e.y+=dy;}}));}
  else{const i=store.doc.slides.indexOf(store.slide),delta=['ArrowLeft','ArrowUp'].includes(event.key)?-1:1;store.activate(store.doc.slides[clamp(i+delta,0,store.doc.slides.length-1)].id);}return;
 }
 if(event.key==='Enter'&&store.selected.length===1&&store.selected[0].type==='text'){event.preventDefault();startTextEditing(store.selected[0]);}
 if(event.key==='Tab'){event.preventDefault();const els=store.slide.elements.filter(e=>!e.hidden);if(!els.length)return;const current=els.findIndex(e=>store.selection.has(e.id)),next=(current+(event.shiftKey?-1:1)+els.length)%els.length;store.select([els[next].id]);}
}));
document.addEventListener('keyup',event=>{if(event.key===' '){state.space=false;viewport.classList.remove('space-pan');}});
window.addEventListener('beforeunload',event=>{if(store.pending)store.commit();if(store.revision!==lastPersistedRevision&&store.revision>0){persist();event.preventDefault();event.returnValue='';}});
async function initialize(){
 hydrateIcons();try{if(localStorage.getItem('aurelia-dark')==='1')document.body.classList.add('dark');}catch{}
 const saved=await openStorage();if(saved){try{store.doc=validateDocument(saved);store.active=store.doc.slides[0].id;$('#saveText').textContent='Restored from this device';}catch(error){toast('The saved recovery copy could not be read. The example deck is open.',true);}}
 syncUI();renderInspector();resizeStage();await renderer.init();resizeStage();scheduleThumbnails(true);
 new ResizeObserver(()=>resizeStage()).observe(viewport);
 window.Aurelia={version:'1.0.0',store,state,renderer,command,validateDocument,makeElement,exportPPTX,importPPTX,createDemo,startTextEditing,finishTextEditing,resizeStage,renderCount:()=>renderCount,ready:true};
 if(!saved)await persist();lastPersistedRevision=store.revision;
}
initialize().catch(error=>{console.error(error);toast('Initialization failed: '+error.message,true);$('#rendererStatus').textContent='Initialization error';});
