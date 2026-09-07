import {readFile,writeFile,mkdir} from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {resolve,dirname} from 'node:path';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const modules=['core','renderer','demo','icons','pptx','app'];
const parts=await Promise.all(modules.map(async name=>{
  const code=await readFile(resolve(root,`src/${name}.js`),'utf8');
  // The source uses top-level, single-line imports and named declaration exports only.
  const bundled=code.replace(/^import .+;\r?\n/gm,'').replace(/^export /gm,'');
  if(/^\s*(import|export)\s/m.test(bundled))throw new Error(`Unsupported module syntax in ${name}`);
  return `// ─── src/${name}.js ───\n${bundled}`;
}));
const template=await readFile(resolve(root,'dev.html'),'utf8');
const tag='<script type="module" src="./src/app.js"></script>';
if(!template.includes(tag))throw new Error('The development entry point is missing');
const script=`<script id="aurelia-bundle">\n(()=>{\n'use strict';\n\n${parts.join('\n\n')}\n})();\n</script>`;
const output=template.replace(tag,()=>script);
await writeFile(resolve(root,'index.html'),output);
if(process.argv.includes('--site')){
  await mkdir(resolve(root,'_site'),{recursive:true});
  await writeFile(resolve(root,'_site/index.html'),output);
  await writeFile(resolve(root,'_site/.nojekyll'),'');
}
console.log(`Built index.html (${Buffer.byteLength(output)} bytes) from ${modules.length} modules`);
