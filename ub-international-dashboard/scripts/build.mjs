import fs from 'node:fs/promises';
import path from 'node:path';
const root=path.resolve(import.meta.dirname,'..'),dist=path.join(root,'dist');
await fs.mkdir(dist,{recursive:true});
for(const item of ['index.html','src','public'])await fs.cp(path.join(root,item),path.join(dist,item==='public'?'.':item),{recursive:true});
console.log(`Static build created: ${dist}`);
