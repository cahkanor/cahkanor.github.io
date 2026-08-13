const files=['mous','activities','partners','countries','faculties','summary','data-quality','categories'];
let cache;
export async function loadData(){
  if(cache)return cache;
  const entries=await Promise.all(files.map(async name=>{const r=await fetch(`./data/${name}.json`);if(!r.ok)throw new Error(`${name}.json (${r.status})`);return [name.replace('-','_'),await r.json()]}));
  cache=Object.fromEntries(entries);
  cache.countries.sort((a,b)=>b.agreements-a.agreements||a.country.localeCompare(b.country));
  return cache;
}
