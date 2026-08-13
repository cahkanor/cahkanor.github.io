export const esc=v=>String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
export const fmtDate=v=>v?new Intl.DateTimeFormat('en-GB',{day:'2-digit',month:'short',year:'numeric'}).format(new Date(`${v}T00:00:00`)):'—';
export const num=v=>new Intl.NumberFormat('en-US').format(v||0);
export const pct=v=>`${Number(v||0).toFixed(1)}%`;
export const validUrl=v=>{try{const u=new URL(v);return ['http:','https:'].includes(u.protocol)}catch{return false}};
export const slug=v=>String(v||'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
export const countBy=(arr,key)=>Object.entries(arr.reduce((a,x)=>{const k=typeof key==='function'?key(x):x[key];if(k)a[k]=(a[k]||0)+1;return a},{})).sort((a,b)=>b[1]-a[1]);
export function download(name,content,type='application/json'){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([content],{type}));a.download=name;a.click();URL.revokeObjectURL(a.href)}
export function csv(rows,cols){const q=v=>`"${String(v??'').replace(/"/g,'""')}"`;return [cols.map(c=>q(c.label)).join(','),...rows.map(r=>cols.map(c=>q(r[c.key])).join(','))].join('\n')}
