export const id=()=>crypto.randomUUID().replaceAll('-','');
export const ticks=()=>String(BigInt(Date.now())*10000n+621355968000000000n);
const enc=new TextEncoder(),dec=new TextDecoder('utf-8',{fatal:true});
export function b64(s){let a=enc.encode(s),v='';for(let i=0;i<a.length;i+=8192)v+=String.fromCharCode(...a.subarray(i,i+8192));return btoa(v)}
export async function hash(s){return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',enc.encode(s))),b=>b.toString(16).padStart(2,'0')).join('').toUpperCase()}
export async function serialize(r){let body=['NS2',r.note,r.id,r.parent,r.state,r.pin?'1':'0',r.order,r.time,b64(r.text)].join('\n');return body+'\n'+await hash(body)+'\n'}
export async function parse(raw){let a=raw.replace(/^\uFEFF/,'').split(/\r?\n/);if(a.at(-1)==='')a.pop();if(a.length!==10||a[0]!=='NS2'||await hash(a.slice(0,9).join('\n'))!==a[9])throw Error('Revisão inválida ou formato antigo não suportado');let r={note:a[1],id:a[2],parent:a[3],state:a[4],pin:a[5]==='1',order:a[6],time:a[7],text:dec.decode(Uint8Array.from(atob(a[8]),c=>c.charCodeAt(0)))};if(!/^[a-f0-9]{32}$/i.test(r.note)||! /^[a-f0-9]{32}$/i.test(r.id)||(r.parent&&!/^[a-f0-9]{32}$/i.test(r.parent))||r.parent===r.id||!['O','C','D'].includes(r.state)||!/^\d+$/.test(r.time)||!/^\d+$/.test(r.order))throw Error('Identificador inválido');return r}
export function heads(rs){let p=new Set(rs.map(r=>r.parent));return rs.filter(r=>!p.has(r.id)).sort((a,b)=>Number(b.pin)-Number(a.pin)||(BigInt(a.order)<BigInt(b.order)?-1:BigInt(a.order)>BigInt(b.order)?1:a.id.localeCompare(b.id)))}
export function revision(prev,text){let t=ticks();return {note:prev?.note||id(),id:id(),parent:prev?.id||'',state:prev?.state==='D'?'O':prev?.state||'O',pin:prev?.pin||false,order:prev?.order||t,time:t,text}}
