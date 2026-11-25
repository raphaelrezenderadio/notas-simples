import {id,revision,serialize,parse,heads} from './format.js?v=14';
const $=x=>document.getElementById(x), client='38b4d77f-1296-420f-ab3f-e08c53455a8f', redirect='https://raphaelrezenderadio.github.io/notas-simples/', authority='https://login.microsoftonline.com/consumers/oauth2/v2.0/', scope='openid profile offline_access https://graph.microsoft.com/Files.ReadWrite', graph='https://graph.microsoft.com/v1.0';
let mode=localStorage.getItem('ns-mode')||'test', token=null, expiry=0, refresh=null, active=null, running=false, timer, dirty=false;
let diagnosing=false, starting=true;
const sessionKey='ns-ms-session-v1';
function storeSession(){try{sessionStorage.setItem(sessionKey,JSON.stringify({token,expiry,refresh}))}catch{}}
function restoreSession(){try{let x=JSON.parse(sessionStorage.getItem(sessionKey)||'null');if(x){token=x.token||null;expiry=Number(x.expiry)||0;refresh=x.refresh||null}}catch{}}
function clearSession(){token=null;refresh=null;expiry=0;try{sessionStorage.removeItem(sessionKey)}catch{}updateControls()}
restoreSession();
const key=()=> 'ns-mobile-v1-'+mode, load=()=>JSON.parse(localStorage.getItem(key())||'{"rows":[],"pending":[],"draft":null}');let db=load();
function persist(){localStorage.setItem(key(),JSON.stringify(db))}function status(s){$('status').textContent=s}function title(r){return r.text.split(/\r?\n/).find(s=>s.trim())?.trim()||'Nova nota'}
function render(){ $('mode').textContent=mode==='test'?'Teste':'Real';let hs=heads(db.rows), q=$('search').value.toLocaleLowerCase();$('notes').replaceChildren();for(let r of hs.filter(r=>r.state!=='D'&&r.text.toLocaleLowerCase().includes(q))){let b=document.createElement('button');b.className=r.id===active?.id?'selected':'';b.textContent=(r.pin?'• ':'')+title(r)+(hs.filter(x=>x.note===r.note).length>1?' [conflito]':'');let small=document.createElement('small');small.textContent=new Date(Number((BigInt(r.time)-621355968000000000n)/10000n)).toLocaleString('pt-BR');b.append(small);b.onclick=()=>open(r);$('notes').append(b)} $('title').textContent=active?title(active):'Selecione uma nota';}
function commit(){clearTimeout(timer);if(!dirty||!active)return;let r=revision(active,$('text').value);db.rows.push(r);db.pending.push(r.id);db.draft=null;persist();active=r;dirty=false;render();status('Salvo localmente · '+db.pending.length+' pendente(s)');}
function open(r){commit();active=r;$('text').disabled=false;$('text').value=r.text;document.body.classList.add('editing');render()}
$('text').oninput=()=>{try{dirty=true;db.draft={base:active,text:$('text').value};persist();$('title').textContent=title({text:$('text').value});status('Rascunho salvo neste dispositivo');clearTimeout(timer);timer=setTimeout(()=>{try{commit();sync()}catch(e){status(e.message)}},900)}catch(e){status('Falha ao salvar localmente: '+e.message)}};
$('new').onclick=()=>{try{commit();let r=revision(null,'');db.rows.push(r);db.pending.push(r.id);persist();open(r);$('text').focus();sync()}catch(e){status(e.message)}};
$('back').onclick=()=>{commit();if(matchMedia('(max-width:650px)').matches)document.body.classList.remove('editing');else $('list').classList.toggle('hidden')};$('search').oninput=render;$('menu').onclick=()=>$('settings').showModal();$('close').onclick=()=>$('settings').close();$('sync').onclick=()=>sync();
const url64=a=>btoa(String.fromCharCode(...a)).replaceAll('+','-').replaceAll('/','_').replace(/=+$/,'');
async function login(){commit();let verifier=url64(crypto.getRandomValues(new Uint8Array(32))),state=id(),challenge=url64(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(verifier))));sessionStorage.setItem('ns-auth',JSON.stringify({verifier,state}));location.href=authority+'authorize?'+new URLSearchParams({client_id:client,response_type:'code',redirect_uri:redirect,response_mode:'query',scope,state,code_challenge:challenge,code_challenge_method:'S256',prompt:'select_account'})}
function updateControls(){
 const busy=starting||running||diagnosing;
 for(const name of ['login','logout','real']) $(name).disabled=busy;
 $('diagnose').disabled=starting||diagnosing;
 $('account').textContent=starting?'Retomando sessão Microsoft…':
  token||refresh?'Sessão Microsoft neste navegador':'Microsoft desconectado';
}
async function exchange(fields){
 const response=await timedFetch(authority+'token',{
  method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
  body:new URLSearchParams({client_id:client,redirect_uri:redirect,scope,...fields})
 });
 const result=await response.json();
 if(!response.ok){
  const error=Error(result.error_description||result.error||'Falha na autenticação');
  error.oauthCode=result.error;
  throw error;
 }
 const lifetime=Number(result.expires_in);
 if(!result.access_token||!Number.isFinite(lifetime)||lifetime<=0)throw Error('Resposta de autenticação inválida');
 token=result.access_token;
 expiry=Date.now()+lifetime*1000;
 refresh=result.refresh_token||(fields.grant_type==='refresh_token'?refresh:null);
 storeSession();
 updateControls();
}
async function auth(){
 if(token&&Date.now()<expiry-60000)return token;
 if(refresh){
  try{
   await exchange({grant_type:'refresh_token',refresh_token:refresh});
   return token;
  }catch(error){
   if(['invalid_grant','interaction_required','login_required','consent_required'].includes(error.oauthCode)){
    clearSession();
    throw Error('Sessão Microsoft expirada. Entre novamente para sincronizar.');
   }
   throw error;
  }
 }
 clearSession();
 throw Error('Entre com Microsoft para sincronizar');
}
async function timedFetch(url,opt={}){
 const ctrl=new AbortController();
 const timeout=setTimeout(()=>ctrl.abort(),20000);
 try{return await fetch(url,{...opt,signal:ctrl.signal})}
 catch(error){
  if(error.name==='AbortError')throw Error('Tempo limite de conexão');
  if(error instanceof TypeError)throw Error('Falha de conexão');
  throw error;
 }finally{clearTimeout(timeout)}
}
async function req(path,opt={}){
 const url=new URL(path.startsWith('https://')?path:graph+path);
 if(url.origin!=='https://graph.microsoft.com')throw Error('URL Graph inválida');
 const send=async()=>timedFetch(url,{...opt,cache:'no-store',headers:{...opt.headers,Authorization:'Bearer '+await auth()}});
 let response=await send();
 if(response.status===401&&refresh){
  expiry=0;
  response=await send();
 }
 if(!response.ok){
  const error=Error('OneDrive: HTTP '+response.status);
  error.code=response.status;
  throw error;
 }
 return response;
}
const base=()=>mode==='real'?['OUTROS ARQUIVOS','Programador','NotasSimples','Dados']:['NotasSimplesMobileTeste','Dados'];const path=a=>'/me/drive/root:/'+a.map(encodeURIComponent).join('/');
async function ensure(a){
 for(let i=0;i<a.length;i++){
  const parts=a.slice(0,i+1);
  let item;
  try{
   item=await (await req(path(parts))).json();
  }catch(error){
   if(error.code!==404)throw error;
   if(mode==='real')throw Error('Pasta real não encontrada: '+parts.join('/'));
   const parent=i?path(parts.slice(0,-1))+':/children':'/me/drive/root/children';
   try{
    item=await (await req(parent,{method:'POST',headers:{'Content-Type':'application/json'},
     body:JSON.stringify({name:parts.at(-1),folder:{},'@microsoft.graph.conflictBehavior':'fail'})})).json();
   }catch(conflict){
    if(conflict.code!==409)throw conflict;
    item=await (await req(path(parts))).json();
   }
  }
  if(!item.folder)throw Error('Não é pasta: '+parts.at(-1));
 }
}
async function downloadFile(item){
 let url=item['@microsoft.graph.downloadUrl'];
 if(!url&&item.id){
  const meta=await (await req('/me/drive/items/'+encodeURIComponent(item.id))).json();
  url=meta['@microsoft.graph.downloadUrl'];
 }
 if(!url||new URL(url).protocol!=='https:')throw Error('Download indisponível');
 const response=await timedFetch(url,{cache:'no-store',credentials:'omit'});
 if(!response.ok)throw Error('Falha ao ler arquivo do OneDrive: HTTP '+response.status);
 return response.text();
}
async function remoteRows(a){
 let next=path(a)+':/children?$select=id,name,folder,file,@microsoft.graph.downloadUrl&$top=200';
 const out=[];
 const pending=new Set(db.pending);
 const cached=new Map(db.rows.filter(r=>!pending.has(r.id)).map(r=>[r.id+'.rev',r]));
 while(next){
  const page=await (await req(next)).json();
  for(const item of page.value){
   if(item.folder||!item.name.endsWith('.rev'))continue;
   const row=cached.get(item.name)||await parse(await downloadFile(item));
   if(row.id+'.rev'!==item.name)throw Error('Nome de revisão incompatível com o conteúdo');
   out.push(row);
  }
  next=page['@odata.nextLink'];
 }
 return out;
}
async function upload(a,body,normalize=text=>text){
 const itemPath=path(a);
 try{
  await req(itemPath+':/content?@microsoft.graph.conflictBehavior=fail',{
   method:'PUT',headers:{'Content-Type':'text/plain; charset=utf-8'},body
  });
 }catch(error){
  if(error.code!==409&&error.code!==412)throw error;
  const item=await (await req(itemPath)).json();
  if(await normalize(await downloadFile(item))!==await normalize(body)){
   throw Error('Arquivo já existe com conteúdo diferente. Revisão local preservada.');
  }
 }
}
async function sync(){
 if(starting||running||diagnosing)return;
 try{
  commit();
  if(!token&&!refresh){status('Salvo localmente · entre com Microsoft para sincronizar');return}
  if(!navigator.onLine){status('Sem conexão · alterações salvas neste dispositivo');return}
  running=true;
  updateControls();
  status('Sincronizando…');
  const a=base(), historyPath=a.concat('Historico'), textsPath=a.concat('Textos');
  await ensure(historyPath);
  await ensure(textsPath);
  const remote=await remoteRows(historyPath);
  const known=new Map(remote.map(r=>[r.id,r]));
  const rows=new Map(db.rows.map(r=>[r.id,r]));
  for(const pending of [...db.pending]){
   const row=rows.get(pending);
   if(!row)throw Error('Revisão pendente ausente');
   const body=await serialize(row);
   if(known.has(row.id)){
    if(await serialize(known.get(row.id))!==body)throw Error('Conflito de identificador. Revisão local preservada.');
   }else{
    await upload(textsPath.concat(row.note+'.'+row.id+'.txt'),row.text);
    await upload(historyPath.concat(row.id+'.rev'),body,async text=>serialize(await parse(text)));
    known.set(row.id,row);
   }
   db.pending=db.pending.filter(id=>id!==pending);
   persist();
  }
  const merged=new Map(db.rows.map(r=>[r.id,r]));
  for(const row of known.values())merged.set(row.id,row);
  db.rows=[...merged.values()];
  persist();
  if(active&&!dirty&&!heads(db.rows).some(r=>r.id===active.id)){
   const candidates=heads(db.rows).filter(r=>r.note===active.note&&r.state!=='D');
   if(candidates.length===1){active=candidates[0];$('text').value=active.text}
   else{active=null;$('text').value='';$('text').disabled=true;document.body.classList.remove('editing')}
  }
  render();
  status('OneDrive atualizado · '+db.pending.length+' pendente(s)');
 }catch(error){
  status('Não sincronizado: '+error.message+' · dados locais preservados');
 }finally{
  running=false;
  updateControls();
 }
}
$('login').onclick=async()=>{
 if(starting||running||diagnosing)return;
 starting=true;
 updateControls();
 try{await login()}catch(error){starting=false;updateControls();$('account').textContent=error.message;status(error.message)}
};
$('logout').onclick=()=>{
 if(starting||running||diagnosing)return;
 commit();
 clearSession();
 status('Sessão encerrada neste aplicativo');
 $('settings').close();
};
$('real').onclick=()=>{
 if(starting||running||diagnosing)return;
 commit();
 if(mode==='test'&&!confirm('Acessar notas reais? Faça backup dos Dados no Windows antes.'))return;
 mode=mode==='test'?'real':'test';
 localStorage.setItem('ns-mode',mode);
 db=load();
 active=null;
 dirty=false;
 $('text').value='';
 $('text').disabled=true;
 document.body.classList.remove('editing');
 render();
 $('settings').close();
 sync();
};
if(db.draft){let r=revision(db.draft.base,db.draft.text);db.rows.push(r);db.pending.push(r.id);db.draft=null;persist()};render();window.addEventListener('pagehide',()=>{try{commit()}catch{}});window.addEventListener('online',sync);setInterval(()=>{if(document.visibilityState==='visible'&&(token||refresh))sync()},30000);
async function start(){
 updateControls();
 if('serviceWorker'in navigator){
  navigator.serviceWorker.register('./sw.js',{updateViaCache:'none'}).catch(()=>{});
 }
 try{
  const query=new URLSearchParams(location.search);
  if(query.has('code')||query.has('error')){
   const raw=sessionStorage.getItem('ns-auth');
   history.replaceState(null,'',new URL(redirect).pathname);
   sessionStorage.removeItem('ns-auth');
   const saved=JSON.parse(raw||'null');
   if(!saved?.verifier||!saved?.state||saved.state!==query.get('state')){
    throw Error('Não foi possível retomar o login nesta janela. Entre novamente.');
   }
   if(query.has('error'))throw Error('Login Microsoft interrompido. Entre novamente.');
   await exchange({grant_type:'authorization_code',code:query.get('code'),code_verifier:saved.verifier});
  }
 }catch(error){
  status(error.message);
  starting=false;
  updateControls();
  $('account').textContent=error.message;
  return;
 }
 starting=false;
 updateControls();
 if(token||refresh)await sync();
}

// Diagnostic uses only a unique dedicated folder.
$('closeReport').onclick=()=>{if(!diagnosing)$('diagnostic').close()};
$('downloadReport').onclick=()=>{let u=URL.createObjectURL(new Blob([$('report').textContent],{type:'text/plain;charset=utf-8'}));let a=document.createElement('a');a.href=u;a.download='Relatorio_Notas_Mobile.txt';a.click();setTimeout(()=>URL.revokeObjectURL(u),30000)};
$('diagnose').onclick=async()=>{
 if(starting||diagnosing)return;
 if(mode!=='test'){status('Volte ao modo Teste antes de executar o diagnóstico');$('settings').close();return}
 $('settings').close();$('diagnostic').showModal();$('report').textContent='NOTAS SIMPLES MOBILE — DIAGNÓSTICO 1.4\n'+new Date().toISOString()+'\n';$('downloadReport').disabled=true;diagnosing=true;updateControls();$('closeReport').disabled=true;
 const log=s=>{$('report').textContent+=s+'\n'};const check=(ok,name)=>{if(!ok)throw Error(name);log('OK: '+name)};
 let folder=null, passed=false, owned=false;
 try{
  log('Aguardando sincronização anterior, se houver…');
  const deadline=Date.now()+120000;while(running){if(Date.now()>deadline)throw Error('Sincronização anterior ainda em andamento; tente novamente');await new Promise(r=>setTimeout(r,200));}
  running=true;owned=true;
  let a=revision(null,'Diagnóstico fictício: ação, ç, ã, 📝\r\nSegunda linha');let b=revision(a,'Versão A');let c=revision(a,'Versão B');
  check((await parse(await serialize(a))).text===a.text,'UTF-8 e checksum');
  let rejected=false;try{await parse((await serialize(a)).replace('NS2','NS3'))}catch{rejected=true}check(rejected,'Rejeição de revisão inválida');
  check(heads([a,b,c]).length===2,'Preservação local de conflitos');
  await auth();log('OK: autenticação Microsoft');
  let name='NotasSimplesDiagnostico_'+id();let created=await (await req('/me/drive/root/children',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,folder:{},'@microsoft.graph.conflictBehavior':'fail'})})).json();folder=created.id;check(!!folder&&!!created.folder,'Criação de pasta isolada no OneDrive');
  let basePath='/me/drive/items/'+encodeURIComponent(folder);
  for(let r of [a,b,c]){let body=await serialize(r);await req(basePath+':/'+r.id+'.rev:/content',{method:'PUT',headers:{'Content-Type':'text/plain; charset=utf-8'},body});}
  log('OK: gravação de três revisões fictícias');
  let listing=await (await req(basePath+'/children')).json();let recovered=[];
  for(let r of [a,b,c]){let f=listing.value.find(x=>x.name===r.id+'.rev');check(!!f,'Localização da revisão '+(r===a?'original':r===b?'A':'B'));let parsed=await parse(await downloadFile(f));check(await serialize(parsed)===await serialize(r),'Leitura e integridade da revisão');recovered.push(parsed)}
  check(heads(recovered).length===2,'Conflitos preservados após leitura do OneDrive');
  let snapshot='Nota fictícia em TXT: português';let item=await (await req(basePath+':/teste.txt:/content',{method:'PUT',headers:{'Content-Type':'text/plain; charset=utf-8'},body:snapshot})).json();check(await downloadFile(item)===snapshot,'Gravação e leitura TXT');passed=true;
 }catch(e){log('FALHA: '+(e.code?'HTTP '+e.code:String(e.message).replace(/https?:\/\/\S+/g,'[URL omitida]').slice(0,300)))}
 finally{if(folder){try{await req('/me/drive/items/'+encodeURIComponent(folder),{method:'DELETE'});log('OK: pasta fictícia removida (lixeira do OneDrive)')}catch{log('AVISO: pasta fictícia não removida; pode excluir NotasSimplesDiagnostico_ no OneDrive')}}log(passed?'RESULTADO: DIAGNÓSTICO APROVADO':'RESULTADO: DIAGNÓSTICO INCOMPLETO OU REPROVADO');log('NÃO TESTADO: sincronização real Windows/iPhone/Android, execução offline após reinício e aparência.');diagnosing=false;if(owned)running=false;updateControls();$('closeReport').disabled=false;$('downloadReport').disabled=false;}
};

start();
