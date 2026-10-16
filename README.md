# _config.yml
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>CRM Consignado INSS</title>
<link href="https://fonts.googleapis.com/css2?family=Figtree:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.tailwindcss.com/3.4.17"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/dist/umd/supabase.js"></script>
<style>
:root{--bg:#f4f6f5;--card:#fff;--ink:#12302b;--mut:#5d716d;--line:#dbe3e0;--head:#e4ece9;--acc:#0f766e;--gold:#fff4d6;--goldln:#e0a100;box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#0d1917;--card:#142623;--ink:#e6f1ee;--mut:#8fa8a3;--line:#24403b;--head:#1c3732;--acc:#2dd4bf;--gold:#3a2f10;--goldln:#e0a100}}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
body{background:var(--bg);color:var(--ink);font-family:'Figtree',system-ui,sans-serif;margin:0}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px}
.mut{color:var(--mut)}
.btn{display:inline-flex;align-items:center;gap:.4rem;padding:.5rem .85rem;border-radius:10px;border:1px solid var(--line);background:var(--card);font-weight:600;font-size:.875rem;cursor:pointer}
.btn:hover{border-color:var(--acc)}
.btn:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible{outline:2px solid var(--acc);outline-offset:2px}
.btn-p{background:var(--acc);color:#fff;border-color:var(--acc)}
[hidden]{display:none!important}
body.locked>header,body.locked>main{visibility:hidden}
.btn-q{background:var(--goldln);border-color:var(--goldln);color:#2b1f00;font-weight:700;box-shadow:0 2px 10px rgba(224,161,0,.35)}
.chipb{border:1px solid var(--line);border-radius:999px;padding:.3rem .7rem;font-size:.8rem;font-weight:600;background:var(--card);color:var(--ink);cursor:pointer}.chipb:hover:not(:disabled){border-color:var(--acc)}.chipb.on{background:var(--acc);color:#fff;border-color:var(--acc)}.chipb:disabled{opacity:.4;cursor:not-allowed}
@media (prefers-color-scheme:dark){.chipb.on{color:#06201c}}
.sc{display:inline-block;border-radius:6px;padding:.05rem .35rem;font-weight:700;font-size:.72rem;white-space:nowrap}
.sc-banner{border-radius:10px;padding:.45rem .7rem}.sc-b1{font-weight:800;font-size:1rem;line-height:1.3}.sc-b2{font-size:.8rem;margin-top:.1rem;line-height:1.3}
.sc-ouro{background:#fff1c2;color:#6b4a00;border:1px solid #e0a100}.sc-prata{background:#eceff1;color:#37474f;border:1px solid #b0bec5}.sc-bronze{background:#f6e4d6;color:#6d3b1c;border:1px solid #cd8b5a}
#qd{width:min(560px,94vw)}.q-act{padding:.45rem .7rem}
@media (prefers-color-scheme:dark){.btn-p{color:#06201c}}
.inp{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:.5rem .75rem;color:var(--ink);width:100%}
.tbl,.tbl *{box-sizing:border-box}
.tbl{width:100%;table-layout:fixed;border-collapse:separate;border-spacing:0}
.tbl th{position:sticky;top:env(safe-area-inset-top,0px);z-index:5;background:var(--head);color:var(--ink);font-size:.62rem;text-transform:uppercase;letter-spacing:.03em;font-weight:700;text-align:left;padding:.32rem .4rem;line-height:1.15;border-top:1px solid var(--line);border-right:1px solid var(--line);border-bottom:2px solid var(--acc)}
.tbl th:first-child{border-top-left-radius:13px}.tbl th:last-child{border-top-right-radius:13px}
.tbl td{padding:.18rem .4rem;font-size:.72rem;line-height:1.25;vertical-align:middle;overflow-wrap:anywhere;border-right:1px solid var(--line);border-bottom:1px solid var(--line)}
.tbl th:last-child,.tbl td:last-child{border-right:0}
.nw{white-space:nowrap;overflow-wrap:normal}
.chip{display:inline-block;background:var(--head);border-radius:6px;padding:.05rem .35rem;font-size:.68rem;line-height:1.2}
.wrap{overflow:visible}
@media (max-width:900px){.wrap{overflow-x:auto}.tbl{min-width:900px}}
#rows button.mut{padding:.05rem .2rem;font-size:.7rem;line-height:1.2;border-radius:6px}
@media (max-width:1199px){.tbl td{font-size:.62rem;padding-left:.25rem;padding-right:.25rem}.tbl th{padding-left:.25rem;padding-right:.25rem}}
@media (max-width:1279px){.tbl td{font-size:.68rem}.tbl th{font-size:.58rem}.mini{font-size:.62rem}.sel{font-size:.66rem}}
tr.posse td{background:var(--gold)}
tr.posse td:first-child{box-shadow:inset 4px 0 0 var(--goldln)}
tr.done td{opacity:.6}
.mini{border:1px solid var(--line);border-radius:6px;padding:.05rem .28rem;font-size:.66rem;line-height:1.3;font-weight:600;background:var(--card);cursor:pointer;white-space:nowrap}.mini:hover:not(:disabled){border-color:var(--acc)}.mini:disabled{opacity:.4;cursor:not-allowed}
.mini.ico{display:inline-flex;align-items:center;justify-content:center;min-width:1.3rem;height:1.3rem;padding:0 .15rem;font-size:.78rem;line-height:1}.mini.ico.w2{min-width:2rem;font-size:.7rem}.mini.ico.sp{margin-left:.35rem}
@keyframes blink{50%{opacity:.35}}
.blink{animation:blink 1s infinite}
.alert-tag{display:inline-block;background:#dc2626;color:#fff;border-radius:8px;padding:.05rem .4rem;font-size:.68rem;font-weight:700;line-height:1.25}
.wait-tag{display:inline-block;background:var(--line);color:var(--ink);border-radius:8px;padding:.05rem .4rem;font-size:.68rem;font-weight:600;line-height:1.25}
.mini.hot{border-color:#dc2626;color:#dc2626;animation:blink 1s infinite}
tr.due td{box-shadow:inset 0 0 0 9999px rgba(220,38,38,.09)}
@media (prefers-reduced-motion:reduce){.blink,.mini.hot{animation:none}}
.sel{padding:.08rem .15rem;font-size:.7rem;width:100%;min-width:0;border-radius:6px}
.tag{background:var(--goldln);color:#2b1f00;border-radius:6px;padding:.05rem .4rem;font-size:.68rem;font-weight:700;white-space:nowrap}
dialog{background:var(--card);color:var(--ink);border:1px solid var(--line);border-radius:16px;padding:0;width:min(680px,94vw)}
dialog::backdrop{background:rgba(0,0,0,.5)}
#toast{position:fixed;left:50%;bottom:calc(1.25rem + env(safe-area-inset-bottom,0px));transform:translateX(-50%);background:var(--ink);color:var(--bg);padding:.6rem 1rem;border-radius:10px;font-size:.875rem;opacity:0;pointer-events:none;transition:opacity .2s;z-index:50}
#toast.on{opacity:1}
#pt,.po{display:none}
@media print{
 :root{--bg:#fff;--card:#fff;--ink:#000;--mut:#333;--line:#999;--gold:#eee}
 .no-print{display:none!important}#pt{display:block}.po{display:inline}
 .wrap{overflow:visible!important}.tbl{table-layout:auto;min-width:0}.tbl th{position:static}.tbl td,.tbl th{font-size:10px;padding:3px 5px}
 tr{break-inside:avoid}.card{border:0}
}
</style>
</head>
<body>
<header class="no-print" style="background:var(--card);border-bottom:1px solid var(--line)">
 <div class="w-full px-3 sm:px-4 py-3 flex flex-wrap items-center gap-3 justify-between">
  <div><h1 class="text-xl font-bold leading-tight">CRM Consignado INSS</h1>
  <p class="mut text-sm">Base de clientes com prioridade para Posse (GO)</p></div>
  <div class="flex flex-wrap gap-2">
   <button class="btn btn-q" id="bQ">▶ Iniciar Fila de Vendas</button>
   <button class="btn btn-p" id="bImp">Importar</button>
   <button class="btn" id="bExp">Exportar CSV</button>
   <button class="btn" id="bIcs">Exportar Todos os Agendamentos (.ics)</button>
   <button class="btn" id="bPrt">Imprimir</button>
   <button class="btn" id="bShr">Compartilhar</button>
   <button class="btn" id="bCloud">☁ Conectar à nuvem</button>
   <span id="usr" class="mut text-xs self-center" hidden style="max-width:190px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"></span>
   <button class="btn" id="bOut" hidden>Sair</button>
  </div>
 </div>
</header>

<main class="w-full px-3 sm:px-4 py-3 space-y-3">
 <section id="alerts" class="no-print" aria-live="polite"></section>
 <section class="no-print grid grid-cols-2 lg:grid-cols-4 xl:grid-cols-9 gap-2" id="kpis"></section>

 <section class="card no-print p-3 flex flex-wrap gap-2 items-center">
  <input id="q" class="inp sm:max-w-xs" placeholder="Buscar nome, CPF, telefone ou cidade">
  <select id="f" class="inp sm:max-w-[200px]">
   <option value="all">Todos</option><option value="posse">Somente Posse</option>
   <option value="pend">Pendentes (sem status)</option><option value="feito">Contato feito</option>
   <option value="aguard">Aguardando retorno</option><option value="erro">Contato errado</option>
   <option value="fech">Fechamento (consignado/refin.)</option><option value="res">Com algum resultado</option><option value="ouro">Somente Selo Ouro 🥇 (80-100)</option><option value="prata">Somente Selo Prata 🥈 (50-79)</option><option value="bronze">Somente Selo Bronze 🥉 (0-49)</option><option value="ret">Retorno programado vencido</option>
  </select>
  <input id="atd" class="inp sm:max-w-[160px]" placeholder="Seu nome" aria-label="Seu nome"><input id="emp" class="inp sm:max-w-[190px]" placeholder="Empresa" aria-label="Empresa"><span class="mut text-sm ml-auto" id="cnt"></span>
  <button class="btn" id="bClr">Limpar base</button>
 </section>

 <h2 id="pt" class="text-lg font-bold">CRM Consignado INSS — Lista de clientes</h2>

 <section class="card wrap">
  <table class="tbl" id="tb">
   <thead><tr><th>Nome</th><th style="width:5.2%" title="Score preditivo de 0 a 100">Score</th><th style="width:8%">CPF</th><th style="width:8.5%">Telefones</th><th style="width:7%">Benefício / ESP</th><th style="width:6.8%">Margem</th><th style="width:3.6%">Idade</th><th style="width:5.5%">Cidade</th><th style="width:10%">Status Contato</th><th style="width:10%">Resultado Fechamento</th><th style="width:8.5%">Retorno Programado</th><th class="no-print" style="width:13%" title="Passe o mouse sobre os ícones para ver a função">Mensagens</th><th class="no-print" style="width:2.2%"></th></tr></thead>
   <tbody id="rows"></tbody>
  </table>
  <div id="empty" class="p-10 text-center hidden">
   <p class="font-semibold">Nenhum cliente na base</p>
   <p class="mut text-sm mb-3">Cole os dados ou carregue um CSV para começar.</p>
   <button class="btn btn-p no-print" onclick="openImp()">Importar clientes</button>
  </div>
 </section>
</main>

<dialog id="dlg">
 <div class="p-5 space-y-3">
  <h3 class="text-lg font-bold">Importar clientes</h3>
  <p class="mut text-sm">Ordem das colunas: Nome, CPF, Telefone 01, Telefone 02, Benefício, ESP, Margem, Idade, Cidade. Depois de Cidade podem vir Status Contato, Data/Hora Status 1, Resultado Fechamento, Data/Hora Status 2 e Retorno Programado (calculado em 6 meses se vier vazio). Aceita Excel (.xlsx), CSV (; , ou Tab) ou texto colado. Com cabeçalho, as colunas são reconhecidas pelo nome.</p>
  <textarea id="paste" rows="8" class="inp font-mono text-xs" placeholder="Maria Souza;123.456.789-00;62999990000;;1234567890;41;R$ 1.250,50;67;Posse"></textarea>
  <div class="flex flex-wrap items-center gap-3">
   <input type="file" id="file" accept=".xlsx,.xls,.csv,.txt,.tsv" class="text-sm">
   <label class="text-sm flex items-center gap-2"><input type="checkbox" id="repl"> Substituir base atual</label>
  </div>
  <p id="prev" class="text-sm mut"></p>
  <div class="flex justify-end gap-2">
   <button class="btn" id="bCancel">Cancelar</button>
   <button class="btn btn-p" id="bGo">Importar</button>
  </div>
 </div>
</dialog>
<dialog id="obsd"><div class="p-5 space-y-3"><h3 class="text-lg font-bold" id="obsT">Observações</h3><textarea id="obsX" rows="5" class="inp" maxlength="600"></textarea><div class="flex justify-end gap-2"><button class="btn" id="obsC">Cancelar</button><button class="btn btn-p" id="obsS">Salvar</button></div></div></dialog>
<dialog id="qd" aria-labelledby="qT"><div class="p-5 space-y-4"><div class="flex items-center justify-between gap-2"><h3 class="text-lg font-bold" id="qT">Fila de Vendas <span class="mut text-sm font-normal" id="qPos"></span></h3><button class="btn" id="qX" aria-label="Fechar fila">✕</button></div><div id="qc"></div></div></dialog>
<dialog id="cfgd" aria-labelledby="cfgT"><div class="p-5 space-y-3">
 <h3 class="text-lg font-bold" id="cfgT">Conexão com a nuvem (Supabase)</h3>
 <p class="mut text-sm">Informe a URL e a chave <b>anon/public</b> do seu projeto (Supabase → Project Settings → API). Nunca use a chave service_role ou secret.</p>
 <label class="block text-sm font-semibold">Supabase URL<input id="sbU" class="inp mt-1" placeholder="https://xxxx.supabase.co" autocomplete="off" spellcheck="false"></label>
 <label class="block text-sm font-semibold">Anon Key<div class="flex gap-2 mt-1"><input id="sbK" type="password" class="inp" autocomplete="off" spellcheck="false"><button class="btn" id="sbShow" type="button">Mostrar</button></div></label>
 <p id="sbMsg" class="text-sm" role="status"></p>
 <p class="mut text-xs">A chave anon é pública por desenho: quem protege os dados são as regras de segurança (RLS) do banco. Ela fica salva só neste navegador (localStorage, sem criptografia). A senha do operador nunca é armazenada.</p>
 <div class="flex flex-wrap justify-between gap-2"><button class="btn" id="sbOff">Desconectar (modo local)</button><div class="flex flex-wrap gap-2"><button class="btn" id="sbTest">Testar conexão</button><button class="btn" id="sbCancel">Fechar</button><button class="btn btn-p" id="sbSave">Salvar e conectar</button></div></div>
</div></dialog>
<dialog id="lg" aria-labelledby="lgT" style="width:min(400px,94vw)"><form id="lgF" class="p-5 space-y-3">
 <h3 class="text-lg font-bold" id="lgT">🔒 Acesso do operador</h3>
 <p class="mut text-sm">Entre com seu e-mail e senha para acessar a base de clientes.</p>
 <label class="block text-sm font-semibold">E-mail<input id="lgE" type="email" class="inp mt-1" autocomplete="username" required></label>
 <label class="block text-sm font-semibold">Senha<input id="lgP" type="password" class="inp mt-1" autocomplete="current-password" required></label>
 <p id="lgMsg" class="text-sm" role="alert" style="color:#dc2626"></p>
 <div class="flex flex-wrap justify-between gap-2"><button class="btn" id="lgCfg" type="button">Configurar conexão</button><button class="btn btn-p" id="lgGo" type="submit">Entrar</button></div>
</form></dialog>
<div id="toast" role="status"></div>

<script>
const KEY='crm_inss_v1';
let clients=[];
try{clients=JSON.parse(localStorage.getItem(KEY)||'[]')}catch(e){clients=[]}
clients.forEach(c=>{if(c.contatado&&!c.status)c.status='Contato feito';delete c.contatado;c.status=c.status||'';c.resultado=c.resultado||'';c.dt1=c.dt1||'';c.dt2=c.dt2||'';c.ret=c.ret||''});
const $=s=>document.querySelector(s);
const save=()=>{try{localStorage.setItem(KEY,JSON.stringify(clients))}catch(e){}};
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
const brl=n=>n.toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const isPosse=c=>norm(c.cidade).replace(/\s*-?\s*go$/,'')==='posse';
function toast(m){const t=$('#toast');t.textContent=m;t.classList.add('on');clearTimeout(toast.t);toast.t=setTimeout(()=>t.classList.remove('on'),2600)}

/* ---------- Parser ---------- */
function parseTable(text){
 text=text.replace(/^\uFEFF/,'').replace(/\r\n?/g,'\n').trim();
 if(!text)return[];
 const first=text.split('\n')[0];
 const cnt=d=>first.split(d).length-1;
 const delim=[['\t',cnt('\t')],[';',cnt(';')],[',',cnt(',')]].sort((a,b)=>b[1]-a[1])[0][0];
 const rows=[];let row=[],cell='',q=false;
 for(let i=0;i<text.length;i++){
  const ch=text[i];
  if(q){ if(ch==='"'){ if(text[i+1]==='"'){cell+='"';i++} else q=false } else cell+=ch }
  else if(ch==='"')q=true;
  else if(ch===delim){row.push(cell);cell=''}
  else if(ch==='\n'){row.push(cell);rows.push(row);row=[];cell=''}
  else cell+=ch;
 }
 row.push(cell);rows.push(row);
 return rows.filter(r=>r.some(c=>c.trim()));
}
const STATUS=['Contato feito','Aguardando retorno','Contato errado'];
const RESULT=['Realizado consignado','Realizado refinanciamento','Realizado simulação','Sem margem','Sem interesse'];
const ALIAS={'consegui contato':'Contato feito'};
try{const L=JSON.parse(localStorage.getItem(KEY+'_lists')||'{}');(L.s||[]).forEach(x=>STATUS.includes(x)||STATUS.push(x));(L.r||[]).forEach(x=>RESULT.includes(x)||RESULT.push(x))}catch(e){}
function learn(c){for(const[k,l]of[['status',STATUS],['resultado',RESULT]]){if(!c[k])continue;const m=l.find(x=>norm(x)===norm(c[k]));if(m)c[k]=m;else l.push(c[k])}try{localStorage.setItem(KEY+'_lists',JSON.stringify({s:STATUS,r:RESULT}))}catch(e){}}
const pad=n=>String(n).padStart(2,'0');
const brDt=(d,u)=>{const g=m=>d[(u?'getUTC':'get')+m]();return `${pad(g('Date'))}/${pad(g('Month')+1)}/${g('FullYear')} ${pad(g('Hours'))}:${pad(g('Minutes'))}:${pad(g('Seconds'))}`};
const nowBR=()=>brDt(new Date());
const txt=v=>typeof v==='number'?String(Math.round(v)):String(v??'').trim().replace(/^(\d+)\.0+$/,'$1');
const digits=s=>txt(s).replace(/\D/g,'');
function parseDt(v){
 if(v===''||v==null)return'';
 if(typeof v==='string'&&/^\d{5}(\.\d+)?$/.test(v.trim()))v=parseFloat(v);
 if(typeof v==='number')return v>20000?brDt(new Date(Math.floor((v-25569)*86400+1e-6)*1000),true):'';
 const s=String(v).trim(),m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\D+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
 if(m)return `${pad(+m[1])}/${pad(+m[2])}/${m[3]} ${pad(+(m[4]||0))}:${m[5]||'00'}:${m[6]||'00'}`;
 const d=new Date(s.replace(' ','T'));return isNaN(d)?s:brDt(d);
}
function canon(v,list){const n=norm(txt(v));if(!n)return'';return ALIAS[n]&&list===STATUS?ALIAS[n]:(list.find(x=>norm(x)===n)||txt(v))}
function mapHeader(r){
 const map=new Array(15).fill(-1);let tel=0,hits=0;
 r.forEach((h,i)=>{const n=norm(h).replace(/[^a-z0-9]/g,'');let k=-1;
  if(n.includes('nome'))k=0;
  else if(n==='cpf')k=1;
  else if(/^(tel|fone|cel|whats)/.test(n)){k=tel===0?2:3;tel++}
  else if(n.startsWith('benef')||n==='nb')k=4;
  else if(n.startsWith('esp'))k=5;
  else if(n.startsWith('marg'))k=6;
  else if(n.startsWith('idade'))k=7;
  else if(n.startsWith('cidade')||n.startsWith('munic'))k=8;
  else if(n.startsWith('statuscontato')||n==='contato')k=9;
  else if(n.startsWith('retorno'))k=13;
  else if(n.startsWith('obs'))k=14;
  else if(n.startsWith('data'))k=n.includes('retorno')?13:n.includes('status2')?12:10;
  else if(n.startsWith('resultado')||n==='status')k=11;
  if(k>-1&&map[k]<0){map[k]=i;hits++}
 });
 return hits>=3?map:null;
}
function parseMoney(v){
 if(typeof v==='number')return v;
 let s=String(v||'').replace(/[^\d.,-]/g,'');if(!s)return 0;
 if(s.includes(',')&&s.includes('.'))s=s.replace(/\./g,'').replace(',','.');
 else if(s.includes(','))s=s.replace(',','.');
 else if(/^\d{1,3}(\.\d{3})+$/.test(s))s=s.replace(/\./g,'');
 const n=parseFloat(s);return isNaN(n)?0:n;
}
function fmtCpf(s){const d=digits(s).padStart(11,'0').slice(-11);return d.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/,'$1.$2.$3-$4')}
function fmtTel(s){const d=digits(s);if(d.length===11)return d.replace(/(\d{2})(\d{5})(\d{4})/,'($1) $2-$3');if(d.length===10)return d.replace(/(\d{2})(\d{4})(\d{4})/,'($1) $2-$3');return txt(s)}
function buildClients(src){
 const rows=Array.isArray(src)?src.filter(r=>r.some(c=>String(c).trim())):parseTable(src);
 if(!rows.length)return{list:[],skipped:0};
 let map=mapHeader(rows[0]);if(map)rows.shift();else map=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14];
 const g=(r,k)=>map[k]>-1?(r[map[k]]??''):'';
 const list=[];let skipped=0;
 for(const r of rows){
  const nome=txt(g(r,0));if(!nome){skipped++;continue}
  list.push({id:crypto.randomUUID?crypto.randomUUID():Date.now()+Math.random().toString(16),
   nome,cpf:digits(g(r,1))?fmtCpf(g(r,1)):'',t1:fmtTel(g(r,2)),t2:fmtTel(g(r,3)),
   ben:txt(g(r,4)),esp:txt(g(r,5)),margem:parseMoney(g(r,6)),idade:parseInt(txt(g(r,7)))||0,cidade:txt(g(r,8)),
   status:canon(g(r,9),STATUS),resultado:canon(g(r,11),RESULT),dt1:parseDt(g(r,10)),dt2:parseDt(g(r,12)),ret:parseDt(g(r,13)),obs:txt(g(r,14))});
 }
 return{list,skipped};
}

/* ---------- Ordenação ---------- */
/* ---------- Score preditivo (0 a 100) ---------- */
const SC={margemTeto:1000}; // margem (R$) que já vale a pontuação máxima de 40
function scoreParts(c){
 const m=Math.round(Math.min(Math.max(c.margem||0,0)/SC.margemTeto,1)*40);
 const l=isPosse(c)?30:0;
 const id=c.idade?Math.max(0,20-2*(c.idade<50?50-c.idade:c.idade>70?c.idade-70:0)):10;
 const h=(c.status==='Aguardando retorno'||norm(c.resultado)==='realizado simulacao')?10:(c.status==='Contato feito'&&!c.resultado)?5:0;
 return{m,l,id,h,t:Math.min(100,m+l+id+h)};
}
const score=c=>scoreParts(c).t;
const tier=t=>t>=80?{k:'ouro',i:'🥇',n:'Ouro',d:'Alta propensão de fechamento'}:t>=50?{k:'prata',i:'🥈',n:'Prata',d:'Média propensão'}:{k:'bronze',i:'🥉',n:'Bronze',d:'Baixa propensão'};
function scoreBadge(c){const p=scoreParts(c),t=tier(p.t);return `<span class="sc sc-${t.k}" title="${t.n}: ${t.d}\nMargem ${p.m}/40 · Localização ${p.l}/30 · Idade ${p.id}/20 · Histórico ${p.h}/10">${t.i} ${p.t}</span>`}
function scoreTip(c,p,t){
 if(c.status==='Aguardando retorno')return 'Retorno pendente — retome a conversa e ofereça uma simulação.';
 if(t.k==='ouro')return isPosse(c)?(p.m>=30?'Alta margem em Posse — Prioridade Máxima de Fechamento':'Perfil ideal em Posse — alta chance de fechamento, aborde agora.'):'Perfil excelente — alta chance de fechamento, aborde agora.';
 if(t.k==='prata')return isPosse(c)?(p.m<20?'Em Posse, com margem moderada — vale uma simulação rápida.':'Bom perfil em Posse — boa chance com a abordagem certa.'):(p.m>=30?'Boa margem fora de Posse — priorize depois dos clientes Ouro.':'Média propensão — aborde com texto direto.');
 return 'Baixa propensão — abordagem rápida; priorize os clientes de score maior.'+(c.idade&&(c.idade<50||c.idade>70)?' Fora da faixa de idade ideal.':'');
}
function scoreBanner(c){const p=scoreParts(c),t=tier(p.t);return `<div class="sc-banner sc-${t.k}" role="note" title="Margem ${p.m}/40 · Localização ${p.l}/30 · Idade ${p.id}/20 · Histórico ${p.h}/10"><div class="sc-b1">${t.i} Selo ${t.n} - ${p.t} pts</div><div class="sc-b2">${esc(scoreTip(c,p,t))}</div></div>`}
function sorted(list){
 return [...list].sort((a,b)=>{
  const d=score(b)-score(a);if(d)return d;
  const pa=isPosse(a)?1:0,pb=isPosse(b)?1:0;
  if(pa!==pb)return pb-pa;
  return b.margem-a.margem;
 });
}

/* ---------- Datas, regras e mensagens ---------- */
const toDate=s=>{const m=String(s||'').match(/^(\d{2})\/(\d{2})\/(\d{4})(?: (\d{2}):(\d{2}):(\d{2}))?/);return m?new Date(+m[3],+m[2]-1,+m[1],+(m[4]||0),+(m[5]||0),+(m[6]||0)):null};
function addM(s,n){const d=toDate(s);if(!d)return'';const day=d.getDate();d.setDate(1);d.setMonth(d.getMonth()+n);d.setDate(Math.min(day,new Date(d.getFullYear(),d.getMonth()+1,0).getDate()));return brDt(d)}
function rules(c){
 if(c.status==='Contato errado'){c.resultado='';c.dt2='';c.ret=''}
 else if(c.resultado&&c.dt2&&!c.ret)c.ret=addM(c.dt2,6);
}
const cfg=(()=>{try{return JSON.parse(localStorage.getItem(KEY+'_cfg')||'{}')}catch(e){return{}}})();
$('#atd').value=cfg.atd||'';$('#emp').value=cfg.emp||'';
['atd','emp'].forEach(k=>$('#'+k).addEventListener('input',e=>{cfg[k]=e.target.value.trim();try{localStorage.setItem(KEY+'_cfg',JSON.stringify(cfg))}catch(_){}}));
const first=c=>{const n=c.nome.trim().split(/\s+/)[0].toLowerCase();return n.charAt(0).toUpperCase()+n.slice(1)};
function tr(c){
 const n=norm(first(c));
 const g=/^(luca|joshua|zeca)$/.test(n)?'m':/(a|ice|ene|ete|ise|ine|raquel|isabel|beatriz|ester|ruth|nair|iris|ivone)$/.test(n)?'f':/i$/.test(n)?'?':'m';
 return g==='f'?{s:'senhora',o:'a senhora',p:'pra senhora'}:g==='m'?{s:'senhor',o:'o senhor',p:'pro senhor'}:{s:'senhor(a)',o:'o(a) senhor(a)',p:'pro(a) senhor(a)'};
}
const who=()=>`${cfg.atd||'[seu nome]'}, da ${cfg.emp||'[empresa]'}`;
const MSG={
 t:c=>`Oi, ${first(c)}! Aqui é ${who()}. Vi que seu benefício do INSS tem margem livre pra consignado, com a parcela descontada direto no benefício. Quer que eu faça uma simulação sem compromisso? É só me responder aqui.`,
 a:c=>{const t=tr(c);return `Oi, ${t.s} ${first(c)}, tudo bem? Aqui é ${who()}. Tô entrando em contato porque vi aqui que ${t.o} tem um valor liberado no benefício do INSS pra fazer um crédito consignado, com a parcela descontada direto no benefício, sem burocracia. Se ${t.o} quiser, eu faço uma simulação sem compromisso e explico tudo direitinho ${t.p}, com calma. É só me responder por aqui mesmo, tá bom? Fica com Deus, viu?`},
 g:c=>`${first(c)}, só reforçando uma coisa importante: eu nunca peço senha do Meu INSS, código que chega por SMS nem depósito adiantado pra liberar nada. Quem pedir isso é golpe, viu? Nada é feito sem você autorizar, e você pode conferir tudo comigo direto aqui no WhatsApp ou numa agência. Pode me perguntar o que quiser, sem pressa. Aqui é ${who()}.`
};
const waNum=c=>{for(const t of[c.t1,c.t2]){const d=digits(t);if(d.length>=10&&d.length<=11)return'55'+d;if(d.length>=12&&d.length<=13&&d.startsWith('55'))return d}return''};
async function copy(t){
 try{await navigator.clipboard.writeText(t);return true}catch(e){}
 try{const a=document.createElement('textarea');a.value=t;a.style.position='fixed';a.style.opacity='0';document.body.appendChild(a);a.select();const ok=document.execCommand('copy');a.remove();if(ok)return true}catch(e){}
 prompt('Copie a mensagem:',t);return false;
}
function msgCell(c){
 const bad=c.status==='Contato errado',nw=!waNum(c),due=(fuState(c)||{}).st==='due';
 const b=(m,l,off,t)=>{const tt=off?(bad?'Desativado: contato marcado como errado':'Sem telefone válido para WhatsApp'):t;return `<button class="mini ico${m==='wg'?' w2':''}${m==='wt'?' sp':''}${due&&(m==='g'||m==='wg')?' hot':''}" data-m="${m}" data-i="${c.id}" ${off?'disabled':''} title="${tt}" aria-label="${tt}">${l}</button>`};
 return `<div class="flex flex-wrap items-center" style="gap:.15rem">${b('t','📄',bad,'Copiar texto de abordagem (direto e informal)')}${b('a','🎤',bad,'Copiar roteiro de áudio (tratamento senhor/senhora)')}${b('g','🛡️',bad,'Copiar mensagem antigolpe (follow-up de 30 min a 2 h)')}${b('wt','💬',bad||nw,'Abrir WhatsApp com a abordagem pronta')}${b('wg','💬🛡️',bad||nw,'Abrir WhatsApp com a mensagem antigolpe')}</div>`;
}

/* ---------- Follow-up 30 min a 2 h ---------- */
let sig='';
const hms=n=>{n=Math.max(0,Math.floor(n));return `${pad(Math.floor(n/3600))}:${pad(Math.floor(n%3600/60))}:${pad(n%60)}`};
function fuState(c){
 if(c.fu||c.resultado||!c.dt1||!['Contato feito','Aguardando retorno'].includes(c.status))return null;
 const d=toDate(c.dt1);if(!d)return null;
 const min=(Date.now()-d.getTime())/60000;
 if(min<0||min>1440)return null;
 return{min,st:min<30?'wait':min<=120?'due':'late'};
}
const sigNow=()=>{const m={};clients.forEach(c=>{const f=fuState(c);if(f)m[c.id]=f.st});return JSON.stringify(m)};
function tkText(f){const n=f.min*60;return f.st==='wait'?`⏳ follow-up em ${hms(1800-n)}`:f.st==='due'?`⏱ ${hms(n)} decorridos · fecha em ${hms(7200-n)}`:`⏱ ${hms(n)} decorridos`}
function fuBadge(c){
 if(c.fu)return `<div class="text-xs mt-1"><span class="wait-tag">✔ Antigolpe enviado ${esc(c.fu.slice(11,16))}</span> <button class="mini" data-fu="undo" data-i="${c.id}">desfazer</button></div>`;
 const f=fuState(c);if(!f)return'';
 const tag={wait:'',due:'<span class="alert-tag blink">⚠ Hora do Follow-up Antigolpe!</span>',late:'<span class="wait-tag">Janela de follow-up encerrada</span>'}[f.st];
 return `${tag?`<div class="mt-1">${tag}</div>`:''}<div class="mt-1 font-semibold" style="font-size:.66rem" data-tk="${c.id}">${tkText(f)}</div>`;
}
function renderAlerts(){
 const due=[];let wait=0;
 clients.forEach(c=>{const f=fuState(c);if(!f)return;if(f.st==='due')due.push([c,f]);else if(f.st==='wait')wait++});
 due.sort((a,b)=>b[1].min-a[1].min);
 {const fresh=due.filter(x=>!prevDue.has(x[0].id)).map(x=>x[0]);prevDue=new Set(due.map(x=>x[0].id));if(fresh.length)notify(fresh)}
 document.title=(due.length?`(${due.length}) `:'')+'CRM Consignado INSS';
 const el=$('#alerts');el.classList.toggle('hidden',!clients.length);
 const bt=(m,l,c)=>`<button class="mini hot" data-m="${m}" data-i="${c.id}" ${m==='wg'&&!waNum(c)?'disabled':''}>${l}</button>`;
 el.innerHTML=`<div class="card p-4" ${due.length?'style="border-color:#dc2626"':''}>
  <div class="flex items-center justify-between gap-2 flex-wrap mb-2"><h2 class="font-bold">🔔 Alertas de Follow-up <span class="alert-tag ${due.length?'blink':''}">${due.length}</span></h2>
  <span class="mut text-xs">Janela: 30 min a 2 h após a Data/Hora Status 1${wait?` · ${wait} aguardando a janela abrir`:''}</span>
  <div class="flex flex-wrap gap-4 text-xs w-full"><label class="flex items-center gap-1 cursor-pointer"><input type="checkbox" data-opt="snd" ${cfg.snd?'checked':''}> 🔊 Aviso sonoro</label><label class="flex items-center gap-1 cursor-pointer"><input type="checkbox" data-opt="ntf" ${cfg.ntf?'checked':''}> 🔔 Notificação do navegador</label></div></div>
  ${due.length?`<ul>${due.map(([c,f])=>`<li class="py-2 flex flex-wrap items-center gap-x-4 gap-y-2 justify-between" style="border-top:1px solid var(--line)"><div><span class="font-semibold">${esc(c.nome)}</span> ${isPosse(c)?'<span class="tag">Posse</span>':''}<div class="text-xs mut">${esc(c.t1||c.t2)} · ${esc(c.status)}</div></div><div class="text-sm font-semibold" data-tk="${c.id}">${tkText(f)}</div><div class="flex gap-1 flex-wrap">${bt('g','Copiar antigolpe',c)}${bt('wg','WhatsApp antigolpe',c)}<button class="mini" data-fu="mark" data-i="${c.id}">Marcar enviado</button></div></li>`).join('')}</ul>`:'<p class="mut text-sm">Nenhum follow-up pendente agora.</p>'}</div>`;
}
function tick(){
 if(sigNow()!==sig&&!(document.activeElement&&document.activeElement.tagName==='SELECT')){render();return}
 document.querySelectorAll('[data-tk]').forEach(el=>{const c=clients.find(x=>x.id===el.dataset.tk),f=c&&fuState(c);if(f)el.textContent=tkText(f)});
}
let prevDue=new Set(),ac;
const saveCfg=()=>{try{localStorage.setItem(KEY+'_cfg',JSON.stringify(cfg))}catch(_){}};
function beep(){
 try{ac=ac||new(window.AudioContext||window.webkitAudioContext)();if(ac.state==='suspended')ac.resume();
  [880,1100].forEach((fq,i)=>{const o=ac.createOscillator(),g=ac.createGain(),t=ac.currentTime+i*.25;o.frequency.value=fq;o.connect(g);g.connect(ac.destination);g.gain.setValueAtTime(.0001,t);g.gain.exponentialRampToValueAtTime(.25,t+.02);g.gain.exponentialRampToValueAtTime(.0001,t+.2);o.start(t);o.stop(t+.22)});
 }catch(e){}
}
function notify(list){
 if(authState==='locked')return;
 if(cfg.snd)beep();
 if(cfg.ntf&&'Notification' in window&&Notification.permission==='granted'){
  try{new Notification('Hora do Follow-up Antigolpe!',{body:list.length===1?`${list[0].nome}: janela de 30 min a 2 h aberta.`:`${list.length} clientes na janela de follow-up.`,tag:'followup'})}catch(e){}
 }
}
$('#alerts').addEventListener('change',async e=>{
 const k=e.target.dataset.opt;if(!k)return;
 if(k==='ntf'&&e.target.checked){
  if(!('Notification' in window)){toast('Este navegador não permite notificações.');e.target.checked=false;return}
  let p=Notification.permission;if(p==='default')p=await Notification.requestPermission();
  if(p!=='granted'){toast('Permissão negada. Libere as notificações nas configurações do navegador.');e.target.checked=false;return}
  try{new Notification('Notificações ativadas',{body:'Você será avisado quando um cliente entrar na janela de follow-up.'})}catch(_){}
 }
 cfg[k]=e.target.checked;saveCfg();
 if(e.target.checked&&k==='snd')beep();
 toast((k==='snd'?'Aviso sonoro ':'Notificação ')+(e.target.checked?'ativado.':'desativado.'));
});
setInterval(tick,1000);

/* ---------- Agenda (Google Agenda e .ics) ---------- */
const agOk=c=>c.status!=='Contato errado'&&!!toDate(c.ret);
function evOf(c){
 const d=toDate(c.ret);if(!d)return null;
 if(!d.getHours()&&!d.getMinutes()&&!d.getSeconds())d.setHours(9);
 const desc=[`Cliente: ${c.nome}`,`Telefone: ${[c.t1,c.t2].filter(Boolean).join(' / ')||'-'}`,`CPF: ${c.cpf||'-'}`,`Benefício: ${c.ben||'-'}${c.esp?' (ESP '+c.esp+')':''}`,`Cidade: ${c.cidade||'-'}`,`Último resultado: ${c.resultado||'-'}`,`Observações: ${c.obs||'-'}`].join('\n');
 return{title:`Retorno Consignado - ${c.nome}`,start:d,end:new Date(d.getTime()+30*60000),desc};
}
const icsDt=d=>`${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}T${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
const icsEsc=t=>String(t).replace(/\\/g,'\\\\').replace(/;/g,'\\;').replace(/,/g,'\\,').replace(/\r?\n/g,'\\n');
const fold=l=>{const o=[];for(let i=0;i<l.length;i+=36)o.push((i?' ':'')+l.slice(i,i+36));return o.join('\r\n')};
function icsOf(list){
 const st=new Date().toISOString().replace(/[-:]/g,'').replace(/\.\d+/,'');
 const L=['BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//CRM Consignado INSS//PT-BR','CALSCALE:GREGORIAN','METHOD:PUBLISH'];
 list.forEach(c=>{const v=evOf(c);L.push('BEGIN:VEVENT',`UID:${c.id}-retorno@crm-consignado`,`DTSTAMP:${st}`,`DTSTART:${icsDt(v.start)}`,`DTEND:${icsDt(v.end)}`,`SUMMARY:${icsEsc(v.title)}`,`DESCRIPTION:${icsEsc(v.desc)}`,'BEGIN:VALARM','TRIGGER:-PT30M','ACTION:DISPLAY',`DESCRIPTION:${icsEsc(v.title)}`,'END:VALARM','END:VEVENT')});
 L.push('END:VCALENDAR');return L.map(fold).join('\r\n')+'\r\n';
}
function download(name,type,content){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([content],{type}));a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
const slug=t=>norm(t).replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'').slice(0,40)||'cliente';
function gcalUrl(c){const v=evOf(c);return `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(v.title)}&dates=${icsDt(v.start)}/${icsDt(v.end)}&details=${encodeURIComponent(v.desc.slice(0,1200))}`}
function agCell(c){
 const bad=c.status==='Contato errado';
 const b=(a,l,t)=>`<button class="mini" data-ag="${a}" data-i="${c.id}" title="${t}">${l}</button>`;
 return `<div class="flex items-center justify-between gap-1"><span class="${bad?'mut':'nw'}">${bad?'Agendamento desativado':esc((c.ret||'').slice(0,10))}</span><span class="no-print">${b('o','📝','Observações do cliente')}</span></div>${!bad&&agOk(c)?`<div class="no-print flex flex-wrap gap-1 mt-1">${b('g','+ Google Agenda','Abrir no Google Agenda')}${b('i','Baixar .ics','Baixar evento .ics')}</div>`:''}${c.obs?`<div class="mut" style="font-size:.66rem">📝 ${esc(c.obs)}</div>`:''}`;
}
let obsId=null;
$('#rows').addEventListener('click',e=>{
 const b=e.target.closest('[data-ag]');if(!b)return;
 const c=clients.find(x=>x.id===b.dataset.i),a=b.dataset.ag;
 if(a==='o'){obsId=c.id;$('#obsT').textContent='Observações: '+c.nome;$('#obsX').value=c.obs||'';$('#obsd').showModal();return}
 if(!agOk(c)){toast('Agendamento desativado para este cliente.');return}
 if(a==='g')window.open(gcalUrl(c),'_blank','noopener');
 else{download(`retorno-${slug(c.nome)}.ics`,'text/calendar;charset=utf-8',icsOf([c]));toast('Arquivo .ics baixado.')}
});
$('#obsC').onclick=()=>$('#obsd').close();
$('#obsS').onclick=()=>{const c=clients.find(x=>x.id===obsId);if(c){c.obs=$('#obsX').value.trim();save();render()}$('#obsd').close()};
$('#bIcs').onclick=()=>{
 const l=view().filter(agOk);
 if(!l.length){toast('Nenhum retorno programado ativo na lista atual.');return}
 download(`agenda-retornos-${new Date().toISOString().slice(0,10)}.ics`,'text/calendar;charset=utf-8',icsOf(l));
 toast(`${l.length} agendamentos exportados (lista exibida).`);
};

/* ---------- Fila de Vendas (modo passa-prato) ---------- */
let qList=[],qIdx=0,qSeen=new Set();
function qBtn(){const n=clients.filter(FILTERS.pend).length;$('#bQ').textContent=`▶ Iniciar Fila de Vendas (${n})`}
function qOpen(){qList=sorted(clients).filter(FILTERS.pend);qSeen=new Set();qList.length?qShow(0):qDone();$('#qd').showModal()}
function qDone(){
 qIdx=qList.length;$('#qPos').textContent='· fila concluída';
 $('#qc').innerHTML=clients.length?`<div class="p-6 text-center space-y-2"><div style="font-size:2.2rem" aria-hidden="true">🎉</div><p class="text-lg font-bold">Parabéns! Todos os clientes da fila foram trabalhados.</p>${qSeen.size?`<p class="mut text-sm">Você trabalhou ${qSeen.size} cliente${qSeen.size>1?'s':''} nesta sessão.</p>`:''}<div class="flex justify-center gap-2 pt-2">${qList.length?'<button class="btn" data-qn="-1">◀ Voltar</button>':''}<button class="btn btn-p" data-qc="1">Fechar</button></div></div>`:'<div class="p-6 text-center"><p class="font-semibold">Nenhum cliente na base.</p><p class="mut text-sm">Importe uma lista para iniciar a fila.</p></div>';
}
function qAdvance(){
 let n=-1;
 for(let i=qIdx+1;i<qList.length;i++)if(!qList[i].status){n=i;break}
 if(n<0)for(let i=0;i<qIdx&&i<qList.length;i++)if(!qList[i].status){n=i;break}
 n<0?qDone():qShow(n);
}
function qPick(ch){
 const c=clients.find(x=>x.id===(ch.dataset.qs||ch.dataset.qr)),k=ch.dataset.qs?'s':'r',v=ch.dataset.v;
 const val=(k==='s'?c.status:c.resultado)===v?'':v;
 if(!applyChange(c,k,val)||!val){qShow(qIdx);return}
 qSeen.add(c.id);toast(`${first(c)}: ${val}. Próximo cliente…`);qAdvance();
}
function qNav(d){const n=qIdx+d;if(n>=0&&n<qList.length)qShow(n)}
function qShow(i){
 qIdx=i;const c=qList[i],el=$('#qc');
 $('#qPos').textContent=`· Cliente ${i+1} de ${qList.length} da fila`;
 const bad=c.status==='Contato errado',nw=!waNum(c);
 const tl=t=>{const d=digits(t);return d.length>=10&&d.length<=11?`<a class="underline" target="_blank" rel="noopener" href="https://wa.me/55${d}">${esc(t)}</a>`:esc(t)};
 const f=(l,v,cls)=>`<div class="${cls||''}"><div class="mut text-xs">${l}</div><div class="font-semibold">${v||'-'}</div></div>`;
 const qb=(m,ic,lb,off,t)=>`<button class="btn q-act" data-m="${m}" data-i="${c.id}" ${off?'disabled':''} title="${off?(bad?'Desativado: contato marcado como errado':'Sem telefone válido para WhatsApp'):t}" aria-label="${t}"><span aria-hidden="true">${ic}</span> ${lb}</button>`;
 const grp=(l,k,list,v,dt,off)=>`<div><div class="mut text-xs mb-1">${l}</div><div class="flex flex-wrap gap-1" role="group" aria-label="${l}">${[...list,...(v&&!list.includes(v)?[v]:[])].map(o=>`<button class="chipb${o===v?' on':''}" data-${k}="${c.id}" data-v="${esc(o)}" aria-pressed="${o===v}" ${off?'disabled':''}>${esc(o)}</button>`).join('')}</div>${dt?`<div class="mut text-xs mt-1">${esc(dt)}</div>`:''}</div>`;
 el.innerHTML=`<div class="rounded-xl p-4" style="background:${isPosse(c)?'var(--gold)':'var(--head)'};border:1px solid ${isPosse(c)?'var(--goldln)':'var(--line)'}">
  ${scoreBanner(c)}<div class="flex items-start justify-between gap-2" style="margin-top:.6rem"><div class="text-xl font-bold leading-tight">${esc(c.nome)}</div><div class="flex flex-wrap items-center justify-end gap-1">${isPosse(c)?'<span class="tag">★ Posse · Prioridade</span>':''}</div></div>
  <div class="mt-3 grid grid-cols-2 gap-3">
   ${f('Margem',`<span class="text-xl" style="color:var(--acc)">${brl(c.margem)}</span>`)}
   ${f('Cidade',esc(c.cidade))}
   ${f('CPF',esc(c.cpf))}
   ${f('Benefício',esc(c.ben)+(c.esp?` <span class="mut font-normal">(ESP ${esc(c.esp)})</span>`:''))}
   ${f('Telefones',[c.t1,c.t2].filter(Boolean).map(tl).join('<br>'),'col-span-2')}
  </div></div>
  <div class="flex flex-wrap gap-2">${qb('t','📄','Texto',bad,'Copiar texto de abordagem (direto e informal)')}${qb('a','🎤','Áudio',bad,'Copiar roteiro de áudio (senhor/senhora)')}${qb('g','🛡️','Antigolpe',bad,'Copiar mensagem antigolpe (follow-up de 30 min a 2 h)')}${qb('wt','💬','WhatsApp direto',bad||nw,'Abrir WhatsApp com a abordagem pronta')}</div>
  <div class="space-y-3">${grp('Status Contato','qs',STATUS,c.status,c.dt1,false)}${grp('Resultado Fechamento','qr',RESULT,c.resultado,c.dt2,bad)}</div>
  <div class="flex items-center justify-between gap-2"><button class="btn" data-qn="-1" ${i===0?'disabled':''} title="Cliente anterior (←)">◀ Anterior</button><span class="mut text-xs">Atalhos: ← → navegar · Esc fechar</span><button class="btn" data-qn="1" ${i>=qList.length-1?'disabled':''} title="Próximo cliente (→)">Próximo ▶</button></div>`;
}
$('#bQ').onclick=qOpen;$('#qX').onclick=()=>$('#qd').close();
$('#qd').addEventListener('click',e=>{
 if(e.target.id==='qd'){$('#qd').close();return}
 if(e.target.closest('[data-qc]')){$('#qd').close();return}
 const n=e.target.closest('[data-qn]');if(n&&!n.disabled){qNav(+n.dataset.qn);return}
 const ch=e.target.closest('[data-qs],[data-qr]');if(ch&&!ch.disabled){qPick(ch);return}
 onMsg(e);
});
document.addEventListener('keydown',e=>{
 if(!$('#qd').open||(e.key!=='ArrowLeft'&&e.key!=='ArrowRight'))return;
 if(/^(SELECT|INPUT|TEXTAREA)$/.test(e.target.tagName))return;
 e.preventDefault();qNav(e.key==='ArrowRight'?1:-1);
});

/* ---------- Nuvem (Supabase): conexão e login ---------- */
const SBK='crm_sb_cfg',SBS='crm_sb_auth';
let sbCfg=null,sb=null,authState='local',sbUser=null,failN=0,lockUntil=0;
try{sbCfg=JSON.parse(localStorage.getItem(SBK)||'null')}catch(e){}
const jwtRole=k=>{try{return JSON.parse(atob(k.split('.')[1].replace(/-/g,'+').replace(/_/g,'/'))).role}catch(e){return''}};
function sbValidate(u,k){
 u=(u||'').trim().replace(/\/+$/,'');k=(k||'').trim();
 if(!/^https:\/\/[a-z0-9.-]+\.[a-z]{2,}$/i.test(u))return{err:'A URL deve começar com https:// (ex.: https://xxxx.supabase.co).'};
 if(/^sb_secret_/.test(k)||jwtRole(k)==='service_role')return{err:'Essa é uma chave secreta (service_role/secret). Use apenas a chave anon/public.'};
 if(!(/^sb_publishable_/.test(k)||k.split('.').length===3))return{err:'A chave anon parece inválida. Copie a chave anon/public do painel do Supabase.'};
 return{u,k};
}
function sbInit(){
 sb=null;if(!sbCfg||!(window.supabase&&supabase.createClient))return false;
 try{
  sb=supabase.createClient(sbCfg.u,sbCfg.k,{auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:false,storageKey:SBS}});
  sb.auth.onAuthStateChange((ev,ses)=>setAuth(ses));
  return true;
 }catch(e){sb=null;return false}
}
function setAuth(ses){sbUser=ses&&ses.user?ses.user:null;authState=!sbCfg?'local':sbUser?'in':'locked';authUI()}
function authUI(){
 const locked=authState==='locked',lg=$('#lg');
 document.body.classList.toggle('locked',locked);
 $('#bCloud').textContent=!sbCfg?'☁ Conectar à nuvem':authState==='in'?'☁ Nuvem ✓':'☁ Nuvem';
 $('#usr').hidden=authState!=='in';$('#usr').textContent=sbUser?sbUser.email:'';$('#bOut').hidden=authState!=='in';
 if(locked){
  if($('#qd').open)$('#qd').close();
  if(!sb)$('#lgMsg').textContent='Não foi possível carregar a biblioteca do Supabase. Verifique a internet ou se o ambiente onde a página está aberta permite conexões externas.';
  if(!lg.open&&!$('#cfgd').open)lg.showModal();
 }else if(lg.open)lg.close();
}
async function sbBoot(){
 if(!sbCfg){setAuth(null);return}
 if(!sbInit()){authState='locked';authUI();return}
 try{const{data}=await sb.auth.getSession();setAuth(data.session)}catch(e){setAuth(null)}
}
const sbErr=e=>{const m=(e&&e.message)||'';
 return /invalid login credentials/i.test(m)?'E-mail ou senha incorretos.':/email not confirmed/i.test(m)?'E-mail ainda não confirmado. Use o link enviado ao seu e-mail.':/rate limit|too many/i.test(m)?'Muitas tentativas. Aguarde alguns minutos.':/fetch|network/i.test(m)?'Falha de conexão com o Supabase.':(m||'Não foi possível entrar.')};
$('#lgF').addEventListener('submit',async e=>{
 e.preventDefault();
 const msg=$('#lgMsg'),btn=$('#lgGo');
 if(Date.now()<lockUntil){msg.textContent=`Muitas tentativas. Aguarde ${Math.ceil((lockUntil-Date.now())/1000)}s.`;return}
 if(!sb&&!sbInit()){msg.textContent='Biblioteca do Supabase indisponível. Verifique a internet e recarregue.';return}
 btn.disabled=true;msg.textContent='Entrando…';
 try{
  const{data,error}=await sb.auth.signInWithPassword({email:$('#lgE').value.trim(),password:$('#lgP').value});
  if(error){failN++;if(failN>=5){lockUntil=Date.now()+30000;failN=0}msg.textContent=sbErr(error)}
  else{failN=0;msg.textContent='';$('#lgF').reset();setAuth(data.session)}
 }catch(err){msg.textContent='Falha de conexão com o Supabase. Confira a URL, a internet e se o ambiente permite conexões externas.'}
 $('#lgP').value='';btn.disabled=false;
});
$('#lg').addEventListener('cancel',e=>e.preventDefault());
$('#lgCfg').onclick=()=>openCfg();
$('#bOut').onclick=async()=>{
 try{if(sb)await sb.auth.signOut({scope:'local'})}catch(e){}
 try{localStorage.removeItem(SBS)}catch(e){}
 setAuth(null);toast('Sessão encerrada.');
};
function openCfg(){$('#sbU').value=sbCfg?sbCfg.u:'';$('#sbK').value=sbCfg?sbCfg.k:'';$('#sbK').type='password';$('#sbShow').textContent='Mostrar';$('#sbMsg').textContent='';$('#cfgd').showModal()}
const cfgMsg=(t,ok)=>{const m=$('#sbMsg');m.textContent=t;m.style.color=ok?'var(--acc)':'#dc2626'};
$('#bCloud').onclick=openCfg;$('#sbCancel').onclick=()=>$('#cfgd').close();
$('#sbShow').onclick=()=>{const k=$('#sbK'),h=k.type==='password';k.type=h?'text':'password';$('#sbShow').textContent=h?'Ocultar':'Mostrar'};
$('#sbTest').onclick=async()=>{
 const v=sbValidate($('#sbU').value,$('#sbK').value);if(v.err){cfgMsg(v.err);return}
 cfgMsg('Testando…',true);
 const ac=new AbortController(),t=setTimeout(()=>ac.abort(),6000);
 try{
  const r=await fetch(v.u+'/auth/v1/health',{headers:{apikey:v.k},signal:ac.signal});
  r.ok?cfgMsg('Conexão OK: o serviço de autenticação respondeu.',true):cfgMsg(`O servidor respondeu (${r.status}), mas recusou a chave. Confira a Anon Key.`);
 }catch(e){cfgMsg('Não foi possível alcançar o servidor. Confira a URL e a internet. Se a página está hospedada em um ambiente restrito, ele pode bloquear conexões externas.')}
 clearTimeout(t);
};
$('#sbSave').onclick=()=>{
 const v=sbValidate($('#sbU').value,$('#sbK').value);if(v.err){cfgMsg(v.err);return}
 if(sbCfg&&(sbCfg.u!==v.u||sbCfg.k!==v.k))try{localStorage.removeItem(SBS)}catch(e){}
 sbCfg={u:v.u,k:v.k};
 try{localStorage.setItem(SBK,JSON.stringify(sbCfg))}catch(e){cfgMsg('Não foi possível salvar no navegador (armazenamento bloqueado).');return}
 $('#cfgd').close();toast('Credenciais salvas. Faça login para continuar.');sbBoot();
};
$('#sbOff').onclick=async()=>{
 if(!confirm('Desconectar da nuvem e voltar ao modo local? As credenciais salvas serão apagadas deste navegador.'))return;
 try{if(sb)await sb.auth.signOut({scope:'local'})}catch(e){}
 try{localStorage.removeItem(SBK);localStorage.removeItem(SBS)}catch(e){}
 sbCfg=null;sb=null;$('#cfgd').close();setAuth(null);toast('Modo local ativado.');
};

/* ---------- Render ---------- */
const FILTERS={posse:isPosse,pend:c=>!c.status,feito:c=>c.status==='Contato feito',aguard:c=>c.status==='Aguardando retorno',erro:c=>c.status==='Contato errado',fech:c=>/^realizado (consignado|refinanciamento)/.test(norm(c.resultado)),res:c=>!!c.resultado,ouro:c=>score(c)>=80,prata:c=>score(c)>=50&&score(c)<80,bronze:c=>score(c)<50,ret:c=>{const d=toDate(c.ret);return !!d&&d<=new Date()}};
function view(){
 const q=norm($('#q').value),qd=digits(q),f=$('#f').value;
 return sorted(clients).filter(c=>{
  if(FILTERS[f]&&!FILTERS[f](c))return false;
  if(!q)return true;
  return norm([c.nome,c.cpf,c.t1,c.t2,c.cidade,c.ben].join(' ')).includes(q)||(qd.length>=3&&digits(c.cpf+c.t1+c.t2).includes(qd));
 });
}
const opts=(list,v)=>['',...list,...(v&&!list.includes(v)?[v]:[])].map(o=>`<option value="${esc(o)}" ${o===v?'selected':''}>${o?esc(o):'—'}</option>`).join('');
function render(){
 renderAlerts();sig=sigNow();qBtn();
 const tot=clients.length,n=f=>clients.filter(f).length,pend=n(FILTERS.pend);
 const k=[['Total de clientes',tot],['Clientes de Posse',n(isPosse)],['Clientes Selo Ouro 🥇',n(FILTERS.ouro),1],['Pendentes',pend],['Contato feito',n(FILTERS.feito)],['Aguardando retorno',n(FILTERS.aguard)],['Contato errado',n(FILTERS.erro)],['Fechamentos',n(FILTERS.fech)],['Margem total',brl(clients.reduce((s,c)=>s+c.margem,0))]];
 const w=tot?(tot-pend)/tot*100:0;
 $('#kpis').innerHTML=k.map(x=>`<div class="card px-3 py-2"${x[2]?' data-kf="ouro" title="Clique para filtrar os clientes Selo Ouro" style="cursor:pointer;border-color:var(--goldln);background:var(--gold)"':''}><p class="mut text-xs">${x[0]}</p><p class="text-lg font-bold leading-tight">${x[1]}</p></div>`).join('')
  +(tot?`<div class="card p-4 col-span-2 lg:col-span-4 xl:col-span-9"><div class="flex justify-between text-sm mb-2"><span class="mut">Clientes trabalhados (com algum status)</span><span class="font-semibold">${Math.round(w)}%</span></div><div style="background:var(--line);height:8px;border-radius:8px"><div style="background:var(--acc);height:8px;border-radius:8px;width:${w}%"></div></div></div>`:'');
 const v=view();
 $('#cnt').textContent=`${v.length} de ${tot} exibidos`;
 $('#empty').classList.toggle('hidden',tot>0);
 $('#tb').classList.toggle('hidden',tot===0);
 $('#rows').innerHTML=v.map(c=>{
  const wa=digits(c.t1),cell=(a,val,dt,list)=>`<td><select class="inp sel no-print" data-${a}="${c.id}" ${a==='r'&&c.status==='Contato errado'?'disabled':''} aria-label="${a==='s'?'Status do contato':'Resultado do fechamento'}">${opts(list,val)}</select><span class="po">${esc(val)}</span>${dt?`<div class="mut" style="font-size:.66rem">${esc(dt)}</div>`:''}${a==='s'?fuBadge(c):''}</td>`;
  return `<tr class="${isPosse(c)?'posse':''} ${c.resultado?'done':''} ${(fuState(c)||{}).st==='due'?'due':''}">
  <td class="font-semibold nm">${esc(c.nome)}</td><td class="nw">${scoreBadge(c)}</td><td class="nw">${esc(c.cpf)}</td>
  <td class="nw">${wa.length>=10?`<a class="underline" target="_blank" rel="noopener" href="https://wa.me/55${wa}">${esc(c.t1)}</a>`:esc(c.t1)}${c.t2?`<div class="mut">${esc(c.t2)}</div>`:''}</td>
  <td class="nw">${esc(c.ben)}${c.esp?`<div class="mut" title="Espécie do benefício">ESP ${esc(c.esp)}</div>`:''}</td>
  <td class="nw font-semibold">${brl(c.margem)}</td><td class="nw">${c.idade||''}</td>
  <td>${isPosse(c)?'<span class="tag">★ Posse</span>':`<span class="chip" title="${esc(c.cidade)}">${esc(c.cidade)}</span>`}</td>
  ${cell('s',c.status,c.dt1,STATUS)}${cell('r',c.resultado,c.dt2,RESULT)}<td>${agCell(c)}</td><td class="no-print">${msgCell(c)}</td>
  <td class="no-print"><button class="mut" data-d="${c.id}" aria-label="Excluir" title="Excluir">✕</button></td></tr>`}).join('');
}

/* ---------- Eventos ---------- */
$('#rows').addEventListener('click',e=>{
 const d=e.target.closest('[data-d]');
 if(d&&confirm('Excluir este cliente?')){clients=clients.filter(c=>c.id!==d.dataset.d);save();render()}
});
function applyChange(c,k,v){
 if(k==='s'){
  if(v==='Contato errado'&&(c.resultado||c.ret)&&!confirm('Contato errado desativa os agendamentos: o resultado e o retorno programado deste cliente serão removidos. Continuar?'))return false;
  c.status=v;c.dt1=v?nowBR():'';
 }else{
  c.resultado=v;c.dt2=v?nowBR():'';c.ret=v?addM(c.dt2,6):'';
  if(v&&!c.status){c.status='Contato feito';c.dt1=nowBR()}
 }
 rules(c);save();render();return true;
}
$('#rows').addEventListener('change',e=>{
 const el=e.target,id=el.dataset.s||el.dataset.r;if(!id)return;
 const c=clients.find(x=>x.id===id);
 if(!applyChange(c,el.dataset.s?'s':'r',el.value))render();
});
async function onMsg(e){
 const f=e.target.closest('[data-fu]');
 if(f){const c=clients.find(x=>x.id===f.dataset.i);c.fu=f.dataset.fu==='undo'?'':nowBR();save();render();return}
 const b=e.target.closest('[data-m]');if(!b||b.disabled)return;
 const c=clients.find(x=>x.id===b.dataset.i),k=b.dataset.m,text=MSG[k[k.length-1]](c);
 const miss=(!cfg.atd||!cfg.emp)?' Preencha seu nome e empresa no topo para completar a mensagem.':'';
 let ok=true;
 if(k[0]==='w')window.open(`https://wa.me/${waNum(c)}?text=${encodeURIComponent(text)}`,'_blank','noopener');
 else ok=await copy(text);
 if(!ok)return;
 const fu=k.endsWith('g');
 toast((k[0]==='w'?'WhatsApp aberto com a mensagem.':'Mensagem copiada.')+(fu?' Follow-up registrado como enviado.':'')+miss);
 if(fu){c.fu=nowBR();save();render()}
}
$('#rows').addEventListener('click',onMsg);$('#alerts').addEventListener('click',onMsg);
$('#kpis').addEventListener('click',e=>{const k=e.target.closest('[data-kf]');if(k){$('#f').value=k.dataset.kf;render()}});
$('#q').addEventListener('input',render);$('#f').addEventListener('change',render);
$('#bClr').onclick=()=>{if(clients.length&&confirm('Apagar toda a base de clientes?')){clients=[];save();render()}};

/* Importação */
const dlg=$('#dlg');let pending=null;
const src=()=>pending||$('#paste').value;
function openImp(){pending=null;$('#paste').value='';$('#file').value='';$('#prev').textContent='';dlg.showModal()}
$('#bImp').onclick=openImp;$('#bCancel').onclick=()=>dlg.close();
function preview(msg){const r=buildClients(src());$('#prev').textContent=(msg||'')+(r.list.length||r.skipped?` ${r.list.length} clientes reconhecidos, ${r.skipped} linhas ignoradas.`:'')}
$('#paste').addEventListener('input',()=>{pending=null;preview()});
$('#file').addEventListener('change',async e=>{
 const f=e.target.files[0];if(!f)return;
 const buf=await f.arrayBuffer();
 if(/\.xlsx?$/i.test(f.name)){
  if(!window.XLSX){toast('Leitor de Excel indisponível. Salve como CSV e tente de novo.');return}
  const wb=XLSX.read(buf,{type:'array'}),name=wb.SheetNames[0];
  pending=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:true,defval:''});
  $('#paste').value='';preview(`Aba "${name}".`);
 }else{
  let t=new TextDecoder('utf-8').decode(buf);
  if(t.includes('\uFFFD'))t=new TextDecoder('windows-1252').decode(buf);
  pending=null;$('#paste').value=t;preview();
 }
});
$('#bGo').onclick=()=>{
 const r=buildClients(src());
 if(!r.list.length){toast('Nenhum cliente reconhecido. Confira os dados.');return}
 if($('#repl').checked)clients=[];
 const seen=new Set(clients.map(c=>digits(c.cpf)).filter(Boolean));let dup=0,add=0;
 for(const c of r.list){const d=digits(c.cpf);if(d&&seen.has(d)){dup++;continue}if(d)seen.add(d);rules(c);learn(c);clients.push(c);add++}
 save();render();dlg.close();
 toast(`${add} importados, ${dup} duplicados (CPF), ${r.skipped} ignorados`);
};

/* Exportar */
function exportCsv(){
 if(!clients.length){toast('Não há dados para exportar');return}
 const h=['Nome','CPF','Telefone 01','Telefone 02','Beneficio','ESP','Margem','Idade','Cidade','Status Contato','Data/Hora Status 1','Resultado Fechamento','Data/Hora Status 2','Retorno Programado','Observações'];
 const q=v=>`"${String(v??'').replace(/"/g,'""')}"`;
 const lines=[h.map(q).join(';')].concat(sorted(clients).map(c=>[c.nome,c.cpf,c.t1,c.t2,c.ben,c.esp,c.margem.toFixed(2).replace('.',','),c.idade||'',c.cidade,c.status,c.dt1,c.resultado,c.dt2,c.ret,c.obs||''].map(q).join(';')));
 const blob=new Blob(['\uFEFF'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});
 const a=document.createElement('a');a.href=URL.createObjectURL(blob);
 a.download=`clientes-consignado-${new Date().toISOString().slice(0,10)}.csv`;
 document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000);
 toast('CSV exportado');
}
$('#bExp').onclick=exportCsv;
$('#bPrt').onclick=()=>{if(!clients.length){toast('Não há dados para imprimir');return}window.print()};

/* Compartilhar */
$('#bShr').onclick=async()=>{
 if(!clients.length){toast('Não há dados para compartilhar');return}
 const s=sorted(clients).slice(0,10);
 const text=`CRM Consignado INSS\n${clients.length} clientes, ${clients.filter(isPosse).length} de Posse\n\nTop 10 por prioridade:\n`+s.map((c,i)=>`${i+1}. ${c.nome} - ${c.cidade} - ${brl(c.margem)}`).join('\n');
 try{
  if(navigator.share){await navigator.share({title:'CRM Consignado INSS',text});return}
 }catch(e){if(e.name==='AbortError')return}
 try{await navigator.clipboard.writeText(text);toast('Resumo copiado para a área de transferência')}
 catch(e){prompt('Copie o resumo:',text)}
};

render();
sbBoot();
</script>
</body>
</html>
