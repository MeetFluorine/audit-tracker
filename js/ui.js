// Upload handling, results view, filters, table
const $=s=>document.querySelector(s),tick=()=>new Promise(r=>setTimeout(r,0)),esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])),nf=n=>Number(n).toLocaleString('en-IN');
const kb=b=>b>1048576?(b/1048576).toFixed(2)+' MB':(b/1024).toFixed(1)+' KB';
const S={stock:null,audit:null,res:null,tab:'attn',page:0,sort:{k:'rank',d:1},list:[]};
const ORDER=['MISSING IN AUDIT','ITEM CODE MISMATCH','QUANTITY VARIANCE','DUPLICATE SERIAL','EXTRA IN AUDIT','PARSING ISSUE','MATCHED'];
const CLS={'MATCHED':'matched','MISSING IN AUDIT':'missing','EXTRA IN AUDIT':'extra','QUANTITY VARIANCE':'quantity','ITEM CODE MISMATCH':'item','DUPLICATE SERIAL':'duplicate','PARSING ISSUE':'parsing'};
const OTHER=['ITEM CODE MISMATCH','DUPLICATE SERIAL','PARSING ISSUE'];
const TABS=[['attn','Needs attention',r=>r.status!=='MATCHED'],['missing','Missing in audit',r=>r.status==='MISSING IN AUDIT'],['extra','Extra in audit',r=>r.status==='EXTRA IN AUDIT'],['qty','Quantity difference',r=>r.status==='QUANTITY VARIANCE'],['other','Other issues',r=>OTHER.includes(r.status)],['matched','Matched',r=>r.status==='MATCHED'],['all','All records',()=>true]];
const ST2TAB={'MISSING IN AUDIT':'missing','EXTRA IN AUDIT':'extra','QUANTITY VARIANCE':'qty','MATCHED':'matched','ITEM CODE MISMATCH':'other','DUPLICATE SERIAL':'other','PARSING ISSUE':'other'};
const COLS=[['status','Status'],['code','Item code'],['desc','Description'],['serial','Serial no'],['sQty','Stock qty','n'],['aQty','Audit qty','n'],['var','Difference','n'],['inv','Inventory status'],['qual','Quality'],['note','Remark']];
const svg=p=>'<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'+p+'</svg>';
const I={warn:svg('<path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18h.01"/>'),plus:svg('<rect x="7" y="3" width="10" height="18" rx="2"/><path d="M12 9v6M9 12h6"/>'),swap:svg('<path d="M4 8h14l-3-3M20 16H6l3 3"/>'),ok:svg('<circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/>'),info:svg('<circle cx="12" cy="12" r="9"/><path d="M12 7v6M12 16.5h.01"/>'),list:svg('<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M8 9h8M8 13h8M8 17h5"/>'),calc:svg('<rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 7h6M9 12h.01M12 12h.01M15 12h.01M9 16h.01M12 16h.01M15 16h.01"/>')};
const FT=(k,z)=>k==='x'?'<svg width="'+z+'" height="'+z+'" viewBox="0 0 40 40" aria-hidden="true"><rect width="40" height="40" rx="9" fill="#e4f4ea"/><rect x="11" y="8" width="18" height="24" rx="2.5" fill="#1d8a4e"/><path d="M16 15l8 10M24 15l-8 10" stroke="#fff" stroke-width="2.4" stroke-linecap="round"/></svg>':'<svg width="'+z+'" height="'+z+'" viewBox="0 0 40 40" aria-hidden="true"><rect width="40" height="40" rx="9" fill="#fde8e8"/><rect x="11" y="8" width="18" height="24" rx="2.5" fill="#d6342b"/><text x="20" y="23.5" font-size="7.5" font-weight="700" fill="#fff" text-anchor="middle" font-family="Arial,sans-serif">PDF</text></svg>';
$('#ft1').innerHTML=FT('x',46);$('#ft2').innerHTML=FT('p',46);
const CHK='<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" aria-hidden="true"><circle cx="12" cy="12" r="10" fill="currentColor" stroke="none"/><path d="M8 12l3 3 5-6" stroke="#fff"/></svg>';
// state: ok | busy | err
function info(id,kind,name,state,label,stats,msgs,cols){$(id).innerHTML='<div class="frow">'+FT(kind,28)+'<b>'+esc(name)+'</b><span class="st '+(state==='ok'?'ok':state==='err'?'er':'')+'">'+(state==='ok'?CHK:'')+esc(label)+'</span><button class="rm" type="button" aria-label="Remove file" title="Remove file">&times;</button></div>'
 +(stats.length?'<div class="tiles">'+stats.map(s=>'<div><b>'+s[1]+'</b><span>'+s[0]+'</span></div>').join('')+'</div>':'')+(cols?'<details class="cols"><summary>Detected columns</summary>'+esc(cols)+'</details>':'')+msgs.map(m=>'<div class="msg '+m[0]+'">'+esc(m[1])+'</div>').join('')}
function rmFile(n){if(n===1)S.stock=null;else S.audit=null;$('#f'+n).value='';$('#i'+n).innerHTML='';reset();ready()}
$('#i1').onclick=e=>{if(e.target.closest('.rm'))rmFile(1)};$('#i2').onclick=e=>{if(e.target.closest('.rm'))rmFile(2)};
function wire(n,fn){const dz=$('#dz'+n),f=$('#f'+n);$('#b'+n).onclick=()=>f.click();f.onchange=()=>f.files[0]&&fn(f.files[0]);
 dz.ondragover=e=>{e.preventDefault();dz.classList.add('on')};dz.ondragleave=()=>dz.classList.remove('on');dz.ondrop=e=>{e.preventDefault();dz.classList.remove('on');e.dataTransfer.files[0]&&fn(e.dataTransfer.files[0])}}
function ready(){$('#match').disabled=!(S.stock&&S.audit)}
function reset(){S.res=null;$('#res').style.display='none';$('#prog').style.display='none';$('#dl').disabled=true}
function libs(){if(!window.XLSX||!window.pdfjsLib)throw new Error('Required libraries (SheetJS, PDF.js) did not load. Check the internet connection and reload; the files themselves are never uploaded.')}
async function loadStock(file){S.stock=null;reset();ready();info('#i1','x',file.name,'busy','Reading...',[],[]);await tick();
 try{libs();const wb=XLSX.read(await file.arrayBuffer(),{type:'array'}),st=stockFromWb(wb);st.file=file.name;st.size=file.size;S.stock=st;
  const ok=st.recs.filter(r=>!r.bad),sr=ok.filter(r=>r.itemSNo).length,m=[];
  if(st.col.itemSNo===undefined)m.push(['e','Item_SNo column not found. Every row would be treated as Non-SR; serial matching is not possible.']);
  const bad=st.recs.length-ok.length;if(bad)m.push(['w',bad+' stock rows have parsing issues and are flagged, not matched.']);
  const miss=['itemDescription','itemQuality','inventoryStatus'].filter(f=>st.col[f]===undefined);if(miss.length)m.push(['w','Optional columns not found: '+miss.join(', ')]);
  const fatal=st.col.itemSNo===undefined;
  info('#i1','x',file.name,fatal?'err':'ok',fatal?'Cannot match':'File loaded',[['Stock records',nf(st.recs.length)],['Stock serials',nf(sr)],['Quantity items',nf(ok.length-sr)]],m,'Sheet '+st.sheet+': '+st.headers.join(', ')+(st.sub?'. Substore: '+st.sub:''));
  if(fatal)S.stock=null}catch(e){S.stock=null;info('#i1','x',file.name,'err','Could not read',[],[['e',e.message]])}
 chk();ready()}
async function loadAudit(file){S.audit=null;reset();ready();info('#i2','p',file.name,'busy','Reading...',[],[]);await tick();
 try{libs();pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const doc=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;let pg=[],head='';
  for(let n=1;n<=doc.numPages;n++){const t=await(await doc.getPage(n)).getTextContent(),it=t.items.filter(i=>i.str.trim()).map(i=>({x:i.transform[4],y:i.transform[5],w:i.width||0,s:i.str.trim()}));
   if(n===1)head=it.map(i=>i.s).join(' ');pg.push(it);info('#i2','p',file.name,'busy','Reading page '+n+' of '+doc.numPages+'...',[],[]);if(n%3===0)await tick()}
  const recs=auditFromPages(pg);if(!recs.length)throw new Error('PDF was loaded, but no audit records could be detected.');
  const ok=recs.filter(r=>!r.bad),mm=head.match(/Substore Name\s*:\s*(\S+)/i),au={file:file.name,size:file.size,pages:doc.numPages,recs,sub:mm?mm[1]:''};S.audit=au;
  const sr=ok.filter(r=>r.srType==='SR').length,m=[],bad=recs.length-ok.length,mx=Math.max(...recs.map(r=>r.auditNo)),seq=new Set(recs.map(r=>r.auditNo)).size;
  if(bad)m.push(['w',bad+' audit records have parsing issues (see Other issues after matching).']);
  if(mx!==recs.length||seq!==recs.length)m.push(['w','Parsed '+recs.length+' records but highest S.No is '+mx+'. Some rows may not have been detected.']);
  info('#i2','p',file.name,'ok','File loaded',[['Audit records',nf(recs.length)],['Audit serials',nf(sr)],['Quantity items',nf(ok.length-sr)]],m,doc.numPages+' pages'+(au.sub?'. Substore: '+au.sub:''))}
 catch(e){S.audit=null;info('#i2','p',file.name,'err','Could not read',[],[['e',e.message]])}
 chk();ready()}
function chk(){const a=S.stock&&S.stock.sub,b=S.audit&&S.audit.sub;if(a&&b&&norm(a)!==norm(b)){const d=document.createElement('div');d.className='msg w';d.textContent='Substore differs: stock report says '+a+', audit PDF says '+b+'.';$('#i2').appendChild(d)}}
async function step(l,p){$('#pl').textContent=l;$('#pb').style.width=p+'%';await tick()}
$('#match').onclick=async()=>{$('#prog').style.display='block';$('#match').disabled=true;$('#res').style.display='none';
 try{await step('Normalizing identifiers...',20);await step('Separating SR and Non-SR records...',40);await step('Matching serial numbers and item codes...',60);
  const r=matchAll(S.stock,S.audit);await step('Calculating quantity variance...',85);S.res=r;S.res.at=new Date();
  r.rows.forEach(x=>{x.rank=ORDER.indexOf(x.status);x.serial=x.sno||x.aSno;x.blob=(x.code+' '+x.serial+' '+x.desc+' '+x.status+' '+x.detail+' '+x.qual+' '+x.inv+' '+x.note).toLowerCase()});
  S.tab=r.sum.total?'attn':'all';S.sort={k:'rank',d:1};$('#fsort').value='rank:1';populate();summary();tabs();S.page=0;render();$('#res').style.display='block';$('#dl').disabled=false;await step('Complete',100);proc();$('#res').scrollIntoView({behavior:'smooth'})}
 catch(e){await step('Error: '+e.message,0)}ready()};
function summary(){const s=S.res.sum,n=nf,tot=S.res.rows.length,pct=tot?s.matches/tot*100:0;
 $('#rbar').style.width=pct+'%';$('#rcnt').textContent=n(s.matches)+' / '+n(tot);$('#rpct').textContent=pct.toFixed(1)+'% matched';
 const kp=(c,ic,v,l)=>'<div class="kp '+c+'"><span class="ic">'+ic+'</span><div><b>'+n(v)+'</b><span>'+l+'</span></div></div>';
 $('#kpis').innerHTML=kp('r',I.warn,s.missing,'Missing in audit')+kp('o',I.plus,s.extra,'Extra in audit')+kp('p',I.swap,s.qtyVar,'Quantity difference')+kp('g',I.ok,s.matches,'Matched')+kp('t att',I.info,s.total,'Attention (total)');
 $('#strip').innerHTML='<div><span class="ic">'+I.list+'</span><div><small>Serialised items</small><b>'+n(s.srStock)+'</b><span>Total serial numbers (stock)</span></div></div><div><span class="ic">'+I.calc+'</span><div><small>Quantity by item code</small><b>'+n(s.nsCodes)+'</b><span>Item codes compared</span></div></div>'}
function tabs(){$('#tabs').innerHTML=TABS.map(t=>'<button type="button" role="tab" data-t="'+t[0]+'" class="'+(S.tab===t[0]?'on':'')+'" aria-selected="'+(S.tab===t[0])+'">'+t[1]+'<span>('+nf(S.res.rows.filter(t[2]).length)+')</span></button>').join('')}
$('#tabs').onclick=e=>{const b=e.target.closest('button');if(!b)return;S.tab=b.dataset.t;S.page=0;tabs();render()};
$('#fsort').onchange=e=>{const[k,d]=e.target.value.split(':');S.sort={k,d:+d};S.page=0;render()};
$('#ftog').onclick=()=>{const h=$('#fl').hidden=!$('#fl').hidden;$('#ftog').setAttribute('aria-expanded',String(!h))};
function fill(id,set){$(id).innerHTML='<option value="">All</option>'+[...set].sort().map(x=>'<option>'+esc(x)+'</option>').join('')}
function populate(){const q=new Set(),i=new Set();S.res.rows.forEach(r=>{r.qual.split(' / ').forEach(x=>x&&x!=='-'&&q.add(x));r.inv.split(' / ').forEach(x=>x&&i.add(x))});fill('#fq',q);fill('#fi',i);$('#fs').innerHTML='<option value="">All statuses</option>'+ORDER.map(x=>'<option>'+x+'</option>').join('');['#q','#fc','#fn','#ft'].forEach(x=>$(x).value='')}
function filt(){const q=$('#q').value.trim().toLowerCase(),ft=$('#ft').value,fc=norm($('#fc').value),fn=norm($('#fn').value),fq=$('#fq').value,fi=$('#fi').value,tab=TABS.find(t=>t[0]===S.tab)[2],fs=$('#fs').value;
 let l=S.res.rows.filter(r=>tab(r)&&(!fs||r.status===fs)&&(!ft||r.type===ft)&&(!fc||r.code.includes(fc)||r.aCode.includes(fc))&&(!fn||r.serial.includes(fn))&&(!fq||r.qual.split(' / ').includes(fq))&&(!fi||r.inv.split(' / ').includes(fi))&&(!q||r.blob.includes(q)));
 const {k,d}=S.sort;return l.slice().sort((a,b)=>{const x=a[k],y=b[k];return(typeof x==='number'?x-y:String(x).localeCompare(String(y),undefined,{numeric:true}))*d||a.rank-b.rank})}
function render(){S.list=filt();const ps=+$('#ps').value,pages=Math.max(1,Math.ceil(S.list.length/ps));S.page=Math.min(S.page,pages-1);const a=S.page*ps,sl=S.list.slice(a,a+ps),cur=S.sort.k==='rank'?'status':S.sort.k;
 $('#th').innerHTML='<tr>'+COLS.map(c=>'<th data-k="'+c[0]+'" class="'+(c[2]||'')+'" tabindex="0">'+c[1]+(cur===c[0]?(S.sort.d>0?' &#9650;':' &#9660;'):'')+'</th>').join('')+'</tr>';
 const L=i=>' data-label="'+COLS[i][1]+'"';
 $('#tb').innerHTML=sl.map(r=>'<tr><td'+L(0)+'><span class="b s-'+CLS[r.status]+'">'+r.status+'</span></td><td class="m"'+L(1)+'>'+esc(r.code)+'</td><td class="d"'+L(2)+'>'+esc(r.desc)+'</td><td'+L(3)+'>'+esc(r.serial)+'</td><td class="n"'+L(4)+'>'+nf(r.sQty)+'</td><td class="n"'+L(5)+'>'+nf(r.aQty)+'</td><td class="n '+(r.var<0?'neg':r.var>0?'pos':'')+'"'+L(6)+'>'+(r.var>0?'+':'')+nf(r.var)+'</td><td'+L(7)+'>'+esc(r.inv)+'</td><td'+L(8)+'>'+esc(r.qual)+'</td><td class="d"'+L(9)+'>'+esc([r.detail,r.note].filter(Boolean).join(' | '))+'</td></tr>').join('')||'<tr><td style="padding:24px;text-align:center;color:#64748b;display:block">No records match.</td></tr>';
 $('#pc').textContent=S.list.length?'Showing '+nf(a+1)+' to '+nf(a+sl.length)+' of '+nf(S.list.length)+' records':'0 records';$('#pn').textContent='Page '+(S.page+1)+' of '+pages;$('#pp').disabled=S.page===0;$('#pnx').disabled=S.page>=pages-1}
let tm;['q','fc','fn'].forEach(i=>$('#'+i).oninput=()=>{clearTimeout(tm);tm=setTimeout(()=>{S.page=0;render()},120)});
['ft','fq','fi','ps','fs'].forEach(i=>$('#'+i).onchange=()=>{S.page=0;render()});
$('#pp').onclick=()=>{S.page--;render()};$('#pnx').onclick=()=>{S.page++;render()};
$('#th').onclick=e=>{const k=e.target.closest('th');if(!k)return;const key=k.dataset.k;S.sort=S.sort.k===key||(key==='status'&&S.sort.k==='rank')?{k:S.sort.k,d:-S.sort.d}:{k:key,d:1};if(key==='status')S.sort.k='rank';render()};
function proc(){const st=S.stock,au=S.audit,ok=st.recs.filter(r=>!r.bad),aok=au.recs.filter(r=>!r.bad),ss=ok.filter(r=>r.itemSNo).length,as=aok.filter(r=>r.srType==='SR').length,c=(l,a,b,c)=>'<span>'+l+': '+nf(a)+' records | '+nf(b)+' serials | '+nf(c)+' items</span>';
 $('#proc').innerHTML='<h3>Processing summary</h3><div class="pcols"><div><small>Substore</small><b>'+esc(st.sub||au.sub||'-')+'</b></div><div><small>Stock file</small><b>'+esc(st.file)+'</b></div><div><small>Audit file</small><b>'+esc(au.file)+'</b></div><div><small>Processed</small><b>'+S.res.at.toLocaleString('en-GB')+'</b></div><div class="fd"><small>File details</small>'+c('Stock',st.recs.length,ss,ok.length-ss)+c('Audit',au.recs.length,as,aok.length-as)+'</div></div>'}
$('#clr').onclick=()=>{S.stock=S.audit=null;reset();S.list=[];['#i1','#i2','#kpis','#strip','#tabs','#tb','#th','#proc'].forEach(i=>$(i).innerHTML='');['#f1','#f2','#q','#fc','#fn'].forEach(i=>$(i).value='');ready()};
wire(1,loadStock);wire(2,loadAudit);
