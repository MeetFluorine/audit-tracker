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
function info(id,name,size,stats,msgs,cols){$(id).innerHTML='<div class="file"><b>'+esc(name)+'</b><span>'+(size?kb(size):'')+'</span></div>'+(stats.length?'<div class="mini">'+stats.map(s=>'<div><b>'+s[1]+'</b><span>'+s[0]+'</span></div>').join('')+'</div>':'')+(cols?'<details class="cols"><summary>Detected columns</summary>'+esc(cols)+'</details>':'')+msgs.map(m=>'<div class="msg '+m[0]+'">'+esc(m[1])+'</div>').join('')}
function wire(n,fn){const dz=$('#dz'+n),f=$('#f'+n);$('#b'+n).onclick=()=>f.click();f.onchange=()=>f.files[0]&&fn(f.files[0]);
 dz.ondragover=e=>{e.preventDefault();dz.classList.add('on')};dz.ondragleave=()=>dz.classList.remove('on');dz.ondrop=e=>{e.preventDefault();dz.classList.remove('on');e.dataTransfer.files[0]&&fn(e.dataTransfer.files[0])}}
function ready(){$('#match').disabled=!(S.stock&&S.audit);$('#c1').classList.toggle('ok',!!S.stock);$('#c2').classList.toggle('ok',!!S.audit)}
function reset(){S.res=null;$('#res').style.display='none';$('#prog').style.display='none';$('#dl').disabled=true}
function libs(){if(!window.XLSX||!window.pdfjsLib)throw new Error('Required libraries (SheetJS, PDF.js) did not load. Check the internet connection and reload; the files themselves are never uploaded.')}
async function loadStock(file){S.stock=null;reset();ready();info('#i1',file.name,file.size,[],[['w','Reading stock report...']]);await tick();
 try{libs();const wb=XLSX.read(await file.arrayBuffer(),{type:'array'}),st=stockFromWb(wb);st.file=file.name;st.size=file.size;S.stock=st;
  const ok=st.recs.filter(r=>!r.bad),sr=ok.filter(r=>r.itemSNo).length,m=[['o','File loaded']];
  if(st.col.itemSNo===undefined)m.push(['e','Item_SNo column not found. Every row would be treated as Non-SR; serial matching is not possible.']);
  const bad=st.recs.length-ok.length;if(bad)m.push(['w',bad+' stock rows have parsing issues and are flagged, not matched.']);
  const miss=['itemDescription','itemQuality','inventoryStatus'].filter(f=>st.col[f]===undefined);if(miss.length)m.push(['w','Optional columns not found: '+miss.join(', ')]);
  info('#i1',file.name,file.size,[['Records',nf(st.recs.length)],['Serialised',nf(sr)],['Quantity items',nf(ok.length-sr)]],m,'Sheet '+st.sheet+': '+st.headers.join(', ')+(st.sub?'. Substore: '+st.sub:''));
  if(st.col.itemSNo===undefined)S.stock=null}catch(e){S.stock=null;info('#i1',file.name,file.size,[],[['e',e.message]])}
 chk();ready()}
async function loadAudit(file){S.audit=null;reset();ready();info('#i2',file.name,file.size,[],[['w','Reading audit report...']]);await tick();
 try{libs();pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const doc=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;let raw=[],head='';
  for(let n=1;n<=doc.numPages;n++){const t=await(await doc.getPage(n)).getTextContent(),it=t.items.filter(i=>i.str.trim()).map(i=>({x:i.transform[4],y:i.transform[5],s:i.str.trim()}));
   if(n===1)head=it.map(i=>i.s).join(' ');raw.push(...pageRows(it));info('#i2',file.name,file.size,[],[['w','Reading page '+n+' of '+doc.numPages+'...']]);if(n%3===0)await tick()}
  if(!raw.length)throw new Error('PDF was loaded, but no audit records could be detected.');
  const recs=auditFromRaw(raw),ok=recs.filter(r=>!r.bad),mm=head.match(/Substore Name\s*:\s*(\S+)/i),au={file:file.name,size:file.size,pages:doc.numPages,recs,sub:mm?mm[1]:''};S.audit=au;
  const sr=ok.filter(r=>r.srType==='SR').length,m=[['o','File loaded']],bad=recs.length-ok.length,mx=Math.max(...recs.map(r=>r.auditNo)),seq=new Set(recs.map(r=>r.auditNo)).size;
  if(bad)m.push(['w',bad+' audit records have parsing issues (see Other issues after matching).']);
  if(mx!==recs.length||seq!==recs.length)m.push(['w','Parsed '+recs.length+' records but highest S.No is '+mx+'. Some rows may not have been detected.']);
  info('#i2',file.name,file.size,[['Records',nf(recs.length)],['Serialised',nf(sr)],['Quantity items',nf(ok.length-sr)]],m,doc.numPages+' pages'+(au.sub?'. Substore: '+au.sub:''))}
 catch(e){S.audit=null;info('#i2',file.name,file.size,[],[['e',e.message]])}
 chk();ready()}
function chk(){const a=S.stock&&S.stock.sub,b=S.audit&&S.audit.sub;if(a&&b&&norm(a)!==norm(b)){const d=document.createElement('div');d.className='msg w';d.textContent='Substore differs: stock report says '+a+', audit PDF says '+b+'.';$('#i2').appendChild(d)}}
async function step(l,p){$('#pl').textContent=l;$('#pb').style.width=p+'%';await tick()}
$('#match').onclick=async()=>{$('#prog').style.display='block';$('#match').disabled=true;$('#res').style.display='none';
 try{await step('Normalizing identifiers...',20);await step('Separating SR and Non-SR records...',40);await step('Matching serial numbers and item codes...',60);
  const r=matchAll(S.stock,S.audit);await step('Calculating quantity variance...',85);S.res=r;S.res.at=new Date();
  r.rows.forEach(x=>{x.rank=ORDER.indexOf(x.status);x.serial=x.sno||x.aSno;x.blob=(x.code+' '+x.serial+' '+x.desc+' '+x.status+' '+x.detail+' '+x.qual+' '+x.inv+' '+x.note).toLowerCase()});
  S.tab=r.sum.total?'attn':'all';S.sort={k:'rank',d:1};$('#fsort').value='rank:1';populate();summary();tabs();S.page=0;render();$('#res').style.display='block';$('#dl').disabled=false;await step('Complete',100);proc();$('#res').scrollIntoView({behavior:'smooth'})}
 catch(e){await step('Error: '+e.message,0)}ready()};
function summary(){const s=S.res.sum,n=nf,vs=s.nsVar>0?'+':'',c={};S.res.rows.forEach(r=>c[r.status]=(c[r.status]||0)+1);const used=ORDER.filter(x=>c[x]);
 const li=(k,v,t)=>v?'<li><i class="seg-'+k+'"></i><b>'+n(v)+'</b><span>'+t+'</span></li>':'';
 $('#verdict').innerHTML='<div class="verdict"><div><small>Result for '+esc(S.stock.sub||S.audit.sub||'this substore')+'</small><h2>'+(s.total?n(s.total)+'<span>records need attention</span>':'No differences<span>stock and audit agree</span>')+'</h2></div><ul>'
  +li('missing',s.missing,'in stock but not found in the audit')+li('extra',s.extra,'in the audit but not in stock')+li('quantity',s.qtyVar,'item codes with a quantity difference')+li('item',s.mismatch+s.dup+s.issues,'other issues (code mismatch, duplicate serial, unreadable row)')+li('matched',s.matches,'matched')
  +'</ul><div class="db">'+used.map(x=>'<i class="seg-'+CLS[x]+'" style="width:'+(c[x]/S.res.rows.length*100)+'%" title="'+x+': '+n(c[x])+'"></i>').join('')+'</div></div>';
 const pn=(t,rows)=>'<div class="card"><h3>'+t+'</h3><dl>'+rows.map(r=>'<dt>'+r[0]+'</dt><dd>'+r[1]+'</dd>').join('')+'</dl></div>';
 $('#totals').innerHTML=pn('Serialised items',[['In stock report',n(s.srStock)],['In audit',n(s.srAudit)],['Matched',n(s.srMatched)],['Missing in audit',n(s.srMissing)],['Extra in audit',n(s.srExtra)]])
  +pn('Quantity items (by item code)',[['Item codes compared',n(s.nsCodes)],['Stock quantity',n(s.nsStockQty)],['Audit quantity',n(s.nsAuditQty)],['Difference (audit minus stock)',vs+n(s.nsVar)]])}
function tabs(){$('#tabs').innerHTML=TABS.map(t=>'<button type="button" role="tab" data-t="'+t[0]+'" class="'+(S.tab===t[0]?'on':'')+'" aria-selected="'+(S.tab===t[0])+'">'+t[1]+'<span>'+nf(S.res.rows.filter(t[2]).length)+'</span></button>').join('')}
$('#tabs').onclick=e=>{const b=e.target.closest('button');if(!b)return;S.tab=b.dataset.t;S.page=0;tabs();render()};
$('#fsort').onchange=e=>{const[k,d]=e.target.value.split(':');S.sort={k,d:+d};S.page=0;render()};
$('#ftog').onclick=()=>{const h=$('#fl').hidden=!$('#fl').hidden;$('#ftog').setAttribute('aria-expanded',String(!h))};
function fill(id,set){$(id).innerHTML='<option value="">All</option>'+[...set].sort().map(x=>'<option>'+esc(x)+'</option>').join('')}
function populate(){const q=new Set(),i=new Set();S.res.rows.forEach(r=>{r.qual.split(' / ').forEach(x=>x&&x!=='-'&&q.add(x));r.inv.split(' / ').forEach(x=>x&&i.add(x))});fill('#fq',q);fill('#fi',i);['#q','#fc','#fn','#ft'].forEach(x=>$(x).value='')}
function filt(){const q=$('#q').value.trim().toLowerCase(),ft=$('#ft').value,fc=norm($('#fc').value),fn=norm($('#fn').value),fq=$('#fq').value,fi=$('#fi').value,tab=TABS.find(t=>t[0]===S.tab)[2];
 let l=S.res.rows.filter(r=>tab(r)&&(!ft||r.type===ft)&&(!fc||r.code.includes(fc)||r.aCode.includes(fc))&&(!fn||r.serial.includes(fn))&&(!fq||r.qual.split(' / ').includes(fq))&&(!fi||r.inv.split(' / ').includes(fi))&&(!q||r.blob.includes(q)));
 const {k,d}=S.sort;return l.slice().sort((a,b)=>{const x=a[k],y=b[k];return(typeof x==='number'?x-y:String(x).localeCompare(String(y),undefined,{numeric:true}))*d||a.rank-b.rank})}
function render(){S.list=filt();const ps=+$('#ps').value,pages=Math.max(1,Math.ceil(S.list.length/ps));S.page=Math.min(S.page,pages-1);const a=S.page*ps,sl=S.list.slice(a,a+ps),cur=S.sort.k==='rank'?'status':S.sort.k;
 $('#th').innerHTML='<tr>'+COLS.map(c=>'<th data-k="'+c[0]+'" class="'+(c[2]||'')+'" tabindex="0">'+c[1]+(cur===c[0]?(S.sort.d>0?' &#9650;':' &#9660;'):'')+'</th>').join('')+'</tr>';
 const L=i=>' data-label="'+COLS[i][1]+'"';
 $('#tb').innerHTML=sl.map(r=>'<tr><td'+L(0)+'><span class="b s-'+CLS[r.status]+'">'+r.status+'</span></td><td class="m"'+L(1)+'>'+esc(r.code)+'</td><td class="d"'+L(2)+'>'+esc(r.desc)+'</td><td class="m"'+L(3)+'>'+esc(r.serial)+'</td><td class="n"'+L(4)+'>'+nf(r.sQty)+'</td><td class="n"'+L(5)+'>'+nf(r.aQty)+'</td><td class="n"'+L(6)+'><b>'+(r.var>0?'+':'')+nf(r.var)+'</b></td><td'+L(7)+'>'+esc(r.inv)+'</td><td'+L(8)+'>'+esc(r.qual)+'</td><td class="d"'+L(9)+'>'+esc([r.detail,r.note].filter(Boolean).join(' | '))+'</td></tr>').join('')||'<tr><td style="padding:24px;text-align:center;color:#5a6775;display:block">No records match.</td></tr>';
 $('#pc').textContent=S.list.length?'Showing '+nf(a+1)+' to '+nf(a+sl.length)+' of '+nf(S.list.length)+' records':'0 records';$('#pn').textContent='Page '+(S.page+1)+' of '+pages;$('#pp').disabled=S.page===0;$('#pnx').disabled=S.page>=pages-1}
let tm;['q','fc','fn'].forEach(i=>$('#'+i).oninput=()=>{clearTimeout(tm);tm=setTimeout(()=>{S.page=0;render()},120)});
['ft','fq','fi','ps'].forEach(i=>$('#'+i).onchange=()=>{S.page=0;render()});
$('#pp').onclick=()=>{S.page--;render()};$('#pnx').onclick=()=>{S.page++;render()};
$('#th').onclick=e=>{const k=e.target.closest('th');if(!k)return;const key=k.dataset.k;S.sort=S.sort.k===key||(key==='status'&&S.sort.k==='rank')?{k:S.sort.k,d:-S.sort.d}:{k:key,d:1};if(key==='status')S.sort.k='rank';render()};
function proc(){const s=S.res.sum;$('#proc').innerHTML='<b>Processing summary</b><dl><dt>Stock file</dt><dd>'+esc(S.stock.file)+'</dd><dt>Audit file</dt><dd>'+esc(S.audit.file)+'</dd><dt>Substore</dt><dd>'+esc(S.stock.sub||S.audit.sub||'-')+'</dd><dt>Stock records</dt><dd>'+nf(s.stockRecords)+'</dd><dt>Audit records</dt><dd>'+nf(s.auditRecords)+'</dd><dt>Processed</dt><dd>'+S.res.at.toLocaleString('en-GB')+'</dd></dl>'}
$('#clr').onclick=()=>{S.stock=S.audit=null;reset();S.list=[];['#i1','#i2','#verdict','#totals','#tabs','#tb','#th','#proc'].forEach(i=>$(i).innerHTML='');['#f1','#f2','#q','#fc','#fn'].forEach(i=>$(i).value='');ready()};
wire(1,loadStock);wire(2,loadAudit);
