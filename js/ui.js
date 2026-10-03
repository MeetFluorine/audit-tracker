// Upload handling, dashboard, filters, table
const $=s=>document.querySelector(s),tick=()=>new Promise(r=>setTimeout(r,0)),esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])),nf=n=>Number(n).toLocaleString('en-IN');
const kb=b=>b>1048576?(b/1048576).toFixed(2)+' MB':(b/1024).toFixed(1)+' KB';
const S={stock:null,audit:null,res:null,view:'var',page:0,sort:{k:'rank',d:1},list:[]};
const ORDER=['MISSING IN AUDIT','ITEM CODE MISMATCH','QUANTITY VARIANCE','DUPLICATE SERIAL','EXTRA IN AUDIT','PARSING ISSUE','MATCHED'];
const CLS={'MATCHED':'matched','MISSING IN AUDIT':'missing','EXTRA IN AUDIT':'extra','QUANTITY VARIANCE':'quantity','ITEM CODE MISMATCH':'item','DUPLICATE SERIAL':'duplicate','PARSING ISSUE':'parsing'};
const COLS=[['status','Status'],['type','Match Type'],['code','Item Code'],['desc','Item Description'],['inv','Inventory Status'],['sQty','Stock Qty','n'],['aQty','Audit Qty','n'],['var','Variance','n'],['qual','Item Quality'],['sno','Item SNo'],['note','Detail / Note']];
function info(id,rows,msgs){$(id).innerHTML='<dl>'+rows.map(r=>'<dt>'+r[0]+'</dt><dd>'+esc(r[1])+'</dd>').join('')+'</dl>'+msgs.map(m=>'<div class="msg '+m[0]+'">'+esc(m[1])+'</div>').join('')}
function wire(n,fn){const dz=$('#dz'+n),f=$('#f'+n);$('#b'+n).onclick=()=>f.click();f.onchange=()=>f.files[0]&&fn(f.files[0]);
 dz.ondragover=e=>{e.preventDefault();dz.classList.add('on')};dz.ondragleave=()=>dz.classList.remove('on');dz.ondrop=e=>{e.preventDefault();dz.classList.remove('on');e.dataTransfer.files[0]&&fn(e.dataTransfer.files[0])}}
function ready(){$('#match').disabled=!(S.stock&&S.audit);}
function libs(){if(!window.XLSX||!window.pdfjsLib)throw new Error('Required libraries (SheetJS, PDF.js) did not load. Check the internet connection and reload; the files themselves are never uploaded.')}
async function loadStock(file){S.stock=null;S.res=null;$('#res').style.display='none';$('#dl').disabled=true;ready();info('#i1',[['File',file.name]],[['w','Reading stock report...']]);await tick();
 try{libs();const wb=XLSX.read(await file.arrayBuffer(),{type:'array'}),st=stockFromWb(wb);st.file=file.name;st.size=file.size;S.stock=st;
  const ok=st.recs.filter(r=>!r.bad),sr=ok.filter(r=>r.itemSNo).length,m=[['o','File loaded']];
  if(st.col.itemSNo===undefined)m.push(['e','Item_SNo column not found. Every row would be treated as Non-SR; serial matching is not possible.']);
  const bad=st.recs.length-ok.length;if(bad)m.push(['w',bad+' stock rows have parsing issues and are flagged, not matched.']);
  const miss=['itemDescription','itemQuality','inventoryStatus'].filter(f=>st.col[f]===undefined);if(miss.length)m.push(['w','Optional columns not found: '+miss.join(', ')]);
  info('#i1',[['File',file.name],['Size',kb(file.size)],['Sheet',st.sheet],['Records',nf(st.recs.length)],['Substore',st.sub||'not detected'],['SR records',nf(sr)],['Non-SR records',nf(ok.length-sr)],['Detected columns',st.headers.join(', ')]],m);
  if(st.col.itemSNo===undefined)S.stock=null}catch(e){S.stock=null;info('#i1',[['File',file.name]],[['e',e.message]])}
 chk();ready()}
async function loadAudit(file){S.audit=null;S.res=null;$('#res').style.display='none';$('#dl').disabled=true;ready();info('#i2',[['File',file.name]],[['w','Reading audit report...']]);await tick();
 try{libs();pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const doc=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;let raw=[],head='';
  for(let n=1;n<=doc.numPages;n++){const t=await(await doc.getPage(n)).getTextContent(),it=t.items.filter(i=>i.str.trim()).map(i=>({x:i.transform[4],y:i.transform[5],s:i.str.trim()}));
   if(n===1)head=it.map(i=>i.s).join(' ');raw.push(...pageRows(it));info('#i2',[['File',file.name],['Reading page',n+' of '+doc.numPages]],[['w','Reading audit report...']]);if(n%3===0)await tick()}
  if(!raw.length)throw new Error('PDF was loaded, but no audit records could be detected.');
  const recs=auditFromRaw(raw),ok=recs.filter(r=>!r.bad),mm=head.match(/Substore Name\s*:\s*(\S+)/i),au={file:file.name,size:file.size,pages:doc.numPages,recs,sub:mm?mm[1]:''};S.audit=au;
  const sr=ok.filter(r=>r.srType==='SR').length,m=[['o','File loaded']],bad=recs.length-ok.length,mx=Math.max(...recs.map(r=>r.auditNo)),seq=new Set(recs.map(r=>r.auditNo)).size;
  if(bad)m.push(['w',bad+' audit records have parsing issues (see the table after matching).']);
  if(mx!==recs.length||seq!==recs.length)m.push(['w','Parsed '+recs.length+' records but highest S.No is '+mx+'. Some rows may not have been detected.']);
  info('#i2',[['File',file.name],['Size',kb(file.size)],['Pages',doc.numPages],['Records detected',nf(recs.length)],['Substore',au.sub||'not detected'],['SR records',nf(sr)],['Non-SR records',nf(ok.length-sr)]],m)}
 catch(e){S.audit=null;info('#i2',[['File',file.name]],[['e',e.message]])}
 chk();ready()}
function chk(){const a=S.stock&&S.stock.sub,b=S.audit&&S.audit.sub;if(a&&b&&norm(a)!==norm(b)){const d=document.createElement('div');d.className='msg w';d.textContent='Substore differs: stock report says '+a+', audit PDF says '+b+'.';$('#i2').appendChild(d)}}
async function step(l,p){$('#pl').textContent=l;$('#pb').style.width=p+'%';await tick()}
$('#match').onclick=async()=>{$('#prog').style.display='block';$('#match').disabled=true;$('#res').style.display='none';
 try{await step('Normalizing identifiers...',20);await step('Separating SR and Non-SR records...',40);await step('Matching serial numbers and item codes...',60);
  const r=matchAll(S.stock,S.audit);await step('Calculating quantity variance...',85);S.res=r;S.res.at=new Date();
  r.rows.forEach(x=>{x.rank=ORDER.indexOf(x.status);x.blob=(x.code+' '+x.sno+' '+x.aSno+' '+x.desc+' '+x.status+' '+x.detail+' '+x.qual+' '+x.inv+' '+x.note).toLowerCase()});
  populate();cards();S.page=0;render();$('#res').style.display='block';$('#dl').disabled=false;await step('Complete',100);proc()}
 catch(e){await step('Error: '+e.message,0)}$('#match').disabled=false};
function cards(){const s=S.res.sum,c=[['Stock Records',nf(s.stockRecords),''],['Audit Records',nf(s.auditRecords),''],['Total Matches',nf(s.matches),'g'],['Missing in Audit',nf(s.missing),'r'],['Extra in Audit',nf(s.extra),'o'],['Quantity Variance (item codes)',nf(s.qtyVar),'a'],['Item Code Mismatch',nf(s.mismatch),'pu'],['Duplicate Serial',nf(s.dup),'bl'],['Parsing Issues',nf(s.issues),''],['SR Stock Count',nf(s.srStock),''],['SR Audit Count',nf(s.srAudit),''],['SR Matched',nf(s.srMatched),'g'],['SR Missing',nf(s.srMissing),'r'],['SR Extra',nf(s.srExtra),'o'],['Non-SR Stock Qty',nf(s.nsStockQty),''],['Non-SR Audit Qty',nf(s.nsAuditQty),''],['Non-SR Variance (Audit - Stock)',(s.nsVar>0?'+':'')+nf(s.nsVar),'a'],['Total Variance Records',nf(s.total),'r']];
 $('#cards').innerHTML=c.map(x=>'<div class="k '+x[2]+'"><b>'+x[1]+'</b><span>'+x[0]+'</span></div>').join('')}
function fill(id,set,all){$(id).innerHTML='<option value="">'+all+'</option>'+[...set].sort().map(x=>'<option>'+esc(x)+'</option>').join('')}
function populate(){const q=new Set(),i=new Set();S.res.rows.forEach(r=>{r.qual.split(' / ').forEach(x=>x&&x!=='-'&&q.add(x));r.inv.split(' / ').forEach(x=>x&&i.add(x))});
 fill('#fq',q,'All');fill('#fi',i,'All');$('#fs').innerHTML='<option value="">All</option>'+ORDER.map(x=>'<option>'+x+'</option>').join('')+''}
function filt(){const q=$('#q').value.trim().toLowerCase(),fs=$('#fs').value,ft=$('#ft').value,fc=norm($('#fc').value),fn=norm($('#fn').value),fq=$('#fq').value,fi=$('#fi').value;
 let l=S.res.rows.filter(r=>(S.view==='all'||r.status!=='MATCHED')&&(!fs||r.status===fs)&&(!ft||r.type===ft)&&(!fc||r.code.includes(fc)||r.aCode.includes(fc))&&(!fn||r.sno.includes(fn)||r.aSno.includes(fn))&&(!fq||r.qual.split(' / ').includes(fq))&&(!fi||r.inv.split(' / ').includes(fi))&&(!q||r.blob.includes(q)));
 const {k,d}=S.sort;l=l.slice().sort((a,b)=>{const x=a[k],y=b[k];return(typeof x==='number'?x-y:String(x).localeCompare(String(y),undefined,{numeric:true}))*d||a.rank-b.rank});return l}
function render(){S.list=filt();const ps=+$('#ps').value,pages=Math.max(1,Math.ceil(S.list.length/ps));S.page=Math.min(S.page,pages-1);const a=S.page*ps,sl=S.list.slice(a,a+ps);
 $('#th').innerHTML='<tr>'+COLS.map(c=>'<th data-k="'+c[0]+'" class="'+(c[2]||'')+'" tabindex="0">'+c[1]+((S.sort.k==='rank'?'status':S.sort.k)===c[0]?(S.sort.d>0?' ▲':' ▼'):'')+'</th>').join('')+'</tr>';
 $('#tb').innerHTML=sl.map(r=>'<tr><td><span class="b s-'+CLS[r.status]+'">'+r.status+'</span></td><td>'+r.type+'</td><td class="m">'+esc(r.code)+'</td><td class="d">'+esc(r.desc)+'</td><td>'+esc(r.inv)+'</td><td class="n">'+nf(r.sQty)+'</td><td class="n">'+nf(r.aQty)+'</td><td class="n"><b>'+(r.var>0?'+':'')+nf(r.var)+'</b></td><td>'+esc(r.qual)+'</td><td class="m">'+esc(r.sno||r.aSno)+'</td><td class="d">'+esc([r.detail,r.note].filter(Boolean).join(' | '))+'</td></tr>').join('')||'<tr><td colspan="11" style="padding:20px;text-align:center;color:#5b6773">No records match the current filters.</td></tr>';
 $('#pc').textContent=S.list.length?'Showing '+nf(a+1)+' to '+nf(a+sl.length)+' of '+nf(S.list.length)+' records':'0 records';$('#pn').textContent='Page '+(S.page+1)+' of '+pages;$('#pp').disabled=S.page===0;$('#pnx').disabled=S.page>=pages-1}
let tm;['q','fc','fn'].forEach(i=>$('#'+i).oninput=()=>{clearTimeout(tm);tm=setTimeout(()=>{S.page=0;render()},120)});
['fs','ft','fq','fi','ps'].forEach(i=>$('#'+i).onchange=()=>{S.page=0;render()});
$('#pp').onclick=()=>{S.page--;render()};$('#pnx').onclick=()=>{S.page++;render()};
$('#vv').onclick=()=>{S.view='var';$('#vv').classList.add('on');$('#va').classList.remove('on');S.page=0;render()};$('#va').onclick=()=>{S.view='all';$('#va').classList.add('on');$('#vv').classList.remove('on');S.page=0;render()};
$('#th').onclick=e=>{const k=e.target.closest('th');if(!k)return;const key=k.dataset.k;S.sort=S.sort.k===key?{k:key,d:-S.sort.d}:{k:key,d:1};if(key==='status')S.sort.k='rank';render()};
function proc(){const s=S.res.sum,d=S.res.at;$('#proc').innerHTML='<b>Processing summary</b><dl><dt>Stock file</dt><dd>'+esc(S.stock.file)+'</dd><dt>Audit file</dt><dd>'+esc(S.audit.file)+'</dd><dt>Substore</dt><dd>'+esc(S.stock.sub||S.audit.sub||'-')+'</dd><dt>Stock records</dt><dd>'+nf(s.stockRecords)+'</dd><dt>Audit records</dt><dd>'+nf(s.auditRecords)+'</dd><dt>Matched</dt><dd>'+nf(s.matches)+'</dd><dt>Variances</dt><dd>'+nf(s.total)+'</dd><dt>Processed</dt><dd>'+d.toLocaleString('en-GB')+'</dd></dl>'}
$('#clr').onclick=()=>{S.stock=S.audit=S.res=null;S.list=[];['#i1','#i2','#cards','#tb','#th','#proc'].forEach(i=>$(i).innerHTML='');['#f1','#f2'].forEach(i=>$(i).value='');['#q','#fc','#fn'].forEach(i=>$(i).value='');$('#res').style.display='none';$('#prog').style.display='none';$('#dl').disabled=true;ready()};
wire(1,loadStock);wire(2,loadAudit);
