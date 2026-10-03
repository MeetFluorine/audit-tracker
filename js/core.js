// Parsing, normalisation and matching engine (no DOM access)
const NOSER=new Set(['','-','--','NULL','N/A','NA','NONE','NIL']);
const norm=v=>String(v==null?'':v).normalize('NFKC').replace(/[\u200B-\u200D\u2060\uFEFF]/g,'').replace(/\u00A0/g,' ').replace(/\s+/g,' ').trim().toUpperCase();
const hasSer=s=>!NOSER.has(s);
// Rebuild logical audit rows from positioned PDF text. Rows are vertically centred on the S.No, so every
// text item is assigned to the nearest S.No; columns are identified by x position (layout of the i360 report).
function pageRows(items){
 const sn=items.filter(i=>i.x<62&&/^\d{1,5}$/.test(i.s)),snSet=new Set(sn);
 const rows=sn.map(i=>({no:+i.s,y:i.y,c:{dt:[],d:[],m:[],s:[],k:[],t:[],q:[],l:[],u:[],p:[]}}));
 for(const i of items){if(snSet.has(i)||i.x<62)continue;let b=null,bd=1e9;
  for(const r of rows){const d=Math.abs(r.y-i.y);if(d<bd){bd=d;b=r}}
  if(!b||bd>20)continue;const x=i.x;
  b.c[x<100?'dt':x<560?'d':x<588?'m':x<700?'s':x<770?'k':x<815?'t':x<860?'q':x<895?'l':x<920?'u':'p'].push(i)}
 return rows.map(r=>{const j=(a,v)=>a.sort((p,q)=>v?(q.y-p.y||p.x-q.x):p.x-q.x).map(i=>i.s).join(' ').trim(),c=r.c;
  return{auditNo:r.no,date:j(c.dt,1),desc:j(c.d,1),model:j(c.m),serial:j(c.s),code:j(c.k),type:j(c.t),qty:j(c.q),quality:j(c.l),uom:j(c.u),cat:j(c.p,1)}})}
function auditFromRaw(raw){return raw.map(r=>{const t=r.type.replace(/\s+/g,' ').toUpperCase(),sr=t==='SR'?'SR':t==='NON SR'?'NON-SR':'',ser=norm(r.serial),code=norm(r.code),q=parseFloat(r.qty.replace(/,/g,''));
 const bad=!sr?'Unrecognised SR / Non SR value "'+r.type+'"':(!code||code==='-')?'Missing Item Code':isNaN(q)?'Invalid quantity "'+r.qty+'"':(sr==='SR'&&!hasSer(ser))?'SR record without serial number':'';
 return{auditNo:r.auditNo,auditDate:r.date,itemDescription:r.desc,serialNumber:hasSer(ser)?ser:'',itemCode:code,srType:sr,quantity:isNaN(q)?0:q,quality:r.quality,uom:r.uom,productCategory:r.cat,bad}})}
function stockFromWb(wb){
 const key=h=>String(h).toLowerCase().replace(/[^a-z0-9]/g,'');
 const A={itemCode:['itemcode'],itemQty:['itemqnty','itemqty','itemquantity'],itemSNo:['itemsno','itemserialno','itemserialnumber','serialnumber','serialno'],itemDescription:['itemdescription'],inventoryStatus:['inventorystatus'],itemQuality:['itemquality'],itemUom:['itemuom'],itemLotNo:['itemlotno'],location:['location'],sub:['substorename']};
 for(const name of wb.SheetNames){
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:false,defval:''});
  for(let h=0;h<Math.min(rows.length,30);h++){
   const keys=rows[h].map(key),col={};
   for(const f in A){const i=keys.findIndex(k=>A[f].includes(k));if(i>=0)col[f]=i}
   if(col.itemCode===undefined||col.itemQty===undefined)continue;
   const g=(r,f)=>col[f]===undefined?'':String(r[col[f]]==null?'':r[col[f]]).trim();
   const recs=[],subs={};
   for(let i=h+1;i<rows.length;i++){const r=rows[i],code=norm(g(r,'itemCode')),qs=g(r,'itemQty'),sno=norm(g(r,'itemSNo'));
    if(!code&&!qs&&!hasSer(sno))continue;
    const q=parseFloat(qs.replace(/,/g,''));
    recs.push({row:i+1,itemCode:code,inventoryStatus:g(r,'inventoryStatus'),itemDescription:g(r,'itemDescription'),itemUom:g(r,'itemUom'),itemQty:isNaN(q)?0:q,itemLotNo:g(r,'itemLotNo'),itemQuality:g(r,'itemQuality'),itemSNo:hasSer(sno)?sno:'',location:g(r,'location'),bad:!code?'Missing Item_Code':isNaN(q)?'Invalid Item_Qnty "'+qs+'"':''});
    const s=g(r,'sub');if(s)subs[s]=(subs[s]||0)+1}
   const sub=Object.keys(subs).sort((a,b)=>subs[b]-subs[a])[0]||'';
   return{sheet:name,col,headers:Object.keys(col).map(f=>rows[h][col[f]]),recs,sub};
  }}
 throw new Error('Unable to identify the Item_Code and Item_Qnty columns in any sheet of this Excel file.');
}
function matchAll(st,au){
 const out=[],rd=n=>Math.round(n*1e6)/1e6;
 const mk=o=>{const r=Object.assign({status:'',detail:'',type:'SR',code:'',desc:'',inv:'',sQty:0,aQty:0,qual:'',sno:'',aCode:'',aSno:'',note:''},o);r.var=r.status==='PARSING ISSUE'?0:rd(r.aQty-r.sQty);out.push(r)};
 const grp=(a,k)=>{const m=new Map();for(const r of a){const x=m.get(r[k]);x?x.push(r):m.set(r[k],[r])}return m};
 const sSR=[],sNS=[],aSR=[],aNS=[];
 for(const r of st.recs){if(r.bad)mk({status:'PARSING ISSUE',type:r.itemSNo?'SR':'NON-SR',code:r.itemCode,desc:r.itemDescription,inv:r.inventoryStatus,sQty:r.itemQty,qual:r.itemQuality,sno:r.itemSNo,note:'Stock row '+r.row+': '+r.bad});else(r.itemSNo?sSR:sNS).push(r)}
 for(const r of au.recs){if(r.bad)mk({status:'PARSING ISSUE',type:r.srType||'?',code:r.itemCode,desc:r.itemDescription,aQty:r.quantity,qual:r.quality,aSno:r.serialNumber,aCode:r.itemCode,note:'Audit S.No '+r.auditNo+': '+r.bad});else(r.srType==='SR'?aSR:aNS).push(r)}
 const sM=grp(sSR,'itemSNo'),aM=grp(aSR,'serialNumber');
 for(const r of sSR){const a=aM.get(r.itemSNo),sd=sM.get(r.itemSNo).length>1,ad=a?a.length>1:false;
  const b={type:'SR',code:r.itemCode,desc:r.itemDescription,inv:r.inventoryStatus,sQty:r.itemQty,qual:r.itemQuality,sno:r.itemSNo};
  if(!a){mk(Object.assign(b,sd?{status:'DUPLICATE SERIAL',detail:'DUPLICATE SERIAL IN STOCK ('+sM.get(r.itemSNo).length+'x); NOT FOUND IN AUDIT'}:{status:'MISSING IN AUDIT',detail:'SERIAL NOT FOUND IN AUDIT'}));continue}
  b.aSno=a[0].serialNumber;b.aCode=a[0].itemCode;b.aQty=a.reduce((t,x)=>t+x.quantity,0);
  if(sd||ad){b.status='DUPLICATE SERIAL';b.detail=[sd?'DUPLICATE SERIAL IN STOCK ('+sM.get(r.itemSNo).length+'x)':'',ad?'DUPLICATE SERIAL IN AUDIT ('+a.length+'x)':''].filter(Boolean).join('; ')}
  else if(a[0].itemCode!==r.itemCode){b.status='ITEM CODE MISMATCH';b.detail='SERIAL MATCH / ITEM CODE MISMATCH (stock '+r.itemCode+', audit '+a[0].itemCode+')'}
  else if(rd(b.aQty)!==rd(r.itemQty)){b.status='QUANTITY VARIANCE';b.detail='SERIAL MATCHED, QUANTITY DIFFERS'}
  else b.status='MATCHED';
  mk(b)}
 for(const r of aSR){if(sM.has(r.serialNumber))continue;const n=aM.get(r.serialNumber).length;
  mk({status:n>1?'DUPLICATE SERIAL':'EXTRA IN AUDIT',detail:n>1?'DUPLICATE SERIAL IN AUDIT ('+n+'x); NOT IN STOCK':'SERIAL NOT IN STOCK',type:'SR',code:r.itemCode,desc:r.itemDescription,aQty:r.quantity,qual:r.quality,aCode:r.itemCode,aSno:r.serialNumber})}
 const sG=grp(sNS,'itemCode'),aG=grp(aNS,'itemCode'),uq=(a,f)=>[...new Set(a.map(r=>r[f]).filter(x=>x&&x!=='-'))];
 for(const c of [...sG.keys(),...[...aG.keys()].filter(k=>!sG.has(k))]){
  const s=sG.get(c)||[],a=aG.get(c)||[],sQ=s.reduce((t,r)=>t+r.itemQty,0),aQ=a.reduce((t,r)=>t+r.quantity,0),nt=[];
  for(const[l,v]of[['description',uq(s,'itemDescription')],['quality',uq(s,'itemQuality')],['UOM',uq(s,'itemUom')],['audit UOM',uq(a,'uom')]])if(v.length>1)nt.push(l+': '+v.join(' | '));
  const note=[s.length>1?s.length+' stock rows aggregated':'',a.length>1?a.length+' audit rows aggregated':'',nt.length?'MULTIPLE ATTRIBUTE VALUES ('+nt.join('; ')+')':''].filter(Boolean).join('; ');
  const d=rd(aQ-sQ);let status,detail;
  if(!s.length){status='EXTRA IN AUDIT';detail='ONLY IN AUDIT'}else if(!a.length){status='MISSING IN AUDIT';detail='NOT FOUND IN AUDIT'}
  else if(d===0){status='MATCHED';detail=''}else{status='QUANTITY VARIANCE';detail=d<0?'SHORT IN AUDIT':'EXCESS IN AUDIT'}
  mk({status,detail,type:'NON-SR',code:c,desc:(s[0]||a[0]).itemDescription,inv:uq(s,'inventoryStatus').join(' / '),sQty:sQ,aQty:aQ,qual:(s.length?uq(s,'itemQuality'):uq(a,'quality')).join(' / '),aCode:a.length?c:'',note})}
 const R=(f)=>out.filter(f).length,nsS=sNS.reduce((t,r)=>t+r.itemQty,0),nsA=aNS.reduce((t,r)=>t+r.quantity,0);
 const sum={stockRecords:st.recs.length,auditRecords:au.recs.length,srStock:sSR.length,srAudit:aSR.length,srMatched:R(r=>r.type==='SR'&&r.status==='MATCHED'),srMissing:R(r=>r.type==='SR'&&r.status==='MISSING IN AUDIT'),srExtra:R(r=>r.type==='SR'&&r.status==='EXTRA IN AUDIT'),
  nsStockQty:rd(nsS),nsAuditQty:rd(nsA),nsVar:rd(nsA-nsS),matches:R(r=>r.status==='MATCHED'),missing:R(r=>r.status==='MISSING IN AUDIT'),extra:R(r=>r.status==='EXTRA IN AUDIT'),qtyVar:R(r=>r.status==='QUANTITY VARIANCE'),mismatch:R(r=>r.status==='ITEM CODE MISMATCH'),dup:R(r=>r.status==='DUPLICATE SERIAL'),issues:R(r=>r.status==='PARSING ISSUE'),total:R(r=>r.status!=='MATCHED'),nsCodes:sG.size+[...aG.keys()].filter(k=>!sG.has(k)).length};
 return{rows:out,sum};
}
