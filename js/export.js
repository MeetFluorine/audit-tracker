// Excel export (SheetJS)
const rnd=n=>Math.round(n*1e6)/1e6;
function stamp(d){const p=n=>String(n).padStart(2,'0');return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate())}
// Plain category for every non-matched record
function cat(r){
 if(['PARSING ISSUE','ITEM CODE MISMATCH','DUPLICATE SERIAL'].includes(r.status))return'other';
 if(r.type==='SR')return r.status==='MISSING IN AUDIT'?'missSR':r.status==='EXTRA IN AUDIT'?'extraSR':'other';
 if(r.type==='NON-SR'){if(r.status==='MISSING IN AUDIT')return'missNS';if(r.status==='EXTRA IN AUDIT')return'extraNS';return r.var<0?'missNS':'extraNS'}
 return'other'}
const CATNAME={missSR:'Missing in audit - serialised',missNS:'Missing in audit - non-serialised',extraSR:'Extra in audit - serialised',extraNS:'Extra in audit - non-serialised',other:'Other issue'};
const CATORD=['missSR','missNS','extraSR','extraNS','other'];
function remark(r){const c=cat(r);
 if(c==='missNS')return r.status==='MISSING IN AUDIT'?'Not counted in audit':'Audit is short by '+nf(-r.var);
 if(c==='extraNS')return r.status==='EXTRA IN AUDIT'?'Not in system stock':'Audit has '+nf(r.var)+' more than system stock';
 if(c==='missSR')return'Serial not found in audit';if(c==='extraSR')return'Serial not in system stock';
 return[r.detail,r.note].filter(Boolean).join('. ')}
function sheet(wb,name,rows,widths,hdr){const ws=XLSX.utils.json_to_sheet(rows.length?rows:[Object.fromEntries(hdr.map(h=>[h,'']))].slice(0,0),{header:hdr});ws['!cols']=widths.map(w=>({wch:w}));if(rows.length)ws['!autofilter']={ref:ws['!ref']};XLSX.utils.book_append_sheet(wb,ws,name)}
function codeSummary(R,stock,audit){
 const g={},G=c=>g[c]||(g[c]={code:c,desc:'',sr:false,ns:false,sys:0,aud:0,miss:0,extra:0});
 stock.recs.filter(r=>!r.bad).forEach(r=>{const x=G(r.itemCode);x.sys+=r.itemQty;x[r.itemSNo?'sr':'ns']=true;x.desc=x.desc||r.itemDescription});
 audit.recs.filter(r=>!r.bad).forEach(r=>{const x=G(r.itemCode);x.aud+=r.quantity;x[r.srType==='SR'?'sr':'ns']=true;x.desc=x.desc||r.itemDescription});
 R.rows.forEach(r=>{const c=cat(r);
  if(c==='missSR')G(r.code).miss+=r.sQty;else if(c==='extraSR')G(r.aCode||r.code).extra+=r.aQty;
  else if(c==='missNS')G(r.code).miss+=r.status==='MISSING IN AUDIT'?r.sQty:-r.var;
  else if(c==='extraNS')G(r.code).extra+=r.status==='EXTRA IN AUDIT'?r.aQty:r.var;
  else if(r.status==='ITEM CODE MISMATCH'){G(r.code).miss+=r.sQty;G(r.aCode).extra+=r.aQty}});
 return Object.values(g).map(x=>{x.type=x.sr&&x.ns?'Mixed':x.sr?'Serialised':'Non-serialised';x.sys=rnd(x.sys);x.aud=rnd(x.aud);x.miss=rnd(x.miss);x.extra=rnd(x.extra);x.matched=rnd(x.sys-x.miss);x.diff=rnd(x.aud-x.sys);
  x.result=!x.miss&&!x.extra?(x.diff?'Check':'OK'):[x.miss?'Missing '+nf(x.miss):'',x.extra?'Extra '+nf(x.extra):''].filter(Boolean).join(', ');return x})
  .sort((a,b)=>(!!(b.miss||b.extra||b.diff))-(!!(a.miss||a.extra||a.diff))||a.code.localeCompare(b.code))}
function buildWorkbook(R,stock,audit){
 const v=R.rows.filter(r=>r.status!=='MATCHED').map(r=>(r.c=cat(r),r)).sort((a,b)=>CATORD.indexOf(a.c)-CATORD.indexOf(b.c)||a.code.localeCompare(b.code)||(a.serial||'').localeCompare(b.serial||''));
 const by=k=>v.filter(r=>r.c===k),mS=by('missSR'),mN=by('missNS'),eS=by('extraSR'),eN=by('extraNS'),oth=by('other'),wb=XLSX.utils.book_new(),sub=stock.sub||audit.sub||'SUBSTORE';
 const eAll=[...eS,...eN],units=a=>rnd(a.reduce((t,r)=>t+Math.abs(r.var),0));
 const codes=codeSummary(R,stock,audit),H=['Item_Code','Item_Description','Type','System_Stock','Scanned_in_Audit','Matched','Missing','Extra','Difference (Audit - System)','Result'];
 const tot=(label,note,list)=>{const t=k=>rnd(list.reduce((s,x)=>s+x[k],0));return[label,note,'',t('sys'),t('aud'),t('matched'),t('miss'),t('extra'),rnd(t('aud')-t('sys')),'']};
 const ser=codes.filter(x=>x.type==='Serialised'),non=codes.filter(x=>x.type==='Non-serialised'),mix=codes.filter(x=>x.type==='Mixed');
 const A=[['PV INVENTORY VARIANCE REPORT'],['Substore',sub],['Stock file',stock.file],['Audit file',audit.file],['Generated',R.at.toLocaleString('en-GB')],[],
  ['TOTALS: system stock vs scanned in audit'],H,tot('SERIALISED',ser.length+' item codes',ser),tot('NON-SERIALISED',non.length+' item codes',non)];
 if(mix.length)A.push(tot('MIXED',mix.length+' item codes',mix));
 A.push(tot('GRAND TOTAL',codes.length+' item codes',codes),[],
  ['VARIANCE COUNTS'],
  ['Missing, serialised','',mS.length,'serial numbers in system stock but not scanned in audit (sheet: Missing Serialised)'],
  ['Missing, non-serialised','',mN.length,'item codes where audit quantity is below system stock; '+nf(units(mN))+' units short (sheet: Missing Non-Serialised)'],
  ['Extra in audit','',eAll.length,eS.length+' serial numbers not in system stock, plus '+eN.length+' item codes where audit is above system stock ('+nf(units(eN))+' units) (sheet: Extra in Audit)'],
  ['Other issues','',oth.length,'code mismatch, duplicate serial, unreadable row'],
  ['Total variance records','',v.length,'all of the above (sheet: Overall Variance)'],[],
  ['ITEM CODE SUMMARY: for every item code, system stock vs scanned in audit'],H);
 const hdrRow=A.length;
 codes.forEach(x=>A.push([x.code,x.desc,x.type,x.sys,x.aud,x.matched,x.miss,x.extra,x.diff,x.result]));
 const ws=XLSX.utils.aoa_to_sheet(A);ws['!cols']=[18,58,15,13,17,10,10,10,14,22].map(w=>({wch:w}));
 ws['!autofilter']={ref:'A'+hdrRow+':J'+A.length};
 for(let r=0;r<A.length;r++)for(let c=3;c<9;c++){const cell=ws[XLSX.utils.encode_cell({r,c})];if(cell&&cell.t==='n')cell.z=c===8?'+#,##0;-#,##0;0':'#,##0'}
 XLSX.utils.book_append_sheet(wb,ws,'Summary');
 const N=(base,n)=>base+' ('+n+')';
 sheet(wb,N('Overall Variance',v.length),v.map(r=>({Variance_Type:CATNAME[r.c],Item_Code:r.code,Item_Description:r.desc,Item_SNo:r.serial,Inventory_Status:r.inv,Item_Quality:r.qual,Stock_Qnty:r.sQty,Audit_Qnty:r.aQty,Difference:r.var,Remark:remark(r)})),[32,14,58,22,18,14,11,11,11,50],['Variance_Type','Item_Code','Item_Description','Item_SNo','Inventory_Status','Item_Quality','Stock_Qnty','Audit_Qnty','Difference','Remark']);
 sheet(wb,N('Missing Serialised',mS.length),mS.map(r=>({Item_Code:r.code,Item_Description:r.desc,Item_SNo:r.sno,Inventory_Status:r.inv,Item_Quality:r.qual,Item_Qnty:r.sQty})),[14,60,22,20,14,10],['Item_Code','Item_Description','Item_SNo','Inventory_Status','Item_Quality','Item_Qnty']);
 sheet(wb,N('Missing Non-Serialised',mN.length),mN.map(r=>({Item_Code:r.code,Item_Description:r.desc,Inventory_Status:r.inv,Item_Quality:r.qual,System_Stock:r.sQty,Scanned_in_Audit:r.aQty,Short_By:-r.var,Remark:remark(r)})),[14,60,18,14,13,16,10,36],['Item_Code','Item_Description','Inventory_Status','Item_Quality','System_Stock','Scanned_in_Audit','Short_By','Remark']);
 sheet(wb,N('Extra in Audit',eAll.length),eAll.map(r=>({Type:r.c==='extraSR'?'Serialised':'Non-serialised',Item_Code:r.code,Item_Description:r.desc,Item_SNo:r.aSno,Item_Quality:r.qual,System_Stock:r.sQty,Scanned_in_Audit:r.aQty,Extra_By:r.aQty-r.sQty,Remark:remark(r)})),[15,14,60,22,14,13,16,10,40],['Type','Item_Code','Item_Description','Item_SNo','Item_Quality','System_Stock','Scanned_in_Audit','Extra_By','Remark']);
 if(oth.length)sheet(wb,N('Other Issues',oth.length),oth.map(r=>({Item_Code:r.code,Item_SNo:r.serial,Item_Description:r.desc,Stock_Qnty:r.sQty,Audit_Qnty:r.aQty,Remark:remark(r)})),[14,22,60,12,12,80],['Item_Code','Item_SNo','Item_Description','Stock_Qnty','Audit_Qnty','Remark']);
 return{wb,name:'PV_Inventory_Variance_'+sub.replace(/[^\w-]/g,'_')+'_'+stamp(R.at)+'.xlsx'}}
$('#dl').onclick=()=>{try{const o=buildWorkbook(S.res,S.stock,S.audit);XLSX.writeFile(o.wb,o.name)}catch(e){alert('Export failed: '+e.message)}};
