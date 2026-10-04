// Excel export (ExcelJS, styled)
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
// ---- Workbook styling (taken from the reference formatted report) ----
// Summary sheet: merged + centred purple title, Substore / Generated lines, then the table (header on row 6).
// Every sheet: the data is a real Excel Table (banded rows + filter buttons) with all values centred.
const XS={font:'Calibri',size:12,bannerFill:'FF7030A0',bannerText:'FFFFFFFF',numFmt:'#,##0',tableStyle:'TableStyleLight9',title:'PV INVENTORY VARIANCE REPORT'};
let CTX={sub:'',gen:''},TBL=0;
// Writes one sheet: rows are objects keyed by the header names (same data as before), hdr gives column order.
// opt.title => Summary layout (banner + Substore/Generated, table starts at row 6); otherwise table starts at A1.
function sheet(wb,name,rows,widths,hdr,opt){opt=opt||{};
 const ws=wb.addWorksheet(name),n=hdr.length,F={name:XS.font,size:XS.size},hr=opt.title?6:1;
 ws.columns=widths.map(w=>({width:w}));
 if(opt.title){
  for(let c=1;c<=n;c++){const cell=ws.getCell(1,c);cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:XS.bannerFill}};
   cell.font={name:XS.font,size:XS.size,bold:true,color:{argb:XS.bannerText}};cell.alignment={horizontal:'center'}}
  ws.getCell(1,1).value=opt.title;ws.mergeCells(1,1,1,n);
  [['Substore',CTX.sub],['Generated',CTX.gen]].forEach((p,i)=>p.forEach((v,c)=>{const cell=ws.getCell(2+i,c+1);cell.value=v;cell.font=F}))}
 const body=rows.map(o=>hdr.map(h=>{const v=o[h];return v===''||v==null?null:v}));
 // an Excel Table needs at least one data row, so an empty list gets one blank row
 ws.addTable({name:'Table'+(++TBL),ref:'A'+hr,headerRow:true,totalsRow:false,style:{theme:XS.tableStyle,showRowStripes:true},
  columns:hdr.map(h=>({name:h,filterButton:true})),rows:body.length?body:[hdr.map(()=>null)]});
 for(let r=hr;r<=hr+Math.max(body.length,1);r++)for(let c=1;c<=n;c++){const cell=ws.getCell(r,c);cell.font=F;cell.alignment={horizontal:'center'};
  if(r>hr&&opt.numCols&&opt.numCols.includes(c-1)&&typeof cell.value==='number')cell.numFmt=XS.numFmt}
 return ws}
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
 const by=k=>v.filter(r=>r.c===k),mS=by('missSR'),mN=by('missNS'),eS=by('extraSR'),eN=by('extraNS'),oth=by('other'),wb=new ExcelJS.Workbook(),sub=stock.sub||audit.sub||'SUBSTORE';
 const eAll=[...eS,...eN];
 CTX={sub:sub,gen:R.at.toLocaleString('en-GB')};
 const codes=codeSummary(R,stock,audit),H=['Item_Code','Item_Description','Type','System_Stock','Scanned_in_Audit','Matched','Missing','Extra','Result'];
 sheet(wb,'Summary',codes.map(x=>({Item_Code:x.code,Item_Description:x.desc,Type:x.type,System_Stock:x.sys,Scanned_in_Audit:x.aud,Matched:x.matched,Missing:x.miss,Extra:x.extra,Result:x.result})),[11.7,40,13,14,17.8,10,8.9,6.9,17.7],H,{title:XS.title,numCols:[3,4,5,6,7]});
 const N=(base,n)=>base+' ('+n+')';
 sheet(wb,N('Overall Variance',v.length),v.map(r=>({Variance_Type:CATNAME[r.c],Item_Code:r.code,Item_Description:r.desc,Item_SNo:r.serial,Inventory_Status:r.inv,Item_Quality:r.qual,Stock_Qnty:r.sQty,Audit_Qnty:r.aQty,Difference:r.var,Remark:remark(r)})),[32.8,14.8,46,22.8,18.8,17.6,12.2,12.3,11.8,33.3],['Variance_Type','Item_Code','Item_Description','Item_SNo','Inventory_Status','Item_Quality','Stock_Qnty','Audit_Qnty','Difference','Remark']);
 sheet(wb,N('Missing Serialised',mS.length),mS.map(r=>({Item_Code:r.code,Item_Description:r.desc,Item_SNo:r.sno,Inventory_Status:r.inv,Item_Quality:r.qual,Item_Qnty:r.sQty})),[14.8,60.8,22.8,20.8,14.8,11.5],['Item_Code','Item_Description','Item_SNo','Inventory_Status','Item_Quality','Item_Qnty']);
 sheet(wb,N('Missing Non-Serialised',mN.length),mN.map(r=>({Item_Code:r.code,Item_Description:r.desc,Inventory_Status:r.inv,Item_Quality:r.qual,System_Stock:r.sQty,Scanned_in_Audit:r.aQty,Short_By:-r.var,Remark:remark(r)})),[14.8,60.8,18.8,14.8,13.8,16.8,10.8,36.8],['Item_Code','Item_Description','Inventory_Status','Item_Quality','System_Stock','Scanned_in_Audit','Short_By','Remark']);
 sheet(wb,N('Extra in Audit',eAll.length),eAll.map(r=>({Type:r.c==='extraSR'?'Serialised':'Non-serialised',Item_Code:r.code,Item_Description:r.desc,Item_SNo:r.aSno,Item_Quality:r.qual,System_Stock:r.sQty,Scanned_in_Audit:r.aQty,Extra_By:r.aQty-r.sQty,Remark:remark(r)})),[15.8,14.8,60.8,22.8,14.8,14,17.8,10.8,40.8],['Type','Item_Code','Item_Description','Item_SNo','Item_Quality','System_Stock','Scanned_in_Audit','Extra_By','Remark']);
 if(oth.length)sheet(wb,N('Other Issues',oth.length),oth.map(r=>({Item_Code:r.code,Item_SNo:r.serial,Item_Description:r.desc,Stock_Qnty:r.sQty,Audit_Qnty:r.aQty,Remark:remark(r)})),[14.8,22.8,60.8,12.8,12.8,80.8],['Item_Code','Item_SNo','Item_Description','Stock_Qnty','Audit_Qnty','Remark']);
 return{wb,name:'PV_Inventory_Variance_'+sub.replace(/[^\w-]/g,'_')+'_'+stamp(R.at)+'.xlsx'}}
$('#dl').onclick=async()=>{try{if(!window.ExcelJS)throw new Error('Excel export library (ExcelJS) did not load. Check the internet connection and reload.');
 const o=buildWorkbook(S.res,S.stock,S.audit),buf=await o.wb.xlsx.writeBuffer(),a=document.createElement('a');
 a.href=URL.createObjectURL(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));a.download=o.name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
 catch(e){alert('Export failed: '+e.message)}};
