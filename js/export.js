// Excel export (SheetJS)
function stamp(){const d=S.res.at,p=n=>String(n).padStart(2,'0');return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate())}

const XR=r=>({Item_Code:r.code,Inventory_Status:r.inv,Item_Description:r.desc,Item_Qnty:r.sQty,Item_Quality:r.qual,Item_SNo:r.type==='SR'?r.sno:'',Match_Type:r.type,Audit_Qnty:r.aQty,Variance_Qty:r.var,Variance_Status:r.status,Variance_Detail:r.detail,Audit_Item_SNo:r.aSno,Audit_Item_Code:r.aCode,Note:r.note});
$('#dl').onclick=()=>{try{const R=S.res,v=R.rows.filter(r=>r.status!=='MATCHED').sort((a,b)=>a.rank-b.rank),s=R.sum,wb=XLSX.utils.book_new();
 const add=(n,rows)=>{const ws=XLSX.utils.json_to_sheet(rows.map(XR),{header:Object.keys(XR(R.rows[0]||{code:'',inv:'',desc:'',qual:'',type:'',sno:'',aSno:'',aCode:'',note:'',status:'',detail:'',sQty:0,aQty:0,var:0}))});ws['!cols']=[14,18,50,10,12,22,10,10,12,20,36,22,16,50].map(w=>({wch:w}));ws['!freeze']={xSplit:0,ySplit:1};XLSX.utils.book_append_sheet(wb,ws,n)};
 add('Variance_Report',v);add('SR_Missing',v.filter(r=>r.type==='SR'&&r.status==='MISSING IN AUDIT'));add('SR_Extra',v.filter(r=>r.type==='SR'&&r.status==='EXTRA IN AUDIT'));add('NonSR_Variance',v.filter(r=>r.type==='NON-SR'));
 const a=[['Substore Name',S.stock.sub||S.audit.sub],['Stock File',S.stock.file],['Audit File',S.audit.file],['Stock Record Count',s.stockRecords],['Audit Record Count',s.auditRecords],['SR Stock Count',s.srStock],['SR Audit Count',s.srAudit],['SR Matched',s.srMatched],['SR Missing',s.srMissing],['SR Extra',s.srExtra],['Non-SR Stock Qty',s.nsStockQty],['Non-SR Audit Qty',s.nsAuditQty],['Non-SR Variance (Audit - Stock)',s.nsVar],['Item Code Mismatch',s.mismatch],['Duplicate Serial Records',s.dup],['Parsing Issues',s.issues],['Total Variances',s.total],['Generated Date/Time',R.at.toLocaleString('en-GB')]];
 const ws=XLSX.utils.aoa_to_sheet([['Metric','Value'],...a]);ws['!cols']=[{wch:34},{wch:60}];XLSX.utils.book_append_sheet(wb,ws,'Summary');
 XLSX.writeFile(wb,'PV_Inventory_Variance_'+(S.stock.sub||S.audit.sub||'SUBSTORE').replace(/[^\w-]/g,'_')+'_'+stamp()+'.xlsx')}catch(e){alert('Export failed: '+e.message)}};

