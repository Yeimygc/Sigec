var teamYears=Object.create(null);
function teamEscape(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function parseTeamSheets(sheets){
 const years=Object.create(null);
 const bucket=y=>years[y]||(years[y]={consolidated:[],periods:Object.create(null)});
 const num=v=>typeof v==='number'&&Number.isFinite(v)?v:0;
 for(const [name,rows] of Object.entries(sheets)){
  if(/^20\d\d$/.test(name)){
   const year=bucket(name),heads=rows[1]||[];
   year.consolidated=rows.slice(2).filter(r=>typeof r[0]==='string'&&r[0].trim()&&heads.some((v,i)=>/^Q[1-4]$/.test(v)&&typeof r[i]==='number')).map(r=>{
    const metric={item:r[0].trim()};heads.forEach((v,i)=>{if(/^Q[1-4]$/.test(v))metric[v]=typeof r[i]==='number'?r[i]:null;});
    metric.Total=Object.entries(metric).filter(([k])=>/^Q[1-4]$/.test(k)).reduce((s,[k,v])=>s+num(v),0);return metric;
   });continue;
  }
  const match=name.match(/^(Ene|Abr|Jul|Oct)\w*\s+(20\d{2}|\d{2})\b/i);if(!match)continue;
  const year=bucket(match[2].length===2?'20'+match[2]:match[2]);
  const q={ene:'Q1',abr:'Q2',jul:'Q3',oct:'Q4'}[match[1].toLowerCase()];
  const heads=rows[1]||[],months=rows[2]||[];
  const monthly=months.some(v=>String(v).toLowerCase()==='total');const group=Object.create(null);
  for(let c=1;c<heads.length;c++){
   const person=heads[c];if(typeof person!=='string'||!person.trim()||/^total$/i.test(person.trim()))continue;
   if(monthly&&!months[c])continue;
   const metrics=Object.create(null);
   for(const r of rows.slice(monthly?3:2)){
    if(typeof r[0]!=='string'||!r[0].trim())continue;
    const values=monthly?r.slice(c,c+3):[r[c]];
    if(!values.some(v=>typeof v==='number'))continue;
    const numbers=values.map(v=>typeof v==='number'?v:null);
    metrics[r[0].trim()]={months:numbers,total:numbers.reduce((s,v)=>s+num(v),0)};
   }
   if(Object.keys(metrics).length)group[person.trim()]={months:monthly?months.slice(c,c+3).map(String):["Total del trimestre"],metrics};
  }
  year.periods[q]=group;
 }
 for(const year of Object.values(years)){
  if(!year.consolidated.length){const metrics=new Map();for(const [q,people] of Object.entries(year.periods))for(const person of Object.values(people))for(const [item,m] of Object.entries(person.metrics)){if(!metrics.has(item))metrics.set(item,{item,Total:0});const r=metrics.get(item);r[q]=(r[q]||0)+m.total;r.Total+=m.total;}year.consolidated=[...metrics.values()];}
 }
 if(!Object.values(years).some(y=>y.consolidated.length))throw Error('El archivo no contiene las hojas anuales o trimestrales esperadas.');return years;
}
function selectTeamYear(){
 const y=document.getElementById('teamYear').value;dashboardData=teamYears[y]||{consolidated:[],periods:{}};
 const p=document.getElementById('dashPeriod'),old=p.value;p.replaceChildren(new Option('Acumulado '+y,'Total'));
 const quarters=new Set(Object.keys(dashboardData.periods));dashboardData.consolidated.forEach(r=>Object.keys(r).forEach(k=>{if(/^Q[1-4]$/.test(k)&&r[k]!==null)quarters.add(k);}));
 [...quarters].sort().forEach(q=>p.add(new Option(q+' · '+y,q)));if([...p.options].some(o=>o.value===old))p.value=old;
 populateAssistants();populateMetrics();renderDashboard();
}
document.addEventListener('DOMContentLoaded',()=>{
 const input=document.getElementById('teamExcel'),status=document.getElementById('teamExcelStatus');
 input.addEventListener('change',async()=>{const file=input.files[0];if(!file)return;input.disabled=true;status.textContent='Leyendo base…';
 try{if(typeof XLSX==='undefined')throw Error('No se pudo cargar el lector de Excel. Comprueba tu conexión a Internet.');
 const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});const sheets={};wb.SheetNames.forEach(name=>sheets[name]=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:null,raw:true}));
 const parsed=parseTeamSheets(sheets);teamYears=parsed;const select=document.getElementById('teamYear');select.replaceChildren();Object.keys(parsed).filter(y=>parsed[y].consolidated.length).sort().reverse().forEach(y=>select.add(new Option(y,y)));select.disabled=false;selectTeamYear();
 status.textContent='Base cargada: '+file.name+' · '+new Date().toLocaleString('es-CO')+'. Solo disponible durante esta sesión.';
 }catch(e){status.textContent='No se pudo actualizar: '+e.message+(Object.keys(teamYears).length?' Se conserva la última base válida.':'');}finally{input.disabled=false;input.value='';}});
 document.getElementById('clearTeamExcel').onclick=()=>{teamYears=Object.create(null);dashboardData={consolidated:[],periods:{}};const y=document.getElementById('teamYear');y.replaceChildren();y.disabled=true;document.getElementById('dashPeriod').replaceChildren(new Option('Acumulado','Total'));populateAssistants();populateMetrics();renderDashboard();status.textContent='Debes cargar la base de datos de Gestión de equipo.';};
});
