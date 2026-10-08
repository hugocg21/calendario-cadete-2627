(()=>{'use strict';
const API='https://mxflgcwrydqalcvfqhmd.supabase.co/functions/v1/calendar-api';
const escapeHtml=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z0-9]/g,'');
let originals=[];
async function fetchOriginals(){const r=await fetch(API+'?resource=excel',{cache:'no-store'});if(!r.ok)throw Error('Error leyendo tablas');originals=(await r.json()).tables||[];if(view==='statistics')renderOriginals()}
const top=document.createElement('style');top.textContent='.original-table{overflow-x:auto;margin:12px 0 28px;border:1px solid var(--line);border-radius:12px;background:var(--card)}.original-table table{min-width:1400px;white-space:nowrap}.original-table th{font-size:11px;padding:7px 6px}.original-table td{font-size:12px;padding:8px 6px}.original-table tr.total td{font-weight:900;background:var(--home)}.original-actions{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0}.original-actions select{padding:10px;border-radius:9px;background:var(--card);color:var(--text);border:1px solid var(--line)}';document.head.append(top);
const host=document.createElement('section');host.id='originalStats';host.style.display='none';content.after(host);
function makeTable(t){const grouped=t.headers[0],sub=t.headers[1];return '<h3>'+escapeHtml(t.title)+'</h3><div class="original-table"><table><thead><tr>'+grouped.map((x,i)=>'<th>'+escapeHtml(x)+'</th>').join('')+'</tr><tr>'+sub.map(x=>'<th>'+escapeHtml(x)+'</th>').join('')+'</tr></thead><tbody>'+t.rows.map(r=>'<tr class="'+(norm(r[1])==='TOTALES'?'total':'')+'">'+r.map(v=>'<td>'+escapeHtml(v)+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>'}
function renderOriginals(){const selected=document.getElementById('originalGame')?.value||'';host.style.display=view==='statistics'?'block':'none';if(view!=='statistics')return;const ids=[...new Set(originals.map(x=>x.game_id))],gs=games.filter(g=>ids.includes(g.id));host.innerHTML='<h2>📋 Tablas originales del Excel</h2><p>Se muestran las 22 columnas originales, sus porcentajes y las filas de totales, sin recalcular los valores.</p><div class="original-actions"><select id="originalGame"><option value="">Selecciona un partido importado</option>'+gs.map(g=>'<option value="'+g.id+'">J'+g.round+' · '+escapeHtml(T[g.home])+' vs '+escapeHtml(T[g.away])+'</option>').join('')+'</select></div><div id="originalTables"></div>';const sel=document.getElementById('originalGame');sel.value=gs.some(g=>String(g.id)===selected)?selected:gs.length?String(gs[0].id):'';const draw=()=>{const chosen=originals.filter(t=>String(t.game_id)===sel.value).sort((a,b)=>a.team_code==='NAV'?-1:b.team_code==='NAV'?1:0);document.getElementById('originalTables').innerHTML=chosen.length?chosen.map(makeTable).join(''):'<div class="notice">Todavía no hay tablas originales importadas.</div>'};sel.onchange=draw;draw()}
const originalRender=renderStats;
renderStats=function(){originalRender();renderOriginals()};
const file=document.createElement('input');file.type='file';file.accept='.xlsx,.xls';file.hidden=true;document.body.append(file);
const importBtn=document.createElement('button');importBtn.className='primary';importBtn.textContent='📥 Importar tabla original del Excel';importBtn.onclick=()=>file.click();document.querySelector('.tabs').append(importBtn);
function locate(a,team){const target=norm(T[team]);return a.findIndex(row=>{const first=norm(row[0]);return first.length>10&&(first===target||(first.includes(target)&&first.length<target.length+8))})}
function parseWorkbook(wb,g){const sheet=wb.Sheets[wb.SheetNames[0]],a=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false});return [g.home,g.away].map(code=>{const start=locate(a,code);if(start<0)throw Error('No se encontró '+T[code]+' en el Excel');const header=a.findIndex((r,i)=>i>start&&i<start+15&&norm(r[1])==='NOMBRE'&&norm(r[2])==='MIN');if(header<0)throw Error('No se encontraron encabezados de '+T[code]+' (fila de equipo '+(start+1)+'));const rows=[];for(let i=header+1;i<a.length;i++){const r=a[i];if(norm(r[1])==='TOTALES'){rows.push(Array.from({length:22},(_,j)=>String(r[j]??'')));break}if(/^\d+$/.test(String(r[0]).trim())&&String(r[1]).trim())rows.push(Array.from({length:22},(_,j)=>String(r[j]??'')))}if(!rows.some(r=>norm(r[1])==='TOTALES'))throw Error('No se encontraron totales de '+T[code]);return {team_code:code,title:String(a[start][0]),headers:[a[header-1],a[header]].map(r=>Array.from({length:22},(_,j)=>String(r[j]??''))),rows}})}
function detectMatch(wb){
 const a=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:'',raw:false});
 const titles=a.slice(0,15).flat().filter(v=>typeof v==='string'&&/\bvs\.?\b/i.test(v));
 const matches=[];
 for(const title of titles){
  const m=String(title).match(/(?:ESTAD[IÍ]STICAS\s*[-:]\s*)?(.+?)\s+vs\.?\s+(.+?)(?:\s+-\s+Cadete\b|\s+-\s+JUEGOS\b|\s+-\s+FBPA\b|$)/i);
  if(!m)continue;
  const local=norm(m[1]),visitante=norm(m[2]);
  const findCode=name=>Object.keys(T).find(code=>norm(T[code])===name)||Object.keys(T).find(code=>norm(T[code]).includes(name)&&name.length>=10||name.includes(norm(T[code]))&&norm(T[code]).length>=10);
  const home=findCode(local),away=findCode(visitante);
  if(home&&away){const possible=games.filter(g=>g.home===home&&g.away===away);for(const g of possible)if(!matches.some(x=>x.id===g.id))matches.push(g)}
 }
 if(!matches.length)throw Error('No se encontró en el calendario un partido con los equipos local y visitante del Excel.');
 if(matches.length===1)return {game:matches[0],ambiguous:false};
 return {games:matches,ambiguous:true};
}
file.onchange=async()=>{
 const f=file.files[0];file.value='';if(!f)return;
 try{
  if(!window.XLSX)throw Error('No se cargó el lector Excel');
  const wb=XLSX.read(await f.arrayBuffer(),{type:'array'}),detected=detectMatch(wb);
  let g=detected.game;
  if(detected.ambiguous){
   const options=detected.games.map(x=>'J'+x.round+' · '+T[x.home]+' vs '+T[x.away]).join('\\n');
   const selected=prompt('Hay varios partidos con esos equipos. Indica la jornada correcta:\\n'+options);
   if(!selected)return;
   const num=Number((selected.match(/\\d+/)||[])[0]);
   g=detected.games.find(x=>x.round===num);
   if(!g)throw Error('La jornada elegida no corresponde a ese enfrentamiento.');
  }
  const teams=parseWorkbook(wb,g);
  if(!confirm('Partido detectado automáticamente:\\nJornada '+g.round+' · '+T[g.home]+' vs '+T[g.away]+'\\n\\n'+teams.map(t=>t.title+': '+(t.rows.length-1)+' jugadores').join('\\n')+'\\n\\n¿Guardar las tablas completas en Supabase?'))return;
  const code=prompt('Código privado de edición');if(!code)return;
  const r=await fetch(API,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'import_excel_original',code,game_id:g.id,teams})}),j=await r.json();
  if(!r.ok)throw Error(j.error||'Error de importación');
  await fetchOriginals();alert('Jornada '+g.round+': tablas originales guardadas correctamente.');
 }catch(e){alert('No se ha importado el Excel: '+e.message)}
};
const observer=new MutationObserver(()=>{if(view==='statistics'&&host.style.display==='none')renderOriginals();else if(view!=='statistics'&&host.style.display!=='none')host.style.display='none'});observer.observe(document.querySelector('.tabs'),{attributes:true,subtree:true,attributeFilter:['class']});
fetchOriginals().catch(console.error);
})();