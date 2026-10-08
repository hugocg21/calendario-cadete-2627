import {Component,OnInit,inject} from '@angular/core';
import {CommonModule} from '@angular/common';
import {FormsModule} from '@angular/forms';
import * as XLSX from 'xlsx';
import {CalendarService} from './calendar.service';
import {Game,ExcelTable,PlayerStats,TeamStanding} from './models';
import {TEAMS,LOGOS} from './teams';

type Tab='coach'|'league'|'standings'|'statistics';
const norm=(v:unknown)=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z0-9]/g,'');
@Component({selector:'app-root',standalone:true,imports:[CommonModule,FormsModule],templateUrl:'./app.component.html'})
export class AppComponent implements OnInit{
 private api=inject(CalendarService);
 readonly teams=TEAMS;readonly logos=LOGOS;
 games:Game[]=[];tables:ExcelTable[]=[];playerStats:PlayerStats[]=[];
 tab:Tab='coach';search='';round='';selectedGame='';sync='Cargando calendario…';error='';
 editing:Game|null=null;editCode='';saving=false;importing=false;dark=false;
 editDate='';editTime='';editVenue='';editHomeScore:number|null=null;editAwayScore:number|null=null;
 ngOnInit(){try{this.dark=localStorage.getItem('cm1_2627_theme')==='dark'}catch{}this.refresh();}
 setTab(tab:Tab){this.tab=tab;}
 print(){window.print();}
 toggleTheme(){this.dark=!this.dark;try{localStorage.setItem('cm1_2627_theme',this.dark?'dark':'light')}catch{}}
 async refresh(){this.error='';this.sync='Sincronizando…';try{
 const result=await Promise.all([this.api.games(),this.api.excel(),this.api.stats()]);
 this.games=result[0].games;this.tables=result[1].tables;this.playerStats=result[2].stats;
 this.sync='✓ Sincronizado · '+new Date().toLocaleTimeString('es-ES',{hour:'2-digit',minute:'2-digit'});
 if(!this.selectedGame&&this.importedGames.length)this.selectedGame=String(this.importedGames[0].id);
 }catch(e){this.error=this.message(e);this.sync='Error de sincronización';}}
 message(e:unknown){return e instanceof Error?e.message:String(e);}
 name(code:string){return this.teams[code]||code;}
 played(g:Game){return g.home_score!==null&&g.away_score!==null;}
 get filtered(){const q=norm(this.search);return this.games.filter(g=>(this.tab!=='coach'||g.home==='NAV'||g.away==='NAV')&&(!this.round||g.round===Number(this.round))&&(!q||norm(this.name(g.home)+' '+this.name(g.away)+' '+g.venue).includes(q))).sort((a,b)=>a.round-b.round||a.id-b.id);}
 get rounds(){return Array.from({length:22},(_,i)=>i+1);}
 get importedGames(){const ids=new Set(this.tables.map(t=>t.game_id));return this.games.filter(g=>ids.has(g.id)).sort((a,b)=>a.round-b.round);}
 get selectedTables(){return this.tables.filter(t=>String(t.game_id)===this.selectedGame).sort((a,b)=>{const g=this.games.find(x=>String(x.id)===this.selectedGame);return g?(a.team_code===g.home?-1:b.team_code===g.home?1:0):0;});}
 isTotal(row:string[]){return norm(row[1])==='TOTALES';}
 get standings():TeamStanding[]{
 const data:Record<string,TeamStanding>={};
 for(const code of Object.keys(this.teams))data[code]={team:code,PJ:0,G:0,P:0,PF:0,PC:0,PTS:0,DIF:0};
 for(const g of this.games.filter(x=>this.played(x))){const h=data[g.home],a=data[g.away],hs=g.home_score!,as=g.away_score!;
 h.PJ++;a.PJ++;h.PF+=hs;h.PC+=as;a.PF+=as;a.PC+=hs;
 if(hs>as){h.G++;a.P++;h.PTS+=2;a.PTS++;}else{a.G++;h.P++;a.PTS+=2;h.PTS++;}}
 for(const t of Object.values(data))t.DIF=t.PF-t.PC;
 const sorted=Object.values(data).sort((a,b)=>b.PTS-a.PTS);
 const result:TeamStanding[]=[];
 for(let i=0;i<sorted.length;){let j=i+1;while(j<sorted.length&&sorted[j].PTS===sorted[i].PTS)j++;
 const group=sorted.slice(i,j);
 if(group.length>1){const codes=new Set(group.map(x=>x.team));const mini:Record<string,{pts:number;dif:number;pf:number}>={};
 for(const t of group)mini[t.team]={pts:0,dif:0,pf:0};
 for(const g of this.games.filter(x=>this.played(x)&&codes.has(x.home)&&codes.has(x.away))){const h=mini[g.home],a=mini[g.away],hs=g.home_score!,as=g.away_score!;h.pf+=hs;a.pf+=as;h.dif+=hs-as;a.dif+=as-hs;if(hs>as){h.pts+=2;a.pts++;}else{a.pts+=2;h.pts++;}}
 group.sort((a,b)=>mini[b.team].pts-mini[a.team].pts||mini[b.team].dif-mini[a.team].dif||mini[b.team].pf-mini[a.team].pf||b.DIF-a.DIF||b.PF-a.PF||this.name(a.team).localeCompare(this.name(b.team),'es'));}
 result.push(...group);i=j;}return result;
 }
 get navTotals(){const result=new Map<string,{name:string;pj:number;pts:number;reb:number;ast:number;val:number}>();
 for(const row of this.playerStats.filter(x=>x.team_code==='NAV')){let r=result.get(row.player_name);if(!r){r={name:row.player_name,pj:0,pts:0,reb:0,ast:0,val:0};result.set(row.player_name,r);}r.pj++;r.pts+=row.points;r.reb+=row.rebounds_total;r.ast+=row.assists;r.val+=row.efficiency;}
 return [...result.values()].sort((a,b)=>b.pts-a.pts);}
 openEditor(g:Game){this.editing=g;this.editDate=g.game_date;this.editTime=(g.game_time||'').slice(0,5);this.editVenue=g.venue;this.editHomeScore=g.home_score;this.editAwayScore=g.away_score;this.editCode='';}
 async save(){const g=this.editing;if(!g)return;const hs=this.editHomeScore,awayPoints=this.editAwayScore;
 if((hs===null)!==(awayPoints===null)){alert('Introduce ambos marcadores o deja ambos vacíos.');return;}
 if(hs!==null&&awayPoints!==null&&(!Number.isInteger(hs)||!Number.isInteger(awayPoints)||hs<0||awayPoints<0||hs===awayPoints)){alert('Marcador no válido: introduce puntos enteros, positivos y sin empate.');return;}
 this.saving=true;try{await this.api.edit({code:this.editCode,id:g.id,home_score:hs,away_score:awayPoints,game_date:this.editDate,game_time:this.editTime,venue:this.editVenue.trim()});this.editing=null;await this.refresh();}catch(e){alert(this.message(e));}finally{this.saving=false;}}
 private locate(rows:unknown[][],code:string){const target=norm(this.name(code));return rows.findIndex(row=>norm(row[0])===target);}
 private parse(wb:XLSX.WorkBook,g:Game):Omit<ExcelTable,'game_id'>[]{
 const sheet=wb.Sheets[wb.SheetNames[0]],a=XLSX.utils.sheet_to_json<unknown[]>(sheet,{header:1,defval:'',raw:false});
 return [g.home,g.away].map(code=>{const start=this.locate(a,code);if(start<0)throw Error('No se encontró '+this.name(code)+' en el Excel');
 const head=a.findIndex((r,i)=>i>start&&i<start+15&&norm(r[1])==='NOMBRE'&&norm(r[2])==='MIN');
 if(head<0)throw Error('No se encontraron encabezados de '+this.name(code));
 const rows:string[][]=[];for(let i=head+1;i<a.length;i++){const r=a[i];if(norm(r[1])==='TOTALES'){rows.push(Array.from({length:22},(_,j)=>String(r[j]??'')));break;}
 if(/^\d+$/.test(String(r[0]).trim())&&String(r[1]).trim())rows.push(Array.from({length:22},(_,j)=>String(r[j]??'')));}
 if(!rows.some(r=>norm(r[1])==='TOTALES'))throw Error('No se encontraron totales de '+this.name(code));
 return {team_code:code,title:String(a[start][0]),headers:[a[head-1],a[head]].map(r=>Array.from({length:22},(_,j)=>String(r[j]??''))),rows};});
 }
 private detect(wb:XLSX.WorkBook):Game[]{const sheet=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json<unknown[]>(sheet,{header:1,defval:'',raw:false});
 const titles=rows.slice(0,15).flat().map(v=>norm(v)).filter(v=>v.includes('VS'));
 const result=this.games.filter(g=>titles.some(t=>t.includes(norm(this.name(g.home))+'VS'+norm(this.name(g.away)))));
 if(!result.length)throw Error('No se encontró un partido del calendario con los equipos local y visitante del Excel.');
 return result;}
 async importExcel(event:Event){const input=event.target as HTMLInputElement,file=input.files?.[0];input.value='';if(!file)return;this.importing=true;
 try{const wb=XLSX.read(await file.arrayBuffer(),{type:'array'}),matches=this.detect(wb);let g=matches[0];
 if(matches.length>1){const value=prompt('Varios partidos coinciden. Introduce la jornada:\n'+matches.map(x=>'J'+x.round+' · '+this.name(x.home)+' vs '+this.name(x.away)).join('\n'));if(!value)return;g=matches.find(x=>x.round===Number(value.replace(/\D/g,'')))!;if(!g)throw Error('Jornada no válida');}
 const teams=this.parse(wb,g);
 const scores=teams.map(t=>({code:t.team_code,score:t.rows.find(r=>this.isTotal(r))?.[3]??'?'}));
 if(!confirm('Partido detectado: J'+g.round+' · '+this.name(g.home)+' vs '+this.name(g.away)+'\nResultado Excel: '+scores.map(x=>x.score).join(' - ')+'\n¿Guardar las tablas originales completas?'))return;
 const code=prompt('Código privado de edición');if(!code)return;
 const result=await this.api.import(g.id,teams,code);await this.refresh();this.tab='statistics';this.selectedGame=String(g.id);
 alert('Excel importado correctamente.'+(result.score_updated?' Marcador actualizado automáticamente.':' Marcador existente conservado.'));
 }catch(e){alert('No se ha importado el Excel: '+this.message(e));}finally{this.importing=false;}}
}
