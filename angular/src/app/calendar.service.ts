import { Injectable,inject } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import {Game,ExcelTable,PlayerStats} from './models';
@Injectable({providedIn:'root'}) export class CalendarService {
 private http=inject(HttpClient);
 private api='https://mxflgcwrydqalcvfqhmd.supabase.co/functions/v1/calendar-api';
 games(){return firstValueFrom(this.http.get<{games:Game[]}>(this.api));}
 excel(){return firstValueFrom(this.http.get<{tables:ExcelTable[]}>(this.api+'?resource=excel'));}
 stats(){return firstValueFrom(this.http.get<{stats:PlayerStats[]}>(this.api+'?resource=stats'));}
 edit(data:Record<string,unknown>){return firstValueFrom(this.http.post<{game:Game}>(this.api,data));}
 import(game_id:number,teams:Omit<ExcelTable,'game_id'>[],code:string){return firstValueFrom(this.http.post<{ok:boolean;score_updated:boolean;home_score:number;away_score:number}>(this.api,{action:'import_excel_original',game_id,teams,code}));}
}
