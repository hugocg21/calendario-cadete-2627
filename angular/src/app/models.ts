export interface Game {id:number;round:number;home:string;away:string;game_date:string;game_time:string;venue:string;home_score:number|null;away_score:number|null;updated_at?:string}
export interface ExcelTable {game_id:number;team_code:string;title:string;headers:string[][];rows:string[][]}
export interface PlayerStats {game_id:number;team_code:string;player_name:string;points:number;rebounds_total:number;assists:number;efficiency:number}
export interface TeamStanding {team:string;PJ:number;G:number;P:number;PF:number;PC:number;PTS:number;DIF:number}
