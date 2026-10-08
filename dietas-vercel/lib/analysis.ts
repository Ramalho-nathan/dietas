export type DietRecord = { id:string; row:number; sheet:string; hospital:string; floor:string; request:string; type:string; observation:string; meal:string; canceled:boolean; day:string; requested:number; released:number|null; departed:number|null; delivered:number|null; systemMinutes:number|null; invalid:boolean; inferred:boolean };
export type SourceSheet={name:string;rows:unknown[][]};
export type ImportResult={records:DietRecord[]; rejected:{sheet:string;row:number;reason:string}[]; warnings:string[]; filename:string; sourceSheets?:SourceSheet[]};
export const normalize=(s:unknown)=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
const DAY=86400;
function dateSeconds(v:unknown):number|null {
 if(v instanceof Date&&!isNaN(v.getTime()))return Date.UTC(v.getUTCFullYear(),v.getUTCMonth(),v.getUTCDate())/1000;
 if(typeof v==='number'&&isFinite(v))return (Math.floor(v)-25569)*DAY;
 const t=String(v??'').trim();const br=t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);const iso=t.match(/^(\d{4})-(\d{2})-(\d{2})/);
 if(!br&&!iso)return null;const [y,m,d]=br?[+br[3],+br[2],+br[1]]:[+iso![1],+iso![2],+iso![3]];
 const a=new Date(Date.UTC(y,m-1,d));return a.getUTCFullYear()===y&&a.getUTCMonth()===m-1&&a.getUTCDate()===d?a.getTime()/1000:null;
}
function seconds(v:unknown,duration=false):number|null {
 if(v===null||v===undefined||v==='')return null;
 if(v instanceof Date&&!isNaN(v.getTime()))return v.getUTCHours()*3600+v.getUTCMinutes()*60+v.getUTCSeconds();
 if(typeof v==='number'&&isFinite(v)&&v>=0)return Math.round((duration?v:v%1)*DAY);
 const m=String(v).trim().match(/^(\d{1,3}):(\d{2})(?::(\d{2}))?$/);if(!m||+m[2]>59||+(m[3]||0)>59||(!duration&&+m[1]>23))return null;
 return +m[1]*3600+ +m[2]*60+ +(m[3]||0);
}
export const dateKey=(s:number)=>new Date(s*1000).toISOString().slice(0,10);
export const clock=(s:number|null)=>s===null?'—':new Date(s*1000).toISOString().slice(11,19);
export const brDate=(key:string)=>key.split('-').reverse().join('/');
export function mealFromObservation(v:string){
 const t=v.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase();const hit:string[]=[];
 const defs:[string,RegExp][]=[['Café da manhã',/CAFE (DA MANHA|COMPLETO)/],['Almoço',/ALMOCO/],['Lanche da tarde',/LANCHE DA TARDE/],['Jantar',/JANTAR/],['Ceia',/CEIA/],['Lanche sem período',/LANCHE/]];
 for(const [name,regex]of defs)if(regex.test(t)&&!(name==='Lanche sem período'&&hit.includes('Lanche da tarde')))hit.push(name);
 return hit.length===1?hit[0]:hit.length>1?'Múltiplas refeições':'Não informada';
}
export function parseSheets(sheets:{name:string;rows:unknown[][]}[],filename:string):ImportResult{
 const records:DietRecord[]=[];const rejected:ImportResult['rejected']=[];const warnings=new Set<string>();let found=false;
 for(const sh of sheets){const hi=sh.rows.slice(0,15).findIndex(r=>r.some(c=>normalize(c)==='datadasolicitacao')&&r.some(c=>normalize(c)==='horadasolicitacao'));if(hi<0)continue;
 const keys=sh.rows[hi].map(normalize);const required=['datadasolicitacao','horadasolicitacao','unidade','andar','solicitacao','tipo'];const missing=required.filter(k=>!keys.includes(k));
 if(missing.length)throw new Error('Colunas obrigatórias ausentes: '+missing.join(', ')+'. Use o relatório original de registros.');found=true;
 for(const k of ['horaliberacao','saidaparaentrega','entregadadieta','tempodoatendimento'])if(!keys.includes(k))warnings.add('Coluna ausente: '+k+'. Os indicadores dessa etapa ficarão indisponíveis.');
 for(let i=hi+1;i<sh.rows.length;i++){const row=sh.rows[i];if(!row.some(v=>v!==null&&v!==undefined&&v!==''))continue;
 const get=(key:string)=>row[keys.indexOf(key)];const text=(key:string)=>String(get(key)??'').trim();const d=dateSeconds(get('datadasolicitacao'));const s=seconds(get('horadasolicitacao'));
 if(d===null||s===null){rejected.push({sheet:sh.name,row:i+1,reason:'Data ou hora da solicitação inválida'});continue;}
 const stamp=(k:string)=>{const t=seconds(get(k));return t===null?null:d+t+(t<s?DAY:0)};
 const requested=d+s,released=stamp('horaliberacao'),departed=stamp('saidaparaentrega');const total=seconds(get('tempodoatendimento'),true);const de=seconds(get('entregadadieta'));
 const delivered=de===null?null:total!==null?Math.floor((requested+total)/DAY)*DAY+de:stamp('entregadadieta');
 const invalid=Boolean((released!==null&&departed!==null&&departed<released)||(departed!==null&&delivered!==null&&delivered<departed)||(released!==null&&delivered!==null&&delivered<released)||(total!==null&&delivered!==null&&Math.abs(delivered-requested-total)>2));
 for(const k of ['horaliberacao','saidaparaentrega','entregadadieta'])if(text(k)&&seconds(get(k))===null)warnings.add('Há horários inválidos na coluna '+k+'; foram tratados como ausentes.');
 const observation=text('observacao');records.push({id:sh.name+':'+(i+1),row:i+1,sheet:sh.name,hospital:text('unidade')||'Não informado',floor:text('andar')||'Não informado',request:text('solicitacao')||'Não informada',type:text('tipo')||'Não informado',observation,meal:mealFromObservation(observation),canceled:normalize(get('cancelado'))==='sim',day:dateKey(requested),requested,released,departed,delivered,systemMinutes:total===null?null:total/60,invalid,inferred:[released,departed,delivered].some(t=>t!==null&&dateKey(t)!==dateKey(requested))});
 }}
 if(!found)throw new Error('Não encontrei as colunas Data da Solicitação e Hora da Solicitação. Envie o registro original em .xlsx.');
 if(!records.length)throw new Error('Nenhuma solicitação com data e horário válidos foi encontrada.');
 records.sort((a,b)=>a.requested-b.requested||a.id.localeCompare(b.id));return {records,rejected,warnings:[...warnings],filename,sourceSheets:sheets};
}
export type Filters={month:string;day:string;hospital:string;floor:string;type:string};
const mean=(a:number[])=>a.length?a.reduce((s,v)=>s+v,0)/a.length:null;
export function analyze(records:DietRecord[],filters:Filters){
 const scope=records.filter(r=>(!filters.hospital||r.hospital===filters.hospital)&&(!filters.floor||r.floor===filters.floor)&&(!filters.type||r.type===filters.type));
 const inPeriod=(key:string)=>key.startsWith(filters.month)&&(!filters.day||key===filters.day);
 const cohort=scope.filter(r=>inPeriod(r.day)),active=cohort.filter(r=>!r.canceled);
 const lib=scope.filter(r=>!r.canceled&&r.released!==null&&inPeriod(dateKey(r.released)));
 const delivered=scope.filter(r=>!r.canceled&&r.delivered!==null&&inPeriod(dateKey(r.delivered)));
 const releaseTimes=active.filter(r=>!r.invalid&&r.released!==null).map(r=>(r.released!-r.requested)/60);
 const deliveryTimes=active.filter(r=>!r.invalid&&r.released!==null&&r.delivered!==null).map(r=>(r.delivered!-r.released!)/60);
 const totals=active.filter(r=>r.systemMinutes!==null).map(r=>r.systemMinutes!);
 const days:string[]=[];const [year,month]=filters.month.split('-').map(Number);const num=new Date(Date.UTC(year,month,0)).getUTCDate();for(let i=1;i<=num;i++){const key=`${filters.month}-${String(i).padStart(2,'0')}`;if(!filters.day||filters.day===key)days.push(key);}
 const hourly=Array.from({length:24},(_,hour)=>({name:`${String(hour).padStart(2,'0')}h`,hour,count:0}));const grid=days.map(day=>({day,counts:Array(24).fill(0) as number[]}));
 for(const r of lib){const h=new Date(r.released!*1000).getUTCHours();hourly[h].count++;const row=grid.find(x=>x.day===dateKey(r.released!));if(row)row.counts[h]++;}
 const daily=days.map(day=>{const a=active.filter(r=>r.day===day);const good=a.filter(r=>!r.invalid);return {day,name:day.slice(8),requests:cohort.filter(r=>r.day===day).length,canceled:cohort.filter(r=>r.day===day&&r.canceled).length,releases:lib.filter(r=>dateKey(r.released!)===day).length,deliveries:delivered.filter(r=>dateKey(r.delivered!)===day).length,releaseMinutes:mean(good.filter(r=>r.released!==null).map(r=>(r.released!-r.requested)/60)),deliveryMinutes:mean(good.filter(r=>r.released!==null&&r.delivered!==null).map(r=>(r.delivered!-r.released!)/60)),totalMinutes:mean(a.filter(r=>r.systemMinutes!==null).map(r=>r.systemMinutes!))}});
 const peaks=grid.flatMap(g=>{const n=Math.max(...g.counts);return n?g.counts.flatMap((v,h)=>v===n?[{day:g.day,hour:h,count:n,records:lib.filter(r=>dateKey(r.released!)===g.day&&new Date(r.released!*1000).getUTCHours()===h)}]:[]):[]});
 const requestHours=days.flatMap(day=>Array.from({length:24},(_,hour)=>{
  const bucket=cohort.filter(r=>r.day===day&&new Date(r.requested*1000).getUTCHours()===hour);
  const valid=bucket.filter(r=>!r.canceled&&!r.invalid&&r.released!==null);
  return {day,hour,name:`${String(hour).padStart(2,'0')}h`,requests:bucket.length,validReleases:valid.length,releaseMinutes:mean(valid.map(r=>(r.released!-r.requested)/60)),canceled:bucket.filter(r=>r.canceled).length,missingRelease:bucket.filter(r=>!r.canceled&&r.released===null).length,invalid:bucket.filter(r=>!r.canceled&&r.invalid).length};
 }));
 const countBy=(arr:DietRecord[],key:'type'|'floor'|'meal'|'request')=>{const m=new Map<string,number>();for(const r of arr)m.set(r[key],(m.get(r[key])||0)+1);return [...m].map(([name,count])=>({name,count})).sort((a,b)=>b.count-a.count||a.name.localeCompare(b.name));};
 return {cohort,active,lib,delivered,daily,hourly,grid,peaks,requestHours,types:countBy(lib,'type'),floors:countBy(lib,'floor'),meals:countBy(active,'meal'),requestsByType:countBy(cohort,'request'),releaseMean:mean(releaseTimes),deliveryMean:mean(deliveryTimes),totalMean:mean(totals),releaseN:releaseTimes.length,deliveryN:deliveryTimes.length,totalN:totals.length,missingRelease:active.filter(r=>r.released===null).length,missingDeparture:active.filter(r=>r.departed===null).length,missingDelivery:active.filter(r=>r.delivered===null).length,invalid:active.filter(r=>r.invalid).length,unknownMeals:active.filter(r=>r.meal==='Não informada').length};
}
