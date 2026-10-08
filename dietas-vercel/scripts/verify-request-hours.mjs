import assert from 'node:assert/strict';
import {parseSheets,analyze} from '../lib/analysis.ts';
const headers=['Data da Solicitação','Hora da Solicitação','Unidade','Andar','Solicitação','Tipo','Hora Liberação','Saída para Entrega','Entrega da Dieta','Cancelado'];
const rows=Array.from({length:10},(_,i)=>['02/10/2026',`12:${String(i).padStart(2,'0')}:00`,'A','1','Admissão','Geral',`12:${String(i+10).padStart(2,'0')}:00`]);
const filters={month:'2026-10',day:'',hospital:'',floor:'',type:''};
const input=parseSheets([{name:'Teste',rows:[headers,...rows]}],'teste.xlsx');
let result=analyze(input.records,filters),bucket=result.requestHours.find(x=>x.day==='2026-10-02'&&x.hour===12);
assert.equal(bucket.requests,10);assert.equal(bucket.validReleases,10);assert.equal(bucket.releaseMinutes,10);
const extra=parseSheets([{name:'Teste',rows:[headers,...rows,
 ['02/10/2026','12:20:00','A','1','Admissão','Geral','12:40:00',null,null,'Sim'],
 ['02/10/2026','12:30:00','A','1','Admissão','Geral'],
 ['02/10/2026','12:40:00','A','1','Admissão','Geral','12:50:00','12:49:00'],
 ['02/10/2026','13:00:00','A','1','Admissão','Geral','13:20:00'],
 ['02/10/2026','23:55:00','A','1','Admissão','Geral','00:05:00'],
 ['03/10/2026','12:00:00','B','2','Admissão','Leve','12:30:00']]}],'teste.xlsx');
result=analyze(extra.records,filters);bucket=result.requestHours.find(x=>x.day==='2026-10-02'&&x.hour===12);
assert.equal(bucket.requests,13);assert.equal(bucket.validReleases,10);assert.equal(bucket.releaseMinutes,10);assert.equal(bucket.canceled,1);assert.equal(bucket.missingRelease,1);assert.equal(bucket.invalid,1);
assert.equal(result.requestHours.find(x=>x.day==='2026-10-02'&&x.hour===23).releaseMinutes,10);
assert.equal(result.requestHours.find(x=>x.day==='2026-10-02'&&x.hour===14).releaseMinutes,null);
assert.equal(result.requestHours.reduce((sum,x)=>sum+x.requests,0),result.cohort.length);
const weighted=result.requestHours.reduce((sum,x)=>sum+(x.releaseMinutes??0)*x.validReleases,0)/result.requestHours.reduce((sum,x)=>sum+x.validReleases,0);
assert.equal(weighted,result.releaseMean);
const scoped=analyze(extra.records,{...filters,day:'2026-10-03',hospital:'B',floor:'2',type:'Leve'});assert.equal(scoped.requestHours.length,24);assert.equal(scoped.requestHours[12].requests,1);assert.equal(scoped.requestHours[12].releaseMinutes,30);
console.log('Dia/hora conferidos: 10 solicitações e média de 10 min; exclusões, limites de faixa, virada de dia e filtros aprovados.');
