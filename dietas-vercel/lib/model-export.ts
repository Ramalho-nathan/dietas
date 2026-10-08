import JSZip from 'jszip';
import templateData from './excel-model.json';
import {nativeHourChart} from './hour-chart';
import {analyze,normalize,dateKey,type ImportResult,type Filters,type DietRecord} from './analysis';

// Only layout, styles, headers and native chart definitions are shipped.
// Every exported data cell comes from the current import and current filters.
type Value=string|number|null|{formula:string;value:string|number|null};
type Grid=Map<string,Value>;
const template=templateData as unknown as {files:Record<string,string>;sheets:{shell:string;rows:Record<string,{height:string;styles:number[]}>;static:Record<string,string>}[]};
const esc=(x:unknown)=>String(x??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
const column=(n:number):string=>n<26?String.fromCharCode(65+n):column(Math.floor(n/26)-1)+column(n%26);
const serial=(s:number|null)=>s===null?null:s/86400+25569;
const daySerial=(s:number|null)=>s===null?null:Math.floor(s/86400)+25569;
const mean=(v:number[])=>v.length?v.reduce((s,n)=>s+n,0)/v.length:null;
const formula=(formula:string,value:string|number|null):Value=>({formula,value});
const names=['Painel','Horários','Picos por dia','Base organizada','Fonte original'];
const standardNamespace=(xml:string)=>xml.replace(/xmlns:x="http:\/\/schemas.openxmlformats.org\/spreadsheetml\/2006\/main"/g,'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"').replace(/<(\/?)x:/g,'<$1');
function relativeRelationships(path:string,xml:string){
 if(!path.endsWith('.rels'))return xml;
 const from=path.includes('/_rels/')?path.split('/_rels/')[0].split('/'):[];
 return xml.replace(/Target="\/([^"]+)"/g,(_,target:string)=>{
  const to=target.split('/');let common=0;while(common<from.length&&from[common]===to[common])common++;
  return `Target="${'../'.repeat(from.length-common)+to.slice(common).join('/')}"`;
 });
}

function sheetXML(index:number,grid:Grid,last:number,width:number,styles:(row:number)=>number[],height:(row:number)=>string){
 const layout=template.sheets[index];
 const xmlCell=(ref:string,value:Value|undefined,style:number)=>{
  const a=`r="${ref}" s="${style}"`;
  if(value===undefined||value===null)return `<c ${a}/>`;
  if(typeof value==='object')return `<c ${a}${typeof value.value==='string'?' t="str"':''}><f>${esc(value.formula)}</f>${value.value===null?'':`<v>${esc(value.value)}</v>`}</c>`;
  return typeof value==='number'?`<c ${a}><v>${value}</v></c>`:`<c ${a} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
 };
 const rows=[];
 for(let r=1;r<=last;r++){
  const stylesRow=styles(r);
  rows.push(`<row r="${r}" ht="${height(r)}" customHeight="1">${Array.from({length:width},(_,c)=>xmlCell(column(c)+r,grid.get(column(c)+r),stylesRow[c]??2)).join('')}</row>`);
 }
 return layout.shell.replace('__DATA__',`<sheetData xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${rows.join('')}</sheetData>`);
}

export async function buildModelExcel(result:ImportResult,filters:Filters,chartDay?:string):Promise<Uint8Array>{
 const a=analyze(result.records,filters);
 const hourAnalysis=analyze(result.records,{...filters,day:''});
 const scope=result.records.filter(r=>(!filters.hospital||r.hospital===filters.hospital)&&(!filters.floor||r.floor===filters.floor)&&(!filters.type||r.type===filters.type));
 const grids:Grid[]=template.sheets.map(s=>new Map(Object.entries(s.static) as [string,Value][]));
 const [panel,hours,peaks,base,source]=grids;
 const row=(g:Grid,r:number,v:Value[])=>v.forEach((x,c)=>g.set(column(c)+r,x));
 const modelRow=(i:number,r:number)=>{const rows=template.sheets[i].rows as Record<string,{height:string;styles:number[]}>;return rows[String(r)]??{height:'21',styles:Array(26).fill(2) as number[]}};
 const [year,month]=filters.month.split('-').map(Number);
 const start=filters.day?Math.floor(Date.parse(filters.day+'T00:00:00Z')/86400000)+25569:Date.UTC(year,month-1,1)/86400000+25569;
 const end=filters.day?start+1:Date.UTC(year,month,1)/86400000+25569;
 const monthTitle=new Intl.DateTimeFormat('pt-BR',{month:'long',year:'numeric',timeZone:'UTC'}).format(new Date(Date.UTC(year,month-1,1)));
 panel.set('A2',`Solicitações de dietas | ${monthTitle[0].toUpperCase()+monthTitle.slice(1)}${filters.day?' | '+filters.day.split('-').reverse().join('/'):''}`);
 panel.set('A5',`Solicitações em ${monthTitle.split(' ')[0]}`);
 panel.set('A8',`Liberações em ${monthTitle.split(' ')[0]}`);
 panel.set('A9',`Entregas em ${monthTitle.split(' ')[0]}`);
 panel.set('D9',`Tempos: não cancelados pela data da solicitação. ${a.invalid} divergências sinalizadas no período.`);
 const others=scope.filter(r=>!r.day.startsWith(filters.month)).length;
 panel.set('D10',`${others} registros de outros meses preservados na base, fora do painel. Filtros: ${[filters.hospital||'todos os hospitais',filters.floor||'todos os andares',filters.type||'todos os tipos'].join('; ')}.`);
 const lastBase=Math.max(6,scope.length+5);
 const range=(c:string)=>`'Base organizada'!$${c}$6:$${c}$${lastBase}`;
 const conditions=(dateCol:string,day?:string)=>day?`${range(dateCol)},${day}`:`${range(dateCol)},">=${start}",${range(dateCol)},"<${end}"`;
 const count=(dateCol:string,day?:string,extra='')=>`COUNTIFS(${conditions(dateCol,day)}${extra})`;
 const active=`,${range('H')},"Ativo"`;
 const good=(delivery=false)=>`${active},${range('P')},"<>Divergência > 2 s",${range('K')},">0"${delivery?`,${range('M')},">0"`:''}`;
 [a.cohort.length,a.cohort.length-a.active.length,a.active.length,a.lib.length,a.delivered.length].forEach((value,i)=>{
  const f=i===0?count('I'):i===1?count('I',undefined,`,${range('H')},"Cancelado"`):count(i===2?'I':i===3?'S':'U',undefined,active);
  panel.set('B'+(i+5),formula(f,value));
 });
 for(let i=0;i<a.daily.length;i++){
  const d=a.daily[i],r=i+15,day=`A${r}`,records=a.active.filter(x=>x.day===d.day&&!x.invalid);
  const releaseN=records.filter(x=>x.released!==null).length,deliveryN=records.filter(x=>x.released!==null&&x.delivered!==null).length;
  const avg=(c:string,extras:string,value:number|null)=>formula(`IFERROR(SUMIFS(${range(c)},${conditions('I',day)}${extras})/${count('I',day,extras)},"")`,value);
  row(panel,r,[daySerial(Date.parse(d.day+'T00:00:00Z')/1000),formula(count('I',day),d.requests),formula(count('I',day,`,${range('H')},"Cancelado"`),d.canceled),formula(count('S',day,active),d.releases),formula(count('U',day,active),d.deliveries),avg('V',good(),d.releaseMinutes),avg('W',good(true),d.deliveryMinutes),avg('N',`${active},${range('N')},"<>"`,d.totalMinutes),formula(count('I',day,good()),releaseN),formula(count('I',day,good(true)),deliveryN)]);
  row(panel,i+51,[d.day.slice(8)+'/'+d.day.slice(5,7),formula(`D${r}`,d.releases),formula(`B${r}`,d.requests),formula(`F${r}`,d.releaseMinutes),formula(`G${r}`,d.deliveryMinutes)]);
 }
 for(let i=0;i<24;i++)row(panel,86+i,[`${String(i).padStart(2,'0')}–${String(i+1).padStart(2,'0')}h`,formula(`${count('S',undefined,active+`,${range('T')},${i}`)}`,a.hourly[i].count)]);
 for(let i=0;i<10;i++){
  const type=a.types[i],floor=a.floors[i];
  row(panel,116+i,[type?.name??null,type?formula(count('S',undefined,active+`,${range('E')},A${116+i}`),type.count):null,null,floor?.name??null,floor?formula(count('S',undefined,active+`,${range('C')},D${116+i}`),floor.count):null]);
 }
 a.meals.forEach((v,i)=>row(panel,131+i,[v.name,formula(count('I',undefined,active+`,${range('G')},A${131+i}`),v.count)]));
 a.grid.forEach((g,i)=>row(hours,i+7,[daySerial(Date.parse(g.day+'T00:00:00Z')/1000),...g.counts]));
 hours.set('A43','Solicitações por dia e hora e média até a liberação');
 hours.set('A44','Faixas pela hora da solicitação. Média: não canceladas com horários válidos; inclui liberações no dia seguinte.');
 row(hours,45,['Data','Hora','Qtd.','N válido','Média (min)','Cancel.','Sem lib.','Incons.','Faixa']);
 hourAnalysis.requestHours.forEach((h,i)=>{
  const lower=Date.parse(h.day+'T00:00:00Z')/86400000+25569+h.hour/24,upper=lower+1/24;
  const criteria=`${range('J')},">=${lower}",${range('J')},"<${upper}"`;
  const n=`COUNTIFS(${criteria}${good()})`;
  row(hours,46+i,[daySerial(Date.parse(h.day+'T00:00:00Z')/1000),h.hour/24,formula(`COUNTIFS(${criteria})`,h.requests),formula(n,h.validReleases),formula(`IFERROR(SUMIFS(${range('V')},${criteria}${good()})/${n},"")`,h.releaseMinutes),h.canceled,h.missingRelease,h.invalid,h.name]);
 });
 // One editable date selector drives a fixed 24-hour chart range.
 const selectedDay=filters.day||(hourAnalysis.daily.some(d=>d.day===chartDay)?chartDay!:hourAnalysis.daily.find(d=>d.requests>0)?.day||hourAnalysis.daily[0].day);
 const selectedHours=hourAnalysis.requestHours.filter(h=>h.day===selectedDay);
 panel.set('L139','FILTRO DO GRÁFICO • ESCOLHA O DIA');
 panel.set('M141',daySerial(Date.parse(selectedDay+'T00:00:00Z')/1000));
 panel.set('L140','Clique no campo azul abaixo e escolha uma data.');
 panel.set('L144','Barras: solicitações · Linha: média em minutos · Faixas de 00h a 23h.');
 panel.set('M143',formula('"Solicitações e média até liberar | "&TEXT(M141,"dd/mm/yyyy")','Solicitações e média até liberar | '+selectedDay.split('-').reverse().join('/')));
 panel.set('Q141','◀ Clique na data para abrir a lista de dias');
 hours.set('AA5','Dias disponíveis para o gráfico');
 hourAnalysis.daily.forEach((d,i)=>hours.set('AA'+(7+i),daySerial(Date.parse(d.day+'T00:00:00Z')/1000)));
 hours.set('AA43','Dados do gráfico: dia escolhido no Painel');
 ['Hora','Solicitações','Média (min)','Faixa'].forEach((v,i)=>hours.set(column(26+i)+'45',v));
 const lastHour=45+hourAnalysis.requestHours.length;
 const hr=(c:string)=>`$${c}$46:$${c}$${lastHour}`;
 row(panel,146,['Hora','Solicitações','Média (min)']);
 for(let i=0;i<24;i++){
  const r=46+i,h=selectedHours[i];
  row(panel,147+i,[h.name,formula(`'Horários'!AB${r}`,h.requests),formula(`'Horários'!AC${r}`,h.releaseMinutes)]);hours.set('AA'+r,i/24);hours.set('AD'+r,h.name);
  const criteria=`${hr('A')},'Painel'!$M$136,${hr('B')},AA${r}`;
  hours.set('AB'+r,formula(`SUMIFS(${hr('C')},${criteria})`,h.requests));
  hours.set('AC'+r,formula(`IF(SUMIFS(${hr('D')},${criteria})=0,"",SUMIFS(${hr('E')},${criteria}))`,h.releaseMinutes));
 }
 a.peaks.forEach((p,i)=>row(peaks,i+7,[daySerial(Date.parse(p.day+'T00:00:00Z')/1000),p.hour/24,p.count]));
 const compositionHeader=Math.max(48,a.peaks.length+10);
 for(let c=0;c<8;c++)peaks.delete(column(c)+'48');
 const headers=Array.from({length:8},(_,c)=>(template.sheets[2].static as Record<string,string>)[column(c)+'48']);
 row(peaks,compositionHeader,headers);
 let compositionRow=compositionHeader+1;
 for(const peak of a.peaks){
  const groups=new Map<string,number>();
  for(const r of peak.records){const key=JSON.stringify([r.hospital,r.floor,r.request,r.type,r.meal]);groups.set(key,(groups.get(key)??0)+1)}
  for(const [key,n]of [...groups].sort(([x],[y])=>x<y?-1:x>y?1:0))row(peaks,compositionRow++,[daySerial(Date.parse(peak.day+'T00:00:00Z')/1000),peak.hour/24,...JSON.parse(key),n]);
 }
 base.set('A1',`Fonte: ${result.filename}. Linha identifica o registro original; dados organizados conforme os filtros de hospital, andar e tipo.`);
 const sourceMap=new Map((result.sourceSheets??[]).map(s=>[s.name,s]));
 const rawValue=(r:DietRecord,key:string)=>{const sheet=sourceMap.get(r.sheet);if(!sheet)return '';const hi=sheet.rows.slice(0,15).findIndex(x=>x.some(v=>normalize(v)==='datadasolicitacao'));return sheet.rows[r.row-1]?.[sheet.rows[hi].map(normalize).indexOf(key)]??''};
 for(let i=0;i<scope.length;i++){
  const x=scope[i],r=i+6;
  const elapsed=(from:number|null,to:number|null)=>!x.invalid&&from!==null&&to!==null?(to-from)/60:null;
  const metric=(col:string,from:string,to:string,value:number|null)=>formula(`IF(OR(${from}${r}="",${to}${r}="",P${r}="Divergência > 2 s"),"",(${to}${r}-${from}${r})*1440)`,value);
  row(base,r,[x.row,x.hospital,x.floor,x.request,x.type,x.observation,x.meal,x.canceled?'Cancelado':'Ativo',daySerial(x.requested),serial(x.requested),serial(x.released),serial(x.departed),serial(x.delivered),x.systemMinutes,x.inferred?'Dia seguinte / inferido':'Mesmo dia',x.invalid?'Divergência > 2 s':'Conferido',String(rawValue(x,'cancelado')),String(rawValue(x,'motivodocancelamento')),formula(`IF(K${r}="","",INT(K${r}))`,daySerial(x.released)),formula(`IF(K${r}="","",HOUR(K${r}))`,x.released===null?null:new Date(x.released*1000).getUTCHours()),formula(`IF(M${r}="","",INT(M${r}))`,daySerial(x.delivered)),metric('V','J','K',elapsed(x.requested,x.released)),metric('W','K','M',elapsed(x.released,x.delivered)),metric('X','K','L',elapsed(x.released,x.departed)),metric('Y','L','M',elapsed(x.departed,x.delivered))]);
 }
 // The full source is retained locally and copied only into the user's download.
 const sources=result.sourceSheets??[];
 let sourceRow=3,sourceWidth=26;
 const sourceValue=(v:unknown):Value=>v instanceof Date?(v.getUTCFullYear()<1901?v.toISOString().slice(11,19):v.toISOString().slice(0,10).split('-').reverse().join('/')):typeof v==='number'?v:v===undefined||v===null?null:String(v);
 for(let i=0;i<sources.length;i++){
  const s=sources[i];const hi=s.rows.slice(0,15).findIndex(x=>x.some(v=>normalize(v)==='datadasolicitacao')&&x.some(v=>normalize(v)==='horadasolicitacao'));
  if(hi<0)continue;
  if(i===0){row(source,1,[String(s.rows[0]?.[0]??'Registros')]);if(s.rows[hi].length!==26)row(source,2,s.rows[hi].map(v=>v===null||v===undefined||v===''?'Coluna sem título':sourceValue(v)))}
  else {row(source,sourceRow++,[`Aba de origem: ${s.name}`]);row(source,sourceRow++,s.rows[hi].map(sourceValue))}
  sourceWidth=Math.max(sourceWidth,s.rows[hi].length);
  for(const values of s.rows.slice(hi+1))row(source,sourceRow++,values.map(sourceValue));
 }
 if(!sources.length)source.set('A3','Fonte original indisponível nesta sessão. Importe novamente o arquivo para preservá-la.');
 const zip=new JSZip();
 for(const [path,content]of Object.entries(template.files))zip.file(path,relativeRelationships(path,standardNamespace(content)));
 let styleXML=await zip.file('xl/styles.xml')!.async('string');
 styleXML=styleXML.replace('<cellXfs count="14">','<cellXfs count="15">').replace('</cellXfs>','<xf numFmtId="200" fontId="2" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs>');
 zip.file('xl/styles.xml',styleXML);
 const stylePanel=(r:number)=>{const st=[...modelRow(0,r>=15&&r<15+a.daily.length?15:r>=51&&r<51+a.daily.length?51:r>=86&&r<110?86:r>=116&&r<126?116:r>=131&&r<131+a.meals.length?131:r).styles];if(r===141){for(let c=12;c<=14;c++)st[c]=14;}if(r===146){for(let c=0;c<3;c++)st[c]=6;}if(r>=147&&r<=170){st[0]=2;st[1]=2;st[2]=9;}return st;};
 const xmls=[
  sheetXML(0,panel,174,26,stylePanel,r=>r===141?'32':r===146?'36':modelRow(0,r).height),
  sheetXML(1,hours,Math.max(69,hourAnalysis.requestHours.length+45),30,r=>{const st=r===45?Array(30).fill(6):r>=46?[7,13,2,2,9,2,2,2,2]:[...modelRow(1,r>=7&&r<7+a.grid.length?7:r).styles];st[26]=r>=46?13:7;st[28]=9;return st;},r=>r===45?'36':modelRow(1,r).height),
  sheetXML(2,peaks,Math.max(compositionRow+1,compositionHeader+2),8,r=>modelRow(2,r===compositionHeader?48:r>compositionHeader?49:r>=7&&r<7+a.peaks.length?7:r).styles,r=>r===compositionHeader?'36':modelRow(2,r===48?1:r).height),
  sheetXML(3,base,lastBase,25,r=>modelRow(3,r>=6?6:r).styles,r=>modelRow(3,r>=6?6:r).height),
  sheetXML(4,source,Math.max(3,sourceRow-1),sourceWidth,r=>modelRow(4,r>=3?3:r).styles,r=>modelRow(4,r>=3?3:r).height),
 ];
 xmls[0]=xmls[0].replace(/<(?:x:)?pageMargins/,'<mergeCells count="1"><mergeCell ref="M141:O141"/></mergeCells><pageMargins');
 xmls[0]=xmls[0].replace(/<(?:x:)?pageMargins/,`<dataValidations count="1"><dataValidation type="list" allowBlank="0" showErrorMessage="1" errorTitle="Dia inválido" error="Escolha um dia da lista." sqref="M141"><formula1>DiasGrafico</formula1></dataValidation></dataValidations><pageMargins`);
 xmls.forEach((x,i)=>zip.file(`xl/worksheets/sheet${i+1}.xml`,standardNamespace(x)));
 const table=(path:string,ref:string)=>zip.file(path,standardNamespace((template.files as Record<string,string>)[path].replace(/ref="[^"]*"/,`ref="${ref}"`)));
 table('xl/tables/table1.xml',`A${compositionHeader}:H${Math.max(compositionHeader+1,compositionRow-1)}`);
 table('xl/tables/table2.xml',`A5:Y${lastBase}`);
 // Populate chart caches as well as references so Excel opens with visible charts.
 for(let i=1;i<=6;i++){
  const path=`xl/drawings/charts/chart${i}.xml`;let xml=(template.files as Record<string,string>)[path];
  const dailyEnd=50+a.daily.length,mealEnd=130+Math.max(1,a.meals.length);
  if(i<=2)xml=xml.replace(/\$80/g,`$${dailyEnd}`);
  if(i===6)xml=xml.replace(/\$137/g,`$${mealEnd}`);
  xml=xml.replace(/<c:(strRef|numRef)>([\s\S]*?)<\/c:\1>/g,(_,kind:string,body:string)=>{
   const ref=body.match(/<c:f>(.*?)<\/c:f>/)?.[1],m=ref?.match(/\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+)/);
   if(!m)return `<c:${kind}>${body}</c:${kind}>`;
   const points=[];const length=Number(m[3])-Number(m[2])+1;
   for(let j=0;j<length;j++){let v=panel.get(m[1]+(Number(m[2])+j));if(typeof v==='object'&&v!==null)v=v.value;if(v!==undefined&&v!==null)points.push(`<c:pt idx="${j}"><c:v>${esc(v)}</c:v></c:pt>`)}
   const cache=kind==='strRef'?'strCache':'numCache';const inner=`${kind==='numRef'?'<c:formatCode>General</c:formatCode>':''}<c:ptCount val="${length}"/>${points.join('')}`;
   return `<c:${kind}>${body.replace(new RegExp(`<c:${cache}>[\\s\\S]*?<\\/c:${cache}>`),`<c:${cache}>${inner}</c:${cache}>`)}</c:${kind}>`;
  });
  zip.file(path,xml);
 }
 zip.file('xl/drawings/charts/chart7.xml',nativeHourChart(selectedDay,46,selectedHours,true));
 const drawingPath='xl/drawings/drawing1.xml';let drawing=await zip.file(drawingPath)!.async('string');
 const anchor=drawing.match(/<(?:xdr:)?twoCellAnchor[^>]*>[\s\S]*?<\/(?:xdr:)?twoCellAnchor>/)?.[0];
 if(!anchor)throw new Error('Não foi possível posicionar o novo gráfico no modelo.');
 // Separate from/to anchors: use the same width and a new block below chart six.
 let position=0;let extra=anchor.replace(/(<(?:xdr:)?row>)\d+(<\/(?:xdr:)?row>)/g,(_,open:string,close:string)=>`${open}${position++===0?145:172}${close}`);
 extra=extra.replace(/r:id="[^"]+"/,'r:id="hourChart7"').replace(/(<(?:xdr:)?cNvPr[^>]*\bid=")\d+"/,(_,prefix:string)=>prefix+'7"').replace(/\bname="[^"]*"/,'name="Solicitações por dia e hora"');
 drawing=drawing.replace(/<\/(\w+:)?wsDr>/,`${extra}</$1wsDr>`);zip.file(drawingPath,drawing);
 const relPath='xl/drawings/_rels/drawing1.xml.rels';const rel=await zip.file(relPath)!.async('string');
 zip.file(relPath,rel.replace('</Relationships>','<Relationship Id="hourChart7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart7.xml"/></Relationships>'));
 const content=await zip.file('[Content_Types].xml')!.async('string');zip.file('[Content_Types].xml',content.replace('</Types>','<Override PartName="/xl/drawings/charts/chart7.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'));
 let workbook=(template.files as Record<string,string>)['xl/workbook.xml'];
 const dayName=`<definedName name="DiasGrafico">'Horários'!$AA$7:$AA$${6+hourAnalysis.daily.length}</definedName>`;
 if(/<\/(?:x:)?definedNames>/.test(workbook))workbook=workbook.replace(/<\/(?:x:)?definedNames>/,dayName+'$&');
 else workbook=workbook.replace(/<\/(?:x:)?sheets>/,'$&<definedNames>'+dayName+'</definedNames>');
 workbook=workbook.replace(/<(?:x:)?calcPr\b[^>]*\/>/,'');
 workbook=workbook.replace(/<\/(x:)?workbook>/,'<calcPr xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></$1workbook>');
 zip.file('xl/workbook.xml',standardNamespace(workbook));
 return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
