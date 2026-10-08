import {analyze,parseSheets,brDate,clock,type Filters,type ImportResult,type DietRecord} from './analysis';
export async function readExcel(file:File):Promise<ImportResult>{
 if(!file.name.toLowerCase().endsWith('.xlsx'))throw new Error('Selecione um arquivo .xlsx. Arquivos .xls e PDF não são compatíveis nesta versão.');
 if(file.size>25*1024*1024)throw new Error('O limite por arquivo é 25 MB. Exporte somente o período que deseja analisar.');
 const ExcelJS=(await import('exceljs')).default;const w=new ExcelJS.Workbook();await w.xlsx.load(await file.arrayBuffer());
 function value(v:unknown):unknown{if(v&&typeof v==='object'&&!(v instanceof Date)){const o=v as {result?:unknown;text?:string;richText?:{text:string}[]};if('result'in o)return value(o.result);if(o.richText)return o.richText.map(x=>x.text).join('');if(o.text)return o.text;}return v;}
 return parseSheets(w.worksheets.map(s=>{const rows:unknown[][]=[];for(let i=1;i<=s.rowCount;i++){const r=s.getRow(i);const a:unknown[]=[];for(let j=1;j<=s.columnCount;j++)a.push(value(r.getCell(j).value));rows.push(a);}return{name:s.name,rows}}),file.name);
}
export function recordRows(records:DietRecord[]){return records.map(r=>[r.sheet,r.row,r.hospital,r.floor,r.request,r.type,r.meal,r.canceled?'Cancelado':'Não cancelado',brDate(r.day),clock(r.requested),r.released===null?null:brDate(new Date(r.released*1000).toISOString().slice(0,10))+' '+clock(r.released),r.departed===null?null:brDate(new Date(r.departed*1000).toISOString().slice(0,10))+' '+clock(r.departed),r.delivered===null?null:brDate(new Date(r.delivered*1000).toISOString().slice(0,10))+' '+clock(r.delivered),!r.invalid&&r.released!==null?(r.released-r.requested)/60:null,!r.invalid&&r.released!==null&&r.delivered!==null?(r.delivered-r.released)/60:null,r.systemMinutes,r.invalid?'Sequência inconsistente':'',r.observation]);}
export const recordHeaders=['Aba de origem','Linha de origem','Hospital','Andar','Solicitação','Tipo de dieta / item','Refeição explícita','Status','Data solicitação','Hora solicitação','Liberação (data inferida)','Saída (data inferida)','Entrega (data inferida)','Solicitação → liberação (min)','Liberação → entrega (min)','Total sistema (min)','Conferência','Observação'];
export async function svgToPng(svg:SVGSVGElement){
 const box=svg.getBoundingClientRect();if(!box.width||!box.height)throw new Error('Não foi possível capturar um gráfico. Abra a visão geral e tente novamente.');
 const clone=svg.cloneNode(true) as SVGSVGElement;clone.setAttribute('xmlns','http://www.w3.org/2000/svg');clone.setAttribute('width',String(box.width));clone.setAttribute('height',String(box.height));
 const originals=[svg,...svg.querySelectorAll('*')],copies=[clone,...clone.querySelectorAll('*')];originals.forEach((o,i)=>{const style=getComputedStyle(o);for(const prop of ['fill','stroke','stroke-width','font-family','font-size','font-weight','opacity'])copies[i].setAttribute(prop,style.getPropertyValue(prop));});
 const image=new Image();image.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(new XMLSerializer().serializeToString(clone));await image.decode();
 const canvas=document.createElement('canvas');canvas.width=Math.ceil(box.width*2);canvas.height=Math.ceil(box.height*2);const ctx=canvas.getContext('2d');if(!ctx)throw new Error('Seu navegador não conseguiu exportar os gráficos.');ctx.fillStyle='#ffffff';ctx.fillRect(0,0,canvas.width,canvas.height);ctx.drawImage(image,0,0,canvas.width,canvas.height);return canvas.toDataURL('image/png');
}
export async function exportExcel(result:ImportResult,filters:Filters,chartDay?:string){
 const {buildModelExcel}=await import('./model-export');
 const bytes=await buildModelExcel(result,filters,chartDay);
 const url=URL.createObjectURL(new Blob([bytes as BlobPart],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
 const link=document.createElement('a');link.href=url;
 const monthNames=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
 const [year,month]=filters.month.split('-');link.download=`Analise_Dietas_${monthNames[Number(month)-1]}_${year}${filters.day?'_'+filters.day.slice(8):''}.xlsx`;
 document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);
}
