"use strict";
const KEY="diagnostic-journal:v3", OLD="delivery-to-service-centers-v1", MAIN="Журнал-диагностики.json";
const $=id=>document.getElementById(id), clean=(v,n=1000)=>String(v??"").trim().slice(0,n);
const PRI={normal:"Обычный",urgent:"Срочно",critical:"Критично"}, PRIORITY={normal:0,urgent:1,critical:2};
const kinds=["centers","requests","trips"],newStore=()=>({version:3,centers:[],requests:[],trips:[],removed:{centers:[],requests:[],trips:[]}});
const uid=()=>globalThis.crypto?.randomUUID?.()||Date.now().toString(36)+"-"+Math.random().toString(36).slice(2);
const iso=()=>new Date().toISOString(),today=()=>{const t=new Date();return [t.getFullYear(),String(t.getMonth()+1).padStart(2,"0"),String(t.getDate()).padStart(2,"0")].join("-")};
const fmt=d=>{const v=String(d??"").slice(0,10);return isDay(v)?v.split("-").reverse().join("."):"—"}, time=s=>s?new Date(s).toLocaleString("ru-RU",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"}):"—";
const prettyDate=fmt;
function validYandexLink(value) {
  try {
    const url = new URL(value);
    return url.protocol === "https:" && /^(www\.)?yandex\.(ru|com|by)$/.test(url.hostname) && /^\/maps(?:\/|$)/.test(url.pathname);
  } catch { return false; }
}
function mapUrl(center) {
  if (center.map && validYandexLink(center.map)) return center.map;
  const url = new URL("https://yandex.ru/maps/");
  url.searchParams.set("text", center.address);
  return url.href;
}
const withoutMeta = ({_v, ...fields}) => Object.fromEntries(Object.entries(fields).sort(([a],[b]) => a<b?-1:a>b?1:0));
function baseClock(record) {
  const text = JSON.stringify(Object.entries(withoutMeta(record)).sort(([a],[b]) => a.localeCompare(b)));
  let h = 2166136261, g = 5381;
  for (let i=0;i<text.length;i++) { h = Math.imul(h ^ text.charCodeAt(i), 16777619); g = Math.imul(g, 33) ^ text.charCodeAt(i); }
  return {["b"+(h>>>0).toString(16)+"_"+(g>>>0).toString(16)]:1};
}
function joinClock(a,b) {
  const c={...a};
  for (const [k,n] of Object.entries(b)) c[k]=Math.max(c[k]||0,n);
  return c;
}
function advance(clock) {
  return {...clock,[actor]:(clock[actor]||0)+1};
}
function clockGE(a,b) { return Object.entries(b).every(([k,n]) => (a[k]||0)>=n); }
function readClock(clock) {
  if (!clock || typeof clock!=="object" || Array.isArray(clock) || !Object.keys(clock).length ||
      Object.entries(clock).some(([k,n])=>!/^([ab][a-zA-Z0-9_-]{1,105})$/.test(k) || !Number.isSafeInteger(n) || n<=0 || n>=Number.MAX_SAFE_INTEGER))
    throw Error("Некорректная история изменений в файле.");
  return {...clock};
}
function trackChanges(prev,next) {
  for (const k of kinds) {
    const old=new Map(prev[k].map(r=>[r.id,r]));
    for (const row of next[k]) {
      const before=old.get(row.id);
      if (!before || JSON.stringify(withoutMeta(row))!==JSON.stringify(withoutMeta(before))) row._v=advance(before?._v||{});
      old.delete(row.id);
    }
    for (const r of old.values()) next.removed[k].push({id:r.id,label:r.name||(k==="trips"?"Рейс № "+r.no+" · "+fmt(r.date):"№ "+r.ticket+" · "+r.item),_v:advance(r._v)});
  }
}
function indexed(s,k) {
  return new Map([...s[k].map(row=>[row.id,{row,dead:false}]),...s.removed[k].map(row=>[row.id,{row,dead:true}])]);
}
function sameValue(a,b) {
  return a.dead===b.dead && (a.dead || JSON.stringify(withoutMeta(a.row))===JSON.stringify(withoutMeta(b.row)));
}
function digest(s) {
  return JSON.stringify(kinds.map(k=>Array.from(indexed(s,k)).sort(([a],[b])=>a.localeCompare(b)).map(([id,e])=>[id,e.dead,withoutMeta(e.row),Object.entries(e.row._v).sort()])));
}
function planMerge(local,incoming) {
  const plan={records:{},conflicts:[],adds:0,edits:0,deletes:0};
  for (const k of kinds) {
    const left=indexed(local,k),right=indexed(incoming,k), dest=new Map(left);
    for (const [id,b] of right) {
      const a=left.get(id);
      if (!a) { dest.set(id,b); if(!b.dead) plan.adds++; continue; }
      const ge=clockGE(a.row._v,b.row._v),le=clockGE(b.row._v,a.row._v);
      if (sameValue(a,b)) {dest.set(id,{...a,row:{...a.row,_v:joinClock(a.row._v,b.row._v)}}); continue;}
      if (ge&&!le) continue;
      if (le&&!ge) {dest.set(id,b); if(b.dead) plan.deletes++; else plan.edits++; continue;}
      plan.conflicts.push({k,id,a,b});
    }
    plan.records[k]=dest;
  }
  return plan;
}
function downloadBlob(blob, filename) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 60000);
}
function xmlText(value) {
  return String(value ?? "").replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[char]);
}
function excelColumn(index) { let result = ""; for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result; return result; }
function workbookFiles(trip) {
  return {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Лист водителя" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="0"/></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="16"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF18375B"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    "xl/worksheets/sheet1.xml": makeWorksheet(trip)
  };
}
const CRC_TABLE = Array.from({ length: 256 }, (_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});
function crc32(bytes) { let value = 0xffffffff; for (const byte of bytes) value = CRC_TABLE[(value ^ byte) & 255] ^ (value >>> 8); return (value ^ 0xffffffff) >>> 0; }
function zipFiles(files) {
  const encoder = new TextEncoder();
  const local = [], central = [];
  let offset = 0;
  function u16(view, index, number) { view.setUint16(index, number, true); }
  function u32(view, index, number) { view.setUint32(index, number >>> 0, true); }
  for (const [path, content] of Object.entries(files)) {
    const name = encoder.encode(path);
    const body = encoder.encode(content);
    const checksum = crc32(body);
    const file = new Uint8Array(30 + name.length + body.length);
    const view = new DataView(file.buffer);
    u32(view, 0, 0x04034b50); u16(view, 4, 20); u16(view, 6, 0x0800); u16(view, 8, 0);
    u32(view, 14, checksum); u32(view, 18, body.length); u32(view, 22, body.length); u16(view, 26, name.length);
    file.set(name, 30); file.set(body, 30 + name.length); local.push(file);
    const directory = new Uint8Array(46 + name.length);
    const dir = new DataView(directory.buffer);
    u32(dir, 0, 0x02014b50); u16(dir, 4, 20); u16(dir, 6, 20); u16(dir, 8, 0x0800);
    u32(dir, 16, checksum); u32(dir, 20, body.length); u32(dir, 24, body.length); u16(dir, 28, name.length); u32(dir, 42, offset);
    directory.set(name, 46); central.push(directory); offset += file.length;
  }
  const centralLength = central.reduce((n, part) => n + part.length, 0);
  const end = new Uint8Array(22);
  const e = new DataView(end.buffer);
  u32(e, 0, 0x06054b50); u16(e, 8, central.length); u16(e, 10, central.length); u32(e, 12, centralLength); u32(e, 16, offset);
  return new Blob([...local, ...central, end], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}


const templateText=["Мне разъяснено, что в соответствии с требованиями «Закона о защите прав потребителей» мои требования, изложенные в заявлении, подлежат удовлетворению, в случае, если заявленные мною недостатки обусловлены причинами, возникшими до момента его приобретения (заводской брак и т.п.), а в отношении технически сложных и дорогостоящих изделий – в случае, если эти недостатки будут признаны существенными. Также мне разъяснено, что если в результате экспертизы будет установлено, что недостатки в изделии возникли вследствие обстоятельств, за которые не отвечает Продавец, я буду обязан возместить Продавцу расходы на проведение экспертизы, а также связанные с ее проведением расходы на хранение и транспортировку изделия (ст.18) Закона РФ «О защите прав потребителя».", "Стоимость транспортировки и хранения, при не гарантийном случае составляет 1500 рублей.", "Я проинформирован о том, что в случае согласия на не гарантийный платный ремонт, услуга оплачивается мною самостоятельно в сервисном центре."];
const validId=s=>typeof s==="string" && !!s.trim() && s===s.trim() && s.length<=120;
const isDay=s=>typeof s==="string" && /^\d{4}-\d\d-\d\d$/.test(s) && !Number.isNaN(Date.parse(s+"T12:00Z")) && new Date(s+"T12:00Z").toISOString().slice(0,10)===s;
function need(ok,msg="Файл журнала повреждён или имеет другой формат."){if(!ok)throw Error(msg)}
function arr(a){need(Array.isArray(a));return a}
function strs(obj,keys){need(obj&&typeof obj==="object");for(const k of keys.split(" "))need(typeof obj[k]==="string");}
const sfields="name address phone map hours entry rules", rfields="ticket item centerId priority due note createdAt", afields="name street phone purchase price months item fault cause demand other attachments date store director storeAddress";
const snapfields="centerId name address phone map hours entry ticket item note priority due";
function decode(raw){
 need(raw&&[1,2,3].includes(raw.version),"Неизвестный формат копии. Выберите JSON журнала.");
 const old=raw.version<3,out=newStore(), known={}; 
 for(const k of kinds){
  if(old&&k==="trips"){known.trips=new Set();continue}
  const ids=new Set();known[k]=ids;
  for(const x of arr(raw[k])){
   need(x&&validId(x.id)&&!ids.has(x.id));ids.add(x.id);
   let r={...x};delete r._v;
   if(k==="centers"){
    if(old){const c={id:r.id,name:clean(r.name,120),address:clean(r.address,300),phone:clean(r.phone,60),map:clean(r.map,700)};r={...c,hours:"",entry:"",rules:""}}
    strs(r,sfields);need(r.name&&r.address&&r.phone&&(!r.map||validYandexLink(r.map)));
   }else if(k==="requests"){
    if(old){need(["active","done"].includes(r.status));r={id:r.id,ticket:r.ticket,item:r.item,centerId:r.centerId,priority:r.priority,due:r.due||"",note:r.note||"",createdAt:r.createdAt||"",legacy:r.status==="done"?"done":r.inTrip?"selected":"",closed:r.status==="done",calls:[],documents:[],events:[]}}
    strs(r,rfields+" legacy");need(r.ticket&&r.item&&known.centers.has(r.centerId)&&Object.hasOwn(PRI,r.priority)&&typeof r.closed==="boolean"&&(!r.due||isDay(r.due)));
    const callIds=new Set();for(const c of arr(r.calls)){strs(c,"id date result note closedAt");need(validId(c.id)&&!callIds.has(c.id)&&isDay(c.date)&&["","ready","later","missed"].includes(c.result));callIds.add(c.id)}
    for(const doc of arr(r.documents)){strs(doc,afields+" id at");need(validId(doc.id))}
    for(const e of arr(r.events)){strs(e,"day text");need(isDay(e.day))}
   }else{
    need(isDay(r.date)&&Number.isSafeInteger(r.no)&&r.no>0&&["draft","sent","done"].includes(r.status));strs(r,"sentAt doneAt");const rid=new Set();
    for(const l of arr(r.lines)){need(known.requests.has(l.rid)&&!rid.has(l.rid)&&["drop","pick"].includes(l.kind)&&["","ok","no"].includes(l.outcome));rid.add(l.rid);strs(l,"reason");if(r.status!=="draft"){strs(l.snap,snapfields);need(!l.snap.map||validYandexLink(l.snap.map));}if(r.status==="done")need(l.outcome!=="")}
   }
   let v;if(raw.version===1){const legacy={...x};delete legacy._v;v=baseClock(legacy)}else v=readClock(x._v);
   out[k].push({...r,_v:v});
  }
  if(raw.version>=2) for(const x of arr(raw.removed?.[k])){need(x&&validId(x.id)&&!ids.has(x.id));ids.add(x.id);out.removed[k].push({id:x.id,label:clean(x.label,250),_v:readClock(x._v)})}
  known[k]=new Set(out[k].map(r=>r.id));
 }
 return out;
}
function combine(left,right,choices=[]){
 const plan=planMerge(left,right),s=newStore();
 for(const [i,c] of plan.conflicts.entries()){need(["a","b"].includes(choices[i]),"Выберите вариант для каждого расхождения.");const e=c[choices[i]];plan.records[c.k].set(c.id,{dead:e.dead,row:{...e.row,_v:advance(joinClock(c.a.row._v,c.b.row._v))}})}
 // Retain records still needed by an independently changed trip or request.
 function rescue(k,id){const v=plan.records[k].get(id);if(v&&!v.dead)return;const copy=[left,right].flatMap(s=>s[k]).find(x=>x.id===id);need(copy,"Утеряна связь с записью: "+id);plan.records[k].set(id,{dead:false,row:{...copy,_v:advance(joinClock(copy._v,v?.row._v||{}))}})}
 for(const e of plan.records.trips.values())if(!e.dead)for(const l of e.row.lines)rescue("requests",l.rid);
 for(const e of plan.records.requests.values())if(!e.dead)rescue("centers",e.row.centerId);
 for(const k of kinds)for(const e of plan.records[k].values())(e.dead?s.removed[k]:s[k]).push(e.row);
 return decode(s);
}
let S=newStore(),actor="a"+uid(),disk=null,meta={dirty:false,last:"",bound:false},locked=false,tab="cases",filter="open",term="",tripDay=today(),callFilter="due",caseID="",dialogDirty=false,dialogEditing=false,baselineLocal="",submitFn=null,afterClose=null,toTimer=0;
const req=(s,id)=>s.requests.find(x=>x.id===id),center=(s,id)=>s.centers.find(x=>x.id===id),trip=(s,id)=>s.trips.find(x=>x.id===id);
const ongoing=(s,rid)=>s.trips.filter(t=>t.status!=="done"&&t.lines.some(l=>l.rid===rid)).sort((a,b)=>a.date.localeCompare(b.date));
const pendingCall=r=>r.calls.find(c=>!c.result);
const dirLabel=k=>k==="pick"?"Забрать из СЦ":"Отвезти в СЦ",doneLabel=k=>k==="pick"?"Забрали из СЦ":"Отвезли в СЦ";
const tsLabel={draft:"Планируется",sent:"Передан водителю",done:"Завершён"},resLabel={ready:"Готово к выдаче",later:"Ещё не готово",missed:"Не дозвонились"};
function snapshot(s,rid){const r=req(s,rid),c=center(s,r.centerId);return Object.fromEntries(snapfields.split(" ").map(k=>[k,["name","address","phone","map","hours","entry"].includes(k)?c[k]:r[k]]))}
function groups(s,t){
 const map=new Map();
 for(const l of t.lines){const ss=l.snap||snapshot(s,l.rid);if(!map.has(ss.centerId))map.set(ss.centerId,{c:ss,items:[],rank:-1,limit:"9999",order:map.size});const g=map.get(ss.centerId);g.items.push({...l,snap:ss});g.rank=Math.max(g.rank,PRIORITY[ss.priority]||0);if(ss.due)g.limit=g.limit<ss.due?g.limit:ss.due}
 for(const g of map.values())g.items.sort((a,b)=>(PRIORITY[b.snap.priority]-PRIORITY[a.snap.priority])||(a.snap.due||"9999").localeCompare(b.snap.due||"9999"));
 return [...map.values()].sort((a,b)=>b.rank-a.rank||a.limit.localeCompare(b.limit)||a.order-b.order);
}
function newTrip(s,date){need(isDay(date),"Укажите дату рейса.");const no=Math.max(0,...s.trips.filter(t=>t.date===date).map(t=>t.no))+1,t={id:uid(),date,no,status:"draft",sentAt:"",doneAt:"",lines:[]};s.trips.push(t);return t}
function chooseTrip(s,date,choice){if(choice==="new")return newTrip(s,date);if(!choice||choice==="auto")return s.trips.find(t=>t.date===date&&t.status==="draft")||newTrip(s,date);const t=trip(s,choice);need(t?.status==="draft"&&t.date===date,"Выбранный рейс уже недоступен.");return t}
function addLine(s,rid,kind,date,choice){const r=req(s,rid);need(r&&!ongoing(s,rid).length,"Для инструмента уже назначен рейс. Сначала измените существующее назначение.");need(["drop","pick"].includes(kind)&&isDay(date),"Проверьте действие и дату рейса.");const t=chooseTrip(s,date,choice);t.lines.push({rid,kind,outcome:"",reason:"",snap:null});r.closed=false;if(r.legacy==="selected")r.legacy="";return t}
function sendTrip(s,id){const t=trip(s,id);need(t?.status==="draft"&&t.lines.length,"В рейсе нет инструментов.");for(const l of t.lines)l.snap=snapshot(s,l.rid);t.status="sent";t.sentAt=iso()}
function completeTrip(s,id,answers){const t=trip(s,id);need(t?.status==="sent","Сначала передайте лист водителю.");for(const l of t.lines){const a=answers[l.rid];need(a&&["ok","no"].includes(a.outcome)&&(a.outcome!=="no"||a.reason.trim()),"Укажите результат по всем инструментам и причину невыполнения.");l.outcome=a.outcome;l.reason=a.reason;const r=req(s,l.rid);if(l.kind==="pick"&&l.outcome==="ok")r.closed=!pendingCall(r)}t.status="done";t.doneAt=iso()}
function makeWorksheet(t){
 const g=groups(S,t),rows=[["ЛИСТ ВОДИТЕЛЯ · "+fmt(t.date)+" · рейс № "+t.no],["Остановок: "+g.length+" · инструментов: "+t.lines.length],[],["Точка","Действие","Сервисный центр","Адрес / приёмка","Телефон СЦ","Обращение","Инструмент","Комментарий","Срок"]];
 g.forEach((v,i)=>v.items.forEach(l=>{const r=l.snap;rows.push([String(i+1),dirLabel(l.kind),r.name,[r.address,r.hours,r.entry].filter(Boolean).join("\n"),r.phone,r.ticket,r.item,[r.note,r.priority!=="normal"?PRI[r.priority]:""].filter(Boolean).join("\n"),r.due?fmt(r.due):""])}));
 const ws=rows.map((row,i)=>'<row r="'+(i+1)+'">'+row.map((v,j)=>'<c r="'+excelColumn(j)+(i+1)+'" s="'+(i===0?1:i===3?2:0)+'" t="inlineStr"><is><t xml:space="preserve">'+xmlText(v)+'</t></is></c>').join("")+'</row>').join("");
 return '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>'+[9,17,25,44,19,17,31,34,15].map((v,i)=>'<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+v+'" customWidth="1"/>').join("")+'</cols><sheetData>'+ws+'</sheetData><autoFilter ref="A4:I'+rows.length+'"/><mergeCells count="2"><mergeCell ref="A1:I1"/><mergeCell ref="A2:I2"/></mergeCells><pageMargins left=".3" right=".3" top=".5" bottom=".5" header=".1" footer=".1"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>';
}
function loadLocal(){try{disk=localStorage.getItem(KEY);const a=localStorage.getItem(KEY+".actor");if(a&&/^a[a-z\d-]{3,105}$/i.test(a))actor=a;else localStorage.setItem(KEY+".actor",actor);
 if(disk){const raw=JSON.parse(disk);S=decode(raw);meta={...meta,...raw.local}}
 else {const src=localStorage.getItem(OLD);if(src){S=decode(JSON.parse(src));write(S,{...meta,dirty:true});}}
 }catch(e){locked=true;problem="Журнал не удалось прочитать: "+e.message}}
function write(s,m){need(!locked,"Хранилище недоступно.");need(localStorage.getItem(KEY)===disk,"В другой вкладке уже изменили журнал. Обновите страницу.");const raw=JSON.stringify({...s,local:m});localStorage.setItem(KEY,raw);disk=raw;S=s;meta=m}
function change(fn){const n=structuredClone(S);fn(n);trackChanges(S,n);const checked=decode(n);write(checked,{...meta,dirty:true});refresh();queueSync()}
const h=xmlText,tag=(s,cl="")=>'<span class="tag '+cl+'">'+h(s)+'</span>',b=(s,a,id="",cl="",aux="")=>'<button type="button" class="'+cl+'" data-act="'+a+'" data-id="'+h(id)+'" data-aux="'+h(aux)+'">'+h(s)+'</button>';
const mapLink=c=>'<a class="btn" href="'+h(mapUrl(c))+'" target="_blank" rel="noopener noreferrer">Яндекс Карты</a>';
const small=s=>'<span class="muted">'+h(s)+'</span>',blank=(a,text,action="")=>'<div class="empty"><h3>'+h(a)+'</h3><p>'+h(text)+'</p>'+action+'</div>';
const field=(name,label,value="",opt={})=>'<label class="f '+(opt.full?'full':'')+'">'+h(label)+(opt.multi?'<textarea name="'+name+'" maxlength="'+(opt.max||2000)+'" '+(opt.req?'required':'')+'>'+h(value)+'</textarea>':'<input name="'+name+'" type="'+(opt.type||'text')+'" value="'+h(value)+'" maxlength="'+(opt.max||300)+'" '+(opt.req?'required':'')+' '+(opt.extra||'')+'>')+'</label>';
const options=(arr,val)=>arr.map(([id,n])=>'<option value="'+h(id)+'" '+(id===val?'selected':'')+'>'+h(n)+'</option>').join("");
const sel=(name,label,items,val="",full=false)=>'<label class="f '+(full?'full':'')+'">'+h(label)+'<select name="'+name+'">'+options(items,val)+'</select></label>';
const area=(s,cl="")=>'<div class="info '+cl+'">'+s+'</div>',optDay={type:"date",req:true};
function note(s){$("toast").textContent=s;clearTimeout(toTimer);toTimer=setTimeout(()=>$("toast").textContent="",4200)}
function closeDlg(){afterClose=null;$("modal").close();dialogDirty=false;dialogEditing=false;submitFn=null;syncGuard()}
function leaveDlg(){if(dialogDirty&&!confirm("Закрыть форму без сохранения введённых изменений?"))return;const f=afterClose;closeDlg();if(f)f()}
function dialog(title,body,opts={}){const m=$("modal");dialogDirty=false;dialogEditing=!!opts.submit;syncGuard();baselineLocal=digest(S);afterClose=opts.back||null;submitFn=opts.submit||null;
 $("modalTitle").textContent=title;$("modalKicker").textContent=opts.kicker||"Журнал диагностики";
 $("modalContent").innerHTML=opts.submit?'<form id="editor">'+body+'<p class="formErr" role="alert"></p><div class="actions">'+b("Отмена","cancel")+'<button class="primary" type="submit">'+h(opts.save||"Сохранить")+'</button></div></form>':body+'<p class="formErr" role="alert"></p>';if(!m.open)m.showModal();m.scrollTop=0}
function formObj(f){return Object.fromEntries([...new FormData(f)].map(([k,v])=>[k,typeof v==="string"?v.trim():v]))}
const err=s=>{const el=$("modal").open&&$("modalContent").querySelector(".formErr");if(el)el.textContent=s;else note(s)};
function timeline(r){
 let lines=r.events.map(e=>({day:e.day,s:e.text,d:""}));
 for(const t of S.trips){for(const l of t.lines.filter(l=>l.rid===r.id)){
 const label=t.status==="done"?(l.outcome==="ok"?doneLabel(l.kind):"Не выполнено: "+dirLabel(l.kind)):t.status==="sent"?"У водителя: "+dirLabel(l.kind):"Запланировано: "+dirLabel(l.kind);
 lines.push({day:t.date,s:label,d:"Рейс № "+t.no+" · "+(l.snap?.name||center(S,r.centerId).name)+(l.reason?" · "+l.reason:"")})}}
 for(const c of r.calls)lines.push({day:c.result?localDay(c.closedAt):c.date,s:c.result?"Звонок: "+resLabel[c.result]:"Запланирован звонок в СЦ",d:[c.result?"План: "+fmt(c.date):"",c.note].filter(Boolean).join(" · ")});
 for(const d of r.documents)lines.push({day:localDay(d.at),s:"Оформлено заявление",d:"Дата заявления: "+fmt(d.date)});
 if(r.legacy==="done")lines.push({day:"",s:"Выполнено в прежнем журнале",d:"Дата рейса не указана."});
 return lines.sort((a,b)=>a.day.localeCompare(b.day));
}
function localDay(s){if(!s)return "";const t=new Date(s);if(isNaN(t))return "";return [t.getFullYear(),String(t.getMonth()+1).padStart(2,"0"),String(t.getDate()).padStart(2,"0")].join("-")}
function nextSteps(r){const tasks=ongoing(S,r.id).map(t=>({date:t.date,s:dirLabel(t.lines.find(l=>l.rid===r.id).kind),cl:t.lines.find(l=>l.rid===r.id).kind}));const c=pendingCall(r);if(c)tasks.push({date:c.date,s:"Позвонить в СЦ",cl:"call"});return tasks.sort((a,b)=>a.date.localeCompare(b.date))}
const priTag=r=>r.priority!=="normal"?tag(PRI[r.priority],r.priority):"";
const dataCell=(name,body)=>'<td data-label="'+h(name)+'">'+body+'</td>';
function caseRows(){const q=term.toLowerCase().replace(/^№\s*/,"");const rs=S.requests.filter(r=>(q||filter==="all"||(filter==="closed"?r.closed:!r.closed))&&(!q||[r.ticket,r.item,center(S,r.centerId).name,r.note].some(s=>s.toLowerCase().includes(q)))).sort((a,b)=>Number(a.closed)-Number(b.closed)||(PRIORITY[b.priority]-PRIORITY[a.priority])||(nextSteps(a)[0]?.date||a.due||"9999").localeCompare(nextSteps(b)[0]?.date||b.due||"9999")||a.ticket.localeCompare(b.ticket,"ru",{numeric:true}));
 if(!rs.length)return blank(q?"Ничего не найдено":filter==="closed"?"Завершённых обращений пока нет":"Нет обращений в этом списке",q?"Поиск идёт и по активным, и по завершённым обращениям.":"Добавьте обращение или откройте прежнюю копию журнала.","");
 const buckets=new Map();for(const r of rs){if(!buckets.has(r.centerId))buckets.set(r.centerId,[]);buckets.get(r.centerId).push(r)}
 return [...buckets].map(([cid,items],i)=>{const c=center(S,cid),priority=items.every(r=>r.closed)?"done":items.reduce((a,r)=>PRIORITY[r.priority]>PRIORITY[a]?r.priority:a,"normal");
 return '<article class="stop-card priority-'+priority+'"><div class="stop-main"><span aria-hidden="true" class="stop-number">'+String(i+1).padStart(2,"0")+'</span><div class="stop-content"><div class="stop-first-line"><strong class="stop-name">'+h(c.name)+'</strong><span class="badge badge-'+priority+'">'+h(priority==="done"?"Завершено":PRI[priority])+'</span></div><p class="stop-address">'+h(c.address)+'</p><div class="stop-meta"><span>'+h(c.phone)+'</span><a href="'+h(mapUrl(c))+'" target="_blank" rel="noopener noreferrer">Яндекс Карты</a></div></div></div><div class="card-tickets">'+items.map(r=>{const ns=nextSteps(r);return '<div class="case-line"><div>'+b("№ "+r.ticket,"open",r.id,"link")+' · <strong>'+h(r.item)+'</strong></div><div class="tags">'+(ns.length?ns.map(x=>tag(x.s+" · "+fmt(x.date),x.cl)).join(" "):tag(r.closed?"Завершено":r.legacy==="selected"?"Из прежнего листа":"Не назначено",r.closed?"ok":""))+'</div></div>'}).join("")+'</div></article>'}).join("")+small("Обращений: "+rs.length+" · Нажмите на номер, чтобы открыть историю.");
}
function card(r){const c=center(S,r.centerId),hist=timeline(r),steps=nextSteps(r),pending=pendingCall(r),ot=ongoing(S,r.id)[0];
 return '<article class="panel pad"><div class="detailsgrid"><div class="sidebar"><p class="kicker">Инструмент</p><h2>'+h(r.item)+'</h2><p>'+tag(r.closed?"Завершено":"В работе",r.closed?"ok":"")+' '+priTag(r)+'</p><p><strong>'+h(c.name)+'</strong><br>'+h(c.address)+'<br>'+h(c.phone)+'</p><p>'+mapLink(c)+'</p><p>'+small("Доставить до: "+fmt(r.due))+'</p>'+(r.note?'<p class="pre">'+h(r.note)+'</p>':'')+'</div><div><h3>Следующее действие</h3><div class="stack">'+(steps.length?steps.map(x=>'<div>'+tag(x.s,x.cl)+' &nbsp;'+fmt(x.date)+'</div>').join(""):small("Пока ничего не назначено."))+'</div><div class="btnrow" style="margin:12px 0">'+(ot?b("Открыть рейс","goTrip",ot.id):b("Едем в СЦ","plan",r.id,"","drop")+b("Забираем из СЦ","plan",r.id,"","pick"))+(pending?b("Результат звонка","result",r.id):b("Позвонить в СЦ","call",r.id))+'</div><h3>История инструмента</h3><ul class="timeline">'+hist.map(x=>'<li><time>'+h(x.day?fmt(x.day):"Без даты")+'</time><strong>'+h(x.s)+'</strong>'+small(x.d)+'</li>').join("")+'</ul></div></div></article><section class="panel pad"><div class="sectionline"><h3>Заявления</h3><div class="btnrow">'+b("Оформить заявление","statement",r.id,"primary")+b("Бланк от руки","hand",r.id)+'</div></div>'+r.documents.map((d,i)=>'<div class="itemrow"><span>Заявление от '+fmt(d.date)+' · '+h(d.name||"без ФИО")+'</span>'+b("Печать","printDoc",r.id,"",String(i))+'</div>').join("")+'</section><div class="btnrow">'+b("Изменить обращение","editCase",r.id)+(!ot&&!pending?b(r.closed?"Вернуть в работу":"Закрыть обращение","toggleClosed",r.id):"")+(!r.calls.length&&!r.documents.length&&!S.trips.some(t=>t.lines.some(l=>l.rid===r.id))?b("Удалить","deleteCase",r.id,"quiet danger"):"")+'</div>';
}
function tripBody(t,history=false){const gs=groups(S,t);return '<article class="panel" id="trip-'+h(t.id)+'"><div class="panelhead"><div><p class="kicker">'+fmt(t.date)+'</p><h2>Рейс № '+t.no+' &nbsp;'+tag(tsLabel[t.status],t.status==="done"?"ok":t.status==="sent"?"sent":"")+'</h2><div class="muted">Остановок: '+gs.length+' · инструментов: '+t.lines.length+'</div></div><div class="btnrow">'+(t.status==="draft"?b("Добавить инструмент","addToTrip",t.id)+(t.lines.length?b("Передать водителю","dispatch",t.id,"primary"):b("Удалить пустой рейс","deleteTrip",t.id,"quiet danger")):t.status==="sent"?b("Завершить рейс","finish",t.id,"primary"):"")+(t.lines.length?b("Печать листа","printTrip",t.id)+b("Excel","xlsx",t.id):"")+'</div></div>'+gs.map((g,i)=>'<div class="stop"><div class="stophead"><div><h3>'+String(i+1).padStart(2,"0")+' · '+h(g.c.name)+'</h3>'+h(g.c.address)+' · '+h(g.c.phone)+(g.c.hours||g.c.entry?'<div class="muted">'+h([g.c.hours,g.c.entry].filter(Boolean).join(" · "))+'</div>':'')+'</div>'+mapLink(g.c)+'</div>'+g.items.map(l=>'<div class="itemrow"><div>'+b("№ "+l.snap.ticket,"open",l.rid,"link")+' &nbsp; <strong>'+h(l.snap.item)+'</strong>'+(l.snap.note?'<div class="muted pre">'+h(l.snap.note)+'</div>':'')+'</div><div class="tags">'+tag(dirLabel(l.kind),l.kind)+priTag(l.snap)+(t.status==="draft"?b("Убрать","unplan",t.id,"quiet",l.rid):l.outcome?tag(l.outcome==="ok"?doneLabel(l.kind):"Не выполнено",l.outcome==="ok"?"ok":"bad"):"")+(l.reason?small(l.reason):"")+'</div></div>').join("")+'</div>').join("")+'</article>'}
function tripList(){const ts=S.trips.filter(t=>t.date===tripDay).sort((a,b)=>a.no-b.no);const old=S.requests.filter(r=>r.legacy==="selected"&&!ongoing(S,r.id).length);return (old.length?area("Из прежнего листа ожидают назначения: "+old.length+" &nbsp; "+b("Назначить на "+fmt(tripDay),"legacySelected"),"warning"):"")+(ts.map(t=>tripBody(t)).join("")||blank("На эту дату рейсов нет","Создайте рейс или назначьте поездку из карточки обращения.",b("Создать рейс","newTrip","","primary")))}
function historyView(){const by=new Map();for(const t of S.trips.filter(t=>t.status==="done").sort((a,b)=>b.date.localeCompare(a.date)||a.no-b.no)){if(!by.has(t.date))by.set(t.date,[]);by.get(t.date).push(t)}
 const old=S.requests.filter(r=>r.legacy==="done");return [...by].map(([d,ts],i)=>'<details class="day" '+(i?"":"open")+'><summary class="dayhead">'+fmt(d)+' &nbsp;·&nbsp; рейсов: '+ts.length+' · инструментов: '+ts.reduce((a,t)=>a+t.lines.length,0)+'</summary>'+ts.map(t=>tripBody(t,true)).join("")+'</details>').join("")+(old.length?'<section class="panel pad"><h3>Дата рейса не указана</h3><p class="muted">Отметьте старые записи из одного бумажного листа и назначьте им дату.</p>'+old.map(r=>'<label class="checkrow"><input class="legacyCheck" type="checkbox" value="'+h(r.id)+'">№ '+h(r.ticket)+' · '+h(r.item)+'</label>').join("")+b("Указать дату выбранным","legacyDate")+'</section>':"")+(!by.size&&!old.length?blank("Завершённых рейсов пока нет","После подтверждения результатов рейс появится здесь под датой выезда."):"")}
function callsView(){const ls=S.requests.flatMap(r=>r.calls.map(c=>({r,c}))).filter(({c})=>callFilter==="past"?!!c.result:!c.result&&(callFilter==="all"||c.date<=today())).sort((a,b)=>callFilter==="past"?b.c.closedAt.localeCompare(a.c.closedAt):a.c.date.localeCompare(b.c.date));
 return ls.length?'<div class="panel">'+ls.map(({r,c})=>{const sc=center(S,r.centerId);return '<div class="itemrow"><div>'+tag(c.result?resLabel[c.result]:c.date<today()?"Просрочен":"Позвонить",c.result?"ok":c.date<today()?"critical":"call")+' &nbsp; '+fmt(c.date)+'<h3>'+h(sc.name)+' · '+h(sc.phone)+'</h3>'+b("№ "+r.ticket,"open",r.id,"link")+' · '+h(r.item)+(c.note?'<p class="muted">'+h(c.note)+'</p>':'')+'</div>'+(c.result?small("Звонок: "+fmt(localDay(c.closedAt))):b("Записать результат","result",r.id,"primary"))+'</div>'}).join("")+'</div>':blank(callFilter==="due"?"На сегодня звонков нет":"Список пуст","Назначьте «Позвонить в СЦ» из карточки обращения.")}
function centersView(){return S.centers.length?'<div class="cluster">'+[...S.centers].sort((a,b)=>a.name.localeCompare(b.name,"ru")).map(c=>'<article class="centerbox"><p class="kicker">Сервисный центр</p><h2>'+h(c.name)+'</h2><p>'+h(c.address)+'</p><p>'+h(c.phone)+'</p>'+[c.hours,c.entry,c.rules].filter(Boolean).map(s=>'<p class="muted pre">'+h(s)+'</p>').join("")+'<div class="btnrow">'+mapLink(c)+b("Изменить","center",c.id)+b("Удалить","deleteCenter",c.id,"quiet danger")+'</div></article>').join("")+'</div>':blank("Добавьте первый сервисный центр","Сохраните адрес, телефон и часы приёмки один раз.",b("Добавить СЦ","center","","primary"))}
function refresh(){
 const titles={cases:"Маршрутный лист",trips:"Рейсы",calls:"Кому позвонить",history:"Выполненные рейсы",centers:"Адреса СЦ",card:"Обращение № "+(req(S,caseID)?.ticket||"")};if(tab==="card"&&!req(S,caseID))tab="cases";
 $("pageTitle").textContent=titles[tab];$("pageKicker").textContent=new Intl.DateTimeFormat("ru-RU",{weekday:"long",day:"numeric",month:"long",year:"numeric"}).format(new Date());
 classicView();
 document.querySelectorAll("nav button").forEach(v=>{const on=v.dataset.id===(tab==="card"?"cases":tab);v.classList.toggle("active",on);if(on)v.setAttribute("aria-current","page");else v.removeAttribute("aria-current")});
 $("pageButtons").innerHTML=tab==="card"?b("К обращениям","tab","cases"):tab==="centers"?b("+ Добавить СЦ","center","","primary"):tab==="trips"?b("+ Новый рейс","newTrip","","primary"):b("+ Новое обращение","newCase","","primary");
 const fl=(id,s,cur,action)=>b(s,action,id,cur===id?"activeFilter":"");
 $("filters").innerHTML=tab==="cases"&&!term?[["open","В работе"],["closed","Завершённые"],["all","Все"]].map(([i,n])=>fl(i,n,filter,"filter")).join(""):tab==="trips"?'<label>Дата выезда <input id="tripDay" type="date" value="'+tripDay+'"></label>':tab==="calls"?[["due","На сегодня"],["all","Все предстоящие"],["past","История звонков"]].map(([i,n])=>fl(i,n,callFilter,"callfilter")).join(""):"";
 $("content").innerHTML=tab==="cases"?caseRows():tab==="card"?card(req(S,caseID)):tab==="trips"?tripList():tab==="history"?historyView():tab==="calls"?callsView():centersView();renderStatus();
}
const centersOpts=()=>[["","Выберите СЦ"],...[...S.centers].sort((a,b)=>a.name.localeCompare(b.name,"ru")).map(c=>[c.id,c.name+" — "+c.address])];
const tripOpts=date=>[["auto","Первый планируемый или новый"],...S.trips.filter(t=>t.date===date&&t.status==="draft").sort((a,b)=>a.no-b.no).map(t=>[t.id,"Рейс № "+t.no]),["new","Создать отдельный рейс"]];
function tripInputs(date,kind="drop",choice="auto"){return sel("kind","Действие",[["drop","Едем в СЦ"],["pick","Забираем из СЦ"]],kind)+field("tripDate","Дата выезда",date,optDay)+sel("tripId","Рейс на эту дату",tripOpts(date),choice,true)}
function saveRequest(v,id=""){
 need(v.ticket&&v.item&&center(S,v.centerId),"Укажите номер, инструмент и сервисный центр.");if(!id&&S.requests.some(x=>x.ticket.toLowerCase()===v.ticket.toLowerCase()&&x.item.toLowerCase()===v.item.toLowerCase())&&!confirm("Уже есть обращение с таким номером и инструментом. Добавить ещё один инструмент?"))return false;
 let rid=id||uid();change(s=>{let row=req(s,rid);if(!row){row={id:rid,createdAt:iso(),legacy:"",closed:false,events:[{day:today(),text:"Создано обращение"}],calls:[],documents:[]};s.requests.push(row)}Object.assign(row,{ticket:v.ticket,item:v.item,centerId:v.centerId,priority:v.priority,due:v.due,note:v.note});if(!id&&v.step!=="none"){if(v.step==="call")addCall(row,v.stepDate,"");else addLine(s,rid,v.step,v.stepDate,v.tripId)}});tab="card";caseID=rid;term="";$("search").value="";refresh();note("Обращение сохранено.");
}
function requestForm(id="",draft=null){const r=draft||req(S,id)||{priority:"normal",centerId:"",step:"none"};
 const inline=()=>{const data=formObj($("editor"));centerForm("",c=>{if(c)data.centerId=c;requestForm(id,data);dialogDirty=true;syncGuard()})};
 dialog(id?"Изменить обращение":"Новое обращение",'<div class="grid">'+field("ticket","Номер обращения",r.ticket,{req:true,max:80})+field("item","Инструмент и модель",r.item,{req:true,max:160})+sel("centerId","Сервисный центр",centersOpts(),r.centerId,true)+'<div class="full"><div id="suggestions" class="muted"></div>'+b("+ Новый СЦ","inlineCenter","","quiet")+' <span id="formMap"></span></div>'+sel("priority","Приоритет",Object.entries(PRI),r.priority)+field("due","Доставить до (необязательно)",r.due,{type:"date"})+field("note","Комментарий водителю",r.note,{multi:true,full:true,max:800})+(id?"":sel("step","Следующий шаг",[["none","Назначить позже"],["drop","Едем в СЦ"],["pick","Забираем из СЦ"],["call","Позвонить в СЦ"]],r.step||"none",true)+'<div id="stepWhen" class="full grid" hidden>'+field("stepDate","Дата действия",r.stepDate||today(),{type:"date"})+sel("tripId","Рейс",tripOpts(r.stepDate||today()),r.tripId||"auto")+'</div>')+'</div>',{submit:(v)=>{
 return saveRequest(v,id);
 },save:id?"Сохранить":"Создать обращение"});inlineAction=inline;updateRequestHints();
}
let inlineAction=null;
function updateRequestHints(){const f=$("editor");if(!f?.elements.namedItem("item")||!$("suggestions"))return;const item=f.elements.namedItem("item").value.toLowerCase();const suggested=S.centers.filter(c=>c.rules.split(/[,;\n]/).some(x=>x.trim()&&item.includes(x.trim().toLowerCase())));$("suggestions").innerHTML=suggested.length?"Подходящие СЦ: "+suggested.map(c=>b(c.name,"suggest",c.id,"quiet")).join(""):"";const c=center(S,f.elements.centerId.value);$("formMap").innerHTML=c?mapLink(c):"";if(f.elements.step){$("stepWhen").hidden=f.elements.step.value==="none";f.elements.stepDate.required=!$("stepWhen").hidden;f.elements.tripId.parentElement.hidden=f.elements.step.value==="call"}}
function centerForm(id="",back=null){const c=center(S,id)||{};dialog(id?"Изменить сервисный центр":"Новый сервисный центр",'<div class="grid">'+field("name","Название СЦ",c.name,{req:true,max:120})+field("phone","Телефон СЦ",c.phone,{req:true,max:60,type:"tel"})+field("address","Полный адрес с городом",c.address,{req:true,full:true,max:300})+field("hours","Часы приёмки",c.hours)+field("entry","Вход / место передачи",c.entry)+field("rules","Для каких инструментов или брендов",c.rules,{full:true})+'<p class="fieldnote full">Через запятую, например: Makita, перфоратор. По этим словам журнал предложит центр при вводе инструмента.</p>'+field("map","Ссылка на точку в Яндекс Картах (необязательно)",c.map,{full:true,max:700})+'</div>'+b("Проверить адрес на карте","testMap","","quiet"),{submit:v=>{
 need(v.name&&v.phone&&v.address,"Заполните название, телефон и адрес.");need(!v.map||validYandexLink(v.map),"Укажите ссылку вида https://yandex.ru/maps/… или оставьте поле пустым.");const cid=id||uid();change(s=>{const found=center(s,cid);if(found)Object.assign(found,v);else s.centers.push({id:cid,...v})});if(back){closeDlg();back(cid);return false}note("Сервисный центр сохранён.")},back:back?()=>back(null):null})}
function planForm(rid,kind){dialog("Назначить поездку",'<p>'+h(req(S,rid).item)+' · № '+h(req(S,rid).ticket)+'</p><div class="grid">'+tripInputs(today(),kind)+'</div>',{submit:v=>{let t;change(s=>{t=addLine(s,rid,v.kind,v.tripDate,v.tripId)});note("Назначен рейс № "+t.no+" на "+fmt(t.date))},save:"Добавить в рейс"})}
function addToTrip(t){const rs=S.requests.filter(r=>!ongoing(S,r.id).length);if(!rs.length){note("Нет свободных обращений. Создайте новое обращение.");return}
 dialog("Добавить инструмент в рейс",'<div class="grid">'+sel("rid","Обращение",rs.map(r=>[r.id,"№ "+r.ticket+" · "+r.item]),rs[0].id,true)+sel("kind","Действие",[["drop","Едем в СЦ"],["pick","Забираем из СЦ"]],"drop")+'</div>',{kicker:fmt(t.date)+" · рейс № "+t.no,submit:v=>change(s=>addLine(s,v.rid,v.kind,t.date,t.id)),save:"Добавить"})}
function finishForm(t){dialog("Результаты рейса",'<p class="muted">Подтвердите результат для каждого инструмента. Датой поездки останется '+fmt(t.date)+'.</p>'+t.lines.map((l,i)=>'<fieldset><legend>№ '+h(l.snap.ticket)+' · '+h(l.snap.item)+'</legend><div class="grid">'+sel("result"+i,dirLabel(l.kind),[["","Выберите результат"],["ok",doneLabel(l.kind)],["no","Не выполнено"]],"")+field("reason"+i,"Причина, если не выполнено","",{max:500})+'</div></fieldset>').join(""),{submit:v=>{const a=Object.fromEntries(t.lines.map((l,i)=>[l.rid,{outcome:v["result"+i],reason:v["reason"+i]}]));change(s=>completeTrip(s,t.id,a));note("Рейс сохранён в истории за "+fmt(t.date));},save:"Подтвердить и завершить"})}
function addCall(r,date,comment){need(isDay(date),"Укажите дату звонка.");need(!pendingCall(r),"Звонок уже запланирован.");r.calls.push({id:uid(),date,result:"",note:comment,closedAt:""});r.closed=false}
function callForm(r){dialog("Позвонить в СЦ",'<p class="muted">'+h(center(S,r.centerId).name)+' · '+h(center(S,r.centerId).phone)+'</p><div class="grid">'+field("date","Дата звонка",today(),optDay)+field("note","Что уточнить","",{full:true,multi:true})+'</div>',{kicker:"Обращение № "+r.ticket,submit:v=>change(s=>addCall(req(s,r.id),v.date,v.note)),save:"Назначить звонок"})}
function resultForm(r){const c=pendingCall(r);need(c,"Нет запланированного звонка.");const already=ongoing(S,r.id)[0];dialog("Результат звонка",'<p class="muted">'+h(center(S,r.centerId).name)+' · '+h(center(S,r.centerId).phone)+'</p><div class="grid">'+sel("out","Что ответили в СЦ",Object.entries(resLabel),"ready",true)+field("comment","Комментарий","",{multi:true,full:true})+field("tripDate","Дата забора",today(),optDay)+sel("tripId","Рейс",tripOpts(today()),"auto")+'</div>'+(already?area("У инструмента уже есть рейс № "+already.no+" на "+fmt(already.date)+". Для готового инструмента предварительно проверьте это назначение."):""),{submit:v=>{change(s=>{const row=req(s,r.id),active=pendingCall(row);need(active?.id===c.id,"Звонок изменён.");need(isDay(v.tripDate),"Введите дату следующего действия.");active.result=v.out;active.closedAt=iso();active.note=[active.note,v.comment].filter(Boolean).join(" · ");
 if(v.out==="ready"){if(already){const l=already.lines.find(l=>l.rid===r.id);need(l.kind==="pick","Сначала завершите или снимите назначенную доставку.");}else addLine(s,r.id,"pick",v.tripDate,v.tripId)}else addCall(row,v.tripDate,"Повторный звонок");});note("Результат звонка записан.")},save:"Записать результат"})}
const storeDefault={store:'ООО «СТД Петрович»',director:"Мовчану Е.М.",storeAddress:"Московская обл., г.о. Люберцы, пгт. Томилино, км 23-й Новорязанского шоссе, 19А"};
const demands=[["repair","Устранить возникшие неисправности (ремонт)"],["refund","Вернуть денежные средства"],["replace","Заменить на новый"],["other","Другое"]];
function statementForm(r){const d={...storeDefault,...r.documents.at(-1),date:today(),item:r.item};dialog("Оформить заявление",'<div class="grid">'+field("name","ФИО заявителя",d.name,{full:true,req:true})+field("street","Адрес заявителя",d.street)+field("phone","Телефон заявителя",d.phone,{type:"tel"})+field("date","Дата заявления",d.date,optDay)+field("purchase","Дата покупки",d.purchase,{type:"date"})+field("item","Полное наименование и марка товара",d.item,{full:true,req:true})+field("price","Цена, руб. (с копейками)",d.price,{type:"number",extra:'min="0" step="0.01"'})+field("months","Гарантия, месяцев",d.months,{type:"number",extra:'min="0" step="1"'})+field("fault","Недостатки со слов клиента",d.fault,{multi:true,full:true})+field("cause","Предполагаемая причина со слов клиента",d.cause,{multi:true,full:true})+sel("demand","Требование клиента",[["","Не отмечено"],...demands],d.demand||"",true)+field("other","Уточнение к требованию «Другое»",d.other,{full:true})+field("attachments","Приложенные документы",d.attachments,{full:true})+'</div><details><summary>Реквизиты магазина для этого заявления</summary><div class="grid">'+field("store","Организация",d.store,{req:true})+field("director","Директору",d.director,{req:true})+field("storeAddress","Адрес магазина",d.storeAddress,{full:true,req:true})+'</div></details>',{kicker:"Обращение № "+r.ticket,save:"Сохранить и распечатать",submit:v=>{const doc={...v,id:uid(),at:iso()};change(s=>req(s,r.id).documents.push(doc));closeDlg();printStatement(doc);return false}})}
function printStatement(d,hand=false){const write=(s,extra="")=>'<div class="write '+extra+'">'+h(s||" ")+'</div>';const part=d.price!==""&&d.price!=null?Number(d.price).toFixed(2).split("."):["________","____"];
 const text=demands.map(([k,v],i)=>'<span class="'+(d.demand===k?"underline":"")+'">'+(i+1)+'. '+h(v)+(k==="replace"?' (в течение пятнадцати дней со дня получения товара. Ст. 18 пункт 1 «Закона о защите прав потребителя»)':k==="other"?' '+h(d.other||"__________"):"")+'</span>').join("; ");
 const html='<article class="statement"><div class="to">Директору '+h(d.store)+'<br>'+h(d.director)+'<br>'+h(d.storeAddress)+'<br>от '+write(d.name)+'проживающего по адресу: '+write(d.street)+'Телефон: '+write(d.phone)+'</div><h1>Заявление</h1><p>'+(d.purchase?fmt(d.purchase):"«____» __________ 20____ г.")+' я приобрёл(а) в вашем магазине:</p>'+write(d.item)+'<p>По цене '+h(part[0])+' рублей '+h(part[1])+' коп., с установленным гарантийным сроком '+h(d.months||"________")+' месяцев, что подтверждается кассовым чеком и гарантийным талоном.</p><p>Данное изделие, по моему мнению, имеет следующие недостатки:</p>'+write(d.fault,"extra")+'<p>Причиной возникновения указанных недостатков считаю:</p>'+write(d.cause)+'<p>На основании вышеизложенного прошу (нужное подчеркнуть): '+text+'.</p><p>В подтверждение факта покупки в вашем предприятии прикладываю следующие документы:</p>'+write(d.attachments)+'<div class="fine">'+templateText.map(v=>'<p>'+h(v)+'</p>').join("")+'</div><div class="sign"><span>Дата: '+(hand?"________________":fmt(d.date))+'</span><span>Подпись: ____________________</span></div></article>';printPaper(html)}
function printPaper(html){$("paper").innerHTML=html;window.print()}
function driverPrint(t){const gs=groups(S,t);printPaper('<div class="pphead"><h1>Лист водителя · рейс № '+t.no+'</h1><p>Дата выезда: '+fmt(t.date)+' · остановок: '+gs.length+' · инструментов: '+t.lines.length+'</p></div>'+gs.map((g,i)=>'<section class="printstop"><h3>'+(i+1)+'. '+h(g.c.name)+'</h3><b>'+h(g.c.address)+'</b><br>Телефон: '+h(g.c.phone)+'<br>'+h([g.c.hours,g.c.entry].filter(Boolean).join(" · "))+'<table class="printrows"><thead><tr><th>Действие</th><th>Обращение</th><th>Инструмент</th><th>Комментарий / срок</th></tr></thead><tbody>'+g.items.map(l=>'<tr>'+[dirLabel(l.kind),l.snap.ticket,l.snap.item,[l.snap.note,l.snap.due?"До "+fmt(l.snap.due):"",l.snap.priority!=="normal"?PRI[l.snap.priority]:""].filter(Boolean).join("\n")].map(x=>'<td>'+h(x)+'</td>').join("")+'</tr>').join("")+'</tbody></table></section>').join("")+'<p>Передал: ____________________ &nbsp; Водитель: ____________________</p>')}

// Local changes are committed first. The shared JSON is replaced only after a verified read and merge.
let folder=null,problem="",busy=false,again=false,timer=0,permitted=false,conflict=null,seen=false;
const enc=s=>JSON.stringify(s,null,2),fileStamp=()=>new Date().toISOString().replace(/[:.]/g,"-");
function unload(e){e.preventDefault();e.returnValue=""}
function syncGuard(){window.removeEventListener("beforeunload",unload);if(meta.dirty||dialogDirty||deskDirty||busy)window.addEventListener("beforeunload",unload)}
function status(){return locked?problem:conflict?"Нужно сверить две версии журнала":problem?problem:busy?"Проверяем общую папку…":meta.bound&&permitted?(meta.dirty?"Изменения ожидают записи в общую папку":"Общая папка · сохранено "+time(meta.last)):meta.bound?"Для общей папки нужен доступ":meta.dirty?"Сохранено на этом компьютере · перед сменой ПК сохраните копию":"Журнал на этом компьютере"}
function renderStatus(){const bad=!!(locked||problem||conflict),warn=!!(meta.dirty||meta.bound&&!permitted);
 $("statusBar").classList.toggle("pending",warn);$("statusBar").classList.toggle("error",bad);$("headState").textContent=meta.bound&&permitted&&!problem&&!conflict?"Общая папка":"Журнал на этом ПК";
 $("statusBar").innerHTML='<span><i class="statusDot '+(bad?"err":warn?"warn":"")+'"></i>'+h(status())+'</span><div class="btnrow">'+(conflict?b("Сверить версии","review","","quiet"):meta.bound?(permitted?b("Проверить сейчас","sync","","quiet"):b("Восстановить доступ","grant","","quiet")):(meta.dirty?b("Сохранить копию","export","","quiet"):"")+b("Открыть копию","import","","quiet")+(window.showDirectoryPicker?b("Подключить папку","connect","","quiet"):""))+'</div>';syncGuard()}
function storageView(){dialog("Сохранение журнала",'<label class="f">Оформление на этом компьютере<select id="classicTheme">'+options([["blue","Классическое · синее"],["green","Диспетчерская · зелёное"]],document.documentElement.dataset.theme||"blue")+'</select></label>'+area(h(status()),problem||meta.dirty?"warning":"")+'<h3>Общая папка для двух компьютеров</h3><p>По очереди открывайте сайт на обоих ПК. В Chrome или Edge на каждом выберите одну и ту же сетевую папку. После подключения журнал проверяет общий файл при открытии, возвращении на вкладку и во время работы. Перед переходом на другой ПК дождитесь надписи «Общая папка · сохранено».</p>'+(folder?'<p><strong>Папка: '+h(folder.name)+'</strong></p>':'')+'<div class="btnrow">'+(window.showDirectoryPicker?b(folder?"Сменить папку":"Выбрать общую папку","connect","","primary"):"")+((folder||meta.bound)?b(permitted?"Проверить сейчас":"Разрешить доступ",permitted?"sync":"grant")+b("Отключить","disconnect"):"")+(conflict?b("Сверить версии","review"):"")+'</div><h3>Копия в файле JSON</h3><p>Если браузер не даёт выбрать папку, сохраните копию после смены, а на другом компьютере откройте её перед работой. При загрузке журнал покажет, какие записи добавятся. В общей папке также хранятся три предыдущие версии файла в каталоге «Резервные копии».</p><div class="btnrow">'+b("Сохранить копию","export")+b("Открыть копию","import")+'</div><details><summary>Восстановление прежнего состояния</summary><p>Для отмены более поздних изменений выберите один из файлов из каталога «Резервные копии».</p>'+b("Вернуться к резервной копии","restore")+'</details><p class="fieldnote">Файл данных может содержать сведения заявителей. Выберите рабочую папку с ограниченным доступом. На GitHub загружаются только три файла сайта.</p>')}
async function remembered(mode,value){
 if(!window.indexedDB)throw Error("Браузер не поддерживает запоминание папки.");const db=await new Promise((res,rej)=>{const q=indexedDB.open("sc-journal-folders",1);q.onupgradeneeded=()=>q.result.createObjectStore("handles");q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error)});
 try{return await new Promise((res,rej)=>{const tx=db.transaction("handles",mode==="get"?"readonly":"readwrite"),store=tx.objectStore("handles"),q=mode==="get"?store.get("main"):mode==="set"?store.put(value,"main"):store.delete("main");let result;q.onsuccess=()=>result=q.result;tx.oncomplete=()=>res(result);tx.onerror=tx.onabort=()=>rej(tx.error||Error("Не удалось запомнить папку."))})}finally{db.close()}
}
async function readMain(dir){let f;try{f=await dir.getFileHandle(MAIN)}catch(e){if(e.name==="NotFoundError")return null;throw e}return await (await f.getFile()).text()}
async function safeFile(f,text){const stream=await f.createWritable();try{await stream.write(text);await stream.close()}catch(e){try{await stream.abort()}catch{}throw e}}
function queueSync(delay=250){if(!folder||!permitted||locked||conflict)return;clearTimeout(timer);timer=setTimeout(()=>syncNow(),delay)}
async function syncNow(){
 if(busy){again=true;return}if(!folder||!permitted||locked||conflict||dialogEditing)return;
 const dir=folder;busy=true;renderStatus();let created=false;
 try{
  const before=await readMain(dir);if(dir!==folder)return;
  if(before===null&&(seen||meta.seen))throw Error("Общий файл исчез. Верните файл в папку или подключите новую папку.");
  let obj=null,remote=newStore();if(before!==null){try{obj=JSON.parse(before);remote=decode(obj)}catch(e){throw Error("Общий файл не прочитан и оставлен без изменений. Проверьте его или верните исправную копию в папку. "+e.message)}}
  if(dialogEditing)return;const start=digest(S),plan=planMerge(S,remote);if(plan.conflicts.length){conflict={s:remote,dir};problem="";return}
  const merged=combine(S,remote),md=digest(merged),now=iso();let stamp=obj?.savedAt||now;
  if(before===null||md!==digest(remote)){
   const seq=Number.isSafeInteger(obj?.sequence)?obj.sequence+1:1;
   if(before!==null){const bdir=await dir.getDirectoryHandle("Резервные копии",{create:true});await safeFile(await bdir.getFileHandle("Журнал-"+(1+((seq-2)%3+3)%3)+".json",{create:true}),before)}
   if(dir!==folder||dialogEditing||digest(S)!==start){again=true;return}
   const f=await dir.getFileHandle(MAIN,{create:true}),w=await f.createWritable({mode:"exclusive"});created=before===null;
   try{const latest=await (await f.getFile()).text();if(latest!==(before??"")||dir!==folder||dialogEditing||digest(S)!==start){await w.abort();again=true;return}
    const next=enc({...merged,sequence:seq,savedAt:now});await w.write(next);await w.close();stamp=now;
   }catch(e){try{await w.abort()}catch{}throw e}
  }
  if(dir!==folder)return;seen=true;problem="";
  if(digest(S)!==start||dialogEditing){again=true;return}
  if(start!==md||meta.dirty||!meta.seen||!meta.last)write(merged,{...meta,dirty:false,bound:true,seen:true,last:stamp});
  if(start!==md)refresh();
 }catch(e){if(dir===folder){problem="Нет синхронизации: "+e.message+(created?" Проверьте общий файл.":"");if(e.name==="NotAllowedError")permitted=false;}}
 finally{busy=false;renderStatus();if(again){again=false;queueSync(350)}}
}
async function connect(useStored=false){
 need(!busy,"Дождитесь окончания записи.");let dir=useStored&&folder;if(!dir){need(typeof window.showDirectoryPicker==="function","Для общей папки откройте сайт по HTTPS в Chrome или Edge. Можно использовать копии JSON.");dir=await window.showDirectoryPicker({id:"diagnostics-work",mode:"readwrite"})}
 need(await dir.requestPermission({mode:"readwrite"})==="granted","Нет доступа к папке. Разрешите чтение и запись.");
 const same=folder&&await dir.isSameEntry(folder);await remembered("set",dir);write(S,{...meta,bound:true,seen:same?!!meta.seen:false});folder=dir;seen=!!meta.seen;permitted=true;problem="";conflict=null;if($("modal").open)closeDlg();await syncNow()
}
async function disconnect(){need(!busy,"Дождитесь окончания записи.");if(!confirm("Отключить общую папку на этом компьютере? Локальные записи останутся."))return;await remembered("clear");write(S,{...meta,bound:false,seen:false});folder=null;permitted=false;conflict=null;seen=false;problem="";closeDlg();refresh()}
async function saveCopy(){
 need(!locked,"Хранилище не удалось прочитать. Не сохраняйте пустую копию поверх прежней.");const original=digest(S),payload=enc({...S,savedAt:iso()}),name="Журнал-диагностики-"+fileStamp()+".json";let acknowledged=false;
 if(window.showSaveFilePicker){const f=await window.showSaveFilePicker({suggestedName:name,types:[{description:"Копия журнала",accept:{"application/json":[".json"]}}]});if(folder){try{const shared=await folder.getFileHandle(MAIN);need(!(await shared.isSameEntry(f)),"Выберите имя копии, отличное от основного общего файла.")}catch(e){if(e.name!=="NotFoundError")throw e}}await safeFile(f,payload);acknowledged=true}
 else{downloadBlob(new Blob([payload],{type:"application/json"}),name);acknowledged=confirm("Проверьте загрузки. Файл копии уже сохранён?")}
 if(acknowledged&&!meta.bound&&digest(S)===original)write(S,{...meta,dirty:false,last:iso()});renderStatus();if(acknowledged)note("Копия журнала сохранена.")
}
function openCopy(){const f=document.createElement("input");f.type="file";f.accept=".json,application/json";f.onchange=async()=>{try{if(!f.files.length)return;const data=decode(JSON.parse(await f.files[0].text()));review(data,false,f.files[0].name)}catch(e){err("Не удалось открыть копию: "+e.message)}};f.click()}
const kindNames={centers:"Сервисный центр",requests:"Обращение",trips:"Рейс"};
function shortEntry(e){const v=e.row;return e.dead?"Удалено · "+v.label:v.name||v.ticket&&("№ "+v.ticket+" · "+v.item)||("Рейс № "+v.no+" · "+fmt(v.date))}
function review(incoming,shared=false,filename=MAIN){
 const p=planMerge(S,incoming);dialog(shared?"Сверить изменения":"Открыть копию журнала",area(h(filename)+"<br>Новых записей: "+p.adds+" · обновлений: "+p.edits+" · удалений: "+p.deletes+" · расхождений: "+p.conflicts.length)+'<p>Независимые записи с обоих компьютеров будут объединены.</p>'+p.conflicts.map((c,i)=>'<fieldset><legend>'+h(kindNames[c.k])+'</legend>'+sel("choice"+i,"Какую версию оставить?",[["","Выберите версию"],["a","Этот компьютер: "+shortEntry(c.a)],["b","Файл: "+shortEntry(c.b)]],"",true)+'<details><summary>Сравнить содержимое</summary>'+[["Этот компьютер",c.a],["В файле",c.b]].map(([label,e])=>'<p>'+label+'</p><pre>'+h(JSON.stringify(withoutMeta(e.row),null,2))+'</pre>').join("")+'</details></fieldset>').join(""),{save:shared?"Сохранить решение":"Объединить записи",submit:v=>{const merged=combine(S,incoming,p.conflicts.map((_,i)=>v["choice"+i]));write(merged,{...meta,dirty:meta.bound||digest(merged)!==digest(incoming),last:iso()});if(shared)conflict=null;problem="";refresh();queueSync();note("Данные объединены.")}})
}

function gotoTab(id){tab=id;term="";$("search").value="";refresh()}
function goTrip(t){tripDay=t.date;gotoTab("trips");$("trip-"+t.id)?.scrollIntoView({block:"nearest"})}
function legacyDateForm(){const ids=[...document.querySelectorAll(".legacyCheck:checked")].map(x=>x.value);need(ids.length,"Сначала отметьте записи из одного рейса.");dialog("Восстановить дату рейса",'<p>Записей в бумажном листе: '+ids.length+'</p><div class="grid">'+field("date","Дата фактического выезда",tripDay,optDay)+sel("kind","Что делал водитель",[["drop","Отвёз в СЦ"],["pick","Забрал из СЦ"]],"drop")+'</div>',{save:"Добавить в историю",submit:v=>{change(s=>{const t=newTrip(s,v.date);t.status="done";t.doneAt=iso();for(const id of ids){const r=req(s,id);need(r?.legacy==="done","Выберите старые выполненные записи.");t.lines.push({rid:id,kind:v.kind,snap:snapshot(s,id),outcome:"ok",reason:""});r.legacy=""}});note("Дата выезда сохранена.")}})}
async function act(a,id,aux){const r=req(S,id),t=trip(S,id);
 switch(a){
 case "tab":gotoTab(id);break;
 case "filter":filter=id;refresh();break;
 case "callfilter":callFilter=id;refresh();break;
 case "newCase":gotoTab("cases");$("qTicket").focus();$("compose").scrollIntoView({block:"start"});break;
 case "deskCenter":centerForm("",cid=>{if(cid){$("qCenter").value=cid;deskDirty=true;deskWatch();syncGuard()}});break;
 case "deskSuggest":$("qCenter").value=id;deskDirty=true;deskWatch();syncGuard();break;
 case "open":caseID=id;tab="card";refresh();break;
 case "editCase":requestForm(id);break;
 case "deleteCase":if(confirm("Удалить это обращение?")){change(s=>{const x=req(s,id);need(x&&!x.calls.length&&!x.documents.length&&!s.trips.some(t=>t.lines.some(l=>l.rid===id)),"Для обращения уже есть связанные записи.");s.requests=s.requests.filter(v=>v.id!==id)});gotoTab("cases")}break;
 case "toggleClosed":change(s=>{const x=req(s,id);need(!ongoing(s,id).length&&!pendingCall(x),"Сначала завершите назначенные действия.");x.closed=!x.closed;x.events.push({day:today(),text:x.closed?"Обращение закрыто":"Возвращено в работу"})});break;
 case "center":centerForm(id);break;
 case "inlineCenter":inlineAction();break;
 case "suggest":$("editor").elements.centerId.value=id;dialogDirty=true;syncGuard();updateRequestHints();break;
 case "testMap":{const c=formObj($("editor"));need(c.address||validYandexLink(c.map),"Сначала введите адрес.");need(!c.map||validYandexLink(c.map),"Проверьте ссылку на Яндекс Карты.");window.open(mapUrl(c),"_blank","noopener,noreferrer");break}
 case "deleteCenter":need(!S.requests.some(r=>r.centerId===id),"СЦ используется в обращениях. Его можно изменить.");if(confirm("Удалить адрес этого СЦ?"))change(s=>{s.centers=s.centers.filter(c=>c.id!==id)});break;
 case "goTrip":goTrip(t);break;
 case "newTrip":dialog("Новый рейс",'<div class="grid">'+field("date","Дата выезда",tripDay,optDay)+'</div>',{save:"Создать рейс",submit:v=>{let n;change(s=>{n=newTrip(s,v.date)});goTrip(n)}});break;
 case "addToTrip":addToTrip(t);break;
 case "plan":planForm(id,aux);break;
 case "unplan":change(s=>{const x=trip(s,id);need(x.status==="draft","Рейс уже передан водителю.");x.lines=x.lines.filter(l=>l.rid!==aux)});break;
 case "dispatch":dialog("Передать лист водителю?",'<p>Рейс № '+t.no+' от '+fmt(t.date)+'. В листе '+t.lines.length+' инструмент(ов). Для истории будут сохранены адреса и комментарии на момент передачи.</p>',{save:"Передать водителю",submit:()=>{change(s=>sendTrip(s,id));note("Лист передан водителю.")}});break;
 case "finish":finishForm(t);break;
 case "deleteTrip":if(confirm("Удалить пустой рейс?"))change(s=>{const x=trip(s,id);need(x.status==="draft"&&!x.lines.length);s.trips=s.trips.filter(t=>t.id!==id)});break;
 case "printTrip":driverPrint(t);break;
 case "xlsx":downloadBlob(zipFiles(workbookFiles(t)),"Лист-водителя-"+t.date+"-рейс-"+t.no+".xlsx");break;
 case "legacySelected":dialog("Перенести прежний лист",'<p>В рейс на '+fmt(tripDay)+' войдут инструменты, отмеченные в старом листе водителя.</p>'+sel("kind","Направление",[["drop","Едем в СЦ"],["pick","Забираем из СЦ"]],"drop"),{save:"Назначить рейс",submit:v=>change(s=>{const ids=s.requests.filter(r=>r.legacy==="selected"&&!ongoing(s,r.id).length);need(ids.length);const t=newTrip(s,tripDay);for(const r of ids)addLine(s,r.id,v.kind,tripDay,t.id)})});break;
 case "legacyDate":legacyDateForm();break;
 case "call":callForm(r);break;
 case "result":resultForm(r);break;
 case "statement":statementForm(r);break;
 case "hand":printStatement({...storeDefault,item:r.item},true);break;
 case "printDoc":printStatement(r.documents[Number(aux)]);break;
 case "cancel":leaveDlg();break;
 case "storage":storageView();break;
 case "connect":await connect();break;
 case "grant":await connect(true);break;
 case "sync":if($("modal").open)closeDlg();await syncNow();break;
 case "disconnect":await disconnect();break;
 case "export":await saveCopy();break;
 case "import":openCopy();break;
 case "review":if(conflict)review(conflict.s,true);break;
 case "restore":restoreFromBackup();break;
 }
}
function restoreFromBackup(){const input=document.createElement("input");input.type="file";input.accept=".json,application/json";input.onchange=async()=>{try{if(!input.files.length)return;const src=decode(JSON.parse(await input.files[0].text()));dialog("Вернуть журнал к выбранной копии",area(h(input.files[0].name)+"<br>Обращений в копии: "+src.requests.length+" · рейсов: "+src.trips.length,"warning")+'<p>Все текущие записи будут приведены к состоянию из этой копии. Более поздние изменения будут сняты. Сначала сохраните текущую версию кнопкой «Сохранить копию».</p>'+field("word","Для подтверждения введите ВЕРНУТЬ","",{req:true}),{save:"Восстановить журнал",submit:v=>{need(v.word==="ВЕРНУТЬ","Введите слово подтверждения.");const n=newStore();for(const k of kinds){const left=indexed(S,k),right=indexed(src,k);for(const id of new Set([...left.keys(),...right.keys()])){const old=left.get(id),target=right.get(id),clock=advance(joinClock(old?.row._v||{},target?.row._v||{}));if(target&&!target.dead)n[k].push({...target.row,_v:clock});else n.removed[k].push({id,label:target?.row.label||old?.row.label||shortEntry(old),_v:clock})}}write(decode(n),{...meta,dirty:true});conflict=null;refresh();queueSync();note("Журнал восстановлен из копии.")}})}catch(e){err(e.message)}};input.click()}

let deskDirty=false;
function palette(name,remember=false){const mode=name==="green"?"green":"blue";document.documentElement.dataset.theme=mode;document.querySelector('meta[name="theme-color"]').content=mode==="green"?"#103544":"#10243f";if(remember){try{localStorage.setItem(KEY+".theme",mode)}catch{note("Оформление выбрано для этой вкладки.")}}}
function updateOptions(el,list){const old=el.value,key=JSON.stringify(list);if(el.dataset.key===key)return;el.innerHTML=options(list,old);el.dataset.key=key;if(list.some(([k])=>k===old))el.value=old}
function deskWatch(){const f=$("deskForm"),v=f.elements.step.value,d=$("qDate");
 $("qWhen").hidden=v==="none";d.required=v!=="none";$("qTripWrap").hidden=v==="none"||v==="call";$("qDateLabel").textContent=v==="call"?"Дата звонка":"Дата выезда";
 $("deskSubmit").textContent=v==="call"?"Назначить звонок":v==="none"?"Создать обращение":"Добавить в маршрут";
 updateOptions($("qCenter"),centersOpts());updateOptions($("qTrip"),tripOpts(d.value));
 const item=$("qItem").value.toLowerCase(),cs=S.centers.filter(c=>c.rules.split(/[,;\n]/).some(v=>v.trim()&&item.includes(v.trim().toLowerCase())));
 $("deskHints").innerHTML=cs.length?"Подходит: "+cs.map(c=>b(c.name,"deskSuggest",c.id,"inline-action")).join(" · "):"";
 const c=center(S,$("qCenter").value);$("deskMap").hidden=!c;$("deskMap").innerHTML=c?'<strong>'+h(c.name)+'</strong><span>'+h(c.address)+'</span><span>'+h(c.phone)+'</span><a href="'+h(mapUrl(c))+'" target="_blank" rel="noopener noreferrer">Показать в Яндекс Картах</a>':""
}
function classicView(){const copy={
 cases:["Маршрут","Пункты доставки","Обращения сгруппированы по сервисным центрам.","Доставки на диагностику и задания водителю."],
 trips:["Лист водителя","Рейсы на выбранную дату","Назначьте инструменты, передайте лист и отметьте результат.","Планирование выездов и передача листа водителю."],
 calls:["Задачи","Звонки в сервисные центры","Запишите ответ СЦ и назначьте забор или повторный звонок.","План звонков и ответы сервисных центров."],
 history:["Архив","История рейсов","Откройте дату выезда, чтобы увидеть адреса, инструменты и результаты.","Выполненные поездки по датам выезда."],
 centers:["Справочник","Адреса сервисных центров","Контакты и адреса подставляются в новые обращения.","Контакты, время приёмки и специализация сервисных центров."],
 card:["Обращение","Карточка инструмента","Действия, даты и заявления по обращению.","Доставка, звонки и возврат инструмента."]}[tab];
 $("queueKicker").textContent=copy[0];$("queueTitle").textContent=copy[1];$("queueNote").textContent=copy[2];$("pageNote").textContent=copy[3];
 const active=S.trips.filter(t=>t.status!=="done");$("countOpen").textContent=S.requests.filter(r=>!r.closed).length;$("countTools").textContent=active.reduce((n,t)=>n+t.lines.length,0);$("countStops").textContent=active.reduce((n,t)=>n+groups(S,t).length,0);$("countCenters").textContent=S.centers.length;
 $("compose").hidden=tab!=="cases";$("work").classList.toggle("solo",tab!=="cases");if(tab==="cases")deskWatch();
}
// Event wiring.
function boot(){
 loadLocal();let theme="blue";try{theme=localStorage.getItem(KEY+".theme")||theme}catch{}palette(theme);
 $("qDate").value=today();refresh();const desk=$("deskForm");
 function deskInput(){deskDirty=true;syncGuard();deskWatch();$("deskError").textContent=""}
 desk.addEventListener("input",deskInput);desk.addEventListener("change",deskInput);
 desk.addEventListener("submit",e=>{e.preventDefault();try{if(saveRequest(formObj(desk))===false)return;desk.reset();$("qDate").value=today();deskDirty=false;deskWatch();syncGuard()}catch(ex){$("deskError").textContent=ex.message}});
 document.addEventListener("change",e=>{if(e.target.id==="classicTheme")palette(e.target.value,true)});
 document.addEventListener("click",e=>{const btn=e.target.closest("[data-act]");if(!btn)return;act(btn.dataset.act,btn.dataset.id||"",btn.dataset.aux||"").catch(e=>{if(e.name!=="AbortError")err(e.message)})});
 $("close").onclick=leaveDlg;$("modal").addEventListener("cancel",e=>{e.preventDefault();leaveDlg()});
 $("modal").addEventListener("close",()=>{if(!$("modal").open)queueSync()});
 $("search").addEventListener("input",e=>{term=e.target.value.trim();tab="cases";refresh()});
 document.addEventListener("input",e=>{if(e.target.closest("#editor")){dialogDirty=true;syncGuard();if(e.target.name==="item")updateRequestHints()}});
 document.addEventListener("change",e=>{const n=e.target,f=n.form;if(n.id==="tripDay"&&isDay(n.value)){tripDay=n.value;refresh()}if(f?.id!=="editor")return;dialogDirty=true;syncGuard();
 if(["step","centerId","stepDate"].includes(n.name))updateRequestHints();
 if(["stepDate","tripDate"].includes(n.name)&&f.elements.tripId)f.elements.tripId.innerHTML=options(tripOpts(n.value),"auto");
 if(n.name==="out"){f.elements.tripDate.parentElement.firstChild.textContent=n.value==="ready"?"Дата забора":"Дата следующего звонка";f.elements.tripId.parentElement.hidden=n.value!=="ready"}
 });
 document.addEventListener("submit",e=>{if(e.target.id!=="editor")return;e.preventDefault();try{need(submitFn&&digest(S)===baselineLocal,"Данные изменились. Закройте форму и откройте её снова.");if(submitFn(formObj(e.target),e.target)!==false)closeDlg()}catch(ex){err(ex.message)}});
 window.addEventListener("storage",e=>{if(e.key!==KEY&&e.key!==null)return;if(dialogEditing||deskDirty){problem="Журнал изменился в другой вкладке. Закройте форму и обновите страницу.";renderStatus();return}location.reload()});
 function wake(){if(!document.hidden)queueSync(50)}window.addEventListener("focus",wake);document.addEventListener("visibilitychange",wake);setInterval(wake,12000);
 if(meta.bound&&!locked){(async()=>{try{folder=await remembered("get");if(folder){permitted=await folder.queryPermission({mode:"readwrite"})==="granted";seen=!!meta.seen;if(permitted)await syncNow()}}catch(e){problem="Не удалось открыть общую папку: "+e.message}renderStatus()})()}
}
boot();
