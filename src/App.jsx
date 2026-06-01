@@ -45,140 +45,169 @@ async function getFB() {
  if(fbApp) return {db:fbDb,storage:fbStorage};
  const {initializeApp}   = await import("https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js");
  const {getDatabase,ref,set,onValue,push,get} = await import("https://www.gstatic.com/firebasejs/10.12.0/firebase-database.js");
  const {getStorage,ref:sRef,uploadBytes,getDownloadURL,listAll} = await import("https://www.gstatic.com/firebasejs/10.12.0/firebase-storage.js");
  fbApp = initializeApp(FB_CONFIG);
  fbDb  = {db:getDatabase(fbApp),ref,set,onValue,push,get};
  fbStorage = {storage:getStorage(fbApp),sRef,uploadBytes,getDownloadURL,listAll};
  return {db:fbDb,storage:fbStorage};
}

// ╔══════════════════════════════════════════════════════╗
// ║           FISCAL YEAR  Abril → Marzo                 ║
// ╚══════════════════════════════════════════════════════╝
const FM=["Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic","Ene","Feb","Mar"];
const SM=["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
const getFY =(m,y)=>m>=4?y:y-1;
const getFMI=(m)  =>m>=4?m-4:m+8;
const getFQ =(m)  =>Math.floor(getFMI(m)/3)+1;
const fyLbl =(fy) =>`FY ${fy}/${String(fy+1).slice(2)}`;
const QNames=["Q1 Abr–Jun","Q2 Jul–Sep","Q3 Oct–Dic","Q4 Ene–Mar"];

// ╔══════════════════════════════════════════════════════╗
// ║           COLUMN DETECTION                           ║
// ╚══════════════════════════════════════════════════════╝
const CA={
  ing: ["ingreso tot","ingresos"],
  cst: ["costo tot","costos material"],
  cli: ["cliente"],
  clinm:["descripción cliente","descripcion cliente"],
  f2:  ["jerarquia de producto 2","jerarquía de producto 2"],
  f1:  ["jerarquia de producto","jerarquía de producto"],
  art: ["articulo","artículo"],
  suc: ["sucursal"],
  sec: ["descrip. gr.clie.","descrip.gr.clie."],
  rep: ["representante de ventas"],
  vol: ["vol.ventas","vol ventas"],
  abc: ["indicador abc"],
  fecha:["fecha factura"],
  per: ["período/año","periodo/año"],
  ing: ["ingreso tot","ingresos","importe venta","venta neta","valor neto","net sales","revenue"],
  cst: ["costo tot","costos material","costo total","coste total","cost","costos"],
  cli: ["cliente","cod cliente","codigo cliente","código cliente","customer","sold to"],
  clinm:["descripción cliente","descripcion cliente","nombre cliente","razon social","razón social","customer name"],
  f2:  ["jerarquia de producto 2","jerarquía de producto 2","subfamilia","sub familia","product hierarchy 2"],
  f1:  ["jerarquia de producto","jerarquía de producto","familia","product hierarchy"],
  art: ["articulo","artículo","material","codigo material","código material","item"],
  suc: ["sucursal","centro","branch"],
  sec: ["descrip. gr.clie.","descrip.gr.clie.","grupo cliente","sector","industria"],
  rep: ["representante de ventas","vendedor","ejecutivo","sales representative"],
  vol: ["vol.ventas","vol ventas","volumen ventas","cantidad","qty","quantity"],
  abc: ["indicador abc","abc"],
  fecha:["fecha factura","fecha de factura","billing date","invoice date"],
  per: ["período/año","periodo/año","periodo ano","periodo año","period"],
};
const normHeader=(v)=>String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
// Exact match override — always wins over dc() fuzzy search
function dcExact(headers,patterns){
  const normalized=headers.map(h=>normHeader(h));
  for(const p of patterns){
    const found=headers.find(h=>h.trim().toLowerCase()===p.toLowerCase());
    if(found)return found;
    const ix=normalized.findIndex(h=>h===normHeader(p));
    if(ix!==-1)return headers[ix];
  }
  return null;
}
function dc(headers,f){const lo=headers.map(h=>String(h??"").toLowerCase().trim());const al=CA[f]||[];const i=lo.findIndex(h=>al.some(a=>h.includes(a)));return i!==-1?headers[i]:null;}
function dc(headers,f){const lo=headers.map(h=>normHeader(h));const al=(CA[f]||[]).map(normHeader);const i=lo.findIndex(h=>al.some(a=>h===a||h.includes(a)));return i!==-1?headers[i]:null;}
function previewRows(ws,rowCount=25){
  const ref=ws["!ref"];
  if(!ref)return[];
  const range=XLSX.utils.decode_range(ref);
  range.e.r=Math.min(range.e.r,range.s.r+rowCount-1);
  return XLSX.utils.sheet_to_json(ws,{header:1,defval:null,blankrows:false,range:XLSX.utils.encode_range(range)});
}
function detectHeaderRow(rawRows){
  let best={idx:0,score:-1};
  for(let i=0;i<rawRows.length;i++){
    const headers=(rawRows[i]||[]).map(v=>String(v??"").trim());
    if(!headers.some(Boolean))continue;
    const score=Object.keys(CA).reduce((sum,f)=>sum+(dc(headers,f)?1:0),0);
    if(score>best.score)best={idx:i,score};
  }
  return best;
}
function sucFromName(n){const u=n.toUpperCase();if(u.includes("K27")||u.includes("AREQUIPA"))return"K27";if(u.includes("K11")||u.includes("HUANCAYO"))return"K11";if(u.includes("K23")||u.includes("TACNA"))return"K23";if(u.includes("K01")||u.includes("LIMA")||u.includes("CALLAO"))return"K01";return"K??"}
function periodFromName(n){const u=n.toUpperCase().replace(/[_\-\.]/g," ");const m=u.match(/(\d{4})\s+(\d{1,2})/);if(m)return{y:parseInt(m[1]),mo:parseInt(m[2])};const mi=SM.findIndex(s=>u.includes(s));if(mi!==-1){const ym=u.match(/(\d{4})/);if(ym)return{y:parseInt(ym[1]),mo:mi+1}}return null}

async function parseXLSX(file){
  const ab=await file.arrayBuffer();
  const wb=XLSX.read(ab,{type:"array",cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{defval:null});
  const label=file.name.replace(/\.[^.]+$/,"");
  const sfn=sucFromName(label);
  const pfn=periodFromName(label);

  // Some SAP/Excel exports include title/filter rows before the real header.
  // Detect the row with the most known business columns, then parse from there.
  const rawRows=previewRows(ws);
  if(!rawRows.length)return null;
  const headerInfo=detectHeaderRow(rawRows);
  const rows=XLSX.utils.sheet_to_json(ws,{defval:null,range:headerInfo.idx});
  if(!rows.length)return null;
  const headers=Object.keys(rows[0]);
  const c={};for(const f of Object.keys(CA))c[f]=dc(headers,f);

  // ── Exact-match overrides — these columns must match precisely ────────────
  c.ing  = dcExact(headers,["Ingreso TOT","INGRESO TOT"]) || c.ing;
  c.cst  = dcExact(headers,["Costo TOT","COSTO TOT"])    || c.cst;
  c.sec  = dcExact(headers,["Descrip. Gr.Clie.","Descrip.Gr.Clie.","DESCRIP. GR.CLIE."]) || c.sec;
  c.rep  = dcExact(headers,["Representante de ventas","Representante De Ventas"]) || c.rep;
  c.clinm= dcExact(headers,["Descripción Cliente","Descripcion Cliente"]) || c.clinm;
  c.suc  = dcExact(headers,["Sucursal","SUCURSAL"]) || c.suc;
  c.cli  = dcExact(headers,["Cliente","CLIENTE","Cod. Cliente","Código Cliente","Codigo Cliente"]) || c.cli;
  c.clinm= dcExact(headers,["Descripción Cliente","Descripcion Cliente","Nombre Cliente","Razón Social","Razon Social"]) || c.clinm;
  c.suc  = dcExact(headers,["Sucursal","SUCURSAL","Centro"]) || c.suc;
  c.f2   = dcExact(headers,["Jerarquia de producto 2","Jerarquía de producto 2"]) || c.f2;
  c.f1   = dcExact(headers,["Jerarquia de producto","Jerarquía de producto"]) || c.f1;
  c.per  = dcExact(headers,["Período/Año","Periodo/Año","Período/año"]) || c.per;
  c.fecha= dcExact(headers,["Fecha factura","Fecha Factura"]) || c.fecha;
  // Fecha contab fallback for period detection
  const cFechaContab = dcExact(headers,["Fecha de contabilización","Fecha de Contabilización"]) ||
    headers.find(h=>h.toLowerCase().includes("contabiliz"));
    headers.find(h=>normHeader(h).includes("contabiliz"));

  const recs=[];
  rows.forEach(r=>{
    const ing=pNum(r[c.ing]);
    const cst=pNum(r[c.cst]);
    // Skip pure rebate/zero rows that have no client and no ingreso
    if(ing===0 && cst===0 && !r[c.cli]) return;

    // Period detection — priority: Período/Año column (format 2026004) → fecha factura → fecha contab → filename
    let mo=pfn?.mo??4, yr=pfn?.y??2026;

    // 1. Try Período/Año (value like 2026004 = year 2026 month 04)
    if(r[c.per]){
      const s=String(Math.round(parseFloat(r[c.per]))||0).replace(/\D/g,"");
      if(s.length>=5){yr=parseInt(s.slice(0,4));mo=parseInt(s.slice(4))||mo;}
    }
    // 2. Try Fecha factura
    const fv=r[c.fecha];
    if(fv){const d=fv instanceof Date?fv:new Date(fv);if(!isNaN(d)&&d.getFullYear()>2000){mo=d.getMonth()+1;yr=d.getFullYear();}}
    // 3. Fallback to Fecha de contabilización
    else if(cFechaContab&&r[cFechaContab]){
      const fc=r[cFechaContab];
      const d=fc instanceof Date?fc:new Date(fc);
      if(!isNaN(d)&&d.getFullYear()>2000){mo=d.getMonth()+1;yr=d.getFullYear();}
    }

    const fy=getFY(mo,yr),fmi=getFMI(mo),fq=getFQ(mo);
    recs.push({ing,cst,mgn:ing-cst,vol:pNum(r[c.vol]),
      cli:r[c.cli]?String(r[c.cli]).trim():"0",
      clinm:r[c.clinm]?String(r[c.clinm]).trim():"Sin nombre",
      f2:r[c.f2]?String(r[c.f2]).trim():"Sin familia",
      f1:r[c.f1]?String(r[c.f1]).trim():"Sin familia",
      art:r[c.art]?String(r[c.art]).trim():"?",
      suc:r[c.suc]?String(r[c.suc]).trim():sfn,
      sec:(()=>{const v=r[c.sec]?String(r[c.sec]).trim():"";return(v&&isNaN(v)&&v.length>1)?v:"Otros";})(),
      rep:r[c.rep]?String(r[c.rep]).trim():"0",
      abc:r[c.abc]?String(r[c.abc]).trim():"N/A",
      mo,yr,fy,fmi,fq,fml:FM[fmi],fyl:fyLbl(fy),file:label});
  });
  return{label,suc:sfn,recs,cols:c,headers:headers.slice(0,20)};
  return{label,suc:sfn,recs,cols:c,headers:headers.slice(0,20),headerRow:headerInfo.idx+1};
}

// Robust number parser — handles locale formats, strings, nulls
function pNum(v){
  if(v===null||v===undefined||v==="")return 0;
  if(typeof v==="number")return isFinite(v)?v:0;
  const s=String(v).trim().replace(/[^\d\.,\-]/g,"");
  // Handle European format 1.234,56 → 1234.56
  if(s.match(/\d{1,3}(\.\d{3})+,\d+/))return parseFloat(s.replace(/\./g,"").replace(",","."));
  // Handle 1,234.56
  if(s.match(/\d{1,3}(,\d{3})+\.\d+/))return parseFloat(s.replace(/,/g,""));
  // Handle simple comma decimal 1234,56
  if(s.match(/^\d+,\d+$/))return parseFloat(s.replace(",","."));
  return parseFloat(s)||0;
}
const SEED={"meta":{"sucursal":"K27","sucursalNm":"Arequipa DCP","periodo":"Abril 2026","mes":4,"year":2026,"fyLabel":"FY 2026/27","ing":56030.15,"cst":42780.46,"mgn":13249.69,"lineas":409,"clientes":30,"articulos":197,"vol":1006},"clients":[{"id":6013813,"nm":"INTERNATIONAL CAMIONES DEL PERU","sector":"Automotriz","ing":17000.25,"cst":12809.5,"lineas":96,"vol":220,"mgn":4190.75,"mgnP":24.7,"pct":30.36,"cum":30.36,"seg":"TOP"},{"id":6027645,"nm":"COMERCIALIZADORA MIRIANDI E.I.R.L.","sector":"Canteras","ing":14305.91,"cst":8714.06,"lineas":8,"vol":33,"mgn":5591.85,"mgnP":39.1,"pct":25.55,"cum":55.91,"seg":"TOP"},{"id":6003804,"nm":"SOCIEDAD MINERA CERRO VERDE S.A.A.","sector":"Gran Minería","ing":7481.69,"cst":4766.36,"lineas":9,"vol":5,"mgn":2715.33,"mgnP":36.3,"pct":13.36,"cum":69.27,"seg":"MED"},{"id":6003931,"nm":"SOUTHERN PERU COPPER CORPORATION","sector":"Gran Minería","ing":2944.82,"cst":1793.4,"lineas":2,"vol":2,"mgn":1151.42,"mgnP":39.1,"pct":5.26,"cum":74.53,"seg":"MED"},{"id":6003863,"nm":"UNIMAQ S.A","sector":"Automotriz","ing":2839.85,"cst":2845.27,"lineas":27,"vol":69,"mgn":-5.42,"mgnP":-0.2,"pct":5.07,"cum":79.6,"seg":"MED"},{"id":6008872,"nm":"LADRILLERA EL DIAMANTE S.A.C.","sector":"Construccion","ing":1920.58,"cst":1827.88,"lineas":18,"vol":136,"mgn":92.7,"mgnP":4.8,"pct":3.43,"cum":83.03,"seg":"MED"},{"id":6002240,"nm":"REPUESTOS DAVID DIESEL E.I.R.L.","sector":"Automotriz","ing":1781.32,"cst":1770.8,"lineas":11,"vol":76,"mgn":10.52,"mgnP":0.6,"pct":3.18,"cum":86.21,"seg":"MED"},{"id":6003685,"nm":"TRACTO CAMIONES USA S.A.C.","sector":"Automotriz","ing":1359.38,"cst":1118.86,"lineas":46,"vol":142,"mgn":240.52,"mgnP":17.7,"pct":2.43,"cum":88.64,"seg":"MED"},{"id":6016191,"nm":"LUBRICANTES PIEYYRA E.I.R.L.","sector":"Automotriz","ing":1285.95,"cst":1198.53,"lineas":40,"vol":101,"mgn":87.42,"mgnP":6.8,"pct":2.3,"cum":90.94,"seg":"LOW"},{"id":6002717,"nm":"PERURAIL S.A.","sector":"Automotriz","ing":1196.88,"cst":1135.56,"lineas":6,"vol":52,"mgn":61.32,"mgnP":5.1,"pct":2.14,"cum":93.08,"seg":"LOW"},{"id":6023634,"nm":"FUNDO HNOS DOMINGUEZ S.A.C.","sector":"Agropecuario","ing":838.86,"cst":1049.95,"lineas":5,"vol":4,"mgn":-211.09,"mgnP":-25.2,"pct":1.5,"cum":94.58,"seg":"LOW"},{"id":6002578,"nm":"TRANSPORTES HAGEMSA S.A.C.","sector":"Automotriz","ing":763.84,"cst":777.4,"lineas":5,"vol":22,"mgn":-13.56,"mgnP":-1.8,"pct":1.36,"cum":95.94,"seg":"LOW"},{"id":6021821,"nm":"CAMFRA E.I.R.L.","sector":"Automotriz","ing":468.36,"cst":471.72,"lineas":3,"vol":13,"mgn":-3.36,"mgnP":-0.7,"pct":0.84,"cum":96.78,"seg":"LOW"},{"id":6003898,"nm":"EPIROC PERU SOCIEDAD ANONIMA","sector":"Min. Sub & Loc. Dril","ing":391.25,"cst":274.55,"lineas":3,"vol":5,"mgn":116.7,"mgnP":29.8,"pct":0.7,"cum":97.48,"seg":"LOW"},{"id":6022633,"nm":"AR MAQUINARIAS DEL PERU SOCIEDAD","sector":"Min. Sub & Loc. Dril","ing":321.46,"cst":291.91,"lineas":1,"vol":50,"mgn":29.55,"mgnP":9.2,"pct":0.57,"cum":98.05,"seg":"LOW"},{"id":6026321,"nm":"MULTISERVICIOS Y OPERACIONES JBN","sector":"Comercio","ing":286.0,"cst":196.83,"lineas":3,"vol":3,"mgn":89.17,"mgnP":31.2,"pct":0.51,"cum":98.56,"seg":"LOW"},{"id":6026883,"nm":"HAZVOZ E.I.R.L.","sector":"Utilities / Telecom","ing":201.97,"cst":183.42,"lineas":15,"vol":42,"mgn":18.55,"mgnP":9.2,"pct":0.36,"cum":98.92,"seg":"LOW"},{"id":6013017,"nm":"CONSORCIO EMPRESA DEPURADORA DE","sector":"Utilities / Telecom","ing":191.1,"cst":64.58,"lineas":2,"vol":2,"mgn":126.52,"mgnP":66.2,"pct":0.34,"cum":99.26,"seg":"LOW"},{"id":6000430,"nm":"RD RENTAL S.A.C","sector":"Min. Sub & Loc. Dril","ing":190.74,"cst":174.68,"lineas":9,"vol":21,"mgn":16.06,"mgnP":8.4,"pct":0.34,"cum":99.6,"seg":"LOW"},{"id":6020591,"nm":"AGROGANADO LA ABRIL E.I.R.L.","sector":"Agropecuario","ing":85.18,"cst":57.02,"lineas":12,"vol":79,"mgn":28.16,"mgnP":33.1,"pct":0.15,"cum":99.75,"seg":"LOW"},{"id":6026791,"nm":"CONSTRUCTORA E INMOBILIARIA IBER","sector":"Construccion","ing":59.31,"cst":52.72,"lineas":4,"vol":25,"mgn":6.59,"mgnP":11.1,"pct":0.11,"cum":99.86,"seg":"LOW"},{"id":6014498,"nm":"FERTICA S.A.C.","sector":"Utilities / Telecom","ing":52.34,"cst":66.17,"lineas":2,"vol":2,"mgn":-13.83,"mgnP":-26.4,"pct":0.09,"cum":99.95,"seg":"LOW"},{"id":6025555,"nm":"CONTRATISTAS GENERALES ARRIOLA","sector":"Construccion","ing":42.13,"cst":25.84,"lineas":4,"vol":23,"mgn":16.29,"mgnP":38.7,"pct":0.08,"cum":100.03,"seg":"LOW"},{"id":6017393,"nm":"CONSULTORA DE SISTEMAS INTEGRADO","sector":"Utilities / Telecom","ing":24.71,"cst":18.97,"lineas":2,"vol":2,"mgn":5.74,"mgnP":23.2,"pct":0.04,"cum":100.07,"seg":"LOW"},{"id":6026900,"nm":"H.S. ELECTROMECANICA S.A.C.","sector":"Min. Sub & Loc. Dril","ing":19.54,"cst":11.64,"lineas":2,"vol":2,"mgn":7.9,"mgnP":40.4,"pct":0.03,"cum":100.1,"seg":"LOW"},{"id":6003218,"nm":"PORTUARIA PERUANO SUIZA S.A.C","sector":"Portuario","ing":14.7,"cst":8.72,"lineas":2,"vol":2,"mgn":5.98,"mgnP":40.7,"pct":0.03,"cum":100.13,"seg":"LOW"}],"familia":[{"f1":"Filtración","f2":"Filtración Aftermark","ing":13259.8,"cst":13886.24,"vol":587,"mgn":-626.44,"mgnP":-4.7},{"f1":"Grupo Electrógeno","f2":"Generador Diesel","ing":13127.91,"cst":8714.06,"vol":1,"mgn":4413.85,"mgnP":33.6},{"f1":"Partes Motores","f2":"Partes Mot. HD","ing":12800.72,"cst":9698.57,"vol":214,"mgn":3102.15,"mgnP":24.2},{"f1":"Partes Motores","f2":"Partes Mot.HHP","ing":9226.64,"cst":5443.47,"vol":8,"mgn":3783.17,"mgnP":41.0},{"f1":"Partes Recon","f2":"Partes Recon HD","ing":4556.09,"cst":3367.57,"vol":5,"mgn":1188.52,"mgnP":26.1},{"f1":"Servicio CS","f2":"Servicio CS","ing":1178.0,"cst":0.0,"vol":82,"mgn":1178.0,"mgnP":100.0},{"f1":"Partes Motores","f2":"Partes Mot. MR","ing":658.24,"cst":478.59,"vol":30,"mgn":179.65,"mgnP":27.3},{"f1":"Lubricantes","f2":"Lubricantes","ing":391.25,"cst":274.55,"vol":5,"mgn":116.7,"mgnP":29.8},{"f1":"Partes Motores","f2":"Partes Compra Mot. B","ing":334.03,"cst":168.87,"vol":14,"mgn":165.16,"mgnP":49.4},{"f1":"Seguridad Industrial","f2":"Paños varios","ing":321.46,"cst":291.91,"vol":50,"mgn":29.55,"mgnP":9.2},{"f1":"Filtración","f2":"Filtración Accesorio","ing":155.83,"cst":148.68,"vol":9,"mgn":7.15,"mgnP":4.6},{"f1":"Filtración","f2":"Filtración Mineria","ing":20.18,"cst":307.95,"vol":1,"mgn":-287.77,"mgnP":-1426.0}],"sector":[{"sec":"Automotriz","ing":25590.01,"lineas":214,"clientes":10},{"sec":"Canteras","ing":14305.91,"lineas":8,"clientes":1},{"sec":"Gran Minería","ing":10426.51,"lineas":27,"clientes":3},{"sec":"Construccion","ing":2022.02,"lineas":72,"clientes":5},{"sec":"Min. Sub & Loc. Dril","ing":1677.2,"lineas":49,"clientes":3},{"sec":"Utilities / Telecom","ing":1578.72,"lineas":17,"clientes":3},{"sec":"Comercio","ing":286.0,"lineas":3,"clientes":1},{"sec":"Agropecuario","ing":85.18,"lineas":12,"clientes":1},{"sec":"Portuario","ing":19.54,"lineas":2,"clientes":1}],"byday":[{"dia":1,"ing":37836.6,"lineas":123},{"dia":2,"ing":6462.83,"lineas":8},{"dia":4,"ing":2854.87,"lineas":24},{"dia":5,"ing":5012.65,"lineas":19},{"dia":6,"ing":3863.2,"lineas":32}],"reps":[{"rep":"2193","ing":23336.44,"clientes":7,"lineas":117},{"rep":"2184","ing":14305.91,"clientes":1,"lineas":6},{"rep":"3622","ing":7481.69,"clientes":1,"lineas":6},{"rep":"3768","ing":4794.19,"clientes":7,"lineas":36},{"rep":"184","ing":2944.82,"clientes":1,"lineas":2},{"rep":"7692","ing":1203.81,"clientes":7,"lineas":18},{"rep":"3192","ing":1124.43,"clientes":1,"lineas":12},{"rep":"8626","ing":838.86,"clientes":1,"lineas":5}],"articles":[{"art":"G1C60D6EIC","f1":"Grupo Electrógeno","f2":"Generador Diesel","ing":13127.91,"cst":8714.06,"vol":1,"mgn":4413.85,"mgnP":33.6},{"art":"CM553894900","f1":"Partes Motores","f2":"Partes Mot. HD","ing":4408.32,"cst":2701.28,"vol":1,"mgn":1707.04,"mgnP":38.7},{"art":"LF4054","f1":"Filtración","f2":"Filtración Aftermark","ing":4834.9,"cst":4760.92,"vol":287,"mgn":73.98,"mgnP":1.5},{"art":"CM642951100","f1":"Partes Motores","f2":"Partes Mot. HD","ing":3102.75,"cst":2422.57,"vol":1,"mgn":680.18,"mgnP":21.9},{"art":"LF4054F","f1":"Filtración","f2":"Filtración Aftermark","ing":2989.79,"cst":3050.94,"vol":97,"mgn":-61.15,"mgnP":-2.0},{"art":"AF4501","f1":"Filtración","f2":"Filtración Aftermark","ing":2929.62,"cst":2989.25,"vol":104,"mgn":-59.63,"mgnP":-2.0},{"art":"CM637714700","f1":"Partes Motores","f2":"Partes Mot.HHP","ing":2800.02,"cst":1690.65,"vol":1,"mgn":1109.37,"mgnP":39.6},{"art":"CM341175600RX","f1":"Partes Recon","f2":"Partes Recon HD","ing":2684.72,"cst":2073.38,"vol":4,"mgn":611.34,"mgnP":22.8},{"art":"LF9009","f1":"Filtración","f2":"Filtración Aftermark","ing":1714.61,"cst":1776.68,"vol":39,"mgn":-62.07,"mgnP":-3.6},{"art":"4387765","f1":"Partes Motores","f2":"Partes Mot. HD","ing":1461.7,"cst":1117.82,"vol":1,"mgn":343.88,"mgnP":23.5},{"art":"LF4005F","f1":"Filtración","f2":"Filtración Aftermark","ing":1575.19,"cst":1580.66,"vol":50,"mgn":-5.47,"mgnP":-0.3},{"art":"LF16015","f1":"Filtración","f2":"Filtración Aftermark","ing":1362.94,"cst":1416.64,"vol":32,"mgn":-53.7,"mgnP":-3.9},{"art":"3966323","f1":"Partes Motores","f2":"Partes Mot. HD","ing":1295.12,"cst":990.79,"vol":30,"mgn":304.33,"mgnP":23.5},{"art":"4387766","f1":"Partes Motores","f2":"Partes Mot. HD","ing":1191.5,"cst":911.28,"vol":1,"mgn":280.22,"mgnP":23.5},{"art":"LF3349","f1":"Filtración","f2":"Filtración Aftermark","ing":1027.64,"cst":1056.13,"vol":25,"mgn":-28.49,"mgnP":-2.8}],"abc":[{"ind":"A","ing":16553.93,"lineas":112,"vol":402},{"ind":"B","ing":2446.16,"lineas":32,"vol":87},{"ind":"C","ing":3644.12,"lineas":20,"vol":48},{"ind":"D","ing":234.89,"lineas":4,"vol":6}]};

// ╔══════════════════════════════════════════════════════╗
// ║           FORMATTING & UTILS                         ║
// ╚══════════════════════════════════════════════════════╝
const f$=(n,d=0)=>{if(n==null)return"—";const a=Math.abs(n);const s=a>=1e6?`$${(a/1e6).toFixed(2)}M`:a>=1e3?`$${(a/1e3).toFixed(1)}K`:`$${a.toFixed(d)}`;return n<0?`(${s})`:s};
const fP=(n,d=1)=>n!=null?`${n>0?"+":""}${n.toFixed(d)}%`:"—";
const fN=(n)=>n?.toLocaleString("es-PE")??"—";
const clamp=(v,a,b)=>Math.min(Math.max(v,a),b);
const sortBy=(arr,key,asc=false)=>[...arr].sort((a,b)=>asc?a[key]-b[key]:b[key]-a[key]);
@@ -261,71 +290,81 @@ export default function App(){
  const [pivotSort, setPivotSort]   = useState("desc");

  useEffect(()=>{localStorage.setItem("cummins_c",JSON.stringify(cartera));},[cartera]);
  useEffect(()=>{localStorage.setItem("cummins_r",JSON.stringify(repNames));},[repNames]);

  // Seed records
  const seedRecs = useMemo(()=>{
    const recs=[];
    SEED.clients.forEach(cl=>{
      const chunk=Math.max(1,Math.min(cl.lineas,5));
      for(let i=0;i<chunk;i++){
        recs.push({ing:cl.ing/chunk,cst:cl.cst/chunk,mgn:cl.mgn/chunk,vol:Math.round(cl.vol/chunk),
          cli:String(cl.id),clinm:cl.nm,f2:"—",f1:"—",art:"—",suc:"K27",sec:cl.sector,rep:"0",
          abc:"N/A",mo:4,yr:2026,fy:2026,fmi:0,fq:1,fml:"Abr",fyl:"FY 2026/27",file:"2026_04_K27"});
      }
    });
    return recs;
  },[]);

  useEffect(()=>{if(allRecs.length===0)setAllRecs(seedRecs);},[]);
  useEffect(()=>{if(files.length===0&&allRecs.length>0)setFiles([{label:"2026_04_K27",count:SEED.meta.lineas,suc:"K27"}]);},[allRecs]);

  // — File ingestion — clears seedRecs on first real upload —
  const ingest=useCallback(async fs=>{
    setLoading(true);
    const excelFiles=Array.from(fs||[]).filter(f=>f.name?.match(/\.xlsx?$/i));
    const newRecs=[],newFiles=[];
    for(const f of fs){
      if(!f.name.match(/\.xlsx?$/i))continue;
      const p=await parseXLSX(f);
      if(!p)continue;
      newRecs.push(...p.recs);
      newFiles.push({label:p.label,count:p.recs.length,suc:p.suc,cols:p.cols,headers:p.headers});
    try{
      if(!excelFiles.length){alert("No se encontraron archivos Excel válidos (.xlsx o .xls).");return;}
      for(const f of excelFiles){
        const p=await parseXLSX(f);
        if(!p)continue;
        newRecs.push(...p.recs);
        newFiles.push({label:p.label,count:p.recs.length,suc:p.suc,cols:p.cols,headers:p.headers,headerRow:p.headerRow});
      }
      if(!newFiles.length){alert("No se pudo leer ninguna hoja con datos en los archivos Excel.");return;}
      setAllRecs(prev=>{
        // Remove seed placeholder records on first real load
        const withoutSeed = prev.filter(r=>r.file!=="2026_04_K27");
        const ex = new Set(withoutSeed.map(r=>r.file));
        const fresh = newRecs.filter(r=>!ex.has(r.file));
        return [...withoutSeed,...fresh];
      });
      setFiles(prev=>{
        const withoutSeed = prev.filter(f=>f.label!=="2026_04_K27");
        return [...withoutSeed,...newFiles.filter(f=>!withoutSeed.find(e=>e.label===f.label))];
      });
      if(!newRecs.length){
        alert("El Excel se leyó, pero no se encontraron filas con datos comerciales. Revisa el diagnóstico de columnas para ver qué campos no se detectaron.");
      }
    }catch(e){
      console.error("Error leyendo Excel",e);
      alert(`No se pudo leer el Excel: ${e?.message||e}`);
    }finally{
      setLoading(false);
    }
    if(!newRecs.length){setLoading(false);return;}
    setAllRecs(prev=>{
      // Remove seed placeholder records on first real load
      const withoutSeed = prev.filter(r=>r.file!=="2026_04_K27");
      const ex = new Set(withoutSeed.map(r=>r.file));
      const fresh = newRecs.filter(r=>!ex.has(r.file));
      return [...withoutSeed,...fresh];
    });
    setFiles(prev=>{
      const withoutSeed = prev.filter(f=>f.label!=="2026_04_K27");
      return [...withoutSeed,...newFiles.filter(f=>!withoutSeed.find(e=>e.label===f.label))];
    });
    setLoading(false);
  },[]);

  // — Firebase upload —
  const uploadToFirebase=useCallback(async(fs)=>{
    if(!FB_CONFIGURED){alert("Configura Firebase en el código primero (sección FB_CONFIG).");return;}
    setFbStatus("uploading");
    try{
      const {db,storage}=await getFB();
      for(const f of fs){
        const fileRef=storage.sRef(storage.storage,`excels/${f.name}`);
        await storage.uploadBytes(fileRef,f);
        const url=await storage.getDownloadURL(fileRef);
        await db.push(db.ref(db.db,"files"),{name:f.name,url,ts:Date.now()});
      }
      setFbStatus("ok");
      setTimeout(()=>setFbStatus("idle"),3000);
    }catch(e){console.error(e);setFbStatus("error");setTimeout(()=>setFbStatus("idle"),4000);}
  },[]);

  // — Sector normalizer — collapses spelling variants —
  const normSec=useCallback((s)=>{
    if(!s)return"Otros";
    const u=s.trim();
    if(/cnstr|const/i.test(u))return"Construcción";
    if(/alquil/i.test(u))return"Alquiler Maq.";
@@ -1294,74 +1333,74 @@ export default function App(){
              {n:5,t:"Pegar en el código",d:'Abre el archivo .jsx → busca "FB_CONFIG" al inicio → reemplaza los valores con los de tu firebaseConfig → guarda'},
              {n:6,t:"Deploy en GitHub Pages (hosting gratuito)",d:"Sube el proyecto a un repositorio GitHub → Settings → Pages → Source: main branch → Deploy → URL pública lista"},
            ].map(s=>(
              <div key={s.n} style={{display:"flex",gap:12,alignItems:"flex-start"}}>
                <div style={{width:28,height:28,borderRadius:6,background:B.red,color:"#fff",fontWeight:900,fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{s.n}</div>
                <div>
                  <div style={{fontWeight:700,color:B.white,fontSize:12,marginBottom:3}}>{s.t}</div>
                  <div style={{fontSize:11,color:B.gray2,lineHeight:1.6}}>{s.d}</div>
                </div>
              </div>
            ))}
          </div>
        </Pnl>
      )}

      {/* Upload zone */}
      <Pnl>
        <ST sub={FB_CONFIGURED?"Sube aquí y todos los usuarios verán los datos":"Configura Firebase primero para habilitar"}>Subir Excels a Firebase</ST>
        <div style={{border:`2px dashed ${FB_CONFIGURED?B.red:B.border}`,borderRadius:10,padding:"32px",textAlign:"center",opacity:FB_CONFIGURED?1:.5}}>
          <div style={{fontSize:32,marginBottom:10}}>☁️</div>
          <div style={{fontWeight:700,color:B.white,marginBottom:6}}>Arrastra Excels aquí para subir a la nube</div>
          <div style={{fontSize:11,color:B.gray3,marginBottom:14}}>Los archivos quedarán disponibles para todos los usuarios del dashboard</div>
          <button onClick={()=>{if(FB_CONFIGURED)document.getElementById("fb-upload").click();}} style={{background:FB_CONFIGURED?B.red:"transparent",border:`1px solid ${B.red}`,color:B.white,borderRadius:6,padding:"8px 20px",fontSize:11,cursor:FB_CONFIGURED?"pointer":"not-allowed",fontFamily:"inherit",fontWeight:700}}>
            {fbStatus==="uploading"?"⚙ Subiendo…":fbStatus==="ok"?"✅ Subido":fbStatus==="error"?"❌ Error":"Seleccionar archivos"}
          </button>
          <input id="fb-upload" type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onChange={e=>{const fs=Array.from(e.target.files);uploadToFirebase(fs);ingest(fs);}}/>
          <input id="fb-upload" type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onChange={e=>{const fs=Array.from(e.target.files);uploadToFirebase(fs);ingest(fs);e.target.value="";}}/>
        </div>
      </Pnl>

      {/* Column diagnostics */}
      {files.filter(f=>f.cols).length>0&&(
        <Pnl>
          <ST sub="Columnas detectadas en cada archivo — si 'Ingresos' dice null, el nombre de columna no coincide">🔍 Diagnóstico de columnas detectadas</ST>
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            {files.filter(f=>f.cols).slice(0,5).map((f,i)=>(
              <div key={i} style={{background:B.panel,borderRadius:8,padding:"12px 14px",border:`1px solid ${B.border}`}}>
                <div style={{fontWeight:700,color:B.white,fontSize:11,marginBottom:8}}>{f.label}</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:8}}>
                  {Object.entries(f.cols||{}).map(([field,col])=>(
                    <span key={field} style={{fontSize:10,padding:"2px 8px",borderRadius:4,
                      background:col?B.greenDim:B.redDim,
                      border:`1px solid ${col?B.green:B.red}`,
                      color:col?B.greenLt:B.red}}>
                      {field}: {col?`"${col}"`:"❌ NO DETECTADO"}
                    </span>
                  ))}
                </div>
                <div style={{fontSize:9,color:B.gray3}}>
                  Primeras columnas del archivo: {(f.headers||[]).join(" · ")}
                  Fila de encabezados detectada: {f.headerRow||1} · Primeras columnas del archivo: {(f.headers||[]).join(" · ")}
                </div>
              </div>
            ))}
          </div>
          <div style={{marginTop:12,padding:"10px 14px",background:B.amberDim,border:`1px solid ${B.amber}`,borderRadius:6,fontSize:11,color:B.amberLt,lineHeight:1.6}}>
            ⚠️ Si <strong>ing</strong> aparece como ❌ NO DETECTADO, el campo de ingresos no se encontró.<br/>
            Mira las columnas del archivo y escribe el nombre exacto al equipo para actualizar el detector.
          </div>
        </Pnl>
      )}

      {/* Files loaded */}
      <Pnl>
        <ST sub="Archivos actualmente cargados en esta sesión">Datos en memoria</ST>
        {files.length===0?<div style={{color:B.gray3,fontSize:11}}>Ningún archivo cargado</div>:(
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead><tr style={{borderBottom:`1px solid ${B.border}`}}>
              {["Archivo","Sucursal","Registros"].map(h=><th key={h} style={{padding:"5px 10px",textAlign:"left",color:B.gray3,fontSize:9,letterSpacing:".08em",textTransform:"uppercase"}}>{h}</th>)}
            </tr></thead>
            <tbody>
              {files.map((f,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${B.border}22`}}>
                  <td style={{padding:"7px 10px",color:B.white,fontWeight:600}}>{f.label}</td>
                  <td style={{padding:"7px 10px"}}><span style={{background:B.redBg,color:B.red,borderRadius:3,padding:"2px 7px",fontSize:10,fontWeight:700}}>{f.suc}</span></td>
                  <td style={{padding:"7px 10px",color:B.gray3}}>{fN(f.count)} líneas</td>
@@ -1451,51 +1490,51 @@ export default function App(){
            <button key={v} onClick={()=>setView(v)} style={{
              background:view===v?"rgba(0,0,0,.3)":"transparent",
              border:"none",borderBottom:view===v?"2px solid #fff":"2px solid transparent",
              color:view===v?B.white:"rgba(255,255,255,.58)",fontFamily:"inherit",
              fontSize:10,fontWeight:view===v?800:500,letterSpacing:".06em",
              padding:"15px 12px 13px",cursor:"pointer",transition:"all .12s",whiteSpace:"nowrap",flexShrink:0,
            }}>{v}</button>
          ))}
        </nav>

        {/* KPI mini */}
        <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",paddingLeft:14,borderLeft:"1px solid rgba(255,255,255,.25)",marginLeft:10,flexShrink:0}}>
          <span style={{fontWeight:900,color:B.white,fontSize:13,lineHeight:1}}>{f$(kpi.ing)}</span>
          <span style={{color:kpi.mgnP>20?"#86EFAC":kpi.mgnP>10?"#FDE68A":"#FCA5A5",fontWeight:700,fontSize:9,marginTop:2,letterSpacing:".04em"}}>{fP(kpi.mgnP)} margen</span>
        </div>

        {/* Upload btn */}
        <button onClick={()=>document.getElementById("fi-top").click()} style={{
          background:"rgba(0,0,0,.3)",border:"1px solid rgba(255,255,255,.38)",
          color:B.white,borderRadius:5,padding:"6px 13px",fontSize:10,
          cursor:"pointer",fontFamily:"inherit",fontWeight:800,letterSpacing:".06em",
          flexShrink:0,marginLeft:10,whiteSpace:"nowrap",
        }}>
          {loading?"⚙ Cargando…":"＋ EXCELS"}
        </button>
        <input id="fi-top" type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onChange={e=>ingest(Array.from(e.target.files))}/>
        <input id="fi-top" type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onChange={e=>{ingest(Array.from(e.target.files));e.target.value="";}}/>
      </div>

      {/* ── FILES BAR — fixed height, scrollable ── */}
      {files.length>0&&(
        <div style={{background:B.surface,borderBottom:`1px solid ${B.border}`,padding:"4px 16px",display:"flex",gap:6,alignItems:"center",flexShrink:0,overflowX:"auto",height:32,msOverflowStyle:"none",scrollbarWidth:"none"}}>
          <span style={{fontSize:9,color:B.gray3,letterSpacing:".12em",fontWeight:700,flexShrink:0}}>DATOS:</span>
          {files.map((f,i)=>(
            <span key={i} style={{fontSize:9,background:B.card,border:`1px solid ${B.border}`,borderRadius:3,padding:"2px 7px",color:B.gray3,borderLeft:`2px solid ${B.red}`,whiteSpace:"nowrap",flexShrink:0}}>
              {f.label} · {fN(f.count)}
            </span>
          ))}
          <span style={{fontSize:9,color:B.gray3,marginLeft:4,flexShrink:0}}>{fN(files.reduce((s,f)=>s+f.count,0))} registros</span>
        </div>
      )}

      {/* ── MAIN LAYOUT ── */}
      <div style={{display:"flex",flex:1,overflow:"hidden",minHeight:0}}>
        {/* Sidebar */}
        <div style={{width:208,flexShrink:0,background:B.surface,borderRight:`1px solid ${B.border}`,padding:14,overflowY:"auto",minHeight:0}}>
          <Sidebar/>
        </div>
        {/* Content */}
        <div className="fade-in" key={view} style={{flex:1,padding:"18px 22px",overflowY:"auto",minHeight:0,minWidth:0}}>
          {renderView()}
        </div>
