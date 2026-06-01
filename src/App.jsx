import { useState, useMemo, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { AreaChart, Area, BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, ScatterChart, Scatter, ZAxis, ComposedChart, RadarChart, Radar, PolarGrid, PolarAngleAxis } from "recharts";

// ╔══════════════════════════════════════════════════════╗
// ║           CUMMINS BRAND SYSTEM                       ║
// ╚══════════════════════════════════════════════════════╝
const B = {
  red:"#CC0000",redDk:"#990000",redBg:"#180000",redDim:"#280000",
  white:"#FFFFFF",gray1:"#E0E0E0",gray2:"#A0A0A0",gray3:"#666666",
  gray4:"#3A3A3A",gray5:"#242424",bg:"#0C0C0C",surface:"#121212",
  panel:"#181818",card:"#1C1C1C",border:"#2A2A2A",borderHi:"#404040",
  green:"#16A34A",greenDim:"#0A1A0A",greenLt:"#4ADE80",
  amber:"#D97706",amberDim:"#1C1000",amberLt:"#FCD34D",
  blue:"#2563EB",blueDim:"#071020",blueLt:"#93C5FD",
  teal:"#0D9488",tealDim:"#041410",
  chart:["#CC0000","#E57373","#D97706","#16A34A","#2563EB","#7C3AED","#DB2777","#0D9488","#CA8A04","#EA580C"],
};

const CumminsLogo = ({h=38})=>(
  <svg width={h*1.2} height={h} viewBox="0 0 96 80">
    <polygon points="48,2 90,22 90,58 48,78 6,58 6,22" fill={B.red}/>
    <text x="48" y="44" textAnchor="middle" fill="#fff" fontFamily="Arial Black,sans-serif" fontSize="16" fontWeight="900" letterSpacing=".5">CUMMINS</text>
    <text x="48" y="58" textAnchor="middle" fill="rgba(255,255,255,.65)" fontFamily="Arial,sans-serif" fontSize="7" letterSpacing="2">ZONA SUR</text>
  </svg>
);

// ╔══════════════════════════════════════════════════════╗
// ║  FIREBASE CONFIG  ← REEMPLAZA CON TUS CREDENCIALES  ║
// ╚══════════════════════════════════════════════════════╝
const FB_CONFIG = {
  apiKey:            "TU_API_KEY",
  authDomain:        "TU_PROJECT.firebaseapp.com",
  databaseURL:       "https://TU_PROJECT-default-rtdb.firebaseio.com",
  projectId:         "TU_PROJECT",
  storageBucket:     "TU_PROJECT.appspot.com",
  messagingSenderId: "TU_SENDER_ID",
  appId:             "TU_APP_ID",
};
const FB_CONFIGURED = !FB_CONFIG.apiKey.includes("TU_");

// Firebase lazy loader
let fbApp=null, fbDb=null, fbStorage=null;
async function getFB() {
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
  ing: ["ingreso tot","ingreso","ingresos","neto","ventas netas","venta neta","ventas","znetwr","importe neto","importe","valor neto","net value","netval"],
  cst: ["costo tot","costo total","costo","costos","cost of","coste","valor costo","costo merc"],
  cli: ["cliente","kunnr","cod. cliente","cod cliente","código cliente","codigo cliente"],
  clinm:["descripción cliente","descripcion cliente","nombre cliente","razon social","razón social","nombre"],
  f2:  ["jerarquia de producto 2","jerarquía de producto 2","jer. producto 2","familia 2","subfamilia"],
  f1:  ["jerarquia de producto","jerarquía de producto","jer. producto","familia","linea","línea"],
  art: ["articulo","artículo","matnr","código","codigo","código artículo","cod. art","material","part number"],
  suc: ["sucursal","descrip.sucursal","centro","werks","centro suministrador"],
  sec: ["descrip. gr.clie.","descrip.gr.clie.","descripcion grupo cliente","grupo de clientes","grupo clientes","sector","industria","segmento","descrip. gr"],
  rep: ["representante de ventas","representante ventas","ejecutivo","ejecutivo de ventas","vendedor","asesor","asesor comercial"],
  vol: ["vol.ventas","vol ventas","volumen","cantidad","unidades","qty","volume"],
  abc: ["indicador abc","ind. abc","ind.abc","abc"],
  fecha:["fecha factura","fecha de factura","fecha contab","fecha de contabilización","fecha contabilización","fecha","date","fec. factura"],
  per: ["período/año","periodo/año","período año","periodo año","periodo","period","mes"],
};
function dc(headers,f){const lo=headers.map(h=>String(h??"").toLowerCase().trim());const al=CA[f]||[];const i=lo.findIndex(h=>al.some(a=>h.includes(a)));return i!==-1?headers[i]:null;}
function sucFromName(n){const u=n.toUpperCase();if(u.includes("K27")||u.includes("AREQUIPA"))return"K27";if(u.includes("K11")||u.includes("HUANCAYO"))return"K11";if(u.includes("K23")||u.includes("TACNA"))return"K23";if(u.includes("K01")||u.includes("LIMA")||u.includes("CALLAO"))return"K01";return"K??"}
function periodFromName(n){const u=n.toUpperCase().replace(/[_\-\.]/g," ");const m=u.match(/(\d{4})\s+(\d{1,2})/);if(m)return{y:parseInt(m[1]),mo:parseInt(m[2])};const mi=SM.findIndex(s=>u.includes(s));if(mi!==-1){const ym=u.match(/(\d{4})/);if(ym)return{y:parseInt(ym[1]),mo:mi+1}}return null}

async function parseXLSX(file){
  const ab=await file.arrayBuffer();
  const wb=XLSX.read(ab,{type:"array",cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{defval:null});
  if(!rows.length)return null;
  const headers=Object.keys(rows[0]);
  const c={};for(const f of Object.keys(CA))c[f]=dc(headers,f);
  const label=file.name.replace(/\.[^.]+$/,"");
  const pfn=periodFromName(label), sfn=sucFromName(label);

  // Also detect "Fecha de contabilización" as date fallback
  const cFechaContab = headers.find(h=>h.toLowerCase().includes("contabiliz"));

  const recs=[];
  rows.forEach(r=>{
    const ing=parseFloat(r[c.ing])||0;
    const cst=parseFloat(r[c.cst])||0;
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
    recs.push({ing,cst,mgn:ing-cst,vol:parseFloat(r[c.vol])||0,
      cli:r[c.cli]?String(r[c.cli]).trim():"0",
      clinm:r[c.clinm]?String(r[c.clinm]).trim():"Sin nombre",
      f2:r[c.f2]?String(r[c.f2]).trim():"Sin familia",
      f1:r[c.f1]?String(r[c.f1]).trim():"Sin familia",
      art:r[c.art]?String(r[c.art]).trim():"?",
      suc:r[c.suc]?String(r[c.suc]).trim():sfn,
      sec:r[c.sec]?String(r[c.sec]).trim():"Otros",
      rep:r[c.rep]?String(r[c.rep]).trim():"0",
      abc:r[c.abc]?String(r[c.abc]).trim():"N/A",
      mo,yr,fy,fmi,fq,fml:FM[fmi],fyl:fyLbl(fy),file:label});
  });
  return{label,suc:sfn,recs,cols:c,headers:headers.slice(0,20)};
}

// ╔══════════════════════════════════════════════════════╗
// ║           SEED DATA  K27 Abril 2026                  ║
// ╚══════════════════════════════════════════════════════╝
const SEED={"meta":{"sucursal":"K27","sucursalNm":"Arequipa DCP","periodo":"Abril 2026","mes":4,"year":2026,"fyLabel":"FY 2026/27","ing":56030.15,"cst":42780.46,"mgn":13249.69,"lineas":409,"clientes":30,"articulos":197,"vol":1006},"clients":[{"id":6013813,"nm":"INTERNATIONAL CAMIONES DEL PERU","sector":"Automotriz","ing":17000.25,"cst":12809.5,"lineas":96,"vol":220,"mgn":4190.75,"mgnP":24.7,"pct":30.36,"cum":30.36,"seg":"TOP"},{"id":6027645,"nm":"COMERCIALIZADORA MIRIANDI E.I.R.L.","sector":"Canteras","ing":14305.91,"cst":8714.06,"lineas":8,"vol":33,"mgn":5591.85,"mgnP":39.1,"pct":25.55,"cum":55.91,"seg":"TOP"},{"id":6003804,"nm":"SOCIEDAD MINERA CERRO VERDE S.A.A.","sector":"Gran Minería","ing":7481.69,"cst":4766.36,"lineas":9,"vol":5,"mgn":2715.33,"mgnP":36.3,"pct":13.36,"cum":69.27,"seg":"MED"},{"id":6003931,"nm":"SOUTHERN PERU COPPER CORPORATION","sector":"Gran Minería","ing":2944.82,"cst":1793.4,"lineas":2,"vol":2,"mgn":1151.42,"mgnP":39.1,"pct":5.26,"cum":74.53,"seg":"MED"},{"id":6003863,"nm":"UNIMAQ S.A","sector":"Automotriz","ing":2839.85,"cst":2845.27,"lineas":27,"vol":69,"mgn":-5.42,"mgnP":-0.2,"pct":5.07,"cum":79.6,"seg":"MED"},{"id":6008872,"nm":"LADRILLERA EL DIAMANTE S.A.C.","sector":"Construccion","ing":1920.58,"cst":1827.88,"lineas":18,"vol":136,"mgn":92.7,"mgnP":4.8,"pct":3.43,"cum":83.03,"seg":"MED"},{"id":6002240,"nm":"REPUESTOS DAVID DIESEL E.I.R.L.","sector":"Automotriz","ing":1781.32,"cst":1770.8,"lineas":11,"vol":76,"mgn":10.52,"mgnP":0.6,"pct":3.18,"cum":86.21,"seg":"MED"},{"id":6003685,"nm":"TRACTO CAMIONES USA S.A.C.","sector":"Automotriz","ing":1359.38,"cst":1118.86,"lineas":46,"vol":142,"mgn":240.52,"mgnP":17.7,"pct":2.43,"cum":88.64,"seg":"MED"},{"id":6016191,"nm":"LUBRICANTES PIEYYRA E.I.R.L.","sector":"Automotriz","ing":1285.95,"cst":1198.53,"lineas":40,"vol":101,"mgn":87.42,"mgnP":6.8,"pct":2.3,"cum":90.94,"seg":"LOW"},{"id":6002717,"nm":"PERURAIL S.A.","sector":"Automotriz","ing":1196.88,"cst":1135.56,"lineas":6,"vol":52,"mgn":61.32,"mgnP":5.1,"pct":2.14,"cum":93.08,"seg":"LOW"},{"id":6023634,"nm":"FUNDO HNOS DOMINGUEZ S.A.C.","sector":"Agropecuario","ing":838.86,"cst":1049.95,"lineas":5,"vol":4,"mgn":-211.09,"mgnP":-25.2,"pct":1.5,"cum":94.58,"seg":"LOW"},{"id":6002578,"nm":"TRANSPORTES HAGEMSA S.A.C.","sector":"Automotriz","ing":763.84,"cst":777.4,"lineas":5,"vol":22,"mgn":-13.56,"mgnP":-1.8,"pct":1.36,"cum":95.94,"seg":"LOW"},{"id":6021821,"nm":"CAMFRA E.I.R.L.","sector":"Automotriz","ing":468.36,"cst":471.72,"lineas":3,"vol":13,"mgn":-3.36,"mgnP":-0.7,"pct":0.84,"cum":96.78,"seg":"LOW"},{"id":6003898,"nm":"EPIROC PERU SOCIEDAD ANONIMA","sector":"Min. Sub & Loc. Dril","ing":391.25,"cst":274.55,"lineas":3,"vol":5,"mgn":116.7,"mgnP":29.8,"pct":0.7,"cum":97.48,"seg":"LOW"},{"id":6022633,"nm":"AR MAQUINARIAS DEL PERU SOCIEDAD","sector":"Min. Sub & Loc. Dril","ing":321.46,"cst":291.91,"lineas":1,"vol":50,"mgn":29.55,"mgnP":9.2,"pct":0.57,"cum":98.05,"seg":"LOW"},{"id":6026321,"nm":"MULTISERVICIOS Y OPERACIONES JBN","sector":"Comercio","ing":286.0,"cst":196.83,"lineas":3,"vol":3,"mgn":89.17,"mgnP":31.2,"pct":0.51,"cum":98.56,"seg":"LOW"},{"id":6026883,"nm":"HAZVOZ E.I.R.L.","sector":"Utilities / Telecom","ing":201.97,"cst":183.42,"lineas":15,"vol":42,"mgn":18.55,"mgnP":9.2,"pct":0.36,"cum":98.92,"seg":"LOW"},{"id":6013017,"nm":"CONSORCIO EMPRESA DEPURADORA DE","sector":"Utilities / Telecom","ing":191.1,"cst":64.58,"lineas":2,"vol":2,"mgn":126.52,"mgnP":66.2,"pct":0.34,"cum":99.26,"seg":"LOW"},{"id":6000430,"nm":"RD RENTAL S.A.C","sector":"Min. Sub & Loc. Dril","ing":190.74,"cst":174.68,"lineas":9,"vol":21,"mgn":16.06,"mgnP":8.4,"pct":0.34,"cum":99.6,"seg":"LOW"},{"id":6020591,"nm":"AGROGANADO LA ABRIL E.I.R.L.","sector":"Agropecuario","ing":85.18,"cst":57.02,"lineas":12,"vol":79,"mgn":28.16,"mgnP":33.1,"pct":0.15,"cum":99.75,"seg":"LOW"},{"id":6026791,"nm":"CONSTRUCTORA E INMOBILIARIA IBER","sector":"Construccion","ing":59.31,"cst":52.72,"lineas":4,"vol":25,"mgn":6.59,"mgnP":11.1,"pct":0.11,"cum":99.86,"seg":"LOW"},{"id":6014498,"nm":"FERTICA S.A.C.","sector":"Utilities / Telecom","ing":52.34,"cst":66.17,"lineas":2,"vol":2,"mgn":-13.83,"mgnP":-26.4,"pct":0.09,"cum":99.95,"seg":"LOW"},{"id":6025555,"nm":"CONTRATISTAS GENERALES ARRIOLA","sector":"Construccion","ing":42.13,"cst":25.84,"lineas":4,"vol":23,"mgn":16.29,"mgnP":38.7,"pct":0.08,"cum":100.03,"seg":"LOW"},{"id":6017393,"nm":"CONSULTORA DE SISTEMAS INTEGRADO","sector":"Utilities / Telecom","ing":24.71,"cst":18.97,"lineas":2,"vol":2,"mgn":5.74,"mgnP":23.2,"pct":0.04,"cum":100.07,"seg":"LOW"},{"id":6026900,"nm":"H.S. ELECTROMECANICA S.A.C.","sector":"Min. Sub & Loc. Dril","ing":19.54,"cst":11.64,"lineas":2,"vol":2,"mgn":7.9,"mgnP":40.4,"pct":0.03,"cum":100.1,"seg":"LOW"},{"id":6003218,"nm":"PORTUARIA PERUANO SUIZA S.A.C","sector":"Portuario","ing":14.7,"cst":8.72,"lineas":2,"vol":2,"mgn":5.98,"mgnP":40.7,"pct":0.03,"cum":100.13,"seg":"LOW"}],"familia":[{"f1":"Filtración","f2":"Filtración Aftermark","ing":13259.8,"cst":13886.24,"vol":587,"mgn":-626.44,"mgnP":-4.7},{"f1":"Grupo Electrógeno","f2":"Generador Diesel","ing":13127.91,"cst":8714.06,"vol":1,"mgn":4413.85,"mgnP":33.6},{"f1":"Partes Motores","f2":"Partes Mot. HD","ing":12800.72,"cst":9698.57,"vol":214,"mgn":3102.15,"mgnP":24.2},{"f1":"Partes Motores","f2":"Partes Mot.HHP","ing":9226.64,"cst":5443.47,"vol":8,"mgn":3783.17,"mgnP":41.0},{"f1":"Partes Recon","f2":"Partes Recon HD","ing":4556.09,"cst":3367.57,"vol":5,"mgn":1188.52,"mgnP":26.1},{"f1":"Servicio CS","f2":"Servicio CS","ing":1178.0,"cst":0.0,"vol":82,"mgn":1178.0,"mgnP":100.0},{"f1":"Partes Motores","f2":"Partes Mot. MR","ing":658.24,"cst":478.59,"vol":30,"mgn":179.65,"mgnP":27.3},{"f1":"Lubricantes","f2":"Lubricantes","ing":391.25,"cst":274.55,"vol":5,"mgn":116.7,"mgnP":29.8},{"f1":"Partes Motores","f2":"Partes Compra Mot. B","ing":334.03,"cst":168.87,"vol":14,"mgn":165.16,"mgnP":49.4},{"f1":"Seguridad Industrial","f2":"Paños varios","ing":321.46,"cst":291.91,"vol":50,"mgn":29.55,"mgnP":9.2},{"f1":"Filtración","f2":"Filtración Accesorio","ing":155.83,"cst":148.68,"vol":9,"mgn":7.15,"mgnP":4.6},{"f1":"Filtración","f2":"Filtración Mineria","ing":20.18,"cst":307.95,"vol":1,"mgn":-287.77,"mgnP":-1426.0}],"sector":[{"sec":"Automotriz","ing":25590.01,"lineas":214,"clientes":10},{"sec":"Canteras","ing":14305.91,"lineas":8,"clientes":1},{"sec":"Gran Minería","ing":10426.51,"lineas":27,"clientes":3},{"sec":"Construccion","ing":2022.02,"lineas":72,"clientes":5},{"sec":"Min. Sub & Loc. Dril","ing":1677.2,"lineas":49,"clientes":3},{"sec":"Utilities / Telecom","ing":1578.72,"lineas":17,"clientes":3},{"sec":"Comercio","ing":286.0,"lineas":3,"clientes":1},{"sec":"Agropecuario","ing":85.18,"lineas":12,"clientes":1},{"sec":"Portuario","ing":19.54,"lineas":2,"clientes":1}],"byday":[{"dia":1,"ing":37836.6,"lineas":123},{"dia":2,"ing":6462.83,"lineas":8},{"dia":4,"ing":2854.87,"lineas":24},{"dia":5,"ing":5012.65,"lineas":19},{"dia":6,"ing":3863.2,"lineas":32}],"reps":[{"rep":"2193","ing":23336.44,"clientes":7,"lineas":117},{"rep":"2184","ing":14305.91,"clientes":1,"lineas":6},{"rep":"3622","ing":7481.69,"clientes":1,"lineas":6},{"rep":"3768","ing":4794.19,"clientes":7,"lineas":36},{"rep":"184","ing":2944.82,"clientes":1,"lineas":2},{"rep":"7692","ing":1203.81,"clientes":7,"lineas":18},{"rep":"3192","ing":1124.43,"clientes":1,"lineas":12},{"rep":"8626","ing":838.86,"clientes":1,"lineas":5}],"articles":[{"art":"G1C60D6EIC","f1":"Grupo Electrógeno","f2":"Generador Diesel","ing":13127.91,"cst":8714.06,"vol":1,"mgn":4413.85,"mgnP":33.6},{"art":"CM553894900","f1":"Partes Motores","f2":"Partes Mot. HD","ing":4408.32,"cst":2701.28,"vol":1,"mgn":1707.04,"mgnP":38.7},{"art":"LF4054","f1":"Filtración","f2":"Filtración Aftermark","ing":4834.9,"cst":4760.92,"vol":287,"mgn":73.98,"mgnP":1.5},{"art":"CM642951100","f1":"Partes Motores","f2":"Partes Mot. HD","ing":3102.75,"cst":2422.57,"vol":1,"mgn":680.18,"mgnP":21.9},{"art":"LF4054F","f1":"Filtración","f2":"Filtración Aftermark","ing":2989.79,"cst":3050.94,"vol":97,"mgn":-61.15,"mgnP":-2.0},{"art":"AF4501","f1":"Filtración","f2":"Filtración Aftermark","ing":2929.62,"cst":2989.25,"vol":104,"mgn":-59.63,"mgnP":-2.0},{"art":"CM637714700","f1":"Partes Motores","f2":"Partes Mot.HHP","ing":2800.02,"cst":1690.65,"vol":1,"mgn":1109.37,"mgnP":39.6},{"art":"CM341175600RX","f1":"Partes Recon","f2":"Partes Recon HD","ing":2684.72,"cst":2073.38,"vol":4,"mgn":611.34,"mgnP":22.8},{"art":"LF9009","f1":"Filtración","f2":"Filtración Aftermark","ing":1714.61,"cst":1776.68,"vol":39,"mgn":-62.07,"mgnP":-3.6},{"art":"4387765","f1":"Partes Motores","f2":"Partes Mot. HD","ing":1461.7,"cst":1117.82,"vol":1,"mgn":343.88,"mgnP":23.5},{"art":"LF4005F","f1":"Filtración","f2":"Filtración Aftermark","ing":1575.19,"cst":1580.66,"vol":50,"mgn":-5.47,"mgnP":-0.3},{"art":"LF16015","f1":"Filtración","f2":"Filtración Aftermark","ing":1362.94,"cst":1416.64,"vol":32,"mgn":-53.7,"mgnP":-3.9},{"art":"3966323","f1":"Partes Motores","f2":"Partes Mot. HD","ing":1295.12,"cst":990.79,"vol":30,"mgn":304.33,"mgnP":23.5},{"art":"4387766","f1":"Partes Motores","f2":"Partes Mot. HD","ing":1191.5,"cst":911.28,"vol":1,"mgn":280.22,"mgnP":23.5},{"art":"LF3349","f1":"Filtración","f2":"Filtración Aftermark","ing":1027.64,"cst":1056.13,"vol":25,"mgn":-28.49,"mgnP":-2.8}],"abc":[{"ind":"A","ing":16553.93,"lineas":112,"vol":402},{"ind":"B","ing":2446.16,"lineas":32,"vol":87},{"ind":"C","ing":3644.12,"lineas":20,"vol":48},{"ind":"D","ing":234.89,"lineas":4,"vol":6}]};

// ╔══════════════════════════════════════════════════════╗
// ║           FORMATTING & UTILS                         ║
// ╚══════════════════════════════════════════════════════╝
const f$=(n,d=0)=>{if(n==null)return"—";const a=Math.abs(n);const s=a>=1e6?`$${(a/1e6).toFixed(2)}M`:a>=1e3?`$${(a/1e3).toFixed(1)}K`:`$${a.toFixed(d)}`;return n<0?`(${s})`:s};
const fP=(n,d=1)=>n!=null?`${n>0?"+":""}${n.toFixed(d)}%`:"—";
const fN=(n)=>n?.toLocaleString("es-PE")??"—";
const clamp=(v,a,b)=>Math.min(Math.max(v,a),b);
const sortBy=(arr,key,asc=false)=>[...arr].sort((a,b)=>asc?a[key]-b[key]:b[key]-a[key]);

// ╔══════════════════════════════════════════════════════╗
// ║           AUTO-ANALYSIS ENGINE                       ║
// ╚══════════════════════════════════════════════════════╝
function autoAnalysis(clients,familia,reps,kpi,repNames){
  const alerts=[],opps=[],actions=[];
  familia.filter(f=>f.mgnP<0&&Math.abs(f.ing)>50).forEach(f=>{
    alerts.push({lvl:"CRÍTICO",icon:"🔴",t:`${f.f2}: margen ${f.mgnP.toFixed(0)}%`,b:`Pérdida de ${f$(Math.abs(f.mgn))} en ${f.vol} unidades. Actualizar precio de lista en SAP.`});
  });
  const top2pct=clients.slice(0,2).reduce((s,c)=>s+c.pct,0);
  if(top2pct>50) alerts.push({lvl:"RIESGO",icon:"🟡",t:`Concentración: ${top2pct.toFixed(0)}% en 2 clientes`,b:`${clients[0]?.nm.split(" ").slice(0,2).join(" ")} + ${clients[1]?.nm.split(" ").slice(0,2).join(" ")}. Si uno falla, el mes cae.`});
  clients.filter(c=>c.mgnP<-1&&c.ing>300).forEach(c=>alerts.push({lvl:"MARGEN",icon:"🟠",t:`${c.nm.split(" ").slice(0,3).join(" ")}: margen negativo`,b:`${f$(c.ing)} facturado con pérdida ${f$(c.mgn)}. Revisar descuentos.`}));
  const topRep=reps[0];
  if(topRep&&(topRep.ing/kpi.ing)*100>40) alerts.push({lvl:"RIESGO",icon:"🟡",t:`Rep ${repNames[topRep.rep]||topRep.rep}: ${((topRep.ing/kpi.ing)*100).toFixed(0)}% del ingreso`,b:"Dependencia alta de un solo ejecutivo."});
  familia.filter(f=>f.mgnP>30&&f.ing>1000).forEach(f=>opps.push({icon:"💎",t:`${f.f2}: ${f.mgnP.toFixed(0)}% margen`,b:`Alta rentabilidad. Penetrar más clientes con este producto.`}));
  clients.filter(c=>c.mgnP>35&&c.ing>300&&c.lineas<12).forEach(c=>opps.push({icon:"🚀",t:`${c.nm.split(" ").slice(0,3).join(" ")}: rentable, poco trabajado`,b:`${c.mgnP.toFixed(0)}% margen en solo ${c.lineas} líneas SAP.`}));
  familia.filter(f=>f.vol>50&&f.mgnP<5&&f.mgnP>=0).forEach(f=>opps.push({icon:"📦",t:`${f.f2}: subir precio ${(5-f.mgnP).toFixed(0)}%`,b:`Alto volumen (${f.vol} uds) con margen bajo. +$${(f.ing*0.03).toFixed(0)} si se sube 3%.`}));
  if(alerts.some(a=>a.lvl==="CRÍTICO")) actions.push(`Revisión urgente de precios: ${familia.filter(f=>f.mgnP<0).map(f=>f.f2).join(", ")} — presentar a Pricing esta semana`);
  actions.push(`Visita a ${clients[0]?.nm.split(" ").slice(0,2).join(" ")} para confirmar pipeline siguiente mes (${clients[0]?.pct.toFixed(0)}% del ingreso)`);
  const minCli=clients.filter(c=>c.sector.toLowerCase().includes("miner")&&c.lineas<10);
  if(minCli.length) actions.push(`Ampliar oferta a ${minCli.map(c=>c.nm.split(" ")[0]).slice(0,2).join(", ")} — clientes mineros con pocas líneas`);
  clients.filter(c=>c.mgnP<-0.5&&c.ing>500).forEach(c=>actions.push(`Revisar condición comercial de ${c.nm.split(" ").slice(0,2).join(" ")} — margen negativo`));
  actions.push(`Prospectar nuevos clientes Minería/Construcción para reducir concentración Automotriz`);
  return{alerts,opps,actions:actions.slice(0,6)};
}

// ╔══════════════════════════════════════════════════════╗
// ║           RECHARTS TOOLTIP                           ║
// ╚══════════════════════════════════════════════════════╝
const TT=({active,payload,label})=>{
  if(!active||!payload?.length)return null;
  return(
    <div style={{background:B.gray5,border:`1px solid ${B.borderHi}`,borderRadius:8,padding:"10px 14px",fontSize:11,boxShadow:"0 8px 32px #0009"}}>
      <p style={{color:B.red,fontWeight:800,fontSize:11,marginBottom:5,letterSpacing:".04em"}}>{label}</p>
      {payload.map((p,i)=>(
        <p key={i} style={{color:p.color||B.white,margin:"2px 0",display:"flex",justifyContent:"space-between",gap:16}}>
          <span style={{color:B.gray2,fontSize:10}}>{p.name}</span>
          <strong>{Math.abs(p.value??0)>100?f$(p.value):typeof p.value==="number"?p.value.toFixed(1):p.value}</strong>
        </p>
      ))}
    </div>
  );
};

const VIEWS=["RESUMEN","CLIENTES","PRODUCTOS","SECTORES","TENDENCIAS","EJECUTIVOS","TABLA DINÁMICA","ANÁLISIS","FIREBASE"];

// ╔══════════════════════════════════════════════════════╗
// ║           MAIN APP                                   ║
// ╚══════════════════════════════════════════════════════╝
export default function App(){
  const [allRecs,  setAllRecs]  = useState([]);
  const [files,    setFiles]    = useState([]);
  const [view,     setView]     = useState("RESUMEN");
  const [loading,  setLoading]  = useState(false);
  const [dragging, setDragging] = useState(false);
  const [fbStatus, setFbStatus] = useState("idle"); // idle|uploading|ok|error

  // — Filters (local per-user, not shared) —
  const [fSuc, setFSuc] = useState([]);
  const [fFY,  setFY]   = useState([]);
  const [fQ,   setFQ]   = useState([]);
  const [fSec, setFSec] = useState([]);
  const [fFam, setFFam] = useState([]);
  const [fRep, setFRep] = useState([]);
  const [fSeg, setFSeg] = useState([]);
  const [search,setSearch]=useState("");

  // — Cartera & rep names (persisted locally) —
  const [cartera,  setCartera]  = useState(()=>{try{return JSON.parse(localStorage.getItem("cummins_c")||"{}")}catch{return{}}});
  const [repNames, setRepNames] = useState(()=>{try{return JSON.parse(localStorage.getItem("cummins_r")||"{}")}catch{return{}}});
  const [selCli,   setSelCli]   = useState(null);

  // — Pivot config —
  const [pivotRows, setPivotRows]   = useState("sec");
  const [pivotCols, setPivotCols]   = useState("fyl");
  const [pivotVal,  setPivotVal]    = useState("ing");
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
    const newRecs=[],newFiles=[];
    for(const f of fs){
      if(!f.name.match(/\.xlsx?$/i))continue;
      const p=await parseXLSX(f);
      if(!p)continue;
      newRecs.push(...p.recs);
      newFiles.push({label:p.label,count:p.recs.length,suc:p.suc,cols:p.cols,headers:p.headers});
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
    if(/empre.*ext|extern/i.test(u))return"Empresa Externa";
    if(/mediana.*min|min.*med/i.test(u))return"Mediana Minería";
    if(/gran.*min|min.*gran/i.test(u))return"Gran Minería";
    if(/peq.*min|min.*peq/i.test(u))return"Pequeña Minería";
    if(/min.*sub|sub.*min|loc.*dril|dril/i.test(u))return"Min. Sub & Dril.";
    if(/util|telec/i.test(u))return"Utilities/Telecom";
    if(/agrop|agric|fundo/i.test(u))return"Agropecuario";
    if(/autom|camion|transport/i.test(u))return"Automotriz";
    if(/canter/i.test(u))return"Canteras";
    if(/comerc/i.test(u))return"Comercio";
    if(/marino|portu/i.test(u))return"Portuario";
    if(/indust/i.test(u))return"Industrial";
    if(/pesca/i.test(u))return"Pesca";
    if(/gas|oil/i.test(u))return"Oil & Gas";
    if(/person|natural/i.test(u))return"Persona Natural";
    if(/divers/i.test(u))return"Clientes Div.";
    return u.length>20?u.slice(0,19)+"…":u;
  },[]);

  // — Filter options —
  const opts=useMemo(()=>({
    sucs:[...new Set(allRecs.map(r=>r.suc))].sort(),
    fys: [...new Set(allRecs.map(r=>r.fyl))].sort(),
    secs:[...new Set(allRecs.map(r=>normSec(r.sec)))].sort(),
    reps:[...new Set(allRecs.map(r=>r.rep).filter(r=>r&&r!=="0"))].sort(),
    fams:[...new Set(allRecs.map(r=>r.f2).filter(f=>f&&f!=="—"))].sort(),
  }),[allRecs,normSec]);

  // — Apply filters — use normSec for sector matching —
  const filtered=useMemo(()=>allRecs.filter(r=>{
    if(fSuc.length&&!fSuc.includes(r.suc))return false;
    if(fFY.length &&!fFY.includes(r.fyl)) return false;
    if(fQ.length  &&!fQ.some(q=>parseInt(q)===r.fq))return false;
    if(fSec.length&&!fSec.includes(normSec(r.sec)))return false;
    if(fFam.length&&!fFam.includes(r.f2)) return false;
    if(fRep.length&&!fRep.includes(r.rep))return false;
    if(search){const q=search.toLowerCase();if(!r.clinm.toLowerCase().includes(q)&&!r.cli.includes(q)&&!r.art.toLowerCase().includes(q))return false;}
    return true;
  }),[allRecs,fSuc,fFY,fQ,fSec,fFam,fRep,search,normSec]);

  const useSeed=files.length===1&&files[0]?.label==="2026_04_K27"&&!fSuc.length&&!fFY.length&&!fQ.length&&!fSec.length&&!fFam.length&&!fRep.length&&!search;

  // — KPIs —
  const kpi=useMemo(()=>{
    if(useSeed)return{ing:SEED.meta.ing,cst:SEED.meta.cst,mgn:SEED.meta.mgn,mgnP:SEED.meta.mgn/SEED.meta.ing*100,vol:SEED.meta.vol,lineas:SEED.meta.lineas,clis:SEED.meta.clientes};
    const ing=filtered.reduce((s,r)=>s+r.ing,0);
    const cst=filtered.reduce((s,r)=>s+r.cst,0);
    const mgn=ing-cst;
    return{ing,cst,mgn,mgnP:ing?mgn/ing*100:0,vol:filtered.reduce((s,r)=>s+r.vol,0),lineas:filtered.length,clis:new Set(filtered.map(r=>r.cli)).size};
  },[filtered,useSeed]);

  // — Aggregations — always run useMemo (React rules), then pick seed vs computed
  const clientsComputed=useMemo(()=>{
    const m={};
    filtered.forEach(r=>{
      if(!m[r.cli])m[r.cli]={id:r.cli,nm:r.clinm,sector:r.sec,ing:0,cst:0,mgn:0,vol:0,lineas:0,rep:r.rep};
      const c=m[r.cli];c.ing+=r.ing;c.cst+=r.cst;c.mgn+=r.mgn;c.vol+=r.vol;c.lineas++;
    });
    const arr=sortBy(Object.values(m),"ing").map(c=>({...c,mgnP:c.ing?c.mgn/c.ing*100:0}));
    const ti=arr.reduce((s,c)=>s+c.ing,0);let cum=0;
    return arr.map(c=>{c.pct=ti?c.ing/ti*100:0;cum+=c.pct;c.cum=cum;c.seg=cum<=60?"TOP":cum<=90?"MED":"LOW";return c;});
  },[filtered]);

  const familiaComputed=useMemo(()=>{
    const m={};
    filtered.forEach(r=>{if(!m[r.f2])m[r.f2]={f2:r.f2,f1:r.f1,ing:0,cst:0,vol:0,mgn:0};const f=m[r.f2];f.ing+=r.ing;f.cst+=r.cst;f.vol+=r.vol;f.mgn+=r.mgn;});
    return sortBy(Object.values(m),"ing").map(f=>({...f,mgnP:f.ing?f.mgn/f.ing*100:0}));
  },[filtered]);

  const sectorComputed=useMemo(()=>{
    const m={};
    filtered.forEach(r=>{
      const k=normSec(r.sec);
      if(!m[k])m[k]={sec:k,ing:0,lineas:0,clis:new Set()};
      m[k].ing+=r.ing;m[k].lineas++;m[k].clis.add(r.cli);
    });
    return sortBy(Object.values(m),"ing").filter(s=>s.ing>0).map(s=>({...s,clientes:s.clis.size}));
  },[filtered,normSec]);

  const repsComputed=useMemo(()=>{
    const m={};
    filtered.forEach(r=>{if(!r.rep||r.rep==="0")return;if(!m[r.rep])m[r.rep]={rep:r.rep,ing:0,clis:new Set(),lineas:0};m[r.rep].ing+=r.ing;m[r.rep].clis.add(r.cli);m[r.rep].lineas++;});
    return sortBy(Object.values(m),"ing").filter(r=>r.ing>0).map(r=>({...r,clientes:r.clis.size}));
  },[filtered]);

  const articlesComputed=useMemo(()=>{
    const m={};
    filtered.forEach(r=>{if(!m[r.art])m[r.art]={art:r.art,f1:r.f1,f2:r.f2,ing:0,cst:0,vol:0,mgn:0,lineas:0};const a=m[r.art];a.ing+=r.ing;a.cst+=r.cst;a.vol+=r.vol;a.mgn+=r.mgn;a.lineas++;});
    return sortBy(Object.values(m),"ing").slice(0,30).map(a=>({...a,mgnP:a.ing?a.mgn/a.ing*100:0}));
  },[filtered]);

  // Pick seed (accurate) or computed
  const clients  = useSeed ? SEED.clients  : clientsComputed;
  const familia  = useSeed ? SEED.familia  : familiaComputed;
  const sector   = useSeed ? SEED.sector   : sectorComputed;
  const repsAgg  = useSeed ? SEED.reps     : repsComputed;
  const articles = useSeed ? SEED.articles : articlesComputed;

  const fys=[...new Set(allRecs.map(r=>r.fyl))].sort();

  const fiscalTrend=useMemo(()=>FM.map((m,i)=>{
    const row={mes:m};
    fys.forEach(fy=>{row[fy]=filtered.filter(r=>r.fyl===fy&&r.fmi===i).reduce((s,r)=>s+r.ing,0);});
    return row;
  }),[filtered,fys]);

  const fyTotals=useMemo(()=>fys.map(fy=>{
    const rows=filtered.filter(r=>r.fyl===fy);
    const ing=rows.reduce((s,r)=>s+r.ing,0),mgn=rows.reduce((s,r)=>s+r.mgn,0);
    return{fy,ing,mgn,mgnP:ing?mgn/ing*100:0,lineas:rows.length,clis:new Set(rows.map(r=>r.cli)).size};
  }),[filtered,fys]);

  const quarterData=useMemo(()=>[1,2,3,4].map((q,i)=>{
    const row={q:QNames[i]};
    fys.forEach(fy=>{row[fy]=filtered.filter(r=>r.fyl===fy&&r.fq===q).reduce((s,r)=>s+r.ing,0);});
    return row;
  }),[filtered,fys]);

  // — Pivot table engine —
  const pivotData=useMemo(()=>{
    const PDIMS={
      sec:{label:"Sector",get:(r)=>r.sec},
      f1: {label:"Familia",get:(r)=>r.f1},
      f2: {label:"Sub-familia",get:(r)=>r.f2},
      fyl:{label:"Año Fiscal",get:(r)=>r.fyl},
      fml:{label:"Mes Fiscal",get:(r)=>r.fml},
      fq: {label:"Trimestre",get:(r)=>QNames[r.fq-1]||String(r.fq)},
      suc:{label:"Sucursal",get:(r)=>r.suc},
      rep:{label:"Ejecutivo",get:(r)=>r.rep},
      abc:{label:"ABC",get:(r)=>r.abc},
    };
    const source=useSeed
      ? seedRecs.map(r=>({...r,seg:SEED.clients.find(c=>String(c.id)===r.cli)?.seg||"LOW"}))
      : filtered;
    const rowDimFn=PDIMS[pivotRows]?.get||(r=>r.sec);
    const colDimFn=PDIMS[pivotCols]?.get||(r=>r.fyl);
    const matrix={};
    const colSet=new Set();
    source.forEach(r=>{
      const rv=rowDimFn(r),cv=colDimFn(r);
      if(!rv||!cv)return;
      colSet.add(cv);
      if(!matrix[rv])matrix[rv]={};
      matrix[rv][cv]=(matrix[rv][cv]||0)+(pivotVal==="lineas"?1:(r[pivotVal]||0));
    });
    const cols=[...colSet].sort();
    let rows=Object.entries(matrix).map(([rowKey,vals])=>{
      const total=cols.reduce((s,c)=>s+(vals[c]||0),0);
      return{rowKey,vals,total};
    });
    rows=pivotSort==="desc"?rows.sort((a,b)=>b.total-a.total):rows.sort((a,b)=>a.total-b.total);
    return{rows,cols};
  },[filtered,useSeed,seedRecs,pivotRows,pivotCols,pivotVal,pivotSort]);

  const PIVOT_DIMS_META={
    sec:"Sector",f1:"Familia",f2:"Sub-familia",fyl:"Año Fiscal",
    fml:"Mes Fiscal",fq:"Trimestre",suc:"Sucursal",rep:"Ejecutivo",abc:"ABC",
  };
  const PIVOT_VALS_META={
    ing:{label:"Ingresos $",fmt:f$},cst:{label:"Costo $",fmt:f$},
    mgn:{label:"Margen $",fmt:f$},vol:{label:"Volumen",fmt:fN},lineas:{label:"Líneas",fmt:fN},
  };

  const analysis=useMemo(()=>autoAnalysis(clients,familia,repsAgg,kpi,repNames),[clients,familia,repsAgg,kpi,repNames]);

  // — Helpers —
  const toggle=(s,v)=>s(p=>p.includes(v)?p.filter(x=>x!==v):[...p,v]);
  const clearAll=()=>{setFSuc([]);setFY([]);setFQ([]);setFSec([]);setFFam([]);setFRep([]);setFSeg([]);setSearch("");};
  const hasFilters=fSuc.length||fFY.length||fQ.length||fSec.length||fFam.length||fRep.length||fSeg.length||search;

  const filteredClients=useMemo(()=>clients.filter(c=>{
    if(fSeg.length&&!fSeg.includes(c.seg))return false;
    if(search){const q=search.toLowerCase();return c.nm.toLowerCase().includes(q)||String(c.id).includes(q);}
    return true;
  }),[clients,fSeg,search]);

  const cliMonthly=useMemo(()=>{
    if(!selCli)return[];
    const rows=filtered.filter(r=>r.cli===String(selCli.id));
    return FM.map((m,i)=>({mes:m,ing:rows.filter(r=>r.fmi===i).reduce((s,r)=>s+r.ing,0)}));
  },[filtered,selCli]);

  // ── UI COMPONENTS ─────────────────────────────────────────

  const Pill=({label,active,color=B.red,onClick})=>(
    <button onClick={onClick} style={{
      background:active?`${color}20`:"transparent",border:`1px solid ${active?color:B.border}`,
      color:active?color:B.gray3,borderRadius:20,padding:"3px 10px",fontSize:10,
      cursor:"pointer",fontFamily:"inherit",fontWeight:active?700:400,
      letterSpacing:".04em",transition:"all .12s",whiteSpace:"nowrap",lineHeight:"16px",
    }}>{label}</button>
  );

  const KCard=({icon,label,value,sub,color=B.red,warn=false})=>(
    <div style={{background:warn?B.redBg:B.card,border:`1px solid ${warn?B.red:B.border}`,
      borderRadius:10,padding:"14px 16px",borderTop:`2px solid ${warn?B.red:color}`,
      flex:"1 1 150px",minWidth:150,maxWidth:260}}>
      <div style={{fontSize:9,color:B.gray3,letterSpacing:".12em",textTransform:"uppercase",marginBottom:5,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{icon} {label}</div>
      <div style={{fontSize:17,fontWeight:800,color:warn?B.red:color,lineHeight:1.2,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{value}</div>
      {sub&&<div style={{fontSize:10,color:B.gray3,marginTop:4,lineHeight:1.4}}>{sub}</div>}
    </div>
  );

  const Pnl=({children,style={}})=>(
    <div style={{background:B.card,border:`1px solid ${B.border}`,borderRadius:10,padding:16,...style}}>{children}</div>
  );

  const ST=({children,sub,action})=>(
    <div style={{marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
      <div>
        <div style={{fontSize:11,fontWeight:700,letterSpacing:".1em",textTransform:"uppercase",color:B.white,display:"flex",alignItems:"center",gap:7}}>
          <span style={{display:"inline-block",width:3,height:12,background:B.red,borderRadius:2,flexShrink:0}}/>
          {children}
        </div>
        {sub&&<div style={{fontSize:10,color:B.gray3,marginTop:2,paddingLeft:10}}>{sub}</div>}
      </div>
      {action}
    </div>
  );

  const MBadge=({v})=>{const c=v<0?B.red:v>25?B.green:B.amber;const bg=v<0?B.redDim:v>25?B.greenDim:B.amberDim;return <span style={{background:bg,color:c,borderRadius:3,padding:"2px 6px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{fP(v)}</span>};
  const SBadge=({s})=>{const c=s==="TOP"?B.amberLt:s==="MED"?B.blueLt:B.gray3;const bg=s==="TOP"?B.amberDim:s==="MED"?B.blueDim:B.gray4;return <span style={{background:bg,color:c,borderRadius:3,padding:"2px 6px",fontSize:9,fontWeight:700,border:`1px solid ${c}33`}}>{s}</span>};

  // ── SIDEBAR ────────────────────────────────────────────────
  const Sidebar=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div>
        <div style={{fontSize:9,color:B.gray3,letterSpacing:".14em",fontWeight:700,marginBottom:5}}>BÚSQUEDA</div>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Cliente, artículo, código…"
          style={{width:"100%",background:B.panel,border:`1px solid ${B.border}`,borderRadius:5,color:B.white,fontFamily:"inherit",fontSize:11,padding:"6px 9px",outline:"none"}}/>
      </div>
      {[
        {lbl:"SUCURSAL", opts:opts.sucs, active:fSuc, set:setFSuc, color:B.red},
        {lbl:"AÑO FISCAL",opts:opts.fys,  active:fFY,  set:setFY,   color:B.blue},
        {lbl:"TRIMESTRE", opts:["1","2","3","4"],labels:["Q1 Abr","Q2 Jul","Q3 Oct","Q4 Ene"],active:fQ,set:setFQ,color:B.amber},
        {lbl:"SECTOR",    opts:opts.secs, active:fSec, set:setFSec, color:B.teal},
        {lbl:"FAMILIA",   opts:opts.fams, active:fFam, set:setFFam, color:B.red},
        {lbl:"EJECUTIVO", opts:opts.reps, active:fRep, set:setFRep, color:B.amber,nm:repNames},
        {lbl:"SEGMENTO",  opts:["TOP","MED","LOW"],active:fSeg,set:setFSeg,color:B.amber},
      ].map(({lbl,opts:o,labels,active,set,color,nm})=>(
        <div key={lbl}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
            <span style={{fontSize:9,color:B.gray3,letterSpacing:".1em",fontWeight:700}}>{lbl}</span>
            {active.length>0&&<button onClick={()=>set([])} style={{background:"none",border:"none",color:B.red,cursor:"pointer",fontSize:9,fontFamily:"inherit",padding:0}}>✕</button>}
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:3,maxHeight:lbl==="SECTOR"||lbl==="FAMILIA"?80:9999,overflowY:lbl==="SECTOR"||lbl==="FAMILIA"?"auto":"visible",paddingBottom:2}}>
            {(o||[]).map((opt,i)=>{
              const display=nm&&nm[opt]?nm[opt].split(" ")[0]:labels?labels[i]:String(opt).length>13?String(opt).slice(0,12)+"…":opt;
              return <Pill key={opt} label={display} active={active.includes(opt)} color={color} onClick={()=>toggle(set,opt)}/>;
            })}
          </div>
        </div>
      ))}
      {hasFilters&&<button onClick={clearAll} style={{width:"100%",background:B.redBg,border:`1px solid ${B.red}44`,color:B.red,borderRadius:6,padding:"6px 10px",fontSize:10,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>✕ Limpiar todos los filtros</button>}
      <div style={{borderTop:`1px solid ${B.border}`,paddingTop:12,display:"flex",flexDirection:"column",gap:6}}>
        <div style={{fontSize:9,color:B.gray3,letterSpacing:".1em",fontWeight:700,marginBottom:2}}>DATOS ACTIVOS</div>
        {[{l:"Líneas",v:fN(useSeed?SEED.meta.lineas:filtered.length),c:B.red},{l:"Clientes",v:fN(useSeed?SEED.meta.clientes:new Set(filtered.map(r=>r.cli)).size),c:B.white},{l:"Ingresos",v:f$(kpi.ing),c:B.amberLt}].map(x=>(
          <div key={x.l} style={{background:B.panel,borderRadius:6,padding:"7px 10px",border:`1px solid ${B.border}`}}>
            <div style={{fontSize:9,color:B.gray3}}>{x.l}</div>
            <div style={{color:x.c,fontWeight:800,fontSize:13,marginTop:1}}>{x.v}</div>
          </div>
        ))}
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════════════
  // VIEWS
  // ══════════════════════════════════════════════════════════

  const ViewResumen=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <KCard icon="💰" label="Ingresos Netos" value={f$(kpi.ing)} sub={`${fN(kpi.lineas)} líneas · ${fN(kpi.clis)} clientes`} color={B.red}/>
        <KCard icon="📊" label="Costo Total" value={f$(kpi.cst)} sub="costo de ventas" color={B.gray3}/>
        <KCard icon="📈" label="Margen Bruto" value={f$(kpi.mgn)} sub={`${kpi.mgnP.toFixed(1)}% sobre ingresos`} color={kpi.mgnP>20?B.green:kpi.mgnP>10?B.amber:B.red} warn={kpi.mgnP<10}/>
        <KCard icon="🏢" label="Clientes Activos" value={kpi.clis} sub={`${fN(kpi.vol)} unidades`} color={B.amber}/>
        <KCard icon="📅" label="Período" value={useSeed?"Abr 2026":fys.slice(-1)[0]||"—"} sub={useSeed?"FY 2026/27 · K27 Arequipa":fys.join(" / ")} color={B.blue}/>
      </div>
      {analysis.alerts.filter(a=>a.lvl==="CRÍTICO").map((a,i)=>(
        <div key={i} style={{background:B.redBg,border:`1px solid ${B.red}`,borderRadius:8,padding:"11px 15px",display:"flex",gap:12,alignItems:"flex-start"}}>
          <span style={{fontSize:20,flexShrink:0}}>{a.icon}</span>
          <div><div style={{color:B.red,fontWeight:800,fontSize:12}}>{a.t}</div><div style={{color:B.gray1,fontSize:11,marginTop:2,lineHeight:1.5}}>{a.b}</div></div>
        </div>
      ))}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
        <Pnl>
          <ST sub={useSeed?"Ingresos por día de facturación (Abril 2026)":"Ingresos por mes fiscal"}>
            {useSeed?"Curva de facturación mensual":"Tendencia por mes fiscal"}
          </ST>
          <ResponsiveContainer width="100%" height={180}>
            {useSeed?(
              <AreaChart data={SEED.byday}>
                <defs><linearGradient id="rg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={B.red} stopOpacity={.35}/><stop offset="95%" stopColor={B.red} stopOpacity={0}/></linearGradient></defs>
                <CartesianGrid strokeDasharray="3 3" stroke={B.border}/>
                <XAxis dataKey="dia" stroke={B.gray3} tick={{fontSize:10}} tickFormatter={d=>`Día ${d}`}/>
                <YAxis stroke={B.gray3} tick={{fontSize:9}} tickFormatter={f$} width={60}/>
                <Tooltip content={<TT/>}/>
                <Area type="monotone" dataKey="ing" stroke={B.red} fill="url(#rg)" strokeWidth={2} name="Ingresos"/>
              </AreaChart>
            ):(
              <AreaChart data={FM.map((m,i)=>({mes:m,ing:filtered.filter(r=>r.fmi===i).reduce((s,r)=>s+r.ing,0)}))}>
                <defs><linearGradient id="rg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={B.red} stopOpacity={.35}/><stop offset="95%" stopColor={B.red} stopOpacity={0}/></linearGradient></defs>
                <CartesianGrid strokeDasharray="3 3" stroke={B.border}/>
                <XAxis dataKey="mes" stroke={B.gray3} tick={{fontSize:10}}/>
                <YAxis stroke={B.gray3} tick={{fontSize:9}} tickFormatter={f$} width={60}/>
                <Tooltip content={<TT/>}/>
                <Area type="monotone" dataKey="ing" stroke={B.red} fill="url(#rg)" strokeWidth={2} name="Ingresos"/>
              </AreaChart>
            )}
          </ResponsiveContainer>
        </Pnl>
        <Pnl>
          <ST sub="Concentración TOP/MED/LOW (click para filtrar)">Segmentación Pareto</ST>
          {(()=>{
            const topC=clients.filter(c=>c.seg==="TOP");
            const medC=clients.filter(c=>c.seg==="MED");
            const lowC=clients.filter(c=>c.seg==="LOW");
            const topPct=topC.reduce((s,c)=>s+(c.pct||0),0);
            const medPct=medC.reduce((s,c)=>s+(c.pct||0),0);
            const lowPct=lowC.reduce((s,c)=>s+(c.pct||0),0);
            return [{s:"TOP",n:topC.length,pct:topPct,c:B.amberLt,cs:topC},{s:"MED",n:medC.length,pct:medPct,c:B.red,cs:medC},{s:"LOW",n:lowC.length,pct:lowPct,c:B.gray3,cs:lowC}].map(seg=>(
              <div key={seg.s} style={{marginBottom:10,cursor:"pointer"}} onClick={()=>toggle(setFSeg,seg.s)}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
                  <span style={{color:seg.c,fontWeight:700}}>{seg.s} · {seg.n} clientes</span>
                  <span style={{color:B.white}}>{seg.pct.toFixed(1)}%</span>
                </div>
                <div style={{height:7,background:B.border,borderRadius:4,overflow:"hidden"}}>
                  <div style={{height:"100%",width:`${clamp(seg.pct,0,100)}%`,background:seg.c,borderRadius:4,transition:"width .5s"}}/>
                </div>
                <div style={{display:"flex",gap:4,flexWrap:"wrap",marginTop:4}}>
                  {seg.cs.slice(0,3).map(c=>(
                    <span key={c.id} style={{fontSize:9,color:B.gray3,background:B.panel,borderRadius:3,padding:"1px 5px",border:`1px solid ${B.border}`}}>{c.nm.split(" ").slice(0,2).join(" ")}</span>
                  ))}
                </div>
              </div>
            ));
          })()}
        </Pnl>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
        <Pnl>
          <ST sub="% del ingreso total por industria">Mix sectorial</ST>
          <ResponsiveContainer width="100%" height={200}>
            <PieChart>
              <Pie data={sector.filter(s=>s.ing>0)} dataKey="ing" nameKey="sec" cx="50%" cy="50%" outerRadius={80} innerRadius={38} label={({percent})=>`${(percent*100).toFixed(0)}%`} labelLine={false}>
                {sector.filter(s=>s.ing>0).map((_,i)=><Cell key={i} fill={B.chart[i%B.chart.length]}/>)}
              </Pie>
              <Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:9}}/>
            </PieChart>
          </ResponsiveContainer>
        </Pnl>
        <Pnl>
          <ST sub="Ingresos y margen por familia de producto">Rentabilidad por línea</ST>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={familia.filter(f=>f.ing>0).slice(0,8)} layout="vertical" margin={{left:0,right:20,top:4,bottom:4}}>
              <CartesianGrid strokeDasharray="3 3" stroke={B.border} horizontal={false}/>
              <XAxis type="number" stroke={B.gray3} tick={{fontSize:9}} tickFormatter={f$} width={60}/>
              <YAxis type="category" dataKey="f2" stroke={B.gray3} tick={{fontSize:9}} width={140}/>
              <Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/>
              <Bar dataKey="ing" name="Ingresos" fill={B.red} radius={[0,3,3,0]}/>
              <Bar dataKey="mgn" name="Margen" radius={[0,3,3,0]}>
                {familia.filter(f=>f.ing>0).slice(0,8).map((f,i)=><Cell key={i} fill={f.mgnP<0?B.gray4:B.greenLt}/>)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Pnl>
      </div>
    </div>
  );

  const ViewClientes=()=>(
    <div style={{display:"grid",gridTemplateColumns:selCli?"1fr 350px":"1fr",gap:14,alignItems:"start"}}>
      <Pnl style={{overflow:"auto"}}>
        <ST sub={`${filteredClients.length} clientes · click para perfil completo`}>Cartera de clientes</ST>
        <div style={{display:"flex",gap:5,marginBottom:10,flexWrap:"wrap"}}>
          {["ALL","TOP","MED","LOW"].map(s=>(
            <Pill key={s} label={s==="ALL"?"Todos los segmentos":s} active={s==="ALL"?!fSeg.length:fSeg.includes(s)} color={s==="TOP"?B.amberLt:s==="MED"?B.blueLt:B.gray3} onClick={()=>s==="ALL"?setFSeg([]):toggle(setFSeg,s)}/>
          ))}
        </div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,tableLayout:"fixed"}}>
          <colgroup>
            <col style={{width:28}}/><col style={{width:"26%"}}/><col style={{width:"13%"}}/>
            <col style={{width:78}}/><col style={{width:74}}/><col style={{width:58}}/>
            <col style={{width:36}}/><col style={{width:42}}/><col/>
          </colgroup>
          <thead><tr style={{borderBottom:`1px solid ${B.border}`}}>
            {["#","Cliente","Sector","Ingresos","Margen","Mgn%","Lín.","Seg.","Ejecutivo"].map(h=>(
              <th key={h} style={{padding:"6px 8px",textAlign:h==="Ingresos"||h==="Margen"?"right":"left",color:B.gray3,fontSize:9,letterSpacing:".07em",textTransform:"uppercase"}}>{h}</th>
            ))}
          </tr></thead>
          <tbody>
            {filteredClients.map((c,i)=>(
              <tr key={c.id} onClick={()=>setSelCli(selCli?.id===c.id?null:c)}
                style={{cursor:"pointer",borderBottom:`1px solid ${B.border}22`,
                  background:selCli?.id===c.id?`${B.red}12`:"transparent",
                  borderLeft:selCli?.id===c.id?`2px solid ${B.red}`:"2px solid transparent"}}>
                <td style={{padding:"7px 8px",color:B.gray3,fontSize:9}}>{i+1}</td>
                <td style={{padding:"7px 8px",overflow:"hidden"}}>
                  <div style={{fontWeight:700,color:B.white,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.nm}</div>
                  <div style={{fontSize:9,color:B.gray3}}>{c.id}</div>
                </td>
                <td style={{padding:"7px 8px",color:B.gray3,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.sector}</td>
                <td style={{padding:"7px 8px",fontWeight:700,color:B.red,textAlign:"right"}}>{f$(c.ing)}</td>
                <td style={{padding:"7px 8px",color:c.mgn>=0?B.greenLt:B.red,fontWeight:600,textAlign:"right"}}>{f$(c.mgn)}</td>
                <td style={{padding:"7px 8px",textAlign:"center"}}><MBadge v={c.mgnP}/></td>
                <td style={{padding:"7px 8px",color:B.gray3,textAlign:"center"}}>{c.lineas}</td>
                <td style={{padding:"7px 8px",textAlign:"center"}}><SBadge s={c.seg}/></td>
                <td style={{padding:"7px 8px",color:B.amberLt,fontSize:10,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  {(()=>{const k=Object.keys(cartera).find(k=>cartera[k]?.includes(String(c.id)));return k?`★ ${repNames[k]||k}`:repNames[c.rep]||c.rep||"—"})()}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Pnl>

      {selCli&&(
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          <Pnl style={{borderTop:`2px solid ${selCli.seg==="TOP"?B.amberLt:selCli.seg==="MED"?B.blueLt:B.gray3}`}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
              <div>
                <div style={{fontWeight:800,fontSize:13,color:B.white,lineHeight:1.3}}>{selCli.nm}</div>
                <div style={{fontSize:10,color:B.gray3,marginTop:2}}>{selCli.id} · {selCli.sector}</div>
              </div>
              <button onClick={()=>setSelCli(null)} style={{background:"none",border:"none",color:B.gray3,cursor:"pointer",fontSize:16,lineHeight:1,padding:4}}>✕</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7}}>
              {[{l:"Ingresos",v:f$(selCli.ing),c:B.red},{l:"Margen",v:f$(selCli.mgn),c:selCli.mgn>=0?B.greenLt:B.red},
                {l:"Margen %",v:fP(selCli.mgnP),c:selCli.mgnP<0?B.red:selCli.mgnP>25?B.green:B.amber},
                {l:"% del total",v:`${selCli.pct.toFixed(1)}%`,c:B.amberLt},
                {l:"Líneas SAP",v:selCli.lineas},{l:"Volumen",v:`${fN(selCli.vol)} u`},
              ].map(({l,v,c})=>(
                <div key={l} style={{background:B.panel,borderRadius:6,padding:"8px 10px",border:`1px solid ${B.border}`}}>
                  <div style={{fontSize:9,color:B.gray3,marginBottom:2}}>{l}</div>
                  <div style={{fontSize:13,fontWeight:800,color:c||B.white}}>{v}</div>
                </div>
              ))}
            </div>
            {selCli.mgn<0&&<div style={{marginTop:8,padding:"6px 10px",background:B.redBg,border:`1px solid ${B.red}`,borderRadius:5,fontSize:10,color:B.red}}>⚠️ Margen negativo — revisar condición comercial y descuentos</div>}
            {selCli.pct>20&&<div style={{marginTop:6,padding:"6px 10px",background:B.amberDim,border:`1px solid ${B.amber}`,borderRadius:5,fontSize:10,color:B.amberLt}}>🔑 Cliente estratégico: {selCli.pct.toFixed(1)}% del ingreso total</div>}
          </Pnl>

          <Pnl>
            <ST sub="Ingresos por mes fiscal">Historial del cliente</ST>
            <ResponsiveContainer width="100%" height={95}>
              <AreaChart data={cliMonthly}>
                <defs><linearGradient id="cg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={B.red} stopOpacity={.3}/><stop offset="95%" stopColor={B.red} stopOpacity={0}/></linearGradient></defs>
                <XAxis dataKey="mes" stroke={B.gray3} tick={{fontSize:8}}/><Tooltip content={<TT/>}/>
                <Area type="monotone" dataKey="ing" stroke={B.red} fill="url(#cg)" strokeWidth={2} name="Ingresos" dot={false}/>
              </AreaChart>
            </ResponsiveContainer>
          </Pnl>

          <Pnl>
            <ST sub="Sobreescribe la cartera de SAP — guardado localmente">Asignar ejecutivo</ST>
            <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
              {opts.reps.map(rep=>{
                const assigned=cartera[rep]?.includes(String(selCli.id));
                return(
                  <button key={rep} onClick={()=>setCartera(prev=>{
                    const n={...prev};if(!n[rep])n[rep]=[];
                    if(assigned)n[rep]=n[rep].filter(id=>id!==String(selCli.id));
                    else{Object.keys(n).forEach(r=>{n[r]=(n[r]||[]).filter(id=>id!==String(selCli.id));});n[rep]=[...n[rep],String(selCli.id)];}
                    return n;
                  })} style={{background:assigned?B.amberDim:"transparent",border:`1px solid ${assigned?B.amberLt:B.border}`,color:assigned?B.amberLt:B.gray3,borderRadius:5,padding:"4px 10px",fontSize:10,cursor:"pointer",fontFamily:"inherit",fontWeight:assigned?700:400}}>
                    {assigned?"★ ":""}{repNames[rep]||`Rep ${rep}`}
                  </button>
                );
              })}
            </div>
          </Pnl>
        </div>
      )}
    </div>
  );

  const ViewProductos=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
        <Pnl>
          <ST sub="Click en familia para filtrar">Familias de producto</ST>
          {familia.filter(f=>Math.abs(f.ing)>0).map((f,i)=>(
            <div key={f.f2} style={{marginBottom:9,paddingLeft:9,borderLeft:`2px solid ${f.mgnP<0?B.red:B.chart[i%B.chart.length]}`,cursor:"pointer"}} onClick={()=>toggle(setFFam,f.f2)}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:2,gap:8}}>
                <span style={{color:B.white,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>{f.f2}</span>
                <div style={{display:"flex",gap:7,flexShrink:0}}>
                  <span style={{color:B.red,fontWeight:700}}>{f$(f.ing)}</span>
                  <MBadge v={f.mgnP}/>
                </div>
              </div>
              <div style={{height:4,background:B.border,borderRadius:2,overflow:"hidden"}}>
                <div style={{height:"100%",width:`${familia[0]?.ing?clamp(f.ing/familia[0].ing*100,0,100):0}%`,background:f.mgnP<0?B.red:B.chart[i%B.chart.length],borderRadius:2}}/>
              </div>
            </div>
          ))}
        </Pnl>

        <Pnl style={{overflow:"auto"}}>
          <ST sub={`Top ${articles.length} artículos por ingreso`}>Artículos más vendidos</ST>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,tableLayout:"fixed"}}>
            <colgroup><col style={{width:28}}/><col style={{width:"30%"}}/><col/><col style={{width:72}}/><col style={{width:58}}/><col style={{width:38}}/></colgroup>
            <thead><tr style={{borderBottom:`1px solid ${B.border}`}}>
              {["#","Código","Sub-familia","Ingresos","Mgn%","Vol"].map(h=>(
                <th key={h} style={{padding:"5px 7px",textAlign:h==="Ingresos"?"right":"left",color:B.gray3,fontSize:9,letterSpacing:".07em",textTransform:"uppercase"}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {articles.map((a,i)=>(
                <tr key={a.art} style={{borderBottom:`1px solid ${B.border}22`}}>
                  <td style={{padding:"6px 7px",color:B.gray3,fontSize:9}}>{i+1}</td>
                  <td style={{padding:"6px 7px",color:B.red,fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.art}</td>
                  <td style={{padding:"6px 7px",color:B.gray3,fontSize:10,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.f2||a.f1}</td>
                  <td style={{padding:"6px 7px",fontWeight:700,color:B.white,textAlign:"right"}}>{f$(a.ing)}</td>
                  <td style={{padding:"6px 7px",textAlign:"center"}}><MBadge v={a.mgnP}/></td>
                  <td style={{padding:"6px 7px",color:B.gray3,textAlign:"right"}}>{a.vol}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </Pnl>
      </div>

      <Pnl>
        <ST sub="Ingresos vs Margen por familia">Comparativa rentabilidad</ST>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={familia.filter(f=>f.ing>0)} margin={{top:5,right:20,left:0,bottom:50}}>
            <CartesianGrid strokeDasharray="3 3" stroke={B.border}/>
            <XAxis dataKey="f2" stroke={B.gray3} tick={{fontSize:9,angle:-20,textAnchor:"end"}} interval={0} height={64}/>
            <YAxis stroke={B.gray3} tick={{fontSize:9}} tickFormatter={f$} width={62}/>
            <Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/>
            <Bar dataKey="ing" name="Ingresos" fill={B.red} radius={[4,4,0,0]}/>
            <Bar dataKey="mgn" name="Margen" radius={[4,4,0,0]}>
              {familia.filter(f=>f.ing>0).map((f,i)=><Cell key={i} fill={f.mgnP<0?B.gray4:B.greenLt}/>)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </Pnl>

      <Pnl>
        <ST sub="Clasificación ABC de artículos (A=alto movimiento)">Análisis ABC</ST>
        <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
          {(useSeed?SEED.abc:[]).map((a,i)=>(
            <div key={a.ind} style={{flex:"1 1 120px",background:B.panel,borderRadius:8,padding:"12px 14px",border:`1px solid ${B.border}`,borderTop:`2px solid ${B.chart[i]}`}}>
              <div style={{fontSize:22,fontWeight:900,color:B.chart[i],lineHeight:1}}>{a.ind}</div>
              <div style={{fontSize:15,fontWeight:800,color:B.white,marginTop:6}}>{f$(a.ing)}</div>
              <div style={{fontSize:10,color:B.gray3,marginTop:2}}>{a.lineas} líneas · {fN(a.vol)} uds</div>
            </div>
          ))}
        </div>
      </Pnl>
    </div>
  );

  const ViewSectores=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
        <Pnl>
          <ST sub="Click para filtrar por sector">Ingresos por industria</ST>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={sector.filter(s=>s.ing>0)} margin={{top:5,right:20,left:0,bottom:50}}>
              <CartesianGrid strokeDasharray="3 3" stroke={B.border}/>
              <XAxis dataKey="sec" stroke={B.gray3} tick={{fontSize:9,angle:-20,textAnchor:"end"}} interval={0} height={64}/>
              <YAxis stroke={B.gray3} tick={{fontSize:9}} tickFormatter={f$} width={62}/>
              <Tooltip content={<TT/>}/>
              <Bar dataKey="ing" name="Ingresos" radius={[4,4,0,0]} onClick={d=>d&&toggle(setFSec,d.sec)}>
                {sector.filter(s=>s.ing>0).map((_,i)=><Cell key={i} fill={B.chart[i%B.chart.length]} cursor="pointer"/>)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Pnl>
        <Pnl>
          <ST>Mix de participación</ST>
          <ResponsiveContainer width="100%" height={240}>
            <PieChart>
              <Pie data={sector.filter(s=>s.ing>0)} dataKey="ing" nameKey="sec" cx="50%" cy="50%" outerRadius={95} innerRadius={50} label={({name,percent})=>`${(percent*100).toFixed(0)}%`} labelLine={{stroke:B.border}}>
                {sector.filter(s=>s.ing>0).map((_,i)=><Cell key={i} fill={B.chart[i%B.chart.length]}/>)}
              </Pie>
              <Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:9}}/>
            </PieChart>
          </ResponsiveContainer>
        </Pnl>
      </div>
      {sector.filter(s=>s.ing>0).map((s,si)=>{
        const cs=clients.filter(c=>c.sector===s.sec&&c.ing>0);
        if(!cs.length)return null;
        return(
          <Pnl key={si} style={{borderLeft:`3px solid ${B.chart[si%B.chart.length]}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,cursor:"pointer"}} onClick={()=>toggle(setFSec,s.sec)}>
              <div style={{color:B.chart[si%B.chart.length],fontWeight:800,fontSize:12}}>{s.sec}</div>
              <div style={{display:"flex",gap:14,fontSize:11}}>
                <span style={{color:B.red,fontWeight:700}}>{f$(s.ing)}</span>
                <span style={{color:B.gray3}}>{s.clientes} clientes · {s.lineas} líneas</span>
              </div>
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
              {cs.map(c=>(
                <div key={c.id} onClick={e=>{e.stopPropagation();setSelCli(c);setView("CLIENTES");}} style={{background:B.panel,borderRadius:7,padding:"6px 10px",border:`1px solid ${B.border}`,cursor:"pointer"}}
                  onMouseEnter={e=>e.currentTarget.style.borderColor=B.red} onMouseLeave={e=>e.currentTarget.style.borderColor=B.border}>
                  <div style={{fontSize:10,color:B.white,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:160}}>{c.nm}</div>
                  <div style={{display:"flex",gap:7,marginTop:2,alignItems:"center"}}>
                    <span style={{fontSize:10,color:B.red,fontWeight:700}}>{f$(c.ing)}</span>
                    <MBadge v={c.mgnP}/><SBadge s={c.seg}/>
                  </div>
                </div>
              ))}
            </div>
          </Pnl>
        );
      })}
    </div>
  );

  const ViewTendencias=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      {fys.length>1?(
        <>
          <Pnl>
            <ST sub="Comparativa mes a mes · Año Fiscal Abr→Mar">Evolución fiscal mensual</ST>
            <ResponsiveContainer width="100%" height={240}>
              <LineChart data={fiscalTrend} margin={{top:5,right:20,left:0,bottom:5}}>
                <CartesianGrid strokeDasharray="3 3" stroke={B.border}/>
                <XAxis dataKey="mes" stroke={B.gray3} tick={{fontSize:10}}/>
                <YAxis stroke={B.gray3} tick={{fontSize:9}} tickFormatter={f$} width={62}/>
                <Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/>
                {fys.map((fy,i)=><Line key={fy} type="monotone" dataKey={fy} stroke={B.chart[i%B.chart.length]} strokeWidth={2} dot={{r:3}} connectNulls/>)}
              </LineChart>
            </ResponsiveContainer>
          </Pnl>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
            <Pnl>
              <ST sub="Q1=Abr-Jun · Q2=Jul-Sep · Q3=Oct-Dic · Q4=Ene-Mar">Comparativa trimestral</ST>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={quarterData} margin={{top:5,right:20,left:0,bottom:5}}>
                  <CartesianGrid strokeDasharray="3 3" stroke={B.border}/>
                  <XAxis dataKey="q" stroke={B.gray3} tick={{fontSize:10}}/>
                  <YAxis stroke={B.gray3} tick={{fontSize:9}} tickFormatter={f$} width={62}/>
                  <Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/>
                  {fys.map((fy,i)=><Bar key={fy} dataKey={fy} fill={B.chart[i%B.chart.length]} radius={[4,4,0,0]}/>)}
                </BarChart>
              </ResponsiveContainer>
            </Pnl>
            <Pnl>
              <ST sub="YoY por año fiscal">Totales anuales</ST>
              {fyTotals.map((fy,i)=>{
                const prev=fyTotals[i-1];const d=prev?(fy.ing-prev.ing)/prev.ing*100:null;
                return(
                  <div key={fy.fy} style={{marginBottom:10,background:B.panel,borderRadius:8,padding:"11px 14px",border:`1px solid ${B.border}`,borderLeft:`3px solid ${B.chart[i%B.chart.length]}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <span style={{color:B.chart[i%B.chart.length],fontWeight:800,fontSize:12}}>{fy.fy}</span>
                      {d!=null&&<span style={{color:d>=0?B.greenLt:B.red,fontWeight:700,fontSize:11}}>{d>=0?"▲":"▼"} {Math.abs(d).toFixed(1)}% YoY</span>}
                    </div>
                    <div style={{display:"flex",gap:20,marginTop:6}}>
                      <div><div style={{fontSize:9,color:B.gray3}}>INGRESOS</div><div style={{fontSize:14,fontWeight:800,color:B.red}}>{f$(fy.ing)}</div></div>
                      <div><div style={{fontSize:9,color:B.gray3}}>MARGEN</div><div style={{fontSize:14,fontWeight:800,color:fy.mgnP>15?B.greenLt:B.amberLt}}>{fP(fy.mgnP)}</div></div>
                      <div><div style={{fontSize:9,color:B.gray3}}>CLIENTES</div><div style={{fontSize:14,fontWeight:800,color:B.white}}>{fy.clis}</div></div>
                    </div>
                  </div>
                );
              })}
            </Pnl>
          </div>
        </>
      ):(
        <Pnl style={{textAlign:"center",padding:"48px 32px"}}>
          <div style={{fontSize:36,marginBottom:12}}>📂</div>
          <div style={{color:B.white,fontWeight:700,fontSize:14,marginBottom:8}}>Carga más Excels para activar comparativas</div>
          <div style={{color:B.gray3,fontSize:12,lineHeight:1.7}}>Agrega archivos de otros meses o sucursales.<br/>Nombres sugeridos: <code style={{color:B.red}}>2025_04_K27.xlsx</code> · <code style={{color:B.red}}>2026_01_K11.xlsx</code></div>
        </Pnl>
      )}
    </div>
  );

  const ViewEjecutivos=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <Pnl>
        <ST sub="Asigna nombres reales a los códigos SAP · guardado en tu navegador">Configurar ejecutivos</ST>
        <div style={{display:"flex",flexWrap:"wrap",gap:10}}>
          {opts.reps.map(rep=>(
            <div key={rep} style={{display:"flex",alignItems:"center",gap:8,background:B.panel,borderRadius:7,padding:"7px 12px",border:`1px solid ${B.border}`}}>
              <span style={{color:B.gray3,fontSize:10,minWidth:52,flexShrink:0}}>Rep {rep}</span>
              <span style={{color:B.border}}>→</span>
              <input value={repNames[rep]||""} placeholder="Nombre…" onChange={e=>setRepNames(p=>({...p,[rep]:e.target.value}))}
                style={{background:"transparent",border:"none",borderBottom:`1px solid ${B.border}`,color:B.amberLt,fontFamily:"inherit",fontSize:11,padding:"2px 4px",outline:"none",width:155}}/>
            </div>
          ))}
        </div>
      </Pnl>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
        <Pnl>
          <ST sub="Click para filtrar por ejecutivo">Ingresos por ejecutivo</ST>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={repsAgg.map(r=>({...r,nm:repNames[r.rep]||`Rep ${r.rep}`}))} margin={{top:5,right:20,left:0,bottom:50}}
              onClick={d=>d?.activePayload&&toggle(setFRep,d.activePayload[0]?.payload?.rep)}>
              <CartesianGrid strokeDasharray="3 3" stroke={B.border}/>
              <XAxis dataKey="nm" stroke={B.gray3} tick={{fontSize:9,angle:-15,textAnchor:"end"}} interval={0} height={64}/>
              <YAxis stroke={B.gray3} tick={{fontSize:9}} tickFormatter={f$} width={62}/>
              <Tooltip content={<TT/>}/>
              <Bar dataKey="ing" name="Ingresos" radius={[4,4,0,0]} cursor="pointer">
                {repsAgg.map((_,i)=><Cell key={i} fill={B.chart[i%B.chart.length]}/>)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Pnl>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {repsAgg.map((r,i)=>(
            <div key={r.rep} style={{background:B.panel,borderRadius:9,padding:"11px 14px",border:`1px solid ${B.border}`,borderLeft:`3px solid ${B.chart[i%B.chart.length]}`,cursor:"pointer"}} onClick={()=>toggle(setFRep,r.rep)}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
                <div style={{fontWeight:800,color:B.chart[i%B.chart.length],fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{repNames[r.rep]||`Rep ${r.rep}`}</div>
                <div style={{color:B.red,fontWeight:700,fontSize:12,flexShrink:0,marginLeft:8}}>{f$(r.ing)}</div>
              </div>
              <div style={{display:"flex",gap:14,fontSize:10,color:B.gray3}}>
                <span>{r.clientes} clientes</span>
                <span>{r.lineas} líneas</span>
                <span>{(r.ing/kpi.ing*100).toFixed(1)}% del total</span>
              </div>
              <div style={{marginTop:7,height:3,background:B.border,borderRadius:2}}>
                <div style={{height:"100%",width:`${repsAgg[0]?.ing?clamp(r.ing/repsAgg[0].ing*100,0,100):0}%`,background:B.chart[i%B.chart.length],borderRadius:2}}/>
              </div>
              <div style={{marginTop:6,display:"flex",flexWrap:"wrap",gap:4}}>
                {clients.filter(c=>c.rep===r.rep||cartera[r.rep]?.includes(String(c.id))).slice(0,5).map(c=>(
                  <span key={c.id} style={{fontSize:9,color:B.gray3,background:B.card,borderRadius:3,padding:"1px 5px",border:`1px solid ${B.border}`,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:100}}>{c.nm.split(" ").slice(0,2).join(" ")}</span>
                ))}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );

  const ViewPivot=()=>{
    const {rows:pivRows,cols:pivCols}=pivotData;
    const vFmt=PIVOT_VALS_META[pivotVal]?.fmt||fN;
    const grandTotal=pivRows.reduce((s,r)=>s+r.total,0);
    return(
      <div style={{display:"flex",flexDirection:"column",gap:14}}>
        {/* Config bar */}
        <Pnl>
          <ST sub="Configura filas, columnas y valor — se actualiza en tiempo real">Constructor de tabla dinámica</ST>
          <div style={{display:"flex",gap:16,flexWrap:"wrap",alignItems:"flex-start"}}>
            {[{label:"FILAS",val:pivotRows,set:setPivotRows},{label:"COLUMNAS",val:pivotCols,set:setPivotCols}].map(({label,val,set})=>(
              <div key={label}>
                <div style={{fontSize:9,color:B.gray3,letterSpacing:".1em",fontWeight:700,marginBottom:6}}>{label}</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                  {Object.entries(PIVOT_DIMS_META).map(([k,d])=>(
                    <Pill key={k} label={d} active={val===k} color={B.blue} onClick={()=>set(k)}/>
                  ))}
                </div>
              </div>
            ))}
            <div>
              <div style={{fontSize:9,color:B.gray3,letterSpacing:".1em",fontWeight:700,marginBottom:6}}>VALOR</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                {Object.entries(PIVOT_VALS_META).map(([k,d])=>(
                  <Pill key={k} label={d.label} active={pivotVal===k} color={B.red} onClick={()=>setPivotVal(k)}/>
                ))}
              </div>
            </div>
            <div>
              <div style={{fontSize:9,color:B.gray3,letterSpacing:".1em",fontWeight:700,marginBottom:6}}>ORDEN</div>
              <div style={{display:"flex",gap:4}}>
                <Pill label="Mayor → Menor" active={pivotSort==="desc"} color={B.amber} onClick={()=>setPivotSort("desc")}/>
                <Pill label="Menor → Mayor" active={pivotSort==="asc"} color={B.amber} onClick={()=>setPivotSort("asc")}/>
              </div>
            </div>
          </div>
        </Pnl>

        {/* Pivot table */}
        <Pnl style={{overflow:"auto"}}>
          <ST sub={`${pivRows.length} filas · ${pivCols.length} columnas · ${PIVOT_VALS_META[pivotVal]?.label}`}>
            {PIVOT_DIMS_META[pivotRows]||"—"} × {PIVOT_DIMS_META[pivotCols]||"—"}
          </ST>
          {pivRows.length===0?(
            <div style={{color:B.gray3,fontSize:12,padding:"20px 0",textAlign:"center"}}>Sin datos con los filtros actuales</div>
          ):(
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:500}}>
              <thead>
                <tr style={{borderBottom:`2px solid ${B.red}`}}>
                  <th style={{padding:"7px 10px",textAlign:"left",color:B.gray3,fontSize:9,letterSpacing:".08em",textTransform:"uppercase",background:B.panel,position:"sticky",left:0,zIndex:1,minWidth:160}}>{PIVOT_DIMS_META[pivotRows]||'—'}</th>
                  {pivCols.map(c=>(
                    <th key={c} style={{padding:"7px 10px",textAlign:"right",color:B.gray3,fontSize:9,letterSpacing:".07em",textTransform:"uppercase",background:B.panel,whiteSpace:"nowrap"}}>{c}</th>
                  ))}
                  <th style={{padding:"7px 10px",textAlign:"right",color:B.red,fontSize:9,letterSpacing:".07em",textTransform:"uppercase",background:B.panel,whiteSpace:"nowrap"}}>TOTAL</th>
                  <th style={{padding:"7px 10px",textAlign:"right",color:B.gray3,fontSize:9,background:B.panel}}>%</th>
                </tr>
              </thead>
              <tbody>
                {pivRows.map((row,ri)=>{
                  const rowPct=grandTotal?row.total/grandTotal*100:0;
                  return(
                    <tr key={ri} style={{borderBottom:`1px solid ${B.border}22`}} onMouseEnter={e=>e.currentTarget.style.background=`${B.red}08`} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                      <td style={{padding:"7px 10px",fontWeight:600,color:B.white,background:B.panel,position:"sticky",left:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:200}}>{row.rowKey}</td>
                      {pivCols.map(c=>{
                        const v=row.vals[c]||0;
                        const colTotal=pivRows.reduce((s,r)=>s+(r.vals[c]||0),0);
                        const heat=colTotal?v/colTotal:0;
                        return(
                          <td key={c} style={{padding:"7px 10px",textAlign:"right",fontWeight:v>0?600:400,color:v>0?B.white:B.gray4,background:v>0?`rgba(204,0,0,${heat*0.18})`:"transparent",whiteSpace:"nowrap"}}>{v?vFmt(v):"—"}</td>
                        );
                      })}
                      <td style={{padding:"7px 10px",textAlign:"right",fontWeight:800,color:B.red,whiteSpace:"nowrap"}}>{vFmt(row.total)}</td>
                      <td style={{padding:"7px 10px",textAlign:"right"}}>
                        <div style={{display:"flex",alignItems:"center",gap:6,justifyContent:"flex-end"}}>
                          <div style={{width:40,height:5,background:B.border,borderRadius:3,overflow:"hidden"}}>
                            <div style={{height:"100%",width:`${clamp(rowPct,0,100)}%`,background:B.red,borderRadius:3}}/>
                          </div>
                          <span style={{color:B.gray2,fontSize:10,minWidth:36,textAlign:"right"}}>{rowPct.toFixed(1)}%</span>
                        </div>
                      </td>
                    </tr>
                  );
                })}
                {/* Grand total row */}
                <tr style={{borderTop:`2px solid ${B.red}`,background:B.panel}}>
                  <td style={{padding:"8px 10px",fontWeight:800,color:B.red,position:"sticky",left:0,background:B.panel,fontSize:11}}>TOTAL GENERAL</td>
                  {pivCols.map(c=>{
                    const colTotal=pivRows.reduce((s,r)=>s+(r.vals[c]||0),0);
                    return <td key={c} style={{padding:"8px 10px",textAlign:"right",fontWeight:800,color:B.amberLt,whiteSpace:"nowrap"}}>{colTotal?vFmt(colTotal):"—"}</td>;
                  })}
                  <td style={{padding:"8px 10px",textAlign:"right",fontWeight:900,color:B.red,fontSize:12,whiteSpace:"nowrap"}}>{vFmt(grandTotal)}</td>
                  <td style={{padding:"8px 10px",textAlign:"right",color:B.gray3,fontSize:10}}>100%</td>
                </tr>
              </tbody>
            </table>
          )}
        </Pnl>

        {/* Pivot chart */}
        {pivRows.length>0&&pivCols.length<=4&&(
          <Pnl>
            <ST sub="Visualización de la tabla dinámica">Gráfico de la tabla</ST>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={pivRows.slice(0,12).map(r=>({name:r.rowKey.length>18?r.rowKey.slice(0,17)+"…":r.rowKey,...Object.fromEntries(pivCols.map(c=>[c,r.vals[c]||0]))}))} margin={{top:5,right:20,left:0,bottom:60}}>
                <CartesianGrid strokeDasharray="3 3" stroke={B.border}/>
                <XAxis dataKey="name" stroke={B.gray3} tick={{fontSize:9,angle:-20,textAnchor:"end"}} interval={0} height={64}/>
                <YAxis stroke={B.gray3} tick={{fontSize:9}} tickFormatter={v=>PIVOT_VALS_META[pivotVal]?.fmt===f$?f$(v):fN(v)} width={64}/>
                <Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/>
                {pivCols.map((c,i)=><Bar key={c} dataKey={c} fill={B.chart[i%B.chart.length]} radius={[3,3,0,0]}/>)}
              </BarChart>
            </ResponsiveContainer>
          </Pnl>
        )}
      </div>
    );
  };

  const ViewAnalisis=()=>{
    const{alerts,opps,actions}=analysis;
    return(
      <div style={{display:"flex",flexDirection:"column",gap:14}}>
        <div style={{background:`linear-gradient(135deg,${B.redBg} 0%,${B.panel} 100%)`,border:`1px solid ${B.red}44`,borderRadius:10,padding:"16px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",gap:14,flexWrap:"wrap"}}>
          <div>
            <div style={{fontWeight:900,fontSize:15,color:B.white,letterSpacing:".04em"}}>ANÁLISIS EJECUTIVO AUTOMÁTICO</div>
            <div style={{fontSize:10,color:B.gray3,marginTop:3}}>Motor integrado · sin IA externa · actualizado con los filtros activos · {new Date().toLocaleDateString("es-PE")}</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:22,fontWeight:900,color:B.red}}>{f$(kpi.ing)}</div>
            <div style={{fontSize:11,color:kpi.mgnP>20?B.greenLt:kpi.mgnP>10?B.amberLt:B.red,fontWeight:700}}>{fP(kpi.mgnP)} margen bruto</div>
          </div>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
          <Pnl style={{borderLeft:`3px solid ${B.red}`}}>
            <ST>🔴 Alertas críticas</ST>
            {alerts.length===0?<div style={{color:B.greenLt,fontSize:11}}>✅ Sin alertas con los filtros actuales</div>:
            alerts.map((a,i)=>(
              <div key={i} style={{marginBottom:12,paddingLeft:9,borderLeft:`2px solid ${B.red}44`}}>
                <div style={{display:"flex",gap:6,alignItems:"flex-start",marginBottom:3}}>
                  <span style={{fontSize:13,flexShrink:0}}>{a.icon}</span>
                  <span style={{fontWeight:700,fontSize:11,color:B.white,lineHeight:1.3}}>{a.t}</span>
                </div>
                <div style={{fontSize:10,color:B.gray2,lineHeight:1.6,paddingLeft:19}}>{a.b}</div>
              </div>
            ))}
          </Pnl>
          <Pnl style={{borderLeft:`3px solid ${B.green}`}}>
            <ST>🟢 Oportunidades detectadas</ST>
            {opps.length===0?<div style={{color:B.gray3,fontSize:11}}>Sin oportunidades detectadas actualmente</div>:
            opps.map((o,i)=>(
              <div key={i} style={{marginBottom:12,paddingLeft:9,borderLeft:`2px solid ${B.green}44`}}>
                <div style={{display:"flex",gap:6,alignItems:"flex-start",marginBottom:3}}>
                  <span style={{fontSize:13,flexShrink:0}}>{o.icon}</span>
                  <span style={{fontWeight:700,fontSize:11,color:B.white,lineHeight:1.3}}>{o.t}</span>
                </div>
                <div style={{fontSize:10,color:B.gray2,lineHeight:1.6,paddingLeft:19}}>{o.b}</div>
              </div>
            ))}
          </Pnl>
        </div>

        <Pnl style={{borderLeft:`3px solid ${B.amberLt}`}}>
          <ST>🎯 Prioridades para esta semana</ST>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))",gap:10}}>
            {actions.map((a,i)=>(
              <div key={i} style={{background:B.panel,borderRadius:8,padding:"11px 13px",border:`1px solid ${B.border}`,display:"flex",gap:10,alignItems:"flex-start"}}>
                <div style={{fontSize:16,fontWeight:900,color:B.redBg,background:B.red,borderRadius:4,width:26,height:26,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,lineHeight:1}}>{i+1}</div>
                <div style={{fontSize:11,color:B.white,lineHeight:1.6}}>{a}</div>
              </div>
            ))}
          </div>
        </Pnl>

        <Pnl>
          <ST sub="Eje X=Ingresos · Eje Y=Margen% · Tamaño=líneas SAP">Mapa de rentabilidad</ST>
          <ResponsiveContainer width="100%" height={260}>
            <ScatterChart margin={{top:10,right:20,left:0,bottom:10}}>
              <CartesianGrid strokeDasharray="3 3" stroke={B.border}/>
              <XAxis type="number" dataKey="ing" name="Ingresos" stroke={B.gray3} tick={{fontSize:9}} tickFormatter={f$} width={62}/>
              <YAxis type="number" dataKey="mgnP" name="Margen%" stroke={B.gray3} tick={{fontSize:9}} tickFormatter={v=>`${v.toFixed(0)}%`} width={42}/>
              <ZAxis type="number" dataKey="lineas" range={[40,450]}/>
              <Tooltip cursor={{strokeDasharray:"3 3"}} content={({active,payload})=>{
                if(!active||!payload?.length)return null;
                const d=payload[0]?.payload;
                return<div style={{background:B.panel,border:`1px solid ${B.border}`,borderRadius:8,padding:"9px 13px",fontSize:11}}>
                  <p style={{color:B.red,fontWeight:700,marginBottom:4}}>{d?.nm}</p>
                  <p style={{color:B.white}}>Ingresos: {f$(d?.ing)}</p>
                  <p style={{color:d?.mgnP<0?B.red:B.greenLt}}>Margen: {fP(d?.mgnP)}</p>
                  <p style={{color:B.gray3}}>{d?.lineas} líneas · {d?.sector}</p>
                </div>;
              }}/>
              <Scatter data={clients.filter(c=>c.ing>0)} fill={B.red}>
                {clients.filter(c=>c.ing>0).map((c,i)=><Cell key={i} fill={c.mgnP<0?B.red:c.mgnP>30?B.greenLt:B.amberLt}/>)}
              </Scatter>
            </ScatterChart>
          </ResponsiveContainer>
          <div style={{display:"flex",gap:16,marginTop:6,fontSize:10,color:B.gray3}}>
            <span><span style={{color:B.greenLt}}>●</span> Margen &gt;30%</span>
            <span><span style={{color:B.amberLt}}>●</span> Margen 0–30%</span>
            <span><span style={{color:B.red}}>●</span> Margen negativo</span>
          </div>
        </Pnl>
      </div>
    );
  };

  const ViewFirebase=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      {/* Status banner */}
      <div style={{background:FB_CONFIGURED?B.greenDim:B.amberDim,border:`1px solid ${FB_CONFIGURED?B.green:B.amber}`,borderRadius:8,padding:"12px 16px",display:"flex",gap:10,alignItems:"center"}}>
        <span style={{fontSize:20}}>{FB_CONFIGURED?"✅":"⚠️"}</span>
        <div>
          <div style={{fontWeight:700,color:FB_CONFIGURED?B.greenLt:B.amberLt,fontSize:12}}>
            {FB_CONFIGURED?"Firebase configurado — listo para compartir":"Firebase no configurado aún"}
          </div>
          <div style={{fontSize:10,color:B.gray2,marginTop:2}}>
            {FB_CONFIGURED?"Los Excels que subas aquí estarán disponibles para todos los usuarios.":"Sigue los pasos de abajo para activar el modo compartido gratuito."}
          </div>
        </div>
      </div>

      {!FB_CONFIGURED&&(
        <Pnl style={{borderLeft:`3px solid ${B.amber}`}}>
          <ST>📋 Pasos para activar Firebase (5 minutos, gratis)</ST>
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            {[
              {n:1,t:"Crear proyecto Firebase",d:"Ve a console.firebase.google.com → 'Crear proyecto' → nombre: cummins-zona-sur → desactiva Analytics → Crear"},
              {n:2,t:"Activar Realtime Database",d:"Menú izquierdo → Realtime Database → Crear base de datos → us-central1 → modo prueba → Crear"},
              {n:3,t:"Activar Storage",d:"Menú izquierdo → Storage → Comenzar → modo prueba → Listo"},
              {n:4,t:"Obtener credenciales",d:"Configuración (⚙️) → General → 'Tus apps' → botón </> (Web) → nombre: zona-sur → Registrar → copiar firebaseConfig"},
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
                </tr>
              ))}
              <tr style={{borderTop:`1px solid ${B.border}`,background:B.panel}}>
                <td style={{padding:"7px 10px",fontWeight:800,color:B.red}}>TOTAL</td>
                <td/>
                <td style={{padding:"7px 10px",fontWeight:800,color:B.amberLt}}>{fN(files.reduce((s,f)=>s+f.count,0))} líneas</td>
              </tr>
            </tbody>
          </table>
        )}
      </Pnl>

      {/* Architecture explanation */}
      <Pnl>
        <ST>📐 Arquitectura del sistema compartido</ST>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:12}}>
          {[
            {icon:"📂",t:"Tú subes",d:"Arrastras el Excel aquí → se sube a Firebase Storage automáticamente"},
            {icon:"☁️",t:"Firebase guarda",d:"El archivo queda en la nube con una URL pública. Gratis hasta 1 GB"},
            {icon:"👥",t:"Todos ven",d:"Cualquier persona con el link del dashboard descarga y procesa los datos"},
            {icon:"🔒",t:"Filtros locales",d:"Cada usuario tiene sus propios filtros. No interfieren entre sí"},
            {icon:"🌐",t:"GitHub Pages",d:"Hosting gratuito y permanente. URL tipo: tu-usuario.github.io/cummins"},
            {icon:"♾️",t:"Sin costo",d:"Firebase free tier: 1GB storage, 10GB/mes descarga. Más que suficiente"},
          ].map(x=>(
            <div key={x.t} style={{background:B.panel,borderRadius:8,padding:"12px 14px",border:`1px solid ${B.border}`}}>
              <div style={{fontSize:20,marginBottom:6}}>{x.icon}</div>
              <div style={{fontWeight:700,color:B.white,fontSize:11,marginBottom:4}}>{x.t}</div>
              <div style={{fontSize:10,color:B.gray3,lineHeight:1.5}}>{x.d}</div>
            </div>
          ))}
        </div>
      </Pnl>
    </div>
  );

  const renderView=()=>{
    switch(view){
      case "RESUMEN":         return <ViewResumen/>;
      case "CLIENTES":        return <ViewClientes/>;
      case "PRODUCTOS":       return <ViewProductos/>;
      case "SECTORES":        return <ViewSectores/>;
      case "TENDENCIAS":      return <ViewTendencias/>;
      case "EJECUTIVOS":      return <ViewEjecutivos/>;
      case "TABLA DINÁMICA":  return <ViewPivot/>;
      case "ANÁLISIS":        return <ViewAnalisis/>;
      case "FIREBASE":        return <ViewFirebase/>;
      default:                return <ViewResumen/>;
    }
  };

  // ══════════════════════════════════════════════════════════
  // SHELL
  // ══════════════════════════════════════════════════════════
  return(
    <div style={{minHeight:"100vh",background:B.bg,color:B.white,fontFamily:"'Helvetica Neue',Helvetica,Arial,sans-serif",display:"flex",flexDirection:"column"}}>
      <style>{`
        *{box-sizing:border-box;margin:0;padding:0}
        body{background:${B.bg};font-family:'Helvetica Neue',Helvetica,Arial,sans-serif}
        ::-webkit-scrollbar{width:4px;height:4px}
        ::-webkit-scrollbar-track{background:transparent}
        ::-webkit-scrollbar-thumb{background:${B.border};border-radius:2px}
        ::-webkit-scrollbar-thumb:hover{background:${B.red}}
        @keyframes fadeIn{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:translateY(0)}}
        .fade-in{animation:fadeIn .25s ease}
        input::placeholder{color:${B.gray4}}
        td,th{overflow:hidden;text-overflow:ellipsis}
        .recharts-wrapper,.recharts-surface{overflow:visible!important}
      `}</style>

      {/* ── TOPBAR ── */}
      <div style={{background:B.red,padding:"0 16px",display:"flex",alignItems:"center",height:54,flexShrink:0,boxShadow:"0 2px 16px rgba(120,0,0,.5)",gap:0}}>
        {/* Brand */}
        <div style={{display:"flex",alignItems:"center",gap:10,paddingRight:18,borderRight:"1px solid rgba(255,255,255,.25)",marginRight:14,flexShrink:0}}>
          <CumminsLogo h={32}/>
          <div style={{lineHeight:1.2}}>
            <div style={{fontWeight:900,fontSize:13,color:B.white,letterSpacing:".06em"}}>CUMMINS PERÚ</div>
            <div style={{fontSize:8,color:"rgba(255,255,255,.6)",letterSpacing:".2em",marginTop:1}}>ZONA SUR · COMERCIAL</div>
          </div>
        </div>

        {/* Nav */}
        <nav style={{display:"flex",flex:1,overflowX:"auto",gap:0,msOverflowStyle:"none",scrollbarWidth:"none"}}>
          {VIEWS.map(v=>(
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
      </div>

      {/* Drag overlay */}
      {dragging&&(
        <div onDrop={e=>{e.preventDefault();setDragging(false);ingest(Array.from(e.dataTransfer.files));}}
          onDragOver={e=>e.preventDefault()} onDragLeave={()=>setDragging(false)}
          style={{position:"fixed",inset:0,background:"rgba(204,0,0,.12)",border:"3px dashed #CC0000",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",backdropFilter:"blur(4px)"}}>
          <div style={{textAlign:"center",color:B.white}}>
            <div style={{fontSize:52,marginBottom:12}}>📂</div>
            <div style={{fontSize:18,fontWeight:900,letterSpacing:".04em"}}>Suelta los archivos Excel aquí</div>
            <div style={{fontSize:12,color:"rgba(255,255,255,.6)",marginTop:6}}>Se cargarán automáticamente</div>
          </div>
        </div>
      )}
      <div style={{position:"fixed",inset:0,pointerEvents:"none",zIndex:998}} onDragEnter={()=>setDragging(true)}/>
    </div>
  );
}
