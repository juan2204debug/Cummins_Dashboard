import { useState, useMemo, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { AreaChart,Area,BarChart,Bar,LineChart,Line,XAxis,YAxis,
  CartesianGrid,Tooltip,Legend,ResponsiveContainer,PieChart,Pie,Cell,
  ScatterChart,Scatter,ZAxis,ComposedChart } from "recharts";

// ═══════════════════════════════════════════════
// BRAND
// ═══════════════════════════════════════════════
const B={
  red:"#CC0000",redDk:"#990000",redBg:"#150000",redDim:"#220000",
  white:"#FFFFFF",g1:"#E0E0E0",g2:"#A8A8A8",g3:"#606060",g4:"#303030",
  bg:"#080808",surf:"#101010",panel:"#141414",card:"#181818",
  border:"#252525",borderHi:"#383838",
  green:"#16A34A",greenDim:"#071507",greenLt:"#4ADE80",
  amber:"#D97706",amberDim:"#1A0E00",amberLt:"#FCD34D",
  blue:"#2563EB",blueDim:"#040D1C",blueLt:"#93C5FD",
  teal:"#0D9488",tealDim:"#041210",
  ch:["#CC0000","#2563EB","#16A34A","#D97706","#7C3AED","#0D9488","#DB2777","#EA580C","#0891B2","#65A30D","#E57373","#CA8A04"],
};

const Logo=()=>(<svg width={36} height={30} viewBox="0 0 96 80"><polygon points="48,2 90,22 90,58 48,78 6,58 6,22" fill={B.red}/><text x="48" y="44" textAnchor="middle" fill="#fff" fontFamily="Arial Black,sans-serif" fontSize="15" fontWeight="900" letterSpacing=".5">CUMMINS</text><text x="48" y="57" textAnchor="middle" fill="rgba(255,255,255,.6)" fontFamily="Arial,sans-serif" fontSize="7" letterSpacing="2">ZONA SUR</text></svg>);

// ═══════════════════════════════════════════════
// FISCAL YEAR
// ═══════════════════════════════════════════════
const FM=["Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic","Ene","Feb","Mar"];
const FQN=["Q1 Abr-Jun","Q2 Jul-Sep","Q3 Oct-Dic","Q4 Ene-Mar"];
const getFY=(m,y)=>m>=4?y:y-1;
const getFMI=(m)=>m>=4?m-4:m+8;
const getFQ=(m)=>Math.floor(getFMI(m)/3)+1;
const fyLbl=(fy)=>`FY ${fy}/${String(fy+1).slice(2)}`;

// ═══════════════════════════════════════════════
// UTILS
// ═══════════════════════════════════════════════
const f$=(n,d=0)=>{if(n==null)return"—";const a=Math.abs(n);const s=a>=1e6?`$${(a/1e6).toFixed(2)}M`:a>=1e3?`$${(a/1e3).toFixed(1)}K`:`$${a.toFixed(d)}`;return n<0?`(${s})`:s;};
const fP=(n,d=1)=>n!=null?`${n>0?"+":""}${n.toFixed(d)}%`:"—";
const fN=(n)=>n?.toLocaleString("es-PE")??"—";
const clamp=(v,a,b)=>Math.min(Math.max(v,a),b);
const pNum=(v)=>{if(v===null||v===undefined||v==="")return 0;if(typeof v==="number")return isFinite(v)?v:0;const s=String(v).replace(/[^\d.,\-]/g,"");if(!s)return 0;return parseFloat(s.replace(",","."))||0;};

// ═══════════════════════════════════════════════
// COLUMN DETECTION
// ═══════════════════════════════════════════════
function findCol(headers,exactNames,containsNames=[]){
  for(const e of exactNames){const f=headers.find(h=>h.trim().toLowerCase()===e.toLowerCase());if(f)return f;}
  for(const c of containsNames){const f=headers.find(h=>h.toLowerCase().includes(c.toLowerCase()));if(f)return f;}
  return null;
}

// ═══════════════════════════════════════════════
// PARSER — pre-aggregates per file
// ═══════════════════════════════════════════════
function parseAndAggregate(fileLabel,rows){
  if(!rows.length)return null;
  const headers=Object.keys(rows[0]);
  const cols={
    ing:  findCol(headers,["Ingreso TOT","INGRESO TOT"]),
    cst:  findCol(headers,["Costo TOT","COSTO TOT"]),
    cli:  findCol(headers,["Cliente","CLIENTE"]),
    clinm:findCol(headers,["Descripción Cliente","Descripcion Cliente"]),
    f2:   findCol(headers,["Jerarquia de producto 2","Jerarquía de producto 2"]),
    f1:   findCol(headers,["Jerarquia de producto","Jerarquía de producto"]),
    art:  findCol(headers,["Articulo","Artículo"]),
    suc:  findCol(headers,["Sucursal","SUCURSAL"]),
    sec:  findCol(headers,["Descrip. Gr.Clie.","Descrip.Gr.Clie."]),
    rep:  findCol(headers,["Representante de ventas","Representante De Ventas"]),
    vol:  findCol(headers,["Vol.Ventas","VOL.VENTAS"],["vol.ventas"]),
    per:  findCol(headers,["Período/Año","Periodo/Año"],["período/año","periodo/año"]),
    fecha:findCol(headers,["Fecha factura","Fecha Factura"]),
    fcont:findCol(headers,["Fecha de contabilización"],["contabiliz"]),
  };
  const fnM=fileLabel.match(/(\d{4})_(\d{1,2})/);
  const fnYear=fnM?parseInt(fnM[1]):null;
  const fnMonth=fnM?parseInt(fnM[2]):null;
  const sucursal=fileLabel.match(/K(\d+)/i)?.[0]?.toUpperCase()||"K??";
  const byClient={},byFam={},bySec={},byRep={},byArt={};
  let totalIng=0,totalCst=0,totalVol=0,salesRows=0;
  rows.forEach(r=>{
    const ing=pNum(r[cols.ing]),cst=pNum(r[cols.cst]);
    const cli=r[cols.cli]?String(r[cols.cli]).trim():null;
    if(!cli||cli==="0"||cli.toLowerCase()==="nan")return;
    if(ing===0&&cst===0)return;
    let mo=fnMonth,yr=fnYear;
    if(!mo&&r[cols.per]){const ps=String(Math.round(pNum(r[cols.per]))||0).replace(/\D/g,"");if(ps.length>=5){yr=parseInt(ps.slice(0,4));mo=parseInt(ps.slice(4))||4;}}
    if(!mo){const fv=r[cols.fecha]||r[cols.fcont];if(fv){const d=fv instanceof Date?fv:new Date(fv);if(!isNaN(d)&&d.getFullYear()>2000){mo=d.getMonth()+1;yr=d.getFullYear();}}}
    if(!mo){mo=4;yr=new Date().getFullYear();}
    const fy=getFY(mo,yr),fmi=getFMI(mo),fq=getFQ(mo),fyl=fyLbl(fy),fml=FM[fmi];
    const mgn=ing-cst,vol=pNum(r[cols.vol]);
    const clinm=r[cols.clinm]?String(r[cols.clinm]).trim():"Sin nombre";
    const secRaw=r[cols.sec]?String(r[cols.sec]).trim():"";
    const sec=(secRaw&&isNaN(Number(secRaw))&&secRaw.length>1)?secRaw:"Otros";
    const f2=r[cols.f2]?String(r[cols.f2]).trim():"Sin familia";
    const f1=r[cols.f1]?String(r[cols.f1]).trim():"Sin familia";
    const art=r[cols.art]?String(r[cols.art]).trim():"?";
    const repRaw=r[cols.rep]?String(r[cols.rep]).trim():"0";
    const rep=(repRaw&&repRaw!=="nan"&&repRaw!=="0.0")?repRaw:"0";
    totalIng+=ing;totalCst+=cst;totalVol+=vol;salesRows++;
    if(!byClient[cli])byClient[cli]={id:cli,nm:clinm,sector:sec,rep,ing:0,cst:0,mgn:0,vol:0,lineas:0,months:new Set()};
    const c=byClient[cli];c.ing+=ing;c.cst+=cst;c.mgn+=mgn;c.vol+=vol;c.lineas++;c.months.add(`${yr}/${mo}`);
    if(sec!=="Otros")c.sector=sec;if(rep!=="0")c.rep=rep;
    if(!byFam[f2])byFam[f2]={f2,f1,ing:0,cst:0,mgn:0,vol:0};
    byFam[f2].ing+=ing;byFam[f2].cst+=cst;byFam[f2].mgn+=mgn;byFam[f2].vol+=vol;
    if(!bySec[sec])bySec[sec]={sec,ing:0,lineas:0,clis:new Set()};
    bySec[sec].ing+=ing;bySec[sec].lineas++;bySec[sec].clis.add(cli);
    if(rep!=="0"){if(!byRep[rep])byRep[rep]={rep,ing:0,lineas:0,clis:new Set()};byRep[rep].ing+=ing;byRep[rep].lineas++;byRep[rep].clis.add(cli);}
    if(art!=="?"){if(!byArt[art])byArt[art]={art,f2,f1,ing:0,cst:0,mgn:0,vol:0};byArt[art].ing+=ing;byArt[art].cst+=cst;byArt[art].mgn+=mgn;byArt[art].vol+=vol;}
  });
  if(salesRows===0)return null;
  const clients=Object.values(byClient).map(c=>({...c,meses:c.months.size,mgnP:c.ing?Math.round(c.mgn/c.ing*1000)/10:0,months:undefined})).sort((a,b)=>b.ing-a.ing);
  return{
    label:fileLabel,sucursal,mo:fnMonth,yr:fnYear,
    fy:getFY(fnMonth??4,fnYear??2026),fmi:getFMI(fnMonth??4),
    fml:FM[getFMI(fnMonth??4)],fyl:fyLbl(getFY(fnMonth??4,fnYear??2026)),
    fq:getFQ(fnMonth??4),totalIng,totalCst,totalVol,salesRows,clients,
    familia:Object.values(byFam).map(f=>({...f,mgnP:f.ing?Math.round(f.mgn/f.ing*1000)/10:0})).sort((a,b)=>b.ing-a.ing),
    sector:Object.values(bySec).map(s=>({...s,clientes:s.clis.size,clis:undefined})).sort((a,b)=>b.ing-a.ing),
    reps:Object.values(byRep).map(r=>({...r,clientes:r.clis.size,clis:undefined})).sort((a,b)=>b.ing-a.ing),
    articles:Object.values(byArt).map(a=>({...a,mgnP:a.ing?Math.round(a.mgn/a.ing*1000)/10:0})).sort((a,b)=>b.ing-a.ing).slice(0,30),
    cols:Object.fromEntries(Object.entries(cols).map(([k,v])=>[k,v||null])),
    headers:headers.slice(0,25),
  };
}
async function loadXLSX(file){
  const ab=await file.arrayBuffer();
  const wb=XLSX.read(ab,{type:"array",cellDates:true});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{defval:null});
  return parseAndAggregate(file.name.replace(/\.[^.]+$/,""),rows);
}

// ═══════════════════════════════════════════════
// AUTO ANALYSIS
// ═══════════════════════════════════════════════
function autoAnalysis(clients,familia,reps,kpi,repNames,periodLabel){
  const alerts=[],opps=[],actions=[];
  familia.filter(f=>f.mgnP<0&&Math.abs(f.ing)>100).slice(0,6).forEach(f=>
    alerts.push({icon:"🔴",t:`${f.f2}: margen ${f.mgnP.toFixed(0)}%`,b:`Pérdida de ${f$(Math.abs(f.mgn))} en ${Math.round(f.vol)} uds. Revisar precio SAP.`})
  );
  const top2pct=clients.slice(0,2).reduce((s,c)=>s+(c.pct||0),0);
  if(top2pct>50)alerts.push({icon:"🟡",t:`Concentración: ${top2pct.toFixed(0)}% en 2 clientes`,b:"Riesgo de dependencia. Diversificar cartera."});
  clients.filter(c=>c.mgnP<-2&&c.ing>200).slice(0,3).forEach(c=>
    alerts.push({icon:"🟠",t:`${c.nm.split(" ").slice(0,3).join(" ")}: margen ${c.mgnP.toFixed(0)}%`,b:`${f$(c.ing)} facturado con pérdida de ${f$(Math.abs(c.mgn))}.`})
  );
  familia.filter(f=>f.mgnP>35&&f.ing>500).slice(0,4).forEach(f=>
    opps.push({icon:"💎",t:`${f.f2}: ${f.mgnP.toFixed(0)}% margen`,b:`${f$(f.ing)} en ingresos. Penetrar más clientes con esta línea.`})
  );
  clients.filter(c=>c.mgnP>40&&c.ing>300&&c.meses<3).slice(0,2).forEach(c=>
    opps.push({icon:"🚀",t:`${c.nm.split(" ").slice(0,3).join(" ")}: potencial`,b:`${c.mgnP.toFixed(0)}% margen en solo ${c.meses} mes(es).`})
  );
  if(alerts.some(a=>a.icon==="🔴"))actions.push(`Revisión urgente de precios: ${familia.filter(f=>f.mgnP<0).map(f=>f.f2).slice(0,3).join(", ")}`);
  if(clients[0])actions.push(`Visita confirmada a ${clients[0].nm.split(" ").slice(0,2).join(" ")} — ${clients[0].pct?.toFixed(0)||"?"}% del ingreso`);
  if(reps[0])actions.push(`${repNames[reps[0].rep]||`Rep ${reps[0].rep}`} lidera con ${f$(reps[0].ing)} — revisar pipeline próximo mes`);
  clients.filter(c=>c.mgnP<0&&c.ing>500).slice(0,2).forEach(c=>actions.push(`Revisar condición comercial: ${c.nm.split(" ").slice(0,2).join(" ")}`));
  actions.push("Prospectar nuevos clientes Minería/Construcción para reducir concentración");
  return{alerts:alerts.slice(0,8),opps:opps.slice(0,6),actions:actions.slice(0,5),periodLabel};
}

// ═══════════════════════════════════════════════
// TOOLTIP
// ═══════════════════════════════════════════════
const TT=({active,payload,label})=>{
  if(!active||!payload?.length)return null;
  return(<div style={{background:B.g4,border:`1px solid ${B.borderHi}`,borderRadius:8,padding:"9px 13px",fontSize:11,boxShadow:"0 8px 24px #0009"}}>
    <p style={{color:B.red,fontWeight:800,marginBottom:5}}>{label}</p>
    {payload.map((p,i)=>(<p key={i} style={{color:p.color||B.white,margin:"2px 0",display:"flex",gap:16,justifyContent:"space-between"}}>
      <span style={{color:B.g2,fontSize:10}}>{p.name}</span>
      <strong>{Math.abs(p.value??0)>50?f$(p.value):typeof p.value==="number"?p.value.toFixed(1):p.value}</strong>
    </p>))}
  </div>);
};

const VIEWS=["RESUMEN","CLIENTES","PRODUCTOS","SECTORES","TENDENCIAS","EJECUTIVOS","TABLA DIN.","ANÁLISIS","FIREBASE"];

// ═══════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════
export default function App(){
  const [fileAggs,setFileAggs]=useState([]);
  const [loading,setLoading]=useState(false);
  const [view,setView]=useState("RESUMEN");
  const [dragging,setDragging]=useState(false);

  // Filters
  const [fSuc,setFSuc]=useState([]);
  const [fFY,setFY]=useState([]);
  const [fQ,setFQ]=useState([]);
  const [fSec,setFSec]=useState([]);
  const [fFam,setFFam]=useState([]);
  const [fRep,setFRep]=useState([]);
  const [fSeg,setFSeg]=useState([]);
  const [search,setSearch]=useState("");
  const [selCli,setSelCli]=useState(null);
  const [selRep,setSelRep]=useState(null); // for exec profile

  const [cartera,setCartera]=useState(()=>{try{return JSON.parse(localStorage.getItem("cmn_c")||"{}")}catch{return{}}});
  const [repNames,setRepNames]=useState(()=>{try{return JSON.parse(localStorage.getItem("cmn_r")||"{}")}catch{return{}}});
  const [pRows,setPRows]=useState("sector");
  const [pCols,setPCols]=useState("fyl");
  const [pVal,setPVal]=useState("ing");
  const [pSort,setPSort]=useState("desc");

  useEffect(()=>{localStorage.setItem("cmn_c",JSON.stringify(cartera));},[cartera]);
  useEffect(()=>{localStorage.setItem("cmn_r",JSON.stringify(repNames));},[repNames]);

  // Auto-select sucursal when first file loaded
  useEffect(()=>{
    if(fileAggs.length>0&&fSuc.length===0){
      const sucs=[...new Set(fileAggs.map(f=>f.sucursal))].sort();
      if(sucs.length===1)setFSuc(sucs); // auto-select if only one
    }
  },[fileAggs]);

  // ── Ingest ───────────────────────────────────────────────────────────────
  const ingest=useCallback(async(files)=>{
    setLoading(true);
    const xlsFiles=Array.from(files).filter(f=>f.name.match(/\.xlsx?$/i));
    if(!xlsFiles.length){setLoading(false);return;}
    const results=[];
    for(const f of xlsFiles){
      try{const agg=await loadXLSX(f);if(agg)results.push(agg);}
      catch(e){console.error("Error parsing",f.name,e);}
    }
    if(results.length){
      setFileAggs(prev=>{
        const existing=new Set(prev.map(f=>f.label));
        const fresh=results.filter(r=>!existing.has(r.label));
        return fresh.length?[...prev,...fresh]:prev;
      });
    }
    setLoading(false);
  },[]);

  // ── Filter options ────────────────────────────────────────────────────────
  const opts=useMemo(()=>({
    sucs:[...new Set(fileAggs.map(f=>f.sucursal))].sort(),
    fys:[...new Set(fileAggs.map(f=>f.fyl))].sort(),
    secs:[...new Set(fileAggs.flatMap(f=>f.sector.map(s=>s.sec)))].sort(),
    fams:[...new Set(fileAggs.flatMap(f=>f.familia.map(x=>x.f2)))].sort(),
    reps:[...new Set(fileAggs.flatMap(f=>f.reps.map(r=>r.rep)))].sort(),
  }),[fileAggs]);

  // ── Apply filters ─────────────────────────────────────────────────────────
  const filteredAggs=useMemo(()=>fileAggs.filter(f=>{
    if(fSuc.length&&!fSuc.includes(f.sucursal))return false;
    if(fFY.length&&!fFY.includes(f.fyl))return false;
    if(fQ.length&&!fQ.some(q=>parseInt(q)===f.fq))return false;
    return true;
  }),[fileAggs,fSuc,fFY,fQ]);

  // ── KPIs ──────────────────────────────────────────────────────────────────
  const kpi=useMemo(()=>{
    const ing=filteredAggs.reduce((s,f)=>s+f.totalIng,0);
    const cst=filteredAggs.reduce((s,f)=>s+f.totalCst,0);
    const mgn=ing-cst;
    return{ing,cst,mgn,mgnP:ing?mgn/ing*100:0,
      vol:filteredAggs.reduce((s,f)=>s+f.totalVol,0),
      lineas:filteredAggs.reduce((s,f)=>s+f.salesRows,0),
      clis:new Set(filteredAggs.flatMap(f=>f.clients.map(c=>c.id))).size};
  },[filteredAggs]);

  // Period label for analysis
  const periodLabel=useMemo(()=>{
    const sucs=fSuc.length?fSuc.join("+"):"Todas";
    const fys=fFY.length?fFY.join(" / "):[...new Set(filteredAggs.map(f=>f.fyl))].join(" / ")||"—";
    const qs=fQ.length?` Q${fQ.join("+")}`:"";
    return`${sucs} · ${fys}${qs}`;
  },[fSuc,fFY,fQ,filteredAggs]);

  // ── Aggregations ──────────────────────────────────────────────────────────
  const allClients=useMemo(()=>{
    const m={};
    filteredAggs.forEach(f=>f.clients.forEach(c=>{
      if(fSec.length&&!fSec.includes(c.sector))return;
      if(!m[c.id])m[c.id]={id:c.id,nm:c.nm,sector:c.sector,rep:c.rep,ing:0,cst:0,mgn:0,vol:0,lineas:0,meses:0};
      const x=m[c.id];x.ing+=c.ing;x.cst+=c.cst;x.mgn+=c.mgn;x.vol+=c.vol;x.lineas+=c.lineas;
      x.meses=Math.max(x.meses,c.meses);
      if(c.sector&&c.sector!=="Otros")x.sector=c.sector;
      if(c.rep&&c.rep!=="0")x.rep=c.rep;
    }));
    const arr=Object.values(m).map(c=>({...c,mgnP:c.ing?c.mgn/c.ing*100:0})).sort((a,b)=>b.ing-a.ing);
    const total=arr.reduce((s,c)=>s+c.ing,0);let cum=0;
    return arr.map(c=>{c.pct=total?c.ing/total*100:0;cum+=c.pct;c.cum=cum;c.seg=cum<=60?"TOP":cum<=90?"MED":"LOW";return c;});
  },[filteredAggs,fSec]);

  const allFamilia=useMemo(()=>{
    const m={};
    filteredAggs.forEach(f=>f.familia.forEach(x=>{
      if(fFam.length&&!fFam.includes(x.f2))return;
      if(!m[x.f2])m[x.f2]={f2:x.f2,f1:x.f1,ing:0,cst:0,mgn:0,vol:0};
      m[x.f2].ing+=x.ing;m[x.f2].cst+=x.cst;m[x.f2].mgn+=x.mgn;m[x.f2].vol+=x.vol;
    }));
    return Object.values(m).map(f=>({...f,mgnP:f.ing?f.mgn/f.ing*100:0})).sort((a,b)=>b.ing-a.ing);
  },[filteredAggs,fFam]);

  const allSector=useMemo(()=>{
    const m={};
    filteredAggs.forEach(f=>f.sector.forEach(s=>{
      if(!m[s.sec])m[s.sec]={sec:s.sec,ing:0,lineas:0,clientes:0};
      m[s.sec].ing+=s.ing;m[s.sec].lineas+=s.lineas;m[s.sec].clientes+=s.clientes;
    }));
    return Object.values(m).filter(s=>s.ing>0).sort((a,b)=>b.ing-a.ing);
  },[filteredAggs]);

  // Pie data — cap at 9 sectors, group rest as "Otros"
  const pieData=useMemo(()=>{
    const top=allSector.slice(0,9);
    const rest=allSector.slice(9);
    if(!rest.length)return top;
    const otrosIng=rest.reduce((s,x)=>s+x.ing,0);
    const otrosCli=rest.reduce((s,x)=>s+x.clientes,0);
    return[...top,{sec:"Otros sectores",ing:otrosIng,clientes:otrosCli,lineas:rest.reduce((s,x)=>s+x.lineas,0)}];
  },[allSector]);

  const allReps=useMemo(()=>{
    const m={};
    filteredAggs.forEach(f=>f.reps.forEach(r=>{
      if(fRep.length&&!fRep.includes(r.rep))return;
      if(!m[r.rep])m[r.rep]={rep:r.rep,ing:0,lineas:0,clientes:0};
      m[r.rep].ing+=r.ing;m[r.rep].lineas+=r.lineas;m[r.rep].clientes+=r.clientes;
    }));
    return Object.values(m).sort((a,b)=>b.ing-a.ing);
  },[filteredAggs,fRep]);

  const allArticles=useMemo(()=>{
    const m={};
    filteredAggs.forEach(f=>f.articles.forEach(a=>{
      if(!m[a.art])m[a.art]={art:a.art,f2:a.f2,f1:a.f1,ing:0,cst:0,mgn:0,vol:0};
      m[a.art].ing+=a.ing;m[a.art].cst+=a.cst;m[a.art].mgn+=a.mgn;m[a.art].vol+=a.vol;
    }));
    return Object.values(m).map(a=>({...a,mgnP:a.ing?a.mgn/a.ing*100:0})).sort((a,b)=>b.ing-a.ing).slice(0,25);
  },[filteredAggs]);

  const fyList=[...new Set(filteredAggs.map(f=>f.fyl))].sort();

  const monthlyTrend=useMemo(()=>FM.map((m,i)=>{
    const row={mes:m};
    fyList.forEach(fy=>{row[fy]=filteredAggs.filter(f=>f.fyl===fy&&f.fmi===i).reduce((s,f)=>s+f.totalIng,0);});
    return row;
  }),[filteredAggs,fyList.join()]);

  const fyTotals=useMemo(()=>fyList.map(fy=>{
    const aggs=filteredAggs.filter(f=>f.fyl===fy);
    const ing=aggs.reduce((s,f)=>s+f.totalIng,0),mgn=aggs.reduce((s,f)=>s+(f.totalIng-f.totalCst),0);
    return{fy,ing,mgn,mgnP:ing?mgn/ing*100:0};
  }),[filteredAggs,fyList.join()]);

  const quarterData=useMemo(()=>[1,2,3,4].map((q,i)=>{
    const row={q:FQN[i].split(" ")[0]};
    fyList.forEach(fy=>{row[fy]=filteredAggs.filter(f=>f.fyl===fy&&f.fq===q).reduce((s,f)=>s+f.totalIng,0);});
    return row;
  }),[filteredAggs,fyList.join()]);

  const filteredClients=useMemo(()=>allClients.filter(c=>{
    if(fSeg.length&&!fSeg.includes(c.seg))return false;
    if(search){const q=search.toLowerCase();return c.nm.toLowerCase().includes(q)||c.id.includes(q);}
    return true;
  }),[allClients,fSeg,search]);

  // ── Exec profile data ─────────────────────────────────────────────────────
  const repProfile=useMemo(()=>{
    if(!selRep)return null;
    const repClients=allClients.filter(c=>{
      const assigned=cartera[selRep]?.includes(c.id);
      return assigned||(c.rep===selRep&&!Object.values(cartera).flat().includes(c.id));
    });
    const total=repClients.reduce((s,c)=>s+c.ing,0);
    let cum=0;
    const withPareto=repClients.map(c=>{
      c.repPct=total?c.ing/total*100:0;cum+=c.repPct;c.repCum=cum;
      c.repSeg=cum<=60?"TOP":cum<=90?"MED":"LOW";return c;
    });
    const monthly=FM.map((m,i)=>({mes:m,ing:filteredAggs.filter(f=>f.reps.some(r=>r.rep===selRep)&&f.fmi===i).reduce((s,f)=>{const r=f.reps.find(r=>r.rep===selRep);return s+(r?.ing||0);},0)}));
    return{clients:withPareto,monthly,total,kpi:{ing:total,mgn:repClients.reduce((s,c)=>s+c.mgn,0),lineas:repClients.reduce((s,c)=>s+c.lineas,0),clis:repClients.length}};
  },[selRep,allClients,filteredAggs,cartera]);

  // ── Pivot ─────────────────────────────────────────────────────────────────
  const PDIMS={sector:"Sector",fyl:"Año Fiscal",f2:"Sub-familia",f1:"Familia",rep:"Ejecutivo",suc:"Sucursal",cliente:"Cliente"};
  const PVALS={ing:{l:"Ingresos $",f:f$},mgn:{l:"Margen $",f:f$},vol:{l:"Volumen",f:fN},lineas:{l:"Líneas",f:fN}};

  const pivotData=useMemo(()=>{
    const matrix={},colSet=new Set();
    filteredAggs.forEach(agg=>{
      const colKey=pCols==="fyl"?agg.fyl:pCols==="suc"?agg.sucursal:"?";
      colSet.add(colKey);
      if(pRows==="cliente"){
        agg.clients.forEach(c=>{
          if(!matrix[c.nm])matrix[c.nm]={};
          matrix[c.nm][colKey]=(matrix[c.nm][colKey]||0)+(pVal==="ing"?c.ing:pVal==="mgn"?c.mgn:pVal==="vol"?c.vol:c.lineas);
        });
      } else {
        const src=pRows==="sector"?agg.sector:pRows==="f2"||pRows==="f1"?agg.familia:pRows==="rep"?agg.reps:null;
        if(!src)return;
        src.forEach(row=>{
          const rk=pRows==="sector"?row.sec:pRows==="f2"?row.f2:pRows==="f1"?row.f1:pRows==="rep"?(repNames[row.rep]||`Rep ${row.rep}`):row;
          const val=pVal==="ing"?row.ing:pVal==="mgn"?row.mgn:pVal==="vol"?row.vol:row.lineas||0;
          if(!matrix[rk])matrix[rk]={};
          matrix[rk][colKey]=(matrix[rk][colKey]||0)+val;
        });
      }
    });
    const cols=[...colSet].sort();
    let rows=Object.entries(matrix).map(([k,v])=>({rowKey:k,vals:v,total:cols.reduce((s,c)=>s+(v[c]||0),0)}));
    rows=pSort==="desc"?rows.sort((a,b)=>b.total-a.total):rows.sort((a,b)=>a.total-b.total);
    return{rows:rows.slice(0,100),cols};
  },[filteredAggs,pRows,pCols,pVal,pSort,repNames]);

  const analysis=useMemo(()=>autoAnalysis(allClients,allFamilia,allReps,kpi,repNames,periodLabel),[allClients,allFamilia,allReps,kpi,repNames,periodLabel]);

  // ── Helpers ───────────────────────────────────────────────────────────────
  const toggle=(s,v)=>s(p=>p.includes(v)?p.filter(x=>x!==v):[...p,v]);
  const clearAll=()=>{setFSuc([]);setFY([]);setFQ([]);setFSec([]);setFFam([]);setFRep([]);setFSeg([]);setSearch("");};
  const hasFilters=fSuc.length||fFY.length||fQ.length||fSec.length||fFam.length||fRep.length||fSeg.length||search;
  const hasData=fileAggs.length>0;
  const noSucSelected=hasData&&fSuc.length===0&&opts.sucs.length>1;

  // ══════════════════════════════════════════════
  // UI COMPONENTS
  // ══════════════════════════════════════════════
  const Pill=({label,active,color=B.red,onClick})=>(
    <button onClick={onClick} style={{background:active?`${color}22`:"transparent",border:`1px solid ${active?color:B.border}`,color:active?color:B.g3,borderRadius:20,padding:"3px 10px",fontSize:10,cursor:"pointer",fontFamily:"inherit",fontWeight:active?700:400,letterSpacing:".04em",transition:"all .12s",whiteSpace:"nowrap",lineHeight:"16px"}}>{label}</button>
  );
  const KCard=({icon,label,value,sub,color=B.red,warn=false})=>(
    <div style={{background:warn?B.redBg:B.card,border:`1px solid ${warn?B.red:B.border}`,borderRadius:10,padding:"13px 16px",borderTop:`2px solid ${warn?B.red:color}`,flex:"1 1 145px",minWidth:145,maxWidth:260}}>
      <div style={{fontSize:9,color:B.g3,letterSpacing:".12em",textTransform:"uppercase",marginBottom:5,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{icon} {label}</div>
      <div style={{fontSize:17,fontWeight:800,color:warn?B.red:color,lineHeight:1.2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{value}</div>
      {sub&&<div style={{fontSize:10,color:B.g3,marginTop:4,lineHeight:1.4}}>{sub}</div>}
    </div>
  );
  const Pnl=({children,style={}})=>(<div style={{background:B.card,border:`1px solid ${B.border}`,borderRadius:10,padding:16,...style}}>{children}</div>);
  const ST=({children,sub})=>(<div style={{marginBottom:12}}><div style={{fontSize:11,fontWeight:700,letterSpacing:".1em",textTransform:"uppercase",color:B.white,display:"flex",alignItems:"center",gap:7}}><span style={{display:"inline-block",width:3,height:12,background:B.red,borderRadius:2,flexShrink:0}}/>{children}</div>{sub&&<div style={{fontSize:10,color:B.g3,marginTop:2,paddingLeft:10}}>{sub}</div>}</div>);
  const MBadge=({v})=>{const c=v<0?B.red:v>25?B.green:B.amber;const bg=v<0?B.redDim:v>25?B.greenDim:B.amberDim;return <span style={{background:bg,color:c,borderRadius:3,padding:"2px 6px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{fP(v)}</span>;};
  const SBadge=({s,color})=>{const c=color||(s==="TOP"?B.amberLt:s==="MED"?B.blueLt:B.g3);const bg=s==="TOP"?B.amberDim:s==="MED"?B.blueDim:B.g4;return <span style={{background:bg,color:c,borderRadius:3,padding:"2px 6px",fontSize:9,fontWeight:700}}>{s}</span>;};

  // ── SIDEBAR ────────────────────────────────────────────────────────────────
  const Sidebar=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      {/* Sucursal — highlighted as primary filter */}
      <div style={{background:fSuc.length?`${B.red}18`:B.redBg,border:`1px solid ${fSuc.length?B.red:B.red+"44"}`,borderRadius:8,padding:"10px 12px"}}>
        <div style={{fontSize:9,color:B.red,letterSpacing:".14em",fontWeight:800,marginBottom:7}}>📍 SUCURSAL</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
          {opts.sucs.map(s=><Pill key={s} label={s} active={fSuc.includes(s)} color={B.red} onClick={()=>toggle(setFSuc,s)}/>)}
        </div>
        {!fSuc.length&&<div style={{fontSize:9,color:B.red,marginTop:6,fontStyle:"italic"}}>← Selecciona una sucursal</div>}
      </div>

      {/* Search */}
      <div><div style={{fontSize:9,color:B.g3,letterSpacing:".14em",fontWeight:700,marginBottom:5}}>BÚSQUEDA</div>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Cliente, código..." style={{width:"100%",background:B.panel,border:`1px solid ${B.border}`,borderRadius:5,color:B.white,fontFamily:"inherit",fontSize:11,padding:"6px 9px",outline:"none"}}/>
      </div>

      {[
        {lbl:"AÑO FISCAL",opts:opts.fys,active:fFY,set:setFY,color:B.blue},
        {lbl:"TRIMESTRE",opts:["1","2","3","4"],labels:["Q1 Abr","Q2 Jul","Q3 Oct","Q4 Ene"],active:fQ,set:setFQ,color:B.amber},
        {lbl:"SECTOR",opts:opts.secs,active:fSec,set:setFSec,color:B.teal,maxH:80},
        {lbl:"FAMILIA",opts:opts.fams,active:fFam,set:setFFam,color:B.red,maxH:80},
        {lbl:"EJECUTIVO",opts:opts.reps,active:fRep,set:setFRep,color:B.amber,nm:repNames},
        {lbl:"SEGMENTO",opts:["TOP","MED","LOW"],active:fSeg,set:setFSeg,color:B.amber},
      ].map(({lbl,opts:o,labels,active,set,color,nm,maxH})=>(
        <div key={lbl}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
            <span style={{fontSize:9,color:B.g3,letterSpacing:".1em",fontWeight:700}}>{lbl}</span>
            {active.length>0&&<button onClick={()=>set([])} style={{background:"none",border:"none",color:B.red,cursor:"pointer",fontSize:9,fontFamily:"inherit",padding:0}}>✕</button>}
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:3,maxHeight:maxH||9999,overflowY:maxH?"auto":"visible"}}>
            {(o||[]).map((opt,i)=>{const display=nm&&nm[opt]?nm[opt].split(" ")[0]:labels?labels[i]:String(opt).length>14?String(opt).slice(0,13)+"…":opt;return <Pill key={opt} label={display} active={active.includes(opt)} color={color} onClick={()=>toggle(set,opt)}/>;})}</div>
        </div>
      ))}

      {hasFilters&&<button onClick={clearAll} style={{width:"100%",background:B.redBg,border:`1px solid ${B.red}44`,color:B.red,borderRadius:6,padding:"6px",fontSize:10,cursor:"pointer",fontFamily:"inherit",fontWeight:700}}>✕ Limpiar todos</button>}

      <div style={{borderTop:`1px solid ${B.border}`,paddingTop:12,display:"flex",flexDirection:"column",gap:5}}>
        <div style={{fontSize:9,color:B.g3,letterSpacing:".1em",fontWeight:700,marginBottom:2}}>EN PANTALLA</div>
        {[{l:"Ingresos",v:f$(kpi.ing),c:B.red},{l:"Clientes",v:fN(kpi.clis),c:B.white},{l:"Líneas",v:fN(kpi.lineas),c:B.g2}].map(x=>(
          <div key={x.l} style={{background:B.panel,borderRadius:5,padding:"5px 9px",border:`1px solid ${B.border}`}}>
            <div style={{fontSize:9,color:B.g3}}>{x.l}</div>
            <div style={{color:x.c,fontWeight:800,fontSize:12,marginTop:1}}>{x.v}</div>
          </div>
        ))}
      </div>
    </div>
  );

  // ── NO SUCURSAL BANNER ────────────────────────────────────────────────────
  const NoBanner=()=>(
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",flex:1,gap:20,padding:48}}>
      <div style={{fontSize:40}}>📍</div>
      <div style={{textAlign:"center"}}>
        <div style={{fontWeight:800,fontSize:18,color:B.white,marginBottom:8}}>Selecciona una sucursal</div>
        <div style={{color:B.g3,fontSize:13,marginBottom:20}}>Tienes datos de {opts.sucs.length} sucursales. Para evitar mezclar información, selecciona una a la vez.</div>
        <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
          {opts.sucs.map(s=>(
            <button key={s} onClick={()=>setFSuc([s])} style={{background:B.redBg,border:`2px solid ${B.red}`,color:B.red,borderRadius:10,padding:"12px 28px",fontSize:16,fontWeight:900,cursor:"pointer",fontFamily:"inherit",letterSpacing:".1em",transition:"all .15s"}}
              onMouseEnter={e=>{e.currentTarget.style.background=B.red;e.currentTarget.style.color=B.white;}}
              onMouseLeave={e=>{e.currentTarget.style.background=B.redBg;e.currentTarget.style.color=B.red;}}>
              {s}
            </button>
          ))}
          <button onClick={()=>setFSuc([...opts.sucs])} style={{background:"transparent",border:`1px solid ${B.border}`,color:B.g3,borderRadius:10,padding:"12px 24px",fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>Ver todas juntas</button>
        </div>
      </div>
    </div>
  );

  // ══════════════════════════════════════════════
  // VIEWS
  // ══════════════════════════════════════════════

  const ViewResumen=()=>(<div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
      <KCard icon="💰" label="Ingresos Netos" value={f$(kpi.ing)} sub={`${fN(kpi.lineas)} líneas · ${fN(kpi.clis)} clientes`} color={B.red}/>
      <KCard icon="📊" label="Costo Total" value={f$(kpi.cst)} sub="costo de ventas" color={B.g3}/>
      <KCard icon="📈" label="Margen Bruto" value={f$(kpi.mgn)} sub={`${kpi.mgnP.toFixed(1)}% sobre ingresos`} color={kpi.mgnP>20?B.green:kpi.mgnP>10?B.amber:B.red} warn={kpi.mgnP<5}/>
      <KCard icon="🏢" label="Clientes" value={kpi.clis} sub={`${fN(kpi.vol)} unidades`} color={B.amber}/>
      <KCard icon="📅" label="Período" value={fSuc.join("+")||"Todas"} sub={periodLabel} color={B.blue}/>
    </div>

    {analysis.alerts.filter(a=>a.icon==="🔴").slice(0,2).map((a,i)=>(
      <div key={i} style={{background:B.redBg,border:`1px solid ${B.red}`,borderRadius:8,padding:"10px 14px",display:"flex",gap:10,alignItems:"flex-start"}}>
        <span style={{fontSize:18,flexShrink:0}}>🔴</span>
        <div><div style={{color:B.red,fontWeight:800,fontSize:12}}>{a.t}</div><div style={{color:B.g1,fontSize:11,marginTop:2,lineHeight:1.5}}>{a.b}</div></div>
      </div>
    ))}

    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
      <Pnl><ST sub="Ingresos por mes fiscal">Tendencia fiscal</ST>
        <ResponsiveContainer width="100%" height={180}><LineChart data={monthlyTrend} margin={{top:4,right:16,left:0,bottom:4}}>
          <CartesianGrid strokeDasharray="3 3" stroke={B.border}/><XAxis dataKey="mes" stroke={B.g3} tick={{fontSize:10}}/><YAxis stroke={B.g3} tick={{fontSize:9}} tickFormatter={f$} width={60}/><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/>
          {fyList.map((fy,i)=>(<Line key={fy} type="monotone" dataKey={fy} stroke={B.ch[i%B.ch.length]} strokeWidth={2} dot={{r:3}} connectNulls/>))}
        </LineChart></ResponsiveContainer>
      </Pnl>
      <Pnl><ST sub="Click para filtrar">Pareto de clientes</ST>
        {[{s:"TOP",c:B.amberLt},{s:"MED",c:B.red},{s:"LOW",c:B.g3}].map(({s,c})=>{
          const sc=allClients.filter(x=>x.seg===s);const pct=sc.reduce((sum,x)=>sum+(x.pct||0),0);
          return(<div key={s} style={{marginBottom:10,cursor:"pointer"}} onClick={()=>toggle(setFSeg,s)}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}><span style={{color:c,fontWeight:700}}>{s} · {sc.length} clientes</span><span style={{color:B.white}}>{pct.toFixed(1)}%</span></div>
            <div style={{height:7,background:B.border,borderRadius:4,overflow:"hidden"}}><div style={{height:"100%",width:`${clamp(pct,0,100)}%`,background:c,borderRadius:4,transition:"width .5s"}}/></div>
            <div style={{display:"flex",gap:4,flexWrap:"wrap",marginTop:4}}>{sc.slice(0,3).map(x=>(<span key={x.id} style={{fontSize:9,color:B.g3,background:B.panel,borderRadius:3,padding:"1px 5px",border:`1px solid ${B.border}`}}>{x.nm.split(" ").slice(0,2).join(" ")}</span>))}</div>
          </div>);
        })}
      </Pnl>
    </div>

    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
      <Pnl><ST sub={`Top ${pieData.length} sectores`}>Mix sectorial</ST>
        <ResponsiveContainer width="100%" height={210}><PieChart>
          <Pie data={pieData} dataKey="ing" nameKey="sec" cx="50%" cy="50%" outerRadius={82} innerRadius={40}
            label={({percent})=>percent>0.04?`${(percent*100).toFixed(0)}%`:""} labelLine={{stroke:B.border,strokeWidth:1}}>
            {pieData.map((_,i)=><Cell key={i} fill={B.ch[i%B.ch.length]}/>)}
          </Pie><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:9}} layout="horizontal"/>
        </PieChart></ResponsiveContainer>
      </Pnl>
      <Pnl><ST sub="Ingresos y margen por familia">Rentabilidad por línea</ST>
        <ResponsiveContainer width="100%" height={210}><BarChart data={allFamilia.filter(f=>f.ing>0).slice(0,7)} layout="vertical" margin={{left:0,right:20,top:4,bottom:4}}>
          <CartesianGrid strokeDasharray="3 3" stroke={B.border} horizontal={false}/><XAxis type="number" stroke={B.g3} tick={{fontSize:9}} tickFormatter={f$} width={60}/><YAxis type="category" dataKey="f2" stroke={B.g3} tick={{fontSize:9}} width={138}/><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/>
          <Bar dataKey="ing" name="Ingresos" fill={B.red} radius={[0,3,3,0]}/><Bar dataKey="mgn" name="Margen" radius={[0,3,3,0]}>{allFamilia.filter(f=>f.ing>0).slice(0,7).map((f,i)=><Cell key={i} fill={f.mgnP<0?B.g4:B.greenLt}/>)}</Bar>
        </BarChart></ResponsiveContainer>
      </Pnl>
    </div>
  </div>);

  const ViewClientes=()=>(<div style={{display:"grid",gridTemplateColumns:selCli?"1fr 340px":"1fr",gap:14,alignItems:"start"}}>
    <Pnl style={{overflow:"auto"}}><ST sub={`${filteredClients.length} clientes`}>Cartera de clientes</ST>
      <div style={{display:"flex",gap:5,marginBottom:10,flexWrap:"wrap"}}>
        {["ALL","TOP","MED","LOW"].map(s=>(<Pill key={s} label={s==="ALL"?"Todos":s} active={s==="ALL"?!fSeg.length:fSeg.includes(s)} color={s==="TOP"?B.amberLt:s==="MED"?B.blueLt:B.g3} onClick={()=>s==="ALL"?setFSeg([]):toggle(setFSeg,s)}/>))}
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar..." style={{marginLeft:"auto",background:B.panel,border:`1px solid ${B.border}`,borderRadius:5,color:B.white,fontFamily:"inherit",fontSize:11,padding:"3px 9px",outline:"none",width:160}}/>
      </div>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,tableLayout:"fixed"}}>
        <colgroup><col style={{width:26}}/><col style={{width:"24%"}}/><col style={{width:"12%"}}/><col style={{width:76}}/><col style={{width:72}}/><col style={{width:54}}/><col style={{width:32}}/><col style={{width:40}}/><col/></colgroup>
        <thead><tr style={{borderBottom:`1px solid ${B.border}`}}>{["#","Cliente","Sector","Ingresos","Margen","Mgn%","Lín.","Seg.","Ejecutivo"].map((h,i)=>(<th key={h} style={{padding:"5px 7px",textAlign:i>=3&&i<=4?"right":"left",color:B.g3,fontSize:9,letterSpacing:".07em",textTransform:"uppercase"}}>{h}</th>))}</tr></thead>
        <tbody>{filteredClients.map((c,i)=>(<tr key={c.id} onClick={()=>setSelCli(selCli?.id===c.id?null:c)} style={{cursor:"pointer",borderBottom:`1px solid ${B.border}22`,background:selCli?.id===c.id?`${B.red}10`:"transparent",borderLeft:selCli?.id===c.id?`2px solid ${B.red}`:"2px solid transparent"}}>
          <td style={{padding:"7px 7px",color:B.g3,fontSize:9}}>{i+1}</td>
          <td style={{padding:"7px 7px",overflow:"hidden"}}><div style={{fontWeight:700,color:B.white,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.nm}</div><div style={{fontSize:9,color:B.g3}}>{c.id}</div></td>
          <td style={{padding:"7px 7px",color:B.g3,fontSize:10,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.sector}</td>
          <td style={{padding:"7px 7px",fontWeight:700,color:B.red,textAlign:"right"}}>{f$(c.ing)}</td>
          <td style={{padding:"7px 7px",color:c.mgn>=0?B.greenLt:B.red,fontWeight:600,textAlign:"right"}}>{f$(c.mgn)}</td>
          <td style={{padding:"7px 7px",textAlign:"center"}}><MBadge v={c.mgnP}/></td>
          <td style={{padding:"7px 7px",color:B.g3,textAlign:"center"}}>{c.lineas}</td>
          <td style={{padding:"7px 7px",textAlign:"center"}}><SBadge s={c.seg}/></td>
          <td style={{padding:"7px 7px",color:B.amberLt,fontSize:10,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{(()=>{const k=Object.keys(cartera).find(k=>cartera[k]?.includes(c.id));return k?`★ ${repNames[k]||k}`:repNames[c.rep]||c.rep||"—";})()}</td>
        </tr>))}</tbody>
      </table>
    </Pnl>
    {selCli&&(<div style={{display:"flex",flexDirection:"column",gap:10}}>
      <Pnl style={{borderTop:`2px solid ${selCli.seg==="TOP"?B.amberLt:selCli.seg==="MED"?B.blueLt:B.g3}`}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}><div><div style={{fontWeight:800,fontSize:13,color:B.white}}>{selCli.nm}</div><div style={{fontSize:10,color:B.g3,marginTop:2}}>{selCli.id} · {selCli.sector}</div></div><button onClick={()=>setSelCli(null)} style={{background:"none",border:"none",color:B.g3,cursor:"pointer",fontSize:16,padding:2}}>✕</button></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7}}>
          {[{l:"Ingresos",v:f$(selCli.ing),c:B.red},{l:"Margen",v:f$(selCli.mgn),c:selCli.mgn>=0?B.greenLt:B.red},{l:"Margen %",v:fP(selCli.mgnP),c:selCli.mgnP<0?B.red:selCli.mgnP>25?B.green:B.amber},{l:"% del total",v:`${selCli.pct?.toFixed(1)||"?"}%`,c:B.amberLt},{l:"Líneas",v:selCli.lineas},{l:"Meses activo",v:selCli.meses||"—"}].map(({l,v,c})=>(
            <div key={l} style={{background:B.panel,borderRadius:6,padding:"8px 10px",border:`1px solid ${B.border}`}}><div style={{fontSize:9,color:B.g3,marginBottom:2}}>{l}</div><div style={{fontSize:13,fontWeight:800,color:c||B.white}}>{v}</div></div>
          ))}
        </div>
        {selCli.mgn<0&&<div style={{marginTop:8,padding:"6px 10px",background:B.redBg,border:`1px solid ${B.red}`,borderRadius:5,fontSize:10,color:B.red}}>⚠️ Margen negativo — revisar condición comercial</div>}
        {(selCli.pct||0)>15&&<div style={{marginTop:6,padding:"6px 10px",background:B.amberDim,border:`1px solid ${B.amber}`,borderRadius:5,fontSize:10,color:B.amberLt}}>🔑 Cliente estratégico: {selCli.pct?.toFixed(1)}% del ingreso</div>}
      </Pnl>
      <Pnl><ST sub="Asignar a ejecutivo">Cartera</ST><div style={{display:"flex",flexWrap:"wrap",gap:5}}>
        {opts.reps.map(rep=>{const a=cartera[rep]?.includes(selCli.id);return(<button key={rep} onClick={()=>setCartera(prev=>{const n={...prev};if(!n[rep])n[rep]=[];if(a)n[rep]=n[rep].filter(id=>id!==selCli.id);else{Object.keys(n).forEach(r=>{n[r]=(n[r]||[]).filter(id=>id!==selCli.id);});n[rep]=[...n[rep],selCli.id];}return n;})} style={{background:a?B.amberDim:"transparent",border:`1px solid ${a?B.amberLt:B.border}`,color:a?B.amberLt:B.g3,borderRadius:5,padding:"4px 10px",fontSize:10,cursor:"pointer",fontFamily:"inherit",fontWeight:a?700:400}}>{a?"★ ":""}{repNames[rep]||`Rep ${rep}`}</button>);})}
      </div></Pnl>
    </div>)}
  </div>);

  const ViewProductos=()=>(<div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
      <Pnl><ST sub="Click para filtrar">Familias de producto</ST>{allFamilia.filter(f=>Math.abs(f.ing)>0).map((f,i)=>(<div key={f.f2} style={{marginBottom:9,paddingLeft:9,borderLeft:`2px solid ${f.mgnP<0?B.red:B.ch[i%B.ch.length]}`,cursor:"pointer"}} onClick={()=>toggle(setFFam,f.f2)}><div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:2,gap:8}}><span style={{color:B.white,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>{f.f2}</span><div style={{display:"flex",gap:7,flexShrink:0}}><span style={{color:B.red,fontWeight:700}}>{f$(f.ing)}</span><MBadge v={f.mgnP}/></div></div><div style={{height:4,background:B.border,borderRadius:2,overflow:"hidden"}}><div style={{height:"100%",width:`${allFamilia[0]?.ing?clamp(f.ing/allFamilia[0].ing*100,0,100):0}%`,background:f.mgnP<0?B.red:B.ch[i%B.ch.length],borderRadius:2}}/></div></div>))}</Pnl>
      <Pnl style={{overflow:"auto"}}><ST>Artículos más vendidos</ST>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,tableLayout:"fixed"}}><colgroup><col style={{width:26}}/><col style={{width:"30%"}}/><col/><col style={{width:70}}/><col style={{width:56}}/><col style={{width:36}}/></colgroup>
          <thead><tr style={{borderBottom:`1px solid ${B.border}`}}>{["#","Código","Familia","Ingresos","Mgn%","Vol"].map((h,i)=>(<th key={h} style={{padding:"5px 7px",textAlign:i===3?"right":"left",color:B.g3,fontSize:9,letterSpacing:".07em",textTransform:"uppercase"}}>{h}</th>))}</tr></thead>
          <tbody>{allArticles.map((a,i)=>(<tr key={a.art} style={{borderBottom:`1px solid ${B.border}22`}}><td style={{padding:"6px 7px",color:B.g3,fontSize:9}}>{i+1}</td><td style={{padding:"6px 7px",color:B.red,fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.art}</td><td style={{padding:"6px 7px",color:B.g3,fontSize:10,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.f2}</td><td style={{padding:"6px 7px",fontWeight:700,color:B.white,textAlign:"right"}}>{f$(a.ing)}</td><td style={{padding:"6px 7px",textAlign:"center"}}><MBadge v={a.mgnP}/></td><td style={{padding:"6px 7px",color:B.g3,textAlign:"right"}}>{a.vol.toFixed(0)}</td></tr>))}</tbody>
        </table>
      </Pnl>
    </div>
    <Pnl><ST>Comparativa ingresos vs margen</ST><ResponsiveContainer width="100%" height={200}><BarChart data={allFamilia.filter(f=>f.ing>0)} margin={{top:5,right:20,left:0,bottom:52}}><CartesianGrid strokeDasharray="3 3" stroke={B.border}/><XAxis dataKey="f2" stroke={B.g3} tick={{fontSize:9,angle:-20,textAnchor:"end"}} interval={0} height={64}/><YAxis stroke={B.g3} tick={{fontSize:9}} tickFormatter={f$} width={60}/><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/><Bar dataKey="ing" name="Ingresos" fill={B.red} radius={[4,4,0,0]}/><Bar dataKey="mgn" name="Margen" radius={[4,4,0,0]}>{allFamilia.filter(f=>f.ing>0).map((f,i)=><Cell key={i} fill={f.mgnP<0?B.g4:B.greenLt}/>)}</Bar></BarChart></ResponsiveContainer></Pnl>
  </div>);

  const ViewSectores=()=>(<div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
      <Pnl><ST>Ingresos por industria</ST><ResponsiveContainer width="100%" height={240}><BarChart data={allSector.filter(s=>s.ing>0)} margin={{top:5,right:20,left:0,bottom:52}}><CartesianGrid strokeDasharray="3 3" stroke={B.border}/><XAxis dataKey="sec" stroke={B.g3} tick={{fontSize:9,angle:-20,textAnchor:"end"}} interval={0} height={64}/><YAxis stroke={B.g3} tick={{fontSize:9}} tickFormatter={f$} width={60}/><Tooltip content={<TT/>}/><Bar dataKey="ing" name="Ingresos" radius={[4,4,0,0]} onClick={d=>d&&toggle(setFSec,d.sec)}>{allSector.filter(s=>s.ing>0).map((_,i)=><Cell key={i} fill={B.ch[i%B.ch.length]} cursor="pointer"/>)}</Bar></BarChart></ResponsiveContainer></Pnl>
      <Pnl><ST>Mix</ST><ResponsiveContainer width="100%" height={240}><PieChart><Pie data={pieData} dataKey="ing" nameKey="sec" cx="50%" cy="50%" outerRadius={90} innerRadius={45} label={({percent})=>percent>0.04?`${(percent*100).toFixed(0)}%`:""} labelLine={{stroke:B.border,strokeWidth:1}}>{pieData.map((_,i)=><Cell key={i} fill={B.ch[i%B.ch.length]}/>)}</Pie><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:9}}/></PieChart></ResponsiveContainer></Pnl>
    </div>
    {allSector.filter(s=>s.ing>0).map((s,si)=>{const cs=allClients.filter(c=>c.sector===s.sec&&c.ing>0).slice(0,12);if(!cs.length)return null;return(<Pnl key={si} style={{borderLeft:`3px solid ${B.ch[si%B.ch.length]}`,cursor:"pointer"}} onClick={()=>toggle(setFSec,s.sec)}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><div style={{color:B.ch[si%B.ch.length],fontWeight:800,fontSize:12}}>{s.sec}</div><div style={{display:"flex",gap:14,fontSize:11}}><span style={{color:B.red,fontWeight:700}}>{f$(s.ing)}</span><span style={{color:B.g3}}>{s.clientes} cl · {s.lineas} lín</span></div></div><div style={{display:"flex",flexWrap:"wrap",gap:6}}>{cs.map(c=>(<div key={c.id} onClick={e=>{e.stopPropagation();setSelCli(c);setView("CLIENTES");}} style={{background:B.panel,borderRadius:7,padding:"5px 10px",border:`1px solid ${B.border}`,cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.borderColor=B.red} onMouseLeave={e=>e.currentTarget.style.borderColor=B.border}><div style={{fontSize:10,color:B.white,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:160}}>{c.nm}</div><div style={{display:"flex",gap:6,marginTop:2,alignItems:"center"}}><span style={{fontSize:10,color:B.red,fontWeight:700}}>{f$(c.ing)}</span><MBadge v={c.mgnP}/><SBadge s={c.seg}/></div></div>))}</div></Pnl>);})}
  </div>);

  const ViewTendencias=()=>(<div style={{display:"flex",flexDirection:"column",gap:14}}>
    {fyList.length>1?(<>
      <Pnl><ST sub="Comparativa Abr→Mar">Evolución fiscal mensual</ST><ResponsiveContainer width="100%" height={240}><LineChart data={monthlyTrend} margin={{top:5,right:20,left:0,bottom:5}}><CartesianGrid strokeDasharray="3 3" stroke={B.border}/><XAxis dataKey="mes" stroke={B.g3} tick={{fontSize:10}}/><YAxis stroke={B.g3} tick={{fontSize:9}} tickFormatter={f$} width={62}/><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/>{fyList.map((fy,i)=><Line key={fy} type="monotone" dataKey={fy} stroke={B.ch[i%B.ch.length]} strokeWidth={2} dot={{r:3}} connectNulls/>)}</LineChart></ResponsiveContainer></Pnl>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
        <Pnl><ST>Trimestral</ST><ResponsiveContainer width="100%" height={200}><BarChart data={quarterData} margin={{top:5,right:20,left:0,bottom:5}}><CartesianGrid strokeDasharray="3 3" stroke={B.border}/><XAxis dataKey="q" stroke={B.g3} tick={{fontSize:10}}/><YAxis stroke={B.g3} tick={{fontSize:9}} tickFormatter={f$} width={62}/><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:10}}/>{fyList.map((fy,i)=><Bar key={fy} dataKey={fy} fill={B.ch[i%B.ch.length]} radius={[4,4,0,0]}/>)}</BarChart></ResponsiveContainer></Pnl>
        <Pnl><ST>Totales anuales</ST>{fyTotals.map((fy,i)=>{const prev=fyTotals[i-1];const d=prev?(fy.ing-prev.ing)/prev.ing*100:null;return(<div key={fy.fy} style={{marginBottom:10,background:B.panel,borderRadius:8,padding:"11px 14px",border:`1px solid ${B.border}`,borderLeft:`3px solid ${B.ch[i%B.ch.length]}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{color:B.ch[i%B.ch.length],fontWeight:800,fontSize:12}}>{fy.fy}</span>{d!=null&&<span style={{color:d>=0?B.greenLt:B.red,fontWeight:700,fontSize:11}}>{d>=0?"▲":"▼"}{Math.abs(d).toFixed(1)}% YoY</span>}</div><div style={{display:"flex",gap:20,marginTop:6}}><div><div style={{fontSize:9,color:B.g3}}>INGRESOS</div><div style={{fontSize:14,fontWeight:800,color:B.red}}>{f$(fy.ing)}</div></div><div><div style={{fontSize:9,color:B.g3}}>MARGEN</div><div style={{fontSize:14,fontWeight:800,color:fy.mgnP>15?B.greenLt:B.amberLt}}>{fP(fy.mgnP)}</div></div></div></div>);})}</Pnl>
      </div>
    </>):(<Pnl style={{textAlign:"center",padding:"48px 32px"}}><div style={{fontSize:36,marginBottom:12}}>📊</div><div style={{color:B.white,fontWeight:700,fontSize:14,marginBottom:8}}>Carga más Excels para comparativas</div></Pnl>)}
  </div>);

  // ── EJECUTIVOS — full profile ─────────────────────────────────────────────
  const ViewEjecutivos=()=>(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      {/* Config names */}
      <Pnl><ST sub="Asigna nombres reales — guardado localmente">Configurar ejecutivos</ST>
        <div style={{display:"flex",flexWrap:"wrap",gap:10}}>
          {opts.reps.map(rep=>(<div key={rep} style={{display:"flex",alignItems:"center",gap:8,background:B.panel,borderRadius:7,padding:"7px 12px",border:`1px solid ${selRep===rep?B.red:B.border}`,cursor:"pointer"}} onClick={()=>setSelRep(selRep===rep?null:rep)}>
            <span style={{color:selRep===rep?B.red:B.g3,fontSize:10,minWidth:52,flexShrink:0}}>Rep {rep}</span>
            <span style={{color:B.border}}>→</span>
            <input value={repNames[rep]||""} placeholder="Nombre…" onChange={e=>{e.stopPropagation();setRepNames(p=>({...p,[rep]:e.target.value}));}} onClick={e=>e.stopPropagation()}
              style={{background:"transparent",border:"none",borderBottom:`1px solid ${B.border}`,color:B.amberLt,fontFamily:"inherit",fontSize:11,padding:"2px 4px",outline:"none",width:140}}/>
            <span style={{fontSize:9,color:selRep===rep?B.red:B.g3,marginLeft:4}}>{selRep===rep?"▾ perfil":"ver"}</span>
          </div>))}
        </div>
      </Pnl>

      {/* Overview ranking */}
      {!selRep&&(<>
        <Pnl><ST sub="Click en ejecutivo para ver perfil completo">Ranking de ejecutivos</ST>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={allReps.map(r=>({...r,nm:repNames[r.rep]||`Rep ${r.rep}`}))} margin={{top:5,right:20,left:0,bottom:52}} onClick={d=>d?.activePayload&&setSelRep(d.activePayload[0]?.payload?.rep)}>
              <CartesianGrid strokeDasharray="3 3" stroke={B.border}/><XAxis dataKey="nm" stroke={B.g3} tick={{fontSize:9,angle:-15,textAnchor:"end"}} interval={0} height={64}/><YAxis stroke={B.g3} tick={{fontSize:9}} tickFormatter={f$} width={60}/><Tooltip content={<TT/>}/>
              <Bar dataKey="ing" name="Ingresos" radius={[4,4,0,0]} cursor="pointer">{allReps.map((_,i)=><Cell key={i} fill={B.ch[i%B.ch.length]}/>)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </Pnl>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))",gap:10}}>
          {allReps.map((r,i)=>(<div key={r.rep} style={{background:B.panel,borderRadius:9,padding:"12px 14px",border:`1px solid ${B.border}`,borderLeft:`3px solid ${B.ch[i%B.ch.length]}`,cursor:"pointer"}} onClick={()=>setSelRep(r.rep)}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
              <div style={{fontWeight:800,color:B.ch[i%B.ch.length],fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{repNames[r.rep]||`Rep ${r.rep}`}</div>
              <div style={{color:B.red,fontWeight:700,fontSize:12,flexShrink:0,marginLeft:8}}>{f$(r.ing)}</div>
            </div>
            <div style={{display:"flex",gap:14,fontSize:10,color:B.g3,marginBottom:6}}><span>{r.clientes} clientes</span><span>{r.lineas} líneas</span><span>{kpi.ing?(r.ing/kpi.ing*100).toFixed(1):0}%</span></div>
            <div style={{height:3,background:B.border,borderRadius:2}}><div style={{height:"100%",width:`${allReps[0]?.ing?clamp(r.ing/allReps[0].ing*100,0,100):0}%`,background:B.ch[i%B.ch.length],borderRadius:2}}/></div>
            <div style={{marginTop:8,fontSize:9,color:B.blue}}>↗ Click para ver perfil completo</div>
          </div>))}
        </div>
      </>)}

      {/* Exec profile */}
      {selRep&&repProfile&&(
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          {/* Header */}
          <div style={{background:`linear-gradient(135deg,${B.redBg},${B.panel})`,border:`1px solid ${B.red}`,borderRadius:10,padding:"16px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
            <div>
              <div style={{fontWeight:900,fontSize:16,color:B.white}}>{repNames[selRep]||`Ejecutivo ${selRep}`}</div>
              <div style={{fontSize:10,color:B.g3,marginTop:3}}>Código SAP: {selRep} · {periodLabel}</div>
            </div>
            <div style={{display:"flex",gap:16,alignItems:"center"}}>
              <div style={{textAlign:"right"}}>
                <div style={{fontSize:22,fontWeight:900,color:B.red}}>{f$(repProfile.kpi.ing)}</div>
                <div style={{fontSize:11,color:repProfile.kpi.mgn>0?B.greenLt:B.red,fontWeight:700}}>{fP(repProfile.kpi.ing?repProfile.kpi.mgn/repProfile.kpi.ing*100:0)} margen</div>
              </div>
              <button onClick={()=>setSelRep(null)} style={{background:"none",border:`1px solid ${B.border}`,color:B.g3,borderRadius:6,padding:"6px 12px",cursor:"pointer",fontFamily:"inherit",fontSize:11}}>← Volver</button>
            </div>
          </div>

          {/* KPIs del ejecutivo */}
          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            <KCard icon="💰" label="Ingresos" value={f$(repProfile.kpi.ing)} sub={`${repProfile.kpi.lineas} líneas`} color={B.red}/>
            <KCard icon="📈" label="Margen" value={f$(repProfile.kpi.mgn)} sub={fP(repProfile.kpi.ing?repProfile.kpi.mgn/repProfile.kpi.ing*100:0)} color={repProfile.kpi.mgn>0?B.green:B.red}/>
            <KCard icon="🏢" label="Clientes" value={repProfile.kpi.clis} sub="en cartera" color={B.amber}/>
            <KCard icon="📊" label="% del total" value={fP(kpi.ing?repProfile.kpi.ing/kpi.ing*100:0)} sub="participación" color={B.blue}/>
          </div>

          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
            {/* Monthly trend */}
            <Pnl><ST sub="Ingresos por mes fiscal">Tendencia mensual</ST>
              <ResponsiveContainer width="100%" height={160}><AreaChart data={repProfile.monthly} margin={{top:4,right:16,left:0,bottom:4}}>
                <defs><linearGradient id="rg2" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={B.red} stopOpacity={.35}/><stop offset="95%" stopColor={B.red} stopOpacity={0}/></linearGradient></defs>
                <CartesianGrid strokeDasharray="3 3" stroke={B.border}/><XAxis dataKey="mes" stroke={B.g3} tick={{fontSize:9}}/><YAxis stroke={B.g3} tick={{fontSize:9}} tickFormatter={f$} width={56}/><Tooltip content={<TT/>}/>
                <Area type="monotone" dataKey="ing" stroke={B.red} fill="url(#rg2)" strokeWidth={2} name="Ingresos"/>
              </AreaChart></ResponsiveContainer>
            </Pnl>

            {/* Pareto of their clients */}
            <Pnl><ST sub="Pareto de su cartera">Concentración de cartera</ST>
              {[{s:"TOP",c:B.amberLt},{s:"MED",c:B.red},{s:"LOW",c:B.g3}].map(({s,c})=>{
                const sc=repProfile.clients.filter(x=>x.repSeg===s);
                const pct=sc.reduce((sum,x)=>sum+(x.repPct||0),0);
                return(<div key={s} style={{marginBottom:9}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
                    <span style={{color:c,fontWeight:700}}>{s} · {sc.length} clientes</span>
                    <span style={{color:B.white}}>{pct.toFixed(1)}%</span>
                  </div>
                  <div style={{height:6,background:B.border,borderRadius:3,overflow:"hidden"}}><div style={{height:"100%",width:`${clamp(pct,0,100)}%`,background:c,borderRadius:3}}/></div>
                  <div style={{display:"flex",gap:4,flexWrap:"wrap",marginTop:3}}>{sc.slice(0,3).map(x=>(<span key={x.id} style={{fontSize:9,color:B.g3,background:B.panel,borderRadius:3,padding:"1px 5px",border:`1px solid ${B.border}`}}>{x.nm.split(" ").slice(0,2).join(" ")}</span>))}</div>
                </div>);
              })}
            </Pnl>
          </div>

          {/* Client table */}
          <Pnl style={{overflow:"auto"}}><ST sub={`${repProfile.clients.length} clientes en cartera — ordenados por ingreso`}>Cartera completa</ST>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,tableLayout:"fixed"}}>
              <colgroup><col style={{width:26}}/><col style={{width:"26%"}}/><col style={{width:"13%"}}/><col style={{width:76}}/><col style={{width:72}}/><col style={{width:54}}/><col style={{width:34}}/><col/></colgroup>
              <thead><tr style={{borderBottom:`1px solid ${B.border}`}}>{["#","Cliente","Sector","Ingresos","Margen","Mgn%","Lín.","Seg."].map((h,i)=>(<th key={h} style={{padding:"5px 7px",textAlign:i>=3&&i<=4?"right":"left",color:B.g3,fontSize:9,letterSpacing:".07em",textTransform:"uppercase"}}>{h}</th>))}</tr></thead>
              <tbody>{repProfile.clients.map((c,i)=>(<tr key={c.id} style={{borderBottom:`1px solid ${B.border}22`}} onClick={()=>{setSelCli(c);setView("CLIENTES");}}>
                <td style={{padding:"7px 7px",color:B.g3,fontSize:9,cursor:"pointer"}}>{i+1}</td>
                <td style={{padding:"7px 7px",overflow:"hidden",cursor:"pointer"}}><div style={{fontWeight:700,color:B.white,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.nm}</div><div style={{fontSize:9,color:B.g3}}>{c.id}</div></td>
                <td style={{padding:"7px 7px",color:B.g3,fontSize:10,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.sector}</td>
                <td style={{padding:"7px 7px",fontWeight:700,color:B.red,textAlign:"right"}}>{f$(c.ing)}</td>
                <td style={{padding:"7px 7px",color:c.mgn>=0?B.greenLt:B.red,fontWeight:600,textAlign:"right"}}>{f$(c.mgn)}</td>
                <td style={{padding:"7px 7px",textAlign:"center"}}><MBadge v={c.mgnP}/></td>
                <td style={{padding:"7px 7px",color:B.g3,textAlign:"center"}}>{c.lineas}</td>
                <td style={{padding:"7px 7px",textAlign:"center"}}><SBadge s={c.repSeg} color={c.repSeg==="TOP"?B.amberLt:c.repSeg==="MED"?B.blueLt:B.g3}/></td>
              </tr>))}</tbody>
            </table>
          </Pnl>
        </div>
      )}
    </div>
  );

  const ViewPivot=()=>{const{rows:pr,cols:pc}=pivotData;const vFmt=PVALS[pVal]?.f||fN;const gt=pr.reduce((s,r)=>s+r.total,0);return(<div style={{display:"flex",flexDirection:"column",gap:14}}>
    <Pnl><ST sub="Configura dimensiones — se actualiza al instante">Constructor tabla dinámica</ST>
      <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
        {[{label:"FILAS",val:pRows,set:setPRows},{label:"COLUMNAS",val:pCols,set:setPCols}].map(({label,val,set})=>(<div key={label}><div style={{fontSize:9,color:B.g3,letterSpacing:".1em",fontWeight:700,marginBottom:6}}>{label}</div><div style={{display:"flex",flexWrap:"wrap",gap:4}}>{Object.entries(PDIMS).map(([k,d])=>(<Pill key={k} label={d} active={val===k} color={B.blue} onClick={()=>set(k)}/>))}</div></div>))}
        <div><div style={{fontSize:9,color:B.g3,letterSpacing:".1em",fontWeight:700,marginBottom:6}}>VALOR</div><div style={{display:"flex",flexWrap:"wrap",gap:4}}>{Object.entries(PVALS).map(([k,d])=>(<Pill key={k} label={d.l} active={pVal===k} color={B.red} onClick={()=>setPVal(k)}/>))}</div></div>
        <div><div style={{fontSize:9,color:B.g3,letterSpacing:".1em",fontWeight:700,marginBottom:6}}>ORDEN</div><div style={{display:"flex",gap:4}}><Pill label="Mayor→Menor" active={pSort==="desc"} color={B.amber} onClick={()=>setPSort("desc")}/><Pill label="Menor→Mayor" active={pSort==="asc"} color={B.amber} onClick={()=>setPSort("asc")}/></div></div>
      </div>
    </Pnl>
    <Pnl style={{overflow:"auto"}}><ST sub={`${pr.length} filas · ${pc.length} col · ${PVALS[pVal]?.l}`}>{PDIMS[pRows]||"?"} × {PDIMS[pCols]||"?"}</ST>
      {pr.length===0?(<div style={{color:B.g3,fontSize:12,padding:"20px 0",textAlign:"center"}}>Sin datos</div>):(
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:450}}>
          <thead><tr style={{borderBottom:`2px solid ${B.red}`}}><th style={{padding:"7px 10px",textAlign:"left",color:B.g3,fontSize:9,letterSpacing:".08em",textTransform:"uppercase",background:B.panel,position:"sticky",left:0,zIndex:1,minWidth:140}}>{PDIMS[pRows]}</th>{pc.map(c=><th key={c} style={{padding:"7px 10px",textAlign:"right",color:B.g3,fontSize:9,background:B.panel,whiteSpace:"nowrap"}}>{c}</th>)}<th style={{padding:"7px 10px",textAlign:"right",color:B.red,fontSize:9,background:B.panel}}>TOTAL</th><th style={{padding:"7px 10px",textAlign:"right",color:B.g3,fontSize:9,background:B.panel}}>%</th></tr></thead>
          <tbody>
            {pr.map((row,ri)=>{const rp=gt?row.total/gt*100:0;return(<tr key={ri} style={{borderBottom:`1px solid ${B.border}22`}} onMouseEnter={e=>e.currentTarget.style.background=`${B.red}08`} onMouseLeave={e=>e.currentTarget.style.background="transparent"}><td style={{padding:"7px 10px",fontWeight:600,color:B.white,background:B.panel,position:"sticky",left:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:180}}>{row.rowKey}</td>{pc.map(c=>{const v=row.vals[c]||0;const ct=pr.reduce((s,r)=>s+(r.vals[c]||0),0);const heat=ct?v/ct:0;return(<td key={c} style={{padding:"7px 10px",textAlign:"right",fontWeight:v>0?600:400,color:v>0?B.white:B.g4,background:v>0?`rgba(204,0,0,${heat*0.2})`:"transparent",whiteSpace:"nowrap"}}>{v?vFmt(v):"—"}</td>);})}<td style={{padding:"7px 10px",textAlign:"right",fontWeight:800,color:B.red,whiteSpace:"nowrap"}}>{vFmt(row.total)}</td><td style={{padding:"7px 10px",textAlign:"right"}}><div style={{display:"flex",alignItems:"center",gap:5,justifyContent:"flex-end"}}><div style={{width:32,height:5,background:B.border,borderRadius:3,overflow:"hidden"}}><div style={{height:"100%",width:`${clamp(rp,0,100)}%`,background:B.red,borderRadius:3}}/></div><span style={{color:B.g2,fontSize:10,minWidth:32,textAlign:"right"}}>{rp.toFixed(1)}%</span></div></td></tr>);})}
            <tr style={{borderTop:`2px solid ${B.red}`,background:B.panel}}><td style={{padding:"8px 10px",fontWeight:800,color:B.red,position:"sticky",left:0,background:B.panel}}>TOTAL</td>{pc.map(c=>{const ct=pr.reduce((s,r)=>s+(r.vals[c]||0),0);return <td key={c} style={{padding:"8px 10px",textAlign:"right",fontWeight:800,color:B.amberLt,whiteSpace:"nowrap"}}>{ct?vFmt(ct):"—"}</td>;})} <td style={{padding:"8px 10px",textAlign:"right",fontWeight:900,color:B.red,fontSize:12,whiteSpace:"nowrap"}}>{vFmt(gt)}</td><td style={{padding:"8px 10px",textAlign:"right",color:B.g3,fontSize:10}}>100%</td></tr>
          </tbody>
        </table>
      )}
    </Pnl>
  </div>);};

  // ── ANÁLISIS — 1 page, filters-aware ─────────────────────────────────────
  const ViewAnalisis=()=>{const{alerts,opps,actions}=analysis;return(<div style={{display:"flex",flexDirection:"column",gap:14}}>
    {/* Header */}
    <div style={{background:`linear-gradient(135deg,${B.redBg} 0%,${B.panel} 100%)`,border:`1px solid ${B.red}44`,borderRadius:10,padding:"16px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
      <div>
        <div style={{fontWeight:900,fontSize:15,color:B.white}}>ANÁLISIS EJECUTIVO</div>
        <div style={{fontSize:10,color:B.g3,marginTop:3}}>{periodLabel} · {new Date().toLocaleDateString("es-PE")}</div>
      </div>
      <div style={{display:"flex",gap:20,flexWrap:"wrap"}}>
        {[{l:"Ingresos",v:f$(kpi.ing),c:B.red},{l:"Margen",v:fP(kpi.mgnP),c:kpi.mgnP>20?B.greenLt:kpi.mgnP>10?B.amberLt:B.red},{l:"Clientes",v:fN(kpi.clis),c:B.white},{l:"Líneas",v:fN(kpi.lineas),c:B.g2}].map(x=>(<div key={x.l} style={{textAlign:"right"}}><div style={{fontSize:9,color:B.g3}}>{x.l}</div><div style={{fontSize:14,fontWeight:800,color:x.c}}>{x.v}</div></div>))}
      </div>
    </div>

    {/* 2-col alerts + opps */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
      <Pnl style={{borderLeft:`3px solid ${B.red}`}}>
        <ST>🔴 Alertas críticas</ST>
        {alerts.length===0?<div style={{color:B.greenLt,fontSize:11}}>✅ Sin alertas en este período/sucursal</div>:
        alerts.map((a,i)=>(<div key={i} style={{marginBottom:10,paddingLeft:9,borderLeft:`2px solid ${B.red}44`}}><div style={{display:"flex",gap:6,alignItems:"flex-start",marginBottom:2}}><span style={{fontSize:12,flexShrink:0}}>{a.icon}</span><span style={{fontWeight:700,fontSize:11,color:B.white}}>{a.t}</span></div><div style={{fontSize:10,color:B.g1,lineHeight:1.5,paddingLeft:18}}>{a.b}</div></div>))}
      </Pnl>
      <Pnl style={{borderLeft:`3px solid ${B.green}`}}>
        <ST>🟢 Oportunidades</ST>
        {opps.length===0?<div style={{color:B.g3,fontSize:11}}>Sin oportunidades detectadas</div>:
        opps.map((o,i)=>(<div key={i} style={{marginBottom:10,paddingLeft:9,borderLeft:`2px solid ${B.green}44`}}><div style={{display:"flex",gap:6,alignItems:"flex-start",marginBottom:2}}><span style={{fontSize:12,flexShrink:0}}>{o.icon}</span><span style={{fontWeight:700,fontSize:11,color:B.white}}>{o.t}</span></div><div style={{fontSize:10,color:B.g1,lineHeight:1.5,paddingLeft:18}}>{o.b}</div></div>))}
      </Pnl>
    </div>

    {/* Actions + scatter side by side */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
      <Pnl style={{borderLeft:`3px solid ${B.amberLt}`}}>
        <ST>🎯 Prioridades esta semana</ST>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {actions.map((a,i)=>(<div key={i} style={{background:B.panel,borderRadius:7,padding:"10px 12px",border:`1px solid ${B.border}`,display:"flex",gap:10,alignItems:"flex-start"}}><div style={{fontSize:13,fontWeight:900,color:B.white,background:B.red,borderRadius:4,width:22,height:22,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,lineHeight:1}}>{i+1}</div><div style={{fontSize:11,color:B.white,lineHeight:1.6}}>{a}</div></div>))}
        </div>
      </Pnl>
      <Pnl>
        <ST sub="Eje X=Ingresos · Eje Y=Margen%">Mapa de rentabilidad</ST>
        <ResponsiveContainer width="100%" height={260}><ScatterChart margin={{top:10,right:20,left:0,bottom:10}}>
          <CartesianGrid strokeDasharray="3 3" stroke={B.border}/><XAxis type="number" dataKey="ing" name="Ingresos" stroke={B.g3} tick={{fontSize:9}} tickFormatter={f$} width={60}/><YAxis type="number" dataKey="mgnP" name="Margen%" stroke={B.g3} tick={{fontSize:9}} tickFormatter={v=>`${v.toFixed(0)}%`} width={40}/>
          <ZAxis type="number" dataKey="lineas" range={[40,400]}/>
          <Tooltip cursor={{strokeDasharray:"3 3"}} content={({active,payload})=>{if(!active||!payload?.length)return null;const d=payload[0]?.payload;return<div style={{background:B.panel,border:`1px solid ${B.border}`,borderRadius:8,padding:"9px 13px",fontSize:11}}><p style={{color:B.red,fontWeight:700,marginBottom:4}}>{d?.nm}</p><p style={{color:B.white}}>Ingresos: {f$(d?.ing)}</p><p style={{color:d?.mgnP<0?B.red:B.greenLt}}>Margen: {fP(d?.mgnP)}</p></div>;}}/> 
          <Scatter data={allClients.filter(c=>c.ing>0).slice(0,200)} fill={B.red}>{allClients.filter(c=>c.ing>0).slice(0,200).map((c,i)=><Cell key={i} fill={c.mgnP<0?B.red:c.mgnP>30?B.greenLt:B.amberLt}/>)}</Scatter>
        </ScatterChart></ResponsiveContainer>
        <div style={{display:"flex",gap:16,marginTop:4,fontSize:10,color:B.g3}}>
          <span><span style={{color:B.greenLt}}>●</span> Margen &gt;30%</span><span><span style={{color:B.amberLt}}>●</span> 0-30%</span><span><span style={{color:B.red}}>●</span> Negativo</span>
        </div>
      </Pnl>
    </div>
  </div>);};

  const ViewFirebase=()=>(<div style={{display:"flex",flexDirection:"column",gap:14}}>
    <Pnl style={{borderLeft:`3px solid ${B.amber}`}}><ST>📋 Activar modo compartido con Firebase</ST><div style={{display:"flex",flexDirection:"column",gap:12}}>{[{n:1,t:"Crear proyecto",d:"console.firebase.google.com → Crear proyecto → cummins-zona-sur"},{n:2,t:"Activar Storage",d:"Menú → Storage → Comenzar → modo prueba"},{n:3,t:"Credenciales",d:"⚙️ → General → Tus apps → </> → Registrar → copiar firebaseConfig"},{n:4,t:"Pegar en código",d:"Edita src/App.jsx con tu config y redespliega"}].map(s=>(<div key={s.n} style={{display:"flex",gap:12}}><div style={{width:26,height:26,borderRadius:5,background:B.red,color:B.white,fontWeight:900,fontSize:12,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{s.n}</div><div><div style={{fontWeight:700,color:B.white,fontSize:12,marginBottom:2}}>{s.t}</div><div style={{fontSize:11,color:B.g2,lineHeight:1.5}}>{s.d}</div></div></div>))}</div></Pnl>
    <Pnl><ST sub="Archivos cargados">Datos en memoria</ST>{fileAggs.length===0?<div style={{color:B.g3,fontSize:11}}>Ningún archivo — usa + EXCELS</div>:(
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}><thead><tr style={{borderBottom:`1px solid ${B.border}`}}>{["Archivo","Suc.","FY","Ventas","Clientes","Ingresos"].map((h,i)=>(<th key={h} style={{padding:"5px 10px",textAlign:i===5?"right":"left",color:B.g3,fontSize:9,letterSpacing:".08em",textTransform:"uppercase"}}>{h}</th>))}</tr></thead>
        <tbody>{fileAggs.map((f,i)=>(<tr key={i} style={{borderBottom:`1px solid ${B.border}22`}}><td style={{padding:"7px 10px",color:B.white,fontWeight:600,fontSize:11}}>{f.label}</td><td style={{padding:"7px 10px"}}><span style={{background:B.redBg,color:B.red,borderRadius:3,padding:"2px 7px",fontSize:10,fontWeight:700}}>{f.sucursal}</span></td><td style={{padding:"7px 10px",color:B.g3}}>{f.fyl} {f.fml}</td><td style={{padding:"7px 10px",color:B.g3}}>{f.salesRows}</td><td style={{padding:"7px 10px",color:B.g3}}>{f.clients.length}</td><td style={{padding:"7px 10px",fontWeight:700,color:B.red,textAlign:"right"}}>{f$(f.totalIng)}</td></tr>))}
        <tr style={{borderTop:`1px solid ${B.border}`,background:B.panel}}><td style={{padding:"7px 10px",fontWeight:800,color:B.red}} colSpan={3}>TOTAL</td><td style={{padding:"7px 10px",color:B.g3}}>{fN(fileAggs.reduce((s,f)=>s+f.salesRows,0))}</td><td style={{padding:"7px 10px",color:B.g3}}>{fN(new Set(fileAggs.flatMap(f=>f.clients.map(c=>c.id))).size)}</td><td style={{padding:"7px 10px",fontWeight:900,color:B.red,textAlign:"right"}}>{f$(fileAggs.reduce((s,f)=>s+f.totalIng,0))}</td></tr>
      </tbody></table>
    )}</Pnl>
  </div>);

  const renderView=()=>{
    if(noSucSelected)return<NoBanner/>;
    switch(view){
      case "RESUMEN":    return <ViewResumen/>;
      case "CLIENTES":   return <ViewClientes/>;
      case "PRODUCTOS":  return <ViewProductos/>;
      case "SECTORES":   return <ViewSectores/>;
      case "TENDENCIAS": return <ViewTendencias/>;
      case "EJECUTIVOS": return <ViewEjecutivos/>;
      case "TABLA DIN.": return <ViewPivot/>;
      case "ANÁLISIS":   return <ViewAnalisis/>;
      case "FIREBASE":   return <ViewFirebase/>;
      default:           return <ViewResumen/>;
    }
  };

  return(<div style={{minHeight:"100vh",background:B.bg,color:B.white,fontFamily:"'Helvetica Neue',Helvetica,Arial,sans-serif",display:"flex",flexDirection:"column"}}>
    <style>{`*{box-sizing:border-box;margin:0;padding:0}body{background:${B.bg}}::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:${B.border};border-radius:2px}::-webkit-scrollbar-thumb:hover{background:${B.red}}@keyframes fadeIn{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}.fade{animation:fadeIn .2s ease}input::placeholder{color:${B.g4}}td,th{overflow:hidden;text-overflow:ellipsis}.recharts-wrapper,.recharts-surface{overflow:visible!important}`}</style>

    {/* TOPBAR */}
    <div style={{background:B.red,padding:"0 16px",display:"flex",alignItems:"center",height:52,flexShrink:0,boxShadow:"0 2px 12px rgba(120,0,0,.5)"}}>
      <div style={{display:"flex",alignItems:"center",gap:10,paddingRight:14,borderRight:"1px solid rgba(255,255,255,.25)",marginRight:14,flexShrink:0}}><Logo/><div style={{lineHeight:1.2}}><div style={{fontWeight:900,fontSize:13,color:B.white,letterSpacing:".06em"}}>CUMMINS PERÚ</div><div style={{fontSize:8,color:"rgba(255,255,255,.6)",letterSpacing:".2em",marginTop:1}}>ZONA SUR · COMERCIAL</div></div></div>
      <nav style={{display:"flex",flex:1,overflowX:"auto",gap:0,msOverflowStyle:"none",scrollbarWidth:"none"}}>{VIEWS.map(v=>(<button key={v} onClick={()=>setView(v)} style={{background:view===v?"rgba(0,0,0,.28)":"transparent",border:"none",borderBottom:view===v?"2px solid #fff":"2px solid transparent",color:view===v?B.white:"rgba(255,255,255,.58)",fontFamily:"inherit",fontSize:10,fontWeight:view===v?800:500,letterSpacing:".06em",padding:"14px 11px 12px",cursor:"pointer",transition:"all .12s",whiteSpace:"nowrap",flexShrink:0}}>{v}</button>))}</nav>
      <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",paddingLeft:12,borderLeft:"1px solid rgba(255,255,255,.25)",marginLeft:10,flexShrink:0}}><span style={{fontWeight:900,color:B.white,fontSize:13,lineHeight:1}}>{f$(kpi.ing)}</span><span style={{color:kpi.mgnP>20?"#86EFAC":kpi.mgnP>10?"#FDE68A":"#FCA5A5",fontWeight:700,fontSize:9,marginTop:2}}>{fP(kpi.mgnP)} margen</span></div>
      <button onClick={()=>document.getElementById("fi").click()} style={{background:"rgba(0,0,0,.3)",border:"1px solid rgba(255,255,255,.38)",color:B.white,borderRadius:5,padding:"6px 12px",fontSize:10,cursor:"pointer",fontFamily:"inherit",fontWeight:800,letterSpacing:".06em",flexShrink:0,marginLeft:10,whiteSpace:"nowrap"}}>{loading?"⚙ Cargando…":"＋ EXCELS"}</button>
      <input id="fi" type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onChange={e=>ingest(Array.from(e.target.files))}/>
    </div>

    {/* FILES BAR */}
    {fileAggs.length>0&&(<div style={{background:B.surf,borderBottom:`1px solid ${B.border}`,padding:"4px 16px",display:"flex",gap:6,alignItems:"center",flexShrink:0,overflowX:"auto",height:28,msOverflowStyle:"none",scrollbarWidth:"none"}}>
      <span style={{fontSize:9,color:B.g3,letterSpacing:".12em",fontWeight:700,flexShrink:0}}>DATOS:</span>
      {fileAggs.filter(f=>fSuc.length===0||fSuc.includes(f.sucursal)).map((f,i)=>(<span key={i} style={{fontSize:9,background:B.card,border:`1px solid ${B.border}`,borderRadius:3,padding:"1px 7px",color:B.g3,borderLeft:`2px solid ${B.red}`,whiteSpace:"nowrap",flexShrink:0}}>{f.label} · {f$(f.totalIng)}</span>))}
      <span style={{fontSize:9,color:B.g3,marginLeft:4,flexShrink:0}}>Total: {f$(kpi.ing)}</span>
    </div>)}

    {/* MAIN */}
    {!hasData?(
      <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",flex:1,gap:24,padding:32}}>
        <div style={{textAlign:"center"}}><div style={{fontWeight:900,fontSize:30,color:B.red,lineHeight:1}}>INTELIGENCIA COMERCIAL</div><div style={{color:B.g3,fontSize:13,marginTop:8,letterSpacing:".15em"}}>CUMMINS PERÚ · ZONA SUR</div></div>
        <div onDrop={e=>{e.preventDefault();ingest(Array.from(e.dataTransfer.files));}} onDragOver={e=>e.preventDefault()} onClick={()=>document.getElementById("fi").click()} style={{border:`2px dashed ${dragging?B.red:B.border}`,borderRadius:16,padding:"44px 60px",textAlign:"center",cursor:"pointer",transition:"all .2s",maxWidth:460}}>
          <div style={{fontSize:42,marginBottom:12}}>📂</div><div style={{fontSize:15,fontWeight:700,color:B.white,marginBottom:8}}>Arrastra tus Excels aquí</div>
          <div style={{color:B.g3,fontSize:12,lineHeight:1.7}}>Uno o <strong style={{color:B.red}}>cientos de archivos</strong> a la vez<br/>Pre-agrega al vuelo — sin lag con 100+ Excels<br/><code style={{color:B.red,fontSize:11}}>2026_04_K27.xlsx · 2025_11_K23.xlsx</code></div>
        </div>
      </div>
    ):(
      <div style={{display:"flex",flex:1,overflow:"hidden",minHeight:0}}>
        <div style={{width:204,flexShrink:0,background:B.surf,borderRight:`1px solid ${B.border}`,padding:14,overflowY:"auto",minHeight:0}}><Sidebar/></div>
        <div className="fade" key={view} style={{flex:1,padding:"16px 20px",overflowY:"auto",minHeight:0,minWidth:0}}>{renderView()}</div>
      </div>
    )}

    {dragging&&hasData&&(<div onDrop={e=>{e.preventDefault();setDragging(false);ingest(Array.from(e.dataTransfer.files));}} onDragOver={e=>e.preventDefault()} onDragLeave={()=>setDragging(false)} style={{position:"fixed",inset:0,background:"rgba(204,0,0,.1)",border:"3px dashed #CC0000",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",backdropFilter:"blur(4px)"}}><div style={{textAlign:"center",color:B.white}}><div style={{fontSize:48,marginBottom:10}}>📂</div><div style={{fontSize:18,fontWeight:900}}>Suelta aquí</div></div></div>)}
    <div style={{position:"fixed",inset:0,pointerEvents:"none",zIndex:998}} onDragEnter={()=>setDragging(true)}/>
  </div>);
}
