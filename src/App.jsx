/workspace/Cummins_Dashboard
1549 src/App.jsx
69:const CA={
96:function detectHeaderRow(rawRows){
110:async function parseXLSX(file){
257:export default function App(){
307:  const ingest=useCallback(async fs=>{
1353:      {/* Column diagnostics */}
1	import { useState, useMemo, useCallback, useEffect, useRef } from "react";
     2	import * as XLSX from "xlsx";
     3	import { AreaChart, Area, BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, ScatterChart, Scatter, ZAxis, ComposedChart, RadarChart, Radar, PolarGrid, PolarAngleAxis } from "recharts";
     4	
     5	// ╔══════════════════════════════════════════════════════╗
     6	// ║           CUMMINS BRAND SYSTEM                       ║
     7	// ╚══════════════════════════════════════════════════════╝
     8	const B = {
     9	  red:"#CC0000",redDk:"#990000",redBg:"#180000",redDim:"#280000",
    10	  white:"#FFFFFF",gray1:"#E0E0E0",gray2:"#A0A0A0",gray3:"#666666",
    11	  gray4:"#3A3A3A",gray5:"#242424",bg:"#0C0C0C",surface:"#121212",
    12	  panel:"#181818",card:"#1C1C1C",border:"#2A2A2A",borderHi:"#404040",
    13	  green:"#16A34A",greenDim:"#0A1A0A",greenLt:"#4ADE80",
    14	  amber:"#D97706",amberDim:"#1C1000",amberLt:"#FCD34D",
    15	  blue:"#2563EB",blueDim:"#071020",blueLt:"#93C5FD",
    16	  teal:"#0D9488",tealDim:"#041410",
    17	  chart:["#CC0000","#E57373","#D97706","#16A34A","#2563EB","#7C3AED","#DB2777","#0D9488","#CA8A04","#EA580C"],
    18	};
    19	
    20	const CumminsLogo = ({h=38})=>(
    21	  <svg width={h*1.2} height={h} viewBox="0 0 96 80">
    22	    <polygon points="48,2 90,22 90,58 48,78 6,58 6,22" fill={B.red}/>
    23	    <text x="48" y="44" textAnchor="middle" fill="#fff" fontFamily="Arial Black,sans-serif" fontSize="16" fontWeight="900" letterSpacing=".5">CUMMINS</text>
    24	    <text x="48" y="58" textAnchor="middle" fill="rgba(255,255,255,.65)" fontFamily="Arial,sans-serif" fontSize="7" letterSpacing="2">ZONA SUR</text>
    25	  </svg>
    26	);
    27	
    28	// ╔══════════════════════════════════════════════════════╗
    29	// ║  FIREBASE CONFIG  ← REEMPLAZA CON TUS CREDENCIALES  ║
    30	// ╚══════════════════════════════════════════════════════╝
    31	const FB_CONFIG = {
    32	  apiKey:            "TU_API_KEY",
    33	  authDomain:        "TU_PROJECT.firebaseapp.com",
    34	  databaseURL:       "https://TU_PROJECT-default-rtdb.firebaseio.com",
    35	  projectId:         "TU_PROJECT",
    67	// ║           COLUMN DETECTION                           ║
    68	// ╚══════════════════════════════════════════════════════╝
    69	const CA={
    70	  ing: ["ingreso tot","ingresos","importe venta","venta neta","valor neto","net sales","revenue"],
    71	  cst: ["costo tot","costos material","costo total","coste total","cost","costos"],
    72	  cli: ["cliente","cod cliente","codigo cliente","código cliente","customer","sold to"],
    73	  clinm:["descripción cliente","descripcion cliente","nombre cliente","razon social","razón social","customer name"],
    74	  f2:  ["jerarquia de producto 2","jerarquía de producto 2","subfamilia","sub familia","product hierarchy 2"],
    75	  f1:  ["jerarquia de producto","jerarquía de producto","familia","product hierarchy"],
    76	  art: ["articulo","artículo","material","codigo material","código material","item"],
    77	  suc: ["sucursal","centro","branch"],
    78	  sec: ["descrip. gr.clie.","descrip.gr.clie.","grupo cliente","sector","industria"],
    79	  rep: ["representante de ventas","vendedor","ejecutivo","sales representative"],
    80	  vol: ["vol.ventas","vol ventas","volumen ventas","cantidad","qty","quantity"],
    81	  abc: ["indicador abc","abc"],
    82	  fecha:["fecha factura","fecha de factura","billing date","invoice date"],
    83	  per: ["período/año","periodo/año","periodo ano","periodo año","period"],
    84	};
    85	const normHeader=(v)=>String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
    86	// Exact match override — always wins over dc() fuzzy search
    87	function dcExact(headers,patterns){
    88	  const normalized=headers.map(h=>normHeader(h));
    89	  for(const p of patterns){
    90	    const ix=normalized.findIndex(h=>h===normHeader(p));
    91	    if(ix!==-1)return headers[ix];
    92	  }
    93	  return null;
    94	}
    95	function dc(headers,f){const lo=headers.map(h=>normHeader(h));const al=(CA[f]||[]).map(normHeader);const i=lo.findIndex(h=>al.some(a=>h===a||h.includes(a)));return i!==-1?headers[i]:null;}
    96	function detectHeaderRow(rawRows){
    97	  const max=Math.min(rawRows.length,25);
    98	  let best={idx:0,score:-1};
    99	  for(let i=0;i<max;i++){
   100	    const headers=(rawRows[i]||[]).map(v=>String(v??"").trim());
   101	    if(!headers.some(Boolean))continue;
   102	    const score=Object.keys(CA).reduce((sum,f)=>sum+(dc(headers,f)?1:0),0);
   103	    if(score>best.score)best={idx:i,score};
   104	  }
   105	  return best;
   106	}
   107	function sucFromName(n){const u=n.toUpperCase();if(u.includes("K27")||u.includes("AREQUIPA"))return"K27";if(u.includes("K11")||u.includes("HUANCAYO"))return"K11";if(u.includes("K23")||u.includes("TACNA"))return"K23";if(u.includes("K01")||u.includes("LIMA")||u.includes("CALLAO"))return"K01";return"K??"}
   108	function periodFromName(n){const u=n.toUpperCase().replace(/[_\-\.]/g," ");const m=u.match(/(\d{4})\s+(\d{1,2})/);if(m)return{y:parseInt(m[1]),mo:parseInt(m[2])};const mi=SM.findIndex(s=>u.includes(s));if(mi!==-1){const ym=u.match(/(\d{4})/);if(ym)return{y:parseInt(ym[1]),mo:mi+1}}return null}
   109	
   110	async function parseXLSX(file){
   111	  const ab=await file.arrayBuffer();
   112	  const wb=XLSX.read(ab,{type:"array",cellDates:true});
   113	  const ws=wb.Sheets[wb.SheetNames[0]];
   114	  const label=file.name.replace(/\.[^.]+$/,"");
   115	  const sfn=sucFromName(label);
   116	  const pfn=periodFromName(label);
   117	
   118	  // Some SAP/Excel exports include title/filter rows before the real header.
   119	  // Detect the row with the most known business columns, then parse from there.
   120	  const rawRows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,blankrows:false});
   121	  if(!rawRows.length)return null;
   122	  const headerInfo=detectHeaderRow(rawRows);
   123	  const rows=XLSX.utils.sheet_to_json(ws,{defval:null,range:headerInfo.idx});
   124	  if(!rows.length)return null;
   125	  const headers=Object.keys(rows[0]);
   250	};
   251	
   252	const VIEWS=["RESUMEN","CLIENTES","PRODUCTOS","SECTORES","TENDENCIAS","EJECUTIVOS","TABLA DINÁMICA","ANÁLISIS","FIREBASE"];
   253	
   254	// ╔══════════════════════════════════════════════════════╗
   255	// ║           MAIN APP                                   ║
   256	// ╚══════════════════════════════════════════════════════╝
   257	export default function App(){
   258	  const [allRecs,  setAllRecs]  = useState([]);
   259	  const [files,    setFiles]    = useState([]);
   260	  const [view,     setView]     = useState("RESUMEN");
   261	  const [loading,  setLoading]  = useState(false);
   262	  const [dragging, setDragging] = useState(false);
   263	  const [fbStatus, setFbStatus] = useState("idle"); // idle|uploading|ok|error
   264	
   265	  // — Filters (local per-user, not shared) —
   266	  const [fSuc, setFSuc] = useState([]);
   267	  const [fFY,  setFY]   = useState([]);
   268	  const [fQ,   setFQ]   = useState([]);
   269	  const [fSec, setFSec] = useState([]);
   270	  const [fFam, setFFam] = useState([]);
   271	  const [fRep, setFRep] = useState([]);
   272	  const [fSeg, setFSeg] = useState([]);
   273	  const [search,setSearch]=useState("");
   274	
   275	  // — Cartera & rep names (persisted locally) —
   276	  const [cartera,  setCartera]  = useState(()=>{try{return JSON.parse(localStorage.getItem("cummins_c")||"{}")}catch{return{}}});
   277	  const [repNames, setRepNames] = useState(()=>{try{return JSON.parse(localStorage.getItem("cummins_r")||"{}")}catch{return{}}});
   278	  const [selCli,   setSelCli]   = useState(null);
   279	
   280	  // — Pivot config —
   281	  const [pivotRows, setPivotRows]   = useState("sec");
   282	  const [pivotCols, setPivotCols]   = useState("fyl");
   283	  const [pivotVal,  setPivotVal]    = useState("ing");
   284	  const [pivotSort, setPivotSort]   = useState("desc");
   285	
   286	  useEffect(()=>{localStorage.setItem("cummins_c",JSON.stringify(cartera));},[cartera]);
   287	  useEffect(()=>{localStorage.setItem("cummins_r",JSON.stringify(repNames));},[repNames]);
   288	
   289	  // Seed records
   290	  const seedRecs = useMemo(()=>{
   291	    const recs=[];
   292	    SEED.clients.forEach(cl=>{
   293	      const chunk=Math.max(1,Math.min(cl.lineas,5));
   294	      for(let i=0;i<chunk;i++){
   295	        recs.push({ing:cl.ing/chunk,cst:cl.cst/chunk,mgn:cl.mgn/chunk,vol:Math.round(cl.vol/chunk),
   296	          cli:String(cl.id),clinm:cl.nm,f2:"—",f1:"—",art:"—",suc:"K27",sec:cl.sector,rep:"0",
   297	          abc:"N/A",mo:4,yr:2026,fy:2026,fmi:0,fq:1,fml:"Abr",fyl:"FY 2026/27",file:"2026_04_K27"});
   298	      }
   299	    });
   300	    return recs;
   301	  },[]);
   302	
   303	  useEffect(()=>{if(allRecs.length===0)setAllRecs(seedRecs);},[]);
   304	  useEffect(()=>{if(files.length===0&&allRecs.length>0)setFiles([{label:"2026_04_K27",count:SEED.meta.lineas,suc:"K27"}]);},[allRecs]);
   305	
   306	  // — File ingestion — clears seedRecs on first real upload —
   307	  const ingest=useCallback(async fs=>{
   308	    setLoading(true);
   309	    const newRecs=[],newFiles=[];
   310	    try{
   311	      for(const f of fs){
   312	        if(!f.name.match(/\.xlsx?$/i))continue;
   313	        const p=await parseXLSX(f);
   314	        if(!p)continue;
   315	        newRecs.push(...p.recs);
   316	        newFiles.push({label:p.label,count:p.recs.length,suc:p.suc,cols:p.cols,headers:p.headers,headerRow:p.headerRow});
   317	      }
   318	    }catch(e){
   319	      console.error("Error leyendo Excel",e);
   320	      alert(`No se pudo leer el Excel: ${e?.message||e}`);
   321	      setLoading(false);
   322	      return;
   323	    }
   324	    if(!newFiles.length){setLoading(false);alert("No se encontraron archivos Excel válidos (.xlsx o .xls).");return;}
   325	    setAllRecs(prev=>{
  1348	          </button>
  1349	          <input id="fb-upload" type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onChange={e=>{const fs=Array.from(e.target.files);uploadToFirebase(fs);ingest(fs);}}/>
  1350	        </div>
  1351	      </Pnl>
  1352	
  1353	      {/* Column diagnostics */}
  1354	      {files.filter(f=>f.cols).length>0&&(
  1355	        <Pnl>
  1356	          <ST sub="Columnas detectadas en cada archivo — si 'Ingresos' dice null, el nombre de columna no coincide">🔍 Diagnóstico de columnas detectadas</ST>
  1357	          <div style={{display:"flex",flexDirection:"column",gap:12}}>
  1358	            {files.filter(f=>f.cols).slice(0,5).map((f,i)=>(
  1359	              <div key={i} style={{background:B.panel,borderRadius:8,padding:"12px 14px",border:`1px solid ${B.border}`}}>
  1360	                <div style={{fontWeight:700,color:B.white,fontSize:11,marginBottom:8}}>{f.label}</div>
  1361	                <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:8}}>
  1362	                  {Object.entries(f.cols||{}).map(([field,col])=>(
  1363	                    <span key={field} style={{fontSize:10,padding:"2px 8px",borderRadius:4,
  1364	                      background:col?B.greenDim:B.redDim,
  1365	                      border:`1px solid ${col?B.green:B.red}`,
  1366	                      color:col?B.greenLt:B.red}}>
  1367	                      {field}: {col?`"${col}"`:"❌ NO DETECTADO"}
  1368	                    </span>
  1369	                  ))}
  1370	                </div>
  1371	                <div style={{fontSize:9,color:B.gray3}}>
  1372	                  Fila de encabezados detectada: {f.headerRow||1} · Primeras columnas del archivo: {(f.headers||[]).join(" · ")}
  1373	                </div>
  1374	              </div>
  1375	            ))}
  1538	          style={{position:"fixed",inset:0,background:"rgba(204,0,0,.12)",border:"3px dashed #CC0000",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",backdropFilter:"blur(4px)"}}>
  1539	          <div style={{textAlign:"center",color:B.white}}>
  1540	            <div style={{fontSize:52,marginBottom:12}}>📂</div>
  1541	            <div style={{fontSize:18,fontWeight:900,letterSpacing:".04em"}}>Suelta los archivos Excel aquí</div>
  1542	            <div style={{fontSize:12,color:"rgba(255,255,255,.6)",marginTop:6}}>Se cargarán automáticamente</div>
  1543	          </div>
  1544	        </div>
  1545	      )}
  1546	      <div style={{position:"fixed",inset:0,pointerEvents:"none",zIndex:998}} onDragEnter={()=>setDragging(true)}/>
  1547	    </div>
  1548	  );
  1549	}
