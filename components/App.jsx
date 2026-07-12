import { useState, useEffect, useMemo, useRef, useCallback, Fragment } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Package, Boxes, Upload, FileText,
  Plus, Trash2, Pencil, Check, X, ChevronDown, ChevronRight,
  Menu, Download, AlertTriangle, Search, Save, Image as ImageIcon,
  Zap, BarChart3, Calculator, RefreshCw, TrendingUp, ShieldCheck,
  Tag, IndianRupee, Lock, Unlock, Database, Users, Calendar,
} from "lucide-react";
import { loadData, saveData, subscribeToData, hasSync } from "../lib/db";

/* ═══ BRAND TOKENS ═══ */
const C = {
  zenkyPurple:"#5B2DDA", zenkyPurpleDark:"#4A1FB5", zenkyPink:"#FF4F9A",
  zenkyOrange:"#FF8A1F", sunshineYellow:"#FFC83D", softWhite:"#FFFDF8",
  skyBlue:"#71D7FF", mintGreen:"#9BE76A", darkText:"#2D1B4E",
  lightText:"#7B6B9E", border:"#E8DFF5", bgLight:"#F9F6FF",
};
const F = { display:"'Fredoka','Baloo 2',sans-serif", body:"'Nunito','Poppins',sans-serif", mono:"'Nunito Mono',monospace" };

const MAX_IMAGES = 9;
const MARGIN_OPTIONS = [20, 25, 30, 35, 40, 45];
const MRP_DISCOUNTS  = [50, 55, 60, 65, 70, 75, 80];

const DEFAULT_ADMIN_PIN = "2468"; // used only if no PIN has ever been set in Access Management
const DEFAULT_LOGIN = { username: "admin", password: "zenkybox123" }; // used only until changed in Access Management

// Your requested heads, plus a few standard ones common to small D2C/Amazon
// businesses that are easy to miss (flagged below so you can drop any you
// don't need) — Returns & Refunds, Salary & Wages, Payment Gateway/Bank
// Charges, and Software & Hosting (beyond just AI/Meta subscriptions).
const EXPENSE_HEADS = [
  "Product Procurement","Ad Expenses","Packaging Expenses","Branding Expenses",
  "Travel Expenses","Food Expenses","Taxes","Registration Fee","Trademark Registration Fee",
  "AI Subscription","Meta Subscription",
  "Salary & Wages","Payment Gateway / Bank Charges","Returns & Refunds","Software & Hosting",
  "Other Expenses",
];
const INCOME_HEADS = ["Income from Amazon","Income from Website","Other Income"];
const PAYMENT_MODES = ["Bank Transfer","UPI","Cash","Card"];

const NAV = [
  { id:"dashboard",       label:"Dashboard",         icon:LayoutDashboard, adminOnly:false },
  { id:"combo-readiness", label:"Combo Readiness",   icon:ShieldCheck,     adminOnly:false },
  { id:"bulk-import",     label:"Bulk Import",       icon:Upload,          adminOnly:true  },
  { id:"catalog",         label:"SKU Catalog",       icon:Package,         adminOnly:false },
  { id:"combos",          label:"Gift Combos",       icon:Boxes,           adminOnly:false },
  { id:"upload",          label:"Upload Sales",      icon:TrendingUp,      adminOnly:false },
  { id:"reports",         label:"Reports",           icon:FileText,        adminOnly:false },
  { id:"sales-reports",   label:"ZenkyBox Sales Report", icon:BarChart3,   adminOnly:false },
  { id:"costing",         label:"Costing & Pricing", icon:Calculator,      adminOnly:true  },
  { id:"financials",      label:"Financials",        icon:IndianRupee,     adminOnly:true  },
  { id:"source-data",     label:"Source Data",       icon:Database,        adminOnly:true  },
  { id:"access",          label:"Access Management", icon:Users,           adminOnly:true  },
];

/* ═══ HELPERS ═══ */
function stockStatus(s){ if(s.stock<=0)return"critical"; if(s.stock<=s.reorderLevel)return"low"; return"healthy"; }
function suggestedReorder(s){ if(!s.reorderLevel||s.stock>s.reorderLevel)return 0; return Math.max(s.reorderLevel*2-s.stock,s.reorderLevel); }
function comboReadiness(combo,skuMap){
  if(!combo.components?.length)return{ready:0,bottleneck:null};
  let min=Infinity,bot=null;
  combo.components.forEach(c=>{const st=skuMap[c.sku]?.stock||0;const r=Math.floor(st/(c.qty||1));if(r<min){min=r;bot=c.sku;}});
  return{ready:min===Infinity?0:min,bottleneck:bot};
}
function fmt(n){return Number(n||0).toLocaleString("en-IN");}
function fmtINR(n){return`₹${Number(n||0).toFixed(2)}`;}
function downloadCsv(filename,csv){
  const b=new Blob([csv],{type:"text/csv;charset=utf-8;"});
  const u=URL.createObjectURL(b);const a=document.createElement("a");
  a.href=u;a.download=filename;document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(u);
}

/* Auto-generate SKU: ZB-YYYYMM-NNNNN */
/* Forgiving parser for the "Components" bulk-import column.
   Accepts: "SKU:2;SKU2:1" (correct format), but also tolerates
   comma-separated lists, missing qty, and mixed separators —
   so a data-entry mistake doesn't silently break combo readiness. */
function parseComponentsString(raw){
  if(!raw)return[];
  // Normalize: treat both ; and , as separators between components
  const tokens=String(raw).split(/[;,]/).map(t=>t.trim()).filter(Boolean);
  const merged={};
  tokens.forEach(tok=>{
    const parts=tok.split(":");
    const sku=(parts[0]||"").trim();
    const qty=parts.length>1?(Number(parts[1])||1):1;
    if(!sku)return;
    merged[sku]=(merged[sku]||0)+qty;
  });
  return Object.entries(merged).map(([sku,qty])=>({sku,qty}));
}

function generateSkuCode(existingSkus){
  const now=new Date();
  const prefix=`ZB-${now.getFullYear()}${String(now.getMonth()+1).padStart(2,"0")}`;
  const nums=existingSkus.filter(s=>s.sku.startsWith(prefix)).map(s=>parseInt(s.sku.split("-")[2]||"0")).filter(n=>!isNaN(n));
  const next=(nums.length?Math.max(...nums):0)+1;
  return`${prefix}-${String(next).padStart(5,"0")}`;
}

/* ═══ SALES ANALYTICS HELPERS ═══ */
const MONTH_NAMES=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function monthKey(d){const dt=new Date(d);return isNaN(dt)?"Unknown":`${MONTH_NAMES[dt.getMonth()]} ${dt.getFullYear()}`;}
/* Indian Financial Year: 1 April – 31 March. e.g. a date in Feb 2027 falls in "FY 2026-27". */
function fyLabel(d){
  const dt=new Date(d);if(isNaN(dt))return"Unknown";
  const y=dt.getFullYear(),m=dt.getMonth(); // month 3 = April (0-indexed)
  const startYear=m>=3?y:y-1;
  return`FY ${startYear}-${String(startYear+1).slice(-2)}`;
}
function weekKey(d){
  const dt=new Date(d);if(isNaN(dt))return"Unknown";
  const start=new Date(dt);start.setDate(dt.getDate()-dt.getDay()); // Sunday start
  const end=new Date(start);end.setDate(start.getDate()+6);
  const fmt2=x=>`${x.getDate()} ${MONTH_NAMES[x.getMonth()]}`;
  return`${fmt2(start)} – ${fmt2(end)} ${end.getFullYear()}`;
}

/* Cost of one unit sold, given a matched code (direct SKU or combo code) */
function unitCostOf(code,matchType,skuMap,comboMap){
  if(matchType==="direct"){const s=skuMap[code];return s?Number(s.procurementCost||0):0;}
  if(matchType==="combo"){
    const c=comboMap[code];if(!c)return 0;
    return c.components?.reduce((sum,comp)=>sum+(Number(skuMap[comp.sku]?.procurementCost||0)*(comp.qty||1)),0)||0;
  }
  return 0;
}

/* Detect optional extra columns in an uploaded sales file (best-effort, tolerant of Amazon's varying header names) */
function detectExtraColumns(headerKeys){
  const find=re=>headerKeys.find(k=>re.test(k));
  return{
    dateKey:find(/purchase.?date|order.?date/i),
    buyerKey:find(/buyer.?name|buyer.?email|customer.?name/i),
    orderIdKey:find(/^amazon.?order.?id$|^order.?id$/i),
    cityKey:find(/ship.?city/i),
    stateKey:find(/ship.?state|ship.?region/i),
    priceKey:find(/item.?price|unit.?price|^price$/i),
  };
}

/* pad a numeric-looking value back to 2 digits — Excel silently drops leading
   zeros from split-off timestamp fragments (e.g. "01" becomes 1) */
function pad2(v){const s=String(v??"").trim();return s.length===1?"0"+s:s;}

/* Some Amazon exports (via certain conversion tools) split ISO timestamp columns
   like "purchase-date" and "last-updated-date" into 4 extra cells each, because the
   colons in "2026-07-05T06:29:52+00:00" get treated as a delimiter somewhere in the
   pipeline. This pushes every column after those date fields to the wrong position —
   including sku and quantity — with the header row staying correctly labeled.
   This function detects that exact signature and reconstructs the original columns
   automatically, so the file can be uploaded exactly as exported, with no manual fixing. */
function repairShiftedColumns(headerRow,dataRows){
  const dateIdxs=headerRow.map((h,i)=>/date/i.test(h)?i:-1).filter(i=>i>=0);
  if(!dateIdxs.length)return{repaired:false,rows:dataRows};

  const excesses=dataRows.map(r=>r.length-headerRow.length).filter(e=>e>0);
  if(!excesses.length)return{repaired:false,rows:dataRows}; // already aligned — nothing to do

  const expected=dateIdxs.length*3; // each split date column contributes 3 extra cells (4 parts - 1)
  const consistent=excesses.every(e=>e===expected);
  if(!consistent)return{repaired:false,rows:dataRows,reason:"inconsistent"};

  const repairedRows=dataRows.map(raw=>{
    if(raw.length<=headerRow.length)return raw; // this particular row wasn't affected
    const out=[];let pos=0;
    headerRow.forEach((h,hi)=>{
      if(dateIdxs.includes(hi)){
        const parts=raw.slice(pos,pos+4);pos+=4;
        out.push(`${parts[0]}:${pad2(parts[1])}:${parts[2]}:${pad2(parts[3])}`);
      }else{
        out.push(raw[pos]);pos+=1;
      }
    });
    return out;
  });
  return{repaired:true,rows:repairedRows};
}

function aggregateSalesLines(lines,groupFn,skuMap,comboMap,sortMode="revenue"){
  const groups={};
  lines.forEach(l=>{
    const key=groupFn(l);
    if(!groups[key])groups[key]={key,qty:0,revenue:0,cost:0,earning:0,lines:[],minDate:l.date};
    const g=groups[key];
    g.qty+=l.qty;g.revenue+=l.revenue;g.cost+=l.cost;g.earning+=l.earning;g.lines.push(l);
    if(new Date(l.date)<new Date(g.minDate))g.minDate=l.date;
  });
  const arr=Object.values(groups);
  return sortMode==="chrono"?arr.sort((a,b)=>new Date(a.minDate)-new Date(b.minDate)):arr.sort((a,b)=>b.revenue-a.revenue);
}

/* Tiny dependency-free SVG bar chart — used for MoM / Weekly revenue trends */
function TrendChart({groups}){
  if(groups.length<2)return null; // need at least 2 points for a trend to mean anything
  const W=Math.max(320,groups.length*90),H=180,padL=48,padB=28,padT=10;
  const max=Math.max(1,...groups.map(g=>Math.max(g.revenue,g.cost)));
  const barW=Math.min(28,(W-padL-20)/groups.length/2.4);
  const chartW=W-padL-20;
  const step=chartW/groups.length;
  const scaleY=v=>H-padB-(v/max)*(H-padB-padT);
  return(
    <div className="overflow-x-auto mb-4">
      <svg width={W} height={H} style={{minWidth:"100%"}}>
        {[0,0.5,1].map(f=>(
          <g key={f}>
            <line x1={padL} x2={W-10} y1={scaleY(max*f)} y2={scaleY(max*f)} stroke={C.border} strokeWidth="1"/>
            <text x={padL-6} y={scaleY(max*f)+3} fontSize="9" fill={C.lightText} textAnchor="end" fontFamily={F.mono}>{f===0?"0":fmt(Math.round(max*f))}</text>
          </g>
        ))}
        {groups.map((g,i)=>{
          const cx=padL+step*i+step/2;
          const earnColor=g.earning>=0?C.mintGreen:C.zenkyPink;
          return(
            <g key={g.key}>
              <rect x={cx-barW-2} y={scaleY(g.revenue)} width={barW} height={Math.max(0,H-padB-scaleY(g.revenue))} fill={C.zenkyPurple} rx="2"/>
              <rect x={cx+2} y={scaleY(g.cost)} width={barW} height={Math.max(0,H-padB-scaleY(g.cost))} fill={C.zenkyOrange} rx="2"/>
              <circle cx={cx} cy={scaleY(g.earning)} r="3" fill={earnColor}/>
              <text x={cx} y={H-padB+14} fontSize="9" fill={C.lightText} textAnchor="middle" fontFamily={F.body}>{g.key.length>10?g.key.slice(0,9)+"…":g.key}</text>
            </g>
          );
        })}
      </svg>
      <div className="flex items-center gap-4 mt-1 text-xs" style={{color:C.lightText,fontFamily:F.body}}>
        <span className="inline-flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm inline-block" style={{backgroundColor:C.zenkyPurple}}/>Revenue</span>
        <span className="inline-flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm inline-block" style={{backgroundColor:C.zenkyOrange}}/>Cost</span>
        <span className="inline-flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full inline-block" style={{backgroundColor:C.mintGreen}}/>Earning</span>
      </div>
    </div>
  );
}

/* ═══ SHARED UI ═══ */
/* Reusable sortable-column table header + hook. Click a header to sort by it;
   click again to flip direction. Used across the Reports section. */
function useSortableRows(rows,defaultKey,defaultDir="desc"){
  const [sortKey,setSortKey]=useState(defaultKey);
  const [sortDir,setSortDir]=useState(defaultDir);
  function toggleSort(key){
    if(sortKey===key)setSortDir(sortDir==="asc"?"desc":"asc");
    else{setSortKey(key);setSortDir("desc");}
  }
  const sorted=useMemo(()=>{
    const arr=[...rows];
    arr.sort((a,b)=>{
      let av=a[sortKey],bv=b[sortKey];
      if(typeof av==="string")av=av.toLowerCase();
      if(typeof bv==="string")bv=bv.toLowerCase();
      if(av<bv)return sortDir==="asc"?-1:1;
      if(av>bv)return sortDir==="asc"?1:-1;
      return 0;
    });
    return arr;
  },[rows,sortKey,sortDir]);
  return{sorted,sortKey,sortDir,toggleSort};
}
function SortTH({label,sortKey,activeKey,dir,onClick,className=""}){
  const active=activeKey===sortKey;
  return(
    <th className={`py-2 pr-3 text-left text-xs uppercase font-bold cursor-pointer select-none ${className}`} style={{color:active?C.zenkyPurple:C.lightText}} onClick={()=>onClick(sortKey)}>
      <span className="inline-flex items-center gap-1">{label}{active&&(dir==="asc"?<ChevronDown size={12} style={{transform:"rotate(180deg)"}}/>:<ChevronDown size={12}/>)}</span>
    </th>
  );
}
/* Sortable Code/Name/Qty/Revenue/COGS/Gross-Profit table — used by the P&L
   Statement for both the SKU and Combo revenue breakdowns. */
function PLSortableTable({rows,title}){
  const{sorted,sortKey,sortDir,toggleSort}=useSortableRows(rows,"revenue","desc");
  return(
    <div className="mb-4">
      <div className="text-xs font-bold uppercase mb-2" style={{color:C.lightText}}>{title}</div>
      <div className="overflow-x-auto"><table className="w-full text-sm">
        <thead><tr>
          <SortTH label="Code" sortKey="code" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
          <SortTH label="Name" sortKey="name" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
          <SortTH label="Qty" sortKey="qty" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
          <SortTH label="Revenue" sortKey="revenue" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
          <SortTH label="COGS" sortKey="cogs" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
          <SortTH label="Gross Profit" sortKey="gp" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
        </tr></thead>
        <tbody>{sorted.map(r=>(
          <tr key={r.code} className="border-t" style={{borderColor:C.border}}>
            <td className="py-1.5 pr-3" style={{fontFamily:F.mono,fontWeight:600}}>{r.code}</td>
            <td className="py-1.5 pr-3">{r.name}</td>
            <td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmt(r.qty)}</td>
            <td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmtINR(r.revenue)}</td>
            <td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmtINR(r.cogs)}</td>
            <td className="py-1.5 pr-3 font-bold" style={{fontFamily:F.mono,color:C.mintGreen}}>{fmtINR(r.gp)}</td>
          </tr>
        ))}</tbody>
      </table></div>
    </div>
  );
}

function Stamp({tone="purple",children}){
  const m={purple:{color:C.zenkyPurple,bg:C.bgLight},pink:{color:C.zenkyPink,bg:"#FFE6F2"},orange:{color:C.zenkyOrange,bg:"#FFF3E6"},mint:{color:C.mintGreen,bg:"#F0FDE8"},blue:{color:"#0ea5e9",bg:"#e0f2fe"}};
  const t=m[tone]||m.purple;
  return(<span style={{color:t.color,backgroundColor:t.bg,fontFamily:F.display,letterSpacing:"0.02em"}} className="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-bold uppercase whitespace-nowrap">★ {children}</span>);
}
function statusStamp(s){if(s==="critical")return<Stamp tone="pink">Critical</Stamp>;if(s==="low")return<Stamp tone="orange">Low Stock</Stamp>;return<Stamp tone="mint">Healthy</Stamp>;}

function StockGauge({stock,reorderLevel}){
  const max=Math.max(stock,reorderLevel*2,1);
  const fp=Math.min(100,(stock/max)*100);const tp=Math.min(100,(reorderLevel/max)*100);
  const s=stockStatus({stock,reorderLevel});
  const col=s==="critical"?C.zenkyPink:s==="low"?C.zenkyOrange:C.mintGreen;
  return(<div className="relative h-3 rounded-full overflow-hidden w-full" style={{backgroundColor:"#E8DFF5"}}><div className="h-full rounded-full transition-all" style={{width:`${fp}%`,backgroundColor:col}}/>{reorderLevel>0&&<div className="absolute top-0 bottom-0" style={{left:`${tp}%`,width:"2px",backgroundColor:C.darkText,opacity:0.5}}/>}</div>);
}

function Card({children,className=""}){return<div className={`rounded-2xl border-2 p-4 md:p-5 shadow-sm ${className}`} style={{backgroundColor:C.softWhite,borderColor:C.border}}>{children}</div>;}

function SectionHeader({title,subtitle,action}){
  return(<div className="flex items-start justify-between gap-4 mb-6 flex-wrap"><div><h2 className="text-2xl md:text-3xl font-black" style={{fontFamily:F.display,color:C.zenkyPurple}}>{title}</h2>{subtitle&&<p className="text-sm mt-1" style={{color:C.lightText,fontFamily:F.body}}>{subtitle}</p>}</div>{action}</div>);
}

function Empty({icon:Icon,title,message}){
  return(<div className="flex flex-col items-center justify-center text-center py-12 px-4 rounded-2xl border-2 border-dashed" style={{borderColor:C.zenkyPink,backgroundColor:"#FFF8FC"}}><Icon size={32} style={{color:C.zenkyPink}}/><p className="mt-3 font-bold text-lg" style={{color:C.darkText,fontFamily:F.display}}>{title}</p><p className="text-sm mt-1 max-w-sm" style={{color:C.lightText,fontFamily:F.body}}>{message}</p></div>);
}

function PrimaryButton({children,onClick,type="button",disabled,tone}){
  const bg=tone==="orange"?C.zenkyOrange:tone==="pink"?C.zenkyPink:C.zenkyPurple;
  return(<button type={type} onClick={onClick} disabled={disabled} className="inline-flex items-center gap-2 px-4 py-2.5 rounded-full text-sm font-bold disabled:opacity-50 transition-transform hover:scale-105 whitespace-nowrap" style={{backgroundColor:bg,color:C.softWhite,fontFamily:F.display}}>{children}</button>);
}

function GhostButton({children,onClick,title}){return(<button type="button" title={title} onClick={onClick} className="inline-flex items-center justify-center w-8 h-8 rounded-full border-2 hover:scale-110 transition-transform flex-shrink-0" style={{borderColor:C.border,color:C.zenkyPurple}}>{children}</button>);}

function Input(props){
  return(<input {...props} className={`w-full px-3 py-2.5 rounded-xl border-2 text-sm focus:outline-none transition-colors ${props.className||""}`} style={{borderColor:C.border,fontFamily:F.body,backgroundColor:C.softWhite,...props.style}} onFocus={e=>(e.target.style.borderColor=C.zenkyPurple)} onBlur={e=>(e.target.style.borderColor=C.border)}/>);
}

function Select({children,...props}){
  return(<select {...props} className={`w-full px-3 py-2.5 rounded-xl border-2 text-sm focus:outline-none bg-white ${props.className||""}`} style={{borderColor:C.border,fontFamily:F.body,...props.style}} onFocus={e=>(e.target.style.borderColor=C.zenkyPurple)} onBlur={e=>(e.target.style.borderColor=C.border)}>{children}</select>);
}

function CostInput({label,value,onChange,prefix="₹",readOnly}){
  return(
    <div>
      <label className="text-xs font-bold uppercase block mb-1" style={{color:C.lightText,fontFamily:F.body,letterSpacing:"0.02em"}}>{label}</label>
      <div className="relative">
        <span className="absolute left-3 top-2.5 text-sm font-bold" style={{color:C.lightText}}>{prefix}</span>
        <Input type="number" min="0" step="0.01" value={value} onChange={onChange} readOnly={readOnly} className="pl-7" style={readOnly?{backgroundColor:C.bgLight,color:C.zenkyPurple,fontWeight:"bold"}:{}}/>
      </div>
    </div>
  );
}

/* ═══ IMAGE GALLERY ═══ */
function ImageGallery({images=[],onAddImages,onRemoveImage}){
  const ref=useRef(null);
  const handle=e=>{const files=Array.from(e.target.files||[]).slice(0,MAX_IMAGES-images.length);files.forEach(f=>{const r=new FileReader();r.onload=ev=>onAddImages([ev.target.result]);r.readAsDataURL(f);});if(ref.current)ref.current.value="";};
  return(
    <div>
      <label className="text-xs font-bold uppercase block mb-2" style={{color:C.lightText,fontFamily:F.body,letterSpacing:"0.02em"}}>Product Images ({images.length}/{MAX_IMAGES})</label>
      <div className="grid grid-cols-3 sm:grid-cols-5 gap-2">
        {images.map((img,i)=>(
          <div key={i} className="relative w-full aspect-square rounded-lg overflow-hidden border-2 group" style={{borderColor:C.border}}>
            <img src={img} alt="" className="w-full h-full object-cover"/>
            <button onClick={()=>onRemoveImage(i)} className="absolute inset-0 bg-black bg-opacity-0 group-hover:bg-opacity-40 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-all"><Trash2 size={16} color="white"/></button>
          </div>
        ))}
        {images.length<MAX_IMAGES&&(
          <label className="w-full aspect-square rounded-lg border-2 border-dashed flex items-center justify-center cursor-pointer" style={{borderColor:C.zenkyPink,backgroundColor:"#FFF8FC"}}>
            <input ref={ref} type="file" multiple accept="image/*" onChange={handle} className="hidden"/>
            <div className="text-center"><ImageIcon size={16} style={{color:C.zenkyPink,margin:"0 auto"}}/><div className="text-xs font-bold mt-1" style={{color:C.zenkyPink}}>{images.length===0?"Add images":"Add more"}</div></div>
          </label>
        )}
      </div>
    </div>
  );
}

/* ═══ SIDEBAR ═══ */
/* ═══ LOGIN GATE ═══ */
function LoginScreen({onLogin}){
  const [username,setUsername]=useState("");
  const [password,setPassword]=useState("");
  const [error,setError]=useState("");

  function submit(){
    if(!onLogin(username.trim(),password)){
      setError("Incorrect username or password.");
    }
  }

  return(
    <div className="flex items-center justify-center h-screen p-6" style={{backgroundColor:C.bgLight,fontFamily:F.body}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fredoka:wght@400;500;600;700&family=Nunito:wght@400;500;600;700&display=swap');`}</style>
      <div className="w-full max-w-sm">
        <div className="text-center mb-6">
          <div style={{fontFamily:F.display,color:C.zenkyPurple,fontSize:"32px",fontWeight:"black"}}>✨ ZenkyBox</div>
          <div className="text-xs mt-1 font-bold uppercase" style={{color:C.lightText,letterSpacing:"0.05em"}}>📦 Inventory Hub</div>
        </div>
        <div className="rounded-2xl border-2 p-6" style={{backgroundColor:C.softWhite,borderColor:C.border}}>
          <h2 className="font-black text-lg mb-4 text-center" style={{fontFamily:F.display,color:C.darkText}}>Sign in to continue</h2>
          <div className="space-y-3">
            <div>
              <label className="text-xs font-bold uppercase block mb-1" style={{color:C.lightText}}>Username</label>
              <Input value={username} onChange={e=>{setUsername(e.target.value);setError("");}} onKeyDown={e=>e.key==="Enter"&&submit()} autoFocus/>
            </div>
            <div>
              <label className="text-xs font-bold uppercase block mb-1" style={{color:C.lightText}}>Password</label>
              <Input type="password" value={password} onChange={e=>{setPassword(e.target.value);setError("");}} onKeyDown={e=>e.key==="Enter"&&submit()}/>
            </div>
            {error&&<p className="text-xs font-bold" style={{color:"#dc2626"}}>{error}</p>}
            <PrimaryButton onClick={submit}><Lock size={15}/>Sign In</PrimaryButton>
          </div>
        </div>
        <p className="text-xs text-center mt-4" style={{color:C.lightText}}>Access is restricted. Contact your ZenkyBox admin for credentials.</p>
      </div>
    </div>
  );
}


function Sidebar({view,setView,open,setOpen,synced,role,canBeAdmin,currentUserName,onUnlock,onLock,onLogout}){
  const [showPinBox,setShowPinBox]=useState(false);
  const [pinInput,setPinInput]=useState("");
  const visibleNav=NAV.filter(item=>!item.adminOnly||role==="admin");

  function handleUnlock(){
    if(onUnlock(pinInput)){setShowPinBox(false);setPinInput("");}
  }

  return(
    <>
      {open&&<div className="fixed inset-0 z-20 md:hidden" style={{backgroundColor:"rgba(0,0,0,0.3)"}} onClick={()=>setOpen(false)}/>}
      <aside className={`fixed md:static z-30 top-0 left-0 h-full w-64 flex flex-col transition-transform duration-300 ${open?"translate-x-0":"-translate-x-full"} md:translate-x-0`} style={{backgroundColor:C.zenkyPurple}}>
        <div className="px-5 pt-5 pb-3 border-b" style={{borderColor:"rgba(255,255,255,0.15)"}}>
          <div className="text-xl font-black" style={{fontFamily:F.display,color:C.softWhite}}>ZenkyBox</div>
          <div className="text-xs mt-1 font-bold uppercase" style={{fontFamily:F.body,color:C.skyBlue,letterSpacing:"0.05em"}}>📦 Inventory Hub</div>
          <div className="mt-2 text-xs flex items-center gap-1.5" style={{color:synced?"#9BE76A":"rgba(255,255,255,0.5)"}}>
            <div className={`w-1.5 h-1.5 rounded-full ${synced?"bg-green-400":"bg-gray-400"}`}/>
            {synced?"Synced across devices":"Local only"}
          </div>
        </div>
        <nav className="flex-1 px-3 py-3 space-y-0.5 overflow-y-auto">
          {visibleNav.map(item=>{
            const Icon=item.icon;const active=view===item.id;
            return(
              <button key={item.id} onClick={()=>{setView(item.id);setOpen(false);}} className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-bold transition-all" style={{fontFamily:F.body,color:active?C.zenkyPurple:"rgba(255,255,255,0.85)",backgroundColor:active?C.sunshineYellow:"transparent"}}>
                <Icon size={17}/>{item.label}
              </button>
            );
          })}
        </nav>
        <div className="px-4 py-3 border-t" style={{borderColor:"rgba(255,255,255,0.15)"}}>
          {role==="admin"?(
            <div className="flex items-center justify-between">
              <span className="inline-flex items-center gap-1.5 text-xs font-bold" style={{color:C.sunshineYellow,fontFamily:F.body}}><Unlock size={13}/>Admin mode</span>
              <button onClick={onLock} className="text-xs font-bold underline" style={{color:"rgba(255,255,255,0.7)"}}>Lock</button>
            </div>
          ):!canBeAdmin?null:showPinBox?(
            <div className="space-y-2">
              <Input type="password" placeholder="Admin PIN" value={pinInput} onChange={e=>setPinInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleUnlock()} style={{backgroundColor:"rgba(255,255,255,0.9)"}}/>
              <div className="flex gap-2">
                <button onClick={handleUnlock} className="flex-1 py-1.5 rounded-lg text-xs font-bold" style={{backgroundColor:C.sunshineYellow,color:C.zenkyPurple,fontFamily:F.display}}>Unlock</button>
                <button onClick={()=>{setShowPinBox(false);setPinInput("");}} className="text-xs font-bold" style={{color:"rgba(255,255,255,0.7)"}}>Cancel</button>
              </div>
            </div>
          ):(
            <button onClick={()=>setShowPinBox(true)} className="inline-flex items-center gap-1.5 text-xs font-bold" style={{color:"rgba(255,255,255,0.7)",fontFamily:F.body}}><Lock size={13}/>Staff mode — unlock admin</button>
          )}
        </div>
        <div className="px-5 py-2.5 text-xs border-t text-center" style={{borderColor:"rgba(255,255,255,0.15)"}}>
          {currentUserName&&<div className="mb-1" style={{color:"rgba(255,255,255,0.5)"}}>Signed in as <strong style={{color:"rgba(255,255,255,0.8)"}}>{currentUserName}</strong></div>}
          <button onClick={onLogout} className="font-bold underline" style={{color:"rgba(255,255,255,0.7)"}}>Log Out</button>
        </div>
        <div className="px-5 py-3 text-xs border-t text-center font-semibold" style={{borderColor:"rgba(255,255,255,0.15)",color:"rgba(255,255,255,0.6)",fontFamily:F.body}}>
          💝 Thoughtful Gifts. Joyful Moments.
        </div>
      </aside>
    </>
  );
}

/* ═══ DASHBOARD ═══ */
function Dashboard({skus,comboList}){
  const [query,setQuery]=useState("");
  const low=skus.filter(s=>stockStatus(s)!=="healthy").length;
  const crit=skus.filter(s=>stockStatus(s)==="critical").length;
  const ready=comboList.filter(c=>c.ready>0).length;
  const stats=[{label:"Total SKUs",value:skus.length,color:C.zenkyPurple},{label:"Low / Critical",value:low,color:low?C.zenkyOrange:C.mintGreen},{label:"Critical",value:crit,color:crit?C.zenkyPink:C.mintGreen},{label:"Combos Ready",value:`${ready}/${comboList.length}`,color:ready?C.mintGreen:C.zenkyOrange}];
  const filtered=useMemo(()=>{
    const list=[...skus].sort((a,b)=>({critical:0,low:1,healthy:2}[stockStatus(a)]-({critical:0,low:1,healthy:2}[stockStatus(b)])));
    if(!query.trim())return list;const q=query.toLowerCase();return list.filter(s=>s.sku.toLowerCase().includes(q)||s.name.toLowerCase().includes(q));
  },[skus,query]);
  return(
    <div>
      <SectionHeader title="Dashboard" subtitle="Live snapshot of stock health and combo readiness."/>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
        {stats.map(s=><Card key={s.label}><div className="text-xs font-bold uppercase" style={{color:C.lightText,fontFamily:F.body,letterSpacing:"0.03em"}}>{s.label}</div><div className="text-3xl font-black mt-1" style={{fontFamily:F.display,color:s.color}}>{s.value}</div></Card>)}
      </div>
      {skus.length===0?<Empty icon={Package} title="No SKUs yet" message="Add products to start tracking."/>:(
        <Card className="mb-6">
          <div className="flex items-center justify-between gap-3 mb-4 flex-wrap">
            <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>Stock Ledger</h3>
            <div className="relative w-full sm:w-56"><Search size={14} className="absolute left-3 top-3" style={{color:C.lightText}}/><Input placeholder="Search…" value={query} onChange={e=>setQuery(e.target.value)} className="pl-9"/></div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">SKU</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Name</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase hidden lg:table-cell">Initial Stock</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Current Stock</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase hidden lg:table-cell">Sold</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase w-24 hidden sm:table-cell">Health</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Status</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase hidden md:table-cell">Reorder</th></tr></thead>
              <tbody>{filtered.map(s=>{
                const rq=suggestedReorder(s);
                const initial=s.initialStock??s.stock;
                const sold=Math.max(0,initial-s.stock);
                return(
                  <tr key={s.sku} className="border-t" style={{borderColor:C.border}}>
                    <td className="py-2 pr-3" style={{fontFamily:F.mono,fontWeight:600,color:C.darkText}}>{s.sku}</td>
                    <td className="py-2 pr-3">{s.name}</td>
                    <td className="py-2 pr-3 hidden lg:table-cell" style={{fontFamily:F.mono,color:C.lightText}}>{fmt(initial)}</td>
                    <td className="py-2 pr-3 font-bold" style={{fontFamily:F.mono}}>{fmt(s.stock)}</td>
                    <td className="py-2 pr-3 hidden lg:table-cell font-bold" style={{fontFamily:F.mono,color:sold>0?C.zenkyOrange:C.lightText}}>{fmt(sold)}</td>
                    <td className="py-2 pr-3 hidden sm:table-cell"><StockGauge stock={s.stock} reorderLevel={s.reorderLevel}/></td>
                    <td className="py-2 pr-3">{statusStamp(stockStatus(s))}</td>
                    <td className="py-2 pr-3 hidden md:table-cell" style={{fontFamily:F.mono,color:rq?C.zenkyPink:C.lightText,fontWeight:600}}>{rq?`+${fmt(rq)}`:"—"}</td>
                  </tr>
                );
              })}</tbody>
            </table>
          </div>
        </Card>
      )}
      <Card>
        <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Gift Combo Readiness</h3>
        {comboList.length===0?<Empty icon={Boxes} title="No combos" message="Create gift combos to track readiness."/>:(
          <div className="grid sm:grid-cols-2 gap-3">
            {comboList.map(c=>(
              <div key={c.id} className="rounded-xl border-2 p-3 flex items-center justify-between gap-3" style={{borderColor:C.border,backgroundColor:C.bgLight}}>
                <div><div className="font-bold text-sm" style={{color:C.darkText,fontFamily:F.display}}>{c.name}</div><div className="text-xs" style={{fontFamily:F.mono,color:C.lightText}}>{c.sku}</div>{c.ready<=0&&c.bottleneck&&<div className="text-xs mt-0.5 font-bold" style={{color:C.zenkyPink}}>Bottleneck: {c.bottleneck}</div>}</div>
                {c.ready>0?<Stamp tone="mint">Ready ×{c.ready}</Stamp>:<Stamp tone="pink">Short</Stamp>}
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

/* ═══ COMBO READINESS TAB ═══ */
function ComboReadinessView({skus,combos}){
  const [filterSku,setFilterSku]=useState("");
  const skuMap=useMemo(()=>Object.fromEntries(skus.map(s=>[s.sku,s])),[skus]);

  // Enrich combos with readiness info
  const enriched=useMemo(()=>combos.map(combo=>{
    const{ready,bottleneck}=comboReadiness(combo,skuMap);
    return{...combo,ready,bottleneck};
  }),[combos,skuMap]);

  // Filter combos by SKU search
  const filtered=useMemo(()=>{
    if(!filterSku.trim())return enriched;
    const q=filterSku.toLowerCase();
    return enriched.filter(c=>c.components?.some(comp=>comp.sku.toLowerCase().includes(q)||(skuMap[comp.sku]?.name||"").toLowerCase().includes(q)));
  },[enriched,filterSku,skuMap]);

  // Which SKU is highlighted
  const highlightSku=filterSku.trim().toUpperCase();

  function exportReadinessCsv(){
    let csv="Combo Code,Combo Name,Ready Count,Status,Component SKU,Component Name,Need Qty,Stock,Reorder Level,Component Status\n";
    enriched.forEach(c=>{
      c.components?.forEach(comp=>{
        const s=skuMap[comp.sku];
        const st=s?stockStatus(s):"missing";
        csv+=`${c.sku},"${c.name}",${c.ready},${c.ready>0?"Ready":"Short"},${comp.sku},"${s?.name||"Unknown"}",${comp.qty},${s?.stock||0},${s?.reorderLevel||0},${st}\n`;
      });
    });
    downloadCsv("combo_readiness.csv",csv);
  }

  if(combos.length===0) return(<div><SectionHeader title="Combo Readiness"/><Empty icon={ShieldCheck} title="No combos yet" message="Create gift combos first to see readiness."/></div>);

  return(
    <div>
      <SectionHeader
        title="Combo Readiness"
        subtitle="Per-combo breakdown showing which SKUs are short and by how much."
        action={<PrimaryButton onClick={exportReadinessCsv}><Download size={15}/>Export CSV</PrimaryButton>}
      />

      {/* Filter bar */}
      <div className="mb-5 relative">
        <Search size={14} className="absolute left-3 top-3" style={{color:C.lightText}}/>
        <Input placeholder="Filter by SKU code or name…" value={filterSku} onChange={e=>setFilterSku(e.target.value)} className="pl-9"/>
        {filterSku&&<p className="text-xs mt-1" style={{color:C.lightText}}>Showing {filtered.length} of {combos.length} combos containing "{filterSku}"</p>}
      </div>

      {/* Combo cards */}
      <div className="space-y-4">
        {filtered.map(combo=>{
          const isReady=combo.ready>0;
          return(
            <Card key={combo.id} className={isReady?"border-green-200":"border-red-200"} style={{borderColor:isReady?"#bbf7d0":"#fecaca"}}>
              {/* Combo header */}
              <div className="flex items-start justify-between gap-3 mb-3 flex-wrap">
                <div>
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="font-black text-lg" style={{fontFamily:F.display,color:C.darkText}}>{combo.name}</span>
                    {isReady?<Stamp tone="mint">✓ Ready to ship ×{combo.ready}</Stamp>:<Stamp tone="pink">✗ Cannot ship</Stamp>}
                  </div>
                  <div className="text-xs mt-0.5" style={{fontFamily:F.mono,color:C.lightText}}>{combo.sku} · {combo.components?.length} component{combo.components?.length!==1?"s":""}</div>
                </div>
                <div className="text-right">
                  <div className="text-2xl font-black" style={{fontFamily:F.display,color:isReady?C.mintGreen:C.zenkyPink}}>{combo.ready}</div>
                  <div className="text-xs" style={{color:C.lightText}}>units ready</div>
                </div>
              </div>

              {/* Component breakdown */}
              <div className="rounded-xl overflow-hidden border" style={{borderColor:C.border}}>
                <table className="w-full text-sm">
                  <thead style={{backgroundColor:C.bgLight}}>
                    <tr>
                      <th className="py-2 px-3 text-left font-bold text-xs uppercase" style={{color:C.lightText}}>SKU Code</th>
                      <th className="py-2 px-3 text-left font-bold text-xs uppercase hidden sm:table-cell" style={{color:C.lightText}}>Product</th>
                      <th className="py-2 px-3 text-left font-bold text-xs uppercase" style={{color:C.lightText}}>Need</th>
                      <th className="py-2 px-3 text-left font-bold text-xs uppercase" style={{color:C.lightText}}>Stock</th>
                      <th className="py-2 px-3 text-left font-bold text-xs uppercase hidden md:table-cell" style={{color:C.lightText}}>Reorder</th>
                      <th className="py-2 px-3 text-left font-bold text-xs uppercase" style={{color:C.lightText}}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {combo.components?.map(comp=>{
                      const s=skuMap[comp.sku];
                      const status=s?stockStatus(s):"missing";
                      const isOk=status==="healthy";
                      const isBottleneck=comp.sku===combo.bottleneck&&!isReady;
                      const isHighlighted=highlightSku&&comp.sku.toUpperCase().includes(highlightSku);
                      const canMake=s?Math.floor(s.stock/(comp.qty||1)):0;
                      return(
                        <tr key={comp.sku} className="border-t" style={{borderColor:C.border,backgroundColor:isHighlighted?"#FFF3E6":isBottleneck?"#FFF0F5":"transparent"}}>
                          <td className="py-2.5 px-3">
                            <div className="flex items-center gap-1.5">
                              {isBottleneck&&<AlertTriangle size={12} style={{color:C.zenkyPink,flexShrink:0}}/>}
                              <span style={{fontFamily:F.mono,fontWeight:700,color:isBottleneck?C.zenkyPink:C.darkText}}>{comp.sku}</span>
                            </div>
                          </td>
                          <td className="py-2.5 px-3 hidden sm:table-cell" style={{color:C.darkText}}>{s?.name||<span style={{color:C.zenkyPink}}>SKU not found</span>}</td>
                          <td className="py-2.5 px-3" style={{fontFamily:F.mono,color:C.darkText}}>{comp.qty} unit{comp.qty!==1?"s":""}</td>
                          <td className="py-2.5 px-3">
                            <div>
                              <span style={{fontFamily:F.mono,fontWeight:700,color:isOk?C.mintGreen:C.zenkyPink}}>{fmt(s?.stock||0)}</span>
                              <div className="text-xs mt-0.5" style={{color:C.lightText}}>Can make: {canMake}</div>
                            </div>
                          </td>
                          <td className="py-2.5 px-3 hidden md:table-cell" style={{fontFamily:F.mono,color:C.lightText}}>{fmt(s?.reorderLevel||0)}</td>
                          <td className="py-2.5 px-3">{status==="missing"?<Stamp tone="pink">Missing</Stamp>:statusStamp(status)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Bottleneck alert */}
              {!isReady&&combo.bottleneck&&(
                <div className="mt-3 flex items-center gap-2 text-sm p-2.5 rounded-xl" style={{backgroundColor:"#FFF0F5",color:C.zenkyPink}}>
                  <AlertTriangle size={14} className="flex-shrink-0"/>
                  <span className="font-bold">Bottleneck SKU: {combo.bottleneck}{skuMap[combo.bottleneck]?` — ${skuMap[combo.bottleneck].name}`:""}</span>
                  <span className="font-normal">· Restock to unblock this combo.</span>
                </div>
              )}
            </Card>
          );
        })}
      </div>
    </div>
  );
}

/* ═══ CATALOG ═══ */
function Catalog({skus,setSkus,showToast,role,logActivity}){
  const blank={sku:"",name:"",stock:"",reorderLevel:"",procurementCost:"",images:[]};
  const [form,setForm]=useState(blank);
  const [editingSku,setEditingSku]=useState(null);
  const [editValues,setEditValues]=useState({});
  const [pendingDelete,setPendingDelete]=useState(null);
  const [query,setQuery]=useState("");
  const [clearConfirmText,setClearConfirmText]=useState("");
  const [showClearBox,setShowClearBox]=useState(false);

  function clearAllSkus(){
    if(clearConfirmText.trim().toUpperCase()!=="DELETE"){showToast("error",'Type "DELETE" exactly to confirm.');return;}
    const count=skus.length;
    setSkus([]);setShowClearBox(false);setClearConfirmText("");showToast("success","All SKUs cleared. 🗑️");
    logActivity?.("Cleared all SKUs",`${count} SKUs removed`);
  }

  const filtered=skus.filter(s=>{if(!query.trim())return true;const q=query.toLowerCase();return s.sku.toLowerCase().includes(q)||s.name.toLowerCase().includes(q);});

  function autoGenSku(){setForm({...form,sku:generateSkuCode(skus)});}

  function addSku(){
    const code=form.sku.trim(),name=form.name.trim();
    if(!code||!name){showToast("error","SKU code and name are required.");return;}
    if(skus.some(s=>s.sku===code)){showToast("error",`"${code}" already exists.`);return;}
    const initStock=Number(form.stock)||0;
    setSkus([...skus,{sku:code,name,stock:initStock,initialStock:initStock,reorderLevel:Number(form.reorderLevel)||0,procurementCost:Number(form.procurementCost)||0,images:form.images}]);
    setForm(blank);showToast("success",`Added ${code}. ✨`);
    logActivity?.("SKU added",`${code} — ${name}`);
  }

  function saveEdit(code){
    setSkus(skus.map(s=>s.sku===code?{...s,name:editValues.name,stock:Number(editValues.stock)||0,reorderLevel:Number(editValues.reorderLevel)||0,procurementCost:Number(editValues.procurementCost)||0,images:editValues.images||[]}:s));
    setEditingSku(null);showToast("success","Saved. ✨");
    logActivity?.("SKU edited",code);
  }

  function exportCatalogCsv(){
    let csv="SKU,Name,Stock,Reorder Level,Procurement Cost (INR),Status\n";
    skus.forEach(s=>{csv+=`${s.sku},"${s.name}",${s.stock},${s.reorderLevel},${s.procurementCost||0},${stockStatus(s)}\n`;});
    downloadCsv("sku_catalog.csv",csv);
  }

  return(
    <div>
      <SectionHeader title="SKU Catalog" subtitle="Manage products, stock, reorder levels, and procurement costs."
        action={
          <div className="flex items-center gap-2">
            <PrimaryButton onClick={exportCatalogCsv}><Download size={15}/>Export</PrimaryButton>
            {role==="admin"&&skus.length>0&&(
              <button onClick={()=>setShowClearBox(!showClearBox)} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-full text-xs font-bold border-2 transition-colors" style={{borderColor:"#fecaca",color:"#dc2626",fontFamily:F.display}}>
                <Trash2 size={13}/>Clear All
              </button>
            )}
          </div>
        }
      />

      {showClearBox&&(
        <Card className="mb-5" style={{borderColor:"#fecaca",backgroundColor:"#fff5f5"}}>
          <div className="flex items-start gap-3">
            <AlertTriangle size={20} style={{color:"#dc2626",flexShrink:0,marginTop:2}}/>
            <div className="flex-1">
              <p className="font-bold text-sm mb-1" style={{color:"#991b1b",fontFamily:F.display}}>Delete all {skus.length} SKUs?</p>
              <p className="text-xs mb-3" style={{color:"#991b1b",fontFamily:F.body}}>This cannot be undone. Any combos referencing these SKUs will show "SKU not found" until you re-import. Useful for wiping test data before a fresh Bulk Import.</p>
              <div className="flex items-center gap-2 flex-wrap">
                <Input placeholder='Type "DELETE" to confirm' value={clearConfirmText} onChange={e=>setClearConfirmText(e.target.value)} className="max-w-xs" style={{borderColor:"#fecaca"}}/>
                <button onClick={clearAllSkus} className="px-4 py-2.5 rounded-xl text-sm font-bold text-white" style={{backgroundColor:"#dc2626",fontFamily:F.display}}>Delete All SKUs</button>
                <button onClick={()=>{setShowClearBox(false);setClearConfirmText("");}} className="text-sm font-bold" style={{color:C.lightText,fontFamily:F.body}}>Cancel</button>
              </div>
            </div>
          </div>
        </Card>
      )}

      <Card className="mb-6">
        <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Add a New SKU</h3>
        <div className="space-y-4">
          {/* SKU code + auto-generate */}
          <div className="flex gap-2">
            <div className="flex-1"><Input placeholder="SKU code (e.g. ZB-202407-00001)" value={form.sku} onChange={e=>setForm({...form,sku:e.target.value})}/></div>
            <button onClick={autoGenSku} className="flex items-center gap-1.5 px-3 py-2 rounded-xl border-2 text-xs font-bold whitespace-nowrap transition-colors hover:border-purple-400" style={{borderColor:C.border,color:C.zenkyPurple,fontFamily:F.body}}>
              <Zap size={13}/>Auto Generate
            </button>
          </div>
          {/* Name + stock + reorder + cost */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
            <Input placeholder="Product name" className="col-span-2 md:col-span-1" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>
            <Input placeholder="Stock" type="number" value={form.stock} onChange={e=>setForm({...form,stock:e.target.value})}/>
            <Input placeholder="Reorder level" type="number" value={form.reorderLevel} onChange={e=>setForm({...form,reorderLevel:e.target.value})}/>
            <div className="relative"><span className="absolute left-3 top-2.5 text-sm" style={{color:C.lightText}}>₹</span><Input placeholder="Procurement cost" type="number" className="pl-6" value={form.procurementCost} onChange={e=>setForm({...form,procurementCost:e.target.value})}/></div>
          </div>
          <ImageGallery images={form.images} onAddImages={imgs=>setForm({...form,images:[...form.images,...imgs]})} onRemoveImage={i=>setForm({...form,images:form.images.filter((_,j)=>j!==i)})}/>
        </div>
        <div className="mt-4"><PrimaryButton onClick={addSku}><Plus size={16}/>Add SKU</PrimaryButton></div>
      </Card>

      {skus.length===0?<Empty icon={Package} title="Catalog is empty" message="Add your first SKU above."/>:(
        <Card>
          <div className="flex items-center justify-between mb-4 gap-3 flex-wrap">
            <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>{skus.length} SKU{skus.length!==1?"s":""}</h3>
            <div className="relative w-full sm:w-56"><Search size={14} className="absolute left-3 top-3" style={{color:C.lightText}}/><Input placeholder="Search" value={query} onChange={e=>setQuery(e.target.value)} className="pl-9"/></div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr style={{color:C.lightText}}><th className="py-2 pr-2 text-left font-bold text-xs uppercase hidden sm:table-cell">Img</th><th className="py-2 pr-2 text-left font-bold text-xs uppercase">SKU</th><th className="py-2 pr-2 text-left font-bold text-xs uppercase">Name</th><th className="py-2 pr-2 text-left font-bold text-xs uppercase">Stock</th><th className="py-2 pr-2 text-left font-bold text-xs uppercase hidden md:table-cell">Reorder</th><th className="py-2 pr-2 text-left font-bold text-xs uppercase hidden md:table-cell">Cost</th><th className="py-2 pr-2 text-left font-bold text-xs uppercase">Status</th><th/></tr></thead>
              <tbody>
                {filtered.map(s=>{
                  const isEdit=editingSku===s.sku;
                  return(
                    <tr key={s.sku} className="border-t" style={{borderColor:C.border}}>
                      <td className="py-2 pr-2 hidden sm:table-cell">{s.images?.[0]&&<img src={s.images[0]} alt="" className="w-8 h-8 rounded object-cover"/>}</td>
                      <td className="py-2 pr-2" style={{fontFamily:F.mono,fontWeight:600,color:C.darkText,whiteSpace:"nowrap"}}>{s.sku}</td>
                      <td className="py-2 pr-2">{isEdit?<Input value={editValues.name} onChange={e=>setEditValues({...editValues,name:e.target.value})}/>:s.name}</td>
                      <td className="py-2 pr-2">{isEdit?<Input type="number" value={editValues.stock} onChange={e=>setEditValues({...editValues,stock:e.target.value})} className="w-20"/>:<span style={{fontFamily:F.mono,fontWeight:600}}>{fmt(s.stock)}</span>}</td>
                      <td className="py-2 pr-2 hidden md:table-cell">{isEdit?<Input type="number" value={editValues.reorderLevel} onChange={e=>setEditValues({...editValues,reorderLevel:e.target.value})} className="w-20"/>:<span style={{fontFamily:F.mono}}>{fmt(s.reorderLevel)}</span>}</td>
                      <td className="py-2 pr-2 hidden md:table-cell">{isEdit?<Input type="number" value={editValues.procurementCost} onChange={e=>setEditValues({...editValues,procurementCost:e.target.value})} className="w-24"/>:<span style={{fontFamily:F.mono}}>{fmtINR(s.procurementCost||0)}</span>}</td>
                      <td className="py-2 pr-2">{statusStamp(stockStatus(s))}</td>
                      <td className="py-2 pr-2">
                        {role!=="admin"?null:(
                          <div className="flex items-center gap-1">
                            {isEdit?(<><GhostButton title="Save" onClick={()=>saveEdit(s.sku)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setEditingSku(null)}><X size={13}/></GhostButton></>)
                            :pendingDelete===s.sku?(<><GhostButton title="Confirm" onClick={()=>{setSkus(skus.filter(x=>x.sku!==s.sku));setPendingDelete(null);showToast("success",`Removed ${s.sku}.`);logActivity?.("SKU deleted",s.sku);}}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setPendingDelete(null)}><X size={13}/></GhostButton></>)
                            :(<><GhostButton title="Edit" onClick={()=>{setEditingSku(s.sku);setEditValues({name:s.name,stock:s.stock,reorderLevel:s.reorderLevel,procurementCost:s.procurementCost||0,images:s.images||[]});}}><Pencil size={13}/></GhostButton><GhostButton title="Delete" onClick={()=>setPendingDelete(s.sku)}><Trash2 size={13}/></GhostButton></>)}
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Card>
      )}
    </div>
  );
}

/* ═══ COMBOS VIEW ═══ */
function CombosView({skus,combos,setCombos,showToast,role,logActivity}){
  const blank={sku:"",name:"",components:[{sku:"",qty:1}],images:[]};
  const [form,setForm]=useState(blank);
  const [editId,setEditId]=useState(null);
  const [pendingDelete,setPendingDelete]=useState(null);
  const [expanded,setExpanded]=useState({});
  const [clearConfirmText,setClearConfirmText]=useState("");
  const [showClearBox,setShowClearBox]=useState(false);
  const skuMap=useMemo(()=>Object.fromEntries(skus.map(s=>[s.sku,s])),[skus]);

  function clearAllCombos(){
    if(clearConfirmText.trim().toUpperCase()!=="DELETE"){showToast("error",'Type "DELETE" exactly to confirm.');return;}
    const count=combos.length;
    setCombos([]);setShowClearBox(false);setClearConfirmText("");showToast("success","All combos cleared. 🗑️");
    logActivity?.("Cleared all combos",`${count} combos removed`);
  }

  function saveCombo(){
    const code=form.sku.trim(),name=form.name.trim();
    const comps=form.components.filter(c=>c.sku).map(c=>({sku:c.sku,qty:Number(c.qty)||1}));
    if(!code||!name){showToast("error","Bundle code and name required.");return;}
    if(!comps.length){showToast("error","Add at least one component.");return;}
    if(skus.some(s=>s.sku===code)||combos.some(c=>c.sku===code&&c.id!==editId)){showToast("error",`Code "${code}" already used.`);return;}
    if(editId){setCombos(combos.map(c=>c.id===editId?{...c,sku:code,name,components:comps,images:form.images}:c));showToast("success",`Updated ${name}. ✨`);logActivity?.("Combo edited",`${code} — ${name}`);}
    else{setCombos([...combos,{id:Date.now().toString(),sku:code,name,components:comps,images:form.images}]);showToast("success",`Added ${name}. 🎁`);logActivity?.("Combo added",`${code} — ${name}`);}
    setForm(blank);setEditId(null);
  }

  function exportCombosCsv(){
    let csv="Combo Code,Combo Name,Component SKU,Component Name,Qty\n";
    combos.forEach(c=>{c.components?.forEach(comp=>{csv+=`${c.sku},"${c.name}",${comp.sku},"${skuMap[comp.sku]?.name||""}",${comp.qty}\n`;});});
    downloadCsv("combos.csv",csv);
  }

  return(
    <div>
      <SectionHeader title="Gift Combos" subtitle="Define gift bundles from your SKUs."
        action={(
          <div className="flex items-center gap-2">
            {combos.length>0&&<PrimaryButton onClick={exportCombosCsv}><Download size={15}/>Export</PrimaryButton>}
            {role==="admin"&&combos.length>0&&(
              <button onClick={()=>setShowClearBox(!showClearBox)} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-full text-xs font-bold border-2 transition-colors" style={{borderColor:"#fecaca",color:"#dc2626",fontFamily:F.display}}>
                <Trash2 size={13}/>Clear All
              </button>
            )}
          </div>
        )}
      />

      {showClearBox&&(
        <Card className="mb-5" style={{borderColor:"#fecaca",backgroundColor:"#fff5f5"}}>
          <div className="flex items-start gap-3">
            <AlertTriangle size={20} style={{color:"#dc2626",flexShrink:0,marginTop:2}}/>
            <div className="flex-1">
              <p className="font-bold text-sm mb-1" style={{color:"#991b1b",fontFamily:F.display}}>Delete all {combos.length} combos?</p>
              <p className="text-xs mb-3" style={{color:"#991b1b",fontFamily:F.body}}>This cannot be undone. Your SKU catalog stays untouched — only combo bundles are removed. Useful for wiping test data before a fresh Bulk Import.</p>
              <div className="flex items-center gap-2 flex-wrap">
                <Input placeholder='Type "DELETE" to confirm' value={clearConfirmText} onChange={e=>setClearConfirmText(e.target.value)} className="max-w-xs" style={{borderColor:"#fecaca"}}/>
                <button onClick={clearAllCombos} className="px-4 py-2.5 rounded-xl text-sm font-bold text-white" style={{backgroundColor:"#dc2626",fontFamily:F.display}}>Delete All Combos</button>
                <button onClick={()=>{setShowClearBox(false);setClearConfirmText("");}} className="text-sm font-bold" style={{color:C.lightText,fontFamily:F.body}}>Cancel</button>
              </div>
            </div>
          </div>
        </Card>
      )}

      <Card className="mb-6">
        <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>{editId?"Edit Combo":"Create a Gift Combo"}</h3>
        {skus.length===0?<p className="text-sm" style={{color:C.lightText}}>Add SKUs to your Catalog first.</p>:(
          <div className="space-y-4">
            <div className="grid sm:grid-cols-3 gap-2">
              <Input placeholder="Bundle code (e.g. COMBO-A)" value={form.sku} onChange={e=>setForm({...form,sku:e.target.value})}/>
              <Input placeholder="Combo name" className="sm:col-span-2" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>
            </div>
            <div>
              <label className="text-xs font-bold uppercase block mb-2" style={{color:C.lightText,fontFamily:F.body,letterSpacing:"0.02em"}}>Components</label>
              <div className="space-y-2">
                {form.components.map((c,i)=>(
                  <div key={i} className="flex items-center gap-2">
                    <Select value={c.sku} onChange={e=>setForm({...form,components:form.components.map((x,j)=>j===i?{...x,sku:e.target.value}:x)})}>
                      <option value="">Select SKU…</option>
                      {skus.map(s=><option key={s.sku} value={s.sku}>{s.sku} — {s.name}</option>)}
                    </Select>
                    <Input type="number" min="1" placeholder="Qty" value={c.qty} onChange={e=>setForm({...form,components:form.components.map((x,j)=>j===i?{...x,qty:e.target.value}:x)})} className="w-20"/>
                    <GhostButton title="Remove" onClick={()=>setForm({...form,components:form.components.filter((_,j)=>j!==i)})}><X size={13}/></GhostButton>
                  </div>
                ))}
              </div>
              <button onClick={()=>setForm({...form,components:[...form.components,{sku:"",qty:1}]})} className="inline-flex items-center gap-1 text-sm font-bold mt-2" style={{color:C.zenkyOrange,fontFamily:F.body}}><Plus size={13}/>Add component</button>
            </div>
            <ImageGallery images={form.images} onAddImages={imgs=>setForm({...form,images:[...form.images,...imgs]})} onRemoveImage={i=>setForm({...form,images:form.images.filter((_,j)=>j!==i)})}/>
            <div className="flex items-center gap-2">
              <PrimaryButton onClick={saveCombo}><Save size={15}/>{editId?"Save changes":"Create combo"}</PrimaryButton>
              {editId&&<button onClick={()=>{setForm(blank);setEditId(null);}} className="text-sm font-bold" style={{color:C.lightText,fontFamily:F.body}}>Cancel</button>}
            </div>
          </div>
        )}
      </Card>

      {combos.length===0?<Empty icon={Boxes} title="No combos yet" message="Create a gift combo above."/>:(
        <div className="space-y-3">
          {combos.map(c=>{
            const{ready,bottleneck}=comboReadiness(c,skuMap);const isOpen=expanded[c.id];
            return(
              <Card key={c.id}>
                <div className="flex items-center justify-between gap-3 flex-wrap">
                  <button className="flex items-center gap-3 flex-1 text-left min-w-0" onClick={()=>setExpanded({...expanded,[c.id]:!isOpen})}>
                    {isOpen?<ChevronDown size={16} className="flex-shrink-0"/>:<ChevronRight size={16} className="flex-shrink-0"/>}
                    <div className="min-w-0">
                      <div className="font-bold text-sm truncate" style={{color:C.darkText,fontFamily:F.display}}>{c.name}</div>
                      <div className="text-xs" style={{fontFamily:F.mono,color:C.lightText}}>{c.sku} · {c.components?.length} component{c.components?.length!==1?"s":""}</div>
                    </div>
                  </button>
                  <div className="flex items-center gap-2">
                    {ready>0?<Stamp tone="mint">Ready ×{ready}</Stamp>:<Stamp tone="pink">Short</Stamp>}
                    {role==="admin"&&(pendingDelete===c.id?(<><GhostButton title="Confirm" onClick={()=>{setCombos(combos.filter(x=>x.id!==c.id));setPendingDelete(null);logActivity?.("Combo deleted",c.sku);}}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setPendingDelete(null)}><X size={13}/></GhostButton></>)
                    :(<><GhostButton title="Edit" onClick={()=>{setForm({sku:c.sku,name:c.name,components:c.components?.map(x=>({...x}))||[],images:c.images||[]});setEditId(c.id);window.scrollTo(0,0);}}><Pencil size={13}/></GhostButton><GhostButton title="Delete" onClick={()=>setPendingDelete(c.id)}><Trash2 size={13}/></GhostButton></>))}
                  </div>
                </div>
                {isOpen&&(
                  <div className="mt-3 pt-3 border-t space-y-3" style={{borderColor:C.border}}>
                    {c.images?.length>0&&<div className="grid grid-cols-4 sm:grid-cols-6 gap-2">{c.images.map((img,i)=><img key={i} src={img} alt="" className="w-full aspect-square object-cover rounded-lg"/>)}</div>}
                    {c.components?.map(comp=>{const s=skuMap[comp.sku];const isBo=comp.sku===bottleneck&&ready<=0;return(<div key={comp.sku} className="flex items-center justify-between py-1.5 text-sm"><span style={{fontFamily:F.mono,color:isBo?C.zenkyPink:C.darkText,fontWeight:isBo?700:600}}>{comp.sku} {s?`(${s.name})`:"(missing)"}</span><span style={{color:C.lightText,fontFamily:F.mono}}>need {comp.qty} · have {s?fmt(s.stock):0}</span></div>);})}
                  </div>
                )}
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ═══ BULK IMPORT ═══ */
function BulkImportView({skus,combos,setSkus,setCombos,showToast,logActivity}){
  const [stage,setStage]=useState("idle");const [fileName,setFileName]=useState("");
  const [preview,setPreview]=useState({skus:[],combos:[]});const ref=useRef(null);
  const [importMode,setImportMode]=useState("merge"); // "merge" | "replace"
  const [replaceConfirmText,setReplaceConfirmText]=useState("");
  const [showReplaceConfirm,setShowReplaceConfirm]=useState(false);

  // Detect component SKUs referenced by imported combos that don't exist anywhere
  // (not in current catalog, not in the SKUs sheet being imported) — catches
  // formatting mistakes like commas instead of semicolons in the Components column.
  const unmatchedComponents=useMemo(()=>{
    const known=new Set([...skus.map(s=>s.sku),...preview.skus.map(s=>s.sku)]);
    const bad=new Map(); // sku -> [combo names]
    preview.combos.forEach(c=>{
      c.components?.forEach(comp=>{
        if(!known.has(comp.sku)){
          if(!bad.has(comp.sku))bad.set(comp.sku,[]);
          bad.get(comp.sku).push(c.name);
        }
      });
    });
    return Array.from(bad.entries()).map(([sku,combos])=>({sku,combos}));
  },[preview,skus]);
  function parseImportData(skuRows,comboRows=[]){
    const iSkus=skuRows.filter(r=>r.SKU&&r.Name).map(r=>{const stock=Number(r.Stock)||0;return{sku:String(r.SKU).trim(),name:String(r.Name).trim(),stock,initialStock:stock,reorderLevel:Number(r["Reorder Level"])||0,procurementCost:Number(r["Procurement Cost"])||0,images:[]};});
    const iCombos=comboRows.filter(r=>r["Combo Code"]&&r["Combo Name"]).map(r=>({id:Date.now().toString()+Math.random(),sku:String(r["Combo Code"]).trim(),name:String(r["Combo Name"]).trim(),components:parseComponentsString(r.Components),images:[]}));
    setPreview({skus:iSkus,combos:iCombos});setStage("preview");
  }
  function processFile(file){
    if(!file)return;setFileName(file.name);const ext=file.name.split(".").pop().toLowerCase();
    if(ext==="csv"){Papa.parse(file,{header:true,skipEmptyLines:true,complete:res=>parseImportData(res.data),error:()=>showToast("error","Could not parse CSV.")});}
    else if(ext==="xlsx"||ext==="xls"){const r=new FileReader();r.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:"array"});const ss=wb.Sheets["SKUs"]||wb.Sheets[wb.SheetNames[0]];const cs=wb.Sheets["Combos"];parseImportData(XLSX.utils.sheet_to_json(ss,{defval:""}),cs?XLSX.utils.sheet_to_json(cs,{defval:""}):[]);}catch{showToast("error","Could not parse spreadsheet.");}};r.readAsArrayBuffer(file);}
    else showToast("error","Upload a .csv or .xlsx file.");
  }
  function confirmImport(){
    if(importMode==="replace"){
      // Hard block: never allow Replace All to wipe the catalog with an empty parse result —
      // 0 SKUs and 0 combos almost always means the file was missing a sheet (parsing failure),
      // not a genuine intent to empty the catalog.
      if(preview.skus.length===0&&preview.combos.length===0){
        showToast("error","This file parsed as 0 SKUs and 0 combos — refusing to Replace All, since that would wipe your entire catalog. Check the file has both 'SKUs' and 'Combos' sheets, or use Merge instead.");
        return;
      }
      if(!showReplaceConfirm){setShowReplaceConfirm(true);return;} // require an extra click + typed confirmation first
      if(replaceConfirmText.trim().toUpperCase()!=="REPLACE"){showToast("error",'Type "REPLACE" exactly to confirm.');return;}
      // Replace mode: wipe existing data, use only what's in the file
      setSkus(preview.skus);
      setCombos(preview.combos.map(c=>({...c,id:c.id||Date.now().toString()+Math.random()})));
      setStage("imported");setShowReplaceConfirm(false);setReplaceConfirmText("");
      showToast("success",`Replaced catalog: ${preview.skus.length} SKUs & ${preview.combos.length} combos. 🔄`);
      logActivity?.("Bulk Import — Replace All",`${fileName}: ${preview.skus.length} SKUs & ${preview.combos.length} combos`);
      return;
    }
    // Merge mode: update matches, add new
    const ms=[...skus],mc=[...combos];
    preview.skus.forEach(n=>{const i=ms.findIndex(s=>s.sku===n.sku);i>=0?ms[i]={...ms[i],...n,initialStock:ms[i].initialStock??n.initialStock}:ms.push(n);});
    preview.combos.forEach(n=>{const i=mc.findIndex(c=>c.sku===n.sku);i>=0?mc[i]={...mc[i],...n}:mc.push(n);});
    setSkus(ms);setCombos(mc);setStage("imported");showToast("success",`Imported ${preview.skus.length} SKUs & ${preview.combos.length} combos. ✨`);
    logActivity?.("Bulk Import — Merge",`${fileName}: ${preview.skus.length} SKUs & ${preview.combos.length} combos`);
  }
  function reset(){setStage("idle");setPreview({skus:[],combos:[]});setFileName("");setShowReplaceConfirm(false);setReplaceConfirmText("");if(ref.current)ref.current.value="";}
  return(
    <div>
      <SectionHeader title="Bulk Import" subtitle="Upload SKU & combo catalog from Excel or CSV."/>
      <Card className="mb-5"><h3 className="font-bold mb-3" style={{fontFamily:F.display,color:C.darkText}}>📋 Required Format</h3>
        <div className="space-y-2 text-sm" style={{fontFamily:F.body,color:C.darkText}}>
          <div><p className="font-bold mb-1">SKUs sheet:</p><code className="block bg-gray-100 p-2 rounded text-xs" style={{fontFamily:F.mono}}>SKU | Name | Stock | Reorder Level | Procurement Cost</code></div>
          <div><p className="font-bold mb-1">Combos sheet:</p><code className="block bg-gray-100 p-2 rounded text-xs" style={{fontFamily:F.mono}}>Combo Code | Combo Name | Components (SKU001:2;SKU002:1)</code></div>
        </div>
        <div className="flex gap-4 mt-3">
          <button onClick={()=>downloadCsv("sku_import_template.csv","SKU,Name,Stock,Reorder Level,Procurement Cost\nZB-ST-MBK-002,Magnetic Bookmark Set,100,20,45\n")} className="inline-flex items-center gap-1.5 text-xs font-bold" style={{color:C.zenkyOrange,fontFamily:F.body}}><Download size={13}/>Download SKU template</button>
          <button onClick={()=>downloadCsv("combo_import_template.csv","Combo Code,Combo Name,Components\nCOMBO-A,Starter Gift Set,ZB-ST-MBK-002:1;ZB-CP-ZPP-001:2\n")} className="inline-flex items-center gap-1.5 text-xs font-bold" style={{color:C.zenkyOrange,fontFamily:F.body}}><Download size={13}/>Download Combo template</button>
        </div>
      </Card>
      <Card>
        {stage==="idle"&&(<div className="rounded-2xl border-2 border-dashed p-8 text-center" style={{borderColor:C.zenkyPink,backgroundColor:"#FFF8FC"}}><Upload size={32} className="mx-auto mb-3" style={{color:C.zenkyPink}}/><p className="font-bold mb-4" style={{color:C.darkText,fontFamily:F.display}}>Choose your Excel or CSV file</p><input ref={ref} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={e=>processFile(e.target.files[0])}/><PrimaryButton onClick={()=>ref.current?.click()}><Upload size={15}/>Select File</PrimaryButton></div>)}
        {stage==="preview"&&(<div><div className="flex items-center justify-between mb-4 flex-wrap gap-2"><div><h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>Preview</h3><p className="text-sm" style={{color:C.lightText}}>{fileName} — {preview.skus.length} SKUs, {preview.combos.length} combos</p></div><button onClick={reset} className="text-sm font-bold" style={{color:C.lightText}}>Change file</button></div>
          {preview.skus.length>0&&<div className="overflow-x-auto mb-4"><table className="w-full text-xs"><thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold uppercase">SKU</th><th className="py-2 pr-3 text-left font-bold uppercase">Name</th><th className="py-2 pr-3 text-left font-bold uppercase">Stock</th><th className="py-2 pr-3 text-left font-bold uppercase">Cost</th></tr></thead><tbody>{preview.skus.slice(0,5).map(s=><tr key={s.sku} className="border-t" style={{borderColor:C.border}}><td className="py-1.5 pr-3" style={{fontFamily:F.mono,fontWeight:600}}>{s.sku}</td><td className="py-1.5 pr-3">{s.name}</td><td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{s.stock}</td><td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmtINR(s.procurementCost||0)}</td></tr>)}</tbody></table>{preview.skus.length>5&&<p className="text-xs mt-1" style={{color:C.lightText}}>…and {preview.skus.length-5} more</p>}</div>}

          {unmatchedComponents.length>0&&(
            <div className="mb-4 p-3.5 rounded-xl" style={{backgroundColor:"#fff5f5",border:"2px solid #fecaca"}}>
              <div className="flex items-start gap-2">
                <AlertTriangle size={16} style={{color:"#dc2626",flexShrink:0,marginTop:2}}/>
                <div className="flex-1 min-w-0">
                  <p className="font-bold text-sm mb-1.5" style={{color:"#991b1b",fontFamily:F.display}}>
                    {unmatchedComponents.length} component SKU{unmatchedComponents.length!==1?"s":""} not found in catalog
                  </p>
                  <p className="text-xs mb-2" style={{color:"#991b1b"}}>These combos will show "Missing" until fixed. Common cause: the Components column used commas instead of semicolons (e.g. "SKU1,SKU2" instead of "SKU1:1;SKU2:1").</p>
                  <div className="space-y-1 max-h-32 overflow-y-auto">
                    {unmatchedComponents.map(u=>(
                      <div key={u.sku} className="text-xs" style={{fontFamily:F.mono,color:"#991b1b"}}>
                        <strong>{u.sku.length>60?u.sku.slice(0,60)+"…":u.sku}</strong> — used in: {u.combos.join(", ")}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Import mode toggle */}
          <div className="mb-4 p-3 rounded-xl" style={{backgroundColor:C.bgLight}}>
            <label className="text-xs font-bold uppercase block mb-2" style={{color:C.lightText,fontFamily:F.body,letterSpacing:"0.02em"}}>Import Mode</label>
            <div className="grid grid-cols-2 gap-2">
              <button onClick={()=>{setImportMode("merge");setShowReplaceConfirm(false);setReplaceConfirmText("");}} className="p-3 rounded-xl border-2 text-left transition-all" style={{borderColor:importMode==="merge"?C.zenkyPurple:C.border,backgroundColor:importMode==="merge"?"#F3EEFF":C.softWhite}}>
                <div className="font-bold text-sm flex items-center gap-1.5" style={{color:importMode==="merge"?C.zenkyPurple:C.darkText,fontFamily:F.display}}>{importMode==="merge"&&<Check size={14}/>}Merge</div>
                <div className="text-xs mt-0.5" style={{color:C.lightText}}>Update matches, keep everything else</div>
              </button>
              <button onClick={()=>setImportMode("replace")} className="p-3 rounded-xl border-2 text-left transition-all" style={{borderColor:importMode==="replace"?"#dc2626":C.border,backgroundColor:importMode==="replace"?"#fff5f5":C.softWhite}}>
                <div className="font-bold text-sm flex items-center gap-1.5" style={{color:importMode==="replace"?"#dc2626":C.darkText,fontFamily:F.display}}>{importMode==="replace"&&<Check size={14}/>}Replace All</div>
                <div className="text-xs mt-0.5" style={{color:C.lightText}}>⚠️ Wipes existing catalog first</div>
              </button>
            </div>
          </div>

          {importMode==="replace"&&(
            <div className="mb-4 flex items-start gap-2 p-3 rounded-xl text-xs" style={{backgroundColor:"#fff5f5",color:"#991b1b"}}>
              <AlertTriangle size={14} className="flex-shrink-0 mt-0.5"/>
              <span>Replace mode will delete all {skus.length} current SKUs and {combos.length} current combos, then load only what's in this file. Ideal for repeated test uploads.</span>
            </div>
          )}

          {showReplaceConfirm&&(
            <Card className="mb-4" style={{borderColor:"#fecaca",backgroundColor:"#fff5f5"}}>
              <div className="flex items-start gap-3">
                <AlertTriangle size={20} style={{color:"#dc2626",flexShrink:0,marginTop:2}}/>
                <div className="flex-1">
                  <p className="font-bold text-sm mb-1" style={{color:"#991b1b",fontFamily:F.display}}>Final check: replace {skus.length} SKUs & {combos.length} combos with {preview.skus.length} SKUs & {preview.combos.length} combos?</p>
                  <p className="text-xs mb-3" style={{color:"#991b1b"}}>This cannot be undone, and since sync is on, it updates on every device immediately.</p>
                  <div className="flex items-center gap-2 flex-wrap">
                    <Input placeholder='Type "REPLACE" to confirm' value={replaceConfirmText} onChange={e=>setReplaceConfirmText(e.target.value)} className="max-w-xs" style={{borderColor:"#fecaca"}}/>
                    <button onClick={confirmImport} className="px-4 py-2.5 rounded-xl text-sm font-bold text-white" style={{backgroundColor:"#dc2626",fontFamily:F.display}}>Replace Now</button>
                    <button onClick={()=>{setShowReplaceConfirm(false);setReplaceConfirmText("");}} className="text-sm font-bold" style={{color:C.lightText,fontFamily:F.body}}>Cancel</button>
                  </div>
                </div>
              </div>
            </Card>
          )}

          <div className="flex gap-2"><PrimaryButton onClick={confirmImport} tone={importMode==="replace"?"pink":undefined}><Check size={15}/>{importMode==="replace"?"Replace & Import":"Confirm Import"}</PrimaryButton><button onClick={reset} className="text-sm font-bold" style={{color:C.lightText}}>Cancel</button></div></div>)}
        {stage==="imported"&&(<div className="text-center py-8"><Check size={32} className="mx-auto mb-3" style={{color:C.mintGreen}}/><p className="font-bold text-lg" style={{color:C.darkText,fontFamily:F.display}}>Import complete!</p><p className="text-sm mt-1" style={{color:C.lightText}}>{preview.skus.length} SKUs & {preview.combos.length} combos added.</p><div className="mt-5"><PrimaryButton onClick={reset}><Upload size={15}/>Import another</PrimaryButton></div></div>)}
      </Card>
    </div>
  );
}

/* ═══ UPLOAD SALES ═══ */
function UploadView({skus,combos,setSkus,reports,setReports,salesLines,setSalesLines,logActivity,showToast}){
  const [channel,setChannel]=useState("amazon"); // "amazon" | "website"
  const [entryMode,setEntryMode]=useState("file"); // "file" | "manual"
  const [stage,setStage]=useState("idle");const [fileName,setFileName]=useState("");
  const [rawRows,setRawRows]=useState([]);const [extraCols,setExtraCols]=useState({});
  const [repairNote,setRepairNote]=useState(null);
  const [skipDuplicates,setSkipDuplicates]=useState(true);
  const [weekLabel,setWeekLabel]=useState("");const ref=useRef(null);
  const skuMap=useMemo(()=>Object.fromEntries(skus.map(s=>[s.sku,s])),[skus]);
  const comboMap=useMemo(()=>Object.fromEntries(combos.map(c=>[c.sku,c])),[combos]);

  // Every order-id already recorded from a previously applied report
  const existingOrderIds=useMemo(()=>new Set(salesLines.map(l=>l.orderId).filter(Boolean)),[salesLines]);

  // Which order-ids in THIS file have already been counted before
  const duplicateOrderIds=useMemo(()=>{
    if(!rawRows.length)return new Set();
    return new Set(rawRows.filter(r=>r.orderId&&existingOrderIds.has(r.orderId)).map(r=>r.orderId));
  },[rawRows,existingOrderIds]);

  // Rows actually used for stock deduction / sales report, after optionally skipping duplicates
  const effectiveRows=useMemo(()=>{
    if(!skipDuplicates||duplicateOrderIds.size===0)return rawRows;
    return rawRows.filter(r=>!(r.orderId&&duplicateOrderIds.has(r.orderId)));
  },[rawRows,skipDuplicates,duplicateOrderIds]);

  const aggregated=useMemo(()=>{
    const totals={};
    effectiveRows.forEach(r=>{if(!r.sku||isNaN(r.qty))return;totals[r.sku]=(totals[r.sku]||0)+r.qty;});
    return Object.entries(totals).map(([code,qty])=>{const combo=combos.find(c=>c.sku===code),sku=skuMap[code];const mt=combo?"combo":sku?"direct":"unknown";return{code,qty,matchType:mt,matchName:combo?.name||sku?.name||"Not in catalog"};});
  },[effectiveRows,combos,skuMap]);

  function processRows(rows){
    if(!rows?.length){showToast("error","File is empty.");return;}
    const hk=Object.keys(rows[0]);const sk=hk.find(k=>k.toLowerCase().includes("sku"));const qk=hk.find(k=>/quantity|qty|units/i.test(k));
    if(!sk||!qk){showToast("error",'Expected "sku" and "quantity" columns.');return;}
    const extra=detectExtraColumns(hk);
    setExtraCols(extra);
    setRawRows(rows.map(r=>({sku:String(r[sk]||"").trim(),qty:parseFloat(r[qk]),date:extra.dateKey?r[extra.dateKey]:null,orderId:extra.orderIdKey?String(r[extra.orderIdKey]||"").trim():"",buyer:extra.buyerKey?String(r[extra.buyerKey]||"").trim():"",city:extra.cityKey?String(r[extra.cityKey]||"").trim():"",state:extra.stateKey?String(r[extra.stateKey]||"").trim():"",price:extra.priceKey?parseFloat(r[extra.priceKey]):null})).filter(r=>r.sku&&!isNaN(r.qty)));
    setStage("parsed");
  }

  /* Turns raw array-of-arrays (with header row first) into array-of-objects,
     auto-repairing Amazon's colon-split date corruption first if it's present. */
  function fromRawGrid(grid){
    if(!grid?.length)return[];
    const headerRow=grid[0].map(h=>String(h??"").trim());
    let dataRows=grid.slice(1).filter(r=>r.some(v=>v!==""&&v!=null));
    if(channel==="amazon"){
      const{repaired,rows:fixed,reason}=repairShiftedColumns(headerRow,dataRows);
      if(repaired){
        dataRows=fixed;
        setRepairNote({type:"success",msg:"Detected and auto-corrected shifted columns caused by split timestamp fields — uploaded exactly as exported, no manual fixing needed."});
      }else if(reason==="inconsistent"){
        setRepairNote({type:"warning",msg:"Columns looked misaligned but didn't match the usual pattern — proceeding with the file as-is. Double-check the SKU/Quantity preview below before applying."});
      }else{
        setRepairNote(null);
      }
    }else{
      setRepairNote(null);
    }
    return dataRows.map(r=>Object.fromEntries(headerRow.map((h,i)=>[h,r[i]])));
  }

  function handleFile(file){
    if(!file)return;setFileName(file.name);setRepairNote(null);const ext=file.name.split(".").pop().toLowerCase();
    if(ext==="csv"){
      Papa.parse(file,{header:false,skipEmptyLines:true,complete:res=>processRows(fromRawGrid(res.data)),error:()=>showToast("error","Could not parse CSV.")});
    }else if(ext==="xlsx"||ext==="xls"){
      const r=new FileReader();
      r.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:"array"});const grid=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});processRows(fromRawGrid(grid));}catch{showToast("error","Could not parse spreadsheet.");}};
      r.readAsArrayBuffer(file);
    }else showToast("error","Upload a .csv or .xlsx file.");
  }

  function applyReport(){
    if(!weekLabel.trim()){showToast("error","Add a report label.");return;}
    const skippedCount=duplicateOrderIds.size;
    const skuBefore={};skus.forEach(s=>skuBefore[s.sku]=s.stock);
    const updated=Object.fromEntries(skus.map(s=>[s.sku,{...s}]));
    aggregated.forEach(({code,qty,matchType})=>{if(matchType==="combo"){const c=combos.find(x=>x.sku===code);c?.components?.forEach(comp=>{if(updated[comp.sku])updated[comp.sku].stock=Math.max(0,updated[comp.sku].stock-comp.qty*qty);});}else if(matchType==="direct"&&updated[code])updated[code].stock=Math.max(0,updated[code].stock-qty);});
    const newSkus=Object.values(updated);const bMap=Object.fromEntries(skus.map(s=>[s.sku,s])),aMap=Object.fromEntries(newSkus.map(s=>[s.sku,s]));
    const reportId=Date.now().toString();
    const report={id:reportId,label:weekLabel.trim(),fileName,channel,appliedAt:new Date().toISOString(),
      skuLines:newSkus.map(s=>({sku:s.sku,name:s.name,opening:skuBefore[s.sku]??s.stock,sold:(skuBefore[s.sku]??s.stock)-s.stock,closing:s.stock,reorderLevel:s.reorderLevel,status:stockStatus(s)})),
      comboLines:combos.map(c=>({sku:c.sku,name:c.name,readyBefore:comboReadiness(c,bMap).ready,readyAfter:comboReadiness(c,aMap).ready,bottleneck:comboReadiness(c,aMap).bottleneck})),
      unmatched:aggregated.filter(a=>a.matchType==="unknown"),
      skippedDuplicates:skippedCount};

    const newLines=effectiveRows.map((row,i)=>{
      const combo=comboMap[row.sku],sku=skuMap[row.sku];
      const matchType=combo?"combo":sku?"direct":"unknown";
      const name=combo?.name||sku?.name||row.sku;
      const unitCost=unitCostOf(row.sku,matchType,skuMap,comboMap);
      const cost=unitCost*row.qty;
      const revenue=row.price!=null&&!isNaN(row.price)?row.price*row.qty:0;
      return{id:`${reportId}-${i}`,reportId,channel,date:row.date||report.appliedAt,sku:row.sku,name,matchType,qty:row.qty,unitCost,cost,revenue,earning:revenue-cost,orderId:row.orderId||"",buyer:row.buyer||"",city:row.city||"",state:row.state||""};
    });

    setSkus(newSkus);setReports([report,...reports]);setSalesLines([...salesLines,...newLines]);
    logActivity?.("Sales report applied",`[${channel}] "${weekLabel.trim()}" — ${fileName} (${aggregated.length} codes, ${newLines.length} order lines${skippedCount?`, ${skippedCount} duplicate orders skipped`:""})`);
    setStage("applied");showToast("success",skippedCount?`Inventory updated — ${skippedCount} duplicate order(s) skipped. 📊`:"Inventory updated. 📊");
  }
  function reset(){setStage("idle");setRawRows([]);setFileName("");setWeekLabel("");setRepairNote(null);setSkipDuplicates(true);if(ref.current)ref.current.value="";}

  function downloadWebsiteTemplate(){
    const csv="sku,quantity,item-price,purchase-date,buyer-name,buyer-email,ship-city,ship-state\n"+
      "ZB-ST-MBK-002,2,99,2026-07-05T10:30:00,Anita Sharma,anita@example.com,Mumbai,MAHARASHTRA\n"+
      "COMBO-A,1,499,2026-07-05T14:15:00,Rahul Verma,rahul@example.com,Delhi,DELHI\n";
    downloadCsv("website_sales_template.csv",csv);
  }
  function downloadAmazonReference(){
    const csv="amazon-order-id,sku,quantity,item-price,purchase-date,ship-city,ship-state\n"+
      "402-6139183-3052307,ZB-ST-MBK-002,1,99,2026-07-05T06:29:52+00:00,PUNE,MAHARASHTRA\n"+
      "171-7143323-7105914,OE-D495-NPMY,1,499,2026-07-04T18:56:02+00:00,PATNA,BIHAR\n";
    downloadCsv("amazon_sales_reference.csv",csv);
  }

  const hasExtras=extraCols.priceKey||extraCols.buyerKey||extraCols.cityKey||extraCols.orderIdKey;
  return(
    <div>
      <SectionHeader title="Upload Sales Report" subtitle="Amazon reports upload exactly as exported — combos auto-deduct component SKUs."/>

      {/* Channel selector */}
      <div className="flex gap-2 mb-3">
        {[{id:"amazon",label:"Amazon Seller Report"},{id:"website",label:"Website Sales"}].map(c=>(
          <button key={c.id} onClick={()=>{setChannel(c.id);reset();}} className="flex-1 sm:flex-none px-4 py-2.5 rounded-full text-sm font-bold transition-colors" style={{backgroundColor:channel===c.id?C.zenkyPurple:C.softWhite,color:channel===c.id?C.softWhite:C.darkText,border:`2px solid ${channel===c.id?C.zenkyPurple:C.border}`,fontFamily:F.display}}>
            {c.label}
          </button>
        ))}
      </div>

      {/* Entry mode: bulk file upload vs logging one order at a time */}
      <div className="flex gap-1.5 mb-5">
        {[{id:"file",label:"Upload File"},{id:"manual",label:"Add Single Order"}].map(m=>(
          <button key={m.id} onClick={()=>setEntryMode(m.id)} className="px-3.5 py-1.5 rounded-full text-xs font-bold transition-colors" style={{backgroundColor:entryMode===m.id?C.bgLight:"transparent",color:entryMode===m.id?C.zenkyPurple:C.lightText,border:`1.5px solid ${entryMode===m.id?C.zenkyPurple:C.border}`,fontFamily:F.body}}>
            {m.label}
          </button>
        ))}
      </div>

      {entryMode==="manual"?(
        <AddOrderForm channel={channel} skus={skus} combos={combos} setSkus={setSkus} reports={reports} setReports={setReports} salesLines={salesLines} setSalesLines={setSalesLines} logActivity={logActivity} showToast={showToast}/>
      ):(
      <Card>
        {stage==="idle"&&(
          <div>
            <div className="rounded-2xl border-2 border-dashed p-8 text-center" style={{borderColor:C.zenkyPink,backgroundColor:"#FFF8FC"}}>
              <Upload size={32} className="mx-auto mb-3" style={{color:C.zenkyPink}}/>
              <p className="font-bold mb-1" style={{color:C.darkText,fontFamily:F.display}}>
                {channel==="amazon"?"Choose your Amazon sales report — upload it exactly as downloaded":"Choose your website sales export"}
              </p>
              <p className="text-xs mb-4" style={{color:C.lightText}}>
                {channel==="amazon"?"No need to fix or reformat the file first — shifted/corrupted columns are detected and repaired automatically.":"CSV or Excel with sku, quantity, price, buyer, and location columns."}
              </p>
              <input ref={ref} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={e=>handleFile(e.target.files[0])}/>
              <PrimaryButton onClick={()=>ref.current?.click()}><Upload size={15}/>Select File</PrimaryButton>
              <div className="mt-4">
                <button onClick={channel==="amazon"?downloadAmazonReference:downloadWebsiteTemplate} className="inline-flex items-center gap-1.5 text-xs font-bold" style={{color:C.zenkyOrange,fontFamily:F.body}}>
                  <Download size={13}/>{channel==="amazon"?"Download format reference":"Download blank template"}
                </button>
              </div>
            </div>
          </div>
        )}
        {stage==="parsed"&&(<div><div className="flex items-center justify-between mb-4 flex-wrap gap-2"><span className="text-sm" style={{color:C.darkText}}>Parsed <strong>{fileName}</strong> — {aggregated.length} codes</span><button onClick={reset} className="text-sm font-bold" style={{color:C.lightText}}>Change file</button></div>
          {repairNote&&<div className="mb-3 p-2.5 rounded-xl text-xs flex items-center gap-2" style={{backgroundColor:repairNote.type==="success"?"#F0FDE8":"#FFF3E6",color:repairNote.type==="success"?"#166534":"#9a5b0f"}}><Check size={14}/>{repairNote.msg}</div>}
          {hasExtras&&<div className="mb-3 p-2.5 rounded-xl text-xs flex items-center gap-2" style={{backgroundColor:"#F0FDE8",color:"#166534"}}><Check size={14}/>Detected {[extraCols.priceKey&&"price",extraCols.buyerKey&&"buyer",extraCols.cityKey&&"location",extraCols.dateKey&&"date",extraCols.orderIdKey&&"order ID"].filter(Boolean).join(", ")} — will populate ZenkyBox Sales Report.</div>}
          {!hasExtras&&<div className="mb-3 p-2.5 rounded-xl text-xs flex items-center gap-2" style={{backgroundColor:C.bgLight,color:C.lightText}}><AlertTriangle size={14}/>No price/buyer/location columns detected — stock will update, but Sales Report earning/buyer/location breakdowns won't have data for this upload.</div>}
          {duplicateOrderIds.size>0&&(
            <div className="mb-3 p-3 rounded-xl text-xs" style={{backgroundColor:"#fff5f5",border:"2px solid #fecaca"}}>
              <div className="flex items-start gap-2">
                <AlertTriangle size={16} style={{color:"#dc2626",flexShrink:0,marginTop:1}}/>
                <div className="flex-1">
                  <p className="font-bold mb-1" style={{color:"#991b1b"}}>
                    {duplicateOrderIds.size} order{duplicateOrderIds.size!==1?"s":""} already recorded from a previous upload
                  </p>
                  <p style={{color:"#991b1b"}}>
                    These order IDs exist in an earlier applied report — likely the same sales exported through a different report format. Counting them again would double-deduct stock and inflate Sales Report totals.
                  </p>
                  <div className="mt-2 max-h-20 overflow-y-auto text-xs" style={{fontFamily:F.mono,color:"#991b1b"}}>
                    {Array.from(duplicateOrderIds).slice(0,8).join(", ")}{duplicateOrderIds.size>8?` …and ${duplicateOrderIds.size-8} more`:""}
                  </div>
                  <label className="flex items-center gap-2 mt-2.5 cursor-pointer">
                    <input type="checkbox" checked={skipDuplicates} onChange={e=>setSkipDuplicates(e.target.checked)}/>
                    <span className="font-bold" style={{color:"#991b1b"}}>Skip these duplicate orders when applying (recommended)</span>
                  </label>
                </div>
              </div>
            </div>
          )}
          <div className="text-xs mb-2" style={{color:C.lightText}}>
            {skipDuplicates&&duplicateOrderIds.size>0?`Showing ${aggregated.length} codes from ${effectiveRows.length} order lines (${duplicateOrderIds.size} duplicate orders excluded)`:`${aggregated.length} codes from ${effectiveRows.length} order lines`}
          </div>
          <div className="overflow-x-auto mb-4"><table className="w-full text-sm"><thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Code</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Qty</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Type</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Name</th></tr></thead><tbody>{aggregated.map(a=><tr key={a.code} className="border-t" style={{borderColor:C.border}}><td className="py-2 pr-3" style={{fontFamily:F.mono,fontWeight:600}}>{a.code}</td><td className="py-2 pr-3" style={{fontFamily:F.mono}}>{fmt(a.qty)}</td><td className="py-2 pr-3">{a.matchType==="combo"?<Stamp tone="mint">Combo</Stamp>:a.matchType==="direct"?<Stamp tone="purple">SKU</Stamp>:<Stamp tone="pink">Unknown</Stamp>}</td><td className="py-2 pr-3" style={{color:a.matchType==="unknown"?C.zenkyPink:C.darkText}}>{a.matchName}</td></tr>)}</tbody></table></div>
          <div className="flex items-end gap-3 flex-wrap"><div className="w-full sm:w-64"><label className="text-xs font-bold block mb-1" style={{color:C.lightText}}>Report label</label><Input placeholder="e.g. Week of Jun 16–22" value={weekLabel} onChange={e=>setWeekLabel(e.target.value)}/></div><PrimaryButton onClick={applyReport}><Check size={15}/>Apply to inventory</PrimaryButton></div></div>)}
        {stage==="applied"&&(<div className="text-center py-8"><Check size={32} className="mx-auto mb-3" style={{color:C.mintGreen}}/><p className="font-bold text-lg" style={{color:C.darkText,fontFamily:F.display}}>Report applied</p><p className="text-sm mt-1" style={{color:C.lightText}}>View breakdown in Reports tab, or full analytics in ZenkyBox Sales Report.</p><div className="mt-5"><PrimaryButton onClick={reset}><Upload size={15}/>Upload another</PrimaryButton></div></div>)}
      </Card>
      )}
    </div>
  );
}

/* ═══ MANUAL SINGLE-ORDER ENTRY (for website orders without a bulk export) ═══ */
function AddOrderForm({channel,skus,combos,setSkus,reports,setReports,salesLines,setSalesLines,logActivity,showToast}){
  const blank={code:"",qty:1,price:"",orderId:"",buyerName:"",buyerEmail:"",city:"",state:"",date:new Date().toISOString().slice(0,10)};
  const [form,setForm]=useState(blank);
  const skuMap=useMemo(()=>Object.fromEntries(skus.map(s=>[s.sku,s])),[skus]);
  const comboMap=useMemo(()=>Object.fromEntries(combos.map(c=>[c.sku,c])),[combos]);
  const options=useMemo(()=>[
    ...skus.map(s=>({code:s.sku,label:`${s.sku} — ${s.name}`,type:"SKU"})),
    ...combos.map(c=>({code:c.sku,label:`${c.sku} — ${c.name}`,type:"Combo"})),
  ],[skus,combos]);

  function submitOrder(){
    const code=form.code.trim();
    const qty=Number(form.qty)||0;
    if(!code){showToast("error","Choose a SKU or combo.");return;}
    if(qty<=0){showToast("error","Quantity must be greater than 0.");return;}
    const combo=comboMap[code],sku=skuMap[code];
    if(!combo&&!sku){showToast("error",`"${code}" not found in Catalog or Combos.`);return;}
    const matchType=combo?"combo":"direct";
    const name=combo?.name||sku?.name||code;
    const orderId=form.orderId.trim()||`WEB-${Date.now()}`;

    // Deduct stock — same logic as bulk apply: combos reduce every component proportionally
    const skuBefore={};skus.forEach(s=>skuBefore[s.sku]=s.stock);
    const updated=Object.fromEntries(skus.map(s=>[s.sku,{...s}]));
    if(matchType==="combo"){combo.components?.forEach(comp=>{if(updated[comp.sku])updated[comp.sku].stock=Math.max(0,updated[comp.sku].stock-comp.qty*qty);});}
    else if(updated[code]){updated[code].stock=Math.max(0,updated[code].stock-qty);}
    const newSkus=Object.values(updated);
    const bMap=Object.fromEntries(skus.map(s=>[s.sku,s])),aMap=Object.fromEntries(newSkus.map(s=>[s.sku,s]));

    const unitCost=unitCostOf(code,matchType,skuMap,comboMap);
    const cost=unitCost*qty;
    const price=form.price!==""?Number(form.price):null;
    const revenue=price!=null&&!isNaN(price)?price*qty:0;
    const dateIso=form.date?new Date(form.date).toISOString():new Date().toISOString();

    const reportId=Date.now().toString();
    const report={id:reportId,label:`Website Order — ${orderId}`,fileName:"(manual entry)",channel,appliedAt:new Date().toISOString(),
      skuLines:newSkus.map(s=>({sku:s.sku,name:s.name,opening:skuBefore[s.sku]??s.stock,sold:(skuBefore[s.sku]??s.stock)-s.stock,closing:s.stock,reorderLevel:s.reorderLevel,status:stockStatus(s)})),
      comboLines:combos.map(c=>({sku:c.sku,name:c.name,readyBefore:comboReadiness(c,bMap).ready,readyAfter:comboReadiness(c,aMap).ready,bottleneck:comboReadiness(c,aMap).bottleneck})),
      unmatched:[],skippedDuplicates:0};

    const newLine={id:`${reportId}-0`,reportId,channel,date:dateIso,sku:code,name,matchType,qty,unitCost,cost,revenue,earning:revenue-cost,orderId,buyer:form.buyerName.trim(),city:form.city.trim(),state:form.state.trim()};

    setSkus(newSkus);setReports([report,...reports]);setSalesLines([...salesLines,newLine]);
    logActivity?.("Order added manually",`[${channel}] ${orderId} — ${code} ×${qty}`);
    showToast("success",`Order added — ${name} ×${qty}. 📝`);
    setForm({...blank,date:form.date}); // keep the date for quick consecutive entries
  }

  return(
    <Card>
      <h3 className="font-bold text-lg mb-1" style={{fontFamily:F.display,color:C.darkText}}>Add a Single Order</h3>
      <p className="text-xs mb-4" style={{color:C.lightText}}>For {channel==="website"?"website":"Amazon"} orders you don't have a bulk file for — logs the sale, deducts stock, and records it in Sales Report just like an uploaded file would.</p>
      <div className="space-y-3">
        <div>
          <label className="text-xs font-bold uppercase block mb-1" style={{color:C.lightText}}>SKU or Combo</label>
          <Select value={form.code} onChange={e=>setForm({...form,code:e.target.value})}>
            <option value="">Select a SKU or combo…</option>
            {options.map(o=><option key={o.code} value={o.code}>{o.type}: {o.label}</option>)}
          </Select>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
          <Input placeholder="Quantity" type="number" min="1" value={form.qty} onChange={e=>setForm({...form,qty:e.target.value})}/>
          <div className="relative"><span className="absolute left-3 top-2.5 text-sm" style={{color:C.lightText}}>₹</span><Input placeholder="Sale price" type="number" className="pl-6" value={form.price} onChange={e=>setForm({...form,price:e.target.value})}/></div>
          <Input type="date" value={form.date} onChange={e=>setForm({...form,date:e.target.value})}/>
          <Input placeholder="Order ID (optional)" value={form.orderId} onChange={e=>setForm({...form,orderId:e.target.value})}/>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
          <Input placeholder="Buyer name (optional)" value={form.buyerName} onChange={e=>setForm({...form,buyerName:e.target.value})}/>
          <Input placeholder="Buyer email (optional)" value={form.buyerEmail} onChange={e=>setForm({...form,buyerEmail:e.target.value})}/>
        </div>
        <div className="grid grid-cols-2 gap-2">
          <Input placeholder="City (optional)" value={form.city} onChange={e=>setForm({...form,city:e.target.value})}/>
          <Input placeholder="State (optional)" value={form.state} onChange={e=>setForm({...form,state:e.target.value})}/>
        </div>
      </div>
      <div className="mt-4"><PrimaryButton onClick={submitOrder}><Plus size={16}/>Add Order</PrimaryButton></div>
    </Card>
  );
}

/* ═══ REPORTS (Enhanced with full export) ═══ */
function ReportsView({reports,skus,combos}){
  const [openId,setOpenId]=useState(reports[0]?.id??null);
  const skuMap=useMemo(()=>Object.fromEntries(skus.map(s=>[s.sku,s])),[skus]);

  function exportFullInventory(){
    let csv="=== SKU MASTER ===\nSKU,Name,Stock,Reorder Level,Procurement Cost (INR),Status,Suggested Reorder\n";
    skus.forEach(s=>{csv+=`${s.sku},"${s.name}",${s.stock},${s.reorderLevel},${s.procurementCost||0},${stockStatus(s)},${suggestedReorder(s)}\n`;});
    csv+="\n=== COMBO READINESS ===\nCombo Code,Combo Name,Can Make,Status,Bottleneck SKU\n";
    combos.forEach(c=>{const{ready,bottleneck}=comboReadiness(c,skuMap);csv+=`${c.sku},"${c.name}",${ready},${ready>0?"Ready":"Short"},${bottleneck||""}\n`;});
    csv+="\n=== COMBO COMPONENTS ===\nCombo Code,Combo Name,Component SKU,Component Name,Need,In Stock,Status\n";
    combos.forEach(c=>{c.components?.forEach(comp=>{const s=skuMap[comp.sku];csv+=`${c.sku},"${c.name}",${comp.sku},"${s?.name||""}",${comp.qty},${s?.stock||0},${s?stockStatus(s):"missing"}\n`;});});
    downloadCsv("zenkybox_full_inventory.csv",csv);
  }

  function exportReport(r){
    let csv="SKU,Name,Opening,Sold,Closing,Reorder Level,Status\n";
    r.skuLines.forEach(l=>{csv+=`${l.sku},"${l.name}",${l.opening},${l.sold},${l.closing},${l.reorderLevel},${l.status}\n`;});
    csv+="\nCombo Code,Name,Before,After,Bottleneck\n";
    r.comboLines.forEach(l=>{csv+=`${l.sku},"${l.name}",${l.readyBefore},${l.readyAfter},${l.bottleneck||""}\n`;});
    downloadCsv(`${r.label.replace(/\s+/g,"_")}_report.csv`,csv);
  }

  return(
    <div>
      <SectionHeader title="Reports & Exports" subtitle="Inventory snapshots and full data exports."
        action={<PrimaryButton onClick={exportFullInventory} tone="orange"><Download size={15}/>Export Full Inventory</PrimaryButton>}
      />

      {/* Full inventory quick export cards */}
      <div className="grid sm:grid-cols-3 gap-3 mb-6">
        {[{label:"Total SKUs",value:skus.length,sub:"in catalog",color:C.zenkyPurple},
          {label:"Combos",value:combos.length,sub:"gift bundles",color:C.zenkyPink},
          {label:"SKUs at Risk",value:skus.filter(s=>stockStatus(s)!=="healthy").length,sub:"need restock",color:C.zenkyOrange}].map(s=>(
          <Card key={s.label}><div className="text-xs font-bold uppercase" style={{color:C.lightText,fontFamily:F.body}}>{s.label}</div><div className="text-3xl font-black mt-1" style={{fontFamily:F.display,color:s.color}}>{s.value}</div><div className="text-xs" style={{color:C.lightText}}>{s.sub}</div></Card>
        ))}
      </div>

      {reports.length===0?<Empty icon={FileText} title="No reports yet" message="Upload and apply a sales report to generate snapshots."/>:(
        <div className="space-y-3">
          {reports.map(r=>{
            const isOpen=openId===r.id;
            const lowCount=r.skuLines?.filter(l=>l.status!=="healthy").length||0;
            const shortCombos=r.comboLines?.filter(l=>l.readyAfter<=0).length||0;
            return(
              <Card key={r.id}>
                <button className="w-full flex items-center justify-between gap-3 text-left" onClick={()=>setOpenId(isOpen?null:r.id)}>
                  <div className="flex items-center gap-2">{isOpen?<ChevronDown size={16}/>:<ChevronRight size={16}/>}<div><div className="font-bold text-sm flex items-center gap-2" style={{color:C.darkText,fontFamily:F.display}}>{r.label}{r.channel&&<Stamp tone={r.channel==="amazon"?"orange":"blue"}>{r.channel}</Stamp>}</div><div className="text-xs" style={{color:C.lightText,fontFamily:F.mono}}>Applied {new Date(r.appliedAt).toLocaleDateString()} · {r.fileName}</div></div></div>
                  <div className="flex items-center gap-2 flex-wrap justify-end">{lowCount>0&&<Stamp tone="orange">{lowCount} low</Stamp>}{shortCombos>0&&<Stamp tone="pink">{shortCombos} short</Stamp>}{r.unmatched?.length>0&&<Stamp tone="purple">{r.unmatched.length} unmatched</Stamp>}</div>
                </button>
                {isOpen&&(
                  <div className="mt-4 pt-4 border-t" style={{borderColor:C.border}}>
                    <div className="flex items-center justify-between mb-3"><h4 className="font-bold text-sm" style={{color:C.darkText,fontFamily:F.display}}>SKU Breakdown</h4><button onClick={()=>exportReport(r)} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export CSV</button></div>
                    <div className="overflow-x-auto mb-4"><table className="w-full text-sm"><thead><tr style={{color:C.lightText}}><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">SKU</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Name</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Opening</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Sold</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Closing</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Status</th></tr></thead><tbody>{r.skuLines?.map(l=><tr key={l.sku} className="border-t" style={{borderColor:C.border}}><td className="py-1.5 pr-3" style={{fontFamily:F.mono,fontWeight:600}}>{l.sku}</td><td className="py-1.5 pr-3">{l.name}</td><td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmt(l.opening)}</td><td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmt(l.sold)}</td><td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmt(l.closing)}</td><td className="py-1.5 pr-3">{statusStamp(l.status)}</td></tr>)}</tbody></table></div>
                    {r.unmatched?.length>0&&<div className="mt-3 text-xs flex items-start gap-2" style={{color:C.zenkyPink}}><AlertTriangle size={14} className="mt-0.5"/><span>{r.unmatched.length} unmatched: {r.unmatched.map(u=>u.code).join(", ")}</span></div>}
                    {r.skippedDuplicates>0&&<div className="mt-2 text-xs flex items-center gap-2" style={{color:"#166534"}}><Check size={13}/><span>{r.skippedDuplicates} duplicate order{r.skippedDuplicates!==1?"s":""} skipped — already recorded in an earlier report.</span></div>}
                  </div>
                )}
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ═══ COSTING & PRICING ═══ */
function CostingPricingView({skus}){
  const [selectedSku,setSelectedSku]=useState("");
  const [costs,setCosts]=useState({packaging:0,closingFee:0,cob:0,shipping:0,cogs:0,misc:0});
  const [margin,setMargin]=useState(30);

  const sku=useMemo(()=>skus.find(s=>s.sku===selectedSku),[skus,selectedSku]);
  const procurement=Number(sku?.procurementCost||0);

  const totalCost=useMemo(()=>procurement+Number(costs.packaging)+Number(costs.closingFee)+Number(costs.cob)+Number(costs.shipping)+Number(costs.cogs)+Number(costs.misc),[procurement,costs]);

  // Selling price = Total Cost / (1 - margin%)  [margin on selling price basis]
  const sellingPrice=useMemo(()=>margin>=100?0:totalCost/(1-margin/100),[totalCost,margin]);

  // MRP = Selling Price / (1 - discount%) — shows what MRP to set to offer X% discount from MRP
  const mrpOptions=useMemo(()=>MRP_DISCOUNTS.map(d=>({discount:d,mrp:d>=100?0:sellingPrice/(1-d/100)})),[sellingPrice]);

  function exportPricingCsv(){
    if(!sku){return;}
    let csv=`Costing & Pricing Report — ${sku.name} (${sku.sku})\n\n`;
    csv+=`COST BREAKDOWN\nProcurement Cost,${procurement}\nPackaging,${costs.packaging}\nClosing Fee,${costs.closingFee}\nCOB (Returns+Ads),${costs.cob}\nShipping,${costs.shipping}\nCOGS,${costs.cogs}\nMiscellaneous,${costs.misc}\nTotal Cost,${totalCost.toFixed(2)}\n\n`;
    csv+=`SELLING PRICE\nMargin %,${margin}%\nSelling Price,${sellingPrice.toFixed(2)}\n\n`;
    csv+=`MRP OPTIONS\nDiscount Offered,MRP\n`;
    mrpOptions.forEach(o=>csv+=`${o.discount}%,${o.mrp.toFixed(2)}\n`);
    downloadCsv(`pricing_${sku.sku}.csv`,csv);
  }

  const costFields=[{key:"packaging",label:"Packaging"},{key:"closingFee",label:"Closing Fee"},{key:"cob",label:"COB (Returns + Ads)"},{key:"shipping",label:"Shipping"},{key:"cogs",label:"COGS"},{key:"misc",label:"Miscellaneous"}];

  return(
    <div>
      <SectionHeader title="Costing & Pricing" subtitle="Calculate selling price and MRP discount tiers for any SKU."/>

      {/* SKU selector */}
      <Card className="mb-5">
        <label className="text-xs font-bold uppercase block mb-2" style={{color:C.lightText,fontFamily:F.body,letterSpacing:"0.02em"}}>Select SKU</label>
        <Select value={selectedSku} onChange={e=>setSelectedSku(e.target.value)}>
          <option value="">Choose a SKU to calculate pricing…</option>
          {skus.map(s=><option key={s.sku} value={s.sku}>{s.sku} — {s.name} {s.procurementCost?`(₹${s.procurementCost})`:""}</option>)}
        </Select>
        {sku&&<p className="text-xs mt-2" style={{color:C.lightText}}>Procurement cost pulled from catalog: <strong style={{color:C.zenkyPurple}}>₹{procurement}</strong></p>}
      </Card>

      {!selectedSku?<Empty icon={Calculator} title="Select a SKU" message="Choose a SKU above to calculate its costing and pricing."/>:(
        <div className="grid md:grid-cols-2 gap-5">

          {/* Cost breakdown */}
          <div className="space-y-4">
            <Card>
              <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>💰 Cost Breakdown</h3>
              <div className="space-y-3">
                <CostInput label="Procurement Cost" value={procurement} readOnly prefix="₹"/>
                {costFields.map(f=><CostInput key={f.key} label={f.label} value={costs[f.key]} onChange={e=>setCosts({...costs,[f.key]:e.target.value})} prefix="₹"/>)}
              </div>
              <div className="mt-4 pt-4 border-t flex items-center justify-between" style={{borderColor:C.border}}>
                <span className="font-bold text-sm" style={{fontFamily:F.display,color:C.darkText}}>Total Cost</span>
                <span className="text-2xl font-black" style={{fontFamily:F.display,color:C.zenkyPurple}}>{fmtINR(totalCost)}</span>
              </div>
            </Card>

            {/* Margin & selling price */}
            <Card>
              <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>📈 Selling Price</h3>
              <div className="mb-4">
                <label className="text-xs font-bold uppercase block mb-2" style={{color:C.lightText,fontFamily:F.body,letterSpacing:"0.02em"}}>Margin</label>
                <div className="grid grid-cols-3 gap-2">
                  {MARGIN_OPTIONS.map(m=>(
                    <button key={m} onClick={()=>setMargin(m)} className="py-2 rounded-xl border-2 text-sm font-bold transition-all" style={{borderColor:margin===m?C.zenkyPurple:C.border,backgroundColor:margin===m?C.zenkyPurple:"transparent",color:margin===m?C.softWhite:C.darkText,fontFamily:F.display}}>
                      {m}%
                    </button>
                  ))}
                </div>
              </div>
              <div className="rounded-xl p-4 text-center" style={{backgroundColor:C.bgLight}}>
                <div className="text-xs font-bold uppercase mb-1" style={{color:C.lightText,fontFamily:F.body}}>Selling Price ({margin}% margin)</div>
                <div className="text-4xl font-black" style={{fontFamily:F.display,color:C.zenkyPurple}}>{fmtINR(sellingPrice)}</div>
                <div className="text-xs mt-1" style={{color:C.lightText}}>= Total Cost ÷ (1 − {margin}%)</div>
              </div>
            </Card>
          </div>

          {/* MRP tiers */}
          <Card>
            <div className="flex items-center justify-between mb-4">
              <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>🏷️ MRP Discount Tiers</h3>
              <button onClick={exportPricingCsv} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange,fontFamily:F.body}}><Download size={13}/>Export</button>
            </div>
            <p className="text-xs mb-4" style={{color:C.lightText,fontFamily:F.body}}>Set your MRP to offer a specific discount to customers. Selling Price remains {fmtINR(sellingPrice)}.</p>
            <div className="space-y-2">
              {mrpOptions.map(o=>(
                <div key={o.discount} className="flex items-center justify-between p-3 rounded-xl border-2" style={{borderColor:C.border,backgroundColor:C.bgLight}}>
                  <div>
                    <div className="text-sm font-bold" style={{color:C.darkText,fontFamily:F.display}}>{o.discount}% discount to customer</div>
                    <div className="text-xs" style={{color:C.lightText,fontFamily:F.mono}}>MRP = {fmtINR(sellingPrice)} ÷ (1−{o.discount}%)</div>
                  </div>
                  <div className="text-right">
                    <div className="text-xl font-black" style={{fontFamily:F.display,color:C.zenkyPink}}>{fmtINR(o.mrp)}</div>
                    <div className="text-xs" style={{color:C.lightText}}>MRP</div>
                  </div>
                </div>
              ))}
            </div>
            <div className="mt-4 p-3 rounded-xl text-xs" style={{backgroundColor:"#FFF8FC",color:C.lightText,fontFamily:F.body}}>
              💡 <strong>Tip:</strong> Amazon India typically expects MRP printed on packaging. Choose an MRP that allows you to show meaningful discount while maintaining your target selling price.
            </div>
          </Card>
        </div>
      )}
    </div>
  );
}

/* ═══ ZENKYBOX SALES REPORT ═══ */
function SalesReportsView({salesLines,skus,combos}){
  const [tab,setTab]=useState("overall");
  const [channelFilter,setChannelFilter]=useState("all"); // "all" | "amazon" | "website"
  const skuMap=useMemo(()=>Object.fromEntries(skus.map(s=>[s.sku,s])),[skus]);
  const comboMap=useMemo(()=>Object.fromEntries(combos.map(c=>[c.sku,c])),[combos]);

  // Apply the channel filter before anything else touches sales data — every
  // tab below (including Overall/MoM/Weekly/Buyer/Location) automatically
  // respects whichever channel is selected.
  const filteredLines=useMemo(()=>{
    if(channelFilter==="all")return salesLines;
    return salesLines.filter(l=>(l.channel||"amazon")===channelFilter);
  },[salesLines,channelFilter]);

  const TABS=[
    {id:"overall",label:"Overall"},
    {id:"channel",label:"Channel-wise"},
    {id:"mom",label:"Month-over-Month"},
    {id:"weekly",label:"Weekly"},
    {id:"buyer",label:"Buyer-wise"},
    {id:"location",label:"Location-wise"},
  ];

  const totals=useMemo(()=>filteredLines.reduce((a,l)=>({qty:a.qty+l.qty,revenue:a.revenue+l.revenue,cost:a.cost+l.cost,earning:a.earning+l.earning}),{qty:0,revenue:0,cost:0,earning:0}),[filteredLines]);

  // Per-SKU / per-combo breakdown (used in Overall, MoM, Weekly — each just changes the grouping level)
  function skuComboBreakdown(lines){
    const bySku={},byCombo={};
    lines.forEach(l=>{
      const bucket=l.matchType==="combo"?byCombo:bySku;
      if(!bucket[l.sku])bucket[l.sku]={code:l.sku,name:l.name,qty:0,revenue:0,cost:0,earning:0};
      bucket[l.sku].qty+=l.qty;bucket[l.sku].revenue+=l.revenue;bucket[l.sku].cost+=l.cost;bucket[l.sku].earning+=l.earning;
    });
    return{skuRows:Object.values(bySku).sort((a,b)=>b.revenue-a.revenue),comboRows:Object.values(byCombo).sort((a,b)=>b.revenue-a.revenue)};
  }

  function exportBreakdown(title,skuRows,comboRows){
    let csv=`${title}\n\nSKU DETAILS\nSKU,Name,Qty Sold,Cost,Revenue,Earning\n`;
    skuRows.forEach(r=>csv+=`${r.code},"${r.name}",${r.qty},${r.cost.toFixed(2)},${r.revenue.toFixed(2)},${r.earning.toFixed(2)}\n`);
    csv+=`\nCOMBO DETAILS\nCombo Code,Name,Qty Sold,Cost,Revenue,Earning\n`;
    comboRows.forEach(r=>csv+=`${r.code},"${r.name}",${r.qty},${r.cost.toFixed(2)},${r.revenue.toFixed(2)},${r.earning.toFixed(2)}\n`);
    downloadCsv(`${title.replace(/\s+/g,"_")}.csv`,csv);
  }

  function BreakdownTable({rows,label}){
    if(!rows.length)return null;
    return(
      <div className="mb-5">
        <h4 className="text-sm font-bold mb-2" style={{color:C.darkText,fontFamily:F.display}}>{label} ({rows.length})</h4>
        <div className="overflow-x-auto"><table className="w-full text-sm">
          <thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Code</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Name</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Qty</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Cost</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Revenue</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Earning</th></tr></thead>
          <tbody>{rows.map(r=>(
            <tr key={r.code} className="border-t" style={{borderColor:C.border}}>
              <td className="py-2 pr-3" style={{fontFamily:F.mono,fontWeight:600}}>{r.code}</td>
              <td className="py-2 pr-3">{r.name}</td>
              <td className="py-2 pr-3" style={{fontFamily:F.mono}}>{fmt(r.qty)}</td>
              <td className="py-2 pr-3" style={{fontFamily:F.mono}}>{fmtINR(r.cost)}</td>
              <td className="py-2 pr-3" style={{fontFamily:F.mono}}>{fmtINR(r.revenue)}</td>
              <td className="py-2 pr-3 font-bold" style={{fontFamily:F.mono,color:r.earning>=0?C.mintGreen:C.zenkyPink}}>{fmtINR(r.earning)}</td>
            </tr>
          ))}</tbody>
        </table></div>
      </div>
    );
  }

  function GroupedReport({groupFn,emptyMsg,chrono}){
    const groups=useMemo(()=>aggregateSalesLines(filteredLines,groupFn,skuMap,comboMap,chrono?"chrono":"revenue"),[filteredLines,skuMap,comboMap,chrono]);
    if(!groups.length)return<Empty icon={BarChart3} title="No sales data yet" message={emptyMsg}/>;
    return(
      <div>
        {chrono&&groups.length>=2&&(
          <Card className="mb-4">
            <h4 className="text-sm font-bold mb-3" style={{color:C.darkText,fontFamily:F.display}}>Trend</h4>
            <TrendChart groups={groups}/>
          </Card>
        )}
        <div className="space-y-4">
        {groups.map(g=>{
          const{skuRows,comboRows}=skuComboBreakdown(g.lines);
          return(
            <Card key={g.key}>
              <div className="flex items-center justify-between gap-3 mb-3 flex-wrap">
                <div>
                  <div className="font-bold text-lg" style={{fontFamily:F.display,color:C.zenkyPurple}}>{g.key}</div>
                  <div className="text-xs" style={{color:C.lightText}}>{fmt(g.qty)} units · {g.lines.length} order line{g.lines.length!==1?"s":""}</div>
                </div>
                <div className="flex items-center gap-4 text-right">
                  <div><div className="text-xs" style={{color:C.lightText}}>Revenue</div><div className="font-bold" style={{fontFamily:F.mono,color:C.zenkyPurple}}>{fmtINR(g.revenue)}</div></div>
                  <div><div className="text-xs" style={{color:C.lightText}}>Cost</div><div className="font-bold" style={{fontFamily:F.mono,color:C.zenkyOrange}}>{fmtINR(g.cost)}</div></div>
                  <div><div className="text-xs" style={{color:C.lightText}}>Earning</div><div className="font-bold" style={{fontFamily:F.mono,color:g.earning>=0?C.mintGreen:C.zenkyPink}}>{fmtINR(g.earning)}</div></div>
                  <button onClick={()=>exportBreakdown(g.key,skuRows,comboRows)} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export</button>
                </div>
              </div>
              <BreakdownTable rows={skuRows} label="SKU Details"/>
              <BreakdownTable rows={comboRows} label="Combo Details"/>
            </Card>
          );
        })}
        </div>
      </div>
    );
  }

  const overallBreakdown=useMemo(()=>skuComboBreakdown(filteredLines),[filteredLines]);

  return(
    <div>
      <SectionHeader title="ZenkyBox Sales Report" subtitle="SKU & combo performance across every applied sales upload — cost, revenue, and earning."/>

      {/* Channel filter — applies to every tab below */}
      <div className="flex items-center gap-2 mb-4 flex-wrap">
        <span className="text-xs font-bold uppercase" style={{color:C.lightText,letterSpacing:"0.03em"}}>Channel:</span>
        {[{id:"all",label:"All Channels"},{id:"amazon",label:"Amazon"},{id:"website",label:"Website"}].map(c=>(
          <button key={c.id} onClick={()=>setChannelFilter(c.id)} className="px-3 py-1.5 rounded-full text-xs font-bold transition-colors" style={{backgroundColor:channelFilter===c.id?C.zenkyPurple:C.softWhite,color:channelFilter===c.id?C.softWhite:C.darkText,border:`1.5px solid ${channelFilter===c.id?C.zenkyPurple:C.border}`,fontFamily:F.display}}>
            {c.label}
          </button>
        ))}
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-6">
        {[{label:"Units Sold",value:fmt(totals.qty),color:C.zenkyPurple},
          {label:"Revenue",value:fmtINR(totals.revenue),color:C.zenkyPurple},
          {label:"Cost",value:fmtINR(totals.cost),color:C.zenkyOrange},
          {label:"Earning",value:fmtINR(totals.earning),color:totals.earning>=0?C.mintGreen:C.zenkyPink}].map(s=>(
          <Card key={s.label}><div className="text-xs font-bold uppercase" style={{color:C.lightText,fontFamily:F.body}}>{s.label}</div><div className="text-2xl font-black mt-1" style={{fontFamily:F.display,color:s.color}}>{s.value}</div></Card>
        ))}
      </div>

      {filteredLines.length===0&&(
        <div className="mb-5 p-3 rounded-xl text-xs flex items-start gap-2" style={{backgroundColor:C.bgLight,color:C.lightText}}>
          <AlertTriangle size={14} className="flex-shrink-0 mt-0.5"/>
          <span>No sales data{channelFilter!=="all"?` for ${channelFilter}`:""} yet. Revenue/Earning/Buyer/Location figures populate once you upload sales files containing price, buyer, and location columns (see Upload Sales tab). Stock-only uploads still work but won't appear here.</span>
        </div>
      )}

      {/* Sub-tabs */}
      <div className="flex gap-1.5 mb-5 overflow-x-auto pb-1">
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} className="px-3.5 py-2 rounded-full text-xs font-bold whitespace-nowrap transition-colors" style={{backgroundColor:tab===t.id?C.zenkyPurple:C.softWhite,color:tab===t.id?C.softWhite:C.darkText,border:`2px solid ${tab===t.id?C.zenkyPurple:C.border}`,fontFamily:F.display}}>
            {t.label}
          </button>
        ))}
      </div>

      {tab==="overall"&&(
        filteredLines.length===0?null:(
          <Card>
            <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
              <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>All-Time Overall Report</h3>
              <button onClick={()=>exportBreakdown("Overall_Sales_Report",overallBreakdown.skuRows,overallBreakdown.comboRows)} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export</button>
            </div>
            <BreakdownTable rows={overallBreakdown.skuRows} label="SKU Details"/>
            <BreakdownTable rows={overallBreakdown.comboRows} label="Combo Details"/>
          </Card>
        )
      )}
      {tab==="channel"&&<GroupedReport groupFn={l=>l.channel==="website"?"Website":l.channel==="amazon"?"Amazon":"Unlabeled (older upload)"} emptyMsg="Upload sales from Amazon or Website to compare channels here."/>}
      {tab==="mom"&&<GroupedReport groupFn={l=>monthKey(l.date)} emptyMsg="Upload sales with a date column to see month-over-month trends." chrono/>}
      {tab==="weekly"&&<GroupedReport groupFn={l=>weekKey(l.date)} emptyMsg="Upload sales with a date column to see weekly trends." chrono/>}
      {tab==="buyer"&&<GroupedReport groupFn={l=>l.buyer||"Unknown buyer"} emptyMsg="Upload sales with a buyer name/email column to see buyer-wise data."/>}
      {tab==="location"&&<GroupedReport groupFn={l=>[l.city,l.state].filter(Boolean).join(", ")||"Unknown location"} emptyMsg="Upload sales with ship-city/ship-state columns to see location-wise data."/>}
    </div>
  );
}

/* ═══ SOURCE DATA (Admin only) ═══ */
function SourceDataView({activityLog,synced,salesLines,setSalesLines,reports,setReports,skus,setSkus,combos,adminPin,loginCreds,investors,investments,expenses,income,forceSaveNow,logActivity,showToast}){
  const dbUrl=process.env.NEXT_PUBLIC_SUPABASE_URL||"";
  const [scanResult,setScanResult]=useState(null); // {dupeGroups, extraQtyBySku, linesToRemove}
  const [confirmText,setConfirmText]=useState("");
  const [showConfirm,setShowConfirm]=useState(false);

  // Full off-platform backup — a plain JSON file downloaded to this device.
  // This is the ONLY backup layer that survives a whole-project deletion in
  // Supabase, since it lives entirely outside the database. Everything else
  // (in-database snapshots via pg_cron) is destroyed along with the project
  // if the project itself is ever deleted.
  function downloadFullBackup(){
    const backup={
      exportedAt:new Date().toISOString(),
      skus,combos,reports,salesLines,activityLog,
      investors,investments,expenses,income,
      // adminPin/loginCreds deliberately excluded from the downloadable file —
      // don't want credentials sitting in a file that could end up anywhere.
    };
    const json=JSON.stringify(backup,null,2);
    const blob=new Blob([json],{type:"application/json"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    const dateStr=new Date().toISOString().slice(0,10);
    a.href=url;a.download=`zenkybox_backup_${dateStr}.json`;
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    URL.revokeObjectURL(url);
    logActivity?.("Downloaded full backup",`${skus.length} SKUs, ${combos.length} combos, ${salesLines.length} sale lines, ${investors.length} investors`);
    showToast("success","Backup downloaded. Keep this file somewhere safe outside Supabase. 💾");
    if(typeof window!=="undefined")localStorage.setItem("zenkybox-last-backup",new Date().toISOString());
  }

  const lastBackup=typeof window!=="undefined"?localStorage.getItem("zenkybox-last-backup"):null;
  const daysSinceBackup=lastBackup?Math.floor((Date.now()-new Date(lastBackup).getTime())/86400000):null;

  function exportLog(){
    let csv="Date,Action,Detail,Role\n";
    activityLog.forEach(a=>csv+=`${a.date},"${a.action}","${a.detail||""}",${a.role}\n`);
    downloadCsv("zenkybox_activity_log.csv",csv);
  }
  // Group log entries by day for a clean "day-wise" table
  const byDay=useMemo(()=>{
    const groups={};
    activityLog.forEach(a=>{
      const day=new Date(a.date).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
      if(!groups[day])groups[day]=[];
      groups[day].push(a);
    });
    return Object.entries(groups); // insertion order = newest first since log is prepended
  },[activityLog]);

  // Scan for sales lines that share the same order-id + sku — these are almost
  // certainly the same real-world sale counted more than once (e.g. the same
  // orders exported through two different Amazon report types and both uploaded).
  function scanForDuplicates(){
    const groups={};
    salesLines.forEach(line=>{
      if(!line.orderId)return; // can't dedupe without an order id — leave these alone
      const key=`${line.orderId}::${line.sku}`;
      if(!groups[key])groups[key]=[];
      groups[key].push(line);
    });
    const dupeGroups=Object.values(groups).filter(g=>g.length>1);
    const extraQtyBySku={};
    let linesToRemove=[];
    dupeGroups.forEach(group=>{
      const [keep,...extras]=group; // keep the first occurrence, remove the rest
      extras.forEach(l=>{
        extraQtyBySku[l.sku]=(extraQtyBySku[l.sku]||0)+l.qty;
        linesToRemove.push(l.id);
      });
    });
    setScanResult({dupeGroups,extraQtyBySku,linesToRemove,totalDuplicateLines:linesToRemove.length});
    setShowConfirm(false);setConfirmText("");
  }

  function applyCleanup(){
    if(confirmText.trim().toUpperCase()!=="CLEAN"){showToast("error",'Type "CLEAN" exactly to confirm.');return;}
    const removeSet=new Set(scanResult.linesToRemove);
    const newSalesLines=salesLines.filter(l=>!removeSet.has(l.id));
    const newSkus=skus.map(s=>scanResult.extraQtyBySku[s.sku]?{...s,stock:s.stock+scanResult.extraQtyBySku[s.sku]}:s);
    const newActivityLog=[{id:Date.now().toString()+Math.random(),date:new Date().toISOString(),action:"Cleaned duplicate sales data",detail:`Removed ${scanResult.totalDuplicateLines} duplicate order lines, restored stock for ${Object.keys(scanResult.extraQtyBySku).length} SKUs`,role:"admin"},...activityLog].slice(0,300);
    setSalesLines(newSalesLines);
    setSkus(newSkus);
    logActivity?.("Cleaned duplicate sales data",`Removed ${scanResult.totalDuplicateLines} duplicate order lines, restored stock for ${Object.keys(scanResult.extraQtyBySku).length} SKUs`);
    // Write the corrected state to Supabase immediately — don't wait for the
    // normal debounce, which leaves a window for a stale tab to resave old data.
    forceSaveNow?.({skus:newSkus,combos,reports,salesLines:newSalesLines,activityLog:newActivityLog,adminPin,loginCreds,investors,investments,expenses,income});
    showToast("success",`Removed ${scanResult.totalDuplicateLines} duplicate entries and restored stock. ✨`);
    setScanResult(null);setShowConfirm(false);setConfirmText("");
  }

  const [showFlushConfirm,setShowFlushConfirm]=useState(false);
  const [flushConfirmText,setFlushConfirmText]=useState("");

  function flushAllSalesData(){
    if(flushConfirmText.trim().toUpperCase()!=="FLUSH"){showToast("error",'Type "FLUSH" exactly to confirm.');return;}
    const lineCount=salesLines.length,reportCount=reports.length;
    const newSkus=skus.map(s=>({...s,stock:s.initialStock??s.stock}));
    const newActivityLog=[{id:Date.now().toString()+Math.random(),date:new Date().toISOString(),action:"Flushed all sales data",detail:`Cleared ${lineCount} sale lines and ${reportCount} reports; reset stock to initial baseline for all SKUs`,role:"admin"},...activityLog].slice(0,300);
    setSalesLines([]);
    setReports([]);
    // Reset every SKU's stock back to its recorded baseline, undoing every sales
    // deduction ever applied — the clean-slate option when duplicate/corrupted
    // uploads have made the current numbers untrustworthy.
    setSkus(newSkus);
    logActivity?.("Flushed all sales data",`Cleared ${lineCount} sale lines and ${reportCount} reports; reset stock to initial baseline for all SKUs`);
    // Write immediately — this is the fix for "flush didn't stick": waiting for
    // the normal 500ms debounce left a window where a stale open tab elsewhere
    // could resave its old (pre-flush) state and overwrite this. Forcing the
    // write right now, with the exact new values, closes that window.
    forceSaveNow?.({skus:newSkus,combos,reports:[],salesLines:[],activityLog:newActivityLog,adminPin,loginCreds,investors,investments,expenses,income});
    showToast("success","All sales data flushed and stock reset to baseline. 🗑️");
    setShowFlushConfirm(false);setFlushConfirmText("");
  }

  return(
    <div>
      <SectionHeader title="Source Data" subtitle="Database connection and a day-wise record of every change made to this workspace."/>

      <Card className="mb-6">
        <div className="flex items-center gap-2 mb-3">
          <Database size={18} style={{color:C.zenkyPurple}}/>
          <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>Database Connection</h3>
        </div>
        {dbUrl?(
          <div className="flex items-center gap-2 flex-wrap">
            <div className={`w-2 h-2 rounded-full ${synced?"bg-green-500":"bg-gray-400"}`}/>
            <code className="text-sm px-3 py-2 rounded-lg break-all" style={{backgroundColor:C.bgLight,fontFamily:F.mono,color:C.zenkyPurple}}>{dbUrl}</code>
          </div>
        ):(
          <div className="text-sm p-3 rounded-xl" style={{backgroundColor:"#FFF3E6",color:"#9a5b0f"}}>
            No database connected — data is stored in this browser's local storage only (not synced across devices). See SUPABASE_SETUP.md to connect a free cross-device database.
          </div>
        )}
      </Card>

      <Card className="mb-6" style={{borderColor:daysSinceBackup>7?"#fecaca":C.border}}>
        <div className="flex items-center gap-2 mb-3">
          <Download size={18} style={{color:C.zenkyPurple}}/>
          <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>Download Full Backup</h3>
        </div>
        <p className="text-sm mb-3" style={{color:C.darkText,fontFamily:F.body}}>
          Downloads everything — SKUs, combos, sales history, investors, expenses, income — as one JSON file to this device. This is the only backup that survives if the Supabase project itself is ever deleted, since it lives completely outside Supabase.
        </p>
        {lastBackup?(
          <p className="text-xs mb-3" style={{color:daysSinceBackup>7?"#dc2626":C.lightText,fontWeight:daysSinceBackup>7?700:400}}>
            Last backup: {daysSinceBackup===0?"today":`${daysSinceBackup} day${daysSinceBackup!==1?"s":""} ago`}
            {daysSinceBackup>7&&" — overdue, download a fresh one now"}
          </p>
        ):(
          <p className="text-xs mb-3" style={{color:"#dc2626",fontWeight:700}}>No backup has ever been downloaded on this device.</p>
        )}
        <PrimaryButton onClick={downloadFullBackup}><Download size={15}/>Download Backup Now</PrimaryButton>
        <p className="text-xs mt-3" style={{color:C.lightText}}>
          Recommended: download one every 7 days, and store it somewhere outside Supabase/Vercel entirely — email it to yourself, save to Google Drive, or a folder on your computer.
        </p>
      </Card>

      <Card className="mb-6">
        <div className="flex items-center gap-2 mb-3">
          <RefreshCw size={18} style={{color:C.zenkyPurple}}/>
          <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>Clean Duplicate Sales Data</h3>
        </div>
        <p className="text-sm mb-4" style={{color:C.lightText,fontFamily:F.body}}>
          Scans every recorded sale for repeated order-id + SKU combinations — the signature of the same real sale being uploaded more than once (e.g. via two different Amazon report types covering the same dates). Removes the extras and restores any stock that was wrongly deducted twice.
        </p>
        {!scanResult&&<PrimaryButton onClick={scanForDuplicates}><Search size={15}/>Scan for Duplicates</PrimaryButton>}

        {scanResult&&scanResult.totalDuplicateLines===0&&(
          <div className="flex items-center gap-2 p-3 rounded-xl text-sm" style={{backgroundColor:"#F0FDE8",color:"#166534"}}>
            <Check size={16}/>No duplicates found — your sales data looks clean.
          </div>
        )}

        {scanResult&&scanResult.totalDuplicateLines>0&&(
          <div>
            <div className="p-3.5 rounded-xl mb-4" style={{backgroundColor:"#fff5f5",border:"2px solid #fecaca"}}>
              <p className="font-bold text-sm mb-2" style={{color:"#991b1b",fontFamily:F.display}}>
                Found {scanResult.totalDuplicateLines} duplicate order line{scanResult.totalDuplicateLines!==1?"s":""} across {scanResult.dupeGroups.length} order{scanResult.dupeGroups.length!==1?"s":""}
              </p>
              <p className="text-xs mb-3" style={{color:"#991b1b"}}>Stock will be restored for these SKUs (the amount that was over-deducted):</p>
              <div className="space-y-1 mb-3">
                {Object.entries(scanResult.extraQtyBySku).map(([sku,qty])=>(
                  <div key={sku} className="text-xs flex justify-between" style={{fontFamily:F.mono,color:"#991b1b"}}>
                    <span>{sku}</span><span className="font-bold">+{qty} units</span>
                  </div>
                ))}
              </div>
              <p className="text-xs" style={{color:"#991b1b"}}>Note: this fixes current stock and Sales Report totals. Past entries in the Reports tab are historical snapshots and won't be rewritten.</p>
            </div>
            {!showConfirm?(
              <div className="flex gap-2">
                <PrimaryButton onClick={()=>setShowConfirm(true)} tone="pink"><Check size={15}/>Remove Duplicates & Restore Stock</PrimaryButton>
                <button onClick={()=>setScanResult(null)} className="text-sm font-bold" style={{color:C.lightText}}>Cancel</button>
              </div>
            ):(
              <div className="flex items-center gap-2 flex-wrap">
                <Input placeholder='Type "CLEAN" to confirm' value={confirmText} onChange={e=>setConfirmText(e.target.value)} className="max-w-xs" style={{borderColor:"#fecaca"}}/>
                <button onClick={applyCleanup} className="px-4 py-2.5 rounded-xl text-sm font-bold text-white" style={{backgroundColor:"#dc2626",fontFamily:F.display}}>Confirm Cleanup</button>
                <button onClick={()=>{setShowConfirm(false);setConfirmText("");}} className="text-sm font-bold" style={{color:C.lightText,fontFamily:F.body}}>Cancel</button>
              </div>
            )}
          </div>
        )}
      </Card>

      <Card className="mb-6" style={{borderColor:"#fecaca"}}>
        <div className="flex items-center gap-2 mb-3">
          <Trash2 size={18} style={{color:"#dc2626"}}/>
          <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:"#991b1b"}}>Flush All Sales Data (Clean Slate)</h3>
        </div>
        <p className="text-sm mb-1" style={{color:C.darkText,fontFamily:F.body}}>
          For when duplicate/corrupted uploads have made current numbers untrustworthy and you'd rather start fresh than fix them piece by piece. This action:
        </p>
        <ul className="text-sm mb-4 list-disc pl-5 space-y-0.5" style={{color:C.darkText,fontFamily:F.body}}>
          <li>Clears every sale line ({salesLines.length} currently) — Sales Report goes back to zero</li>
          <li>Clears every applied report ({reports.length} currently) — Reports tab goes back to zero</li>
          <li>Resets every SKU's stock back to its original recorded baseline — undoing every deduction ever applied</li>
          <li><strong>Keeps</strong> your SKU Catalog and Gift Combos exactly as they are — nothing about product definitions is touched</li>
        </ul>
        {!showFlushConfirm?(
          <button onClick={()=>setShowFlushConfirm(true)} className="inline-flex items-center gap-2 px-4 py-2.5 rounded-full text-sm font-bold text-white" style={{backgroundColor:"#dc2626",fontFamily:F.display}}>
            <Trash2 size={15}/>Flush All Sales Data
          </button>
        ):(
          <div className="p-3.5 rounded-xl" style={{backgroundColor:"#fff5f5",border:"2px solid #fecaca"}}>
            <p className="font-bold text-sm mb-2" style={{color:"#991b1b",fontFamily:F.display}}>This cannot be undone. Type "FLUSH" to confirm.</p>
            <div className="flex items-center gap-2 flex-wrap">
              <Input placeholder='Type "FLUSH" to confirm' value={flushConfirmText} onChange={e=>setFlushConfirmText(e.target.value)} className="max-w-xs" style={{borderColor:"#fecaca"}}/>
              <button onClick={flushAllSalesData} className="px-4 py-2.5 rounded-xl text-sm font-bold text-white" style={{backgroundColor:"#dc2626",fontFamily:F.display}}>Confirm Flush</button>
              <button onClick={()=>{setShowFlushConfirm(false);setFlushConfirmText("");}} className="text-sm font-bold" style={{color:C.lightText,fontFamily:F.body}}>Cancel</button>
            </div>
          </div>
        )}
      </Card>

      <Card>
        <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
          <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>Day-wise Update Record</h3>
          {activityLog.length>0&&<button onClick={exportLog} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export</button>}
        </div>
        {activityLog.length===0?<Empty icon={Calendar} title="No activity yet" message="Every SKU/combo change, import, and sales upload will be logged here."/>:(
          <div className="space-y-5">
            {byDay.map(([day,entries])=>(
              <div key={day}>
                <div className="text-xs font-bold uppercase mb-2" style={{color:C.lightText,letterSpacing:"0.03em"}}>{day}</div>
                <div className="overflow-x-auto"><table className="w-full text-sm">
                  <thead><tr style={{color:C.lightText}}><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase w-24">Time</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Action</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Detail</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase w-20">Role</th></tr></thead>
                  <tbody>{entries.map(a=>(
                    <tr key={a.id} className="border-t" style={{borderColor:C.border}}>
                      <td className="py-1.5 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{new Date(a.date).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"})}</td>
                      <td className="py-1.5 pr-3 font-bold" style={{color:C.darkText}}>{a.action}</td>
                      <td className="py-1.5 pr-3" style={{color:C.lightText}}>{a.detail}</td>
                      <td className="py-1.5 pr-3"><Stamp tone={a.role==="admin"?"purple":"blue"}>{a.role}</Stamp></td>
                    </tr>
                  ))}</tbody>
                </table></div>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

/* ═══ ACCESS MANAGEMENT (Admin only) ═══ */
function AccessManagementView({role,adminPin,setAdminPin,loginCreds,setLoginCreds,users,setUsers,showToast,logActivity}){
  const [newPin,setNewPin]=useState("");
  const [confirmPin,setConfirmPin]=useState("");
  const [newUsername,setNewUsername]=useState("");
  const [newPassword,setNewPassword]=useState("");
  const [confirmPassword,setConfirmPassword]=useState("");
  const [userForm,setUserForm]=useState({username:"",password:"",name:"",canBeAdmin:false});
  const [editUserId,setEditUserId]=useState(null);
  const [editUserValues,setEditUserValues]=useState({});
  const [deleteUserId,setDeleteUserId]=useState(null);

  function addUser(){
    const uname=userForm.username.trim();
    if(!uname){showToast("error","Username can't be empty.");return;}
    if(userForm.password.length<6){showToast("error","Password must be at least 6 characters.");return;}
    if(uname===(loginCreds?.username||DEFAULT_LOGIN.username)||users.some(u=>u.username===uname)){showToast("error",`Username "${uname}" is already in use.`);return;}
    const user={id:Date.now().toString(),username:uname,password:userForm.password,name:userForm.name.trim()||uname,canBeAdmin:userForm.canBeAdmin};
    setUsers([...users,user]);
    logActivity?.("User added",`${user.name} (${user.username})${user.canBeAdmin?" — can request admin":""}`);
    showToast("success",`Added user ${user.name}. ✨`);
    setUserForm({username:"",password:"",name:"",canBeAdmin:false});
  }
  function saveUserEdit(id){
    if(editUserValues.password&&editUserValues.password.length<6){showToast("error","Password must be at least 6 characters.");return;}
    setUsers(users.map(u=>u.id===id?{...u,name:editUserValues.name.trim()||u.username,canBeAdmin:editUserValues.canBeAdmin,...(editUserValues.password?{password:editUserValues.password}:{})}:u));
    logActivity?.("User edited",editUserValues.name||id);
    showToast("success","Saved. ✨");
    setEditUserId(null);
  }
  function removeUser(id){
    const u=users.find(x=>x.id===id);
    setUsers(users.filter(x=>x.id!==id));
    logActivity?.("User removed",u?.name||id);
    showToast("success","User removed.");
    setDeleteUserId(null);
  }

  function savePin(){
    if(newPin.length<4){showToast("error","PIN must be at least 4 characters.");return;}
    if(newPin!==confirmPin){showToast("error","PINs don't match.");return;}
    setAdminPin(newPin);setNewPin("");setConfirmPin("");
    logActivity?.("Admin PIN changed","Access PIN updated by admin");
    showToast("success","Admin PIN updated. ✨");
  }
  function saveLogin(){
    if(!newUsername.trim()){showToast("error","Username can't be empty.");return;}
    if(newPassword.length<6){showToast("error","Password must be at least 6 characters.");return;}
    if(newPassword!==confirmPassword){showToast("error","Passwords don't match.");return;}
    setLoginCreds({username:newUsername.trim(),password:newPassword});
    setNewUsername("");setNewPassword("");setConfirmPassword("");
    logActivity?.("Login credentials changed",`Username updated to "${newUsername.trim()}"`);
    showToast("success","Login credentials updated. Use them next time you sign in. ✨");
  }
  return(
    <div>
      <SectionHeader title="Access Management" subtitle="Control who can reach critical actions vs. everyday data entry."/>

      <Card className="mb-6">
        <h3 className="font-bold text-lg mb-3" style={{fontFamily:F.display,color:C.darkText}}>How access works</h3>
        <div className="grid sm:grid-cols-2 gap-4">
          <div className="p-4 rounded-xl" style={{backgroundColor:C.bgLight}}>
            <div className="flex items-center gap-2 mb-2"><Lock size={16} style={{color:C.zenkyPurple}}/><span className="font-bold text-sm" style={{fontFamily:F.display,color:C.zenkyPurple}}>Admin</span></div>
            <ul className="text-xs space-y-1" style={{color:C.darkText}}>
              <li>• Edit / delete SKUs & combos, Clear All</li>
              <li>• Bulk Import (including Replace All)</li>
              <li>• Costing & Pricing, Financials</li>
              <li>• Source Data & Access Management</li>
            </ul>
          </div>
          <div className="p-4 rounded-xl" style={{backgroundColor:C.bgLight}}>
            <div className="flex items-center gap-2 mb-2"><Users size={16} style={{color:C.zenkyOrange}}/><span className="font-bold text-sm" style={{fontFamily:F.display,color:C.zenkyOrange}}>Other Users</span></div>
            <ul className="text-xs space-y-1" style={{color:C.darkText}}>
              <li>• Add new SKUs & combos (no edit/delete)</li>
              <li>• Upload sales reports</li>
              <li>• View Dashboard, Combo Readiness, Reports, Sales Report</li>
              <li>• Cannot clear data, bulk-replace, view Financials, or Source Data</li>
            </ul>
          </div>
        </div>
        <p className="text-xs mt-4 p-2.5 rounded-lg" style={{backgroundColor:"#FFF3E6",color:"#9a5b0f"}}>
          ⚠️ Team members below get their own username/password and admin-visibility setting, which is a real improvement — but this is still stored with your workspace data, not enforced by a real authentication server (no password hashing, no session tokens, no audit-proof logging of who did what). Good enough to keep casual visitors out and reduce accidental admin access — not bank-grade security. For that, Supabase Auth would be the next upgrade.
        </p>
      </Card>

      <Card className="mb-6">
        <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Change Login Credentials</h3>
        <p className="text-xs mb-3" style={{color:C.lightText}}>This is the username/password required to open the app at all — separate from the admin PIN below.</p>
        <div className="grid sm:grid-cols-3 gap-3 max-w-2xl">
          <Input placeholder="New username" value={newUsername} onChange={e=>setNewUsername(e.target.value)}/>
          <Input type="password" placeholder="New password (min 6 chars)" value={newPassword} onChange={e=>setNewPassword(e.target.value)}/>
          <Input type="password" placeholder="Confirm password" value={confirmPassword} onChange={e=>setConfirmPassword(e.target.value)}/>
        </div>
        <div className="mt-3"><PrimaryButton onClick={saveLogin}><Save size={15}/>Update Login Credentials</PrimaryButton></div>
        <p className="text-xs mt-3" style={{color:C.lightText}}>Current username is {loginCreds?.username?`"${loginCreds.username}" (set)`:`the default ("${DEFAULT_LOGIN.username}") — change this before sharing your app link.`}</p>
      </Card>

      <Card className="mb-6">
        <h3 className="font-bold text-lg mb-2" style={{fontFamily:F.display,color:C.darkText}}>Team Members</h3>
        <p className="text-xs mb-4" style={{color:C.lightText}}>Add named logins for your team, each with their own username and password. "Can request admin access" controls whether they even see the option to unlock admin — if unchecked, that option is hidden from them entirely, not just PIN-protected.</p>

        <div className="grid sm:grid-cols-3 gap-2 mb-2">
          <Input placeholder="Display name" value={userForm.name} onChange={e=>setUserForm({...userForm,name:e.target.value})}/>
          <Input placeholder="Username" value={userForm.username} onChange={e=>setUserForm({...userForm,username:e.target.value})}/>
          <Input type="password" placeholder="Password (min 6 chars)" value={userForm.password} onChange={e=>setUserForm({...userForm,password:e.target.value})}/>
        </div>
        <label className="flex items-center gap-2 mb-3 cursor-pointer">
          <input type="checkbox" checked={userForm.canBeAdmin} onChange={e=>setUserForm({...userForm,canBeAdmin:e.target.checked})}/>
          <span className="text-sm" style={{color:C.darkText}}>Can request admin access (sees the "unlock admin" option and may enter the PIN)</span>
        </label>
        <PrimaryButton onClick={addUser}><Plus size={16}/>Add Team Member</PrimaryButton>

        {users.length>0&&(
          <div className="mt-5 overflow-x-auto"><table className="w-full text-sm">
            <thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Name</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Username</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Admin Option Visible?</th><th/></tr></thead>
            <tbody>{users.map(u=>{
              const isEdit=editUserId===u.id;
              return(
                <tr key={u.id} className="border-t" style={{borderColor:C.border}}>
                  <td className="py-2 pr-3 font-bold">{isEdit?<Input value={editUserValues.name} onChange={e=>setEditUserValues({...editUserValues,name:e.target.value})}/>:u.name}</td>
                  <td className="py-2 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{u.username}</td>
                  <td className="py-2 pr-3">
                    {isEdit?(
                      <label className="flex items-center gap-2 cursor-pointer"><input type="checkbox" checked={editUserValues.canBeAdmin} onChange={e=>setEditUserValues({...editUserValues,canBeAdmin:e.target.checked})}/><span className="text-xs">Can request admin</span></label>
                    ):(
                      u.canBeAdmin?<Stamp tone="purple">Visible</Stamp>:<Stamp tone="pink">Hidden</Stamp>
                    )}
                  </td>
                  <td className="py-2 pr-3">
                    <div className="flex items-center gap-1 justify-end">
                      {isEdit?(
                        <>
                          <Input type="password" placeholder="New password (optional)" value={editUserValues.password||""} onChange={e=>setEditUserValues({...editUserValues,password:e.target.value})} className="w-40"/>
                          <GhostButton title="Save" onClick={()=>saveUserEdit(u.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setEditUserId(null)}><X size={13}/></GhostButton>
                        </>
                      ):deleteUserId===u.id?(
                        <><GhostButton title="Confirm" onClick={()=>removeUser(u.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setDeleteUserId(null)}><X size={13}/></GhostButton></>
                      ):(
                        <><GhostButton title="Edit" onClick={()=>{setEditUserId(u.id);setEditUserValues({name:u.name,canBeAdmin:u.canBeAdmin,password:""});}}><Pencil size={13}/></GhostButton><GhostButton title="Delete" onClick={()=>setDeleteUserId(u.id)}><Trash2 size={13}/></GhostButton></>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}</tbody>
          </table></div>
        )}
      </Card>

      <Card>
        <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Change Admin PIN</h3>
        <div className="grid sm:grid-cols-2 gap-3 max-w-md">
          <Input type="password" placeholder="New PIN (min 4 characters)" value={newPin} onChange={e=>setNewPin(e.target.value)}/>
          <Input type="password" placeholder="Confirm new PIN" value={confirmPin} onChange={e=>setConfirmPin(e.target.value)}/>
        </div>
        <div className="mt-3"><PrimaryButton onClick={savePin}><Save size={15}/>Update PIN</PrimaryButton></div>
        <p className="text-xs mt-3" style={{color:C.lightText}}>Current PIN is {adminPin?"set (hidden)":`the default (${DEFAULT_ADMIN_PIN}) — change this before sharing your app link.`}</p>
      </Card>
    </div>
  );
}

/* ═══ FINANCIALS ═══ */
function FinancialsView({investors,setInvestors,investments,setInvestments,expenses,setExpenses,income,setIncome,salesLines,skus,combos,reports,activityLog,adminPin,loginCreds,forceSaveNow,logActivity,showToast}){

  const [tab,setTab]=useState("overview");
  const [viewingInvestorId,setViewingInvestorId]=useState(null); // set to show the dedicated Investor Statement page instead of normal tab content
  const TABS=[
    {id:"overview",label:"Overview"},
    {id:"investors",label:"Investors"},
    {id:"expenses",label:"Expenses"},
    {id:"income",label:"Income"},
    {id:"reports",label:"Reports"},
  ];

  // Dedicated financial-data-only backup — separate from the full app backup
  // in Source Data. This one matters more to track closely: unlike SKU/Combo
  // data (which can be re-imported from an Excel file if lost), investors,
  // expenses, and income are entered by hand with no separate source file to
  // fall back on, so losing them means genuinely re-entering everything.
  function downloadFinancialBackup(){
    const backup={exportedAt:new Date().toISOString(),investors,investments,expenses,income};
    const json=JSON.stringify(backup,null,2);
    const blob=new Blob([json],{type:"application/json"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    const dateStr=new Date().toISOString().slice(0,10);
    a.href=url;a.download=`zenkybox_financial_backup_${dateStr}.json`;
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    URL.revokeObjectURL(url);
    logActivity?.("Downloaded financial backup",`${investors.length} investors, ${investments.length} investments, ${expenses.length} expenses, ${income.length} income entries`);
    showToast("success","Financial backup downloaded. Keep it somewhere safe outside Supabase. 💾");
    if(typeof window!=="undefined")localStorage.setItem("zenkybox-last-financial-backup",new Date().toISOString());
  }
  const lastFinBackup=typeof window!=="undefined"?localStorage.getItem("zenkybox-last-financial-backup"):null;
  const daysSinceFinBackup=lastFinBackup?Math.floor((Date.now()-new Date(lastFinBackup).getTime())/86400000):null;

  const totalInvested=useMemo(()=>investments.reduce((s,i)=>s+Number(i.amount||0),0),[investments]);
  const totalIncome=useMemo(()=>income.reduce((s,i)=>s+Number(i.amount||0),0),[income]);
  const totalExpenses=useMemo(()=>expenses.reduce((s,e)=>s+Number(e.amount||0),0),[expenses]);
  const fundBalance=totalInvested+totalIncome-totalExpenses;

  function exportCsv(rows,headers,filename){
    let csv=headers.join(",")+"\n";
    rows.forEach(r=>csv+=headers.map(h=>{const v=r[h];return typeof v==="string"&&v.includes(",")?`"${v}"`:v;}).join(",")+"\n");
    downloadCsv(filename,csv);
  }

  /* ── Overview ── */
  function Overview(){
    const recent=useMemo(()=>{
      const all=[
        ...investments.map(i=>({...i,kind:"Investment",head:i.investorName})),
        ...income.map(i=>({...i,kind:"Income"})),
        ...expenses.map(e=>({...e,kind:"Expense"})),
      ];
      return all.sort((a,b)=>new Date(b.date)-new Date(a.date)).slice(0,10);
    },[]);
    return(
      <div>
        <Card className="mb-6" style={{borderColor:daysSinceFinBackup>7?"#fecaca":C.zenkyPurple}}>
          <div className="flex items-center gap-2 mb-3">
            <Download size={18} style={{color:C.zenkyPurple}}/>
            <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>Financial Data Backup</h3>
          </div>
          <p className="text-sm mb-3" style={{color:C.darkText,fontFamily:F.body}}>
            Downloads your investors, investments, expenses, and income as a JSON file. Unlike your SKU Catalog (which you can always re-import from an Excel file), this data is entered by hand — losing it means starting over. Back it up weekly.
          </p>
          {lastFinBackup?(
            <p className="text-xs mb-3" style={{color:daysSinceFinBackup>7?"#dc2626":C.lightText,fontWeight:daysSinceFinBackup>7?700:400}}>
              Last backup: {daysSinceFinBackup===0?"today":`${daysSinceFinBackup} day${daysSinceFinBackup!==1?"s":""} ago`}
              {daysSinceFinBackup>7&&" — overdue, download a fresh one now"}
            </p>
          ):(
            <p className="text-xs mb-3" style={{color:"#dc2626",fontWeight:700}}>No financial backup has ever been downloaded on this device.</p>
          )}
          <PrimaryButton onClick={downloadFinancialBackup}><Download size={15}/>Download Financial Backup Now</PrimaryButton>
          <p className="text-xs mt-3" style={{color:C.lightText}}>Store it outside Supabase/Vercel — Google Drive, email to yourself, or a folder on your computer.</p>
        </Card>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-6">
          <Card><div className="text-xs font-bold uppercase" style={{color:C.lightText}}>Total Invested</div><div className="text-2xl font-black mt-1" style={{fontFamily:F.display,color:C.zenkyPurple}}>{fmtINR(totalInvested)}</div></Card>
          <Card><div className="text-xs font-bold uppercase" style={{color:C.lightText}}>Total Income</div><div className="text-2xl font-black mt-1" style={{fontFamily:F.display,color:C.mintGreen}}>{fmtINR(totalIncome)}</div></Card>
          <Card><div className="text-xs font-bold uppercase" style={{color:C.lightText}}>Total Expenses</div><div className="text-2xl font-black mt-1" style={{fontFamily:F.display,color:C.zenkyOrange}}>{fmtINR(totalExpenses)}</div></Card>
          <Card style={{borderColor:fundBalance>=0?C.mintGreen:C.zenkyPink}}><div className="text-xs font-bold uppercase" style={{color:C.lightText}}>Fund Balance</div><div className="text-2xl font-black mt-1" style={{fontFamily:F.display,color:fundBalance>=0?C.mintGreen:C.zenkyPink}}>{fmtINR(fundBalance)}</div></Card>
        </div>
        <Card>
          <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Recent Activity</h3>
          {recent.length===0?<Empty icon={IndianRupee} title="No financial activity yet" message="Add investors, income, or expenses to see them here."/>:(
            <div className="overflow-x-auto"><table className="w-full text-sm">
              <thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Date</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Type</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Head</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Amount</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Comment</th></tr></thead>
              <tbody>{recent.map(r=>(
                <tr key={r.id} className="border-t" style={{borderColor:C.border}}>
                  <td className="py-2 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{r.date}</td>
                  <td className="py-2 pr-3"><Stamp tone={r.kind==="Expense"?"pink":r.kind==="Income"?"mint":"purple"}>{r.kind}</Stamp></td>
                  <td className="py-2 pr-3">{r.head}</td>
                  <td className="py-2 pr-3 font-bold" style={{fontFamily:F.mono}}>{fmtINR(r.amount)}</td>
                  <td className="py-2 pr-3" style={{color:C.lightText}}>{r.comment||"—"}</td>
                </tr>
              ))}</tbody>
            </table></div>
          )}
        </Card>
      </div>
    );
  }

  /* ── Investors ── */
  function Investors(){
    const [invForm,setInvForm]=useState({name:"",contact:"",notes:""});
    const [txForm,setTxForm]=useState({investorId:"",amount:"",date:new Date().toISOString().slice(0,10),paymentMode:"",comment:""});
    const [editInvId,setEditInvId]=useState(null);
    const [editInvValues,setEditInvValues]=useState({});
    const [deleteInvId,setDeleteInvId]=useState(null);
    const [editTxId,setEditTxId]=useState(null);
    const [editTxValues,setEditTxValues]=useState({});
    const [deleteTxId,setDeleteTxId]=useState(null);

    function addInvestor(){
      if(!invForm.name.trim()){showToast("error","Investor name required.");return;}
      const inv={id:Date.now().toString(),name:invForm.name.trim(),contact:invForm.contact.trim(),notes:invForm.notes.trim()};
      setInvestors([...investors,inv]);
      logActivity?.("Investor added",inv.name);
      showToast("success",`Added investor ${inv.name}. ✨`);
      setInvForm({name:"",contact:"",notes:""});
    }
    function saveInvestorEdit(id){
      setInvestors(investors.map(i=>i.id===id?{...i,name:editInvValues.name,contact:editInvValues.contact,notes:editInvValues.notes}:i));
      // Keep investment records' cached investorName in sync with a rename
      setInvestments(investments.map(t=>t.investorId===id?{...t,investorName:editInvValues.name}:t));
      logActivity?.("Investor edited",editInvValues.name);
      showToast("success","Saved. ✨");
      setEditInvId(null);
    }
    function removeInvestor(id){
      const linked=investments.filter(t=>t.investorId===id);
      const linkedFromExpense=linked.filter(t=>t.fromExpenseId).length;
      if(linked.length>0){
        const note=linkedFromExpense>0?` (${linkedFromExpense} came from expenses they covered — remove those in the Expenses tab instead)`:"";
        showToast("error",`Can't delete — ${linked.length} investment record${linked.length!==1?"s are":" is"} logged against this investor${note}.`);
        setDeleteInvId(null);return;
      }
      const inv=investors.find(i=>i.id===id);
      setInvestors(investors.filter(i=>i.id!==id));
      logActivity?.("Investor deleted",inv?.name||id);
      showToast("success","Investor removed.");
      setDeleteInvId(null);
    }

    function addInvestment(){
      const investor=investors.find(i=>i.id===txForm.investorId);
      const amount=Number(txForm.amount);
      if(!investor){showToast("error","Choose an investor.");return;}
      if(!amount||amount<=0){showToast("error","Enter a valid amount.");return;}
      const tx={id:Date.now().toString(),investorId:investor.id,investorName:investor.name,amount,date:txForm.date,paymentMode:txForm.paymentMode,comment:txForm.comment.trim()};
      setInvestments([...investments,tx]);
      logActivity?.("Investment logged",`${investor.name} — ${fmtINR(amount)}`);
      showToast("success",`Logged ${fmtINR(amount)} from ${investor.name}. ✨`);
      setTxForm({investorId:"",amount:"",date:txForm.date,paymentMode:"",comment:""});
    }
    function saveTxEdit(id){
      const amount=Number(editTxValues.amount);
      if(!amount||amount<=0){showToast("error","Enter a valid amount.");return;}
      setInvestments(investments.map(t=>t.id===id?{...t,amount,date:editTxValues.date,paymentMode:editTxValues.paymentMode,comment:editTxValues.comment}:t));
      logActivity?.("Investment edited",id);
      showToast("success","Saved. ✨");
      setEditTxId(null);
    }
    function removeTx(id){
      setInvestments(investments.filter(t=>t.id!==id));
      logActivity?.("Investment deleted",id);
      showToast("success","Investment removed.");
      setDeleteTxId(null);
    }

    const perInvestor=useMemo(()=>{
      const totals={};
      investments.forEach(t=>{totals[t.investorId]=(totals[t.investorId]||0)+Number(t.amount||0);});
      return totals;
    },[investments]);

    return(
      <div>
        <Card className="mb-6">
          <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Investor Master</h3>
          <div className="grid sm:grid-cols-3 gap-2 mb-3">
            <Input placeholder="Investor name" value={invForm.name} onChange={e=>setInvForm({...invForm,name:e.target.value})}/>
            <Input placeholder="Contact (phone/email)" value={invForm.contact} onChange={e=>setInvForm({...invForm,contact:e.target.value})}/>
            <Input placeholder="Notes (optional)" value={invForm.notes} onChange={e=>setInvForm({...invForm,notes:e.target.value})}/>
          </div>
          <PrimaryButton onClick={addInvestor}><Plus size={16}/>Add Investor</PrimaryButton>

          {investors.length>0&&(
            <div className="mt-5 overflow-x-auto"><table className="w-full text-sm">
              <thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Name</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Contact</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Notes</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Total Invested</th><th/></tr></thead>
              <tbody>{investors.map(inv=>{
                const isEdit=editInvId===inv.id;
                return(
                  <tr key={inv.id} className="border-t" style={{borderColor:C.border}}>
                    <td className="py-2 pr-3 font-bold">{isEdit?<Input value={editInvValues.name} onChange={e=>setEditInvValues({...editInvValues,name:e.target.value})}/>:<button onClick={()=>setViewingInvestorId(inv.id)} className="underline hover:no-underline" style={{color:C.zenkyPurple}}>{inv.name}</button>}</td>
                    <td className="py-2 pr-3" style={{color:C.lightText}}>{isEdit?<Input value={editInvValues.contact} onChange={e=>setEditInvValues({...editInvValues,contact:e.target.value})}/>:(inv.contact||"—")}</td>
                    <td className="py-2 pr-3" style={{color:C.lightText}}>{isEdit?<Input value={editInvValues.notes} onChange={e=>setEditInvValues({...editInvValues,notes:e.target.value})}/>:(inv.notes||"—")}</td>
                    <td className="py-2 pr-3 font-bold" style={{fontFamily:F.mono,color:C.zenkyPurple}}>{fmtINR(perInvestor[inv.id]||0)}</td>
                    <td className="py-2 pr-3">
                      <div className="flex items-center gap-1 justify-end">
                        {isEdit?(<><GhostButton title="Save" onClick={()=>saveInvestorEdit(inv.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setEditInvId(null)}><X size={13}/></GhostButton></>)
                        :deleteInvId===inv.id?(<><GhostButton title="Confirm" onClick={()=>removeInvestor(inv.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setDeleteInvId(null)}><X size={13}/></GhostButton></>)
                        :(<><GhostButton title="Edit" onClick={()=>{setEditInvId(inv.id);setEditInvValues({name:inv.name,contact:inv.contact,notes:inv.notes});}}><Pencil size={13}/></GhostButton><GhostButton title="Delete" onClick={()=>setDeleteInvId(inv.id)}><Trash2 size={13}/></GhostButton></>)}
                      </div>
                    </td>
                  </tr>
                );
              })}</tbody>
            </table></div>
          )}
        </Card>

        <Card>
          <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Log an Investment</h3>
          {investors.length===0?<p className="text-sm" style={{color:C.lightText}}>Add an investor above first.</p>:(
            <>
              <div className="grid sm:grid-cols-3 gap-2 mb-2">
                <Select value={txForm.investorId} onChange={e=>setTxForm({...txForm,investorId:e.target.value})}>
                  <option value="">Select investor…</option>
                  {investors.map(i=><option key={i.id} value={i.id}>{i.name}</option>)}
                </Select>
                <div className="relative"><span className="absolute left-3 top-2.5 text-sm" style={{color:C.lightText}}>₹</span><Input placeholder="Amount" type="number" className="pl-6" value={txForm.amount} onChange={e=>setTxForm({...txForm,amount:e.target.value})}/></div>
                <Input type="date" value={txForm.date} onChange={e=>setTxForm({...txForm,date:e.target.value})}/>
              </div>
              <div className="grid sm:grid-cols-2 gap-2 mb-3">
                <Select value={txForm.paymentMode} onChange={e=>setTxForm({...txForm,paymentMode:e.target.value})}>
                  <option value="">Payment mode…</option>
                  {PAYMENT_MODES.map(m=><option key={m} value={m}>{m}</option>)}
                </Select>
                <Input placeholder="Comment (optional)" value={txForm.comment} onChange={e=>setTxForm({...txForm,comment:e.target.value})}/>
              </div>
              <PrimaryButton onClick={addInvestment}><Plus size={16}/>Log Investment</PrimaryButton>
            </>
          )}
          {investments.length>0&&(
            <div className="mt-5 overflow-x-auto"><table className="w-full text-sm">
              <thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Date</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Investor</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Amount</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Mode</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Comment</th><th/></tr></thead>
              <tbody>{[...investments].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(t=>{
                const isEdit=editTxId===t.id;
                const isLinked=!!t.fromExpenseId;
                return(
                  <tr key={t.id} className="border-t" style={{borderColor:C.border}}>
                    <td className="py-2 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{isEdit?<Input type="date" value={editTxValues.date} onChange={e=>setEditTxValues({...editTxValues,date:e.target.value})}/>:t.date}</td>
                    <td className="py-2 pr-3 font-bold">{t.investorName}</td>
                    <td className="py-2 pr-3 font-bold" style={{fontFamily:F.mono,color:C.zenkyPurple}}>{isEdit?<Input type="number" value={editTxValues.amount} onChange={e=>setEditTxValues({...editTxValues,amount:e.target.value})}/>:fmtINR(t.amount)}</td>
                    <td className="py-2 pr-3">{isEdit?<Select value={editTxValues.paymentMode} onChange={e=>setEditTxValues({...editTxValues,paymentMode:e.target.value})}><option value="">—</option>{PAYMENT_MODES.map(m=><option key={m} value={m}>{m}</option>)}</Select>:(t.paymentMode?<Stamp tone="blue">{t.paymentMode}</Stamp>:"—")}</td>
                    <td className="py-2 pr-3" style={{color:C.lightText}}>
                      {isEdit?<Input value={editTxValues.comment} onChange={e=>setEditTxValues({...editTxValues,comment:e.target.value})}/>:(
                        <div className="flex items-center gap-1.5">
                          {t.comment||"—"}
                          {isLinked&&<Stamp tone="orange">🔗 auto</Stamp>}
                        </div>
                      )}
                    </td>
                    <td className="py-2 pr-3">
                      {isLinked?(
                        <span className="text-xs" style={{color:C.lightText}}>Edit via Expenses</span>
                      ):(
                        <div className="flex items-center gap-1 justify-end">
                          {isEdit?(<><GhostButton title="Save" onClick={()=>saveTxEdit(t.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setEditTxId(null)}><X size={13}/></GhostButton></>)
                          :deleteTxId===t.id?(<><GhostButton title="Confirm" onClick={()=>removeTx(t.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setDeleteTxId(null)}><X size={13}/></GhostButton></>)
                          :(<><GhostButton title="Edit" onClick={()=>{setEditTxId(t.id);setEditTxValues({date:t.date,amount:t.amount,paymentMode:t.paymentMode||"",comment:t.comment});}}><Pencil size={13}/></GhostButton><GhostButton title="Delete" onClick={()=>setDeleteTxId(t.id)}><Trash2 size={13}/></GhostButton></>)}
                        </div>
                      )}
                    </td>
                  </tr>
                );
              })}</tbody>
            </table></div>
          )}
        </Card>
      </div>
    );
  }

  /* ── Expenses ── */
  function Expenses(){
    const [form,setForm]=useState({date:new Date().toISOString().slice(0,10),head:"",amount:"",spentBy:"",paidTo:"",paymentMode:"",comment:""});
    const [editId,setEditId]=useState(null);
    const [editValues,setEditValues]=useState({});
    const [deleteId,setDeleteId]=useState(null);
    function add(){
      const amount=Number(form.amount);
      if(!form.head){showToast("error","Choose an expense head.");return;}
      if(!amount||amount<=0){showToast("error","Enter a valid amount.");return;}
      const spentByInvestor=investors.find(i=>i.id===form.spentBy);
      const e={id:Date.now().toString(),date:form.date,head:form.head,amount,spentBy:form.spentBy,spentByName:spentByInvestor?.name||"",paidTo:form.paidTo.trim(),paymentMode:form.paymentMode,comment:form.comment.trim()};
      setExpenses([...expenses,e]);
      // If this expense was personally covered by an investor, treat it as
      // capital they've put into the business — auto-log a linked investment
      // so their Total Invested reflects it too. Net effect on Fund Balance
      // is zero (the expense debit is offset by the investment credit), which
      // is correct: no company-pool cash moved, the investor paid the vendor directly.
      if(spentByInvestor){
        const linkedInv={id:`auto-${e.id}`,investorId:spentByInvestor.id,investorName:spentByInvestor.name,amount,date:form.date,paymentMode:form.paymentMode,comment:`Covered expense: ${form.head}`,fromExpenseId:e.id};
        setInvestments([...investments,linkedInv]);
      }
      logActivity?.("Expense added",`${form.head} — ${fmtINR(amount)}${spentByInvestor?` (covered by ${spentByInvestor.name} — added to their investment total)`:""}`);
      showToast("success",spentByInvestor?`Added expense — ${fmtINR(amount)}, credited to ${spentByInvestor.name}'s investment. ✨`:`Added expense — ${fmtINR(amount)}. ✨`);
      setForm({date:form.date,head:"",amount:"",spentBy:"",paidTo:"",paymentMode:"",comment:""});
    }
    function saveEdit(id){
      const amount=Number(editValues.amount);
      if(!editValues.head){showToast("error","Choose an expense head.");return;}
      if(!amount||amount<=0){showToast("error","Enter a valid amount.");return;}
      const spentByInvestor=investors.find(i=>i.id===editValues.spentBy);
      setExpenses(expenses.map(e=>e.id===id?{...e,date:editValues.date,head:editValues.head,amount,spentBy:editValues.spentBy,spentByName:spentByInvestor?.name||"",paidTo:editValues.paidTo,paymentMode:editValues.paymentMode,comment:editValues.comment}:e));
      // Reconcile the linked investment: drop any previous auto-entry for this
      // expense, then re-add one if a "spent by" investor is still set.
      const withoutOld=investments.filter(t=>t.fromExpenseId!==id);
      const newInvestments=spentByInvestor
        ?[...withoutOld,{id:`auto-${id}`,investorId:spentByInvestor.id,investorName:spentByInvestor.name,amount,date:editValues.date,paymentMode:editValues.paymentMode,comment:`Covered expense: ${editValues.head}`,fromExpenseId:id}]
        :withoutOld;
      setInvestments(newInvestments);
      logActivity?.("Expense edited",`${editValues.head} — ${fmtINR(amount)}`);
      showToast("success","Saved. ✨");
      setEditId(null);
    }
    function removeExpense(id){
      setExpenses(expenses.filter(e=>e.id!==id));
      setInvestments(investments.filter(t=>t.fromExpenseId!==id)); // remove any linked auto-investment too
      logActivity?.("Expense deleted",id);
      showToast("success","Expense removed.");
      setDeleteId(null);
    }

    // ── Bulk upload from Excel/CSV ──
    const [entryMode,setEntryMode]=useState("single"); // "single" | "bulk"
    const [bulkStage,setBulkStage]=useState("idle"); // "idle" | "preview" | "done"
    const [bulkFileName,setBulkFileName]=useState("");
    const [bulkRows,setBulkRows]=useState([]);
    const bulkFileRef=useRef(null);

    function downloadExpenseTemplate(){
      const csv="date,head,amount,spentBy,paidTo,paymentMode,comment\n"+
        "2026-05-01,Product Procurement,19991,Rahul,RHS Enterprise,Cash,Initial inventory stock\n"+
        "2026-06-01,Packaging Expenses,350,Rahul,N/A,Cash,Packaging tape x6 rolls\n";
      downloadCsv("expense_bulk_upload_template.csv",csv);
    }

    function processBulkFile(rows){
      if(!rows?.length){showToast("error","File is empty.");return;}
      const parsed=rows.map((r,i)=>{
        const dateRaw=r.date||r.Date;
        const headRaw=String(r.head||r.Head||"").trim();
        const amount=Number(r.amount||r.Amount);
        const spentByRaw=String(r.spentBy||r["Spent By"]||"").trim();
        const paidTo=String(r.paidTo||r["Paid To"]||"").trim();
        const paymentMode=String(r.paymentMode||r["Payment Mode"]||"").trim();
        const comment=String(r.comment||r.Comment||"").trim();
        // Match head case-insensitively against known heads
        const matchedHead=EXPENSE_HEADS.find(h=>h.toLowerCase()===headRaw.toLowerCase())||"";
        // Match spentBy against existing investors by name (case-insensitive)
        const matchedInvestor=investors.find(inv=>inv.name.toLowerCase()===spentByRaw.toLowerCase());
        let date="";
        if(dateRaw instanceof Date)date=dateRaw.toISOString().slice(0,10);
        else if(typeof dateRaw==="string"&&dateRaw){const d=new Date(dateRaw);date=isNaN(d)?"":d.toISOString().slice(0,10);}
        return{
          rowNum:i+2,date,head:matchedHead||headRaw,headMatched:!!matchedHead,amount,
          spentByRaw,spentById:matchedInvestor?.id||"",spentByName:matchedInvestor?.name||"",
          paidTo,paymentMode,comment,
          valid:!!date&&!!matchedHead&&amount>0,
        };
      });
      setBulkRows(parsed);
      setBulkStage("preview");
    }

    function handleBulkFile(file){
      if(!file)return;setBulkFileName(file.name);
      const ext=file.name.split(".").pop().toLowerCase();
      if(ext==="csv"){Papa.parse(file,{header:true,skipEmptyLines:true,complete:res=>processBulkFile(res.data),error:()=>showToast("error","Could not parse CSV.")});}
      else if(ext==="xlsx"||ext==="xls"){const r=new FileReader();r.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:"array"});processBulkFile(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""}));}catch{showToast("error","Could not parse spreadsheet.");}};r.readAsArrayBuffer(file);}
      else showToast("error","Upload a .csv or .xlsx file.");
    }

    function confirmBulkImport(){
      const validRows=bulkRows.filter(r=>r.valid);
      if(!validRows.length){showToast("error","No valid rows to import.");return;}
      const newExpenses=[];const newLinkedInvestments=[];
      validRows.forEach((r,i)=>{
        const id=`bulk-${Date.now()}-${i}`;
        newExpenses.push({id,date:r.date,head:r.head,amount:r.amount,spentBy:r.spentById,spentByName:r.spentByName,paidTo:r.paidTo,paymentMode:r.paymentMode,comment:r.comment});
        if(r.spentById){
          newLinkedInvestments.push({id:`auto-${id}`,investorId:r.spentById,investorName:r.spentByName,amount:r.amount,date:r.date,paymentMode:r.paymentMode,comment:`Covered expense: ${r.head}`,fromExpenseId:id});
        }
      });
      setExpenses([...expenses,...newExpenses]);
      if(newLinkedInvestments.length)setInvestments([...investments,...newLinkedInvestments]);
      logActivity?.("Bulk expenses imported",`${bulkFileName}: ${newExpenses.length} expenses added, ${newLinkedInvestments.length} linked to investors`);
      showToast("success",`Imported ${newExpenses.length} expenses. ✨`);
      setBulkStage("done");
    }

    function resetBulk(){setBulkStage("idle");setBulkRows([]);setBulkFileName("");if(bulkFileRef.current)bulkFileRef.current.value="";}

    return(
      <div>
        <div className="flex gap-1.5 mb-4">
          {[{id:"single",label:"Add Single Expense"},{id:"bulk",label:"Bulk Upload from Excel"}].map(m=>(
            <button key={m.id} onClick={()=>{setEntryMode(m.id);resetBulk();}} className="px-3.5 py-1.5 rounded-full text-xs font-bold transition-colors" style={{backgroundColor:entryMode===m.id?C.bgLight:"transparent",color:entryMode===m.id?C.zenkyPurple:C.lightText,border:`1.5px solid ${entryMode===m.id?C.zenkyPurple:C.border}`,fontFamily:F.body}}>
              {m.label}
            </button>
          ))}
        </div>

        {entryMode==="bulk"&&(
          <Card className="mb-6">
            <h3 className="font-bold text-lg mb-2" style={{fontFamily:F.display,color:C.darkText}}>Bulk Upload Expenses</h3>
            <p className="text-xs mb-4" style={{color:C.lightText}}>Columns expected: date, head, amount, spentBy, paidTo, paymentMode, comment. "head" must match one of the expense heads exactly (case-insensitive); "spentBy" matches an existing investor by name.</p>
            <button onClick={downloadExpenseTemplate} className="inline-flex items-center gap-1.5 text-xs font-bold mb-4" style={{color:C.zenkyOrange,fontFamily:F.body}}><Download size={13}/>Download Template</button>

            {bulkStage==="idle"&&(
              <div className="rounded-2xl border-2 border-dashed p-8 text-center" style={{borderColor:C.zenkyPink,backgroundColor:"#FFF8FC"}}>
                <Upload size={28} className="mx-auto mb-3" style={{color:C.zenkyPink}}/>
                <p className="font-bold mb-4 text-sm" style={{color:C.darkText,fontFamily:F.display}}>Choose your expense file</p>
                <input ref={bulkFileRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={e=>handleBulkFile(e.target.files[0])}/>
                <PrimaryButton onClick={()=>bulkFileRef.current?.click()}><Upload size={15}/>Select File</PrimaryButton>
              </div>
            )}

            {bulkStage==="preview"&&(
              <div>
                <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
                  <span className="text-sm" style={{color:C.darkText}}>Parsed <strong>{bulkFileName}</strong> — {bulkRows.length} rows, {bulkRows.filter(r=>r.valid).length} valid</span>
                  <button onClick={resetBulk} className="text-sm font-bold" style={{color:C.lightText}}>Change file</button>
                </div>
                <div className="overflow-x-auto mb-4 max-h-96 overflow-y-auto"><table className="w-full text-xs">
                  <thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold uppercase">Row</th><th className="py-2 pr-3 text-left font-bold uppercase">Date</th><th className="py-2 pr-3 text-left font-bold uppercase">Head</th><th className="py-2 pr-3 text-left font-bold uppercase">Amount</th><th className="py-2 pr-3 text-left font-bold uppercase">Spent By</th><th className="py-2 pr-3 text-left font-bold uppercase">Status</th></tr></thead>
                  <tbody>{bulkRows.map(r=>(
                    <tr key={r.rowNum} className="border-t" style={{borderColor:C.border}}>
                      <td className="py-1.5 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{r.rowNum}</td>
                      <td className="py-1.5 pr-3" style={{fontFamily:F.mono,color:r.date?C.darkText:C.zenkyPink}}>{r.date||"missing"}</td>
                      <td className="py-1.5 pr-3" style={{color:r.headMatched?C.darkText:C.zenkyPink}}>{r.head||"missing"}{!r.headMatched&&r.head&&" (no match)"}</td>
                      <td className="py-1.5 pr-3" style={{fontFamily:F.mono,color:r.amount>0?C.darkText:C.zenkyPink}}>{r.amount||"missing"}</td>
                      <td className="py-1.5 pr-3">{r.spentByName||(r.spentByRaw?<span style={{color:C.zenkyOrange}}>{r.spentByRaw} (no match)</span>:"—")}</td>
                      <td className="py-1.5 pr-3">{r.valid?<Stamp tone="mint">Ready</Stamp>:<Stamp tone="pink">Skip</Stamp>}</td>
                    </tr>
                  ))}</tbody>
                </table></div>
                {bulkRows.some(r=>!r.valid)&&<p className="text-xs mb-3" style={{color:C.zenkyPink}}>Rows marked "Skip" are missing a date, valid amount, or matching head — they won't be imported. Fix and re-upload, or import the valid ones now and add the rest manually.</p>}
                <div className="flex gap-2"><PrimaryButton onClick={confirmBulkImport}><Check size={15}/>Import {bulkRows.filter(r=>r.valid).length} Valid Expenses</PrimaryButton><button onClick={resetBulk} className="text-sm font-bold" style={{color:C.lightText}}>Cancel</button></div>
              </div>
            )}

            {bulkStage==="done"&&(
              <div className="text-center py-6">
                <Check size={28} className="mx-auto mb-3" style={{color:C.mintGreen}}/>
                <p className="font-bold" style={{color:C.darkText,fontFamily:F.display}}>Import complete!</p>
                <div className="mt-4"><PrimaryButton onClick={resetBulk}><Upload size={15}/>Upload another file</PrimaryButton></div>
              </div>
            )}
          </Card>
        )}

        {entryMode==="single"&&(
        <Card className="mb-6">
          <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Add an Expense</h3>
          <div className="grid sm:grid-cols-3 gap-2 mb-2">
            <Input type="date" value={form.date} onChange={e=>setForm({...form,date:e.target.value})}/>
            <Select value={form.head} onChange={e=>setForm({...form,head:e.target.value})}>
              <option value="">Select head…</option>
              {EXPENSE_HEADS.map(h=><option key={h} value={h}>{h}</option>)}
            </Select>
            <div className="relative"><span className="absolute left-3 top-2.5 text-sm" style={{color:C.lightText}}>₹</span><Input placeholder="Amount" type="number" className="pl-6" value={form.amount} onChange={e=>setForm({...form,amount:e.target.value})}/></div>
          </div>
          <div className="grid sm:grid-cols-3 gap-2 mb-2">
            <Select value={form.spentBy} onChange={e=>setForm({...form,spentBy:e.target.value})}>
              <option value="">Spent by (accountability)…</option>
              {investors.map(i=><option key={i.id} value={i.id}>{i.name}</option>)}
            </Select>
            <Input placeholder="Paid to (name)" value={form.paidTo} onChange={e=>setForm({...form,paidTo:e.target.value})}/>
            <Select value={form.paymentMode} onChange={e=>setForm({...form,paymentMode:e.target.value})}>
              <option value="">Payment mode…</option>
              {PAYMENT_MODES.map(m=><option key={m} value={m}>{m}</option>)}
            </Select>
          </div>
          <div className="mb-3">
            <Input placeholder="Comment (optional)" value={form.comment} onChange={e=>setForm({...form,comment:e.target.value})}/>
          </div>
          {investors.length===0&&<p className="text-xs mb-3" style={{color:C.lightText}}>Tip: add people to Investor Master first so you can pick who's accountable for each expense.</p>}
          <PrimaryButton onClick={add}><Plus size={16}/>Add Expense</PrimaryButton>
        </Card>
        )}
        <Card>
          <div className="flex items-center justify-between mb-4">
            <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>{expenses.length} Expense{expenses.length!==1?"s":""}</h3>
            {expenses.length>0&&<button onClick={()=>exportCsv(expenses,["date","head","amount","spentByName","paidTo","paymentMode","comment"],"expenses.csv")} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export</button>}
          </div>
          {expenses.length===0?<Empty icon={IndianRupee} title="No expenses logged" message="Add your first expense above."/>:(
            <div className="overflow-x-auto"><table className="w-full text-sm">
              <thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Date</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Head</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Amount</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Spent By</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Paid To</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Mode</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Comment</th><th/></tr></thead>
              <tbody>{[...expenses].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(e=>{
                const isEdit=editId===e.id;
                return(
                  <tr key={e.id} className="border-t" style={{borderColor:C.border}}>
                    <td className="py-2 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{isEdit?<Input type="date" value={editValues.date} onChange={ev=>setEditValues({...editValues,date:ev.target.value})}/>:e.date}</td>
                    <td className="py-2 pr-3">{isEdit?<Select value={editValues.head} onChange={ev=>setEditValues({...editValues,head:ev.target.value})}>{EXPENSE_HEADS.map(h=><option key={h} value={h}>{h}</option>)}</Select>:<Stamp tone="orange">{e.head}</Stamp>}</td>
                    <td className="py-2 pr-3 font-bold" style={{fontFamily:F.mono,color:C.zenkyOrange}}>{isEdit?<Input type="number" value={editValues.amount} onChange={ev=>setEditValues({...editValues,amount:ev.target.value})}/>:fmtINR(e.amount)}</td>
                    <td className="py-2 pr-3">{isEdit?<Select value={editValues.spentBy} onChange={ev=>setEditValues({...editValues,spentBy:ev.target.value})}><option value="">—</option>{investors.map(i=><option key={i.id} value={i.id}>{i.name}</option>)}</Select>:(e.spentByName?<Stamp tone="purple">{e.spentByName}</Stamp>:"—")}</td>
                    <td className="py-2 pr-3">{isEdit?<Input value={editValues.paidTo} onChange={ev=>setEditValues({...editValues,paidTo:ev.target.value})}/>:(e.paidTo||"—")}</td>
                    <td className="py-2 pr-3">{isEdit?<Select value={editValues.paymentMode} onChange={ev=>setEditValues({...editValues,paymentMode:ev.target.value})}><option value="">—</option>{PAYMENT_MODES.map(m=><option key={m} value={m}>{m}</option>)}</Select>:(e.paymentMode?<Stamp tone="blue">{e.paymentMode}</Stamp>:"—")}</td>
                    <td className="py-2 pr-3" style={{color:C.lightText}}>{isEdit?<Input value={editValues.comment} onChange={ev=>setEditValues({...editValues,comment:ev.target.value})}/>:(e.comment||"—")}</td>
                    <td className="py-2 pr-3">
                      <div className="flex items-center gap-1 justify-end">
                        {isEdit?(<><GhostButton title="Save" onClick={()=>saveEdit(e.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setEditId(null)}><X size={13}/></GhostButton></>)
                        :deleteId===e.id?(<><GhostButton title="Confirm" onClick={()=>removeExpense(e.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setDeleteId(null)}><X size={13}/></GhostButton></>)
                        :(<><GhostButton title="Edit" onClick={()=>{setEditId(e.id);setEditValues({date:e.date,head:e.head,amount:e.amount,spentBy:e.spentBy||"",paidTo:e.paidTo||"",paymentMode:e.paymentMode||"",comment:e.comment});}}><Pencil size={13}/></GhostButton><GhostButton title="Delete" onClick={()=>setDeleteId(e.id)}><Trash2 size={13}/></GhostButton></>)}
                      </div>
                    </td>
                  </tr>
                );
              })}</tbody>
            </table></div>
          )}
        </Card>
      </div>
    );
  }

  /* ── Income ── */
  function Income(){
    const [form,setForm]=useState({date:new Date().toISOString().slice(0,10),head:"",amount:"",receivedFrom:"",paymentMode:"",comment:""});
    const [editId,setEditId]=useState(null);
    const [editValues,setEditValues]=useState({});
    const [deleteId,setDeleteId]=useState(null);
    const amazonRevenue=useMemo(()=>salesLines.filter(l=>l.channel==="amazon").reduce((s,l)=>s+l.revenue,0),[salesLines]);
    const websiteRevenue=useMemo(()=>salesLines.filter(l=>l.channel==="website").reduce((s,l)=>s+l.revenue,0),[salesLines]);
    function add(){
      const amount=Number(form.amount);
      if(!form.head){showToast("error","Choose an income head.");return;}
      if(!amount||amount<=0){showToast("error","Enter a valid amount.");return;}
      const i={id:Date.now().toString(),date:form.date,head:form.head,amount,receivedFrom:form.receivedFrom.trim(),paymentMode:form.paymentMode,comment:form.comment.trim()};
      setIncome([...income,i]);
      logActivity?.("Income added",`${form.head} — ${fmtINR(amount)}`);
      showToast("success",`Added income — ${fmtINR(amount)}. ✨`);
      setForm({date:form.date,head:"",amount:"",receivedFrom:"",paymentMode:"",comment:""});
    }
    function saveEdit(id){
      const amount=Number(editValues.amount);
      if(!editValues.head){showToast("error","Choose an income head.");return;}
      if(!amount||amount<=0){showToast("error","Enter a valid amount.");return;}
      setIncome(income.map(i=>i.id===id?{...i,date:editValues.date,head:editValues.head,amount,receivedFrom:editValues.receivedFrom,paymentMode:editValues.paymentMode,comment:editValues.comment}:i));
      logActivity?.("Income edited",`${editValues.head} — ${fmtINR(amount)}`);
      showToast("success","Saved. ✨");
      setEditId(null);
    }
    function removeIncome(id){
      setIncome(income.filter(i=>i.id!==id));
      logActivity?.("Income deleted",id);
      showToast("success","Income entry removed.");
      setDeleteId(null);
    }
    return(
      <div>
        {(amazonRevenue>0||websiteRevenue>0)&&(
          <div className="mb-4 p-3 rounded-xl text-xs" style={{backgroundColor:C.bgLight,color:C.lightText}}>
            For reference — ZenkyBox Sales Report currently shows <strong>{fmtINR(amazonRevenue)}</strong> Amazon revenue and <strong>{fmtINR(websiteRevenue)}</strong> Website revenue from uploaded orders. This is separate from what you log here — log actual payout amounts received if they differ (fees, timing, etc.).
          </div>
        )}
        <Card className="mb-6">
          <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Add Income</h3>
          <div className="grid sm:grid-cols-3 gap-2 mb-2">
            <Input type="date" value={form.date} onChange={e=>setForm({...form,date:e.target.value})}/>
            <Select value={form.head} onChange={e=>setForm({...form,head:e.target.value})}>
              <option value="">Select head…</option>
              {INCOME_HEADS.map(h=><option key={h} value={h}>{h}</option>)}
            </Select>
            <div className="relative"><span className="absolute left-3 top-2.5 text-sm" style={{color:C.lightText}}>₹</span><Input placeholder="Amount" type="number" className="pl-6" value={form.amount} onChange={e=>setForm({...form,amount:e.target.value})}/></div>
          </div>
          <div className="grid sm:grid-cols-3 gap-2 mb-3">
            <Input placeholder="Received from (name)" value={form.receivedFrom} onChange={e=>setForm({...form,receivedFrom:e.target.value})}/>
            <Select value={form.paymentMode} onChange={e=>setForm({...form,paymentMode:e.target.value})}>
              <option value="">Payment mode…</option>
              {PAYMENT_MODES.map(m=><option key={m} value={m}>{m}</option>)}
            </Select>
            <Input placeholder="Comment (optional)" value={form.comment} onChange={e=>setForm({...form,comment:e.target.value})}/>
          </div>
          <PrimaryButton onClick={add}><Plus size={16}/>Add Income</PrimaryButton>
        </Card>
        <Card>
          <div className="flex items-center justify-between mb-4">
            <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>{income.length} Income Entr{income.length!==1?"ies":"y"}</h3>
            {income.length>0&&<button onClick={()=>exportCsv(income,["date","head","amount","receivedFrom","paymentMode","comment"],"income.csv")} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export</button>}
          </div>
          {income.length===0?<Empty icon={IndianRupee} title="No income logged" message="Add your first income entry above."/>:(
            <div className="overflow-x-auto"><table className="w-full text-sm">
              <thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Date</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Head</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Amount</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Received From</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Mode</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Comment</th><th/></tr></thead>
              <tbody>{[...income].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(i=>{
                const isEdit=editId===i.id;
                return(
                  <tr key={i.id} className="border-t" style={{borderColor:C.border}}>
                    <td className="py-2 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{isEdit?<Input type="date" value={editValues.date} onChange={e=>setEditValues({...editValues,date:e.target.value})}/>:i.date}</td>
                    <td className="py-2 pr-3">{isEdit?<Select value={editValues.head} onChange={e=>setEditValues({...editValues,head:e.target.value})}>{INCOME_HEADS.map(h=><option key={h} value={h}>{h}</option>)}</Select>:<Stamp tone="mint">{i.head}</Stamp>}</td>
                    <td className="py-2 pr-3 font-bold" style={{fontFamily:F.mono,color:C.mintGreen}}>{isEdit?<Input type="number" value={editValues.amount} onChange={e=>setEditValues({...editValues,amount:e.target.value})}/>:fmtINR(i.amount)}</td>
                    <td className="py-2 pr-3">{isEdit?<Input value={editValues.receivedFrom} onChange={e=>setEditValues({...editValues,receivedFrom:e.target.value})}/>:(i.receivedFrom||"—")}</td>
                    <td className="py-2 pr-3">{isEdit?<Select value={editValues.paymentMode} onChange={e=>setEditValues({...editValues,paymentMode:e.target.value})}><option value="">—</option>{PAYMENT_MODES.map(m=><option key={m} value={m}>{m}</option>)}</Select>:(i.paymentMode?<Stamp tone="blue">{i.paymentMode}</Stamp>:"—")}</td>
                    <td className="py-2 pr-3" style={{color:C.lightText}}>{isEdit?<Input value={editValues.comment} onChange={e=>setEditValues({...editValues,comment:e.target.value})}/>:(i.comment||"—")}</td>
                    <td className="py-2 pr-3">
                      <div className="flex items-center gap-1 justify-end">
                        {isEdit?(<><GhostButton title="Save" onClick={()=>saveEdit(i.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setEditId(null)}><X size={13}/></GhostButton></>)
                        :deleteId===i.id?(<><GhostButton title="Confirm" onClick={()=>removeIncome(i.id)}><Check size={13}/></GhostButton><GhostButton title="Cancel" onClick={()=>setDeleteId(null)}><X size={13}/></GhostButton></>)
                        :(<><GhostButton title="Edit" onClick={()=>{setEditId(i.id);setEditValues({date:i.date,head:i.head,amount:i.amount,receivedFrom:i.receivedFrom||"",paymentMode:i.paymentMode||"",comment:i.comment});}}><Pencil size={13}/></GhostButton><GhostButton title="Delete" onClick={()=>setDeleteId(i.id)}><Trash2 size={13}/></GhostButton></>)}
                      </div>
                    </td>
                  </tr>
                );
              })}</tbody>
            </table></div>
          )}
        </Card>
      </div>
    );
  }

  /* ── Reports (month-over-month) ── */
  function Reports(){
    const [reportTab,setReportTab]=useState("pl");
    const REPORT_TABS=[
      {id:"pl",label:"P&L Statement"},
      {id:"datewise",label:"Date-wise Ledger"},
      {id:"fy",label:"Financial Year"},
      {id:"investorwise",label:"Investor-wise"},
      {id:"monthly",label:"Monthly"},
    ];

    // ── Shared: product/combo revenue+COGS, sourced from Sales Report data (salesLines) ──
    function productCombobreakdown(lines){
      const bySku={},byCombo={};
      lines.forEach(l=>{
        const bucket=l.matchType==="combo"?byCombo:bySku;
        if(!bucket[l.sku])bucket[l.sku]={code:l.sku,name:l.name,qty:0,revenue:0,cogs:0};
        bucket[l.sku].qty+=l.qty;bucket[l.sku].revenue+=l.revenue;bucket[l.sku].cogs+=l.cost;
      });
      return{skuRows:Object.values(bySku).sort((a,b)=>b.revenue-a.revenue),comboRows:Object.values(byCombo).sort((a,b)=>b.revenue-a.revenue)};
    }

    // Operating expenses exclude "Product Procurement" — that cost is already
    // reflected as COGS against units actually SOLD (via each SKU's procurement
    // cost in Sales Report). Counting the same rupees again here as an operating
    // expense would double-count it. Procurement is instead shown separately as
    // a cash-outflow reference, since buying inventory IS real money leaving the
    // business, just not a same-period "expense" in the profitability sense.
    const PROCUREMENT_HEAD="Product Procurement";
    // "Income from Amazon"/"Income from Website" are excluded from Other Income
    // for the same reason: that revenue is already counted via Sales Report.
    const SALES_INCOME_HEADS=["Income from Amazon","Income from Website"];

    function plFor(lines,expenseList,incomeList){
      const{skuRows,comboRows}=productCombobreakdown(lines);
      const totalRevenue=lines.reduce((s,l)=>s+l.revenue,0);
      const totalCOGS=lines.reduce((s,l)=>s+l.cost,0);
      const grossProfit=totalRevenue-totalCOGS;
      const opExByHead={};let totalOpEx=0,procurementTotal=0;
      expenseList.forEach(e=>{
        const amt=Number(e.amount||0);
        if(e.head===PROCUREMENT_HEAD){procurementTotal+=amt;return;}
        opExByHead[e.head]=(opExByHead[e.head]||0)+amt;totalOpEx+=amt;
      });
      const otherIncomeByHead={};let totalOtherIncome=0;
      incomeList.forEach(i=>{
        if(SALES_INCOME_HEADS.includes(i.head))return;
        const amt=Number(i.amount||0);
        otherIncomeByHead[i.head]=(otherIncomeByHead[i.head]||0)+amt;totalOtherIncome+=amt;
      });
      const netProfit=grossProfit-totalOpEx+totalOtherIncome;
      return{skuRows,comboRows,totalRevenue,totalCOGS,grossProfit,opExByHead,totalOpEx,procurementTotal,otherIncomeByHead,totalOtherIncome,netProfit};
    }

    const overallPL=useMemo(()=>plFor(salesLines,expenses,income),[]);

    function exportPL(pl,label){
      let csv=`Profit & Loss Statement — ${label}\n\n`;
      csv+="REVENUE BY PRODUCT (SKU)\nCode,Name,Qty,Revenue,COGS,Gross Profit\n";
      pl.skuRows.forEach(r=>csv+=`${r.code},"${r.name}",${r.qty},${r.revenue.toFixed(2)},${r.cogs.toFixed(2)},${(r.revenue-r.cogs).toFixed(2)}\n`);
      csv+="\nREVENUE BY COMBO\nCode,Name,Qty,Revenue,COGS,Gross Profit\n";
      pl.comboRows.forEach(r=>csv+=`${r.code},"${r.name}",${r.qty},${r.revenue.toFixed(2)},${r.cogs.toFixed(2)},${(r.revenue-r.cogs).toFixed(2)}\n`);
      csv+=`\nTOTAL REVENUE,${pl.totalRevenue.toFixed(2)}\nTOTAL COGS,${pl.totalCOGS.toFixed(2)}\nGROSS PROFIT,${pl.grossProfit.toFixed(2)}\n`;
      csv+="\nOPERATING EXPENSES\nHead,Amount\n";
      Object.entries(pl.opExByHead).forEach(([h,a])=>csv+=`"${h}",${a.toFixed(2)}\n`);
      csv+=`TOTAL OPERATING EXPENSES,${pl.totalOpEx.toFixed(2)}\n`;
      csv+=`\nInventory Purchases (Product Procurement — cash outflow, excluded from Net Profit as it's already counted via COGS),${pl.procurementTotal.toFixed(2)}\n`;
      csv+="\nOTHER INCOME (excludes Amazon/Website — already in Revenue)\nHead,Amount\n";
      Object.entries(pl.otherIncomeByHead).forEach(([h,a])=>csv+=`"${h}",${a.toFixed(2)}\n`);
      csv+=`TOTAL OTHER INCOME,${pl.totalOtherIncome.toFixed(2)}\n`;
      csv+=`\nNET PROFIT / LOSS,${pl.netProfit.toFixed(2)}\n`;
      downloadCsv(`PL_Statement_${label.replace(/\s+/g,"_")}.csv`,csv);
    }

    function PLView({pl,label}){
      return(
        <div>
          <div className="mb-4 p-3 rounded-xl text-xs" style={{backgroundColor:C.bgLight,color:C.lightText}}>
            <strong>How this is calculated (plain English):</strong> Revenue and product cost (COGS) come from your Sales Report — what was actually sold. Gross Profit = Revenue − COGS. Operating Expenses are everything you've logged in Expenses <em>except</em> "Product Procurement" — that's excluded here because it's already counted as COGS against units sold (counting it twice would understate your real profit). Other Income excludes "Income from Amazon/Website" since that revenue is already in Sales Report. Net Profit = Gross Profit − Operating Expenses + Other Income.
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-6">
            <Card><div className="text-xs font-bold uppercase" style={{color:C.lightText}}>Revenue</div><div className="text-xl font-black mt-1" style={{fontFamily:F.display,color:C.zenkyPurple}}>{fmtINR(pl.totalRevenue)}</div></Card>
            <Card><div className="text-xs font-bold uppercase" style={{color:C.lightText}}>COGS</div><div className="text-xl font-black mt-1" style={{fontFamily:F.display,color:C.zenkyOrange}}>{fmtINR(pl.totalCOGS)}</div></Card>
            <Card><div className="text-xs font-bold uppercase" style={{color:C.lightText}}>Gross Profit</div><div className="text-xl font-black mt-1" style={{fontFamily:F.display,color:pl.grossProfit>=0?C.mintGreen:C.zenkyPink}}>{fmtINR(pl.grossProfit)}</div></Card>
            <Card style={{borderColor:pl.netProfit>=0?C.mintGreen:C.zenkyPink}}><div className="text-xs font-bold uppercase" style={{color:C.lightText}}>Net Profit</div><div className="text-xl font-black mt-1" style={{fontFamily:F.display,color:pl.netProfit>=0?C.mintGreen:C.zenkyPink}}>{fmtINR(pl.netProfit)}</div></Card>
          </div>

          <Card className="mb-4">
            <div className="flex items-center justify-between mb-3"><h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>Revenue by Product & Combo</h3><button onClick={()=>exportPL(pl,label)} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export Full P&L</button></div>
            {pl.skuRows.length>0&&<PLSortableTable rows={pl.skuRows.map(r=>({...r,gp:r.revenue-r.cogs}))} title="SKUs"/>}
            {pl.comboRows.length>0&&<PLSortableTable rows={pl.comboRows.map(r=>({...r,gp:r.revenue-r.cogs}))} title="Combos"/>}
            {pl.skuRows.length===0&&pl.comboRows.length===0&&<p className="text-sm" style={{color:C.lightText}}>No sales recorded for this period.</p>}
          </Card>

          <Card className="mb-4">
            <h3 className="font-bold text-lg mb-3" style={{fontFamily:F.display,color:C.darkText}}>Operating Expenses</h3>
            {Object.keys(pl.opExByHead).length===0?<p className="text-sm" style={{color:C.lightText}}>None logged.</p>:(
              <div className="space-y-1">
                {Object.entries(pl.opExByHead).sort((a,b)=>b[1]-a[1]).map(([h,a])=>(
                  <div key={h} className="flex justify-between text-sm py-1 border-t" style={{borderColor:C.border}}><span>{h}</span><span className="font-bold" style={{fontFamily:F.mono,color:C.zenkyOrange}}>{fmtINR(a)}</span></div>
                ))}
                <div className="flex justify-between text-sm pt-2 font-bold" style={{borderTop:`2px solid ${C.border}`}}><span>Total Operating Expenses</span><span style={{fontFamily:F.mono,color:C.zenkyOrange}}>{fmtINR(pl.totalOpEx)}</span></div>
              </div>
            )}
            {pl.procurementTotal>0&&<p className="text-xs mt-3 p-2.5 rounded-lg" style={{backgroundColor:C.bgLight,color:C.lightText}}>Inventory Purchases (Product Procurement): <strong>{fmtINR(pl.procurementTotal)}</strong> — real cash spent buying stock, but not included in Net Profit above since it's already reflected as COGS against units sold.</p>}
          </Card>

          <Card>
            <h3 className="font-bold text-lg mb-3" style={{fontFamily:F.display,color:C.darkText}}>Other Income</h3>
            {Object.keys(pl.otherIncomeByHead).length===0?<p className="text-sm" style={{color:C.lightText}}>None logged.</p>:(
              <div className="space-y-1">
                {Object.entries(pl.otherIncomeByHead).map(([h,a])=>(
                  <div key={h} className="flex justify-between text-sm py-1 border-t" style={{borderColor:C.border}}><span>{h}</span><span className="font-bold" style={{fontFamily:F.mono,color:C.mintGreen}}>{fmtINR(a)}</span></div>
                ))}
              </div>
            )}
          </Card>
        </div>
      );
    }

    // ── Date-wise Ledger (Cash Book style) ──
    function DateWiseLedger(){
      const [fromDate,setFromDate]=useState("");
      const [toDate,setToDate]=useState("");
      const entries=useMemo(()=>{
        const all=[
          ...investments.map(i=>({date:i.date,desc:`Investment — ${i.investorName}`,in:Number(i.amount||0),out:0})),
          ...income.map(i=>({date:i.date,desc:`Income — ${i.head}${i.receivedFrom?` (from ${i.receivedFrom})`:""}`,in:Number(i.amount||0),out:0})),
          ...expenses.map(e=>({date:e.date,desc:`Expense — ${e.head}${e.paidTo?` (to ${e.paidTo})`:""}`,in:0,out:Number(e.amount||0)})),
        ].filter(e=>e.date);
        return all.sort((a,b)=>new Date(a.date)-new Date(b.date));
      },[]);
      const filtered=entries.filter(e=>(!fromDate||e.date>=fromDate)&&(!toDate||e.date<=toDate));
      let running=0;
      const withBalance=filtered.map((e,i)=>{running+=e.in-e.out;return{...e,id:i,balance:running};});
      // Balance is computed chronologically above (as it must be — it's a running
      // total), but the table itself can still be sorted by any column for display;
      // each row keeps the balance value that was true as of its own date.
      const{sorted,sortKey,sortDir,toggleSort}=useSortableRows(withBalance,"date","asc");

      function exportLedger(){
        let csv="Date,Description,Money In,Money Out,Balance\n";
        withBalance.forEach(e=>csv+=`${e.date},"${e.desc}",${e.in.toFixed(2)},${e.out.toFixed(2)},${e.balance.toFixed(2)}\n`);
        downloadCsv("date_wise_ledger.csv",csv);
      }

      return(
        <div>
          <div className="mb-4 p-3 rounded-xl text-xs" style={{backgroundColor:C.bgLight,color:C.lightText}}>
            <strong>What this shows:</strong> every investment, income entry, and expense in date order with a running balance — like a cash book. This tracks actual money in/out as you've logged it, not accrual-based sales revenue (see P&L Statement for that). Click any column heading to sort.
          </div>
          <div className="flex flex-wrap items-end gap-2 mb-4">
            <div><label className="text-xs font-bold uppercase block mb-1" style={{color:C.lightText}}>From</label><Input type="date" value={fromDate} onChange={e=>setFromDate(e.target.value)}/></div>
            <div><label className="text-xs font-bold uppercase block mb-1" style={{color:C.lightText}}>To</label><Input type="date" value={toDate} onChange={e=>setToDate(e.target.value)}/></div>
            {(fromDate||toDate)&&<button onClick={()=>{setFromDate("");setToDate("");}} className="text-xs font-bold" style={{color:C.lightText}}>Clear filter</button>}
            <div className="flex-1"/>
            {withBalance.length>0&&<button onClick={exportLedger} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export</button>}
          </div>
          {withBalance.length===0?<Empty icon={Calendar} title="No entries in this range" message="Add investments, income, or expenses, or widen the date filter."/>:(
            <Card>
              <div className="overflow-x-auto"><table className="w-full text-sm">
                <thead><tr>
                  <SortTH label="Date" sortKey="date" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
                  <SortTH label="Description" sortKey="desc" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
                  <SortTH label="In" sortKey="in" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
                  <SortTH label="Out" sortKey="out" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
                  <SortTH label="Balance" sortKey="balance" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
                </tr></thead>
                <tbody>{sorted.map((e)=>(
                  <tr key={e.id} className="border-t" style={{borderColor:C.border}}>
                    <td className="py-2 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{e.date}</td>
                    <td className="py-2 pr-3">{e.desc}</td>
                    <td className="py-2 pr-3" style={{fontFamily:F.mono,color:e.in?C.mintGreen:C.lightText}}>{e.in?fmtINR(e.in):"—"}</td>
                    <td className="py-2 pr-3" style={{fontFamily:F.mono,color:e.out?C.zenkyOrange:C.lightText}}>{e.out?fmtINR(e.out):"—"}</td>
                    <td className="py-2 pr-3 font-bold" style={{fontFamily:F.mono,color:e.balance>=0?C.zenkyPurple:C.zenkyPink}}>{fmtINR(e.balance)}</td>
                  </tr>
                ))}</tbody>
              </table></div>
            </Card>
          )}
        </div>
      );
    }

    // ── Financial Year-wise ──
    function FinancialYearView(){
      const fyGroups=useMemo(()=>{
        const fys=new Set();
        salesLines.forEach(l=>fys.add(fyLabel(l.date)));
        expenses.forEach(e=>fys.add(fyLabel(e.date)));
        income.forEach(i=>fys.add(fyLabel(i.date)));
        return Array.from(fys).filter(f=>f!=="Unknown").sort();
      },[]);
      if(!fyGroups.length)return<Empty icon={IndianRupee} title="No data yet" message="Add sales, income, or expenses to see financial-year reports."/>;
      return(
        <div>
          <div className="mb-4 p-3 rounded-xl text-xs" style={{backgroundColor:"#FFF3E6",color:"#9a5b0f"}}>
            Indian Financial Year runs 1 April – 31 March. This is for your own tracking — not a substitute for filing with a CA or tax advisor.
          </div>
          <div className="space-y-6">
            {fyGroups.map(fy=>{
              const fyLines=salesLines.filter(l=>fyLabel(l.date)===fy);
              const fyExpenses=expenses.filter(e=>fyLabel(e.date)===fy);
              const fyIncome=income.filter(i=>fyLabel(i.date)===fy);
              const pl=plFor(fyLines,fyExpenses,fyIncome);
              return(
                <div key={fy}>
                  <div className="flex items-center justify-between mb-3">
                    <h3 className="font-black text-xl" style={{fontFamily:F.display,color:C.zenkyPurple}}>{fy}</h3>
                    <button onClick={()=>exportPL(pl,fy)} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export {fy}</button>
                  </div>
                  <PLView pl={pl} label={fy}/>
                </div>
              );
            })}
          </div>
        </div>
      );
    }

    // ── Investor-wise ──
    function InvestorWiseView(){
      const [expandedId,setExpandedId]=useState(null);
      const perInvestor=useMemo(()=>{
        return investors.map(inv=>{
          const invTx=investments.filter(t=>t.investorId===inv.id);
          const coveredExpenses=expenses.filter(e=>e.spentBy===inv.id);
          const totalInvested=invTx.reduce((s,t)=>s+Number(t.amount||0),0);
          const totalCovered=coveredExpenses.reduce((s,e)=>s+Number(e.amount||0),0);
          // Timeline comes ONLY from investments — this already includes both
          // direct transfers and auto-linked "covered expense" entries (v5.5).
          // Do NOT also list the underlying expense record here: that would
          // show the same money twice (once as the investment credit, once as
          // the expense debit) and double the investor's displayed total.
          const timeline=invTx.map(t=>({date:t.date,desc:`Investment${t.fromExpenseId?" (covered expense, auto-logged)":""}`,amount:t.amount,paymentMode:t.paymentMode,comment:t.comment}))
            .sort((a,b)=>new Date(a.date)-new Date(b.date));
          return{...inv,totalInvested,totalCovered,combined:totalInvested,txCount:timeline.length,timeline}; // note: auto-linked investments already fold covered-expense amounts into totalInvested, so "combined" = totalInvested (no separate addition needed, avoids double counting)
        }).sort((a,b)=>b.combined-a.combined);
      },[]);

      const{sorted,sortKey,sortDir,toggleSort}=useSortableRows(perInvestor,"combined","desc");

      function exportInvestor(inv){
        let csv=`Investor Report — ${inv.name}\n\nDate,Description,Amount,Payment Mode,Comment\n`;
        inv.timeline.forEach(t=>csv+=`${t.date},"${t.desc}",${Number(t.amount||0).toFixed(2)},${t.paymentMode||""},"${t.comment||""}"\n`);
        csv+=`\nTotal Invested (includes expenses they covered),${inv.totalInvested.toFixed(2)}\n`;
        downloadCsv(`investor_report_${inv.name.replace(/\s+/g,"_")}.csv`,csv);
      }
      function exportAllInvestors(){
        let csv="Name,Contact,Total Invested,Expenses Covered (₹),# Transactions\n";
        sorted.forEach(inv=>csv+=`"${inv.name}","${inv.contact||""}",${inv.totalInvested.toFixed(2)},${inv.totalCovered.toFixed(2)},${inv.timeline.length}\n`);
        downloadCsv("investors_summary.csv",csv);
      }

      if(!investors.length)return<Empty icon={Users} title="No investors yet" message="Add investors in the Investors tab to see per-investor reports."/>;

      return(
        <div>
          <div className="mb-4 p-3 rounded-xl text-xs" style={{backgroundColor:C.bgLight,color:C.lightText}}>
            <strong>What this shows:</strong> for each investor, their total contribution — direct fund transfers plus any expenses they personally covered (which already count toward their investment total, so nothing here is double-counted). Click a row to see their full transaction history.
          </div>
          <div className="flex justify-end mb-3"><button onClick={exportAllInvestors} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export Summary (All Investors)</button></div>
          <Card>
            <div className="overflow-x-auto"><table className="w-full text-sm">
              <thead><tr>
                <SortTH label="Name" sortKey="name" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
                <SortTH label="Total Invested" sortKey="totalInvested" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
                <SortTH label="Of Which: Covered Expenses" sortKey="totalCovered" activeKey={sortKey} dir={sortDir} onClick={toggleSort}/>
                <SortTH label="Transactions" sortKey="txCount" activeKey={sortKey} dir={sortDir} onClick={toggleSort} className="hidden sm:table-cell"/>
                <th/>
              </tr></thead>
              <tbody>{sorted.map(inv=>{
                const isOpen=expandedId===inv.id;
                return(
                  <Fragment key={inv.id}>
                    <tr className="border-t cursor-pointer" style={{borderColor:C.border}} onClick={()=>setExpandedId(isOpen?null:inv.id)}>
                      <td className="py-2 pr-3 font-bold"><span className="inline-flex items-center gap-1.5">{isOpen?<ChevronDown size={14}/>:<ChevronRight size={14}/>}{inv.name}</span></td>
                      <td className="py-2 pr-3 font-bold" style={{fontFamily:F.mono,color:C.zenkyPurple}}>{fmtINR(inv.totalInvested)}</td>
                      <td className="py-2 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{inv.totalCovered>0?fmtINR(inv.totalCovered):"—"}</td>
                      <td className="py-2 pr-3 hidden sm:table-cell" style={{color:C.lightText}}>{inv.timeline.length}</td>
                      <td className="py-2 pr-3 text-right"><button onClick={ev=>{ev.stopPropagation();exportInvestor(inv);}} className="text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/></button></td>
                    </tr>
                    {isOpen&&(
                      <tr><td colSpan={5} className="pb-3">
                        <div className="rounded-xl p-3" style={{backgroundColor:C.bgLight}}>
                          {inv.timeline.length===0?<p className="text-xs" style={{color:C.lightText}}>No transactions yet.</p>:(
                            <table className="w-full text-xs">
                              <thead><tr style={{color:C.lightText}}><th className="py-1 pr-3 text-left font-bold uppercase">Date</th><th className="py-1 pr-3 text-left font-bold uppercase">Description</th><th className="py-1 pr-3 text-left font-bold uppercase">Amount</th><th className="py-1 pr-3 text-left font-bold uppercase">Comment</th></tr></thead>
                              <tbody>{inv.timeline.map((t,i)=>(
                                <tr key={i} className="border-t" style={{borderColor:C.border}}>
                                  <td className="py-1.5 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{t.date}</td>
                                  <td className="py-1.5 pr-3">{t.desc}</td>
                                  <td className="py-1.5 pr-3 font-bold" style={{fontFamily:F.mono,color:C.zenkyPurple}}>{fmtINR(t.amount)}</td>
                                  <td className="py-1.5 pr-3" style={{color:C.lightText}}>{t.comment||"—"}</td>
                                </tr>
                              ))}</tbody>
                            </table>
                          )}
                        </div>
                      </td></tr>
                    )}
                  </Fragment>
                );
              })}</tbody>
            </table></div>
          </Card>
        </div>
      );
    }

    // ── Monthly (existing) ──
    const monthly=useMemo(()=>{
      const groups={};
      const add=(date,amount,type,head)=>{
        const key=monthKey(date);
        if(!groups[key])groups[key]={key,income:0,expense:0,byHead:{}};
        groups[key][type]+=amount;
        const hk=`${type}:${head}`;
        groups[key].byHead[hk]=(groups[key].byHead[hk]||0)+amount;
      };
      income.forEach(i=>add(i.date,Number(i.amount||0),"income",i.head));
      investments.forEach(i=>add(i.date,Number(i.amount||0),"income",`Investment — ${i.investorName}`));
      expenses.forEach(e=>add(e.date,Number(e.amount||0),"expense",e.head));
      return Object.values(groups).sort((a,b)=>new Date(a.key+" 1")-new Date(b.key+" 1"));
    },[]);

    const byPerson=useMemo(()=>{
      const totals={};
      expenses.forEach(e=>{
        const key=e.spentByName||"Unattributed";
        if(!totals[key])totals[key]={name:key,count:0,amount:0};
        totals[key].count++;totals[key].amount+=Number(e.amount||0);
      });
      return Object.values(totals).sort((a,b)=>b.amount-a.amount);
    },[]);

    function exportMonthly(){
      let csv="Month,Total Income,Total Expense,Net\n";
      monthly.forEach(m=>csv+=`${m.key},${m.income.toFixed(2)},${m.expense.toFixed(2)},${(m.income-m.expense).toFixed(2)}\n`);
      csv+="\nMonth,Type,Head,Amount\n";
      monthly.forEach(m=>Object.entries(m.byHead).forEach(([hk,amt])=>{const[type,head]=hk.split(":");csv+=`${m.key},${type},"${head}",${amt.toFixed(2)}\n`;}));
      downloadCsv("financial_monthly_report.csv",csv);
    }

    function MonthlyView(){
      if(!monthly.length)return<Empty icon={IndianRupee} title="No data yet" message="Add income, investments, or expenses to see monthly reports."/>;
      return(
        <div>
          {byPerson.length>0&&(
            <Card className="mb-6">
              <h3 className="font-bold text-lg mb-4" style={{fontFamily:F.display,color:C.darkText}}>Accountability — Expenses by Person</h3>
              <div className="space-y-2">
                {byPerson.map(p=>(
                  <div key={p.name} className="flex items-center justify-between p-3 rounded-xl" style={{backgroundColor:C.bgLight}}>
                    <div className="flex items-center gap-2">
                      <Stamp tone={p.name==="Unattributed"?"pink":"purple"}>{p.name}</Stamp>
                      <span className="text-xs" style={{color:C.lightText}}>{p.count} expense{p.count!==1?"s":""}</span>
                    </div>
                    <span className="font-bold" style={{fontFamily:F.mono,color:C.zenkyOrange}}>{fmtINR(p.amount)}</span>
                  </div>
                ))}
              </div>
              {byPerson.some(p=>p.name==="Unattributed")&&<p className="text-xs mt-3" style={{color:C.lightText}}>"Unattributed" expenses don't have a "Spent By" person recorded — edit them in the Expenses tab to assign one.</p>}
            </Card>
          )}
          <div className="flex justify-end mb-4"><button onClick={exportMonthly} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export Monthly Report</button></div>
          <div className="space-y-4">
            {monthly.map(m=>(
              <Card key={m.key}>
                <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
                  <div className="font-bold text-lg" style={{fontFamily:F.display,color:C.zenkyPurple}}>{m.key}</div>
                  <div className="flex items-center gap-4 text-right">
                    <div><div className="text-xs" style={{color:C.lightText}}>Income</div><div className="font-bold" style={{fontFamily:F.mono,color:C.mintGreen}}>{fmtINR(m.income)}</div></div>
                    <div><div className="text-xs" style={{color:C.lightText}}>Expense</div><div className="font-bold" style={{fontFamily:F.mono,color:C.zenkyOrange}}>{fmtINR(m.expense)}</div></div>
                    <div><div className="text-xs" style={{color:C.lightText}}>Net</div><div className="font-bold" style={{fontFamily:F.mono,color:m.income-m.expense>=0?C.mintGreen:C.zenkyPink}}>{fmtINR(m.income-m.expense)}</div></div>
                  </div>
                </div>
                <div className="space-y-1">
                  {Object.entries(m.byHead).sort((a,b)=>b[1]-a[1]).map(([hk,amt])=>{
                    const[type,head]=hk.split(":");
                    return(
                      <div key={hk} className="flex justify-between text-xs py-1 border-t" style={{borderColor:C.border}}>
                        <span style={{color:C.darkText}}>{head}</span>
                        <span className="font-bold" style={{fontFamily:F.mono,color:type==="income"?C.mintGreen:C.zenkyOrange}}>{type==="income"?"+":"−"}{fmtINR(amt)}</span>
                      </div>
                    );
                  })}
                </div>
              </Card>
            ))}
          </div>
        </div>
      );
    }

    return(
      <div>
        <div className="flex gap-1.5 mb-6 overflow-x-auto pb-1">
          {REPORT_TABS.map(t=>(
            <button key={t.id} onClick={()=>setReportTab(t.id)} className="px-3.5 py-2 rounded-full text-xs font-bold whitespace-nowrap transition-colors" style={{backgroundColor:reportTab===t.id?C.zenkyPurple:C.softWhite,color:reportTab===t.id?C.softWhite:C.darkText,border:`2px solid ${reportTab===t.id?C.zenkyPurple:C.border}`,fontFamily:F.display}}>
              {t.label}
            </button>
          ))}
        </div>
        {reportTab==="pl"&&<PLView pl={overallPL} label="Overall (All Time)"/>}
        {reportTab==="datewise"&&<DateWiseLedger/>}
        {reportTab==="fy"&&<FinancialYearView/>}
        {reportTab==="investorwise"&&<InvestorWiseView/>}
        {reportTab==="monthly"&&<MonthlyView/>}
      </div>
    );
  }

  // ── Investor Statement — a dedicated, presentable, branded page shown when an investor's name is clicked ──
  function InvestorStatementView(){
    const inv=investors.find(i=>i.id===viewingInvestorId);
    if(!inv)return null;
    // Timeline comes ONLY from investments — this already includes both direct
    // transfers and auto-linked "covered expense" entries (v5.5). The underlying
    // expense record is deliberately NOT also listed here — that would show the
    // same money twice (once as investment credit, once as expense debit) and
    // double the investor's Total Invested figure.
    const timeline=useMemo(()=>{
      return investments.filter(t=>t.investorId===inv.id)
        .map(t=>({date:t.date,desc:`Investment${t.fromExpenseId?" (covered expense, auto-logged)":""}`,paymentMode:t.paymentMode,comment:t.comment,amount:Number(t.amount||0)}))
        .sort((a,b)=>new Date(a.date)-new Date(b.date));
    },[]);
    const total=timeline.reduce((s,t)=>s+t.amount,0);

    function exportStatement(){
      let csv=`ZenkyBox — Investor Statement\nInvestor: ${inv.name}\nGenerated: ${new Date().toLocaleDateString("en-IN")}\n\n`;
      csv+="Date,Description,Payment Mode,Amount,Comment\n";
      timeline.forEach(t=>csv+=`${t.date},"${t.desc}",${t.paymentMode||""},${t.amount.toFixed(2)},"${t.comment||""}"\n`);
      csv+=`\nTotal Invested,${total.toFixed(2)}\n`;
      downloadCsv(`ZenkyBox_Investor_Statement_${inv.name.replace(/\s+/g,"_")}.csv`,csv);
      logActivity?.("Investor statement exported",inv.name);
    }

    return(
      <div>
        <button onClick={()=>setViewingInvestorId(null)} className="inline-flex items-center gap-1 text-sm font-bold mb-5" style={{color:C.lightText}}><ChevronRight size={14} style={{transform:"rotate(180deg)"}}/>Back to Investors</button>

        <div className="rounded-2xl border-2 overflow-hidden" style={{borderColor:C.border,backgroundColor:C.softWhite}}>
          {/* Branded header */}
          <div className="p-6 sm:p-8 text-center" style={{backgroundColor:C.bgLight,borderBottom:`2px solid ${C.border}`}}>
            <img src="/zenkybox-wordmark.png" alt="ZenkyBox" style={{height:"36px",margin:"0 auto 12px"}}/>
            <div className="text-xs font-bold uppercase" style={{color:C.lightText,letterSpacing:"0.08em"}}>zenkybox.in · Investor Statement</div>
          </div>

          <div className="p-6 sm:p-8">
            <div className="flex items-start justify-between flex-wrap gap-4 mb-6">
              <div>
                <div className="text-2xl font-black" style={{fontFamily:F.display,color:C.darkText}}>{inv.name}</div>
                {inv.contact&&<div className="text-sm mt-1" style={{color:C.lightText}}>{inv.contact}</div>}
                {inv.notes&&<div className="text-sm mt-0.5" style={{color:C.lightText}}>{inv.notes}</div>}
              </div>
              <div className="text-right">
                <div className="text-xs font-bold uppercase" style={{color:C.lightText}}>Statement Date</div>
                <div className="text-sm font-bold" style={{color:C.darkText}}>{new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"long",year:"numeric"})}</div>
              </div>
            </div>

            <div className="rounded-xl p-4 mb-6 flex items-center justify-between" style={{backgroundColor:C.bgLight}}>
              <span className="font-bold" style={{fontFamily:F.display,color:C.darkText}}>Total Invested</span>
              <span className="text-2xl font-black" style={{fontFamily:F.display,color:C.zenkyPurple}}>{fmtINR(total)}</span>
            </div>

            <div className="flex items-center justify-between mb-3">
              <h3 className="font-bold text-lg" style={{fontFamily:F.display,color:C.darkText}}>Date-wise Investment History</h3>
              <button onClick={exportStatement} className="inline-flex items-center gap-1.5 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export Statement</button>
            </div>

            {timeline.length===0?(
              <Empty icon={IndianRupee} title="No transactions yet" message="Log an investment for this person, or check back once they've covered an expense."/>
            ):(
              <div className="overflow-x-auto"><table className="w-full text-sm">
                <thead><tr style={{color:C.lightText}}>
                  <th className="py-2 pr-3 text-left text-xs uppercase font-bold">Date</th>
                  <th className="py-2 pr-3 text-left text-xs uppercase font-bold">Description</th>
                  <th className="py-2 pr-3 text-left text-xs uppercase font-bold hidden sm:table-cell">Mode</th>
                  <th className="py-2 pr-3 text-left text-xs uppercase font-bold">Amount</th>
                  <th className="py-2 pr-3 text-left text-xs uppercase font-bold hidden sm:table-cell">Comment</th>
                </tr></thead>
                <tbody>{timeline.map((t,i)=>(
                  <tr key={i} className="border-t" style={{borderColor:C.border}}>
                    <td className="py-2.5 pr-3" style={{fontFamily:F.mono,color:C.lightText}}>{t.date}</td>
                    <td className="py-2.5 pr-3">{t.desc}</td>
                    <td className="py-2.5 pr-3 hidden sm:table-cell">{t.paymentMode?<Stamp tone="blue">{t.paymentMode}</Stamp>:"—"}</td>
                    <td className="py-2.5 pr-3 font-bold" style={{fontFamily:F.mono,color:C.zenkyPurple}}>{fmtINR(t.amount)}</td>
                    <td className="py-2.5 pr-3 hidden sm:table-cell" style={{color:C.lightText}}>{t.comment||"—"}</td>
                  </tr>
                ))}</tbody>
                <tfoot><tr style={{borderTop:`2px solid ${C.border}`}}>
                  <td colSpan={3} className="py-3 pr-3 font-bold hidden sm:table-cell" style={{fontFamily:F.display,color:C.darkText}}>Total</td>
                  <td colSpan={2} className="py-3 pr-3 font-bold sm:hidden" style={{fontFamily:F.display,color:C.darkText}}>Total</td>
                  <td className="py-3 pr-3 font-black" style={{fontFamily:F.mono,color:C.zenkyPurple}}>{fmtINR(total)}</td>
                  <td className="hidden sm:table-cell"/>
                </tr></tfoot>
              </table></div>
            )}
          </div>

          <div className="px-6 sm:px-8 py-4 text-center text-xs" style={{backgroundColor:C.bgLight,color:C.lightText,borderTop:`2px solid ${C.border}`}}>
            💝 Thoughtful Gifts. Joyful Moments. — ZenkyBox.in
          </div>
        </div>
      </div>
    );
  }

  return(
    <div>
      <SectionHeader title="Financials" subtitle="Investors, expenses, income, and overall fund balance."/>
      {viewingInvestorId?<InvestorStatementView/>:(
        <>
          <div className="flex gap-1.5 mb-6 overflow-x-auto pb-1">
            {TABS.map(t=>(
              <button key={t.id} onClick={()=>setTab(t.id)} className="px-3.5 py-2 rounded-full text-xs font-bold whitespace-nowrap transition-colors" style={{backgroundColor:tab===t.id?C.zenkyPurple:C.softWhite,color:tab===t.id?C.softWhite:C.darkText,border:`2px solid ${tab===t.id?C.zenkyPurple:C.border}`,fontFamily:F.display}}>
                {t.label}
              </button>
            ))}
          </div>
          {tab==="overview"&&<Overview/>}
          {tab==="investors"&&<Investors/>}
          {tab==="expenses"&&<Expenses/>}
          {tab==="income"&&<Income/>}
          {tab==="reports"&&<Reports/>}
        </>
      )}
    </div>
  );
}

export default function App(){
  const [skus,setSkus]=useState([]);
  const [combos,setCombos]=useState([]);
  const [reports,setReports]=useState([]);
  const [salesLines,setSalesLines]=useState([]);
  const [activityLog,setActivityLog]=useState([]);
  const [adminPin,setAdminPinState]=useState("");
  const [loginCreds,setLoginCredsState]=useState(null); // master/owner login — always admin-capable, acts as bootstrap account
  const [users,setUsers]=useState([]); // named team members: [{id,username,password,name,canBeAdmin}]
  const [currentUser,setCurrentUser]=useState(null); // {username,name,canBeAdmin,isMaster}
  const [investors,setInvestors]=useState([]);
  const [investments,setInvestments]=useState([]);
  const [expenses,setExpenses]=useState([]);
  const [income,setIncome]=useState([]);
  const [loaded,setLoaded]=useState(false);
  const [canSave,setCanSave]=useState(false); // only true once a load has genuinely succeeded — prevents overwriting real data with an empty state after a failed fetch
  const [loadError,setLoadError]=useState(false);
  const [view,setView]=useState("dashboard");
  const [sidebarOpen,setSidebarOpen]=useState(false);
  const [toast,setToast]=useState(null);
  const [synced,setSynced]=useState(false);
  const [role,setRole]=useState("staff"); // per-device only, not shared across users
  const [isLoggedIn,setIsLoggedIn]=useState(false); // per-device/session app-wide login gate

  // Tracks the newest updated_at timestamp we're aware of — from our own last
  // save OR the last realtime update we accepted. Any incoming realtime event
  // that isn't strictly newer than this gets ignored, since it's either an
  // echo of our own write or a stale/out-of-order delivery. This is what
  // prevents "add a SKU, count flickers back down" — a real race that existed
  // because Supabase broadcasts writes back to the client that made them.
  const lastKnownUpdatedAt=useRef(null);

  function applyLoadedData(data){
    setSkus(data?.skus||[]);setCombos(data?.combos||[]);setReports(data?.reports||[]);
    setSalesLines(data?.salesLines||[]);setActivityLog(data?.activityLog||[]);
    setAdminPinState(data?.adminPin||"");
    setLoginCredsState(data?.loginCreds||null);
    setUsers(data?.users||[]);
    setInvestors(data?.investors||[]);setInvestments(data?.investments||[]);
    setExpenses(data?.expenses||[]);setIncome(data?.income||[]);
  }

  // Load initial data — retries a couple of times on failure before giving up,
  // since the most common failure window is a fresh deploy's cold start.
  useEffect(()=>{
    let cancelled=false;
    (async()=>{
      let attempt=0;
      let succeeded=false;
      while(attempt<3&&!cancelled){
        const result=await loadData();
        if(result.ok){
          applyLoadedData(result.data);
          if(result.updatedAt)lastKnownUpdatedAt.current=result.updatedAt;
          setSynced(hasSync());
          setCanSave(true); // safe to save now — we know the true remote/local state
          setLoadError(false);
          succeeded=true;
          break;
        }
        attempt++;
        if(attempt<3)await new Promise(res=>setTimeout(res,800*attempt)); // brief backoff before retry
      }
      if(!cancelled&&!succeeded){
        // All retries failed — do NOT enable saving. Show a visible warning instead
        // of silently risking an overwrite of real remote data.
        setLoadError(true);
        setSynced(hasSync());
      }
      if(typeof window!=="undefined"&&sessionStorage.getItem("zenkybox-role")==="admin")setRole("admin");
      if(typeof window!=="undefined"&&sessionStorage.getItem("zenkybox-logged-in")==="1"){
        setIsLoggedIn(true);
        try{const savedUser=sessionStorage.getItem("zenkybox-current-user");if(savedUser)setCurrentUser(JSON.parse(savedUser));}catch(e){}
      }
      if(!cancelled)setLoaded(true);
    })();
    return()=>{cancelled=true;};
    // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  // Subscribe to real-time updates (Supabase) — only apply genuinely newer changes
  useEffect(()=>{
    let unsub=null;
    (async()=>{
      unsub=await subscribeToData((data,updatedAt)=>{
        if(!data)return;
        if(updatedAt&&lastKnownUpdatedAt.current&&new Date(updatedAt)<=new Date(lastKnownUpdatedAt.current)){
          return; // stale or self-echo — we already have this or something newer
        }
        if(updatedAt)lastKnownUpdatedAt.current=updatedAt;
        applyLoadedData(data);
      });
    })();
    return()=>{if(unsub)unsub();};
  },[]);

  // Save whenever data changes — gated on canSave, so a failed initial load
  // can never result in writing an empty catalog over real remote data.
  const saveTimeout=useRef(null);
  useEffect(()=>{
    if(!loaded||!canSave)return;
    clearTimeout(saveTimeout.current);
    saveTimeout.current=setTimeout(()=>{
      saveData({skus,combos,reports,salesLines,activityLog,adminPin,loginCreds,users,investors,investments,expenses,income})
        .then(timestamp=>{if(timestamp)lastKnownUpdatedAt.current=timestamp;})
        .catch(()=>{});
    },500);
    return()=>clearTimeout(saveTimeout.current);
  },[skus,combos,reports,salesLines,activityLog,adminPin,loginCreds,users,investors,investments,expenses,income,loaded,canSave]);

  // For destructive/corrective actions (Flush, Cleanup, Clear All, Replace All) —
  // caller supplies the COMPLETE new payload explicitly (not read from state,
  // since state setters called moments earlier haven't re-rendered yet) and this
  // writes it to Supabase immediately, skipping the 500ms debounce entirely.
  // This closes the race window where a stale tab could resave old data before
  // a flush/cleanup's result reaches the database.
  function forceSaveNow(payload){
    clearTimeout(saveTimeout.current);
    saveData(payload).then(timestamp=>{if(timestamp)lastKnownUpdatedAt.current=timestamp;}).catch(()=>{});
  }

  function retryLoad(){
    setLoadError(false);setLoaded(false);
    (async()=>{
      const result=await loadData();
      if(result.ok){
        applyLoadedData(result.data);
        if(result.updatedAt)lastKnownUpdatedAt.current=result.updatedAt;
        setCanSave(true);setLoadError(false);
      }else setLoadError(true);
      setLoaded(true);
    })();
  }

  function showToast(type,msg){setToast({type,msg});setTimeout(()=>setToast(null),3500);}

  function logActivity(action,detail){
    setActivityLog(prev=>[{id:Date.now().toString()+Math.random(),date:new Date().toISOString(),action,detail:detail||"",role},...prev].slice(0,300));
  }

  function handleUnlock(pin){
    const correct=adminPin||DEFAULT_ADMIN_PIN;
    if(pin===correct){
      setRole("admin");
      if(typeof window!=="undefined")sessionStorage.setItem("zenkybox-role","admin");
      showToast("success","Admin unlocked. 🔓");
      return true;
    }
    showToast("error","Incorrect PIN.");
    return false;
  }
  function handleLock(){
    setRole("staff");
    if(typeof window!=="undefined")sessionStorage.removeItem("zenkybox-role");
    showToast("success","Locked to staff mode.");
  }
  function setAdminPin(pin){setAdminPinState(pin);}
  function setLoginCreds(creds){setLoginCredsState(creds);}

  function handleLogin(username,password){
    const master=loginCreds||DEFAULT_LOGIN;
    if(username===master.username&&password===master.password){
      setIsLoggedIn(true);
      const user={username,name:"Owner",canBeAdmin:true,isMaster:true};
      setCurrentUser(user);
      if(typeof window!=="undefined"){
        sessionStorage.setItem("zenkybox-logged-in","1");
        sessionStorage.setItem("zenkybox-current-user",JSON.stringify(user));
      }
      return true;
    }
    const matched=users.find(u=>u.username===username&&u.password===password);
    if(matched){
      setIsLoggedIn(true);
      const user={username:matched.username,name:matched.name||matched.username,canBeAdmin:!!matched.canBeAdmin,isMaster:false};
      setCurrentUser(user);
      if(typeof window!=="undefined"){
        sessionStorage.setItem("zenkybox-logged-in","1");
        sessionStorage.setItem("zenkybox-current-user",JSON.stringify(user));
      }
      return true;
    }
    return false;
  }
  function handleLogout(){
    setIsLoggedIn(false);
    setCurrentUser(null);
    setRole("staff");
    if(typeof window!=="undefined"){
      sessionStorage.removeItem("zenkybox-logged-in");
      sessionStorage.removeItem("zenkybox-role");
      sessionStorage.removeItem("zenkybox-current-user");
    }
  }

  const skuMap=useMemo(()=>Object.fromEntries(skus.map(s=>[s.sku,s])),[skus]);
  const comboList=useMemo(()=>combos.map(c=>({...c,...comboReadiness(c,skuMap)})),[combos,skuMap]);

  // Defense-in-depth: if a non-admin somehow lands on an admin-only view, bounce to Dashboard
  useEffect(()=>{
    const item=NAV.find(n=>n.id===view);
    if(item?.adminOnly&&role!=="admin")setView("dashboard");
  },[view,role]);

  if(!loaded)return(
    <div className="flex items-center justify-center h-screen" style={{backgroundColor:C.softWhite}}>
      <div className="text-center"><div style={{fontFamily:F.display,color:C.zenkyPurple,fontSize:"24px",fontWeight:"black"}}>✨ ZenkyBox</div><div className="text-sm mt-2" style={{color:C.lightText}}>Loading inventory…</div></div>
    </div>
  );

  if(loadError)return(
    <div className="flex items-center justify-center h-screen p-6" style={{backgroundColor:C.softWhite}}>
      <div className="max-w-md text-center">
        <AlertTriangle size={40} style={{color:"#dc2626",margin:"0 auto"}}/>
        <div className="mt-3 font-black text-xl" style={{fontFamily:F.display,color:C.darkText}}>Couldn't load your synced data</div>
        <p className="text-sm mt-2" style={{color:C.lightText,fontFamily:F.body}}>
          The connection to your database failed just now. To protect your data, ZenkyBox refuses to show or save anything until this is confirmed working — this prevents an empty state from accidentally overwriting your real catalog.
        </p>
        <p className="text-xs mt-3" style={{color:C.lightText}}>This is usually temporary (a brief network hiccup). Try again in a moment.</p>
        <button onClick={retryLoad} className="mt-4 inline-flex items-center gap-2 px-5 py-2.5 rounded-full text-sm font-bold" style={{backgroundColor:C.zenkyPurple,color:C.softWhite,fontFamily:F.display}}>
          <RefreshCw size={15}/>Try Again
        </button>
      </div>
    </div>
  );

  if(!isLoggedIn)return<LoginScreen onLogin={handleLogin}/>;

  const currentNavItem=NAV.find(n=>n.id===view);
  const blocked=currentNavItem?.adminOnly&&role!=="admin";

  return(
    <div className="flex h-screen overflow-hidden" style={{backgroundColor:C.bgLight,fontFamily:F.body}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fredoka:wght@400;500;600;700&family=Baloo+2:wght@500;600;700&family=Nunito:wght@400;500;600;700&display=swap');`}</style>
      <Sidebar view={view} setView={setView} open={sidebarOpen} setOpen={setSidebarOpen} synced={synced} role={role} canBeAdmin={currentUser?.canBeAdmin??true} currentUserName={currentUser?.name} onUnlock={handleUnlock} onLock={handleLock} onLogout={handleLogout}/>
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
        {/* Mobile header */}
        <div className="md:hidden flex items-center justify-between px-4 py-3 flex-shrink-0" style={{backgroundColor:C.zenkyPurple}}>
          <span className="font-black text-lg" style={{fontFamily:F.display,color:C.softWhite}}>ZenkyBox</span>
          <button onClick={()=>setSidebarOpen(true)} style={{color:C.softWhite}}><Menu size={22}/></button>
        </div>
        <main className="flex-1 overflow-y-auto p-4 md:p-8 max-w-6xl mx-auto w-full">
          {blocked?(
            <Empty icon={Lock} title="Admin access required" message="Unlock admin mode from the sidebar to view this section."/>
          ):(
            <>
              {view==="dashboard"&&<Dashboard skus={skus} comboList={comboList}/>}
              {view==="combo-readiness"&&<ComboReadinessView skus={skus} combos={combos}/>}
              {view==="bulk-import"&&<BulkImportView skus={skus} combos={combos} setSkus={setSkus} setCombos={setCombos} showToast={showToast} logActivity={logActivity}/>}
              {view==="catalog"&&<Catalog skus={skus} setSkus={setSkus} showToast={showToast} role={role} logActivity={logActivity}/>}
              {view==="combos"&&<CombosView skus={skus} combos={combos} setCombos={setCombos} showToast={showToast} role={role} logActivity={logActivity}/>}
              {view==="upload"&&<UploadView skus={skus} combos={combos} setSkus={setSkus} reports={reports} setReports={setReports} salesLines={salesLines} setSalesLines={setSalesLines} logActivity={logActivity} showToast={showToast}/>}
              {view==="reports"&&<ReportsView reports={reports} skus={skus} combos={combos}/>}
              {view==="sales-reports"&&<SalesReportsView salesLines={salesLines} skus={skus} combos={combos}/>}
              {view==="costing"&&<CostingPricingView skus={skus}/>}
              {view==="financials"&&<FinancialsView investors={investors} setInvestors={setInvestors} investments={investments} setInvestments={setInvestments} expenses={expenses} setExpenses={setExpenses} income={income} setIncome={setIncome} salesLines={salesLines} skus={skus} combos={combos} reports={reports} activityLog={activityLog} adminPin={adminPin} loginCreds={loginCreds} forceSaveNow={forceSaveNow} logActivity={logActivity} showToast={showToast}/>}
              {view==="source-data"&&<SourceDataView activityLog={activityLog} synced={synced} salesLines={salesLines} setSalesLines={setSalesLines} reports={reports} setReports={setReports} skus={skus} setSkus={setSkus} combos={combos} adminPin={adminPin} loginCreds={loginCreds} investors={investors} investments={investments} expenses={expenses} income={income} forceSaveNow={forceSaveNow} logActivity={logActivity} showToast={showToast}/>}
              {view==="access"&&<AccessManagementView role={role} adminPin={adminPin} setAdminPin={setAdminPin} loginCreds={loginCreds} setLoginCreds={setLoginCreds} users={users} setUsers={setUsers} showToast={showToast} logActivity={logActivity}/>}
            </>
          )}
        </main>
      </div>
      {toast&&<div className="fixed bottom-4 right-4 px-5 py-3 rounded-full text-sm font-bold shadow-lg z-50" style={{backgroundColor:toast.type==="error"?C.zenkyPink:C.mintGreen,color:"#fff",fontFamily:F.display}}>{toast.msg}</div>}
    </div>
  );
}
