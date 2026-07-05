import { useState, useEffect, useMemo, useRef, useCallback } from "react";
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
    cityKey:find(/ship.?city/i),
    stateKey:find(/ship.?state|ship.?region/i),
    priceKey:find(/item.?price|unit.?price|^price$/i),
  };
}

function aggregateSalesLines(lines,groupFn,skuMap,comboMap){
  const groups={};
  lines.forEach(l=>{
    const key=groupFn(l);
    if(!groups[key])groups[key]={key,qty:0,revenue:0,cost:0,earning:0,lines:[]};
    groups[key].qty+=l.qty;groups[key].revenue+=l.revenue;groups[key].cost+=l.cost;groups[key].earning+=l.earning;groups[key].lines.push(l);
  });
  return Object.values(groups).sort((a,b)=>b.revenue-a.revenue);
}

/* ═══ SHARED UI ═══ */
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
function Sidebar({view,setView,open,setOpen,synced,role,onUnlock,onLock}){
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
          ):showPinBox?(
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
      // Replace mode: wipe existing data, use only what's in the file
      setSkus(preview.skus);
      setCombos(preview.combos.map(c=>({...c,id:c.id||Date.now().toString()+Math.random()})));
      setStage("imported");
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
  function reset(){setStage("idle");setPreview({skus:[],combos:[]});setFileName("");if(ref.current)ref.current.value="";}
  return(
    <div>
      <SectionHeader title="Bulk Import" subtitle="Upload SKU & combo catalog from Excel or CSV."/>
      <Card className="mb-5"><h3 className="font-bold mb-3" style={{fontFamily:F.display,color:C.darkText}}>📋 Required Format</h3>
        <div className="space-y-2 text-sm" style={{fontFamily:F.body,color:C.darkText}}>
          <div><p className="font-bold mb-1">SKUs sheet:</p><code className="block bg-gray-100 p-2 rounded text-xs" style={{fontFamily:F.mono}}>SKU | Name | Stock | Reorder Level | Procurement Cost</code></div>
          <div><p className="font-bold mb-1">Combos sheet:</p><code className="block bg-gray-100 p-2 rounded text-xs" style={{fontFamily:F.mono}}>Combo Code | Combo Name | Components (SKU001:2;SKU002:1)</code></div>
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
              <button onClick={()=>setImportMode("merge")} className="p-3 rounded-xl border-2 text-left transition-all" style={{borderColor:importMode==="merge"?C.zenkyPurple:C.border,backgroundColor:importMode==="merge"?"#F3EEFF":C.softWhite}}>
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

          <div className="flex gap-2"><PrimaryButton onClick={confirmImport} tone={importMode==="replace"?"pink":undefined}><Check size={15}/>{importMode==="replace"?"Replace & Import":"Confirm Import"}</PrimaryButton><button onClick={reset} className="text-sm font-bold" style={{color:C.lightText}}>Cancel</button></div></div>)}
        {stage==="imported"&&(<div className="text-center py-8"><Check size={32} className="mx-auto mb-3" style={{color:C.mintGreen}}/><p className="font-bold text-lg" style={{color:C.darkText,fontFamily:F.display}}>Import complete!</p><p className="text-sm mt-1" style={{color:C.lightText}}>{preview.skus.length} SKUs & {preview.combos.length} combos added.</p><div className="mt-5"><PrimaryButton onClick={reset}><Upload size={15}/>Import another</PrimaryButton></div></div>)}
      </Card>
    </div>
  );
}

/* ═══ UPLOAD SALES ═══ */
function UploadView({skus,combos,setSkus,reports,setReports,salesLines,setSalesLines,logActivity,showToast}){
  const [stage,setStage]=useState("idle");const [fileName,setFileName]=useState("");
  const [aggregated,setAggregated]=useState([]);const [rawRows,setRawRows]=useState([]);const [extraCols,setExtraCols]=useState({});
  const [weekLabel,setWeekLabel]=useState("");const ref=useRef(null);
  const skuMap=useMemo(()=>Object.fromEntries(skus.map(s=>[s.sku,s])),[skus]);
  const comboMap=useMemo(()=>Object.fromEntries(combos.map(c=>[c.sku,c])),[combos]);

  function processRows(rows){
    if(!rows?.length){showToast("error","File is empty.");return;}
    const hk=Object.keys(rows[0]);const sk=hk.find(k=>k.toLowerCase().includes("sku"));const qk=hk.find(k=>/quantity|qty|units/i.test(k));
    if(!sk||!qk){showToast("error",'Expected "sku" and "quantity" columns.');return;}
    const extra=detectExtraColumns(hk);
    setExtraCols(extra);
    setRawRows(rows.map(r=>({sku:String(r[sk]||"").trim(),qty:parseFloat(r[qk]),date:extra.dateKey?r[extra.dateKey]:null,buyer:extra.buyerKey?String(r[extra.buyerKey]||"").trim():"",city:extra.cityKey?String(r[extra.cityKey]||"").trim():"",state:extra.stateKey?String(r[extra.stateKey]||"").trim():"",price:extra.priceKey?parseFloat(r[extra.priceKey]):null})).filter(r=>r.sku&&!isNaN(r.qty)));
    const totals={};rows.forEach(r=>{const code=String(r[sk]||"").trim(),qty=parseFloat(r[qk]);if(!code||isNaN(qty))return;totals[code]=(totals[code]||0)+qty;});
    setAggregated(Object.entries(totals).map(([code,qty])=>{const combo=combos.find(c=>c.sku===code),sku=skuMap[code];const mt=combo?"combo":sku?"direct":"unknown";return{code,qty,matchType:mt,matchName:combo?.name||sku?.name||"Not in catalog"};}));setStage("parsed");
  }
  function handleFile(file){
    if(!file)return;setFileName(file.name);const ext=file.name.split(".").pop().toLowerCase();
    if(ext==="csv")Papa.parse(file,{header:true,skipEmptyLines:true,complete:res=>processRows(res.data),error:()=>showToast("error","Could not parse CSV.")});
    else if(ext==="xlsx"||ext==="xls"){const r=new FileReader();r.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:"array"});processRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""}));}catch{showToast("error","Could not parse spreadsheet.");}};r.readAsArrayBuffer(file);}
    else showToast("error","Upload a .csv or .xlsx file.");
  }
  function applyReport(){
    if(!weekLabel.trim()){showToast("error","Add a report label.");return;}
    const skuBefore={};skus.forEach(s=>skuBefore[s.sku]=s.stock);
    const updated=Object.fromEntries(skus.map(s=>[s.sku,{...s}]));
    aggregated.forEach(({code,qty,matchType})=>{if(matchType==="combo"){const c=combos.find(x=>x.sku===code);c?.components?.forEach(comp=>{if(updated[comp.sku])updated[comp.sku].stock=Math.max(0,updated[comp.sku].stock-comp.qty*qty);});}else if(matchType==="direct"&&updated[code])updated[code].stock=Math.max(0,updated[code].stock-qty);});
    const newSkus=Object.values(updated);const bMap=Object.fromEntries(skus.map(s=>[s.sku,s])),aMap=Object.fromEntries(newSkus.map(s=>[s.sku,s]));
    const reportId=Date.now().toString();
    const report={id:reportId,label:weekLabel.trim(),fileName,appliedAt:new Date().toISOString(),
      skuLines:newSkus.map(s=>({sku:s.sku,name:s.name,opening:skuBefore[s.sku]??s.stock,sold:(skuBefore[s.sku]??s.stock)-s.stock,closing:s.stock,reorderLevel:s.reorderLevel,status:stockStatus(s)})),
      comboLines:combos.map(c=>({sku:c.sku,name:c.name,readyBefore:comboReadiness(c,bMap).ready,readyAfter:comboReadiness(c,aMap).ready,bottleneck:comboReadiness(c,aMap).bottleneck})),
      unmatched:aggregated.filter(a=>a.matchType==="unknown")};

    // Build detailed sales lines (for ZenkyBox Sales Report analytics)
    const newLines=rawRows.map((row,i)=>{
      const combo=comboMap[row.sku],sku=skuMap[row.sku];
      const matchType=combo?"combo":sku?"direct":"unknown";
      const name=combo?.name||sku?.name||row.sku;
      const unitCost=unitCostOf(row.sku,matchType,skuMap,comboMap);
      const cost=unitCost*row.qty;
      const revenue=row.price!=null&&!isNaN(row.price)?row.price*row.qty:0;
      return{id:`${reportId}-${i}`,reportId,date:row.date||report.appliedAt,sku:row.sku,name,matchType,qty:row.qty,unitCost,cost,revenue,earning:revenue-cost,buyer:row.buyer||"",city:row.city||"",state:row.state||""};
    });

    setSkus(newSkus);setReports([report,...reports]);setSalesLines([...salesLines,...newLines]);
    logActivity?.("Sales report applied",`"${weekLabel.trim()}" — ${fileName} (${aggregated.length} codes, ${newLines.length} order lines)`);
    setStage("applied");showToast("success","Inventory updated. 📊");
  }
  function reset(){setStage("idle");setAggregated([]);setRawRows([]);setFileName("");setWeekLabel("");if(ref.current)ref.current.value="";}
  const hasExtras=extraCols.priceKey||extraCols.buyerKey||extraCols.cityKey;
  return(
    <div>
      <SectionHeader title="Upload Sales Report" subtitle="Amazon weekly sales CSV/XLSX — combos auto-deduct component SKUs."/>
      <Card>
        {stage==="idle"&&(<div className="rounded-2xl border-2 border-dashed p-8 text-center" style={{borderColor:C.zenkyPink,backgroundColor:"#FFF8FC"}}><Upload size={32} className="mx-auto mb-3" style={{color:C.zenkyPink}}/><p className="font-bold mb-4" style={{color:C.darkText,fontFamily:F.display}}>Choose your Amazon weekly sales report</p><input ref={ref} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={e=>handleFile(e.target.files[0])}/><PrimaryButton onClick={()=>ref.current?.click()}><Upload size={15}/>Select File</PrimaryButton></div>)}
        {stage==="parsed"&&(<div><div className="flex items-center justify-between mb-4 flex-wrap gap-2"><span className="text-sm" style={{color:C.darkText}}>Parsed <strong>{fileName}</strong> — {aggregated.length} codes</span><button onClick={reset} className="text-sm font-bold" style={{color:C.lightText}}>Change file</button></div>
          {hasExtras&&<div className="mb-3 p-2.5 rounded-xl text-xs flex items-center gap-2" style={{backgroundColor:"#F0FDE8",color:"#166534"}}><Check size={14}/>Detected {[extraCols.priceKey&&"price",extraCols.buyerKey&&"buyer",extraCols.cityKey&&"location",extraCols.dateKey&&"date"].filter(Boolean).join(", ")} — will populate ZenkyBox Sales Report.</div>}
          {!hasExtras&&<div className="mb-3 p-2.5 rounded-xl text-xs flex items-center gap-2" style={{backgroundColor:C.bgLight,color:C.lightText}}><AlertTriangle size={14}/>No price/buyer/location columns detected — stock will update, but Sales Report earning/buyer/location breakdowns won't have data for this upload.</div>}
          <div className="overflow-x-auto mb-4"><table className="w-full text-sm"><thead><tr style={{color:C.lightText}}><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Code</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Qty</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Type</th><th className="py-2 pr-3 text-left font-bold text-xs uppercase">Name</th></tr></thead><tbody>{aggregated.map(a=><tr key={a.code} className="border-t" style={{borderColor:C.border}}><td className="py-2 pr-3" style={{fontFamily:F.mono,fontWeight:600}}>{a.code}</td><td className="py-2 pr-3" style={{fontFamily:F.mono}}>{fmt(a.qty)}</td><td className="py-2 pr-3">{a.matchType==="combo"?<Stamp tone="mint">Combo</Stamp>:a.matchType==="direct"?<Stamp tone="purple">SKU</Stamp>:<Stamp tone="pink">Unknown</Stamp>}</td><td className="py-2 pr-3" style={{color:a.matchType==="unknown"?C.zenkyPink:C.darkText}}>{a.matchName}</td></tr>)}</tbody></table></div>
          <div className="flex items-end gap-3 flex-wrap"><div className="w-full sm:w-64"><label className="text-xs font-bold block mb-1" style={{color:C.lightText}}>Report label</label><Input placeholder="e.g. Week of Jun 16–22" value={weekLabel} onChange={e=>setWeekLabel(e.target.value)}/></div><PrimaryButton onClick={applyReport}><Check size={15}/>Apply to inventory</PrimaryButton></div></div>)}
        {stage==="applied"&&(<div className="text-center py-8"><Check size={32} className="mx-auto mb-3" style={{color:C.mintGreen}}/><p className="font-bold text-lg" style={{color:C.darkText,fontFamily:F.display}}>Report applied</p><p className="text-sm mt-1" style={{color:C.lightText}}>View breakdown in Reports tab, or full analytics in ZenkyBox Sales Report.</p><div className="mt-5"><PrimaryButton onClick={reset}><Upload size={15}/>Upload another</PrimaryButton></div></div>)}
      </Card>
    </div>
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
                  <div className="flex items-center gap-2">{isOpen?<ChevronDown size={16}/>:<ChevronRight size={16}/>}<div><div className="font-bold text-sm" style={{color:C.darkText,fontFamily:F.display}}>{r.label}</div><div className="text-xs" style={{color:C.lightText,fontFamily:F.mono}}>Applied {new Date(r.appliedAt).toLocaleDateString()} · {r.fileName}</div></div></div>
                  <div className="flex items-center gap-2 flex-wrap justify-end">{lowCount>0&&<Stamp tone="orange">{lowCount} low</Stamp>}{shortCombos>0&&<Stamp tone="pink">{shortCombos} short</Stamp>}{r.unmatched?.length>0&&<Stamp tone="purple">{r.unmatched.length} unmatched</Stamp>}</div>
                </button>
                {isOpen&&(
                  <div className="mt-4 pt-4 border-t" style={{borderColor:C.border}}>
                    <div className="flex items-center justify-between mb-3"><h4 className="font-bold text-sm" style={{color:C.darkText,fontFamily:F.display}}>SKU Breakdown</h4><button onClick={()=>exportReport(r)} className="inline-flex items-center gap-1 text-xs font-bold" style={{color:C.zenkyOrange}}><Download size={13}/>Export CSV</button></div>
                    <div className="overflow-x-auto mb-4"><table className="w-full text-sm"><thead><tr style={{color:C.lightText}}><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">SKU</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Name</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Opening</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Sold</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Closing</th><th className="py-1.5 pr-3 text-left font-bold text-xs uppercase">Status</th></tr></thead><tbody>{r.skuLines?.map(l=><tr key={l.sku} className="border-t" style={{borderColor:C.border}}><td className="py-1.5 pr-3" style={{fontFamily:F.mono,fontWeight:600}}>{l.sku}</td><td className="py-1.5 pr-3">{l.name}</td><td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmt(l.opening)}</td><td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmt(l.sold)}</td><td className="py-1.5 pr-3" style={{fontFamily:F.mono}}>{fmt(l.closing)}</td><td className="py-1.5 pr-3">{statusStamp(l.status)}</td></tr>)}</tbody></table></div>
                    {r.unmatched?.length>0&&<div className="mt-3 text-xs flex items-start gap-2" style={{color:C.zenkyPink}}><AlertTriangle size={14} className="mt-0.5"/><span>{r.unmatched.length} unmatched: {r.unmatched.map(u=>u.code).join(", ")}</span></div>}
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
  const skuMap=useMemo(()=>Object.fromEntries(skus.map(s=>[s.sku,s])),[skus]);
  const comboMap=useMemo(()=>Object.fromEntries(combos.map(c=>[c.sku,c])),[combos]);

  const TABS=[
    {id:"overall",label:"Overall"},
    {id:"mom",label:"Month-over-Month"},
    {id:"weekly",label:"Weekly"},
    {id:"buyer",label:"Buyer-wise"},
    {id:"location",label:"Location-wise"},
  ];

  const totals=useMemo(()=>salesLines.reduce((a,l)=>({qty:a.qty+l.qty,revenue:a.revenue+l.revenue,cost:a.cost+l.cost,earning:a.earning+l.earning}),{qty:0,revenue:0,cost:0,earning:0}),[salesLines]);

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

  function GroupedReport({groupFn,emptyMsg}){
    const groups=useMemo(()=>aggregateSalesLines(salesLines,groupFn,skuMap,comboMap),[salesLines,skuMap,comboMap]);
    if(!groups.length)return<Empty icon={BarChart3} title="No sales data yet" message={emptyMsg}/>;
    return(
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
    );
  }

  const overallBreakdown=useMemo(()=>skuComboBreakdown(salesLines),[salesLines]);

  return(
    <div>
      <SectionHeader title="ZenkyBox Sales Report" subtitle="SKU & combo performance across every applied sales upload — cost, revenue, and earning."/>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-6">
        {[{label:"Units Sold",value:fmt(totals.qty),color:C.zenkyPurple},
          {label:"Revenue",value:fmtINR(totals.revenue),color:C.zenkyPurple},
          {label:"Cost",value:fmtINR(totals.cost),color:C.zenkyOrange},
          {label:"Earning",value:fmtINR(totals.earning),color:totals.earning>=0?C.mintGreen:C.zenkyPink}].map(s=>(
          <Card key={s.label}><div className="text-xs font-bold uppercase" style={{color:C.lightText,fontFamily:F.body}}>{s.label}</div><div className="text-2xl font-black mt-1" style={{fontFamily:F.display,color:s.color}}>{s.value}</div></Card>
        ))}
      </div>

      {salesLines.length===0&&(
        <div className="mb-5 p-3 rounded-xl text-xs flex items-start gap-2" style={{backgroundColor:C.bgLight,color:C.lightText}}>
          <AlertTriangle size={14} className="flex-shrink-0 mt-0.5"/>
          <span>No sales data yet. Revenue/Earning/Buyer/Location figures populate once you upload sales files containing price, buyer, and location columns (see Upload Sales tab). Stock-only uploads still work but won't appear here.</span>
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
        salesLines.length===0?null:(
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
      {tab==="mom"&&<GroupedReport groupFn={l=>monthKey(l.date)} emptyMsg="Upload sales with a date column to see month-over-month trends."/>}
      {tab==="weekly"&&<GroupedReport groupFn={l=>weekKey(l.date)} emptyMsg="Upload sales with a date column to see weekly trends."/>}
      {tab==="buyer"&&<GroupedReport groupFn={l=>l.buyer||"Unknown buyer"} emptyMsg="Upload sales with a buyer name/email column to see buyer-wise data."/>}
      {tab==="location"&&<GroupedReport groupFn={l=>[l.city,l.state].filter(Boolean).join(", ")||"Unknown location"} emptyMsg="Upload sales with ship-city/ship-state columns to see location-wise data."/>}
    </div>
  );
}

/* ═══ SOURCE DATA (Admin only) ═══ */
function SourceDataView({activityLog,synced}){
  const dbUrl=process.env.NEXT_PUBLIC_SUPABASE_URL||"";
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
function AccessManagementView({role,adminPin,setAdminPin,showToast,logActivity}){
  const [newPin,setNewPin]=useState("");
  const [confirmPin,setConfirmPin]=useState("");
  function savePin(){
    if(newPin.length<4){showToast("error","PIN must be at least 4 characters.");return;}
    if(newPin!==confirmPin){showToast("error","PINs don't match.");return;}
    setAdminPin(newPin);setNewPin("");setConfirmPin("");
    logActivity?.("Admin PIN changed","Access PIN updated by admin");
    showToast("success","Admin PIN updated. ✨");
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
              <li>• Costing & Pricing calculator</li>
              <li>• Source Data & Access Management</li>
            </ul>
          </div>
          <div className="p-4 rounded-xl" style={{backgroundColor:C.bgLight}}>
            <div className="flex items-center gap-2 mb-2"><Users size={16} style={{color:C.zenkyOrange}}/><span className="font-bold text-sm" style={{fontFamily:F.display,color:C.zenkyOrange}}>Other Users</span></div>
            <ul className="text-xs space-y-1" style={{color:C.darkText}}>
              <li>• Add new SKUs & combos (no edit/delete)</li>
              <li>• Upload sales reports</li>
              <li>• View Dashboard, Combo Readiness, Reports, Sales Report</li>
              <li>• Cannot clear data, bulk-replace, or view Source Data</li>
            </ul>
          </div>
        </div>
        <p className="text-xs mt-4 p-2.5 rounded-lg" style={{backgroundColor:"#FFF3E6",color:"#9a5b0f"}}>
          ⚠️ This is a lightweight PIN gate stored with your workspace data — enough to prevent accidental changes by staff, but not bank-grade security (there's no per-person login). For real multi-user accounts, Supabase Auth would be the next upgrade.
        </p>
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

export default function App(){
  const [skus,setSkus]=useState([]);
  const [combos,setCombos]=useState([]);
  const [reports,setReports]=useState([]);
  const [salesLines,setSalesLines]=useState([]);
  const [activityLog,setActivityLog]=useState([]);
  const [adminPin,setAdminPinState]=useState("");
  const [loaded,setLoaded]=useState(false);
  const [view,setView]=useState("dashboard");
  const [sidebarOpen,setSidebarOpen]=useState(false);
  const [toast,setToast]=useState(null);
  const [synced,setSynced]=useState(false);
  const [role,setRole]=useState("staff"); // per-device only, not shared across users

  // Load initial data
  useEffect(()=>{
    (async()=>{
      try{
        const data=await loadData();
        if(data){
          setSkus(data.skus||[]);setCombos(data.combos||[]);setReports(data.reports||[]);
          setSalesLines(data.salesLines||[]);setActivityLog(data.activityLog||[]);
          setAdminPinState(data.adminPin||"");
        }
        setSynced(hasSync());
        // Restore this device's unlocked session (not shared — each device unlocks independently)
        if(typeof window!=="undefined"&&sessionStorage.getItem("zenkybox-role")==="admin")setRole("admin");
      }catch(e){console.error(e);}
      setLoaded(true);
    })();
  },[]);

  // Subscribe to real-time updates (Supabase)
  useEffect(()=>{
    let unsub=null;
    (async()=>{
      unsub=await subscribeToData(data=>{
        if(data){
          setSkus(data.skus||[]);setCombos(data.combos||[]);setReports(data.reports||[]);
          setSalesLines(data.salesLines||[]);setActivityLog(data.activityLog||[]);
          setAdminPinState(data.adminPin||"");
        }
      });
    })();
    return()=>{if(unsub)unsub();};
  },[]);

  // Save whenever data changes
  const saveTimeout=useRef(null);
  useEffect(()=>{
    if(!loaded)return;
    clearTimeout(saveTimeout.current);
    saveTimeout.current=setTimeout(()=>{saveData({skus,combos,reports,salesLines,activityLog,adminPin}).catch(()=>{});},500);
    return()=>clearTimeout(saveTimeout.current);
  },[skus,combos,reports,salesLines,activityLog,adminPin,loaded]);

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

  const currentNavItem=NAV.find(n=>n.id===view);
  const blocked=currentNavItem?.adminOnly&&role!=="admin";

  return(
    <div className="flex h-screen overflow-hidden" style={{backgroundColor:C.bgLight,fontFamily:F.body}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fredoka:wght@400;500;600;700&family=Baloo+2:wght@500;600;700&family=Nunito:wght@400;500;600;700&display=swap');`}</style>
      <Sidebar view={view} setView={setView} open={sidebarOpen} setOpen={setSidebarOpen} synced={synced} role={role} onUnlock={handleUnlock} onLock={handleLock}/>
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
              {view==="source-data"&&<SourceDataView activityLog={activityLog} synced={synced}/>}
              {view==="access"&&<AccessManagementView role={role} adminPin={adminPin} setAdminPin={setAdminPin} showToast={showToast} logActivity={logActivity}/>}
            </>
          )}
        </main>
      </div>
      {toast&&<div className="fixed bottom-4 right-4 px-5 py-3 rounded-full text-sm font-bold shadow-lg z-50" style={{backgroundColor:toast.type==="error"?C.zenkyPink:C.mintGreen,color:"#fff",fontFamily:F.display}}>{toast.msg}</div>}
    </div>
  );
}
