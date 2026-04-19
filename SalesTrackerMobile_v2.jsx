import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, LineChart, Line } from "recharts";
import * as XLSX from "xlsx";

// ── Tokens ────────────────────────────────────────────────────────────────────
const STATUS_CONFIG = {
  Win:          { color: "#1D9E75", bg: "#E1F5EE", text: "#085041", dark_bg: "#04342C", dark_text: "#9FE1CB" },
  Loss:         { color: "#E24B4A", bg: "#FCEBEB", text: "#501313", dark_bg: "#4A1B0C", dark_text: "#F0997B" },
  Negotiation:  { color: "#378ADD", bg: "#E6F1FB", text: "#042C53", dark_bg: "#042C53", dark_text: "#85B7EB" },
  "On-bidding": { color: "#EF9F27", bg: "#FAEEDA", text: "#412402", dark_bg: "#412402", dark_text: "#FAC775" },
  Revision:     { color: "#D4537E", bg: "#FBEAF0", text: "#4B1528", dark_bg: "#4B1528", dark_text: "#ED93B1" },
};
const STATUS_OPTIONS = ["Win", "Loss", "Negotiation", "On-bidding", "Revision"];
const SORT_OPTIONS = ["Proposal #", "Cost ↓", "Cost ↑", "Win Rate ↓", "Win Rate ↑", "Status"];

// ── Formatters ────────────────────────────────────────────────────────────────
const fmtM = n => { const v=Number(n||0); return v>=1e6?"₱"+(v/1e6).toFixed(2)+"M":v>=1e3?"₱"+(v/1e3).toFixed(0)+"K":"₱"+v.toFixed(0); };
const fmtFull = n => "₱"+Number(n||0).toLocaleString("en-PH",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmtPct = n => (Number(n||0)*100).toFixed(1)+"%";
const numVal = v => parseFloat(String(v).replace(/[^0-9.-]/g,""))||0;

// ── Business logic ─────────────────────────────────────────────────────────────
const SEED_DATA = [
  {id:1,  status:"Revision",    proposal:"PCORP-001-26", cost:34000000,    markup:0.20, revisions:32000000,   winRate:0,   comments:"Serendra 2 Alveo"},
  {id:2,  status:"Loss",        proposal:"PCORP-002-26", cost:11586094.43, markup:0.25, revisions:0,          winRate:0,   comments:"Alveo"},
  {id:3,  status:"Negotiation", proposal:"PCORP-003-26", cost:7827388.53,  markup:0.25, revisions:3629947.46, winRate:0.7, comments:"ACEN"},
  {id:4,  status:"On-bidding",  proposal:"PCORP-004-26", cost:578970.77,   markup:0.25, revisions:0,          winRate:0.3, comments:""},
  {id:5,  status:"On-bidding",  proposal:"PCORP-005-26", cost:2671382.72,  markup:0.25, revisions:0,          winRate:0.5, comments:"ACEN"},
  {id:6,  status:"Negotiation", proposal:"PCORP-006-26", cost:3685582.20,  markup:0.25, revisions:0,          winRate:0.9, comments:"Park Terraces"},
  {id:7,  status:"On-bidding",  proposal:"PCORP-007-26", cost:67924931.44, markup:0.25, revisions:0,          winRate:0.1, comments:""},
  {id:8,  status:"On-bidding",  proposal:"PCORP-008-26", cost:11067983.77, markup:0.25, revisions:0,          winRate:0.1, comments:"ACEN"},
  {id:9,  status:"Win",         proposal:"PCORP-009-26", cost:1223040,     markup:0.10, revisions:1223040,    winRate:1.0, comments:"Manpower 1Yr / Jangho"},
  {id:10, status:"On-bidding",  proposal:"PCORP-010-26", cost:11477183.54, markup:0.25, revisions:9557337.23, winRate:0.1, comments:""},
  {id:11, status:"On-bidding",  proposal:"PCORP-011-26", cost:1382754.45,  markup:0.25, revisions:2382754.45, winRate:0.1, comments:""},
  {id:12, status:"Loss",        proposal:"PCORP-012-26", cost:4615637.72,  markup:0.25, revisions:0,          winRate:0,   comments:""},
];

let _id = SEED_DATA.length + 1;

function computeRow(r) {
  const mv=(r.cost||0)*(r.markup||0), isWin=r.status==="Win", isLoss=r.status==="Loss";
  return {...r,markupValue:isWin?mv:0,totalRevenue:isWin?(r.cost||0)+mv:0,totalSold:isWin?1:0,
    totalSales:isWin?(r.cost||0)+mv:0,pipelineValue:(!isWin&&!isLoss)?(r.cost||0):0};
}
function autoStatus(wr,rev) {
  if(wr>=1) return "Win"; if(wr>=0.6) return "Negotiation"; if(wr>0) return "On-bidding";
  if(rev>0) return "Revision"; return "Loss";
}
const blankRow = () => ({id:_id++,status:"On-bidding",proposal:`PCORP-${String(_id).padStart(3,"0")}-26`,cost:0,markup:0.25,revisions:0,winRate:0.5,comments:""});

// ── SVG Icon set ──────────────────────────────────────────────────────────────
const Icon = {
  Dashboard:  <svg width="18" height="18" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg>,
  List:       <svg width="18" height="18" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>,
  Plus:       <svg width="22" height="22" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
  Sheets:     <svg width="18" height="18" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><path d="M8 3H5a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2h-3"/><rect x="8" y="1" width="8" height="4" rx="1"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/><line x1="10" y1="9" x2="10" y2="17"/></svg>,
  Export:     <svg width="18" height="18" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
  Sun:        <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" viewBox="0 0 24 24"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>,
  Moon:       <svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" viewBox="0 0 24 24"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>,
  Search:     <svg width="15" height="15" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>,
  Sort:       <svg width="15" height="15" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" viewBox="0 0 24 24"><line x1="3" y1="6" x2="21" y2="6"/><line x1="7" y1="12" x2="17" y2="12"/><line x1="11" y1="18" x2="13" y2="18"/></svg>,
  Sync:       <svg width="15" height="15" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><polyline points="1 4 1 10 7 10"/><polyline points="23 20 23 14 17 14"/><path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4-4.64 4.36A9 9 0 0 1 3.51 15"/></svg>,
  Check:      <svg width="15" height="15" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>,
  X:          <svg width="14" height="14" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  Trash:      <svg width="14" height="14" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>,
  Edit:       <svg width="14" height="14" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>,
  ChevronRight: <svg width="14" height="14" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" viewBox="0 0 24 24"><polyline points="9 18 15 12 9 6"/></svg>,
  Upload:     <svg width="18" height="18" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>,
};

// ── Theme context helper ───────────────────────────────────────────────────────
function useTheme() {
  const [dark, setDark] = useState(false);
  return { dark, toggle: () => setDark(d => !d) };
}

// ── Reusable components ───────────────────────────────────────────────────────
function Field({ label, children, hint }) {
  return (
    <div style={{ marginBottom:14 }}>
      <label style={{ fontSize:10, fontWeight:500, color:"var(--color-text-secondary)", textTransform:"uppercase", letterSpacing:"0.6px", display:"block", marginBottom:5 }}>{label}</label>
      {children}
      {hint && <div style={{ fontSize:10, color:"var(--color-text-tertiary)", marginTop:3 }}>{hint}</div>}
    </div>
  );
}

function Inp({ value, onChange, type="text", placeholder="", min, max, step }) {
  return <input type={type} value={value??""} placeholder={placeholder} min={min} max={max} step={step} onChange={onChange}
    style={{ width:"100%", fontSize:13, padding:"10px 12px", borderRadius:8, border:"0.5px solid var(--color-border-secondary)",
      background:"var(--color-background-secondary)", color:"var(--color-text-primary)", boxSizing:"border-box",
      fontFamily:"inherit", outline:"none" }} />;
}

function Sel({ value, onChange, options }) {
  return <select value={value} onChange={onChange} style={{ width:"100%", fontSize:13, padding:"10px 12px", borderRadius:8,
    border:"0.5px solid var(--color-border-secondary)", background:"var(--color-background-secondary)",
    color:"var(--color-text-primary)", fontFamily:"inherit" }}>
    {options.map(o => <option key={o}>{o}</option>)}
  </select>;
}

function Badge({ status, dark }) {
  const c = STATUS_CONFIG[status] || { bg:"#eee", text:"#333", dark_bg:"#333", dark_text:"#eee" };
  return <span style={{ fontSize:11, fontWeight:500, padding:"3px 10px", borderRadius:20,
    background: dark ? c.dark_bg : c.bg,
    color: dark ? c.dark_text : c.text,
    display:"inline-block", whiteSpace:"nowrap" }}>{status}</span>;
}

function Stat({ label, val, color }) {
  return <div>
    <div style={{ fontSize:10, color:"var(--color-text-tertiary)", textTransform:"uppercase", letterSpacing:"0.3px", marginBottom:2 }}>{label}</div>
    <div style={{ fontSize:12, fontWeight:500, color:color||"var(--color-text-primary)" }}>{val}</div>
  </div>;
}

function KpiCard({ label, val, color, sub, icon }) {
  return (
    <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)",
      borderRadius:12, padding:"13px 14px", display:"flex", flexDirection:"column", gap:3 }}>
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div style={{ fontSize:10, color:"var(--color-text-secondary)", textTransform:"uppercase", letterSpacing:"0.5px" }}>{label}</div>
        {icon && <span style={{ color: color||"var(--color-text-tertiary)", opacity:0.6 }}>{icon}</span>}
      </div>
      <div style={{ fontSize:22, fontWeight:500, color:color||"var(--color-text-primary)", lineHeight:1.1 }}>{val}</div>
      {sub && <div style={{ fontSize:10, color:"var(--color-text-tertiary)" }}>{sub}</div>}
    </div>
  );
}

function WinRateSlider({ value, onChange }) {
  const pct = Math.round(value * 100);
  const color = pct >= 60 ? "#1D9E75" : pct >= 30 ? "#EF9F27" : "#E24B4A";
  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}>
        <span style={{ fontSize:12, color:"var(--color-text-secondary)" }}>Win probability</span>
        <span style={{ fontSize:14, fontWeight:500, color }}>{pct}%</span>
      </div>
      <input type="range" min="0" max="1" step="0.05" value={value}
        onChange={e => onChange(parseFloat(e.target.value))}
        style={{ width:"100%", accentColor: color }} />
      <div style={{ display:"flex", justifyContent:"space-between", fontSize:9, color:"var(--color-text-tertiary)", marginTop:3 }}>
        <span>Loss (0)</span><span>Negotiation (0.6)</span><span>Win (1.0)</span>
      </div>
    </div>
  );
}

function Toast({ toast }) {
  if (!toast) return null;
  const ok = toast.type !== "danger";
  return (
    <div style={{ position:"absolute", top:16, left:"50%", transform:"translateX(-50%)",
      zIndex:300, padding:"9px 18px", borderRadius:20,
      background: ok ? "#E1F5EE" : "#FCEBEB",
      color: ok ? "#085041" : "#501313",
      fontSize:13, fontWeight:500,
      border:`0.5px solid ${ok ? "#5DCAA5" : "#F09595"}`,
      whiteSpace:"nowrap", display:"flex", alignItems:"center", gap:6,
      boxShadow:"0 4px 20px rgba(0,0,0,0.15)" }}>
      <span style={{ fontSize:14 }}>{ok ? <span style={{ color:"#1D9E75" }}>{Icon.Check}</span> : <span style={{ color:"#E24B4A" }}>{Icon.X}</span>}</span>
      {toast.msg}
    </div>
  );
}

function ConfirmDialog({ msg, onConfirm, onCancel }) {
  return (
    <div style={{ position:"absolute", inset:0, background:"rgba(0,0,0,0.5)", display:"flex",
      alignItems:"center", justifyContent:"center", zIndex:200 }}
      onClick={e => { if(e.target===e.currentTarget) onCancel(); }}>
      <div style={{ background:"var(--color-background-primary)", borderRadius:14, padding:"22px 20px",
        width:"min(320px,88%)", border:"0.5px solid var(--color-border-tertiary)" }}>
        <div style={{ fontSize:14, fontWeight:500, marginBottom:8 }}>Delete proposal?</div>
        <div style={{ fontSize:13, color:"var(--color-text-secondary)", marginBottom:20 }}>{msg}</div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
          <button onClick={onCancel} style={{ padding:"11px", borderRadius:9, fontSize:13, cursor:"pointer",
            background:"var(--color-background-secondary)", border:"0.5px solid var(--color-border-secondary)", color:"var(--color-text-secondary)" }}>
            Cancel
          </button>
          <button onClick={onConfirm} style={{ padding:"11px", borderRadius:9, fontSize:13, cursor:"pointer",
            background:"#E24B4A", color:"#fff", border:"none", fontWeight:500 }}>
            Delete
          </button>
        </div>
      </div>
    </div>
  );
}

function ProposalCard({ row, onEdit, onDelete, dark }) {
  const c = STATUS_CONFIG[row.status] || {};
  return (
    <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)",
      borderRadius:12, overflow:"hidden", marginBottom:8,
      borderLeft:`3px solid ${c.color||"#ccc"}` }}>
      <div style={{ padding:"13px 14px" }}>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:9 }}>
          <div style={{ display:"flex", alignItems:"center", gap:7, minWidth:0 }}>
            <Badge status={row.status} dark={dark} />
            <span style={{ fontSize:13, fontWeight:500, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{row.proposal}</span>
          </div>
          <div style={{ display:"flex", gap:5, flexShrink:0, marginLeft:8 }}>
            <button onClick={() => onEdit(row)} style={{ background:"var(--color-background-secondary)", border:"0.5px solid var(--color-border-secondary)",
              borderRadius:7, padding:"5px 10px", fontSize:12, cursor:"pointer", color:"var(--color-text-secondary)",
              display:"flex", alignItems:"center", gap:4 }}>
              {Icon.Edit} Edit
            </button>
            <button onClick={() => onDelete(row)} style={{ background:"none", border:"0.5px solid var(--color-border-secondary)",
              borderRadius:7, padding:"5px 8px", cursor:"pointer", color:"var(--color-text-tertiary)",
              display:"flex", alignItems:"center" }}>
              {Icon.Trash}
            </button>
          </div>
        </div>

        {/* Win rate progress bar */}
        <div style={{ marginBottom:9 }}>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
            <span style={{ fontSize:10, color:"var(--color-text-tertiary)", textTransform:"uppercase", letterSpacing:"0.3px" }}>Win probability</span>
            <span style={{ fontSize:10, fontWeight:500, color: c.color }}>{fmtPct(row.winRate)}</span>
          </div>
          <div style={{ height:4, background:"var(--color-background-secondary)", borderRadius:2, overflow:"hidden" }}>
            <div style={{ height:"100%", width:fmtPct(row.winRate), background:c.color, borderRadius:2, transition:"width 0.3s" }} />
          </div>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"7px 8px" }}>
          <Stat label="Cost"    val={fmtM(row.cost)} />
          <Stat label="Markup"  val={fmtPct(row.markup)} />
          <Stat label="Loss Rate" val={fmtPct(1-row.winRate)} />
          {row.pipelineValue > 0 && <Stat label="Pipeline" val={fmtM(row.pipelineValue)} color="#378ADD" />}
          {row.totalRevenue  > 0 && <Stat label="Revenue"  val={fmtM(row.totalRevenue)}  color="#1D9E75" />}
          {row.revisions     > 0 && <Stat label="Revisions" val={fmtM(row.revisions)}    color={STATUS_CONFIG.Revision.color} />}
        </div>
        {row.comments && (
          <div style={{ marginTop:9, fontSize:11, color:"var(--color-text-secondary)",
            borderTop:"0.5px solid var(--color-border-tertiary)", paddingTop:8,
            display:"flex", gap:5, alignItems:"flex-start" }}>
            <span style={{ flexShrink:0, opacity:0.5 }}>
              <svg width="11" height="11" fill="none" stroke="currentColor" strokeWidth="1.6" viewBox="0 0 24 24"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>
            </span>
            {row.comments}
          </div>
        )}
      </div>
    </div>
  );
}

function BottomSheet({ title, onClose, children }) {
  return (
    <div style={{ position:"absolute", inset:0, background:"rgba(0,0,0,0.42)", display:"flex",
      flexDirection:"column", justifyContent:"flex-end", zIndex:150 }}
      onClick={e => { if(e.target===e.currentTarget) onClose(); }}>
      <div style={{ background:"var(--color-background-primary)", borderRadius:"16px 16px 0 0",
        maxHeight:"90%", display:"flex", flexDirection:"column" }}
        onClick={e => e.stopPropagation()}>
        {/* Handle pill */}
        <div style={{ display:"flex", justifyContent:"center", paddingTop:10 }}>
          <div style={{ width:36, height:4, borderRadius:2, background:"var(--color-border-secondary)" }} />
        </div>
        <div style={{ padding:"10px 16px 12px", display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0, borderBottom:"0.5px solid var(--color-border-tertiary)" }}>
          <span style={{ fontSize:15, fontWeight:500 }}>{title}</span>
          <button onClick={onClose} style={{ background:"var(--color-background-secondary)", border:"none",
            borderRadius:20, width:28, height:28, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center",
            color:"var(--color-text-secondary)" }}>{Icon.X}</button>
        </div>
        <div style={{ overflowY:"auto", flex:1 }}>{children}</div>
      </div>
    </div>
  );
}

function RowForm({ row, onChange, onSave, onCancel, isNew }) {
  const c = computeRow(row);
  const upd = (k, v) => {
    const u = {...row, [k]:v};
    if (k==="winRate"||k==="revisions") u.status = autoStatus(k==="winRate"?v:row.winRate, k==="revisions"?v:row.revisions);
    onChange(u);
  };
  return (
    <div style={{ padding:"16px 16px 8px" }}>
      <Field label="Proposal #"><Inp value={row.proposal} onChange={e=>upd("proposal",e.target.value)} /></Field>
      <Field label="Status"><Sel value={row.status} onChange={e=>upd("status",e.target.value)} options={STATUS_OPTIONS} /></Field>
      <Field label="Cost Proposal (₱)"><Inp type="number" value={row.cost} onChange={e=>upd("cost",numVal(e.target.value))} /></Field>

      <Field label="Markup (e.g. 0.25 = 25%)">
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <Inp type="number" value={row.markup} step="0.01" min="0" max="1" onChange={e=>upd("markup",numVal(e.target.value))} />
          <span style={{ fontSize:13, fontWeight:500, color:"#378ADD", minWidth:42 }}>{fmtPct(row.markup)}</span>
        </div>
      </Field>

      <Field label="Win Rate">
        <WinRateSlider value={row.winRate} onChange={v=>upd("winRate",v)} />
      </Field>

      <Field label="Revisions (₱)"><Inp type="number" value={row.revisions} onChange={e=>upd("revisions",numVal(e.target.value))} /></Field>
      <Field label="Comments"><Inp value={row.comments} onChange={e=>upd("comments",e.target.value)} placeholder="Client, project notes…" /></Field>

      {/* Live computed preview */}
      <div style={{ background:"var(--color-background-secondary)", borderRadius:10, padding:"12px 14px", marginBottom:16 }}>
        <div style={{ fontSize:10, color:"var(--color-text-secondary)", textTransform:"uppercase", letterSpacing:"0.5px", marginBottom:9, fontWeight:500 }}>Live preview</div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"8px 12px" }}>
          {[
            ["Status",     row.status,          STATUS_CONFIG[row.status]?.color],
            ["Revenue",    fmtFull(c.totalRevenue),   "#1D9E75"],
            ["Pipeline",   fmtFull(c.pipelineValue),  "#378ADD"],
            ["Markup Val", fmtFull(c.markupValue),     undefined],
          ].map(([l,v,col])=>(
            <div key={l}>
              <div style={{ fontSize:10, color:"var(--color-text-tertiary)", textTransform:"uppercase", letterSpacing:"0.3px" }}>{l}</div>
              <div style={{ fontSize:12, fontWeight:500, color:col||"var(--color-text-primary)" }}>{v}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, paddingBottom:20 }}>
        <button onClick={onCancel} style={{ padding:"13px", borderRadius:10, fontSize:14, cursor:"pointer",
          background:"var(--color-background-secondary)", border:"0.5px solid var(--color-border-secondary)",
          color:"var(--color-text-secondary)", fontFamily:"inherit" }}>Cancel</button>
        <button onClick={()=>onSave(row)} style={{ padding:"13px", borderRadius:10, fontSize:14, cursor:"pointer",
          background:"#1D9E75", color:"#fff", border:"none", fontWeight:500, fontFamily:"inherit" }}>
          {isNew ? "Add Proposal" : "Save Changes"}
        </button>
      </div>
    </div>
  );
}

function NavBtn({ id, icon, label, active, onClick }) {
  return (
    <button onClick={onClick} style={{ flex:1, padding:"8px 4px 10px", border:"none", background:"none", cursor:"pointer",
      display:"flex", flexDirection:"column", alignItems:"center", gap:3,
      color: active ? "#1D9E75" : "var(--color-text-tertiary)",
      borderTop: active ? "2px solid #1D9E75" : "2px solid transparent",
      transition:"color 0.15s", fontFamily:"inherit" }}>
      <span style={{ lineHeight:1 }}>{icon}</span>
      <span style={{ fontSize:10, fontWeight: active?500:400 }}>{label}</span>
    </button>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const { dark, toggle: toggleDark } = useTheme();
  const [tab, setTab]             = useState("dashboard");
  const [rows, setRows]           = useState(SEED_DATA);
  const [editRow, setEditRow]     = useState(null);
  const [showForm, setShowForm]   = useState(false);
  const [isNew, setIsNew]         = useState(false);
  const [search, setSearch]       = useState("");
  const [filterStatus, setFilter] = useState("All");
  const [sortBy, setSortBy]       = useState("Proposal #");
  const [showSort, setShowSort]   = useState(false);
  const [toast, setToast]         = useState(null);
  const [confirmRow, setConfirmRow] = useState(null);
  const [gsConfig, setGsConfig]   = useState({ id:"", apiKey:"", csvUrl:"", sheetName:"SALES TRACKER" });
  const [syncStatus, setSyncStatus] = useState("idle");
  const [spinning, setSpinning]   = useState(false);
  const [autoSync, setAutoSync]   = useState(false);
  const fileRef = useRef();
  const syncTimer = useRef();

  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get("sales_rows_v3");
        if (r?.value) setRows(JSON.parse(r.value));
        const g = await window.storage.get("gs_config_v3");
        if (g?.value) setGsConfig(JSON.parse(g.value));
      } catch {}
    })();
  }, []);

  const saveRows = useCallback(async next => {
    setRows(next);
    try { await window.storage.set("sales_rows_v3", JSON.stringify(next)); } catch {}
  }, []);
  const saveGs = async cfg => {
    setGsConfig(cfg);
    try { await window.storage.set("gs_config_v3", JSON.stringify(cfg)); } catch {}
  };
  const showToast = (msg, type="success") => {
    setToast({msg,type});
    setTimeout(()=>setToast(null), 3500);
  };

  useEffect(() => {
    clearInterval(syncTimer.current);
    if (autoSync && (gsConfig.id || gsConfig.csvUrl))
      syncTimer.current = setInterval(pullFromSheets, 30000);
    return () => clearInterval(syncTimer.current);
  }, [autoSync, gsConfig]);

  const computed = useMemo(() => rows.map(computeRow), [rows]);

  const stats = useMemo(() => {
    const wins = computed.filter(r=>r.status==="Win").length;
    return {
      n:      rows.length,
      wins,
      wr:     rows.length ? wins/rows.length : 0,
      cost:   computed.reduce((s,r)=>s+(r.cost||0),0),
      rev:    computed.reduce((s,r)=>s+r.totalRevenue,0),
      pipe:   computed.reduce((s,r)=>s+r.pipelineValue,0),
      revise: computed.reduce((s,r)=>s+(r.revisions||0),0),
      markup: rows.length ? rows.reduce((s,r)=>s+(r.markup||0),0)/rows.length : 0,
    };
  }, [computed, rows]);

  const filteredSorted = useMemo(() => {
    let list = filterStatus==="All" ? computed : computed.filter(r=>r.status===filterStatus);
    if (search.trim()) {
      const q = search.toLowerCase();
      list = list.filter(r => r.proposal.toLowerCase().includes(q) || (r.comments||"").toLowerCase().includes(q));
    }
    const sorters = {
      "Proposal #": (a,b) => a.proposal.localeCompare(b.proposal),
      "Cost ↓":     (a,b) => b.cost - a.cost,
      "Cost ↑":     (a,b) => a.cost - b.cost,
      "Win Rate ↓": (a,b) => b.winRate - a.winRate,
      "Win Rate ↑": (a,b) => a.winRate - b.winRate,
      "Status":     (a,b) => STATUS_OPTIONS.indexOf(a.status)-STATUS_OPTIONS.indexOf(b.status),
    };
    return [...list].sort(sorters[sortBy]||sorters["Proposal #"]);
  }, [computed, filterStatus, search, sortBy]);

  const statusCounts = STATUS_OPTIONS.map(s => ({
    name:s, value:computed.filter(r=>r.status===s).length, color:STATUS_CONFIG[s].color,
  }));

  const pipeData = computed.filter(r=>r.pipelineValue>0)
    .sort((a,b)=>b.pipelineValue-a.pipelineValue).slice(0,6)
    .map(r=>({name:r.proposal.replace("PCORP-","#").replace("-26",""),
      val:+(r.pipelineValue/1e6).toFixed(2), color:STATUS_CONFIG[r.status]?.color}));

  const winRateData = [...computed]
    .sort((a,b)=>a.proposal.localeCompare(b.proposal)).slice(0,10)
    .map(r=>({name:r.proposal.replace("PCORP-","#").replace("-26",""), wr:Math.round(r.winRate*100)}));

  // Google Sheets
  async function pullFromSheets() {
    const { id, apiKey, csvUrl, sheetName } = gsConfig;
    if (!id && !csvUrl) { showToast("Configure Google Sheets first","danger"); return; }
    setSyncStatus("pulling"); setSpinning(true);
    try {
      let values;
      if (csvUrl) {
        const resp = await fetch(csvUrl);
        if (!resp.ok) throw new Error("Failed to fetch CSV");
        const text = await resp.text();
        const wb = XLSX.read(text,{type:"string"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        values = XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
      } else {
        const range = encodeURIComponent(`${sheetName||"SALES TRACKER"}!A:M`);
        const resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${range}?key=${apiKey}`);
        const data = await resp.json();
        if (data.error) throw new Error(data.error.message);
        values = data.values||[];
      }
      const hi = values.findIndex(r=>String(r[0]).toLowerCase().trim()==="status");
      if (hi<0) throw new Error("Status header not found");
      let idx=_id;
      const imported = values.slice(hi+1).filter(r=>STATUS_OPTIONS.includes(String(r[0]).trim()))
        .map(r=>({id:idx++,status:String(r[0]).trim(),proposal:String(r[1]||""),cost:numVal(r[2]),
          markup:numVal(r[3]),revisions:numVal(r[7]),winRate:numVal(r[8]),comments:String(r[12]||"")}));
      _id=idx; saveRows(imported);
      setSyncStatus("success"); setSpinning(false);
      showToast(`Pulled ${imported.length} rows`);
      setTimeout(()=>setSyncStatus("idle"),3000);
    } catch(err) {
      setSyncStatus("error"); setSpinning(false);
      showToast("Pull failed: "+err.message,"danger");
      setTimeout(()=>setSyncStatus("idle"),4000);
    }
  }

  async function pushToSheets() {
    const { id, apiKey, sheetName } = gsConfig;
    if (!id||!apiKey) { showToast("Enter Sheet ID and API Key","danger"); return; }
    setSyncStatus("pushing"); setSpinning(true);
    try {
      const header = ["Status","Proposal #","Cost Proposal (₱)","Markup %","Markup Value (₱)","Total Revenue (₱)","Total Sold","Revisions (₱)","Win Rate","Loss Rate","Total Sales (₱)","Pipeline Value (₱)","Comments"];
      const data = computed.map(r=>[r.status,r.proposal,r.cost,r.markup,r.markupValue,r.totalRevenue,r.totalSold,r.revisions,r.winRate,1-r.winRate,r.totalSales,r.pipelineValue,r.comments]);
      const range = `${sheetName||"SALES TRACKER"}!A5:M${5+data.length}`;
      const resp = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED&key=${apiKey}`,
        {method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify({range,majorDimension:"ROWS",values:[header,...data]})}
      );
      const result = await resp.json();
      if (result.error) throw new Error(result.error.code===403?"Permission denied — sheet must be publicly editable":result.error.message);
      setSyncStatus("success"); setSpinning(false);
      showToast("Pushed to Google Sheets");
      setTimeout(()=>setSyncStatus("idle"),3000);
    } catch(err) {
      setSyncStatus("error"); setSpinning(false);
      showToast(err.message,"danger");
      setTimeout(()=>setSyncStatus("idle"),4000);
    }
  }

  function parseAndImport(file) {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result,{type:"array"});
        const ws = wb.Sheets["SALES TRACKER"]||wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        const hi = raw.findIndex(r=>String(r[0]).toLowerCase().includes("status"));
        if (hi<0) { showToast("Header not found","danger"); return; }
        let idx=_id;
        const imported = raw.slice(hi+1).filter(r=>STATUS_OPTIONS.includes(String(r[0]).trim()))
          .map(r=>({id:idx++,status:String(r[0]).trim(),proposal:String(r[1]||""),cost:numVal(r[2]),
            markup:numVal(r[3]),revisions:numVal(r[7]),winRate:numVal(r[8]),comments:String(r[12]||"")}));
        _id=idx; saveRows(imported);
        showToast(`Imported ${imported.length} rows`);
        setTab("tracker");
      } catch(err) { showToast("Import failed: "+err.message,"danger"); }
    };
    reader.readAsArrayBuffer(file);
  }

  function exportXlsx() {
    const header = ["Status","Proposal #","Cost Proposal (₱)","Markup %","Markup Value (₱)","Total Revenue (₱)","Total Sold","Revisions (₱)","Win Rate","Loss Rate","Total Sales (₱)","Pipeline Value (₱)","Comments"];
    const data = computed.map(r=>[r.status,r.proposal,r.cost,r.markup,r.markupValue,r.totalRevenue,r.totalSold,r.revisions,r.winRate,1-r.winRate,r.totalSales,r.pipelineValue,r.comments]);
    const ws = XLSX.utils.aoa_to_sheet([header,...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"SALES TRACKER");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
      ["SALES DASHBOARD — FY 2026"],[],
      ["Total Proposals","Won","Win Rate","Pipeline","Revenue","Total Cost"],
      [stats.n,stats.wins,stats.wr,stats.pipe,stats.rev,stats.cost],
    ]),"DASHBOARD");
    XLSX.writeFile(wb,"SALES_TRACKER_EXPORT.xlsx");
    showToast("Exported successfully");
  }

  const openEdit = r => { setEditRow({...r}); setIsNew(false); setShowForm(true); };
  const openAdd  = () => { setEditRow(blankRow()); setIsNew(true); setShowForm(true); };
  const closeForm = () => { setShowForm(false); setEditRow(null); };
  const saveRow = r => {
    if (isNew) { saveRows([...rows,r]); showToast("Proposal added"); }
    else       { saveRows(rows.map(x=>x.id===r.id?r:x)); showToast("Saved"); }
    closeForm();
  };
  const confirmDelete = r => setConfirmRow(r);
  const doDelete = () => {
    saveRows(rows.filter(r=>r.id!==confirmRow.id));
    showToast("Deleted","danger"); setConfirmRow(null);
  };

  const syncColor = syncStatus==="success"?"#1D9E75":syncStatus==="error"?"#E24B4A":syncStatus==="idle"?"var(--color-text-tertiary)":"#EF9F27";

  const TABS = [
    {id:"dashboard", icon:Icon.Dashboard, label:"Dash"},
    {id:"tracker",   icon:Icon.List,      label:"Tracker"},
    {id:"add",       icon:Icon.Plus,      label:"Add"},
    {id:"sheets",    icon:Icon.Sheets,    label:"Sheets"},
    {id:"export",    icon:Icon.Export,    label:"Files"},
  ];

  return (
    <div style={{ fontFamily:"var(--font-sans),system-ui,sans-serif", display:"flex", flexDirection:"column",
      minHeight:750, background:"var(--color-background-tertiary)", position:"relative",
      color:"var(--color-text-primary)" }}>

      <style>{`@keyframes spin{to{transform:rotate(360deg)}}.spinning{animation:spin 0.8s linear infinite;display:inline-flex;align-items:center;}`}</style>

      <Toast toast={toast} />
      {confirmRow && <ConfirmDialog msg={`Delete "${confirmRow.proposal}"? This cannot be undone.`} onConfirm={doDelete} onCancel={()=>setConfirmRow(null)} />}

      {/* ── Header ── */}
      <div style={{ background:"var(--color-background-primary)", padding:"11px 16px",
        display:"flex", alignItems:"center", justifyContent:"space-between",
        borderBottom:"0.5px solid var(--color-border-tertiary)", flexShrink:0, gap:8 }}>
        <div style={{ minWidth:0 }}>
          <div style={{ fontWeight:500, fontSize:15, lineHeight:1.2 }}>
            Sales Tracker <span style={{ fontWeight:400, color:"var(--color-text-tertiary)", fontSize:13 }}>FY 2026</span>
          </div>
          <div style={{ fontSize:10, color:"var(--color-text-tertiary)", marginTop:2 }}>
            {rows.length} proposals · {fmtM(stats.pipe)} pipeline
          </div>
        </div>
        <div style={{ display:"flex", gap:6, alignItems:"center", flexShrink:0 }}>
          {(gsConfig.id||gsConfig.csvUrl) && (
            <button onClick={pullFromSheets} style={{ background:"none",
              border:`0.5px solid ${syncColor}`, borderRadius:8, padding:"5px 10px",
              fontSize:12, cursor:"pointer", color:syncColor, display:"flex", alignItems:"center", gap:5, fontWeight:500,
              fontFamily:"inherit" }}>
              <span className={spinning?"spinning":""}>{Icon.Sync}</span>
              {syncStatus==="pulling"?"Pulling…":syncStatus==="pushing"?"Pushing…":"Sync"}
            </button>
          )}
          {/* Dark mode toggle */}
          <button onClick={toggleDark} style={{ background:"var(--color-background-secondary)",
            border:"0.5px solid var(--color-border-secondary)", borderRadius:20, padding:"5px 11px",
            cursor:"pointer", display:"flex", alignItems:"center", gap:5, fontSize:12,
            color:"var(--color-text-secondary)", fontFamily:"inherit" }}>
            {dark ? Icon.Sun : Icon.Moon}
            <span>{dark?"Light":"Dark"}</span>
          </button>
        </div>
      </div>

      {/* ── Scrollable content ── */}
      <div style={{ flex:1, overflowY:"auto", overflowX:"hidden", paddingBottom:4 }}>

        {/* DASHBOARD */}
        {tab==="dashboard" && (
          <div style={{ padding:12 }}>
            {/* KPI grid */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:10, marginBottom:12 }}>
              <KpiCard label="Total proposals" val={stats.n} />
              <KpiCard label="Win rate" val={fmtPct(stats.wr)} color="#1D9E75" sub={`${stats.wins} won of ${stats.n}`}
                icon={<svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="1.6" viewBox="0 0 24 24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>} />
              <KpiCard label="Pipeline" val={fmtM(stats.pipe)} color="#378ADD"
                sub={`${computed.filter(r=>["On-bidding","Negotiation"].includes(r.status)).length} active`}
                icon={<svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="1.6" viewBox="0 0 24 24"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 1 0 0 7h5a3.5 3.5 0 1 1 0 7H6"/></svg>} />
              <KpiCard label="Revenue" val={fmtM(stats.rev)} color="#1D9E75"
                icon={<svg width="16" height="16" fill="none" stroke="currentColor" strokeWidth="1.6" viewBox="0 0 24 24"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>} />
              <KpiCard label="Total cost" val={fmtM(stats.cost)} />
              <KpiCard label="Avg markup" val={fmtPct(stats.markup)} color="var(--color-text-secondary)" />
            </div>

            {/* Status filter chips */}
            <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)",
              borderRadius:12, padding:"12px 14px", marginBottom:12 }}>
              <div style={{ fontSize:10, fontWeight:500, color:"var(--color-text-tertiary)", marginBottom:9,
                textTransform:"uppercase", letterSpacing:"0.5px" }}>Tap to filter tracker</div>
              <div style={{ display:"flex", flexWrap:"wrap", gap:7 }}>
                {statusCounts.map(s => (
                  <div key={s.name} onClick={()=>{ setFilter(s.name); setTab("tracker"); }}
                    style={{ display:"flex", alignItems:"center", gap:6, padding:"6px 11px", borderRadius:20,
                      background:STATUS_CONFIG[s.name].bg, cursor:"pointer", userSelect:"none",
                      border:`0.5px solid ${STATUS_CONFIG[s.name].color}22` }}>
                    <span style={{ width:7, height:7, borderRadius:"50%", background:STATUS_CONFIG[s.name].color, flexShrink:0 }} />
                    <span style={{ fontSize:12, fontWeight:500, color:STATUS_CONFIG[s.name].text }}>{s.name}</span>
                    <span style={{ fontSize:11, fontWeight:700, color:STATUS_CONFIG[s.name].color,
                      background:"rgba(0,0,0,0.07)", borderRadius:20, padding:"1px 7px" }}>{s.value}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Charts row */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:12 }}>
              {/* Donut */}
              <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)", borderRadius:12, padding:"12px" }}>
                <div style={{ fontSize:10, fontWeight:500, color:"var(--color-text-secondary)", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>By status</div>
                <ResponsiveContainer width="100%" height={130}>
                  <PieChart>
                    <Pie data={statusCounts.filter(s=>s.value>0)} dataKey="value" cx="50%" cy="50%"
                      outerRadius={55} innerRadius={28} paddingAngle={2}>
                      {statusCounts.filter(s=>s.value>0).map(e=><Cell key={e.name} fill={e.color} />)}
                    </Pie>
                    <Tooltip formatter={(v,n)=>[v+" proposals",n]} contentStyle={{ fontSize:11 }} />
                  </PieChart>
                </ResponsiveContainer>
                <div style={{ display:"flex", flexDirection:"column", gap:4, marginTop:4 }}>
                  {statusCounts.filter(s=>s.value>0).map(s=>(
                    <div key={s.name} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", fontSize:11 }}>
                      <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                        <span style={{ width:8, height:8, borderRadius:2, background:s.color, display:"inline-block" }} />
                        <span style={{ color:"var(--color-text-secondary)" }}>{s.name}</span>
                      </div>
                      <span style={{ fontWeight:500, color:s.color }}>{s.value}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Win rate scatter */}
              <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)", borderRadius:12, padding:"12px" }}>
                <div style={{ fontSize:10, fontWeight:500, color:"var(--color-text-secondary)", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.4px" }}>Win rate %</div>
                <ResponsiveContainer width="100%" height={180}>
                  <LineChart data={winRateData} margin={{top:4,right:4,left:-30,bottom:16}}>
                    <CartesianGrid strokeDasharray="2 2" stroke="var(--color-border-tertiary)" vertical={false} />
                    <XAxis dataKey="name" tick={{fontSize:9,fill:"var(--color-text-tertiary)"}} angle={-35} textAnchor="end" />
                    <YAxis tick={{fontSize:9,fill:"var(--color-text-tertiary)"}} domain={[0,100]} tickFormatter={v=>v+"%"} />
                    <Tooltip formatter={v=>[v+"%","Win Rate"]} contentStyle={{fontSize:11}} />
                    <Line type="monotone" dataKey="wr" stroke="#1D9E75" strokeWidth={2} dot={{r:3,fill:"#1D9E75"}} />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Pipeline bar */}
            {pipeData.length > 0 && (
              <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)", borderRadius:12, padding:"12px 14px" }}>
                <div style={{ fontSize:10, fontWeight:500, color:"var(--color-text-secondary)", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px" }}>Pipeline value (₱M)</div>
                <ResponsiveContainer width="100%" height={150}>
                  <BarChart data={pipeData} margin={{top:4,right:4,left:-24,bottom:16}}>
                    <CartesianGrid strokeDasharray="2 2" stroke="var(--color-border-tertiary)" vertical={false} />
                    <XAxis dataKey="name" tick={{fontSize:10,fill:"var(--color-text-secondary)"}} />
                    <YAxis tick={{fontSize:10,fill:"var(--color-text-secondary)"}} tickFormatter={v=>`${v}M`} />
                    <Tooltip formatter={v=>[`₱${v}M`,"Pipeline"]} contentStyle={{fontSize:11}} />
                    <Bar dataKey="val" radius={[4,4,0,0]}>
                      {pipeData.map((e,i)=><Cell key={i} fill={e.color} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        )}

        {/* TRACKER */}
        {tab==="tracker" && (
          <div style={{ padding:12 }}>
            {/* Search bar */}
            <div style={{ position:"relative", marginBottom:10 }}>
              <div style={{ position:"absolute", left:10, top:"50%", transform:"translateY(-50%)",
                color:"var(--color-text-tertiary)", pointerEvents:"none" }}>{Icon.Search}</div>
              <input value={search} onChange={e=>setSearch(e.target.value)}
                placeholder="Search proposals or comments…"
                style={{ width:"100%", padding:"10px 36px", borderRadius:10, fontSize:13,
                  border:"0.5px solid var(--color-border-secondary)", background:"var(--color-background-primary)",
                  color:"var(--color-text-primary)", boxSizing:"border-box", fontFamily:"inherit", outline:"none" }} />
              {search && (
                <button onClick={()=>setSearch("")} style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)",
                  background:"none", border:"none", cursor:"pointer", color:"var(--color-text-tertiary)", display:"flex", alignItems:"center" }}>
                  {Icon.X}
                </button>
              )}
            </div>

            {/* Filter + Sort row */}
            <div style={{ display:"flex", gap:7, marginBottom:10, alignItems:"center" }}>
              <div style={{ display:"flex", gap:5, overflowX:"auto", flex:1 }}>
                {["All",...STATUS_OPTIONS].map(s => (
                  <button key={s} onClick={()=>setFilter(s)} style={{ flexShrink:0, padding:"6px 13px", borderRadius:20, fontSize:11,
                    cursor:"pointer", fontWeight:filterStatus===s?500:400, whiteSpace:"nowrap", fontFamily:"inherit",
                    background: filterStatus===s ? (s==="All"?"var(--color-text-primary)":STATUS_CONFIG[s]?.bg||"#eee") : "var(--color-background-primary)",
                    color: filterStatus===s ? (s==="All"?"var(--color-background-primary)":STATUS_CONFIG[s]?.text||"#333") : "var(--color-text-secondary)",
                    border: `0.5px solid ${filterStatus===s ? (s==="All"?"var(--color-text-primary)":STATUS_CONFIG[s]?.color||"#ccc") : "var(--color-border-secondary)"}` }}>
                    {s}
                  </button>
                ))}
              </div>
              <button onClick={()=>setShowSort(v=>!v)} style={{ flexShrink:0, background:"var(--color-background-primary)",
                border:"0.5px solid var(--color-border-secondary)", borderRadius:8, padding:"6px 10px",
                cursor:"pointer", color:"var(--color-text-secondary)", display:"flex", alignItems:"center", gap:5, fontSize:12, fontFamily:"inherit" }}>
                {Icon.Sort}
              </button>
            </div>

            {/* Sort dropdown */}
            {showSort && (
              <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-secondary)",
                borderRadius:10, marginBottom:10, overflow:"hidden" }}>
                {SORT_OPTIONS.map(s => (
                  <div key={s} onClick={()=>{ setSortBy(s); setShowSort(false); }}
                    style={{ padding:"10px 14px", fontSize:13, cursor:"pointer", display:"flex", justifyContent:"space-between",
                      background: sortBy===s ? "var(--color-background-secondary)" : "transparent",
                      color: sortBy===s ? "var(--color-text-primary)" : "var(--color-text-secondary)",
                      borderBottom:"0.5px solid var(--color-border-tertiary)" }}>
                    {s} {sortBy===s && <span style={{ color:"#1D9E75" }}>{Icon.Check}</span>}
                  </div>
                ))}
              </div>
            )}

            {/* Results summary */}
            <div style={{ fontSize:11, color:"var(--color-text-tertiary)", marginBottom:8 }}>
              {filteredSorted.length} of {rows.length} proposals{sortBy!=="Proposal #"?` · sorted by ${sortBy}`:""}
            </div>

            {filteredSorted.map(r => (
              <ProposalCard key={r.id} row={r} onEdit={openEdit} onDelete={confirmDelete} dark={dark} />
            ))}

            {filteredSorted.length === 0 && (
              <div style={{ textAlign:"center", padding:"40px 20px", color:"var(--color-text-tertiary)" }}>
                <svg width="36" height="36" fill="none" stroke="currentColor" strokeWidth="1.2" viewBox="0 0 24 24" style={{ marginBottom:10, opacity:0.4 }}><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                <div style={{ fontSize:13 }}>No proposals found{filterStatus!=="All"?` with status "${filterStatus}"`:""}.</div>
                {search && <div style={{ fontSize:12, marginTop:4 }}>Try clearing the search.</div>}
              </div>
            )}

            {/* Totals footer */}
            {filteredSorted.length > 0 && (
              <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)",
                borderRadius:12, padding:"12px 14px", marginTop:4 }}>
                <div style={{ fontSize:10, fontWeight:500, color:"var(--color-text-tertiary)", marginBottom:8,
                  textTransform:"uppercase", letterSpacing:"0.4px" }}>Filtered totals</div>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"6px 12px", fontSize:12 }}>
                  <Stat label="Cost"     val={fmtM(filteredSorted.reduce((s,r)=>s+(r.cost||0),0))} />
                  <Stat label="Revenue"  val={fmtM(filteredSorted.reduce((s,r)=>s+r.totalRevenue,0))} color="#1D9E75" />
                  <Stat label="Pipeline" val={fmtM(filteredSorted.reduce((s,r)=>s+r.pipelineValue,0))} color="#378ADD" />
                  <Stat label="Revisions" val={fmtM(filteredSorted.reduce((s,r)=>s+(r.revisions||0),0))} color={STATUS_CONFIG.Revision.color} />
                </div>
              </div>
            )}
          </div>
        )}

        {/* GOOGLE SHEETS */}
        {tab==="sheets" && (
          <div style={{ padding:12 }}>
            {/* Connection banner */}
            <div style={{ borderRadius:10, padding:"11px 14px", marginBottom:12, display:"flex", alignItems:"center", gap:9,
              background: syncStatus==="success"?"#E1F5EE":syncStatus==="error"?"#FCEBEB":"var(--color-background-secondary)",
              border:`0.5px solid ${syncStatus==="success"?"#5DCAA5":syncStatus==="error"?"#F09595":"var(--color-border-tertiary)"}` }}>
              <span className={spinning?"spinning":""} style={{ color:syncColor, display:"flex" }}>{Icon.Sync}</span>
              <span style={{ fontSize:12, color:syncStatus==="success"?"#085041":syncStatus==="error"?"#501313":"var(--color-text-secondary)" }}>
                {syncStatus==="pulling"?"Pulling from Google Sheets…":syncStatus==="pushing"?"Pushing to Google Sheets…":syncStatus==="success"?"Synced successfully":syncStatus==="error"?"Sync failed — check config below":"Not connected"}
              </span>
            </div>

            <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)", borderRadius:12, padding:16, marginBottom:12 }}>
              <div style={{ fontSize:14, fontWeight:500, marginBottom:14 }}>Google Sheets connection</div>
              <Field label="Spreadsheet ID" hint="From the URL: spreadsheets/d/[THIS PART]/edit">
                <Inp value={gsConfig.id} placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms" onChange={e=>saveGs({...gsConfig,id:e.target.value})} />
              </Field>
              <Field label="API Key" hint="console.cloud.google.com → Sheets API → Credentials → API Key">
                <input type="password" value={gsConfig.apiKey||""} placeholder="AIzaSy…" onChange={e=>saveGs({...gsConfig,apiKey:e.target.value})}
                  style={{ width:"100%",fontSize:13,padding:"10px 12px",borderRadius:8,border:"0.5px solid var(--color-border-secondary)",
                    background:"var(--color-background-secondary)",color:"var(--color-text-primary)",boxSizing:"border-box",fontFamily:"inherit" }} />
              </Field>
              <Field label="Sheet tab name">
                <Inp value={gsConfig.sheetName} placeholder="SALES TRACKER" onChange={e=>saveGs({...gsConfig,sheetName:e.target.value})} />
              </Field>
              <Field label="Published CSV URL (read-only, no API key needed)" hint="File → Share → Publish to web → CSV format">
                <Inp value={gsConfig.csvUrl} placeholder="https://docs.google.com/spreadsheets/d/…/export?format=csv&gid=0"
                  onChange={e=>saveGs({...gsConfig,csvUrl:e.target.value})} />
              </Field>

              {/* Auto-sync toggle */}
              <div style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 12px",
                background:"var(--color-background-secondary)", borderRadius:9, marginBottom:14, cursor:"pointer" }}
                onClick={()=>setAutoSync(v=>!v)}>
                <div style={{ width:38, height:22, borderRadius:11, background:autoSync?"#1D9E75":"var(--color-border-secondary)",
                  position:"relative", transition:"background 0.2s", flexShrink:0 }}>
                  <div style={{ width:18, height:18, borderRadius:9, background:"#fff", position:"absolute",
                    top:2, left:autoSync?18:2, transition:"left 0.2s" }} />
                </div>
                <div>
                  <div style={{ fontSize:12, fontWeight:500 }}>Auto-sync every 30 seconds</div>
                  <div style={{ fontSize:10, color:"var(--color-text-tertiary)" }}>Pulls latest data automatically</div>
                </div>
              </div>

              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
                <button onClick={pullFromSheets} disabled={!gsConfig.id&&!gsConfig.csvUrl}
                  style={{ padding:"12px", borderRadius:9, fontSize:13, cursor:"pointer", fontFamily:"inherit",
                    background:"var(--color-background-secondary)", border:"0.5px solid var(--color-border-secondary)",
                    color:"var(--color-text-primary)", opacity:(!gsConfig.id&&!gsConfig.csvUrl)?0.4:1,
                    display:"flex", alignItems:"center", justifyContent:"center", gap:6 }}>
                  {Icon.Export} Pull data
                </button>
                <button onClick={pushToSheets} disabled={!gsConfig.id||!gsConfig.apiKey}
                  style={{ padding:"12px", borderRadius:9, fontSize:13, cursor:"pointer", fontFamily:"inherit",
                    background:"#1D9E75", color:"#fff", border:"none", fontWeight:500,
                    opacity:(!gsConfig.id||!gsConfig.apiKey)?0.4:1,
                    display:"flex", alignItems:"center", justifyContent:"center", gap:6 }}>
                  {Icon.Upload} Push data
                </button>
              </div>
            </div>

            <div style={{ background:"var(--color-background-secondary)", borderRadius:12, padding:"12px 14px" }}>
              <div style={{ fontSize:11, fontWeight:500, marginBottom:8, textTransform:"uppercase", letterSpacing:"0.4px" }}>Setup guide</div>
              {["Open your Google Sheet","Share → Anyone with link can view","Go to console.cloud.google.com → Enable Sheets API","Create an API Key under Credentials","Copy the Sheet ID from the URL","Paste both above and tap Pull","For push/write: sheet must be publicly editable"].map((s,i)=>(
                <div key={i} style={{ display:"flex", gap:8, marginBottom:6, fontSize:11, color:"var(--color-text-secondary)" }}>
                  <span style={{ color:"#1D9E75", fontWeight:700, flexShrink:0, minWidth:14 }}>{i+1}.</span>
                  <span>{s}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* FILES / IMPORT EXPORT */}
        {tab==="export" && (
          <div style={{ padding:12 }}>
            <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)", borderRadius:12, padding:16, marginBottom:12 }}>
              <div style={{ fontSize:14, fontWeight:500, marginBottom:4 }}>Import Excel</div>
              <div style={{ fontSize:12, color:"var(--color-text-secondary)", marginBottom:12, lineHeight:1.5 }}>
                Upload your SALES_TRACKER .xlsx file — it auto-parses and populates the dashboard.
              </div>
              <div onClick={()=>fileRef.current?.click()}
                style={{ border:"2px dashed var(--color-border-secondary)", borderRadius:10,
                  padding:"28px 16px", textAlign:"center", cursor:"pointer", background:"var(--color-background-secondary)",
                  transition:"border-color 0.15s" }}>
                <div style={{ display:"flex", justifyContent:"center", marginBottom:10, color:"var(--color-text-tertiary)" }}>{Icon.Upload}</div>
                <div style={{ fontSize:13, fontWeight:500, marginBottom:3 }}>Tap to upload Excel file</div>
                <div style={{ fontSize:11, color:"var(--color-text-tertiary)" }}>.xlsx or .xls · Needs "Status" header row</div>
              </div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display:"none" }}
                onChange={e=>{ if(e.target.files[0]) parseAndImport(e.target.files[0]); }} />
            </div>

            <div style={{ background:"var(--color-background-primary)", border:"0.5px solid var(--color-border-tertiary)", borderRadius:12, padding:16 }}>
              <div style={{ fontSize:14, fontWeight:500, marginBottom:12 }}>Export to Excel</div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"8px 10px", marginBottom:14 }}>
                <Stat label="Proposals" val={stats.n} />
                <Stat label="Pipeline" val={fmtM(stats.pipe)} color="#378ADD" />
                <Stat label="Win Rate" val={fmtPct(stats.wr)} color="#1D9E75" />
                <Stat label="Revenue" val={fmtM(stats.rev)} color="#1D9E75" />
                <Stat label="Total Cost" val={fmtM(stats.cost)} />
                <Stat label="Revisions" val={fmtM(stats.revise)} color={STATUS_CONFIG.Revision.color} />
              </div>
              <button onClick={exportXlsx} style={{ width:"100%", padding:"14px", borderRadius:10, fontSize:14,
                cursor:"pointer", background:"#1D9E75", color:"#fff", border:"none", fontWeight:500, fontFamily:"inherit",
                display:"flex", alignItems:"center", justifyContent:"center", gap:8 }}>
                {Icon.Export} Download SALES_TRACKER_EXPORT.xlsx
              </button>
              <div style={{ fontSize:10, color:"var(--color-text-tertiary)", marginTop:7, textAlign:"center" }}>
                Includes SALES TRACKER + DASHBOARD sheets with all computed fields
              </div>
            </div>
          </div>
        )}
      </div>

      {/* ── Bottom Nav ── */}
      <div style={{ background:"var(--color-background-primary)", borderTop:"0.5px solid var(--color-border-tertiary)",
        display:"flex", flexShrink:0 }}>
        {TABS.map(t => (
          <NavBtn key={t.id} id={t.id} icon={t.icon} label={t.label}
            active={t.id!=="add" && tab===t.id}
            onClick={()=>{ if(t.id==="add") openAdd(); else setTab(t.id); }} />
        ))}
      </div>

      {/* ── Add / Edit sheet ── */}
      {showForm && editRow && (
        <BottomSheet title={isNew?"New proposal":"Edit proposal"} onClose={closeForm}>
          <RowForm row={editRow} onChange={setEditRow} onSave={saveRow} onCancel={closeForm} isNew={isNew} />
        </BottomSheet>
      )}
    </div>
  );
}
