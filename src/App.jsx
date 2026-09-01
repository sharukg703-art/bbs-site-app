import React, { useState, useMemo, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { Download, Ruler, LogOut, Save, Plus, Trash2, Pencil, FolderOpen, ClipboardList, Calculator as CalcIcon } from "lucide-react";

/* ---------------------------------------------------------------------
   THEME — "blueprint / drawing sheet"
--------------------------------------------------------------------- */
const Theme = () => (
  <style>{`
    .bbs-root{
      --bp-bg:#0B2E4A; --bp-line:#BFE0F5; --bp-line-dim:#6E9BBE;
      --bp-card:#0F3457; --bp-white:#F3F7FB; --accent:#E2793D; --accent2:#F2C14E; --ok:#7FD6A6; --danger:#E2593D;
      font-family:'IBM Plex Sans',ui-sans-serif,system-ui,sans-serif;
      background:
        linear-gradient(var(--bp-bg),var(--bp-bg)),
        repeating-linear-gradient(0deg, rgba(191,224,245,0.06) 0px, rgba(191,224,245,0.06) 1px, transparent 1px, transparent 24px),
        repeating-linear-gradient(90deg, rgba(191,224,245,0.06) 0px, rgba(191,224,245,0.06) 1px, transparent 1px, transparent 24px);
      color:var(--bp-white);
      min-height:100%;
    }
    .bbs-mono{ font-family:'IBM Plex Mono',ui-monospace,monospace; }
    .bbs-display{ font-family:'Space Grotesk',ui-sans-serif,sans-serif; }
    .bbs-tab{ border:1px solid rgba(191,224,245,0.25); background:rgba(15,58,87,0.4); transition:all .15s ease; }
    .bbs-tab:hover{ background:rgba(191,224,245,0.08); }
    .bbs-tab.active{ background:var(--accent); border-color:var(--accent); color:#1a1200; }
    .bbs-card{ background:rgba(15,58,87,0.55); border:1px solid rgba(191,224,245,0.2); border-radius:4px; }
    .bbs-input{
      background:rgba(243,247,251,0.06); border:1px solid rgba(191,224,245,0.35);
      color:var(--bp-white); border-radius:3px; font-family:'IBM Plex Mono',monospace;
    }
    .bbs-input:focus{ outline:2px solid var(--accent); outline-offset:1px; border-color:var(--accent); }
    .bbs-th{
      background:#123A5C; color:var(--bp-white); font-family:'IBM Plex Mono',monospace;
      font-size:11px; text-transform:uppercase; letter-spacing:.04em;
    }
    .bbs-td{ border-bottom:1px solid rgba(191,224,245,0.15); }
    select.bbs-input{ appearance:none; }
    .scrollbar-thin::-webkit-scrollbar{height:6px; width:6px;}
    .scrollbar-thin::-webkit-scrollbar-thumb{background:rgba(191,224,245,0.3); border-radius:3px;}
    .bbs-btn{ display:inline-flex; align-items:center; gap:6px; padding:8px 14px; border-radius:4px; font-size:13px; font-weight:600; }
    .bbs-btn-primary{ background:var(--accent); color:#1a1200; }
    .bbs-btn-ghost{ background:transparent; border:1px solid rgba(191,224,245,0.35); color:var(--bp-white); }
    .bbs-btn-danger{ background:rgba(226,89,61,0.15); border:1px solid var(--danger); color:#FFD9CE; }
  `}</style>
);

/* ---------------------------------------------------------------------
   CONSTANTS + helpers
--------------------------------------------------------------------- */
const SHAPES = [
  { code: 1, name: "Straight" },
  { code: 2, name: "L-Shape (bent one end)" },
  { code: 3, name: "Bent-up both ends (U / anchored)" },
  { code: 4, name: "Rectangular Tie / Stirrup" },
  { code: 5, name: "Circular / Spiral Tie" },
  { code: 6, name: "Cranked Bar" },
  { code: 7, name: "Triangular Stirrup" },
  { code: 8, name: "Diamond / Cross Tie" },
];
const shapeName = (code) => (SHAPES.find((s) => s.code === Number(code)) || {}).name || "Unknown";
const unitWeight = (dia) => (dia * dia) / 162;
const uid = (prefix) => `${prefix}_${Date.now()}_${Math.floor(Math.random() * 1000)}`;

const MEMBER_COLORS = { Footing: "#E2793D", "Tie Beam": "#F2C14E", Column: "#7FD6A6", Beam: "#7FB0E2", Slab: "#C79BE0" };

/* ---------------------------------------------------------------------
   FORMULA-BASED EXCEL LAYOUT — mirrors the calculator's own math as
   real Excel formulas, so the exported file recalculates like the app.
--------------------------------------------------------------------- */
const INPUT_LABELS = {
  L: "Length L (mm)", W: "Width W (mm)", D: "Depth D (mm)", cover: "Clear Cover (mm)",
  diaMain: "Dia Main (mm)", diaCross: "Dia Cross (mm)", spMain: "Spacing Main (mm)", spCross: "Spacing Cross (mm)",
  Ln: "Clear Span Ln (mm)", b: "Width b (mm)", nBot: "No. Bottom Bars", diaBot: "Dia Bottom (mm)",
  nTop: "No. Top Bars", diaTop: "Dia Top (mm)", Ld: "Anchorage / Dev. Length Ld (mm)", diaTie: "Tie/Stirrup Dia (mm)",
  spTie: "Tie/Stirrup Spacing (mm)", hook: "Hook & Bend Allowance (mm)",
  H: "Clear Height H (mm)", nBars: "No. Main Bars",
  Lx: "Short Span Lx (mm)", Ly: "Long Span Ly (mm)", diaDist: "Distribution Dia (mm)", spDist: "Distribution Spacing (mm)",
};
const MEMBER_LABEL_BY_TAB = { footing: "Footing", tieBeam: "Tie Beam", column: "Column", beam: "Beam", slab: "Slab" };

function colLetter(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Builds the {refs, rowsDef} formula recipe for one member, given its live input object. */
function buildMemberFormulaRecipe(tabId, inputs, inputStartRow) {
  const inputEntries = Object.entries(inputs);
  const refs = {};
  inputEntries.forEach(([k], i) => { refs[k] = `B${inputStartRow + i}`; });

  let rowsDef = [];
  if (tabId === "footing") {
    const { L, W, D, cover, diaMain, diaCross, spMain, spCross } = refs;
    const leg = `(${D}-2*${cover})`, main1 = `(${L}-2*${cover})`, main2 = `(${W}-2*${cover})`;
    rowsDef = [
      { desc: "Footing Main Bars - Bottom (Long dir.)", code: 3, dia: `=${diaMain}`, spacing: `=${spMain}`,
        nBars: `=ROUNDUP((${W}-2*${cover})/${spMain},0)+1`, cutLen: `=${main1}+2*${leg}`,
        shapeDim: `="["&TEXT(${leg},"0")&"]"&CHAR(10)&TEXT(${main1},"0")&CHAR(10)&"["&TEXT(${leg},"0")&"]"` },
      { desc: "Footing Cross Bars - Bottom (Short dir.)", code: 3, dia: `=${diaCross}`, spacing: `=${spCross}`,
        nBars: `=ROUNDUP((${L}-2*${cover})/${spCross},0)+1`, cutLen: `=${main2}+2*${leg}`,
        shapeDim: `="["&TEXT(${leg},"0")&"]"&CHAR(10)&TEXT(${main2},"0")&CHAR(10)&"["&TEXT(${leg},"0")&"]"` },
    ];
  } else if (tabId === "tieBeam" || tabId === "beam") {
    const { Ln, b, D, cover, nBot, diaBot, nTop, diaTop, Ld, diaTie, spTie, hook } = refs;
    const bx = `(${b}-2*${cover})`, dx = `(${D}-2*${cover})`;
    const memberName = tabId === "tieBeam" ? "Tie Beam" : "Beam";
    rowsDef = [
      { desc: `${memberName} Bottom Main Bars`, code: 3, dia: `=${diaBot}`, spacing: "-", nBars: `=${nBot}`,
        cutLen: `=${Ln}+2*${Ld}`, shapeDim: `="["&TEXT(${Ld},"0")&"]"&CHAR(10)&TEXT(${Ln},"0")&CHAR(10)&"["&TEXT(${Ld},"0")&"]"` },
      { desc: `${memberName} Top Main Bars`, code: 3, dia: `=${diaTop}`, spacing: "-", nBars: `=${nTop}`,
        cutLen: `=${Ln}+2*${Ld}`, shapeDim: `="["&TEXT(${Ld},"0")&"]"&CHAR(10)&TEXT(${Ln},"0")&CHAR(10)&"["&TEXT(${Ld},"0")&"]"` },
      { desc: `${memberName} Stirrups (rect. hoop)`, code: 4, dia: `=${diaTie}`, spacing: `=${spTie}`,
        nBars: `=ROUNDUP(${Ln}/${spTie},0)+1`, cutLen: `=2*(${bx}+${dx})+${hook}`,
        shapeDim: `=TEXT(${bx},"0")&" x "&TEXT(${dx},"0")` },
    ];
  } else if (tabId === "column") {
    const { H, b, D, cover, nBars, diaMain, Ld, diaTie, spTie, hook } = refs;
    const bx = `(${b}-2*${cover})`, dx = `(${D}-2*${cover})`;
    rowsDef = [
      { desc: "Column Main Vertical Bars", code: 2, dia: `=${diaMain}`, spacing: "-", nBars: `=${nBars}`,
        cutLen: `=${H}+${Ld}`, shapeDim: `=TEXT(${H},"0")&CHAR(10)&"["&TEXT(${Ld},"0")&"]"` },
      { desc: "Column Ties / Stirrups (rect. hoop)", code: 4, dia: `=${diaTie}`, spacing: `=${spTie}`,
        nBars: `=ROUNDUP(${H}/${spTie},0)+1`, cutLen: `=2*(${bx}+${dx})+${hook}`,
        shapeDim: `=TEXT(${bx},"0")&" x "&TEXT(${dx},"0")` },
    ];
  } else if (tabId === "slab") {
    const { Lx, Ly, cover, diaMain, spMain, diaDist, spDist } = refs;
    const cut1 = `(${Lx}-2*${cover})`, cut2 = `(${Ly}-2*${cover})`;
    rowsDef = [
      { desc: "Slab Main Bars (Short span dir.)", code: 1, dia: `=${diaMain}`, spacing: `=${spMain}`,
        nBars: `=ROUNDUP((${Ly}-2*${cover})/${spMain},0)+1`, cutLen: `=${cut1}`, shapeDim: `=TEXT(${cut1},"0")` },
      { desc: "Slab Distribution Bars (Long dir.)", code: 1, dia: `=${diaDist}`, spacing: `=${spDist}`,
        nBars: `=ROUNDUP((${Lx}-2*${cover})/${spDist},0)+1`, cutLen: `=${cut2}`, shapeDim: `=TEXT(${cut2},"0")` },
    ];
  }
  return { inputEntries, refs, rowsDef };
}

/**
 * Writes one fully formula-driven BBS sheet for a single member into `wb`.
 * Returns the address of the sheet's TOTAL WEIGHT cell (e.g. "'Footing'!M18") for cross-sheet summary references.
 */
function writeMemberFormulaSheet(wb, sheetName, tabId, inputs, letterLines) {
  const inputStartRow = letterLines.length + 3; // leave room for title + letterhead + "INPUTS" header
  const { inputEntries, refs, rowsDef } = buildMemberFormulaRecipe(tabId, inputs, inputStartRow);
  const inputEndRow = inputStartRow + inputEntries.length - 1;
  const libStartRow = inputEndRow + 2;
  const tableTitleRow = libStartRow + SHAPES.length + 2;
  const headerRow = tableTitleRow + 1;
  const dataStartRow = headerRow + 1;
  const dataEndRow = dataStartRow + rowsDef.length - 1;
  const totalRow = dataEndRow + 2;

  const headers = ["S.No", "Description", "Shape Code", "Shape Type", "Shape Sketch", "Full Bar Description",
    "Dia (mm)", "Spacing c/c (mm)", "No. of Bars", "Cutting Length (mm)", "Total Length (m)", "Unit Wt (kg/m)", "Total Weight (kg)"];

  const aoa = [];
  aoa[0] = [`BAR BENDING SCHEDULE (BBS) - ${sheetName.toUpperCase()}`];
  letterLines.forEach((line, i) => { aoa[1 + i] = line; });
  aoa[letterLines.length + 1] = ["INPUTS (edit these — everything below recalculates automatically)"];
  inputEntries.forEach(([k, v], i) => { aoa[inputStartRow - 1 + i] = [INPUT_LABELS[k] || k, v]; });
  aoa[libStartRow - 1] = ["SHAPE CODE LIBRARY (reference — Shape Type looks these up automatically)"];
  SHAPES.forEach((s, i) => { aoa[libStartRow + i] = [s.code, s.name]; });
  aoa[tableTitleRow - 1] = [`${sheetName.toUpperCase()} BAR BENDING SCHEDULE`];
  aoa[headerRow - 1] = headers;
  rowsDef.forEach((r, i) => { aoa[dataStartRow - 1 + i] = [i + 1, r.desc]; });
  aoa[totalRow - 1] = ["TOTAL WEIGHT (kg)"];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 6 }, { wch: 30 }, { wch: 10 }, { wch: 24 }, { wch: 26 }, { wch: 55 },
    { wch: 9 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 12 }, { wch: 11 }, { wch: 13 }];
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } },
    { s: { r: tableTitleRow - 1, c: 0 }, e: { r: tableTitleRow - 1, c: 12 } },
  ];

  const setF = (addr, f) => { ws[addr] = { t: "n", f }; };
  const setN = (addr, v) => { ws[addr] = { t: "n", v }; };
  const setS = (addr, v) => { ws[addr] = { t: "s", v }; };
  const codeRange = `$A$${libStartRow}:$A$${libStartRow + SHAPES.length - 1}`;
  const nameRange = `$B$${libStartRow}:$B$${libStartRow + SHAPES.length - 1}`;

  rowsDef.forEach((r, i) => {
    const row = dataStartRow + i;
    setN(`C${row}`, r.code);
    ws[`D${row}`] = { t: "str", f: `IFERROR(INDEX(${nameRange},MATCH(C${row},${codeRange},0)),"Set code 1-8")` };
    ws[`E${row}`] = { t: "str", f: r.shapeDim.replace(/^=/, "") };
    setF(`G${row}`, r.dia.replace(/^=/, ""));
    if (r.spacing === "-") setS(`H${row}`, "-"); else setF(`H${row}`, r.spacing.replace(/^=/, ""));
    setF(`I${row}`, r.nBars.replace(/^=/, ""));
    setF(`J${row}`, r.cutLen.replace(/^=/, ""));
    setF(`K${row}`, `I${row}*J${row}/1000`);
    setF(`L${row}`, `G${row}^2/162`);
    setF(`M${row}`, `K${row}*L${row}`);
    ws[`F${row}`] = { t: "str", f: `G${row}&"mm dia HYSD bar, "&D${row}&", "&I${row}&" nos, cutting length "&J${row}&"mm, total length "&TEXT(K${row},"0.00")&" m, unit weight "&TEXT(L${row},"0.000")&" kg/m, total weight "&TEXT(M${row},"0.00")&" kg."` };
    ws[`K${row}`].z = "0.00"; ws[`L${row}`].z = "0.000"; ws[`M${row}`].z = "0.00"; ws[`J${row}`].z = "0";
  });
  setF(`M${totalRow}`, `SUM(M${dataStartRow}:M${dataEndRow})`);

  // mark yellow input cells + row heights so shape sketch text is readable
  for (let i = 0; i < inputEntries.length; i++) {
    const addr = `B${inputStartRow + i}`;
    if (!ws[addr]) ws[addr] = { t: "n", v: inputEntries[i][1] };
  }
  ws["!rows"] = ws["!rows"] || [];
  for (let i = 0; i < rowsDef.length; i++) ws["!rows"][dataStartRow - 1 + i] = { hpt: 60 };

  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  return `'${sheetName.slice(0, 31)}'!M${totalRow}`;
}

const DEFAULT_FOOTING = { L: 3400, W: 2100, D: 400, cover: 50, diaMain: 10, diaCross: 10, spMain: 125, spCross: 175 };
const DEFAULT_TIEBEAM = { Ln: 3000, b: 230, D: 300, cover: 25, nBot: 2, diaBot: 12, nTop: 2, diaTop: 12, Ld: 300, diaTie: 8, spTie: 150, hook: 150 };
const DEFAULT_COLUMN = { H: 3000, b: 300, D: 450, cover: 40, nBars: 8, diaMain: 16, Ld: 640, diaTie: 8, spTie: 150, hook: 150 };
const DEFAULT_BEAM = { Ln: 4000, b: 230, D: 450, cover: 25, nBot: 3, diaBot: 16, nTop: 2, diaTop: 12, Ld: 400, diaTie: 8, spTie: 150, hook: 150 };
const DEFAULT_SLAB = { Lx: 3000, Ly: 4000, cover: 20, diaMain: 10, spMain: 150, diaDist: 8, spDist: 200 };
const DEFAULT_CODES = {
  footingMain: 3, footingCross: 3, tieBeamBottom: 3, tieBeamTop: 3, tieBeamStirrup: 4,
  columnMain: 2, columnTie: 4, beamBottom: 3, beamTop: 3, beamStirrup: 4, slabMain: 1, slabDist: 1,
};
const DEFAULT_DPR = { date: "", work: "", manpower: "", material: "", weather: "Sunny", remarks: "" };

async function storageGet(key, shared = false) {
  try {
    const r = await window.storage.get(key, shared);
    return r ? JSON.parse(r.value) : null;
  } catch (e) {
    return null;
  }
}
async function storageSet(key, value, shared = false) {
  try {
    await window.storage.set(key, JSON.stringify(value), shared);
    return true;
  } catch (e) {
    console.error("Storage save failed:", e);
    return false;
  }
}

/* ---------------------------------------------------------------------
   BAR SHAPE SVG
--------------------------------------------------------------------- */
function BarShapeSVG({ code, a, b }) {
  const stroke = "#E2793D", textColor = "#BFE0F5";
  const label = (x, y, t, anchor = "middle") => (
    <text x={x} y={y} fontSize="9" fill={textColor} textAnchor={anchor} fontFamily="IBM Plex Mono, monospace">{t}</text>
  );
  let body = null;
  if (code === 1) {
    body = (<><line x1="15" y1="40" x2="135" y2="40" stroke={stroke} strokeWidth="3" strokeLinecap="round" />{label(75, 58, `${Math.round(a)} mm`)}</>);
  } else if (code === 2) {
    body = (<><line x1="15" y1="40" x2="110" y2="40" stroke={stroke} strokeWidth="3" strokeLinecap="round" /><line x1="110" y1="40" x2="110" y2="15" stroke={stroke} strokeWidth="3" strokeLinecap="round" />{label(60, 55, `${Math.round(a)}`)}{label(122, 25, `${Math.round(b)}`, "start")}</>);
  } else if (code === 3) {
    body = (<><line x1="30" y1="45" x2="30" y2="18" stroke={stroke} strokeWidth="3" strokeLinecap="round" /><line x1="30" y1="45" x2="120" y2="45" stroke={stroke} strokeWidth="3" strokeLinecap="round" /><line x1="120" y1="45" x2="120" y2="18" stroke={stroke} strokeWidth="3" strokeLinecap="round" />{label(20, 14, `${Math.round(b)}`)}{label(130, 14, `${Math.round(b)}`)}{label(75, 62, `${Math.round(a)}`)}</>);
  } else if (code === 4) {
    body = (<><rect x="35" y="16" width="80" height="48" fill="none" stroke={stroke} strokeWidth="3" rx="2" /><line x1="35" y1="16" x2="27" y2="10" stroke={stroke} strokeWidth="2" strokeLinecap="round" /><line x1="115" y1="64" x2="123" y2="70" stroke={stroke} strokeWidth="2" strokeLinecap="round" />{label(75, 78, `${Math.round(a)} x ${Math.round(b)}`)}</>);
  } else if (code === 5) {
    body = (<><circle cx="75" cy="40" r="26" fill="none" stroke={stroke} strokeWidth="3" />{label(75, 72, `dia ${Math.round(a || b || 0)}`)}</>);
  } else if (code === 6) {
    body = (<><polyline points="15,40 60,40 85,20 100,55 135,40" fill="none" stroke={stroke} strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />{label(75, 68, `${Math.round(a)} mm`)}</>);
  } else if (code === 7) {
    body = (<><polygon points="75,14 118,62 32,62" fill="none" stroke={stroke} strokeWidth="3" strokeLinejoin="round" />{label(75, 74, `${Math.round(a)} x ${Math.round(b)}`)}</>);
  } else if (code === 8) {
    body = (<><polygon points="75,12 118,40 75,68 32,40" fill="none" stroke={stroke} strokeWidth="3" strokeLinejoin="round" />{label(75, 78, `${Math.round(a)} x ${Math.round(b)}`)}</>);
  }
  return (<svg viewBox="0 0 150 80" width="130" height="70" style={{ display: "block" }}>{body}</svg>);
}

/* ---------------------------------------------------------------------
   SMALL UI HELPERS
--------------------------------------------------------------------- */
function NumField({ label, value, onChange }) {
  return (
    <label className="flex flex-col gap-1 text-xs">
      <span className="bbs-mono" style={{ color: "#BFE0F5", opacity: 0.85 }}>{label}</span>
      <input type="number" value={value} onChange={(e) => onChange(Number(e.target.value))} className="bbs-input px-2 py-1.5 text-sm" />
    </label>
  );
}
function ShapeCodeField({ value, onChange }) {
  return (
    <select value={value} onChange={(e) => onChange(Number(e.target.value))} className="bbs-input px-1 py-1 text-xs w-full">
      {SHAPES.map((s) => (<option key={s.code} value={s.code}>{s.code} — {s.name}</option>))}
    </select>
  );
}
function TextField({ label, value, onChange, type = "text", placeholder = "" }) {
  return (
    <label className="flex flex-col gap-1 text-xs">
      <span className="bbs-mono" style={{ color: "#BFE0F5", opacity: 0.85 }}>{label}</span>
      <input type={type} value={value} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} className="bbs-input px-2 py-1.5 text-sm" />
    </label>
  );
}

/* ---------------------------------------------------------------------
   LOGIN SCREEN
--------------------------------------------------------------------- */
function LoginScreen({ onLogin }) {
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [siteName, setSiteName] = useState("");
  const [err, setErr] = useState("");

  const submit = () => {
    if (!name.trim() || !email.trim() || !siteName.trim()) { setErr("Fill in all fields to continue."); return; }
    if (!/^\S+@\S+\.\S+$/.test(email)) { setErr("Enter a valid email address."); return; }
    setErr("");
    onLogin({ name: name.trim(), email: email.trim(), siteName: siteName.trim(), loggedInAt: new Date().toISOString() });
  };

  return (
    <div className="bbs-root min-h-screen flex items-center justify-center p-6">
      <Theme />
      <div className="bbs-card p-8 w-full max-w-sm">
        <div className="bbs-mono text-xs tracking-widest mb-1" style={{ color: "#E2793D" }}>STRUCTURAL DRAWING SET</div>
        <h1 className="bbs-display text-2xl font-bold mb-1 flex items-center gap-2"><Ruler size={22} style={{ color: "#F2C14E" }} /> BBS Site App</h1>
        <p className="text-xs mb-6" style={{ color: "#6E9BBE" }}>Sign in to open your site's BBS records and daily progress log.</p>
        <div className="space-y-3">
          <TextField label="Your Name" value={name} onChange={setName} placeholder="Site Engineer" />
          <TextField label="Email" value={email} onChange={setEmail} type="email" placeholder="you@site.com" />
          <TextField label="Site / Project Name" value={siteName} onChange={setSiteName} placeholder="e.g. Block A, Phase 1" />
        </div>
        {err && <div className="text-xs mt-3" style={{ color: "#FFB4A0" }}>{err}</div>}
        <button onClick={submit} className="bbs-btn bbs-btn-primary w-full justify-center mt-5">Enter Site</button>
        <p className="text-xs mt-4 leading-relaxed" style={{ color: "#6E9BBE" }}>
          This is a lightweight sign-in to label your data — not a secured account system. Your name here jus
