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
          This is a lightweight sign-in to label your data — not a secured account system. Your name here just
          tags what you save; it isn't a password.
        </p>
        <p className="text-xs mt-2 leading-relaxed" style={{ color: "#F2C14E" }}>
          Site Records and Daily Progress Reports are SHARED — anyone who opens this app (e.g. your team, or
          anyone with the link if published) sees the same site-wide log. Don't put confidential data in here.
        </p>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   MAIN APP
--------------------------------------------------------------------- */
export default function BBSSiteApp() {
  const [booting, setBooting] = useState(true);
  const [profile, setProfile] = useState(null);
  const [section, setSection] = useState("calculator");
  const [calcTab, setCalcTab] = useState("footing");

  const [footing, setFooting] = useState(DEFAULT_FOOTING);
  const [tieBeam, setTieBeam] = useState(DEFAULT_TIEBEAM);
  const [column, setColumn] = useState(DEFAULT_COLUMN);
  const [beam, setBeam] = useState(DEFAULT_BEAM);
  const [slab, setSlab] = useState(DEFAULT_SLAB);
  const [shapeCodes, setShapeCodes] = useState(DEFAULT_CODES);
  const setCode = (key, val) => setShapeCodes((s) => ({ ...s, [key]: val }));

  const [records, setRecords] = useState([]);
  const [loadedRecordId, setLoadedRecordId] = useState({ footing: null, tieBeam: null, column: null, beam: null, slab: null });
  const [recordLabel, setRecordLabel] = useState({ footing: "", tieBeam: "", column: "", beam: "", slab: "" });

  const [dprEntries, setDprEntries] = useState([]);
  const [dprForm, setDprForm] = useState(DEFAULT_DPR);
  const [editingDprId, setEditingDprId] = useState(null);

  const [saveMsg, setSaveMsg] = useState("");
  const [downloadReady, setDownloadReady] = useState(null); // { url, filename }

  /* ---------- boot: load persisted data ---------- */
  useEffect(() => {
    (async () => {
      const p = await storageGet("bbs-app-profile", false);
      const r = await storageGet("bbs-app-records", true);
      const d = await storageGet("bbs-app-dpr", true);
      if (p) setProfile(p);
      if (r) setRecords(r);
      if (d) setDprEntries(d);
      setBooting(false);
    })();
  }, []);

  const handleLogin = async (p) => {
    setProfile(p);
    await storageSet("bbs-app-profile", p);
  };
  const handleLogout = async () => {
    setProfile(null);
    await storageSet("bbs-app-profile", null);
  };

  /* ---------- computed rows (same formulas as before) ---------- */
  const footingRows = useMemo(() => {
    const { L, W, D, cover, diaMain, diaCross, spMain, spCross } = footing;
    const leg = D - 2 * cover;
    const main1 = L - 2 * cover, n1 = Math.ceil((W - 2 * cover) / spMain) + 1, cut1 = main1 + 2 * leg;
    const main2 = W - 2 * cover, n2 = Math.ceil((L - 2 * cover) / spCross) + 1, cut2 = main2 + 2 * leg;
    const mk = (key, desc, dia, sp, n, cut, a, b) => {
      const totLen = (n * cut) / 1000, uw = unitWeight(dia), totW = totLen * uw;
      return { key, member: "Footing", desc, code: shapeCodes[key], dia, spacing: sp, nBars: n, cutLen: cut, totLen, unitW: uw, totW, a, b };
    };
    return [
      mk("footingMain", "Footing Main Bars – Bottom (Long dir.)", diaMain, spMain, n1, cut1, main1, leg),
      mk("footingCross", "Footing Cross Bars – Bottom (Short dir.)", diaCross, spCross, n2, cut2, main2, leg),
    ];
  }, [footing, shapeCodes]);

  const tieBeamRows = useMemo(() => {
    const { Ln, b, D, cover, nBot, diaBot, nTop, diaTop, Ld, diaTie, spTie, hook } = tieBeam;
    const cutMain = Ln + 2 * Ld, bx = b - 2 * cover, dx = D - 2 * cover;
    const nStir = Math.ceil(Ln / spTie) + 1, cutStir = 2 * (bx + dx) + hook;
    const mk = (key, desc, dia, sp, n, cut, a, bdim) => {
      const totLen = (n * cut) / 1000, uw = unitWeight(dia), totW = totLen * uw;
      return { key, member: "Tie Beam", desc, code: shapeCodes[key], dia, spacing: sp, nBars: n, cutLen: cut, totLen, unitW: uw, totW, a, b: bdim };
    };
    return [
      mk("tieBeamBottom", "Tie Beam Bottom Bars", diaBot, "-", nBot, cutMain, Ln, Ld),
      mk("tieBeamTop", "Tie Beam Top Bars", diaTop, "-", nTop, cutMain, Ln, Ld),
      mk("tieBeamStirrup", "Tie Beam Stirrups (rect. hoop)", diaTie, spTie, nStir, cutStir, bx, dx),
    ];
  }, [tieBeam, shapeCodes]);

  const columnRows = useMemo(() => {
    const { H, b, D, cover, nBars, diaMain, Ld, diaTie, spTie, hook } = column;
    const cutMain = H + Ld, bx = b - 2 * cover, dx = D - 2 * cover;
    const nTies = Math.ceil(H / spTie) + 1, cutTie = 2 * (bx + dx) + hook;
    const mk = (key, desc, dia, sp, n, cut, a, bdim) => {
      const totLen = (n * cut) / 1000, uw = unitWeight(dia), totW = totLen * uw;
      return { key, member: "Column", desc, code: shapeCodes[key], dia, spacing: sp, nBars: n, cutLen: cut, totLen, unitW: uw, totW, a, b: bdim };
    };
    return [
      mk("columnMain", "Column Main Vertical Bars", diaMain, "-", nBars, cutMain, H, Ld),
      mk("columnTie", "Column Ties / Stirrups (rect. hoop)", diaTie, spTie, nTies, cutTie, bx, dx),
    ];
  }, [column, shapeCodes]);

  const beamRows = useMemo(() => {
    const { Ln, b, D, cover, nBot, diaBot, nTop, diaTop, Ld, diaTie, spTie, hook } = beam;
    const cutMain = Ln + 2 * Ld, bx = b - 2 * cover, dx = D - 2 * cover;
    const nStir = Math.ceil(Ln / spTie) + 1, cutStir = 2 * (bx + dx) + hook;
    const mk = (key, desc, dia, sp, n, cut, a, bdim) => {
      const totLen = (n * cut) / 1000, uw = unitWeight(dia), totW = totLen * uw;
      return { key, member: "Beam", desc, code: shapeCodes[key], dia, spacing: sp, nBars: n, cutLen: cut, totLen, unitW: uw, totW, a, b: bdim };
    };
    return [
      mk("beamBottom", "Beam Bottom Main Bars", diaBot, "-", nBot, cutMain, Ln, Ld),
      mk("beamTop", "Beam Top Main Bars", diaTop, "-", nTop, cutMain, Ln, Ld),
      mk("beamStirrup", "Beam Stirrups (rect. hoop)", diaTie, spTie, nStir, cutStir, bx, dx),
    ];
  }, [beam, shapeCodes]);

  const slabRows = useMemo(() => {
    const { Lx, Ly, cover, diaMain, spMain, diaDist, spDist } = slab;
    const cut1 = Lx - 2 * cover, n1 = Math.ceil((Ly - 2 * cover) / spMain) + 1;
    const cut2 = Ly - 2 * cover, n2 = Math.ceil((Lx - 2 * cover) / spDist) + 1;
    const mk = (key, desc, dia, sp, n, cut) => {
      const totLen = (n * cut) / 1000, uw = unitWeight(dia), totW = totLen * uw;
      return { key, member: "Slab", desc, code: shapeCodes[key], dia, spacing: sp, nBars: n, cutLen: cut, totLen, unitW: uw, totW, a: cut, b: 0 };
    };
    return [
      mk("slabMain", "Slab Main Bars (Short span dir.)", diaMain, spMain, n1, cut1),
      mk("slabDist", "Slab Distribution Bars (Long dir.)", diaDist, spDist, n2, cut2),
    ];
  }, [slab, shapeCodes]);

  const allRows = useMemo(() => [...footingRows, ...tieBeamRows, ...columnRows, ...beamRows, ...slabRows],
    [footingRows, tieBeamRows, columnRows, beamRows, slabRows]);
  const memberTotals = useMemo(() => {
    const t = {}; allRows.forEach((r) => { t[r.member] = (t[r.member] || 0) + r.totW; }); return t;
  }, [allRows]);
  const grandTotal = useMemo(() => allRows.reduce((s, r) => s + r.totW, 0), [allRows]);

  /* ---------- per-member config: keeps each member's Save/Load/New fully independent ---------- */
  const getMemberConfig = (tabId) => {
    switch (tabId) {
      case "footing": return { label: "Footing", rows: footingRows, state: footing, setState: setFooting, defaultState: DEFAULT_FOOTING, shapeKeys: ["footingMain", "footingCross"] };
      case "tieBeam": return { label: "Tie Beam", rows: tieBeamRows, state: tieBeam, setState: setTieBeam, defaultState: DEFAULT_TIEBEAM, shapeKeys: ["tieBeamBottom", "tieBeamTop", "tieBeamStirrup"] };
      case "column": return { label: "Column", rows: columnRows, state: column, setState: setColumn, defaultState: DEFAULT_COLUMN, shapeKeys: ["columnMain", "columnTie"] };
      case "beam": return { label: "Beam", rows: beamRows, state: beam, setState: setBeam, defaultState: DEFAULT_BEAM, shapeKeys: ["beamBottom", "beamTop", "beamStirrup"] };
      case "slab": return { label: "Slab", rows: slabRows, state: slab, setState: setSlab, defaultState: DEFAULT_SLAB, shapeKeys: ["slabMain", "slabDist"] };
      default: return null;
    }
  };

  /* ---------- BBS record save / load / delete — kept SEPARATE per member, multiple entries allowed ---------- */
  const saveRecord = (tabId) => {
    const cfg = getMemberConfig(tabId);
    if (!cfg) return;
    const existingId = loadedRecordId[tabId];
    const id = existingId || uid("rec");
    const label = (recordLabel[tabId] || "").trim() || `${cfg.label} Calc ${new Date().toLocaleDateString()}`;
    const totalWeight = cfg.rows.reduce((s, r) => s + r.totW, 0);
    const shapeSubset = {};
    cfg.shapeKeys.forEach((k) => { shapeSubset[k] = shapeCodes[k]; });
    const rec = {
      id, tabId, memberType: cfg.label, label, savedAt: new Date().toISOString(),
      inputs: cfg.state, shapeCodes: shapeSubset, rows: cfg.rows, totalWeight,
    };
    setRecords((prev) => {
      const exists = prev.some((r) => r.id === id);
      const next = exists ? prev.map((r) => (r.id === id ? rec : r)) : [rec, ...prev];
      storageSet("bbs-app-records", next, true);
      return next;
    });
    setLoadedRecordId((prev) => ({ ...prev, [tabId]: id }));
    setRecordLabel((prev) => ({ ...prev, [tabId]: label }));
    setSaveMsg(`${cfg.label} ${existingId ? "record updated." : "saved to Site Records."}`);
    setTimeout(() => setSaveMsg(""), 2500);
    return rec;
  };

  const loadRecord = (rec) => {
    const cfg = getMemberConfig(rec.tabId);
    if (!cfg) return;
    cfg.setState(rec.inputs);
    setShapeCodes((prev) => ({ ...prev, ...rec.shapeCodes }));
    setLoadedRecordId((prev) => ({ ...prev, [rec.tabId]: rec.id }));
    setRecordLabel((prev) => ({ ...prev, [rec.tabId]: rec.label }));
    setSection("calculator");
    setCalcTab(rec.tabId);
  };

  const deleteRecord = (id) => {
    const rec = records.find((r) => r.id === id);
    setRecords((prev) => {
      const next = prev.filter((r) => r.id !== id);
      storageSet("bbs-app-records", next, true);
      return next;
    });
    if (rec && loadedRecordId[rec.tabId] === id) {
      setLoadedRecordId((p) => ({ ...p, [rec.tabId]: null }));
      setRecordLabel((p) => ({ ...p, [rec.tabId]: "" }));
    }
  };

  const newCalculation = (tabId) => {
    const cfg = getMemberConfig(tabId);
    if (!cfg) return;
    cfg.setState(cfg.defaultState);
    setLoadedRecordId((prev) => ({ ...prev, [tabId]: null }));
    setRecordLabel((prev) => ({ ...prev, [tabId]: "" }));
  };

  /* ---------- DPR CRUD ---------- */
  const saveDpr = () => {
    if (!dprForm.date || !dprForm.work.trim()) { setSaveMsg("Add at least a date and work description."); setTimeout(() => setSaveMsg(""), 2500); return; }
    const entry = { id: editingDprId || uid("dpr"), ...dprForm, savedAt: new Date().toISOString() };
    setDprEntries((prev) => {
      const exists = prev.some((e) => e.id === entry.id);
      const next = exists ? prev.map((e) => (e.id === entry.id ? entry : e)) : [entry, ...prev];
      storageSet("bbs-app-dpr", next, true);
      return next;
    });
    setEditingDprId(null);
    setDprForm(DEFAULT_DPR);
  };
  const editDpr = (entry) => { setDprForm({ date: entry.date, work: entry.work, manpower: entry.manpower, material: entry.material, weather: entry.weather, remarks: entry.remarks }); setEditingDprId(entry.id); };
  const cancelEditDpr = () => { setEditingDprId(null); setDprForm(DEFAULT_DPR); };
  const deleteDpr = (id) => {
    setDprEntries((prev) => {
      const next = prev.filter((e) => e.id !== id);
      storageSet("bbs-app-dpr", next, true);
      return next;
    });
    if (editingDprId === id) cancelEditDpr();
  };

  /* ---------- Excel export — formula-driven, mirrors the app's own math live in Excel ---------- */
  const exportExcel = useCallback(() => {
    try {
      const wb = XLSX.utils.book_new();
      const genLetter = [
        [],
        ["Project / Site Name:", profile ? profile.siteName : ""],
        ["Prepared By:", profile ? profile.name : ""],
        ["Date Generated:", new Date().toLocaleString()],
        ["Checked By:", ""],
        [],
      ];

      const totalRefs = {
        Footing: writeMemberFormulaSheet(wb, "Footing", "footing", footing, genLetter),
        "Tie Beam": writeMemberFormulaSheet(wb, "Tie Beam", "tieBeam", tieBeam, genLetter),
        Column: writeMemberFormulaSheet(wb, "Column", "column", column, genLetter),
        Beam: writeMemberFormulaSheet(wb, "Beam", "beam", beam, genLetter),
        Slab: writeMemberFormulaSheet(wb, "Slab", "slab", slab, genLetter),
      };

      // Summary sheet — live formulas pulling each member sheet's own total
      const sumAoa = [
        ["PROJECT BBS SUMMARY"], [],
        ["Project / Site Name:", profile ? profile.siteName : ""], [],
        ["Structural Member", "Total Weight (kg)"],
        ...Object.keys(totalRefs).map((m) => [m, ""]),
        [], ["GRAND TOTAL (kg)", ""],
      ];
      const wsSum = XLSX.utils.aoa_to_sheet(sumAoa);
      wsSum["!cols"] = [{ wch: 22 }, { wch: 18 }];
      const memberStartRow = 6; // 1-indexed row of first member in sumAoa
      Object.values(totalRefs).forEach((ref, i) => {
        wsSum[`B${memberStartRow + i}`] = { t: "n", f: ref, z: "0.00" };
      });
      const lastMemberRow = memberStartRow + Object.keys(totalRefs).length - 1;
      const grandRow = lastMemberRow + 2;
      wsSum[`B${grandRow}`] = { t: "n", f: `SUM(B${memberStartRow}:B${lastMemberRow})`, z: "0.00" };
      XLSX.utils.book_append_sheet(wb, wsSum, "Summary");
      // move Summary to the front
      wb.SheetNames.unshift(wb.SheetNames.pop());

      if (records.length) {
        const recAoa = [["SAVED SITE BBS RECORDS"], [], ["Member Type", "Label", "Saved On", "Total Weight (kg)"],
          ...records.map((r) => [r.memberType, r.label, new Date(r.savedAt).toLocaleString(), Number(r.totalWeight.toFixed(2))])];
        const wsRec = XLSX.utils.aoa_to_sheet(recAoa);
        wsRec["!cols"] = [{ wch: 12 }, { wch: 28 }, { wch: 20 }, { wch: 16 }];
        XLSX.utils.book_append_sheet(wb, wsRec, "Site BBS Records");
      }

      if (dprEntries.length) {
        const dprAoa = [["DAILY PROGRESS REPORTS"], [], ["Date", "Work Carried Out", "Manpower", "Material Used", "Weather", "Remarks"],
          ...dprEntries.map((e) => [e.date, e.work, e.manpower, e.material, e.weather, e.remarks])];
        const wsDpr = XLSX.utils.aoa_to_sheet(dprAoa);
        wsDpr["!cols"] = [{ wch: 12 }, { wch: 32 }, { wch: 18 }, { wch: 24 }, { wch: 10 }, { wch: 30 }];
        XLSX.utils.book_append_sheet(wb, wsDpr, "Daily Progress Reports");
      }

      const filename = `BBS_${(profile && profile.siteName ? profile.siteName.replace(/\s+/g, "_") : "Export")}.xlsx`;
      const b64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
      const dataUri = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${b64}`;
      setDownloadReady({ url: dataUri, filename });
      setSaveMsg("File ready below — tap the green link to save it.");
      setTimeout(() => setSaveMsg(""), 3500);
    } catch (e) {
      console.error("Export failed:", e);
      setSaveMsg("Export failed — see console for details.");
      setTimeout(() => setSaveMsg(""), 3000);
    }
  }, [footing, tieBeam, column, beam, slab, records, dprEntries, profile]);

  /* ---------- Export a SINGLE saved record straight to Excel (one click, right from the record card) ---------- */
  const exportSingleRecordExcel = useCallback((rec) => {
    try {
      const wb = XLSX.utils.book_new();
      const letterLines = [
        [],
        ["Project / Site Name:", profile ? profile.siteName : ""],
        ["Record Label:", rec.label],
        ["Prepared By:", profile ? profile.name : ""],
        ["Saved On:", new Date(rec.savedAt).toLocaleString()],
        [],
      ];
      writeMemberFormulaSheet(wb, rec.memberType, rec.tabId, rec.inputs, letterLines);

      const filename = `${rec.memberType}_${rec.label.replace(/\s+/g, "_")}.xlsx`;
      const b64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
      const dataUri = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${b64}`;
      setDownloadReady({ url: dataUri, filename });
      setSaveMsg("File ready below — tap the green link to save it.");
      setTimeout(() => setSaveMsg(""), 3500);
    } catch (e) {
      console.error("Record export failed:", e);
      setSaveMsg("Export failed for this record — see console.");
      setTimeout(() => setSaveMsg(""), 3000);
    }
  }, [profile]);

  const saveAndExportRecord = (tabId) => {
    const rec = saveRecord(tabId);
    if (rec) exportSingleRecordExcel(rec);
  };

  /* ---------- render guards ---------- */
  if (booting) {
    return (
      <div className="bbs-root min-h-screen flex items-center justify-center">
        <Theme />
        <div className="bbs-mono text-sm" style={{ color: "#BFE0F5" }}>Loading site data…</div>
      </div>
    );
  }
  if (!profile) return <LoginScreen onLogin={handleLogin} />;

  const CALC_TABS = [
    { id: "footing", n: "01", label: "Footing" }, { id: "tieBeam", n: "02", label: "Tie Beam" },
    { id: "column", n: "03", label: "Column" }, { id: "beam", n: "04", label: "Beam" },
    { id: "slab", n: "05", label: "Slab" }, { id: "summary", n: "06", label: "Summary" },
  ];
  const SECTIONS = [
    { id: "calculator", label: "BBS Calculator", icon: CalcIcon },
    { id: "records", label: `Site Records (${records.length})`, icon: FolderOpen },
    { id: "dpr", label: `Daily Progress (${dprEntries.length})`, icon: ClipboardList },
  ];

  function BBSTable({ rows }) {
    return (
      <div className="bbs-card overflow-x-auto scrollbar-thin">
        <table className="w-full text-xs" style={{ borderCollapse: "collapse" }}>
          <thead><tr>{["#", "Description", "Shape Code", "Shape", "Dia", "Spacing", "No.", "Cut Len (mm)", "Len (m)", "Wt/m", "Total (kg)"].map((h) => (<th key={h} className="bbs-th px-2 py-2 text-left whitespace-nowrap">{h}</th>))}</tr></thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={r.key} className="bbs-td" style={{ borderLeft: `3px solid ${MEMBER_COLORS[r.member]}` }}>
                <td className="px-2 py-2 bbs-mono">{i + 1}</td>
                <td className="px-2 py-2">{r.desc}</td>
                <td className="px-2 py-2" style={{ minWidth: 150 }}><ShapeCodeField value={r.code} onChange={(v) => setCode(r.key, v)} /></td>
                <td className="px-2 py-2"><BarShapeSVG code={r.code} a={r.a} b={r.b} /></td>
                <td className="px-2 py-2 bbs-mono">{r.dia}</td>
                <td className="px-2 py-2 bbs-mono">{r.spacing}</td>
                <td className="px-2 py-2 bbs-mono">{r.nBars}</td>
                <td className="px-2 py-2 bbs-mono">{Math.round(r.cutLen)}</td>
                <td className="px-2 py-2 bbs-mono">{r.totLen.toFixed(2)}</td>
                <td className="px-2 py-2 bbs-mono">{r.unitW.toFixed(3)}</td>
                <td className="px-2 py-2 bbs-mono font-bold" style={{ color: "#F2C14E" }}>{r.totW.toFixed(2)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
  function SectionTotal({ rows, label }) {
    const t = rows.reduce((s, r) => s + r.totW, 0);
    return (<div className="flex justify-end mt-2"><div className="bbs-card px-4 py-2 text-sm bbs-mono" style={{ borderColor: "#E2793D" }}>{label} total: <span style={{ color: "#F2C14E", fontWeight: 700 }}>{t.toFixed(2)} kg</span></div></div>);
  }

  function FootingPanel() { const s = footing, u = (p) => setFooting({ ...s, ...p }); return (
    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
      <NumField label="Length L (mm)" value={s.L} onChange={(v) => u({ L: v })} />
      <NumField label="Width W (mm)" value={s.W} onChange={(v) => u({ W: v })} />
      <NumField label="Depth D (mm)" value={s.D} onChange={(v) => u({ D: v })} />
      <NumField label="Clear Cover (mm)" value={s.cover} onChange={(v) => u({ cover: v })} />
      <NumField label="Dia Main (mm)" value={s.diaMain} onChange={(v) => u({ diaMain: v })} />
      <NumField label="Dia Cross (mm)" value={s.diaCross} onChange={(v) => u({ diaCross: v })} />
      <NumField label="Spacing Main (mm)" value={s.spMain} onChange={(v) => u({ spMain: v })} />
      <NumField label="Spacing Cross (mm)" value={s.spCross} onChange={(v) => u({ spCross: v })} />
    </div>); }
  function TieBeamPanel() { const s = tieBeam, u = (p) => setTieBeam({ ...s, ...p }); return (
    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
      <NumField label="Clear Span Ln (mm)" value={s.Ln} onChange={(v) => u({ Ln: v })} />
      <NumField label="Width b (mm)" value={s.b} onChange={(v) => u({ b: v })} />
      <NumField label="Depth D (mm)" value={s.D} onChange={(v) => u({ D: v })} />
      <NumField label="Clear Cover (mm)" value={s.cover} onChange={(v) => u({ cover: v })} />
      <NumField label="No. Bottom Bars" value={s.nBot} onChange={(v) => u({ nBot: v })} />
      <NumField label="Dia Bottom (mm)" value={s.diaBot} onChange={(v) => u({ diaBot: v })} />
      <NumField label="No. Top Bars" value={s.nTop} onChange={(v) => u({ nTop: v })} />
      <NumField label="Dia Top (mm)" value={s.diaTop} onChange={(v) => u({ diaTop: v })} />
      <NumField label="Anchorage Ld (mm)" value={s.Ld} onChange={(v) => u({ Ld: v })} />
      <NumField label="Stirrup Dia (mm)" value={s.diaTie} onChange={(v) => u({ diaTie: v })} />
      <NumField label="Stirrup Spacing (mm)" value={s.spTie} onChange={(v) => u({ spTie: v })} />
      <NumField label="Hook Allowance (mm)" value={s.hook} onChange={(v) => u({ hook: v })} />
    </div>); }
  function ColumnPanel() { const s = column, u = (p) => setColumn({ ...s, ...p }); return (
    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
      <NumField label="Clear Height (mm)" value={s.H} onChange={(v) => u({ H: v })} />
      <NumField label="Width b (mm)" value={s.b} onChange={(v) => u({ b: v })} />
      <NumField label="Depth D (mm)" value={s.D} onChange={(v) => u({ D: v })} />
      <NumField label="Clear Cover (mm)" value={s.cover} onChange={(v) => u({ cover: v })} />
      <NumField label="No. Main Bars" value={s.nBars} onChange={(v) => u({ nBars: v })} />
      <NumField label="Dia Main (mm)" value={s.diaMain} onChange={(v) => u({ diaMain: v })} />
      <NumField label="Dev. Length Ld (mm)" value={s.Ld} onChange={(v) => u({ Ld: v })} />
      <NumField label="Tie Dia (mm)" value={s.diaTie} onChange={(v) => u({ diaTie: v })} />
      <NumField label="Tie Spacing (mm)" value={s.spTie} onChange={(v) => u({ spTie: v })} />
      <NumField label="Hook Allowance (mm)" value={s.hook} onChange={(v) => u({ hook: v })} />
    </div>); }
  function BeamPanel() { const s = beam, u = (p) => setBeam({ ...s, ...p }); return (
    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
      <NumField label="Clear Span Ln (mm)" value={s.Ln} onChange={(v) => u({ Ln: v })} />
      <NumField label="Width b (mm)" value={s.b} onChange={(v) => u({ b: v })} />
      <NumField label="Depth D (mm)" value={s.D} onChange={(v) => u({ D: v })} />
      <NumField label="Clear Cover (mm)" value={s.cover} onChange={(v) => u({ cover: v })} />
      <NumField label="No. Bottom Bars" value={s.nBot} onChange={(v) => u({ nBot: v })} />
      <NumField label="Dia Bottom (mm)" value={s.diaBot} onChange={(v) => u({ diaBot: v })} />
      <NumField label="No. Top Bars" value={s.nTop} onChange={(v) => u({ nTop: v })} />
      <NumField label="Dia Top (mm)" value={s.diaTop} onChange={(v) => u({ diaTop: v })} />
      <NumField label="Anchorage Ld (mm)" value={s.Ld} onChange={(v) => u({ Ld: v })} />
      <NumField label="Stirrup Dia (mm)" value={s.diaTie} onChange={(v) => u({ diaTie: v })} />
      <NumField label="Stirrup Spacing (mm)" value={s.spTie} onChange={(v) => u({ spTie: v })} />
      <NumField label="Hook Allowance (mm)" value={s.hook} onChange={(v) => u({ hook: v })} />
    </div>); }
  function SlabPanel() { const s = slab, u = (p) => setSlab({ ...s, ...p }); return (
    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
      <NumField label="Short Span Lx (mm)" value={s.Lx} onChange={(v) => u({ Lx: v })} />
      <NumField label="Long Span Ly (mm)" value={s.Ly} onChange={(v) => u({ Ly: v })} />
      <NumField label="Clear Cover (mm)" value={s.cover} onChange={(v) => u({ cover: v })} />
      <NumField label="Main Dia (mm)" value={s.diaMain} onChange={(v) => u({ diaMain: v })} />
      <NumField label="Main Spacing (mm)" value={s.spMain} onChange={(v) => u({ spMain: v })} />
      <NumField label="Dist. Dia (mm)" value={s.diaDist} onChange={(v) => u({ diaDist: v })} />
      <NumField label="Dist. Spacing (mm)" value={s.spDist} onChange={(v) => u({ spDist: v })} />
    </div>); }

  const maxMember = Math.max(...Object.values(memberTotals), 1);

  return (
    <div className="bbs-root p-4 md:p-6">
      <Theme />

      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3 mb-4">
        <div>
          <div className="bbs-mono text-xs tracking-widest" style={{ color: "#E2793D" }}>{profile.siteName.toUpperCase()}</div>
          <h1 className="bbs-display text-2xl md:text-3xl font-bold flex items-center gap-2">
            <Ruler size={22} style={{ color: "#F2C14E" }} /> BBS Site App
          </h1>
          <div className="text-xs bbs-mono mt-1" style={{ color: "#6E9BBE" }}>{profile.name} · {profile.email}</div>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={exportExcel} className="bbs-btn bbs-btn-primary"><Download size={15} /> Export Excel</button>
          <button onClick={handleLogout} className="bbs-btn bbs-btn-ghost"><LogOut size={15} /> Logout</button>
        </div>
      </div>

      {saveMsg && (<div className="bbs-card px-3 py-2 text-xs mb-3 bbs-mono" style={{ color: "#7FD6A6", borderColor: "#7FD6A6" }}>{saveMsg}</div>)}

      <div className="bbs-card px-3 py-2 text-xs mb-4 bbs-mono" style={{ color: "#F2C14E", borderColor: "#F2C14E" }}>
        Shared mode: Site Records and Daily Progress here are visible to everyone who opens this app.
      </div>

      {downloadReady && (
        <div className="bbs-card p-4 mb-4" style={{ borderColor: "#7FD6A6", borderWidth: 2 }}>
          <div className="text-sm font-semibold mb-2">Your file is ready — tap to save it:</div>
          <a
            href={downloadReady.url}
            download={downloadReady.filename}
            className="bbs-btn bbs-btn-primary"
            style={{ textDecoration: "none", background: "#7FD6A6", color: "#0B2E4A" }}
          >
            <Download size={15} /> {downloadReady.filename}
          </a>
          <div className="text-xs mt-2" style={{ color: "#BFE0F5" }}>
            If tapping doesn't save automatically, press and hold this link, then choose "Download link" or "Save link as."
          </div>
          <button onClick={() => setDownloadReady(null)} className="bbs-btn bbs-btn-ghost text-xs px-2 py-1 mt-2">Dismiss</button>
        </div>
      )}

      {/* Section nav */}
      <div className="flex flex-wrap gap-2 mb-5">
        {SECTIONS.map((s) => {
          const Icon = s.icon;
          return (
            <button key={s.id} onClick={() => setSection(s.id)} className={`bbs-tab px-3 py-2 rounded text-sm bbs-mono flex items-center gap-2 ${section === s.id ? "active" : ""}`}>
              <Icon size={14} /> {s.label}
            </button>
          );
        })}
      </div>

      {/* ---------------- CALCULATOR SECTION ---------------- */}
      {section === "calculator" && (
        <div>
          <div className="flex flex-wrap gap-2 mb-4">
            {CALC_TABS.map((t) => (
              <button key={t.id} onClick={() => setCalcTab(t.id)} className={`bbs-tab px-3 py-2 rounded text-sm bbs-mono ${calcTab === t.id ? "active" : ""}`}>
                <span style={{ opacity: 0.7 }}>{t.n}</span> &nbsp;{t.label}
              </button>
            ))}
          </div>

          {calcTab !== "summary" && (
            <div className="bbs-card p-4 mb-4 flex flex-col md:flex-row gap-3 items-start md:items-end">
              <div className="flex-1 w-full">
                <TextField
                  label={loadedRecordId[calcTab] ? `Editing ${getMemberConfig(calcTab).label} record — update label if needed` : `Label this ${getMemberConfig(calcTab).label} calculation`}
                  value={recordLabel[calcTab] || ""}
                  onChange={(v) => setRecordLabel((prev) => ({ ...prev, [calcTab]: v }))}
                  placeholder={`e.g. ${getMemberConfig(calcTab).label} F1 – Block A`}
                />
              </div>
              <div className="flex gap-2 flex-wrap">
                <button onClick={() => saveRecord(calcTab)} className="bbs-btn bbs-btn-primary"><Save size={15} /> {loadedRecordId[calcTab] ? "Update Record" : "Save to Site Records"}</button>
                <button onClick={() => saveAndExportRecord(calcTab)} className="bbs-btn bbs-btn-primary"><Download size={15} /> Save &amp; Download Excel</button>
                <button onClick={() => newCalculation(calcTab)} className="bbs-btn bbs-btn-ghost"><Plus size={15} /> New</button>
              </div>
            </div>
          )}

          {calcTab === "footing" && (<div><div className="bbs-card p-4 mb-4"><FootingPanel /></div><BBSTable rows={footingRows} /><SectionTotal rows={footingRows} label="Footing" /></div>)}
          {calcTab === "tieBeam" && (<div><div className="bbs-card p-4 mb-4"><TieBeamPanel /></div><BBSTable rows={tieBeamRows} /><SectionTotal rows={tieBeamRows} label="Tie Beam" /></div>)}
          {calcTab === "column" && (<div><div className="bbs-card p-4 mb-4"><ColumnPanel /></div><BBSTable rows={columnRows} /><SectionTotal rows={columnRows} label="Column" /></div>)}
          {calcTab === "beam" && (<div><div className="bbs-card p-4 mb-4"><BeamPanel /></div><BBSTable rows={beamRows} /><SectionTotal rows={beamRows} label="Beam" /></div>)}
          {calcTab === "slab" && (<div><div className="bbs-card p-4 mb-4"><SlabPanel /></div><BBSTable rows={slabRows} /><SectionTotal rows={slabRows} label="Slab" /></div>)}
          {calcTab === "summary" && (
            <div className="bbs-card p-5">
              <h2 className="bbs-display text-lg font-bold mb-4">Calculation Summary</h2>
              <div className="space-y-3 mb-6">
                {Object.entries(memberTotals).map(([m, w]) => (
                  <div key={m} className="flex items-center gap-3">
                    <div className="w-24 text-xs bbs-mono">{m}</div>
                    <div className="flex-1 h-5 rounded" style={{ background: "rgba(191,224,245,0.1)" }}>
                      <div className="h-5 rounded" style={{ width: `${(w / maxMember) * 100}%`, background: MEMBER_COLORS[m] }} />
                    </div>
                    <div className="w-24 text-right bbs-mono text-sm font-bold">{w.toFixed(2)} kg</div>
                  </div>
                ))}
              </div>
              <div className="border-t pt-4 flex justify-between items-center" style={{ borderColor: "rgba(191,224,245,0.2)" }}>
                <span className="bbs-display text-lg font-bold">Grand Total</span>
                <span className="bbs-mono text-2xl font-bold" style={{ color: "#F2C14E" }}>{grandTotal.toFixed(2)} kg</span>
              </div>
              <p className="text-xs mt-4 bbs-mono" style={{ color: "#6E9BBE" }}>
                This reflects whatever is currently in each tab. Saved Site Records are kept separately per member — see the Site Records section.
              </p>
            </div>
          )}
        </div>
      )}

      {/* ---------------- SITE RECORDS SECTION (grouped separately per member) ---------------- */}
      {section === "records" && (
        <div className="bbs-card p-5">
          <h2 className="bbs-display text-lg font-bold mb-1">Site BBS Records</h2>
          <p className="text-xs mb-5" style={{ color: "#6E9BBE" }}>
            Kept separate per member — calculate Footing many times and every save adds a new Footing entry here, without touching Tie Beam / Column / Beam / Slab. Load to view/edit, or delete anytime.
          </p>
          {records.length === 0 && (<div className="text-sm bbs-mono" style={{ color: "#6E9BBE" }}>No records saved yet. Go to the BBS Calculator, pick a tab, and click "Save to Site Records".</div>)}
          {["Footing", "Tie Beam", "Column", "Beam", "Slab"].map((mt) => {
            const list = records.filter((r) => r.memberType === mt);
            return (
              <div key={mt} className="mb-6">
                <h3 className="bbs-display text-sm font-bold mb-2 flex items-center gap-2">
                  <span style={{ width: 10, height: 10, borderRadius: 2, background: MEMBER_COLORS[mt], display: "inline-block" }} />
                  {mt} <span style={{ color: "#6E9BBE", fontWeight: 400 }}>({list.length})</span>
                </h3>
                {list.length === 0 ? (
                  <div className="text-xs bbs-mono mb-1" style={{ color: "#6E9BBE" }}>No {mt} records yet.</div>
                ) : (
                  <div className="space-y-2">
                    {list.map((r) => (
                      <div key={r.id} className="bbs-card p-3 flex flex-col md:flex-row md:items-center gap-2 md:gap-4" style={{ borderLeft: `3px solid ${MEMBER_COLORS[mt]}` }}>
                        <div className="flex-1">
                          <div className="font-semibold text-sm">{r.label}</div>
                          <div className="text-xs bbs-mono" style={{ color: "#6E9BBE" }}>{new Date(r.savedAt).toLocaleString()}</div>
                        </div>
                        <div className="bbs-mono text-sm font-bold" style={{ color: "#F2C14E" }}>{r.totalWeight.toFixed(2)} kg</div>
                        <div className="flex gap-2">
                          <button onClick={() => exportSingleRecordExcel(r)} className="bbs-btn bbs-btn-primary text-xs px-2 py-1"><Download size={13} /> Excel</button>
                          <button onClick={() => loadRecord(r)} className="bbs-btn bbs-btn-ghost text-xs px-2 py-1"><Pencil size={13} /> Load / Edit</button>
                          <button onClick={() => deleteRecord(r.id)} className="bbs-btn bbs-btn-danger text-xs px-2 py-1"><Trash2 size={13} /> Delete</button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* ---------------- DAILY PROGRESS REPORT SECTION ---------------- */}
      {section === "dpr" && (
        <div>
          <div className="bbs-card p-4 mb-4">
            <h2 className="bbs-display text-lg font-bold mb-3">{editingDprId ? "Edit Entry" : "New Daily Progress Entry"}</h2>
            <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-3">
              <TextField label="Date" type="date" value={dprForm.date} onChange={(v) => setDprForm({ ...dprForm, date: v })} />
              <TextField label="Manpower (e.g. 10 mason, 5 helper)" value={dprForm.manpower} onChange={(v) => setDprForm({ ...dprForm, manpower: v })} />
              <label className="flex flex-col gap-1 text-xs">
                <span className="bbs-mono" style={{ color: "#BFE0F5", opacity: 0.85 }}>Weather</span>
                <select value={dprForm.weather} onChange={(e) => setDprForm({ ...dprForm, weather: e.target.value })} className="bbs-input px-2 py-1.5 text-sm">
                  <option>Sunny</option><option>Cloudy</option><option>Rainy</option><option>Windy</option><option>Hot</option>
                </select>
              </label>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
              <TextField label="Work Carried Out" value={dprForm.work} onChange={(v) => setDprForm({ ...dprForm, work: v })} placeholder="e.g. Footing F1-F4 concreting completed" />
              <TextField label="Material Used" value={dprForm.material} onChange={(v) => setDprForm({ ...dprForm, material: v })} placeholder="e.g. 12 bags cement, 2 ton steel" />
            </div>
            <TextField label="Remarks / Issues" value={dprForm.remarks} onChange={(v) => setDprForm({ ...dprForm, remarks: v })} placeholder="Any delays, site issues, instructions" />
            <div className="flex gap-2 mt-3">
              <button onClick={saveDpr} className="bbs-btn bbs-btn-primary"><Save size={15} /> {editingDprId ? "Update Entry" : "Add Entry"}</button>
              {editingDprId && (<button onClick={cancelEditDpr} className="bbs-btn bbs-btn-ghost">Cancel</button>)}
            </div>
          </div>

          <div className="bbs-card p-4">
            <h3 className="bbs-display text-base font-bold mb-3">Log ({dprEntries.length})</h3>
            {dprEntries.length === 0 && (<div className="text-sm bbs-mono" style={{ color: "#6E9BBE" }}>No entries yet.</div>)}
            <div className="space-y-2">
              {dprEntries.map((e) => (
                <div key={e.id} className="bbs-card p-3" style={{ borderLeft: "3px solid #F2C14E" }}>
                  <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-2">
                    <div>
                      <div className="bbs-mono text-xs" style={{ color: "#F2C14E" }}>{e.date} · {e.weather}</div>
                      <div className="text-sm font-semibold mt-0.5">{e.work}</div>
                      {e.manpower && <div className="text-xs mt-1" style={{ color: "#BFE0F5" }}>Manpower: {e.manpower}</div>}
                      {e.material && <div className="text-xs" style={{ color: "#BFE0F5" }}>Material: {e.material}</div>}
                      {e.remarks && <div className="text-xs mt-1" style={{ color: "#6E9BBE" }}>{e.remarks}</div>}
                    </div>
                    <div className="flex gap-2">
                      <button onClick={() => editDpr(e)} className="bbs-btn bbs-btn-ghost text-xs px-2 py-1"><Pencil size={13} /> Edit</button>
                      <button onClick={() => deleteDpr(e.id)} className="bbs-btn bbs-btn-danger text-xs px-2 py-1"><Trash2 size={13} /> Delete</button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
