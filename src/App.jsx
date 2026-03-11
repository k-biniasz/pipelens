import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";

const ACCENT  = "#E8FF47";
const ACCENT2 = "#FF4D6D";
const ACCENT3 = "#00C9A7";
const ACCENT4 = "#FF922B";
const ACCENT5 = "#CC5DE8";
const DIM_COLORS = ["#E8FF47","#FF4D6D","#00C9A7","#4D9FFF","#FF922B","#CC5DE8","#F06595","#63E6BE","#74C0FC","#A9E34B"];

const FIELD_ALIASES = {
  converted:          ["converted","is_converted","isconverted","opportunity","opp","converted_to_opp","convert","won"],
  lead_status:        ["lead_status","leadstatus","lead status","status"],
  ae_owner:           ["ae","ae name","ae_name","aename","account executive","account_executive","accountexecutive","owner","assigned to","assigned_to","rep","sales rep","sales_rep","salesrep"],
  job_role:           ["job_role","jobrole","role","job role"],
  job_title:          ["job_title","jobtitle","title","job title","contact_title"],
  segment:            ["segment","market_segment","marketsegment","account_segment"],
  industry:           ["industry","industry_vertical"],
  vertical:           ["vertical","sector"],
  company_size:       ["company_size","companysize","employee_count","employees","size","headcount"],
  region:             ["region","geo","geography","territory"],
  lead_source:        ["lead_source","leadsource","source","marketing_channel","channel"],
  account_type:       ["account_type","accounttype","type","customer_type"],
  tcp_regional:       ["tcp (for na&emea&apac)","tcp(fornaemeaapac)","tcp for na&emea&apac","tcp naemeaapac","tcp (na&emea&apac)","tcp"],
  mysql_family:       ["mysql family capture","mysqlfamilycapture","mysql family","mysqlfamily","mysql_family_capture","mysql_family"],
  cloud_vendor:       ["cloud vendor","cloudvendor","cloud_vendor","cloud provider","cloudprovider"],
  existing_database:  ["existing database","existingdatabase","existing_database","current database","current_database","db","database"],
  iqm_notes:          ["iqm_notes","iqm notes","iqmnotes","initial_qualifying_meeting","iqm"],
  iqm_notes_regional: ["iqm note (na&emea&apac)","iqm notes (na&emea&apac)","iqm note (na & emea & apac)","iqm note na&emea&apac","iqm note naemeaapac"],
  sdr_notes:          ["sdr_notes","sdr notes","sdrnotes","sdr_feedback","sdr"],
  sdr_notes_regional: ["sdr note (na&emea&apac)","sdr notes (na&emea&apac)","sdr note (na & emea & apac)","sdr note na&emea&apac","sdr note naemeaapac"],
  budget_notes:       ["budget_notes","budget notes","budgetnotes","budget"],
  authority_notes:    ["authority_notes","authority notes","authoritynotes","authority","decision_maker_notes"],
  need_notes:         ["need_notes","need notes","neednotes","need","pain_notes","pain","primary need","primary_need","primary need/pain","primary need / pain"],
  timing_notes:       ["timing_notes","timing notes","timingnotes","timing","timeline_notes"],
  team_notes:         ["team_notes","team notes","teamnotes","team","stakeholder_notes"],
  ae_feedback:        ["ae_feedback","ae feedback","aefeedback","ae_notes","ae notes","account_executive_notes"],
  iqm_score:          ["iqm_score","iqm score","iqmscore","iqm quality score","iqm_quality_score","iqm quality","meeting quality score","meeting score","iqm rating","iqm_rating","meeting quality"],
  ae_score:           ["ae_score","ae score","aescore","ae quality score","ae_quality_score","ae feedback score","ae_feedback_score","ae rating","ae_rating","account executive score"],
  next_steps:         ["next steps","nextsteps","next_steps","next step","action items","action_items","follow up","followup"],
  linkedin_url:       ["linkedin_url","linkedin url","linkedin","linkedin_profile","profile_url"],
  linkedin_summary:   ["linkedin_summary","linkedin summary","about","bio","profile_summary","linkedin_about"],
  linkedin_skills:    ["linkedin_skills","linkedin skills","skills","member_skills"],
  linkedin_history:   ["linkedin_history","job_history","work_history","experience","linkedin_experience"],
  linkedin_headline:  ["linkedin_headline","headline","linkedin_title"],
  account_name:       ["account_name","accountname","account name","company","company_name","companyname","organization","org","account"],
};

const NOTE_FIELDS = ["iqm_notes","iqm_notes_regional","sdr_notes","sdr_notes_regional","budget_notes","authority_notes","need_notes","timing_notes","team_notes","ae_feedback","next_steps"];
const NOTE_LABELS = {
  iqm_notes:"IQM Notes", iqm_notes_regional:"IQM Note (NA/EMEA/APAC)",
  sdr_notes:"SDR Notes", sdr_notes_regional:"SDR Note (NA/EMEA/APAC)",
  budget_notes:"Budget Notes", authority_notes:"Authority Notes", need_notes:"Need Notes",
  timing_notes:"Timing Notes", team_notes:"Team Notes", ae_feedback:"AE Feedback", next_steps:"Next Steps",
};
const STUCK_LABELS = {
  iqm_notes:"IQM Notes", iqm_notes_regional:"IQM Note (NA/EMEA/APAC)",
  sdr_notes:"SDR Notes", sdr_notes_regional:"SDR Note (NA/EMEA/APAC)",
  budget_notes:"Budget Notes", authority_notes:"Authority Notes",
  need_notes:"Need / Pain Notes", timing_notes:"Timing Notes",
  team_notes:"Team Notes", ae_feedback:"AE Feedback", next_steps:"Next Steps",
};
const STUCK_STATUSES = ["nurture - marketing","nurture - sales","nurture","engaged","meeting completed","working","open","contacted","in progress"];
const LINKEDIN_FIELDS = ["linkedin_url","linkedin_summary","linkedin_skills","linkedin_history","linkedin_headline"];
const DIM_KEYS = ["job_role","job_title","segment","industry","vertical","company_size","region","lead_source","account_type","tcp_regional","mysql_family","cloud_vendor","existing_database"];
const PROFILE_FIELDS = [
  {key:"job_title",label:"Job Title"},{key:"job_role",label:"Job Role"},
  {key:"industry",label:"Industry"},{key:"vertical",label:"Vertical"},
  {key:"tcp_regional",label:"TCP Regional"},{key:"mysql_family",label:"MySQL Family"},
  {key:"cloud_vendor",label:"Cloud Vendor"},{key:"existing_database",label:"Existing DB"},
  {key:"segment",label:"Segment"},{key:"region",label:"Region"},
];

// ── Markdown renderer ────────────────────────────────────────────────────────
function MarkdownBlock({ content, accent }) {
  if (!content) return null;
  const ac = accent || ACCENT3;

  const lines = content.split("\n");
  const elements = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // H1
    if (line.startsWith("# ")) {
      elements.push(
        <h2 key={i} style={{fontSize:17,fontWeight:800,color:"#fff",margin:"24px 0 8px",borderBottom:"1px solid "+ac+"44",paddingBottom:6}}>
          {line.slice(2)}
        </h2>
      );
    }
    // H2
    else if (line.startsWith("## ")) {
      elements.push(
        <h3 key={i} style={{fontSize:14,fontWeight:700,color:ac,margin:"20px 0 6px",letterSpacing:"0.3px"}}>
          {line.slice(3)}
        </h3>
      );
    }
    // H3 / numbered section header like "1. TITLE"
    else if (line.startsWith("### ") || /^\d+\.\s+[A-Z][A-Z\s]+$/.test(line.trim())) {
      const txt = line.startsWith("### ") ? line.slice(4) : line.trim();
      elements.push(
        <div key={i} style={{display:"flex",alignItems:"center",gap:8,margin:"20px 0 8px"}}>
          <div style={{width:3,height:16,background:ac,borderRadius:2,flexShrink:0}}/>
          <span style={{fontSize:12,fontWeight:700,color:ac,textTransform:"uppercase",letterSpacing:"1.5px"}}>{txt}</span>
        </div>
      );
    }
    // Bullet
    else if (/^[-*]\s/.test(line)) {
      const items = [];
      while (i < lines.length && /^[-*]\s/.test(lines[i])) {
        items.push(lines[i].replace(/^[-*]\s/, ""));
        i++;
      }
      elements.push(
        <ul key={"ul"+i} style={{margin:"4px 0 10px",paddingLeft:20,listStyle:"none"}}>
          {items.map((it, idx) => (
            <li key={idx} style={{fontSize:13,color:"#bbb",lineHeight:1.75,position:"relative",paddingLeft:14}}>
              <span style={{position:"absolute",left:0,color:ac}}>›</span>
              <InlineMarkdown text={it}/>
            </li>
          ))}
        </ul>
      );
      continue;
    }
    // Numbered list
    else if (/^\d+\.\s/.test(line) && !/^[0-9]+\.\s+[A-Z][A-Z\s]+$/.test(line.trim())) {
      const items = [];
      while (i < lines.length && /^\d+\.\s/.test(lines[i])) {
        items.push(lines[i].replace(/^\d+\.\s/, ""));
        i++;
      }
      elements.push(
        <ol key={"ol"+i} style={{margin:"4px 0 10px",paddingLeft:20}}>
          {items.map((it, idx) => (
            <li key={idx} style={{fontSize:13,color:"#bbb",lineHeight:1.75,marginBottom:2}}>
              <InlineMarkdown text={it}/>
            </li>
          ))}
        </ol>
      );
      continue;
    }
    // Bold-colon label lines like "**Label:** value"
    else if (/^\*\*[^*]+\*\*:/.test(line)) {
      elements.push(
        <p key={i} style={{fontSize:13,color:"#bbb",lineHeight:1.75,margin:"3px 0"}}>
          <InlineMarkdown text={line}/>
        </p>
      );
    }
    // Horizontal rule
    else if (/^---+$/.test(line.trim())) {
      elements.push(<hr key={i} style={{border:"none",borderTop:"1px solid rgba(255,255,255,0.08)",margin:"16px 0"}}/>);
    }
    // Empty line
    else if (line.trim() === "") {
      elements.push(<div key={i} style={{height:6}}/>);
    }
    // Normal paragraph
    else {
      elements.push(
        <p key={i} style={{fontSize:13,color:"#ccc",lineHeight:1.8,margin:"3px 0"}}>
          <InlineMarkdown text={line}/>
        </p>
      );
    }
    i++;
  }

  return <div style={{padding:"2px 0"}}>{elements}</div>;
}

function InlineMarkdown({ text }) {
  // Handle **bold**, *italic*, `code`
  const parts = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) parts.push(<span key={last}>{text.slice(last, m.index)}</span>);
    const raw = m[0];
    if (raw.startsWith("**")) {
      parts.push(<strong key={m.index} style={{color:"#fff",fontWeight:700}}>{raw.slice(2,-2)}</strong>);
    } else if (raw.startsWith("*")) {
      parts.push(<em key={m.index} style={{color:"#ddd"}}>{raw.slice(1,-1)}</em>);
    } else {
      parts.push(<code key={m.index} style={{background:"rgba(255,255,255,0.08)",borderRadius:3,padding:"1px 5px",fontSize:11,fontFamily:"monospace"}}>{raw.slice(1,-1)}</code>);
    }
    last = m.index + raw.length;
  }
  if (last < text.length) parts.push(<span key={last}>{text.slice(last)}</span>);
  return <>{parts}</>;
}

// ── Utilities ────────────────────────────────────────────────────────────────
function norm(s){ return s?.toLowerCase().replace(/[\s_\-&()]/g,"") ?? ""; }
function detectCol(headers, aliases){
  for(const a of aliases){
    const i = headers.findIndex(h => norm(h)===norm(a));
    if(i !== -1) return headers[i];
  }
  return null;
}
function isConverted(v){
  if(v===null||v===undefined||v==="") return false;
  const s = String(v).toLowerCase().trim();
  return s==="true"||s==="1"||s==="yes"||s==="y"||s==="converted"||s==="opportunity"||s==="won";
}
function isDisqualified(v){
  if(!v) return false;
  const s = String(v).toLowerCase().trim();
  return s==="disqualified"||s==="dq";
}
function isStuck(v){
  if(!v) return false;
  const s = String(v).toLowerCase().trim();
  return STUCK_STATUSES.some(st => s.includes(st));
}
function computeConvByDim(rows, dimCol, convCol){
  const map = {};
  for(const r of rows){
    const d = r[dimCol] ? String(r[dimCol]).trim() : "(Unknown)";
    if(!map[d]) map[d] = {total:0, converted:0};
    map[d].total++;
    if(isConverted(r[convCol])) map[d].converted++;
  }
  return Object.entries(map)
    .map(([label,{total,converted}]) => ({label,total,converted,rate:total>0?converted/total:0}))
    .filter(d => d.total >= 2)
    .sort((a,b) => b.rate-a.rate);
}
function buildLeadProfile(row, colMap){
  return PROFILE_FIELDS
    .map(({key,label}) => {
      const col = colMap[key];
      return (col && row[col] && String(row[col]).trim()) ? label+": "+String(row[col]).trim() : null;
    })
    .filter(Boolean).join(" | ");
}
async function callClaude(system, user, maxTokens){
  const res = await fetch("/api/claude",{
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:maxTokens||2400,system,messages:[{role:"user",content:user}]})
  });
  const data = await res.json();
  if(data.error) throw new Error(data.error.message);
  return data.content?.map(b=>b.text||"").join("") ?? "";
}

// ── UI primitives ─────────────────────────────────────────────────────────────
function Spinner(){
  return <div style={{width:16,height:16,border:"2px solid rgba(255,255,255,0.15)",borderTopColor:ACCENT,borderRadius:"50%",animation:"spin 0.7s linear infinite",flexShrink:0}}/>;
}
function Tag({children,bg,color}){
  return <span style={{background:bg||"rgba(255,255,255,0.06)",color:color||"#888",borderRadius:99,padding:"3px 10px",fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>{children}</span>;
}
function SectionLabel({text,accent}){
  return (
    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:20}}>
      <div style={{width:3,height:18,background:accent||ACCENT,borderRadius:2,flexShrink:0}}/>
      <span style={{fontFamily:"monospace",fontSize:10,letterSpacing:3,color:"#777",textTransform:"uppercase"}}>{text}</span>
    </div>
  );
}
function Bar({pct,color}){
  return (
    <div style={{background:"rgba(255,255,255,0.06)",borderRadius:3,height:5,overflow:"hidden",flex:1}}>
      <div style={{width:Math.min(pct,100)+"%",background:color,height:"100%",borderRadius:3,transition:"width 0.9s cubic-bezier(0.16,1,0.3,1)",boxShadow:"0 0 5px "+color+"99"}}/>
    </div>
  );
}
function DimCard({title,data,color}){
  const top = data.slice(0,8);
  const max = top[0]?.rate||1;
  return (
    <div style={{background:"rgba(255,255,255,0.025)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:14,padding:"20px 22px"}}>
      <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:16}}>
        <div style={{width:7,height:7,borderRadius:"50%",background:color,boxShadow:"0 0 7px "+color}}/>
        <span style={{fontFamily:"monospace",fontSize:10,letterSpacing:2,color:"#666",textTransform:"uppercase"}}>{title}</span>
      </div>
      {top.length===0 ? <p style={{color:"#444",fontSize:12}}>Not enough data</p> : (
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {top.map((d,i) => (
            <div key={d.label}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                <span style={{fontSize:12,color:i===0?"#fff":"#999",fontWeight:i===0?600:400,maxWidth:"64%",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.label}</span>
                <span style={{fontSize:11,color,fontFamily:"monospace",fontWeight:700}}>{(d.rate*100).toFixed(1)+"%"}</span>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <Bar pct={(d.rate/max)*100} color={color}/>
                <span style={{fontSize:10,color:"#555",minWidth:46,textAlign:"right",fontFamily:"monospace"}}>{d.converted+"/"+d.total}</span>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
function AIPanel({title,content,accent,icon,onExport,exporting}){
  return (
    <div style={{background:"rgba(255,255,255,0.02)",border:"1px solid "+accent+"30",borderRadius:14,padding:"22px 26px"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <div style={{display:"flex",alignItems:"center",gap:9}}>
          <span style={{fontSize:16}}>{icon}</span>
          <span style={{fontFamily:"monospace",fontSize:10,letterSpacing:3,color:accent,textTransform:"uppercase"}}>{title}</span>
        </div>
        {onExport && (
          <button onClick={onExport} disabled={exporting} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.12)",color:"#aaa",borderRadius:8,padding:"6px 14px",cursor:"pointer",fontSize:11,display:"flex",alignItems:"center",gap:6,opacity:exporting?0.5:1}}>
            {exporting ? <><Spinner/><span>Exporting...</span></> : <><span>⬇</span><span>Export to Word</span></>}
          </button>
        )}
      </div>
      <MarkdownBlock content={content} accent={accent}/>
    </div>
  );
}
function NavTab({label,active,onClick,badge}){
  return (
    <button onClick={onClick} style={{background:active?ACCENT:"transparent",color:active?"#0c0c14":"#666",border:"1px solid "+(active?ACCENT:"rgba(255,255,255,0.09)"),borderRadius:99,padding:"7px 16px",cursor:"pointer",fontSize:11,fontWeight:active?700:500,display:"flex",alignItems:"center",gap:5,transition:"all 0.15s",fontFamily:"monospace"}}>
      {label}
      {badge!=null && <span style={{background:active?"rgba(0,0,0,0.2)":"rgba(255,255,255,0.08)",borderRadius:99,padding:"1px 6px",fontSize:10}}>{badge}</span>}
    </button>
  );
}
function FieldToggle({fields,selected,onToggle,activeColor,labelMap}){
  const labels = labelMap||NOTE_LABELS;
  return (
    <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
      {fields.map(f => {
        const on = selected.includes(f);
        return (
          <button key={f} onClick={()=>onToggle(f)} style={{background:on?activeColor+"18":"rgba(255,255,255,0.03)",border:"1px solid "+(on?activeColor:"rgba(255,255,255,0.09)"),color:on?activeColor:"#777",borderRadius:99,padding:"5px 13px",cursor:"pointer",fontSize:11,fontWeight:500,transition:"all 0.15s"}}>
            {labels[f]||f.replace(/_/g," ")}
          </button>
        );
      })}
    </div>
  );
}
function ProfileTags({fields,accent}){
  return (
    <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
      {fields.map(f => <Tag key={f.key} bg={accent+"15"} color={accent}>{f.label}</Tag>)}
    </div>
  );
}
function EmptyState({icon,title,subtitle}){
  return (
    <div style={{textAlign:"center",color:"#444",padding:"80px 0",border:"1px dashed rgba(255,255,255,0.06)",borderRadius:14}}>
      <div style={{fontSize:34,marginBottom:10}}>{icon}</div>
      <div style={{fontSize:15,marginBottom:8,color:"#555"}}>{title}</div>
      {subtitle && <div style={{fontSize:12,color:"#333"}}>{subtitle}</div>}
    </div>
  );
}
function RunButton({onClick,disabled,loading,color,label,loadingLabel}){
  return (
    <button onClick={onClick} disabled={disabled||loading} style={{background:color,color:color===ACCENT?"#0c0c14":"#fff",border:"none",borderRadius:10,padding:"10px 22px",cursor:"pointer",fontSize:12,fontWeight:700,display:"flex",alignItems:"center",gap:8,marginBottom:24,opacity:disabled||loading?0.45:1}}>
      {loading ? <><Spinner/><span>{loadingLabel||"Analyzing..."}</span></> : <span>{label}</span>}
    </button>
  );
}



// ── Main App ──────────────────────────────────────────────────────────────────
export default function App(){
  const [rows,setRows]               = useState(null);
  const [headers,setHeaders]         = useState([]);
  const [fileName,setFileName]       = useState("");
  const [error,setError]             = useState("");
  const [dragging,setDragging]       = useState(false);
  const [manualConvCol,setManualConvCol] = useState("");
  const [colMap,setColMap]           = useState({});
  const [activeTab,setActiveTab]     = useState("overview");
  const [notesInsight,setNotesInsight]       = useState(null);
  const [notesLoading,setNotesLoading]       = useState(false);
  const [convInsight,setConvInsight]         = useState(null);
  const [convLoading,setConvLoading]         = useState(false);
  const [disqInsight,setDisqInsight]         = useState(null);
  const [disqLoading,setDisqLoading]         = useState(false);
  const [stuckInsight,setStuckInsight]       = useState(null);
  const [stuckLoading,setStuckLoading]       = useState(false);
  const [linkedinInsight,setLinkedinInsight] = useState(null);
  const [linkedinLoading,setLinkedinLoading] = useState(false);
  const [selectedNoteFields,setSelectedNoteFields]         = useState([]);
  const [selectedStuckFields,setSelectedStuckFields]       = useState([]);
  const [selectedLinkedInFields,setSelectedLinkedInFields] = useState([]);
  const [selectedAE,setSelectedAE]           = useState(null);
  const [aeInsights,setAeInsights]           = useState({});
  const [aeLoadingFor,setAeLoadingFor]       = useState(null);
  const [aeTeamInsight,setAeTeamInsight]     = useState(null);
  const [aeTeamLoading,setAeTeamLoading]     = useState(false);
  const [aeNoteFields,setAeNoteFields]       = useState([]);
  const [negInsight,setNegInsight]           = useState(null);
  const [negLoading,setNegLoading]           = useState(false);
  const [negNoteFields,setNegNoteFields]     = useState([]);
  const [negScoreThreshold,setNegScoreThreshold] = useState(3);
  const [negLinkedInFields,setNegLinkedInFields] = useState([]);
  const [negPersonaInsight,setNegPersonaInsight] = useState(null);
  const [negPersonaLoading,setNegPersonaLoading] = useState(false);
  const [reengageInsight,setReengageInsight]     = useState(null);
  const [reengageLoading,setReengageLoading]     = useState(false);
  const [reengageNoteFields,setReengageNoteFields] = useState([]);
  const [exportError,setExportError]         = useState("");
  const [companyContext,setCompanyContext]     = useState("We are PingCAP, the company behind TiDB — an open-source, MySQL-compatible, distributed SQL database designed for HTAP (Hybrid Transactional and Analytical Processing) workloads. TiDB is built for enterprises that need horizontal scalability, high availability, and strong consistency without sharding complexity. Key differentiators: MySQL compatibility (no app rewrite needed), auto-sharding, real-time analytics alongside transactions, multi-cloud/on-prem deployment, and open source. Primary competitors: CockroachDB, Aurora, Vitess/MySQL, YugabyteDB, SingleStore. ICP: mid-to-large engineering-led companies with high-volume transactional workloads outgrowing MySQL or needing global scale.");
  const [showContext,setShowContext]           = useState(false);

  const parseFile = useCallback((file) => {
    setError(""); setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb   = XLSX.read(e.target.result,{type:"array"});
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws,{defval:""});
        if(!data.length){ setError("File appears empty."); return; }
        const hdrs = Object.keys(data[0]);
        setHeaders(hdrs); setRows(data);
        const detected = {};
        for(const [key,aliases] of Object.entries(FIELD_ALIASES)){
          const col = detectCol(hdrs,aliases);
          if(col) detected[key] = col;
        }
        setColMap(detected);
        const noteDetected = NOTE_FIELDS.filter(f => detected[f]);
        setSelectedNoteFields(noteDetected);
        setSelectedStuckFields(NOTE_FIELDS.filter(f => detected[f]));
        setSelectedLinkedInFields(LINKEDIN_FIELDS.filter(f => detected[f]));
        setAeNoteFields(noteDetected);
        setNegNoteFields(noteDetected);
        setNegLinkedInFields(LINKEDIN_FIELDS.filter(f => detected[f]));
        setReengageNoteFields(noteDetected);
        setNegInsight(null); setNegPersonaInsight(null); setReengageInsight(null);
        setNotesInsight(null); setConvInsight(null); setDisqInsight(null);
        setStuckInsight(null); setLinkedinInsight(null);
        setAeInsights({}); setAeTeamInsight(null); setSelectedAE(null);
    setNegInsight(null); setNegPersonaInsight(null); setReengageInsight(null);
        setActiveTab("overview");
      } catch(e){ setError("Could not parse file. Please upload a valid CSV or Excel file."); }
    };
    reader.readAsArrayBuffer(file);
  },[]);

  const onDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false);
    const f = e.dataTransfer.files[0];
    if(f) parseFile(f);
  },[parseFile]);

  const activeConvCol = colMap.converted||manualConvCol;

  const analysis = useMemo(() => {
    if(!rows||!activeConvCol) return null;
    const statusCol = colMap.lead_status;
    const dimData = DIM_KEYS.map((key,i) => {
      const col = colMap[key]; if(!col) return null;
      return {key,label:key.replace(/_/g," "),col,color:DIM_COLORS[i%DIM_COLORS.length],data:computeConvByDim(rows,col,activeConvCol)};
    }).filter(Boolean);
    const totalMeetings  = rows.length;
    const totalConverted = rows.filter(r => isConverted(r[activeConvCol])).length;
    const overallRate    = totalMeetings>0 ? totalConverted/totalMeetings : 0;
    const allSignals     = dimData.flatMap(d => d.data.slice(0,3).map(item => ({...item,dim:d.label,color:d.color}))).sort((a,b) => b.rate-a.rate).slice(0,8);
    const totalDisq  = statusCol ? rows.filter(r => isDisqualified(r[statusCol])).length : 0;
    const totalStuck = statusCol ? rows.filter(r => !isConverted(r[activeConvCol])&&!isDisqualified(r[statusCol])&&isStuck(r[statusCol])).length : 0;
    return {totalMeetings,totalConverted,overallRate,dimData,allSignals,totalDisq,totalStuck,statusCol};
  },[rows,activeConvCol,colMap]);

  const aeData = useMemo(() => {
    if(!rows||!activeConvCol||!colMap.ae_owner) return null;
    const aeCol = colMap.ae_owner;
    const statusCol = colMap.lead_status;
    const map = {};
    for(const r of rows){
      const ae = String(r[aeCol]||"").trim();
      if(!ae) continue;
      if(!map[ae]) map[ae] = {total:0,converted:0,disq:0,stuck:0,rows:[]};
      map[ae].total++;
      map[ae].rows.push(r);
      if(isConverted(r[activeConvCol])) map[ae].converted++;
      else if(statusCol && isDisqualified(r[statusCol])) map[ae].disq++;
      else if(statusCol && isStuck(r[statusCol])) map[ae].stuck++;
    }
    return Object.entries(map)
      .map(([name,d]) => ({name,...d,rate:d.total>0?d.converted/d.total:0}))
      .filter(ae => ae.total >= 2)
      .sort((a,b) => b.rate-a.rate);
  },[rows,activeConvCol,colMap]);

  const teamAvgRate = aeData&&aeData.length>0 ? aeData.reduce((s,a)=>s+a.rate,0)/aeData.length : 0;

  const detectedNoteFields     = NOTE_FIELDS.filter(f => colMap[f]);
  const detectedLinkedInFields = LINKEDIN_FIELDS.filter(f => colMap[f]);
  const detectedProfileFields  = PROFILE_FIELDS.filter(f => colMap[f.key]);
  const convertedCount = rows ? rows.filter(r => isConverted(r[activeConvCol])).length : 0;
  const disqCount  = analysis?.totalDisq||0;
  const stuckCount = analysis?.totalStuck||0;
  // Low-score leads: rows where iqm_score or ae_score is a number <= threshold
  const negCount = useMemo(() => {
    if(!rows||!activeConvCol) return 0;
    const iqmScoreCol = colMap.iqm_score;
    const aeScoreCol  = colMap.ae_score;
    const statusCol   = colMap.lead_status;
    let n = 0;
    for(const r of rows){
      if(isConverted(r[activeConvCol])) continue;
      if(statusCol && isDisqualified(r[statusCol])){ n++; continue; }
      const iqmV = iqmScoreCol ? parseFloat(r[iqmScoreCol]) : NaN;
      const aeV  = aeScoreCol  ? parseFloat(r[aeScoreCol])  : NaN;
      if((!isNaN(iqmV) && iqmV <= negScoreThreshold) || (!isNaN(aeV) && aeV <= negScoreThreshold)) n++;
    }
    return n;
  },[rows,activeConvCol,colMap,negScoreThreshold]);

  function buildNotesBlock(subset, fields){
    const parts = [];
    for(const f of fields){
      const col = colMap[f]; if(!col) continue;
      const texts = subset.map(r => r[col]).filter(v => v&&String(v).trim().length>5).slice(0,80).map(v => String(v).trim());
      if(texts.length) parts.push("=== "+(NOTE_LABELS[f]||STUCK_LABELS[f]||f)+" ===\n"+texts.join("\n"));
    }
    const profiles = subset.map(r => buildLeadProfile(r,colMap)).filter(p => p.length>5).slice(0,120);
    if(profiles.length) parts.push("=== LEAD PROFILES ===\n"+profiles.join("\n"));
    return parts.join("\n\n");
  }

  const MD_SYSTEM = "Respond using Markdown formatting. Use ## for section headers, ### for sub-headers, **bold** for key terms, and - for bullet lists. Be thorough and specific. Do not truncate any section.";
  const CTX = companyContext.trim()
    ? " COMPANY CONTEXT: "+companyContext.trim()+" — All analysis, positioning, competitive references, ICP definitions, and recommendations must be specific to this company and product. Never reference competitors as aspirational targets — treat them as displacement opportunities."
    : "";

  const runNotesAnalysis = async () => {
    setNotesLoading(true); setNotesInsight(null);
    try {
      const subset = rows.filter(r => isConverted(r[activeConvCol]));
      const text   = buildNotesBlock(subset,selectedNoteFields);
      if(!text.trim()){ setNotesInsight("No content found for converted leads."); setNotesLoading(false); return; }
      const result = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B sales intelligence analyst. Analyze sales notes AND lead profile attributes to extract patterns about who converts and why. Be specific and complete every section fully.",
        "Data from CONVERTED leads:\n\n"+text.slice(0,16000)+"\n\nProvide these sections:\n## 1. Top Conversion Signals\n## 2. Ideal Lead Profile\n## 3. Common Pain Points\n## 4. Budget Patterns\n## 5. Authority Signals\n## 6. Timing Triggers\n## 7. Technology Context\n## 8. What to Listen For on Future Calls",
        3000
      );
      setNotesInsight(result);
    } catch(e){ setNotesInsight("Analysis failed: "+(e.message||"unknown error")); }
    setNotesLoading(false);
  };

  const runConvAnalysis = async () => {
    setConvLoading(true); setConvInsight(null);
    try {
      const subset   = rows.filter(r => isConverted(r[activeConvCol]));
      const profiles = subset.map(r => buildLeadProfile(r,colMap)).filter(p => p.length>5).slice(0,150);
      if(!profiles.length){ setConvInsight("No profile data found."); setConvLoading(false); return; }
      const result = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B demand generation strategist. Analyze lead profile data to define a precise Ideal Customer Profile. Be specific and complete every section fully.",
        "Profile data for "+subset.length+" CONVERTED leads:\n\n"+profiles.join("\n").slice(0,16000)+"\n\nProvide:\n## 1. Top Converting Lead Archetypes\n## 2. Job Title and Role Patterns\n## 3. Industry and Vertical Breakdown\n## 4. Technology Stack Patterns\n## 5. TCP Patterns\n## 6. Segment and Region Patterns\n## 7. Ideal Customer Profile Definition\n## 8. Targeting Criteria for CRM, LinkedIn, and ad platforms",
        3000
      );
      setConvInsight(result);
    } catch(e){ setConvInsight("Analysis failed: "+(e.message||"unknown error")); }
    setConvLoading(false);
  };

  const runDisqAnalysis = async () => {
    setDisqLoading(true); setDisqInsight(null);
    try {
      const statusCol = colMap.lead_status;
      if(!statusCol){ setDisqInsight("No Lead Status column detected."); setDisqLoading(false); return; }
      const subset = rows.filter(r => isDisqualified(r[statusCol]));
      const text   = buildNotesBlock(subset,selectedNoteFields);
      if(!text.trim()){ setDisqInsight("No content found for disqualified leads."); setDisqLoading(false); return; }
      const result = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B sales intelligence analyst. Analyze notes and profile attributes for disqualified leads to find bad-fit patterns. Be specific and complete every section fully.",
        "Data from DISQUALIFIED leads:\n\n"+text.slice(0,16000)+"\n\nProvide:\n## 1. Top Disqualification Reasons\n## 2. Bad-Fit Lead Profiles\n## 3. Red Flag Signals in Notes\n## 4. Technology Red Flags\n## 5. Budget and Authority Issues\n## 6. Bad-Fit Company Profiles\n## 7. Early Screening Questions\n## 8. ICP Exclusion Criteria",
        3000
      );
      setDisqInsight(result);
    } catch(e){ setDisqInsight("Analysis failed: "+(e.message||"unknown error")); }
    setDisqLoading(false);
  };

  const runStuckAnalysis = async () => {
    setStuckLoading(true); setStuckInsight(null);
    try {
      const statusCol   = colMap.lead_status;
      const stuckSubset = statusCol
        ? rows.filter(r => !isConverted(r[activeConvCol])&&!isDisqualified(r[statusCol])&&isStuck(r[statusCol]))
        : rows.filter(r => !isConverted(r[activeConvCol]));
      const convSubset  = rows.filter(r => isConverted(r[activeConvCol]));
      const stuckText   = buildNotesBlock(stuckSubset,selectedStuckFields);
      const convText    = buildNotesBlock(convSubset,selectedStuckFields);
      if(!stuckText.trim()){ setStuckInsight("No content found for stuck leads."); setStuckLoading(false); return; }
      const combined = "--- STUCK LEADS ("+stuckSubset.length+") ---\n\n"+stuckText.slice(0,10000)+"\n\n--- CONVERTED LEADS ("+convSubset.length+") FOR COMPARISON ---\n\n"+convText.slice(0,6000);
      const result = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B sales intelligence analyst comparing stuck pipeline leads against converted leads. Identify exactly what qualification criteria were missing. Be specific and complete every section fully.",
        combined+"\n\nProvide:\n## 1. Missing Qualification (BANT criteria absent or unclear)\n## 2. Next Steps Gaps\n## 3. Primary Need and Pain Gaps\n## 4. Authority Gaps\n## 5. Budget Gaps\n## 6. Profile Differences\n## 7. Common Themes in Stuck Notes\n## 8. What Converted Leads Had That Stuck Leads Did Not\n## 9. Recommended Actions",
        3000
      );
      setStuckInsight(result);
    } catch(e){ setStuckInsight("Analysis failed: "+(e.message||"unknown error")); }
    setStuckLoading(false);
  };

  const runLinkedInAnalysis = async () => {
    setLinkedinLoading(true); setLinkedinInsight(null);
    try {
      const subset = rows.filter(r => isConverted(r[activeConvCol]));
      const parts  = [];
      for(const f of selectedLinkedInFields){
        const col = colMap[f]; if(!col) continue;
        // Cap each field to 50 entries, truncate each entry to 300 chars to keep input tight
        const texts = subset.map(r => r[col]).filter(v => v&&String(v).trim().length>5)
          .slice(0,50).map(v => String(v).trim().slice(0,300));
        if(texts.length) parts.push("=== "+f.replace(/_/g," ").toUpperCase()+"\n"+texts.join("\n"));
      }
      const profiles = subset.map(r => buildLeadProfile(r,colMap)).filter(p => p.length>5).slice(0,60);
      if(profiles.length) parts.push("=== LEAD PROFILES\n"+profiles.join("\n"));
      const text = parts.join("\n\n");
      if(!text.trim()){ setLinkedinInsight("No LinkedIn data found."); setLinkedinLoading(false); return; }
      // Hard-cap input at 12k chars so output has room to breathe within token limits
      const dataBlock = "LinkedIn + profile data from "+subset.length+" CONVERTED leads:\n\n"+text.slice(0,12000);

      const CONCISE = " Be concise and structured. Each bullet point must be one line only. No padding or repetition.";

      // Call 1: Profile & skills sections — capped output
      const part1 = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B demand generation strategist."+CONCISE,
        dataBlock+"\n\nAnalyze the data above and provide ONLY these 4 sections. Each section: max 8 bullet points, one line each.\n\n## 1. Ideal LinkedIn Profile Characteristics\nTop 8 traits (seniority, function, background, education).\n\n## 2. Career Trajectory Patterns\nTop 6 career path patterns seen in the data.\n\n## 3. Top Skills and Keywords\nList as grouped bullets: Database Skills | Cloud Skills | Dev/Infra Skills | Certifications. Max 6 per group.\n\n## 4. Technology Background Patterns\nTop 8 tech stack or tooling patterns. One line each.",
        2000
      );

      // Call 2: Segments & activation — each segment strictly templated
      const part2 = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B demand generation strategist."+CONCISE,
        dataBlock+"\n\nAnalyze the data above and provide ONLY these 4 sections.\n\n## 5. Company Characteristics\nTop 6 firmographic patterns (size, industry, growth stage, tech maturity). One line each.\n\n## 6. LinkedIn Audience Segments\nIdentify the top 4 audience segments. For EACH use exactly this format:\n**Segment: [name]**\n- Titles: [list]\n- Industries: [list]\n- Size: [range]\n- Seniority: [level]\n- Targeting: [LinkedIn criteria]\n- Message: [one-line angle]\n\n## 7. Ad Targeting Recommendations\nTop 5 campaign recommendations. One line each covering objective, audience, and format.\n\n## 8. Messaging Angles\nFor each of the 4 segments above, write exactly 2 message angles. Format: **[Segment]** — [hook] / [pain point] / [CTA]",
        2500
      );

      setLinkedinInsight(part1+"\n\n---\n\n"+part2);
    } catch(e){ setLinkedinInsight("Analysis failed: "+(e.message||"unknown error")); }
    setLinkedinLoading(false);
  };

  const runAEAnalysis = async (ae) => {
    setAeLoadingFor(ae.name);
    try {
      const statusCol = colMap.lead_status;
      const convRows  = ae.rows.filter(r => isConverted(r[activeConvCol]));
      const stuckRows = ae.rows.filter(r => !isConverted(r[activeConvCol])&&(!statusCol||!isDisqualified(r[statusCol])));
      const convText  = buildNotesBlock(convRows,  aeNoteFields);
      const stuckText = buildNotesBlock(stuckRows, aeNoteFields);
      const summary   = "AE: "+ae.name+"\nTotal meetings: "+ae.total+"\nConverted: "+ae.converted+" ("+(ae.rate*100).toFixed(1)+"%)"+"\nStuck/Open: "+ae.stuck+"\nDisqualified: "+ae.disq+"\nTeam avg conversion rate: "+(teamAvgRate*100).toFixed(1)+"%";
      const body      = summary+"\n\n--- CONVERTED LEADS ("+convRows.length+") ---\n\n"+convText.slice(0,8000)+"\n\n--- STUCK/OPEN LEADS ("+stuckRows.length+") ---\n\n"+stuckText.slice(0,6000);
      const result    = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B sales coaching analyst reviewing individual AE performance on marketing-sourced meetings. Be specific, fair, and constructive. Complete every section fully.",
        body+"\n\nProvide:\n## 1. Performance Summary\n## 2. Qualification Strengths\n## 3. Qualification Gaps\n## 4. Note Quality\n## 5. Deal Progression Patterns\n## 6. Coaching Recommendations\n## 7. Leads to Revisit",
        3000
      );
      setAeInsights(prev => ({...prev,[ae.name]:result}));
    } catch(e){
      setAeInsights(prev => ({...prev,[ae.name]:"Analysis failed: "+(e.message||"unknown error")}));
    }
    setAeLoadingFor(null);
  };

  const runAETeamAnalysis = async () => {
    setAeTeamLoading(true); setAeTeamInsight(null);
    if(!aeData||aeData.length===0){ setAeTeamInsight("No AE data found."); setAeTeamLoading(false); return; }
    try {
      const leaderboard = aeData.map(ae =>
        ae.name+" | "+ae.total+" meetings | "+ae.converted+" converted | "+(ae.rate*100).toFixed(1)+"% | "+ae.stuck+" stuck | "+ae.disq+" disq"
      ).join("\n");
      const sorted     = [...aeData].sort((a,b) => b.rate-a.rate);
      const topHalf    = sorted.slice(0, Math.ceil(sorted.length/2));
      const bottomHalf = sorted.slice(Math.ceil(sorted.length/2));
      const topText    = buildNotesBlock(topHalf.flatMap(ae=>ae.rows.filter(r=>isConverted(r[activeConvCol]))).slice(0,40), aeNoteFields);
      const bottomText = buildNotesBlock(bottomHalf.flatMap(ae=>ae.rows.filter(r=>!isConverted(r[activeConvCol]))).slice(0,40), aeNoteFields);
      const body       = "TEAM LEADERBOARD:\n"+leaderboard+"\n\nTeam avg: "+(teamAvgRate*100).toFixed(1)+"%\n\n--- TOP PERFORMERS CONVERTED NOTES ---\n\n"+topText.slice(0,7000)+"\n\n--- LOWER PERFORMERS STUCK NOTES ---\n\n"+bottomText.slice(0,5000);
      const result     = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B sales coaching analyst reviewing a team of AEs on marketing-sourced meetings. Be fair, specific, and constructive. Complete every section fully.",
        body+"\n\nProvide:\n## 1. Leaderboard Summary\n## 2. Performance Spread\n## 3. What Top Performers Do Differently\n## 4. Common Team-Wide Qualification Gaps\n## 5. Systemic Issues\n## 6. Team Coaching Priorities\n## 7. Individual Callouts",
        3500
      );
      setAeTeamInsight(result);
    } catch(e){ setAeTeamInsight("Analysis failed: "+(e.message||"unknown error")); }
    setAeTeamLoading(false);
  };

  // Export handler -- downloads .md file (opens natively in Word, Notion, etc.)
  const handleExport = (exportKey, title, content) => {
    try {
      const blob = new Blob([content], {type:"text/markdown;charset=utf-8"});
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href     = url;
      a.download = title.replace(/[^a-z0-9]/gi,"_").toLowerCase()+".md";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch(e){
      setExportError("Export failed: "+e.message);
    }
  };

  const runNegAnalysis = async () => {
    setNegLoading(true); setNegInsight(null);
    try {
      const statusCol   = colMap.lead_status;
      const iqmScoreCol = colMap.iqm_score;
      const aeScoreCol  = colMap.ae_score;
      const disqSubset  = statusCol ? rows.filter(r => isDisqualified(r[statusCol])) : [];
      const lowScoreSubset = rows.filter(r => {
        if(isConverted(r[activeConvCol])) return false;
        if(statusCol && isDisqualified(r[statusCol])) return false;
        const iqmV = iqmScoreCol ? parseFloat(r[iqmScoreCol]) : NaN;
        const aeV  = aeScoreCol  ? parseFloat(r[aeScoreCol])  : NaN;
        return (!isNaN(iqmV) && iqmV <= negScoreThreshold) || (!isNaN(aeV) && aeV <= negScoreThreshold);
      });
      const convSubset = rows.filter(r => isConverted(r[activeConvCol]));
      const disqText   = buildNotesBlock(disqSubset,     negNoteFields);
      const lowText    = buildNotesBlock(lowScoreSubset, negNoteFields);
      const convText   = buildNotesBlock(convSubset.slice(0,40), negNoteFields);
      let scoreSummary = "";
      if(iqmScoreCol || aeScoreCol){
        const scoreLines = lowScoreSubset.slice(0,60).map(r => {
          const p = [buildLeadProfile(r, colMap)];
          if(iqmScoreCol) p.push("IQM Score: "+(r[iqmScoreCol]||"?"));
          if(aeScoreCol)  p.push("AE Score: "+(r[aeScoreCol]||"?"));
          return p.join(" | ");
        }).filter(Boolean);
        if(scoreLines.length) scoreSummary = "=== LOW SCORE LEAD PROFILES ===\n"+scoreLines.join("\n");
      }
      const hasDisq     = disqText.trim().length > 0;
      const hasLowScore = lowText.trim().length > 0 || scoreSummary.length > 0;
      if(!hasDisq && !hasLowScore){
        setNegInsight("No disqualified leads and no score columns detected.");
        setNegLoading(false); return;
      }
      const inputParts = [];
      if(hasDisq)     inputParts.push("--- DISQUALIFIED LEADS ("+disqSubset.length+") ---\n\n"+disqText.slice(0,6000));
      if(hasLowScore) inputParts.push("--- LOW SCORE LEADS (score <= "+negScoreThreshold+") ("+lowScoreSubset.length+") ---\n\n"+(scoreSummary+"\n\n"+lowText).slice(0,6000));
      inputParts.push("--- CONVERTED LEADS ("+convSubset.length+") FOR CONTRAST ---\n\n"+convText.slice(0,3500));
      const inputBlock = inputParts.join("\n\n");
      const CONCISE = " Be concise: max 6 bullets per section, one line each. No padding.";
      const prompt = inputBlock+"\n\nProvide a focused bad-fit persona analysis. Max 6 bullets per section:\n## 1. Bad-Fit Persona Summary\nWho NOT to target — titles, roles, seniority, background in 4-6 bullets.\n## 2. Top Disqualification Reasons\nMost common reasons leads fail. Max 6 bullets.\n## 3. Bad-Fit Technology Profile\nTech stacks, DBs, or cloud patterns that correlate with poor fit. Max 6 bullets.\n## 4. Red Flag Signals in Notes\nEarly warning phrases or intent signals. Max 6 bullets.\n## 5. Key Differences vs Converted Leads\nWhat converted leads had that bad-fit leads clearly lacked. Max 6 bullets.\n## 6. Exclusion Criteria for Targeting\nSpecific negative criteria for CRM, LinkedIn, paid campaigns. Max 6 bullets.\n## 7. Early Disqualification Questions\nBest 5 questions to surface bad fit quickly.";
      const result = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B sales intelligence analyst building a Negative ICP."+CONCISE,
        prompt, 2500
      );
      setNegInsight(result);
    } catch(e){ setNegInsight("Analysis failed: "+(e.message||"unknown error")); }
    setNegLoading(false);
  };

  const runNegPersonaAnalysis = async () => {
    setNegPersonaLoading(true); setNegPersonaInsight(null);
    try {
      const statusCol  = colMap.lead_status;
      // Bad-fit = disqualified + low-score non-converted
      const iqmScoreCol = colMap.iqm_score;
      const aeScoreCol  = colMap.ae_score;
      const badSubset = rows.filter(r => {
        if(isConverted(r[activeConvCol])) return false;
        if(statusCol && isDisqualified(r[statusCol])) return true;
        const iqmV = iqmScoreCol ? parseFloat(r[iqmScoreCol]) : NaN;
        const aeV  = aeScoreCol  ? parseFloat(r[aeScoreCol])  : NaN;
        return (!isNaN(iqmV) && iqmV <= negScoreThreshold) || (!isNaN(aeV) && aeV <= negScoreThreshold);
      });
      const convSubset = rows.filter(r => isConverted(r[activeConvCol]));

      const buildLinkedInBlock = (subset, label) => {
        const parts = [];
        for(const f of negLinkedInFields){
          const col = colMap[f]; if(!col) continue;
          const texts = subset.map(r => r[col]).filter(v => v&&String(v).trim().length>5)
            .slice(0,40).map(v => String(v).trim().slice(0,250));
          if(texts.length) parts.push("=== "+label+" — "+f.replace(/_/g," ").toUpperCase()+" ===\n"+texts.join("\n"));
        }
        // Also include profile attributes
        const profiles = subset.map(r => buildLeadProfile(r,colMap)).filter(p=>p.length>5).slice(0,40);
        if(profiles.length) parts.push("=== "+label+" — PROFILE ATTRIBUTES ===\n"+profiles.join("\n"));
        return parts.join("\n\n");
      };

      const badBlock  = buildLinkedInBlock(badSubset,  "BAD FIT");
      const convBlock = buildLinkedInBlock(convSubset, "CONVERTED");

      if(!badBlock.trim()){ setNegPersonaInsight("No LinkedIn data found for bad-fit leads. Make sure LinkedIn columns are selected."); setNegPersonaLoading(false); return; }

      const inputBlock = badBlock.slice(0,8000)+"\n\n"+convBlock.slice(0,6000);
      const prompt = inputBlock+"\n\nYou have two groups: BAD FIT leads (did not convert / low score) and CONVERTED leads. Analyze both and produce an EXCLUSION-FOCUSED persona report. For every section, state what to AVOID targeting. Max 6 bullets per section, one line each.\n\n## 1. Titles and Functions to Exclude\nSpecific job titles and functions that appear in bad-fit leads but not in converted. List as exclusion terms.\n## 2. Seniority Levels to Exclude\nWhich seniority bands skew toward bad fit? State the levels to avoid explicitly.\n## 3. Skills and Keywords to Exclude\nSkills, tools, certifications, or keyword patterns on bad-fit profiles. These are LinkedIn negative keywords — list as exclusion terms.\n## 4. Career Backgrounds to Exclude\nPrior roles or career path types that predict bad fit. What history is a red flag?\n## 5. Industries and Verticals to Exclude\nIndustries or verticals where personas consistently do not convert.\n## 6. Exclusion Persona Summary\nWrite a 3-4 sentence description of the person NOT to target, specific enough to use as LinkedIn campaign exclusion criteria.";
      const result = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B persona analyst. Your job is to identify how the PEOPLE who convert differ from those who don't — focus on persona, career, seniority, and background differences, not product fit. Be specific and concise.",
        prompt, 2500
      );
      setNegPersonaInsight(result);
    } catch(e){ setNegPersonaInsight("Analysis failed: "+(e.message||"unknown error")); }
    setNegPersonaLoading(false);
  };

  const runReengageAnalysis = async () => {
    setReengageLoading(true); setReengageInsight(null);
    try {
      const statusCol    = colMap.lead_status;
      const accountCol   = colMap.account_name;
      // Candidates: non-converted, non-disqualified leads with notes suggesting interest or fit
      const candidates = rows.filter(r => {
        if(isConverted(r[activeConvCol])) return false;
        if(statusCol && isDisqualified(r[statusCol])) return false;
        return true;
      });
      // Build per-account summaries if we have an account column, else just use lead profiles
      let accountBlock = "";
      if(accountCol){
        const accountMap = {};
        for(const r of candidates){
          const acct = String(r[accountCol]||"").trim(); if(!acct) continue;
          if(!accountMap[acct]) accountMap[acct] = [];
          accountMap[acct].push(r);
        }
        const accountLines = Object.entries(accountMap).slice(0,200).map(([acct, acctRows]) => {
          const profiles = acctRows.map(r => buildLeadProfile(r,colMap)).filter(Boolean).join(" | ");
          const notes    = reengageNoteFields.flatMap(f => {
            const col = colMap[f]; if(!col) return [];
            return acctRows.map(r => r[col]).filter(v => v&&String(v).trim().length>5).map(v => String(v).trim().slice(0,150));
          }).join(" // ");
          return "ACCOUNT: "+acct+" | Profiles: "+profiles+" | Notes: "+(notes||"(no notes)");
        });
        accountBlock = "=== NON-CONVERTED ACCOUNTS ("+Object.keys(accountMap).length+" accounts) ===\n\n"+accountLines.join("\n");
      } else {
        const leadLines = candidates.slice(0,80).map(r => {
          const profile = buildLeadProfile(r,colMap);
          const notes   = reengageNoteFields.flatMap(f => {
            const col = colMap[f]; if(!col) return [];
            const v = r[col]; return (v&&String(v).trim().length>5) ? [String(v).trim().slice(0,200)] : [];
          }).join(" // ");
          return "LEAD: "+profile+"\n  Notes: "+(notes||"(no notes)");
        }).filter(Boolean);
        accountBlock = "=== NON-CONVERTED LEADS ("+candidates.length+") ===\n\n"+leadLines.join("\n\n");
      }

      if(!accountBlock.trim()){ setReengageInsight("No non-converted leads found."); setReengageLoading(false); return; }

      const convProfiles = rows.filter(r=>isConverted(r[activeConvCol])).slice(0,30).map(r=>buildLeadProfile(r,colMap)).filter(Boolean).join("\n");
      const inputBlock = accountBlock.slice(0,25000)+"\n\n=== CONVERTED LEAD PROFILES FOR REFERENCE ===\n"+convProfiles.slice(0,2000);

      const prompt = inputBlock+"\n\nAnalyze the non-converted "+(accountCol?"accounts":"leads")+" above. Identify which ones show signals suggesting the ACCOUNT is still a good target for re-engagement (SDR outreach or targeted ad campaigns), even though the specific lead we met didn't convert. Look for: genuine interest or pain expressed in notes, right company profile but wrong contact, timing issues that may resolve, budget interest with authority gap, technical fit signals.\n\nProvide:\n## 1. Re-engagement Signal Patterns\nWhat note patterns or profile signals indicate a good account to re-target. Max 6 bullets.\n## 2. "
      +(accountCol ? "Full List of Accounts to Re-engage\nList EVERY account that shows any re-engagement potential — not just the top ones. For each account include: account name, re-engagement signals observed, suggested approach (ads vs SDR outreach), and what to lead with. Group them by priority tier (High / Medium / Low) but include all of them.\n## 3. " : "")
      +"Persona Gap Hypothesis\nFor accounts worth re-targeting, what type of person should we reach next instead — different title, seniority, or function?\n## "+(accountCol?"4":"3")+". Recommended Re-engagement Message Themes\nTop 4 message angles to use in ads or SDR outreach for these accounts.\n## "+(accountCol?"5":"4")+". Accounts to Deprioritize\nWhich accounts show clear signals of no fit and should be removed from targeting entirely.";

      const result = await callClaude(
        MD_SYSTEM+CTX+" You are a B2B revenue intelligence analyst identifying account re-engagement opportunities. Focus on accounts where meeting notes or profile data suggest company-level fit even if the specific contact wasn't the right buyer. Be specific and actionable.",
        prompt, 5000
      );
      setReengageInsight(result);
    } catch(e){ setReengageInsight("Analysis failed: "+(e.message||"unknown error")); }
    setReengageLoading(false);
  };


  const toggleField   = (setter) => (f) => setter(p => p.includes(f) ? p.filter(x=>x!==f) : [...p,f]);
  const toggleAeNote  = (f) => setAeNoteFields(p => p.includes(f) ? p.filter(x=>x!==f) : [...p,f]);
  const reset = () => {
    setRows(null); setHeaders([]); setColMap({}); setFileName(""); setError("");
    setNotesInsight(null); setConvInsight(null); setDisqInsight(null);
    setStuckInsight(null); setLinkedinInsight(null);
    setAeInsights({}); setAeTeamInsight(null); setSelectedAE(null);
    setNegInsight(null); setNegPersonaInsight(null); setReengageInsight(null);
  };

  return (
    <div style={{minHeight:"100vh",background:"#0c0c14",color:"#fff",padding:"36px 28px",fontFamily:"system-ui,sans-serif"}}>
      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes fadeup { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:translateY(0); } }
        .up { animation: fadeup 0.35s ease forwards; }
        button:focus { outline: none; }
      `}</style>
      <div style={{maxWidth:1160,margin:"0 auto"}}>

        <div style={{marginBottom:40,display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:16}}>
          <div>
            <div style={{fontFamily:"monospace",fontSize:10,letterSpacing:3,color:ACCENT,marginBottom:10}}>SALES INTELLIGENCE</div>
            <h1 style={{fontSize:32,fontWeight:800,margin:0,letterSpacing:"-1.5px",lineHeight:1.1}}>Marketing Meeting<br/><span style={{color:ACCENT}}>Intelligence</span></h1>
          </div>
          {rows && <button onClick={reset} style={{background:"transparent",border:"1px solid rgba(255,255,255,0.1)",color:"#666",borderRadius:8,padding:"7px 14px",cursor:"pointer",fontSize:11}}>New File</button>}
        </div>

        {!rows && (
          <div>
            <div onDragOver={e=>{e.preventDefault();setDragging(true);}} onDragLeave={()=>setDragging(false)} onDrop={onDrop} onClick={()=>document.getElementById("mi_file").click()}
              style={{border:"2px dashed "+(dragging?ACCENT:"rgba(255,255,255,0.09)"),borderRadius:20,padding:"68px 36px",textAlign:"center",background:dragging?ACCENT+"07":"rgba(255,255,255,0.01)",transition:"all 0.2s",cursor:"pointer"}}>
              <div style={{fontSize:42,marginBottom:14}}>📊</div>
              <div style={{fontSize:17,fontWeight:600,marginBottom:8}}>Drop your CSV or Excel file</div>
              <div style={{fontSize:12,color:"#4a4a5a",maxWidth:640,margin:"0 auto",lineHeight:2}}>
                Job Title | Job Role | Industry | Vertical | Segment | Region | Lead Source<br/>
                TCP Regional | MySQL Family Capture | Cloud Vendor | Existing Database<br/>
                IQM Notes | SDR Notes | Budget | Authority | Need/Pain | Timing | Team | AE Feedback | Next Steps<br/>
                AE (owner) | Converted | Lead Status | LinkedIn data
              </div>
              <input id="mi_file" type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={e=>{if(e.target.files[0])parseFile(e.target.files[0]);}}/>
              <div style={{marginTop:26,display:"inline-block",background:ACCENT,color:"#0c0c14",fontWeight:700,padding:"10px 26px",borderRadius:99,fontSize:13}}>Browse File</div>
            </div>
            {error && <div style={{color:"#FF6B6B",marginTop:14,fontSize:12}}>{error}</div>}
          </div>
        )}

        {rows && !colMap.converted && (
          <div style={{background:"rgba(255,77,109,0.07)",border:"1px solid rgba(255,77,109,0.22)",borderRadius:12,padding:"16px 20px",marginBottom:22}}>
            <p style={{margin:"0 0 10px",color:ACCENT2,fontWeight:600,fontSize:12}}>Select the Converted / Opportunity column:</p>
            <select value={manualConvCol} onChange={e=>setManualConvCol(e.target.value)} style={{background:"#181826",color:"#fff",border:"1px solid rgba(255,255,255,0.12)",borderRadius:8,padding:"7px 13px",fontSize:12}}>
              <option value="">-- Select --</option>
              {headers.map(h=><option key={h} value={h}>{h}</option>)}
            </select>
          </div>
        )}

        {/* COMPANY CONTEXT PANEL */}
        <div style={{marginBottom:24}}>
          <button onClick={()=>setShowContext(p=>!p)} style={{background:"rgba(255,255,255,0.04)",border:"1px solid rgba(255,255,255,0.09)",color:"#888",borderRadius:8,padding:"6px 14px",cursor:"pointer",fontSize:11,display:"flex",alignItems:"center",gap:6}}>
            <span style={{color:companyContext.trim()?ACCENT3:"#555",fontSize:9}}>{"●"}</span>
            <span>{"Company Context"}</span>
            <span style={{opacity:0.4}}>{showContext ? "▲" : "▼"}</span>
          </button>
          {showContext && (
            <div style={{marginTop:10,background:"rgba(255,255,255,0.025)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:12,padding:"16px 18px"}}>
              <div style={{fontSize:11,color:"#666",marginBottom:8,letterSpacing:1,textTransform:"uppercase"}}>Tell the AI who you are — this shapes every analysis</div>
              <textarea
                value={companyContext}
                onChange={e=>setCompanyContext(e.target.value)}
                rows={5}
                placeholder={"e.g. We sell an open-source distributed SQL database called TiDB. Key differentiators: MySQL compatibility, horizontal scale, HTAP. Competitors: CockroachDB, Aurora, YugabyteDB..."}
                style={{width:"100%",background:"rgba(0,0,0,0.3)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:8,padding:"10px 12px",color:"#ddd",fontSize:12,lineHeight:1.6,resize:"vertical",boxSizing:"border-box",fontFamily:"system-ui,sans-serif"}}
              />
              <div style={{fontSize:11,color:"#444",marginTop:6}}>Include: company name, product, key differentiators, top competitors, and ICP summary. The more specific, the better the analysis.</div>
            </div>
          )}
        </div>

        {rows && (
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:28,flexWrap:"wrap"}}>
            <Tag bg={ACCENT3+"18"} color={ACCENT3}>{fileName}</Tag>
            <Tag>{rows.length.toLocaleString()+" rows"}</Tag>
            {activeConvCol && <Tag>{"conv: "+activeConvCol}</Tag>}
            {colMap.lead_status && <Tag>{"status: "+colMap.lead_status}</Tag>}
            {colMap.ae_owner && <Tag bg={ACCENT5+"18"} color={ACCENT5}>{"ae: "+colMap.ae_owner}</Tag>}
            {detectedNoteFields.length>0 && <Tag bg={ACCENT+"18"} color={ACCENT}>{detectedNoteFields.length+" note cols"}</Tag>}
            {detectedProfileFields.length>0 && <Tag bg={ACCENT3+"18"} color={ACCENT3}>{detectedProfileFields.length+" profile cols"}</Tag>}
          </div>
        )}

        {analysis && (
          <div style={{display:"flex",gap:7,marginBottom:32,flexWrap:"wrap"}}>
            <NavTab label="Overview"            active={activeTab==="overview"} onClick={()=>setActiveTab("overview")}/>
            <NavTab label="Conversion Analysis" active={activeTab==="conv"}     onClick={()=>setActiveTab("conv")}     badge={convertedCount||undefined}/>
            <NavTab label="Notes Analysis"      active={activeTab==="notes"}    onClick={()=>setActiveTab("notes")}    badge={detectedNoteFields.length||undefined}/>
            <NavTab label="LinkedIn"            active={activeTab==="linkedin"} onClick={()=>setActiveTab("linkedin")} badge={detectedLinkedInFields.length||undefined}/>
            <NavTab label="Disqualified"        active={activeTab==="disq"}     onClick={()=>setActiveTab("disq")}     badge={disqCount||undefined}/>
            <NavTab label="Stuck in Pipeline"   active={activeTab==="stuck"}    onClick={()=>setActiveTab("stuck")}    badge={stuckCount||undefined}/>
            <NavTab label="AE Performance"      active={activeTab==="ae"}       onClick={()=>setActiveTab("ae")}       badge={aeData?aeData.length:undefined}/>
            <NavTab label="Who NOT to Target"  active={activeTab==="neg"}      onClick={()=>setActiveTab("neg")}      badge={negCount||undefined}/>
            <NavTab label="Re-engage Accounts" active={activeTab==="reengage"} onClick={()=>setActiveTab("reengage")}  badge={rows?rows.filter(r=>!isConverted(r[activeConvCol])&&!(colMap.lead_status&&isDisqualified(r[colMap.lead_status]))).length:undefined}/>
          </div>
        )}

        {exportError && (
          <div style={{background:"rgba(255,77,109,0.08)",border:"1px solid rgba(255,77,109,0.22)",borderRadius:8,padding:"10px 16px",marginBottom:16,fontSize:12,color:ACCENT2}}>
            {exportError}
          </div>
        )}

        {/* OVERVIEW */}
        {analysis && activeTab==="overview" && (
          <div className="up">
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:32}}>
              {[
                {l:"Total Meetings",v:analysis.totalMeetings.toLocaleString(),c:"#fff"},
                {l:"Converted",v:analysis.totalConverted.toLocaleString(),c:ACCENT3},
                {l:"Conv. Rate",v:(analysis.overallRate*100).toFixed(1)+"%",c:ACCENT},
                {l:"Stuck",v:analysis.totalStuck.toLocaleString(),c:ACCENT4},
                {l:"Disqualified",v:analysis.totalDisq.toLocaleString(),c:ACCENT2},
              ].map(s=>(
                <div key={s.l} style={{background:"rgba(255,255,255,0.03)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:12,padding:"16px 20px",minWidth:120}}>
                  <div style={{fontSize:28,fontWeight:800,fontFamily:"monospace",color:s.c}}>{s.v}</div>
                  <div style={{fontSize:10,color:"#555",textTransform:"uppercase",letterSpacing:"1.5px",marginTop:3}}>{s.l}</div>
                </div>
              ))}
            </div>
            {analysis.allSignals.length>0 && (
              <div style={{background:"linear-gradient(135deg,"+ACCENT+"09,"+ACCENT3+"09)",border:"1px solid "+ACCENT+"1a",borderRadius:14,padding:"20px 24px",marginBottom:28}}>
                <SectionLabel text="Top Conversion Signals" accent={ACCENT}/>
                <div style={{display:"flex",flexDirection:"column",gap:11}}>
                  {analysis.allSignals.map((s,i)=>(
                    <div key={i} style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                      <span style={{fontSize:10,color:s.color,fontFamily:"monospace",textTransform:"uppercase",minWidth:100}}>{s.dim}</span>
                      <span style={{fontWeight:600,fontSize:13,flex:1}}>{s.label}</span>
                      <span style={{fontFamily:"monospace",fontSize:12,color:s.color,fontWeight:700}}>{(s.rate*100).toFixed(1)+"%"}</span>
                      <span style={{fontSize:11,color:"#555"}}>{s.converted+"/"+s.total}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(290px,1fr))",gap:16}}>
              {analysis.dimData.filter(d=>d.data.length>0).map(d=>(
                <DimCard key={d.key} title={d.label} data={d.data} color={d.color}/>
              ))}
            </div>
          </div>
        )}

        {/* CONVERSION ANALYSIS */}
        {analysis && activeTab==="conv" && (
          <div className="up">
            <div style={{background:ACCENT3+"09",border:"1px solid "+ACCENT3+"22",borderRadius:14,padding:"20px 24px",marginBottom:20}}>
              <SectionLabel text="Deep Conversion Profile Analysis" accent={ACCENT3}/>
              <p style={{fontSize:13,color:"#aaa",margin:"0 0 16px",lineHeight:1.7}}>
                {"AI analysis of "}<span style={{color:ACCENT3,fontWeight:700}}>{convertedCount+" converted leads"}</span>{" across all profile attributes to define your precise ICP."}
              </p>
              {detectedProfileFields.length>0 ? <ProfileTags fields={detectedProfileFields} accent={ACCENT3}/> : <p style={{fontSize:12,color:"#555"}}>No profile columns detected.</p>}
            </div>
            <RunButton onClick={runConvAnalysis} disabled={detectedProfileFields.length===0} loading={convLoading} color={ACCENT3} label="Analyze Who Converts"/>
            {convInsight && <AIPanel title="Conversion Profile Intelligence" content={convInsight} accent={ACCENT3} icon="✦" onExport={()=>handleExport("conv","Conversion Profile Intelligence",convInsight)}/>}
          </div>
        )}

        {/* NOTES */}
        {analysis && activeTab==="notes" && (
          <div className="up">
            {detectedNoteFields.length===0 ? (
              <EmptyState icon="📝" title="No notes columns detected" subtitle="Expected: IQM Notes, SDR Notes, Budget, Authority, Need, Timing, Team, AE Feedback, Next Steps"/>
            ) : (
              <>
                <div style={{background:"rgba(255,255,255,0.02)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:14,padding:"20px 24px",marginBottom:20}}>
                  <SectionLabel text="Select Note Columns to Analyze" accent={ACCENT}/>
                  <FieldToggle fields={detectedNoteFields} selected={selectedNoteFields} onToggle={toggleField(setSelectedNoteFields)} activeColor={ACCENT}/>
                  {detectedProfileFields.length>0 && (
                    <div style={{marginTop:14,paddingTop:14,borderTop:"1px solid rgba(255,255,255,0.06)"}}>
                      <div style={{fontSize:11,color:"#666",marginBottom:8}}>Also enriching with profile attributes:</div>
                      <ProfileTags fields={detectedProfileFields} accent="#666"/>
                    </div>
                  )}
                  <div style={{fontSize:11,color:"#555",marginTop:12}}>{"Analyzing "+convertedCount+" converted leads"}</div>
                </div>
                <RunButton onClick={runNotesAnalysis} disabled={selectedNoteFields.length===0} loading={notesLoading} color={ACCENT} label="Analyze Converted Notes + Profiles"/>
                {notesInsight && <AIPanel title="Notes + Profile Analysis" content={notesInsight} accent={ACCENT3} icon="✦" onExport={()=>handleExport("notes","Notes and Profile Analysis",notesInsight)}/>}
              </>
            )}
          </div>
        )}

        {/* LINKEDIN */}
        {analysis && activeTab==="linkedin" && (
          <div className="up">
            {detectedLinkedInFields.length===0 ? (
              <EmptyState icon="🔗" title="No LinkedIn columns detected" subtitle="Expected: LinkedIn Summary, LinkedIn Skills, Job History, LinkedIn Headline"/>
            ) : (
              <>
                <div style={{background:"rgba(255,255,255,0.02)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:14,padding:"20px 24px",marginBottom:20}}>
                  <SectionLabel text="Select LinkedIn Fields" accent="#4D9FFF"/>
                  <FieldToggle fields={detectedLinkedInFields} selected={selectedLinkedInFields} onToggle={toggleField(setSelectedLinkedInFields)} activeColor="#4D9FFF"/>
                  <div style={{fontSize:11,color:"#555",marginTop:12}}>{"Analyzing "+convertedCount+" converted leads"}</div>
                </div>
                <RunButton onClick={runLinkedInAnalysis} disabled={selectedLinkedInFields.length===0} loading={linkedinLoading} color="#4D9FFF" label="Analyze LinkedIn + Tech Context"/>
                {linkedinInsight && <AIPanel title="LinkedIn Profile Intelligence" content={linkedinInsight} accent="#4D9FFF" icon="🔗" onExport={()=>handleExport("linkedin","LinkedIn Profile Intelligence",linkedinInsight)}/>}
              </>
            )}
          </div>
        )}

        {/* DISQUALIFIED */}
        {analysis && activeTab==="disq" && (
          <div className="up">
            {!analysis.statusCol ? (
              <EmptyState icon="🚫" title="No Lead Status column detected" subtitle="Add a Lead Status column with Disqualified as one of the values"/>
            ) : (
              <>
                <div style={{background:ACCENT2+"09",border:"1px solid "+ACCENT2+"20",borderRadius:14,padding:"20px 24px",marginBottom:20}}>
                  <SectionLabel text="Disqualification Theme Analysis" accent={ACCENT2}/>
                  <p style={{fontSize:12,color:"#aaa",margin:"0 0 14px",lineHeight:1.7}}>
                    {"Analyzing "}<span style={{color:ACCENT2,fontWeight:700}}>{disqCount+" disqualified leads"}</span>{" to identify who to stop targeting."}
                  </p>
                  {detectedNoteFields.length>0 && (
                    <>
                      <div style={{fontSize:11,color:"#777",marginBottom:8,textTransform:"uppercase",letterSpacing:1}}>Note columns:</div>
                      <FieldToggle fields={detectedNoteFields} selected={selectedNoteFields} onToggle={toggleField(setSelectedNoteFields)} activeColor={ACCENT2}/>
                    </>
                  )}
                  {detectedProfileFields.length>0 && (
                    <div style={{marginTop:14,paddingTop:14,borderTop:"1px solid rgba(255,255,255,0.06)"}}>
                      <div style={{fontSize:11,color:"#666",marginBottom:8}}>Profile attributes included:</div>
                      <ProfileTags fields={detectedProfileFields} accent={ACCENT2}/>
                    </div>
                  )}
                </div>
                <RunButton onClick={runDisqAnalysis} loading={disqLoading} color={ACCENT2} label="Analyze Disqualification Themes"/>
                {disqInsight && <AIPanel title="Disqualification Theme Analysis" content={disqInsight} accent={ACCENT2} icon="🚫" onExport={()=>handleExport("disq","Disqualification Theme Analysis",disqInsight)}/>}
              </>
            )}
          </div>
        )}

        {/* STUCK */}
        {analysis && activeTab==="stuck" && (
          <div className="up">
            {!analysis.statusCol ? (
              <EmptyState icon="⏸" title="No Lead Status column detected" subtitle="Add a Lead Status column with values like Nurture - Marketing, Nurture - Sales, Engaged, Meeting Completed, or Working"/>
            ) : stuckCount===0 ? (
              <EmptyState icon="⏸" title="No stuck leads found" subtitle="Expected Lead Status values: Nurture - Marketing, Nurture - Sales, Engaged, Meeting Completed, Working"/>
            ) : (
              <>
                <div style={{background:ACCENT4+"09",border:"1px solid "+ACCENT4+"22",borderRadius:14,padding:"20px 24px",marginBottom:20}}>
                  <SectionLabel text="Stuck in Pipeline - Gap Analysis" accent={ACCENT4}/>
                  <p style={{fontSize:13,color:"#aaa",margin:"0 0 16px",lineHeight:1.7}}>
                    {"Comparing "}<span style={{color:ACCENT4,fontWeight:700}}>{stuckCount+" stuck leads"}</span>{" against "}<span style={{color:ACCENT3,fontWeight:700}}>{convertedCount+" converted leads"}</span>{" to identify what qualification is missing."}
                  </p>
                  {detectedNoteFields.length>0 && (
                    <>
                      <div style={{fontSize:11,color:"#777",marginBottom:8,textTransform:"uppercase",letterSpacing:1}}>Select fields to analyze:</div>
                      <FieldToggle fields={detectedNoteFields} selected={selectedStuckFields} onToggle={toggleField(setSelectedStuckFields)} activeColor={ACCENT4} labelMap={STUCK_LABELS}/>
                    </>
                  )}
                  {detectedProfileFields.length>0 && (
                    <div style={{marginTop:14,paddingTop:14,borderTop:"1px solid rgba(255,255,255,0.06)"}}>
                      <div style={{fontSize:11,color:"#666",marginBottom:8}}>Profile attributes included:</div>
                      <ProfileTags fields={detectedProfileFields} accent={ACCENT4}/>
                    </div>
                  )}
                </div>
                <RunButton onClick={runStuckAnalysis} disabled={selectedStuckFields.length===0} loading={stuckLoading} color={ACCENT4} label="Analyze Stuck vs Converted"/>
                {stuckInsight && <AIPanel title="Stuck in Pipeline Gap Analysis" content={stuckInsight} accent={ACCENT4} icon="⏸" onExport={()=>handleExport("stuck","Stuck in Pipeline Gap Analysis",stuckInsight)}/>}
              </>
            )}
          </div>
        )}

        {/* AE PERFORMANCE */}
        {analysis && activeTab==="ae" && (
          <div className="up">
            {!colMap.ae_owner ? (
              <EmptyState icon="👤" title="No AE column detected" subtitle="Add a column named AE, Account Executive, Owner, or Rep with the AE name for each lead"/>
            ) : !aeData||aeData.length===0 ? (
              <EmptyState icon="👤" title="No AE data found" subtitle="Make sure your AE column has names populated"/>
            ) : (
              <>
                <div style={{background:ACCENT5+"09",border:"1px solid "+ACCENT5+"22",borderRadius:14,padding:"20px 24px",marginBottom:24}}>
                  <SectionLabel text="AE Performance Leaderboard" accent={ACCENT5}/>

                  <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:16}}>
                    <div style={{fontSize:12,color:"#666"}}>
                      {"Team avg: "}
                      <span style={{color:ACCENT5,fontWeight:700,fontFamily:"monospace"}}>{(teamAvgRate*100).toFixed(1)+"%"}</span>
                      {" across "+aeData.length+" AEs"}
                    </div>
                  </div>

                  {detectedNoteFields.length>0 && (
                    <div style={{marginBottom:20,paddingBottom:16,borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
                      <div style={{fontSize:11,color:"#555",marginBottom:8,textTransform:"uppercase",letterSpacing:1}}>Note fields used in AI analysis:</div>
                      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                        {detectedNoteFields.map(f=>{
                          const on = aeNoteFields.includes(f);
                          return (
                            <button key={f} onClick={()=>toggleAeNote(f)} style={{background:on?ACCENT5+"18":"rgba(255,255,255,0.03)",border:"1px solid "+(on?ACCENT5:"rgba(255,255,255,0.09)"),color:on?ACCENT5:"#555",borderRadius:99,padding:"4px 11px",cursor:"pointer",fontSize:11,transition:"all 0.15s"}}>
                              {NOTE_LABELS[f]||f.replace(/_/g," ")}
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  <div style={{display:"flex",flexDirection:"column",gap:10}}>
                    {aeData.map((ae,i)=>{
                      const color    = ae.rate>=teamAvgRate ? ACCENT3 : ae.rate<teamAvgRate*0.6 ? ACCENT2 : ACCENT4;
                      const isActive = selectedAE===ae.name;
                      const vsTeam   = (ae.rate-teamAvgRate)*100;
                      return (
                        <div key={ae.name} onClick={()=>setSelectedAE(isActive?null:ae.name)}
                          style={{background:isActive?color+"0d":"rgba(255,255,255,0.02)",border:"1px solid "+(isActive?color+"44":"rgba(255,255,255,0.07)"),borderRadius:12,padding:"14px 18px",cursor:"pointer",transition:"all 0.15s"}}>
                          <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                            <span style={{fontFamily:"monospace",fontSize:11,color:"#444",minWidth:22}}>{"#"+(i+1)}</span>
                            <span style={{fontWeight:700,fontSize:14,flex:1,color:isActive?color:"#ddd"}}>{ae.name}</span>
                            <div style={{display:"flex",gap:14,alignItems:"center",flexWrap:"wrap"}}>
                              <span style={{fontFamily:"monospace",fontSize:17,fontWeight:800,color}}>{(ae.rate*100).toFixed(1)+"%"}</span>
                              <span style={{fontSize:11,color:"#666"}}>{ae.converted+"/"+ae.total+" conv"}</span>
                              {ae.stuck>0&&<span style={{fontSize:11,color:ACCENT4}}>{ae.stuck+" stuck"}</span>}
                              {ae.disq>0&&<span style={{fontSize:11,color:ACCENT2}}>{ae.disq+" disq"}</span>}
                              <span style={{fontSize:10,color:vsTeam>=0?ACCENT3:ACCENT2,fontFamily:"monospace"}}>{(vsTeam>=0?"+":"")+vsTeam.toFixed(1)+"% vs avg"}</span>
                              <div style={{width:72}}><Bar pct={(ae.rate/(aeData[0].rate||1))*100} color={color}/></div>
                            </div>
                            <span style={{fontSize:10,color:"#333",fontFamily:"monospace"}}>{isActive?"▲":"▼"}</span>
                          </div>

                          {isActive && (
                            <div style={{marginTop:16,paddingTop:16,borderTop:"1px solid rgba(255,255,255,0.07)"}} onClick={e=>e.stopPropagation()}>
                              <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:16}}>
                                {[
                                  {l:"Meetings",v:ae.total,c:"#fff"},
                                  {l:"Converted",v:ae.converted,c:ACCENT3},
                                  {l:"Conv Rate",v:(ae.rate*100).toFixed(1)+"%",c:color},
                                  {l:"vs Team Avg",v:(vsTeam>=0?"+":"")+vsTeam.toFixed(1)+"%",c:vsTeam>=0?ACCENT3:ACCENT2},
                                  {l:"Stuck",v:ae.stuck,c:ACCENT4},
                                  {l:"Disq",v:ae.disq,c:ACCENT2},
                                ].map(s=>(
                                  <div key={s.l} style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:"10px 14px",minWidth:80}}>
                                    <div style={{fontSize:16,fontWeight:700,fontFamily:"monospace",color:s.c}}>{s.v}</div>
                                    <div style={{fontSize:9,color:"#555",textTransform:"uppercase",letterSpacing:1,marginTop:2}}>{s.l}</div>
                                  </div>
                                ))}
                              </div>
                              <RunButton onClick={()=>runAEAnalysis(ae)} loading={aeLoadingFor===ae.name} color={color} label={"Run Coaching Analysis for "+ae.name} loadingLabel="Analyzing..."/>
                              {aeInsights[ae.name] && (
                                <div style={{background:"rgba(0,0,0,0.15)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:10,padding:"18px 20px"}}>
                                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
                                    <span style={{fontFamily:"monospace",fontSize:10,letterSpacing:2,color,textTransform:"uppercase"}}>{ae.name+" - Coaching Analysis"}</span>
                                    <button onClick={()=>handleExport("ae_"+ae.name,"AE Coaching - "+ae.name,aeInsights[ae.name])} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#aaa",borderRadius:7,padding:"5px 12px",cursor:"pointer",fontSize:11,display:"flex",alignItems:"center",gap:5}}>
                                      <span>&#8595; Export</span>
                                    </button>
                                  </div>
                                  <MarkdownBlock content={aeInsights[ae.name]} accent={color}/>
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>

                <div style={{background:"rgba(255,255,255,0.02)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:14,padding:"20px 24px"}}>
                  <SectionLabel text="Team-Wide Coaching Analysis" accent={ACCENT5}/>
                  <p style={{fontSize:12,color:"#aaa",margin:"0 0 16px",lineHeight:1.7}}>
                    Compare top and bottom performers to surface systemic patterns, team-wide qualification gaps, and coaching priorities.
                  </p>
                  <RunButton onClick={runAETeamAnalysis} loading={aeTeamLoading} color={ACCENT5} label="Analyze Full Team" loadingLabel="Analyzing team..."/>
                  {aeTeamInsight && <AIPanel title="Team Coaching Analysis" content={aeTeamInsight} accent={ACCENT5} icon="👥" onExport={()=>handleExport("team","Team Coaching Analysis",aeTeamInsight)}/>}
                </div>
              </>
            )}
          </div>
        )}

        {/* WHO NOT TO TARGET */}
        {analysis && activeTab==="neg" && (
          <div className="up">
            <div style={{background:ACCENT2+"09",border:"1px solid "+ACCENT2+"22",borderRadius:14,padding:"20px 24px",marginBottom:24}}>
              <SectionLabel text="Negative ICP — Who NOT to Target" accent={ACCENT2}/>
              <p style={{fontSize:13,color:"#aaa",margin:"0 0 16px",lineHeight:1.7}}>
                Analyze disqualified leads and low-quality meetings to define who to exclude from targeting. Compares against converted leads for contrast.
              </p>

              <div style={{display:"flex",gap:24,flexWrap:"wrap",marginBottom:20}}>
                <div>
                  <div style={{fontSize:11,color:"#555",marginBottom:6,textTransform:"uppercase",letterSpacing:1}}>Sources included:</div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    {analysis.statusCol && (
                      <Tag bg={ACCENT2+"18"} color={ACCENT2}>{disqCount+" disqualified leads"}</Tag>
                    )}
                    {colMap.iqm_score && (
                      <Tag bg={ACCENT4+"18"} color={ACCENT4}>{"IQM Score: "+colMap.iqm_score}</Tag>
                    )}
                    {colMap.ae_score && (
                      <Tag bg={ACCENT4+"18"} color={ACCENT4}>{"AE Score: "+colMap.ae_score}</Tag>
                    )}
                    {!analysis.statusCol && !colMap.iqm_score && !colMap.ae_score && (
                      <span style={{fontSize:12,color:"#555"}}>No disqualified leads or score columns detected. Add a Lead Status column with Disqualified values, or columns named IQM Score or AE Score.</span>
                    )}
                  </div>
                </div>

                {(colMap.iqm_score || colMap.ae_score) && (
                  <div>
                    <div style={{fontSize:11,color:"#555",marginBottom:6,textTransform:"uppercase",letterSpacing:1}}>Low score threshold (include scores at or below):</div>
                    <div style={{display:"flex",gap:6}}>
                      {[1,2,3,4,5].map(n => (
                        <button key={n} onClick={()=>setNegScoreThreshold(n)} style={{background:negScoreThreshold===n?ACCENT2+"22":"rgba(255,255,255,0.03)",border:"1px solid "+(negScoreThreshold===n?ACCENT2:"rgba(255,255,255,0.09)"),color:negScoreThreshold===n?ACCENT2:"#666",borderRadius:8,padding:"5px 12px",cursor:"pointer",fontSize:12,fontWeight:negScoreThreshold===n?700:400}}>
                          {n}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {detectedNoteFields.length > 0 && (
                <div style={{marginBottom:16,paddingBottom:16,borderTop:"1px solid rgba(255,255,255,0.06)",paddingTop:16}}>
                  <div style={{fontSize:11,color:"#555",marginBottom:8,textTransform:"uppercase",letterSpacing:1}}>Note fields to include:</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                    {detectedNoteFields.map(f => {
                      const on = negNoteFields.includes(f);
                      return (
                        <button key={f} onClick={()=>setNegNoteFields(p=>p.includes(f)?p.filter(x=>x!==f):[...p,f])} style={{background:on?ACCENT2+"18":"rgba(255,255,255,0.03)",border:"1px solid "+(on?ACCENT2:"rgba(255,255,255,0.09)"),color:on?ACCENT2:"#555",borderRadius:99,padding:"4px 11px",cursor:"pointer",fontSize:11,transition:"all 0.15s"}}>
                          {NOTE_LABELS[f]||f.replace(/_/g," ")}
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}

              {detectedProfileFields.length > 0 && (
                <div>
                  <div style={{fontSize:11,color:"#666",marginBottom:8}}>Profile attributes always included:</div>
                  <ProfileTags fields={detectedProfileFields} accent={ACCENT2}/>
                </div>
              )}
            </div>

            <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:8}}>
              <RunButton
                onClick={runNegAnalysis}
                loading={negLoading}
                color={ACCENT2}
                label="Analyze Bad-Fit Patterns"
                loadingLabel="Analyzing bad-fit patterns..."
                disabled={!analysis.statusCol && !colMap.iqm_score && !colMap.ae_score}
              />
              {detectedLinkedInFields.length > 0 && (
                <RunButton
                  onClick={runNegPersonaAnalysis}
                  loading={negPersonaLoading}
                  color="#E06030"
                  label="Compare Personas (LinkedIn)"
                  loadingLabel="Comparing personas..."
                  disabled={negLinkedInFields.length===0}
                />
              )}
            </div>

            {detectedLinkedInFields.length > 0 && (
              <div style={{background:"rgba(224,96,48,0.06)",border:"1px solid rgba(224,96,48,0.18)",borderRadius:12,padding:"14px 16px",marginBottom:20}}>
                <div style={{fontSize:11,color:"#E06030",marginBottom:8,textTransform:"uppercase",letterSpacing:1}}>LinkedIn fields for persona comparison:</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                  {detectedLinkedInFields.map(f => {
                    const on = negLinkedInFields.includes(f);
                    return (
                      <button key={f} onClick={()=>setNegLinkedInFields(p=>p.includes(f)?p.filter(x=>x!==f):[...p,f])} style={{background:on?"rgba(224,96,48,0.18)":"rgba(255,255,255,0.03)",border:"1px solid "+(on?"#E06030":"rgba(255,255,255,0.09)"),color:on?"#E06030":"#555",borderRadius:99,padding:"4px 11px",cursor:"pointer",fontSize:11,transition:"all 0.15s"}}>
                        {f.replace(/linkedin_/,"").replace(/_/g," ")}
                      </button>
                    );
                  })}
                </div>
                {negLinkedInFields.length === 0 && <div style={{fontSize:11,color:"#555",marginTop:6}}>Select at least one LinkedIn field to enable persona comparison.</div>}
              </div>
            )}

            {negInsight && <AIPanel title="Bad-Fit Pattern Analysis" content={negInsight} accent={ACCENT2} icon="🚫" onExport={()=>handleExport("neg","Negative ICP - Who NOT to Target",negInsight)}/>}
            {negPersonaInsight && <AIPanel title="Persona Comparison — Converted vs Bad Fit" content={negPersonaInsight} accent="#E06030" icon="👤" onExport={()=>handleExport("neg_persona","Persona Comparison",negPersonaInsight)}/>}
          </div>
        )}


        {/* ACCOUNT RE-ENGAGEMENT */}
        {analysis && activeTab==="reengage" && (
          <div className="up">
            <div style={{background:"rgba(77,159,255,0.06)",border:"1px solid rgba(77,159,255,0.18)",borderRadius:14,padding:"20px 24px",marginBottom:24}}>
              <SectionLabel text="Account Re-engagement Opportunities" accent="#4D9FFF"/>
              <p style={{fontSize:13,color:"#aaa",margin:"0 0 16px",lineHeight:1.7}}>
                Identify non-converted accounts where meeting notes suggest company-level fit — good candidates for SDR re-outreach or targeted ad campaigns even though the contact didn't progress.
              </p>

              {colMap.account_name ? (
                <Tag bg="rgba(77,159,255,0.12)" color="#4D9FFF">{"Account column: "+colMap.account_name}</Tag>
              ) : (
                <div style={{fontSize:12,color:"#666",marginBottom:12}}>No account name column detected — analysis will run at the lead level. Add a column named "Account Name" or "Company" for account-level grouping.</div>
              )}

              {detectedNoteFields.length > 0 && (
                <div style={{marginTop:16,paddingTop:16,borderTop:"1px solid rgba(255,255,255,0.06)"}}>
                  <div style={{fontSize:11,color:"#555",marginBottom:8,textTransform:"uppercase",letterSpacing:1}}>Note fields to include:</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                    {detectedNoteFields.map(f => {
                      const on = reengageNoteFields.includes(f);
                      return (
                        <button key={f} onClick={()=>setReengageNoteFields(p=>p.includes(f)?p.filter(x=>x!==f):[...p,f])} style={{background:on?"rgba(77,159,255,0.15)":"rgba(255,255,255,0.03)",border:"1px solid "+(on?"#4D9FFF":"rgba(255,255,255,0.09)"),color:on?"#4D9FFF":"#555",borderRadius:99,padding:"4px 11px",cursor:"pointer",fontSize:11,transition:"all 0.15s"}}>
                          {NOTE_LABELS[f]||f.replace(/_/g," ")}
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}

              {detectedProfileFields.length > 0 && (
                <div style={{marginTop:12}}>
                  <div style={{fontSize:11,color:"#666",marginBottom:6}}>Profile attributes always included:</div>
                  <ProfileTags fields={detectedProfileFields} accent="#4D9FFF"/>
                </div>
              )}
            </div>

            <RunButton onClick={runReengageAnalysis} loading={reengageLoading} color="#4D9FFF" label="Identify Re-engagement Opportunities" loadingLabel="Analyzing accounts..."/>
            {reengageInsight && <AIPanel title="Account Re-engagement Opportunities" content={reengageInsight} accent="#4D9FFF" icon="🔁" onExport={()=>handleExport("reengage","Account Re-engagement Opportunities",reengageInsight)}/>}
          </div>
        )}


      </div>
    </div>
  );
}
