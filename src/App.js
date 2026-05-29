import { useState, useEffect, useCallback } from "react";
import { db, ref, onValue, set } from "./firebase";
import * as XLSX from "xlsx";

const DEFAULT_USERS = ["陳世恆", "陳奕輝", "江翠媚", "陳庭萱"];
const CATEGORIES_IN  = ["營業收入", "其他收入"];
const CATEGORIES_OUT = ["食材採購", "水電費", "租金", "人事費用", "包裝材料", "其他支出"];

const BATIK_PATTERN = `url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23c8860a' fill-opacity='0.06'%3E%3Cpath d='M30 0c16.569 0 30 13.431 30 30S46.569 60 30 60 0 46.569 0 30 13.431 0 30 0zm0 5C16.193 5 5 16.193 5 30s11.193 25 25 25 25-11.193 25-25S43.807 5 30 5zm0 8c9.389 0 17 7.611 17 17s-7.611 17-17 17-17-7.611-17-17 7.611-17 17-17zm0 5c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E")`;

const USER_COLORS = ["#e07b39","#2d8a6e","#c8393b","#5b7fa6"];
const formatCurrency = (n) => "NT$" + Math.abs(n).toLocaleString("zh-TW");
const today = () => new Date().toISOString().slice(0, 10);
const nowY = new Date().getFullYear();
const nowM = new Date().getMonth() + 1;
function generateId() { return Date.now().toString(36) + Math.random().toString(36).slice(2,6); }
function monthStart(y,m){ return `${y}-${String(m).padStart(2,"0")}-01`; }
function monthEnd(y,m){ const d=new Date(y,m,0).getDate(); return `${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`; }

// ── 雲端儲存 ──

// ── Excel 報表產生 ──
function exportExcel(records, users, rangeLabel){
  const wb = XLSX.utils.book_new();
  const sorted = [...records].sort((a,b)=>a.date.localeCompare(b.date)||a.id.localeCompare(b.id));

  const H = (v,extra={})=>({v, s:{font:{bold:true,name:"Arial",sz:10,color:{rgb:"FFFFFFFF"},...(extra.font||{})},
    fill:{fgColor:{rgb:"FF7A2000"},patternType:"solid"},
    alignment:{horizontal:"center",vertical:"center",...(extra.align||{})},
    border:{top:{style:"thin",color:{rgb:"FFCC9944"}},bottom:{style:"thin",color:{rgb:"FFCC9944"}},
      left:{style:"thin",color:{rgb:"FFCC9944"}},right:{style:"thin",color:{rgb:"FFCC9944"}}},
    ...extra}});

  const C = (v,extra={})=>({v, s:{font:{name:"Arial",sz:10,...(extra.font||{})},
    border:{top:{style:"thin",color:{rgb:"FFEEDDBB"}},bottom:{style:"thin",color:{rgb:"FFEEDDBB"}},
      left:{style:"thin",color:{rgb:"FFEEDDBB"}},right:{style:"thin",color:{rgb:"FFEEDDBB"}}},
    alignment:{vertical:"center",...(extra.align||{})},
    fill:extra.fill||{fgColor:{rgb:"FFFFFFFF"},patternType:"solid"},
    ...extra}});

  const N = (v,color,fill)=>({v,t:"n",s:{font:{name:"Arial",sz:10,bold:true,color:{rgb:color||"FF222222"}},
    numFmt:"#,##0",alignment:{horizontal:"right",vertical:"center"},
    fill:fill?{fgColor:{rgb:fill},patternType:"solid"}:{fgColor:{rgb:"FFFFFFFF"},patternType:"solid"},
    border:{top:{style:"thin",color:{rgb:"FFEEDDBB"}},bottom:{style:"thin",color:{rgb:"FFEEDDBB"}},
      left:{style:"thin",color:{rgb:"FFEEDDBB"}},right:{style:"thin",color:{rgb:"FFEEDDBB"}}}}});

  const FOOT = (v,extra={})=>({v,s:{font:{bold:true,name:"Arial",sz:10,color:{rgb:"FFFFFFFF"},...(extra.font||{})},
    fill:{fgColor:{rgb:"FF3D2000"},patternType:"solid"},
    alignment:{horizontal:extra.right?"right":"center",vertical:"center"},
    numFmt:extra.numFmt||"@",
    border:{top:{style:"medium",color:{rgb:"FFCC9944"}},bottom:{style:"medium",color:{rgb:"FFCC9944"}},
      left:{style:"thin",color:{rgb:"FFCC9944"}},right:{style:"thin",color:{rgb:"FFCC9944"}}}}});

  const totalIncome  = sorted.filter(r=>r.type==="income").reduce((s,r)=>s+r.amount,0);
  const totalExpense = sorted.filter(r=>r.type==="expense").reduce((s,r)=>s+r.amount,0);
  const totalUnpaid  = sorted.filter(r=>r.type==="expense"&&r.paid===false).reduce((s,r)=>s+r.amount,0);

  // ── 工作表1：明細帳 ──
  const title1 = [{v:`mega 記帳  ｜  明細帳　　${rangeLabel}　　共 ${sorted.length} 筆`,
    s:{font:{bold:true,sz:14,name:"Arial",color:{rgb:"FF5C2000"}},alignment:{horizontal:"left"}}}];

  const hdr1 = ["日期","類型","類別","金額（NT$）","付款狀態","記帳人","備註"].map(h=>H(h));

  const rows1 = sorted.map((r,i)=>{
    const isIn = r.type==="income";
    const paid = r.paid!==false;
    const rowFill = i%2===0?"FFFFFFFF":"FFFFF8F0";
    const incFill = i%2===0?"FFF0FBF5":"FFE6F7EF";
    const expFill = paid?(i%2===0?"FFFDF0F0":"FFF9E8E8"):(i%2===0?"FFFFF8E0":"FFFDF3D0");
    const fill = isIn?incFill:expFill;
    return [
      C(r.date,{align:{horizontal:"center"},fill:{fgColor:{rgb:fill},patternType:"solid"}}),
      C(isIn?"收入":"支出",{font:{bold:true,name:"Arial",sz:10,color:{rgb:isIn?"FF0E6B3A":"FF8B1010"}},
        align:{horizontal:"center"},fill:{fgColor:{rgb:fill},patternType:"solid"}}),
      C(r.category,{fill:{fgColor:{rgb:fill},patternType:"solid"}}),
      N(r.amount, isIn?"FF0E6B3A":paid?"FF8B1010":"FFAA6600", fill),
      C(isIn?"—":paid?"✓ 已付款":"⏳ 未付款",{
        font:{name:"Arial",sz:10,color:{rgb:isIn?"FFAAAAAA":paid?"FF0E6B3A":"FFAA6600"}},
        align:{horizontal:"center"},fill:{fgColor:{rgb:fill},patternType:"solid"}}),
      C(r.user,{align:{horizontal:"center"},fill:{fgColor:{rgb:fill},patternType:"solid"}}),
      C(r.note||"—",{font:{name:"Arial",sz:10,color:{rgb:"FF888888"}},fill:{fgColor:{rgb:fill},patternType:"solid"}}),
    ];
  });

  const foot1 = [
    [FOOT(""),FOOT(""),FOOT("收入合計",{right:true}),{...FOOT("",{numFmt:"#,##0"}),v:totalIncome,t:"n",
      s:{...FOOT("").s,font:{bold:true,name:"Arial",sz:10,color:{rgb:"FF5dde8a"}},numFmt:"#,##0",alignment:{horizontal:"right",vertical:"center"}}},
     FOOT(""),FOOT(""),FOOT("")],
    [FOOT(""),FOOT(""),FOOT("支出合計",{right:true}),{...FOOT("",{numFmt:"#,##0"}),v:totalExpense,t:"n",
      s:{...FOOT("").s,font:{bold:true,name:"Arial",sz:10,color:{rgb:"FFff8080"}},numFmt:"#,##0",alignment:{horizontal:"right",vertical:"center"}}},
     FOOT(""),FOOT(""),FOOT("")],
    [FOOT(""),FOOT(""),FOOT("未付款",{right:true}),{...FOOT(""),v:totalUnpaid,t:"n",
      s:{...FOOT("").s,font:{bold:true,name:"Arial",sz:10,color:{rgb:"FFffb400"}},numFmt:"#,##0",alignment:{horizontal:"right",vertical:"center"}}},
     FOOT(""),FOOT(""),FOOT("")],
    [FOOT(""),FOOT(""),FOOT("淨利潤",{right:true}),{...FOOT(""),v:totalIncome-totalExpense,t:"n",
      s:{...FOOT("").s,font:{bold:true,name:"Arial",sz:11,color:{rgb:totalIncome-totalExpense>=0?"FF5dde8a":"FFff7070"}},
        numFmt:"#,##0",alignment:{horizontal:"right",vertical:"center"}}},
     FOOT(""),FOOT(""),FOOT("")],
  ];

  const ws1data = [title1,[],hdr1,...rows1,[],...foot1];
  const ws1 = XLSX.utils.aoa_to_sheet(ws1data);
  ws1["!cols"] = [{wch:13},{wch:8},{wch:14},{wch:16},{wch:12},{wch:10},{wch:26}];
  ws1["!merges"] = [{s:{r:0,c:0},e:{r:0,c:6}}];
  ws1["!rows"] = [{hpt:24},...Array(ws1data.length-1).fill({hpt:20})];
  XLSX.utils.book_append_sheet(wb, ws1, "明細帳");

  // ── 工作表2：月份摘要 ──
  const byMonth = {};
  sorted.forEach(r=>{
    const ym=r.date.slice(0,7);
    if(!byMonth[ym]) byMonth[ym]={ym,income:0,expense:0,unpaid:0,count:0};
    if(r.type==="income") byMonth[ym].income+=r.amount;
    else byMonth[ym].expense+=r.amount;
    if(r.type==="expense"&&r.paid===false) byMonth[ym].unpaid+=r.amount;
    byMonth[ym].count++;
  });
  const months=Object.values(byMonth).sort((a,b)=>a.ym.localeCompare(b.ym));

  const title2=[{v:`mega 記帳  ｜  月份摘要`,
    s:{font:{bold:true,sz:14,name:"Arial",color:{rgb:"FF5C2000"}},alignment:{horizontal:"left"}}}];
  const hdr2=["月份","收入（NT$）","支出（NT$）","淨利潤（NT$）","未付款（NT$）","筆數"].map(h=>H(h));

  const rows2=months.map((m,i)=>{
    const profit=m.income-m.expense;
    const fill=i%2===0?"FFFFFFFF":"FFFFF8F0";
    const f={fgColor:{rgb:fill},patternType:"solid"};
    return [
      C(m.ym.replace("-","/"),{align:{horizontal:"center"},fill:f}),
      N(m.income,"FF0E6B3A",fill),
      N(m.expense,"FF8B1010",fill),
      N(profit,profit>=0?"FF0E6B3A":"FF8B1010",fill),
      N(m.unpaid,"FFAA6600",fill),
      C(m.count,{align:{horizontal:"center"},fill:f}),
    ];
  });

  const totM={income:months.reduce((s,m)=>s+m.income,0),expense:months.reduce((s,m)=>s+m.expense,0),
    unpaid:months.reduce((s,m)=>s+m.unpaid,0),count:sorted.length};
  const foot2=[[FOOT("合計"),
    {...FOOT(""),v:totM.income,t:"n",s:{...FOOT("").s,font:{bold:true,name:"Arial",sz:10,color:{rgb:"FF5dde8a"}},numFmt:"#,##0",alignment:{horizontal:"right",vertical:"center"}}},
    {...FOOT(""),v:totM.expense,t:"n",s:{...FOOT("").s,font:{bold:true,name:"Arial",sz:10,color:{rgb:"FFff8080"}},numFmt:"#,##0",alignment:{horizontal:"right",vertical:"center"}}},
    {...FOOT(""),v:totM.income-totM.expense,t:"n",s:{...FOOT("").s,font:{bold:true,name:"Arial",sz:10,color:{rgb:totM.income-totM.expense>=0?"FF5dde8a":"FFff7070"}},numFmt:"#,##0",alignment:{horizontal:"right",vertical:"center"}}},
    {...FOOT(""),v:totM.unpaid,t:"n",s:{...FOOT("").s,font:{bold:true,name:"Arial",sz:10,color:{rgb:"FFffb400"}},numFmt:"#,##0",alignment:{horizontal:"right",vertical:"center"}}},
    {...FOOT(""),v:totM.count,t:"n",s:{...FOOT("").s,numFmt:"#,##0",alignment:{horizontal:"center",vertical:"center"}}},
  ]];

  const ws2data=[title2,[],hdr2,...rows2,[],...foot2];
  const ws2=XLSX.utils.aoa_to_sheet(ws2data);
  ws2["!cols"]=[{wch:10},{wch:16},{wch:16},{wch:16},{wch:14},{wch:8}];
  ws2["!merges"]=[{s:{r:0,c:0},e:{r:0,c:5}}];
  ws2["!rows"]=[{hpt:24},...Array(ws2data.length-1).fill({hpt:20})];
  XLSX.utils.book_append_sheet(wb,ws2,"月份摘要");

  // ── 工作表3：人員統計 ──
  const title3=[{v:`mega 記帳  ｜  人員統計`,
    s:{font:{bold:true,sz:14,name:"Arial",color:{rgb:"FF5C2000"}},alignment:{horizontal:"left"}}}];
  const hdr3=["人員","收入筆數","收入金額","支出筆數","支出金額","未付款","淨貢獻"].map(h=>H(h));

  const rows3=users.map((u,i)=>{
    const uRecs=sorted.filter(r=>r.user===u);
    const uIn=uRecs.filter(r=>r.type==="income");
    const uOut=uRecs.filter(r=>r.type==="expense");
    const inAmt=uIn.reduce((s,r)=>s+r.amount,0);
    const outAmt=uOut.reduce((s,r)=>s+r.amount,0);
    const unpaid=uOut.filter(r=>r.paid===false).reduce((s,r)=>s+r.amount,0);
    const net=inAmt-outAmt;
    const fill=i%2===0?"FFFFFFFF":"FFFFF8F0";
    const f={fgColor:{rgb:fill},patternType:"solid"};
    return [
      C(u,{font:{bold:true,name:"Arial",sz:10},fill:f}),
      C(uIn.length,{align:{horizontal:"center"},fill:f}),
      N(inAmt,"FF0E6B3A",fill),
      C(uOut.length,{align:{horizontal:"center"},fill:f}),
      N(outAmt,"FF8B1010",fill),
      N(unpaid,"FFAA6600",fill),
      N(net,net>=0?"FF0E6B3A":"FF8B1010",fill),
    ];
  });

  const ws3data=[title3,[],hdr3,...rows3];
  const ws3=XLSX.utils.aoa_to_sheet(ws3data);
  ws3["!cols"]=[{wch:12},{wch:10},{wch:14},{wch:10},{wch:14},{wch:12},{wch:14}];
  ws3["!merges"]=[{s:{r:0,c:0},e:{r:0,c:6}}];
  ws3["!rows"]=[{hpt:24},...Array(ws3data.length-1).fill({hpt:20})];
  XLSX.utils.book_append_sheet(wb,ws3,"人員統計");

  XLSX.writeFile(wb,`mega記帳_${rangeLabel.replace(/\s/g,"")}_${today()}.xlsx`);
}

// ── UI 元件 ──
const inputStyle={width:"100%",background:"#2a1500",border:"1px solid rgba(200,134,10,0.3)",
  borderRadius:10,padding:"10px 14px",color:"#f5e6c8",fontSize:14,boxSizing:"border-box",outline:"none"};

function FormField({label,children}){
  return <div style={{marginBottom:12}}>
    <div style={{fontSize:11,opacity:0.5,marginBottom:5,letterSpacing:1}}>{label}</div>
    {children}
  </div>;
}

function CategoryBar({label,amount,max,color}){
  const pct=max>0?(amount/max)*100:0;
  return <div style={{marginBottom:8}}>
    <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}>
      <span style={{opacity:0.7}}>{label}</span>
      <span style={{color,fontWeight:600}}>NT${amount.toLocaleString("zh-TW")}</span>
    </div>
    <div style={{background:"rgba(255,255,255,0.07)",borderRadius:4,height:5}}>
      <div style={{width:`${pct}%`,background:color,borderRadius:4,height:5,transition:"width .4s ease"}}/>
    </div>
  </div>;
}

function RecordCard({r,getUserColor,onDelete,onTogglePaid}){
  const isIncome=r.type==="income", paid=r.paid!==false;
  return (
    <div style={{background:"#2a1500",
      border:`1px solid ${!isIncome&&!paid?"rgba(255,180,0,0.35)":"rgba(200,134,10,0.15)"}`,
      borderRadius:12,padding:"12px 14px",marginBottom:8,display:"flex",alignItems:"center",gap:12}}>
      <div style={{width:36,height:36,borderRadius:"50%",
        background:isIncome?"rgba(93,222,138,0.15)":paid?"rgba(255,128,128,0.15)":"rgba(255,180,0,0.15)",
        display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>
        {isIncome?"📈":paid?"📉":"⏳"}
      </div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontSize:13,fontWeight:600,marginBottom:2}}>{r.category}</div>
        <div style={{fontSize:11,opacity:0.5,display:"flex",gap:8,flexWrap:"wrap"}}>
          <span style={{color:getUserColor(r.user)}}>● {r.user}</span>
          <span>{r.date}</span>
          {r.note&&<span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.note}</span>}
        </div>
      </div>
      <div style={{textAlign:"right",flexShrink:0,display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4}}>
        <div style={{fontSize:14,fontWeight:700,color:isIncome?"#5dde8a":paid?"#ff8080":"#ffb400"}}>
          {isIncome?"+":"-"}NT${r.amount.toLocaleString("zh-TW")}
        </div>
        {!isIncome&&(
          <button onClick={()=>onTogglePaid(r.id)} style={{fontSize:10,fontWeight:700,padding:"2px 8px",
            borderRadius:10,cursor:"pointer",border:"none",
            background:paid?"rgba(93,222,138,0.2)":"rgba(255,180,0,0.2)",
            color:paid?"#5dde8a":"#ffb400"}}>{paid?"✓ 已付款":"⏳ 未付款"}</button>
        )}
        <button onClick={()=>onDelete(r.id)} style={{background:"none",border:"none",
          color:"rgba(255,255,255,0.18)",cursor:"pointer",fontSize:13,padding:0}}>✕</button>
      </div>
    </div>
  );
}

function MembersModal({users,onSave,onClose}){
  const [draft,setDraft]=useState([...users]);
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",display:"flex",
      alignItems:"center",justifyContent:"center",zIndex:300,padding:20}}>
      <div style={{background:"#241200",borderRadius:18,padding:24,width:"100%",maxWidth:340,
        border:"1px solid rgba(200,134,10,0.3)"}}>
        <div style={{fontSize:16,fontWeight:700,marginBottom:4}}>✏️ 編輯成員名稱</div>
        <div style={{fontSize:12,opacity:0.5,marginBottom:18}}>修改後點擊儲存生效</div>
        {draft.map((name,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
            <div style={{width:28,height:28,borderRadius:"50%",background:USER_COLORS[i]||"#888",
              display:"flex",alignItems:"center",justifyContent:"center",
              fontSize:12,fontWeight:700,flexShrink:0,color:"#fff"}}>{name[0]||"?"}</div>
            <input value={name} onChange={e=>{const d=[...draft];d[i]=e.target.value;setDraft(d);}}
              style={{...inputStyle,marginBottom:0,fontSize:13,padding:"7px 12px"}} maxLength={6}/>
          </div>
        ))}
        <div style={{display:"flex",gap:10,marginTop:18}}>
          <button onClick={onClose} style={{flex:1,background:"#3a2500",border:"1px solid rgba(255,255,255,0.1)",
            color:"#ccc",borderRadius:10,padding:"10px 0",cursor:"pointer",fontSize:14}}>取消</button>
          <button onClick={()=>onSave(draft.map(n=>n.trim()||"未命名"))} style={{
            flex:1,background:"linear-gradient(135deg,#8b2500,#c8560a)",border:"none",
            color:"#fff",borderRadius:10,padding:"10px 0",cursor:"pointer",fontSize:14,fontWeight:700}}>儲存</button>
        </div>
      </div>
    </div>
  );
}

function ExportModal({records,rangedRecords,users,rangeLabel,onClose,onToast}){
  const [exporting,setExporting]=useState(false);
  async function doExport(recs,label){
    setExporting(true);
    try{ exportExcel(recs,users,label); onToast({msg:"📊 報表已下載",ok:true}); }
    catch(e){ onToast({msg:"匯出失敗，請重試",ok:false}); }
    setExporting(false);
    onClose();
  }
  const sheets=[
    {icon:"📋",name:"明細帳",desc:"每筆完整記錄，含日期、類別、金額、付款狀態"},
    {icon:"📅",name:"月份摘要",desc:"各月收入、支出、淨利潤、未付款統計"},
    {icon:"👥",name:"人員統計",desc:"各成員收支筆數、金額彙整"},
  ];
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.8)",display:"flex",
      alignItems:"flex-end",justifyContent:"center",zIndex:300}}>
      <div style={{background:"#1e0f00",borderRadius:"22px 22px 0 0",padding:"6px 20px 36px",
        width:"100%",maxWidth:480,border:"1px solid rgba(200,134,10,0.25)"}}>
        <div style={{width:36,height:4,background:"rgba(255,255,255,0.15)",borderRadius:2,margin:"14px auto 20px"}}/>
        <div style={{fontSize:18,fontWeight:800,marginBottom:2}}>📊 匯出 Excel 報表</div>
        <div style={{fontSize:12,opacity:0.4,marginBottom:20}}>產生含三個工作表的格式化 .xlsx 檔案</div>

        {/* 工作表說明 */}
        <div style={{background:"#2a1500",borderRadius:12,padding:"4px 0",marginBottom:16,
          border:"1px solid rgba(200,134,10,0.15)"}}>
          {sheets.map(({icon,name,desc},i)=>(
            <div key={name} style={{display:"flex",gap:12,padding:"12px 14px",
              borderBottom:i<2?"1px solid rgba(200,134,10,0.1)":"none"}}>
              <span style={{fontSize:20,flexShrink:0,marginTop:1}}>{icon}</span>
              <div>
                <div style={{fontSize:13,fontWeight:700,marginBottom:2}}>工作表：{name}</div>
                <div style={{fontSize:11,opacity:0.45,lineHeight:1.5}}>{desc}</div>
              </div>
            </div>
          ))}
        </div>

        {/* 範圍資訊 */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:20}}>
          {[
            {label:"篩選期間",sub:rangeLabel,count:rangedRecords.length,color:"#e09020"},
            {label:"全部資料",sub:"所有時間",count:records.length,color:"#f5e6c8"},
          ].map(({label,sub,count,color})=>(
            <div key={label} style={{background:"#2a1500",borderRadius:10,padding:"10px 12px",
              border:"1px solid rgba(200,134,10,0.15)"}}>
              <div style={{fontSize:10,opacity:0.4,marginBottom:3}}>{label}</div>
              <div style={{fontSize:12,color,fontWeight:600,marginBottom:1}}>{sub}</div>
              <div style={{fontSize:11,opacity:0.5}}>{count} 筆記錄</div>
            </div>
          ))}
        </div>

        {/* 按鈕 */}
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          <button onClick={()=>doExport(rangedRecords,rangeLabel)} disabled={exporting||rangedRecords.length===0} style={{
            width:"100%",background:"linear-gradient(135deg,#1a5c3a,#2d8a6e)",border:"none",
            color:"#fff",borderRadius:12,padding:"13px 0",cursor:"pointer",fontSize:14,fontWeight:700,
            opacity:rangedRecords.length===0?0.4:1}}>
            ⬇ 匯出篩選期間（{rangedRecords.length} 筆）
          </button>
          <button onClick={()=>doExport(records,"全部")} disabled={exporting||records.length===0} style={{
            width:"100%",background:"linear-gradient(135deg,#8b2500,#c8560a)",border:"none",
            color:"#fff",borderRadius:12,padding:"13px 0",cursor:"pointer",fontSize:14,fontWeight:700,
            opacity:records.length===0?0.4:1}}>
            ⬇ 匯出全部（{records.length} 筆）
          </button>
          <button onClick={onClose} style={{width:"100%",background:"transparent",
            border:"1px solid rgba(255,255,255,0.12)",color:"#888",borderRadius:12,
            padding:"11px 0",cursor:"pointer",fontSize:14}}>取消</button>
        </div>
      </div>
    </div>
  );
}

function MonthQuickPicker({year,month,onChange}){
  const months=[];
  for(let i=0;i<6;i++){let m=nowM-i,y=nowY;while(m<=0){m+=12;y--;}months.push({y,m});}
  return (
    <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:10}}>
      {months.map(({y,m})=>{
        const active=year===y&&month===m;
        return <button key={`${y}-${m}`} onClick={()=>onChange(y,m)} style={{
          background:active?"rgba(200,134,10,0.35)":"#2a1500",
          border:`1px solid ${active?"#c8860a":"rgba(200,134,10,0.2)"}`,
          color:active?"#f5e6c8":"#888",borderRadius:20,padding:"4px 12px",fontSize:12,cursor:"pointer",
        }}>{y===nowY?"":y+"/"}{m}月</button>;
      })}
    </div>
  );
}

function Spinner(){
  return (
    <div style={{minHeight:"100vh",background:"#1a0e05",display:"flex",flexDirection:"column",
      alignItems:"center",justifyContent:"center",gap:16}}>
      <div style={{fontSize:36}}>🍜</div>
      <div style={{fontSize:15,color:"#c8860a",fontWeight:600,letterSpacing:2}}>mega 記帳</div>
      <div style={{color:"#888",fontSize:13}}>載入共用帳本中…</div>
      <div style={{width:40,height:40,border:"3px solid rgba(200,134,10,0.2)",
        borderTop:"3px solid #c8860a",borderRadius:"50%",animation:"spin .8s linear infinite",marginTop:8}}/>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

export default function App(){
  const [loading,  setLoading]  = useState(true);
  const [syncing,  setSyncing]  = useState(false);
  const [lastSync, setLastSync] = useState(null);
  const [users,    setUsers]    = useState(DEFAULT_USERS);
  const [records,  setRecords]  = useState([]);
  const [currentUser, setCurrentUser] = useState(DEFAULT_USERS[0]);
  const [view, setView] = useState("dashboard");
  const [form, setForm] = useState({type:"income",amount:"",category:CATEGORIES_IN[0],note:"",date:today(),paid:true});
  const [rangeMode,   setRangeMode]   = useState("month");
  const [selYear,     setSelYear]     = useState(nowY);
  const [selMonth,    setSelMonth]    = useState(nowM);
  const [customStart, setCustomStart] = useState(monthStart(nowY,nowM));
  const [customEnd,   setCustomEnd]   = useState(today());
  const [filterUser,     setFilterUser]     = useState("全部");
  const [filterType,     setFilterType]     = useState("全部");
  const [filterPaid,     setFilterPaid]     = useState("全部");
  const [filterCategory, setFilterCategory] = useState("全部類別");
  const [toast,       setToast]       = useState(null);
  const [deleteConfirm,setDeleteConfirm]=useState(null);
  const [showMembers, setShowMembers] = useState(false);
  const [showExport,  setShowExport]  = useState(false);

  const getUserColor=(u)=>USER_COLORS[users.indexOf(u)]||"#888";

  useEffect(()=>{
    async function init(){
      const unsubR = onValue(ref(db,"records"), snap=>{
        const val=snap.val();
        setRecords(val?Object.values(val):[]);
        setLastSync(new Date());
        setLoading(false);
      });
      const unsubU = onValue(ref(db,"users"), snap=>{
        const val=snap.val(); if(val){setUsers(val);setCurrentUser(val[0]);}
      });
      return ()=>{ unsubR(); unsubU(); };
    }
    init();
  },[]);



  useEffect(()=>{ if(toast){const t=setTimeout(()=>setToast(null),2500);return()=>clearTimeout(t);} },[toast]);
  useEffect(()=>{ if(!users.includes(currentUser)) setCurrentUser(users[0]); },[users]);

  const saveRecords=useCallback(async(recs)=>{
    setSyncing(true);
    const obj={}; recs.forEach(r=>{obj[r.id]=r;});
    await set(ref(db,"records"),obj);
    setLastSync(new Date());
    setSyncing(false);
  },[]);
  const saveUsers=useCallback(async(u)=>{ await set(ref(db,"users"),u); },[]);

  const rangeStart=rangeMode==="month"?monthStart(selYear,selMonth):rangeMode==="custom"?customStart:"0000-01-01";
  const rangeEnd  =rangeMode==="month"?monthEnd(selYear,selMonth)  :rangeMode==="custom"?customEnd  :"9999-12-31";
  const inRange=r=>r.date>=rangeStart&&r.date<=rangeEnd;

  const rangedRecords=records.filter(inRange);
  const totalIncome  =rangedRecords.filter(r=>r.type==="income").reduce((s,r)=>s+r.amount,0);
  const totalExpense =rangedRecords.filter(r=>r.type==="expense").reduce((s,r)=>s+r.amount,0);
  const unpaidTotal  =rangedRecords.filter(r=>r.type==="expense"&&r.paid===false).reduce((s,r)=>s+r.amount,0);
  const profit=totalIncome-totalExpense;
  const todayRecs   =records.filter(r=>r.date===today());
  const todayIncome =todayRecs.filter(r=>r.type==="income").reduce((s,r)=>s+r.amount,0);
  const todayExpense=todayRecs.filter(r=>r.type==="expense").reduce((s,r)=>s+r.amount,0);

  const filteredRecords=records.filter(inRange)
    .filter(r=>filterUser==="全部"||r.user===filterUser)
    .filter(r=>filterType==="全部"||r.type===filterType)
    .filter(r=>filterCategory==="全部類別"||r.category===filterCategory)
    .filter(r=>{
      if(filterPaid==="全部") return true;
      if(filterPaid==="unpaid") return r.type==="expense"&&r.paid===false;
      if(filterPaid==="paid")   return r.type==="income"||r.paid!==false;
      return true;
    })
    .sort((a,b)=>b.date.localeCompare(a.date)||b.id.localeCompare(a.id));

  function handleFormChange(field,value){
    setForm(prev=>{
      const next={...prev,[field]:value};
      if(field==="type"){next.category=value==="income"?CATEGORIES_IN[0]:CATEGORIES_OUT[0];next.paid=true;}
      return next;
    });
  }

  async function handleAdd(){
    if(!form.amount||isNaN(Number(form.amount))||Number(form.amount)<=0){
      setToast({msg:"請輸入有效金額",ok:false}); return;
    }
    const rec={id:generateId(),...form,amount:Number(form.amount),user:currentUser};
    if(rec.type==="income") rec.paid=true;
    const next=[rec,...records];
    setRecords(next);
    setForm({type:"income",amount:"",category:CATEGORIES_IN[0],note:"",date:today(),paid:true});
    setToast({msg:"記錄已儲存，同步中…",ok:true});
    setView("dashboard");
    await saveRecords(next);
    setToast({msg:"✓ 已同步至雲端",ok:true});
  }

  async function handleDelete(id){
    const next=records.filter(r=>r.id!==id);
    setRecords(next); setDeleteConfirm(null);
    setToast({msg:"已刪除",ok:false});
    await saveRecords(next);
  }

  async function handleTogglePaid(id){
    const next=records.map(r=>r.id===id?{...r,paid:!r.paid}:r);
    setRecords(next);
    await saveRecords(next);
  }

  async function handleSaveMembers(newUsers){
    const next=records.map(r=>{const idx=users.indexOf(r.user);return idx>=0?{...r,user:newUsers[idx]}:r;});
    setRecords(next); setUsers(newUsers); setShowMembers(false);
    setToast({msg:"成員名稱已更新 ✓",ok:true});
    await Promise.all([saveRecords(next),saveUsers(newUsers)]);
  }

  function shiftMonth(delta){
    let m=selMonth+delta,y=selYear;
    if(m>12){m=1;y++;} if(m<1){m=12;y--;}
    setSelYear(y); setSelMonth(m);
  }

  const rangeLabel=rangeMode==="month"?`${selYear}年${selMonth}月`
    :rangeMode==="custom"?`${customStart}～${customEnd}`:"全部時間";

  const catBreakdown=(type)=>{
    const cats=type==="income"?CATEGORIES_IN:CATEGORIES_OUT;
    return cats.map(c=>({c,sum:rangedRecords.filter(r=>r.type===type&&r.category===c).reduce((s,r)=>s+r.amount,0)}))
      .filter(x=>x.sum>0);
  };

  const RangeSelector=({compact=false})=>(
    <div style={{background:"#2a1500",border:"1px solid rgba(200,134,10,0.2)",
      borderRadius:14,padding:compact?"10px 12px":"14px 16px",marginBottom:12}}>
      <div style={{display:"flex",gap:6,marginBottom:10}}>
        {[["month","月份"],["custom","自訂"],["all","全部"]].map(([m,l])=>(
          <button key={m} onClick={()=>setRangeMode(m)} style={{
            flex:1,background:rangeMode===m?"rgba(200,134,10,0.3)":"transparent",
            border:`1px solid ${rangeMode===m?"#c8860a":"rgba(200,134,10,0.2)"}`,
            color:rangeMode===m?"#f5e6c8":"#888",borderRadius:8,
            padding:`${compact?5:6}px 0`,fontSize:12,cursor:"pointer"}}>
            {l}
          </button>
        ))}
      </div>
      {rangeMode==="month"&&<>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:compact?0:10}}>
          <button onClick={()=>shiftMonth(-1)} style={{background:"rgba(200,134,10,0.15)",border:"none",
            color:"#e09020",borderRadius:8,width:30,height:30,cursor:"pointer",fontSize:16}}>‹</button>
          <div style={{flex:1,textAlign:"center",fontWeight:700,fontSize:compact?14:15}}>{rangeLabel}</div>
          <button onClick={()=>shiftMonth(1)} style={{background:"rgba(200,134,10,0.15)",border:"none",
            color:"#e09020",borderRadius:8,width:30,height:30,cursor:"pointer",fontSize:16}}>›</button>
        </div>
        {!compact&&<MonthQuickPicker year={selYear} month={selMonth} onChange={(y,m)=>{setSelYear(y);setSelMonth(m);}}/>}
      </>}
      {rangeMode==="custom"&&(
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <input type="date" value={customStart} onChange={e=>setCustomStart(e.target.value)} style={{
            flex:1,background:"#1a0e05",border:"1px solid rgba(200,134,10,0.3)",
            borderRadius:8,padding:"5px 6px",color:"#f5e6c8",fontSize:11,outline:"none"}}/>
          <span style={{opacity:0.5,fontSize:11}}>至</span>
          <input type="date" value={customEnd} onChange={e=>setCustomEnd(e.target.value)} style={{
            flex:1,background:"#1a0e05",border:"1px solid rgba(200,134,10,0.3)",
            borderRadius:8,padding:"5px 6px",color:"#f5e6c8",fontSize:11,outline:"none"}}/>
        </div>
      )}
      {rangeMode==="all"&&<div style={{opacity:0.5,fontSize:12,textAlign:"center"}}>顯示所有時間的記錄</div>}
    </div>
  );

  if(loading) return <Spinner/>;

  return (
    <div style={{minHeight:"100vh",background:"#1a0e05",backgroundImage:BATIK_PATTERN,
      fontFamily:"'Segoe UI','Noto Sans',sans-serif",color:"#f5e6c8",
      display:"flex",flexDirection:"column",maxWidth:480,margin:"0 auto",position:"relative"}}>

      {/* Header */}
      <div style={{background:"linear-gradient(135deg,#8b2500 0%,#c8560a 60%,#e09020 100%)",
        padding:"14px 14px 12px",boxShadow:"0 4px 20px rgba(0,0,0,0.5)",
        position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{fontSize:18,fontWeight:800,letterSpacing:1}}>mega 記帳</div>
            <div style={{fontSize:10,opacity:0.7,marginTop:1,display:"flex",alignItems:"center",gap:4}}>
              {syncing
                ?<><span style={{width:6,height:6,borderRadius:"50%",background:"#ffb400",display:"inline-block",animation:"pulse 1s infinite"}}/>同步中…</>
                :<><span style={{width:6,height:6,borderRadius:"50%",background:"#5dde8a",display:"inline-block"}}/>
                  {lastSync?"已同步 "+lastSync.toLocaleTimeString("zh-TW",{hour:"2-digit",minute:"2-digit"}):"雲端共用帳本"}</>
              }
            </div>
          </div>
          <div style={{display:"flex",gap:6,alignItems:"center"}}>
            <button onClick={()=>setShowExport(true)} style={{
              background:"rgba(0,0,0,0.25)",border:"1px solid rgba(255,255,255,0.2)",
              color:"#fff",borderRadius:16,padding:"5px 10px",fontSize:11,cursor:"pointer",
              display:"flex",alignItems:"center",gap:4}}>
              📊<span>報表</span>
            </button>
            <button onClick={()=>setShowMembers(true)} style={{
              background:"rgba(0,0,0,0.25)",border:"1px solid rgba(255,255,255,0.2)",
              color:"#fff",borderRadius:16,padding:"5px 10px",fontSize:11,cursor:"pointer"}}>✏️ 成員</button>
            <select value={currentUser} onChange={e=>setCurrentUser(e.target.value)} style={{
              background:"rgba(0,0,0,0.3)",color:"#fff",border:"1px solid rgba(255,255,255,0.3)",
              borderRadius:20,padding:"5px 8px",fontSize:12,cursor:"pointer"}}>
              {users.map(u=><option key={u} value={u} style={{background:"#2a1200"}}>{u}</option>)}
            </select>
          </div>
        </div>
      </div>

      {/* Toast */}
      {toast&&<div style={{position:"fixed",top:68,left:"50%",transform:"translateX(-50%)",
        background:toast.ok?"#1e5c3a":"#5c3a00",color:"#fff",borderRadius:24,
        padding:"8px 22px",fontSize:13,boxShadow:"0 4px 16px rgba(0,0,0,0.4)",
        zIndex:999,animation:"fadeIn .2s ease",whiteSpace:"nowrap"}}>{toast.msg}</div>}

      <div style={{flex:1,overflowY:"auto",padding:"0 0 80px"}}>

        {/* ── 首頁 ── */}
        {view==="dashboard"&&<div><div style={{padding:"16px 16px 0"}}>
          <RangeSelector/>

          {/* 主摘要 */}
          <div style={{background:"linear-gradient(135deg,#2d1a00,#3d2500)",
            border:"1px solid rgba(200,134,10,0.25)",borderRadius:16,
            padding:"16px 20px",marginBottom:12,boxShadow:"0 2px 12px rgba(0,0,0,0.4)"}}>
            <div style={{fontSize:11,opacity:0.6,letterSpacing:2,marginBottom:4}}>{rangeLabel}　淨利潤</div>
            <div style={{fontSize:28,fontWeight:800,color:profit>=0?"#5dde8a":"#ff7070"}}>
              {profit>=0?"+":"-"}{formatCurrency(profit)}
            </div>
            <div style={{display:"flex",gap:16,marginTop:10,flexWrap:"wrap"}}>
              <div><div style={{fontSize:10,opacity:0.5,letterSpacing:1}}>期間收入</div>
                <div style={{fontSize:15,color:"#5dde8a",fontWeight:600}}>{formatCurrency(totalIncome)}</div></div>
              <div><div style={{fontSize:10,opacity:0.5,letterSpacing:1}}>期間支出</div>
                <div style={{fontSize:15,color:"#ff8080",fontWeight:600}}>{formatCurrency(totalExpense)}</div></div>
              {unpaidTotal>0&&<div><div style={{fontSize:10,opacity:0.5,letterSpacing:1}}>未付款</div>
                <div style={{fontSize:15,color:"#ffb400",fontWeight:600}}>{formatCurrency(unpaidTotal)}</div></div>}
              <div><div style={{fontSize:10,opacity:0.5,letterSpacing:1}}>筆數</div>
                <div style={{fontSize:15,color:"#e09020",fontWeight:600}}>{rangedRecords.length}</div></div>
            </div>
          </div>

          {/* 今日 */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
            {[{label:"今日收入",val:todayIncome,color:"#5dde8a",icon:"📈"},
              {label:"今日支出",val:todayExpense,color:"#ff8080",icon:"📉"}].map(({label,val,color,icon})=>(
              <div key={label} style={{background:"#2a1500",border:"1px solid rgba(200,134,10,0.2)",borderRadius:12,padding:"12px 14px"}}>
                <div style={{fontSize:11,opacity:0.5,marginBottom:4}}>{icon} {label}</div>
                <div style={{fontSize:17,fontWeight:700,color}}>{formatCurrency(val)}</div>
              </div>
            ))}
          </div>

          {/* 類別分解 */}
          {(catBreakdown("income").length>0||catBreakdown("expense").length>0)&&(
            <div style={{background:"#2a1500",border:"1px solid rgba(200,134,10,0.15)",
              borderRadius:12,padding:"14px 16px",marginBottom:12}}>
              <div style={{fontSize:11,opacity:0.5,letterSpacing:2,marginBottom:10}}>收支明細分類</div>
              {catBreakdown("income").map(({c,sum})=><CategoryBar key={c} label={c} amount={sum} max={totalIncome} color="#5dde8a"/>)}
              {catBreakdown("expense").map(({c,sum})=><CategoryBar key={c} label={c} amount={sum} max={totalExpense} color="#ff8080"/>)}
            </div>
          )}

          {/* 人員 */}
          <div style={{background:"#2a1500",borderRadius:12,padding:"14px 16px",
            border:"1px solid rgba(200,134,10,0.15)",marginBottom:12}}>
            <div style={{fontSize:11,opacity:0.5,letterSpacing:2,marginBottom:10}}>各人員（期間）</div>
            {users.map(u=>{
              const uRecs=rangedRecords.filter(r=>r.user===u);
              if(!uRecs.length) return null;
              const uIn=uRecs.filter(r=>r.type==="income").reduce((s,r)=>s+r.amount,0);
              const uOut=uRecs.filter(r=>r.type==="expense").reduce((s,r)=>s+r.amount,0);
              const uUnpaid=uRecs.filter(r=>r.type==="expense"&&r.paid===false).reduce((s,r)=>s+r.amount,0);
              return (
                <div key={u} style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
                  <div style={{width:28,height:28,borderRadius:"50%",background:getUserColor(u),
                    display:"flex",alignItems:"center",justifyContent:"center",
                    fontSize:12,fontWeight:700,flexShrink:0,color:"#fff"}}>{u[0]}</div>
                  <div style={{flex:1}}>
                    <div style={{fontSize:13,fontWeight:600}}>{u}</div>
                    <div style={{fontSize:11,opacity:0.5}}>{uRecs.length} 筆</div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    {uIn>0&&<div style={{fontSize:12,color:"#5dde8a"}}>+{formatCurrency(uIn)}</div>}
                    {uOut>0&&<div style={{fontSize:12,color:"#ff8080"}}>-{formatCurrency(uOut)}</div>}
                    {uUnpaid>0&&<div style={{fontSize:11,color:"#ffb400"}}>未付 {formatCurrency(uUnpaid)}</div>}
                  </div>
                </div>
              );
            })}
            {rangedRecords.length===0&&<div style={{opacity:0.4,fontSize:13}}>此區間尚無記錄</div>}
          </div>

          <div style={{fontSize:12,opacity:0.5,letterSpacing:2,marginBottom:8}}>最近記錄</div>
          {records.length===0
            ?<div style={{opacity:0.4,fontSize:13,textAlign:"center",padding:24}}>尚無記錄，點擊 + 新增</div>
            :records.slice(0,5).map(r=><RecordCard key={r.id} r={r} getUserColor={getUserColor}
                onDelete={setDeleteConfirm} onTogglePaid={handleTogglePaid}/>)
          }
        </div></div>}

        {/* ── 新增 ── */}
        {view==="add"&&<div style={{padding:16}}>
          <div style={{fontSize:16,fontWeight:700,marginBottom:16,opacity:0.8}}>
            新增記錄 — <span style={{color:getUserColor(currentUser)}}>{currentUser}</span>
          </div>
          <div style={{display:"flex",gap:8,marginBottom:14}}>
            {["income","expense"].map(t=>(
              <button key={t} onClick={()=>handleFormChange("type",t)} style={{
                flex:1,padding:"10px 0",borderRadius:10,border:"none",cursor:"pointer",
                fontWeight:700,fontSize:14,transition:"all .2s",
                background:form.type===t?(t==="income"?"#1e5c3a":"#6b2020"):"#2a1500",
                color:form.type===t?"#fff":"#888",
                boxShadow:form.type===t?"0 2px 10px rgba(0,0,0,0.4)":"none"}}>
                {t==="income"?"📈 收入":"📉 支出"}
              </button>
            ))}
          </div>
          <FormField label="金額（新台幣 NT$）">
            <input type="number" value={form.amount} placeholder="例：1500"
              onChange={e=>handleFormChange("amount",e.target.value)} style={inputStyle}/>
          </FormField>
          <FormField label="類別">
            <select value={form.category} onChange={e=>handleFormChange("category",e.target.value)} style={inputStyle}>
              {(form.type==="income"?CATEGORIES_IN:CATEGORIES_OUT).map(c=>(
                <option key={c} value={c} style={{background:"#2a1200"}}>{c}</option>
              ))}
            </select>
          </FormField>
          {form.type==="expense"&&(
            <FormField label="付款狀態">
              <div style={{display:"flex",gap:8}}>
                {[[true,"✓ 已付款"],[false,"⏳ 未付款"]].map(([val,label])=>(
                  <button key={String(val)} onClick={()=>handleFormChange("paid",val)} style={{
                    flex:1,padding:"9px 0",borderRadius:10,border:"none",cursor:"pointer",
                    fontWeight:700,fontSize:13,transition:"all .2s",
                    background:form.paid===val?(val?"rgba(93,222,138,0.25)":"rgba(255,180,0,0.25)"):"#2a1500",
                    color:form.paid===val?(val?"#5dde8a":"#ffb400"):"#666",
                    boxShadow:form.paid===val?"0 2px 8px rgba(0,0,0,0.3)":"none"}}>
                    {label}
                  </button>
                ))}
              </div>
            </FormField>
          )}
          <FormField label="備註（選填）">
            <input value={form.note} placeholder="可留空..." onChange={e=>handleFormChange("note",e.target.value)} style={inputStyle}/>
          </FormField>
          <FormField label="日期">
            <input type="date" value={form.date} onChange={e=>handleFormChange("date",e.target.value)} style={inputStyle}/>
          </FormField>
          <div style={{height:8}}/>
          <button onClick={handleAdd} style={{width:"100%",background:"linear-gradient(135deg,#8b2500,#c8560a)",
            color:"#fff",border:"none",borderRadius:12,padding:"14px 0",fontSize:16,fontWeight:700,
            cursor:"pointer",letterSpacing:1,boxShadow:"0 4px 16px rgba(200,86,10,0.4)"}}>確認儲存</button>
        </div>}

        {/* ── 明細 ── */}
        {view==="history"&&<div style={{padding:16}}>
          <RangeSelector compact/>
          <div style={{display:"flex",gap:6,marginBottom:10,flexWrap:"wrap",alignItems:"center"}}>
            <select value={filterUser} onChange={e=>setFilterUser(e.target.value)} style={{
              background:"#2a1500",border:"1px solid rgba(200,134,10,0.3)",color:"#f5e6c8",
              borderRadius:8,padding:"6px 8px",fontSize:12,outline:"none"}}>
              {["全部",...users].map(u=><option key={u} value={u} style={{background:"#2a1200"}}>{u}</option>)}
            </select>
            <select value={filterType} onChange={e=>{ setFilterType(e.target.value); setFilterCategory("全部類別"); }} style={{
              background:"#2a1500",border:"1px solid rgba(200,134,10,0.3)",color:"#f5e6c8",
              borderRadius:8,padding:"6px 8px",fontSize:12,outline:"none"}}>
              {[["全部","全部"],["income","收入"],["expense","支出"]].map(([v,l])=>(
                <option key={v} value={v} style={{background:"#2a1200"}}>{l}</option>
              ))}
            </select>
            {/* 類別篩選 */}
            <select value={filterCategory} onChange={e=>setFilterCategory(e.target.value)} style={{
              background:"#2a1500",border:"1px solid rgba(200,134,10,0.3)",color:"#f5e6c8",
              borderRadius:8,padding:"6px 8px",fontSize:12,outline:"none"}}>
              <option value="全部類別" style={{background:"#2a1200"}}>全部類別</option>
              {filterType!=="expense" && CATEGORIES_IN.map(c=>(
                <option key={c} value={c} style={{background:"#2a1200"}}>📈 {c}</option>
              ))}
              {filterType!=="income" && CATEGORIES_OUT.map(c=>(
                <option key={c} value={c} style={{background:"#2a1200"}}>📉 {c}</option>
              ))}
            </select>
            <select value={filterPaid} onChange={e=>setFilterPaid(e.target.value)} style={{
              background:"#2a1500",border:"1px solid rgba(200,134,10,0.3)",color:"#f5e6c8",
              borderRadius:8,padding:"6px 8px",fontSize:12,outline:"none"}}>
              {[["全部","全部付款"],["paid","已付款"],["unpaid","未付款"]].map(([v,l])=>(
                <option key={v} value={v} style={{background:"#2a1200"}}>{l}</option>
              ))}
            </select>
            <div style={{width:"100%",display:"flex",justifyContent:"flex-end",marginTop:-4}}>
              <span style={{opacity:0.5,fontSize:11}}>共 {filteredRecords.length} 筆</span>
            </div>
          </div>
          {/* 類別小計卡 — 當選了特定類別時顯示 */}
          {filterCategory!=="全部類別"&&filteredRecords.length>0&&(()=>{
            const catTotal=filteredRecords.reduce((s,r)=>s+r.amount,0);
            const isIncomeCat=CATEGORIES_IN.includes(filterCategory);
            return (
              <div style={{background:isIncomeCat?"rgba(93,222,138,0.08)":"rgba(255,128,128,0.08)",
                border:`1px solid ${isIncomeCat?"rgba(93,222,138,0.3)":"rgba(255,128,128,0.3)"}`,
                borderRadius:12,padding:"12px 16px",marginBottom:12,display:"flex",alignItems:"center",gap:12}}>
                <div style={{fontSize:22}}>{isIncomeCat?"📈":"📉"}</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:12,opacity:0.6,marginBottom:2}}>{filterCategory} 合計</div>
                  <div style={{fontSize:20,fontWeight:800,color:isIncomeCat?"#5dde8a":"#ff8080"}}>
                    {isIncomeCat?"+":"-"}{formatCurrency(catTotal)}
                  </div>
                </div>
                <div style={{textAlign:"right",opacity:0.6,fontSize:12}}>{filteredRecords.length} 筆</div>
              </div>
            );
          })()}
          {filteredRecords.length>0&&(
            <div style={{display:"flex",gap:8,marginBottom:12,background:"#2a1500",
              borderRadius:10,padding:"10px 14px",border:"1px solid rgba(200,134,10,0.15)"}}>
              {[
                {label:"收入合計",val:filteredRecords.filter(r=>r.type==="income").reduce((s,r)=>s+r.amount,0),color:"#5dde8a"},
                {label:"支出合計",val:filteredRecords.filter(r=>r.type==="expense").reduce((s,r)=>s+r.amount,0),color:"#ff8080"},
                {label:"未付款",val:filteredRecords.filter(r=>r.type==="expense"&&r.paid===false).reduce((s,r)=>s+r.amount,0),color:"#ffb400"},
              ].map(({label,val,color})=>(
                <div key={label} style={{flex:1}}>
                  <div style={{fontSize:10,opacity:0.5}}>{label}</div>
                  <div style={{fontSize:13,color,fontWeight:700}}>{formatCurrency(val)}</div>
                </div>
              ))}
            </div>
          )}
          {filteredRecords.length===0
            ?<div style={{opacity:0.4,fontSize:13,textAlign:"center",padding:40}}>此區間沒有符合的記錄</div>
            :filteredRecords.map(r=><RecordCard key={r.id} r={r} getUserColor={getUserColor}
                onDelete={setDeleteConfirm} onTogglePaid={handleTogglePaid}/>)
          }
        </div>}
      </div>

      {/* 刪除確認 */}
      {deleteConfirm&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.7)",
          display:"flex",alignItems:"center",justifyContent:"center",zIndex:200,padding:20}}>
          <div style={{background:"#2a1500",borderRadius:16,padding:24,width:"100%",maxWidth:320,
            border:"1px solid rgba(200,134,10,0.3)"}}>
            <div style={{fontSize:16,fontWeight:700,marginBottom:8}}>確認刪除？</div>
            <div style={{opacity:0.6,fontSize:13,marginBottom:20}}>此操作無法復原</div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setDeleteConfirm(null)} style={{flex:1,background:"#3a2500",
                border:"1px solid rgba(255,255,255,0.1)",color:"#ccc",borderRadius:10,
                padding:"10px 0",cursor:"pointer",fontSize:14}}>取消</button>
              <button onClick={()=>handleDelete(deleteConfirm)} style={{flex:1,background:"#6b2020",
                border:"none",color:"#fff",borderRadius:10,padding:"10px 0",
                cursor:"pointer",fontSize:14,fontWeight:700}}>刪除</button>
            </div>
          </div>
        </div>
      )}

      {showMembers&&<MembersModal users={users} onSave={handleSaveMembers} onClose={()=>setShowMembers(false)}/>}
      {showExport&&<ExportModal records={records} rangedRecords={rangedRecords} users={users}
        rangeLabel={rangeLabel} onClose={()=>setShowExport(false)} onToast={setToast}/>}

      {/* 底部導航 */}
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",
        width:"100%",maxWidth:480,
        background:"linear-gradient(180deg,rgba(26,14,5,0) 0%,#1a0e05 20%)",
        padding:"12px 20px 20px",display:"flex",justifyContent:"space-around",alignItems:"center",zIndex:100}}>
        {[{id:"dashboard",icon:"🏠",label:"首頁"},{id:"add",icon:"＋",label:"新增",big:true},{id:"history",icon:"📋",label:"明細"}]
          .map(({id,icon,label,big})=>(
          <button key={id} onClick={()=>setView(id)} style={{
            background:big?"linear-gradient(135deg,#8b2500,#c8560a)":view===id?"rgba(200,134,10,0.2)":"transparent",
            border:"none",color:big?"#fff":view===id?"#e09020":"#888",
            borderRadius:big?"50%":12,width:big?56:60,height:big?56:44,
            display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
            cursor:"pointer",boxShadow:big?"0 4px 20px rgba(200,86,10,0.5)":"none",
            transition:"all .2s",marginTop:big?-12:0}}>
            <span style={{fontSize:big?24:18}}>{icon}</span>
            {!big&&<span style={{fontSize:10,marginTop:2}}>{label}</span>}
          </button>
        ))}
      </div>

      <style>{`
        @keyframes fadeIn{from{opacity:0;transform:translateX(-50%) translateY(-8px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}
        @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.4}}
        input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none}
        select option{color:#f5e6c8}
        ::-webkit-scrollbar{width:4px}::-webkit-scrollbar-track{background:transparent}
        ::-webkit-scrollbar-thumb{background:rgba(200,134,10,0.3);border-radius:4px}
        input[type=date]::-webkit-calendar-picker-indicator{filter:invert(0.7)}
      `}</style>
    </div>
  );
}
