// app.js — 대시보드 조립, 상세 모달, 전체 목록, 캘린더, 이메일, 단톡공지, 계산기, 앱 초기화

import {FSCache, LS, TODAY, addLog, buildMailingData, curSubTab, ddayChip, diff, findEmail, fmt, fsscRaw, getChecks, getEmails, getIgcStore, getLogs, getMailingLog, getOverrides, getSchedules, hydrateMailingRawData, hydrateRawData, isOverseasCountry, jobRaw, loadMailingFssc, loadMailingJob, mailingFsscName, mailingFsscRaw, mailingJobName, mailingJobRaw, mergeAndRender, merged, normalizeStatus, parseDate, readXlsx, renderMiniCalHtml, renderScheduleHtml, saveChecks, saveEmails, saveOverrides, saveSchedules, setCurSubTab, setMailingField, setMailingNotice, setShowHidden, setTxt, showHidden, statusChip} from './core.js';
import {renderMonthRows, renderReqRows, toggleMonthSort} from './tab-expiry.js';
import {renderSuspendRows, toggleUrgentDdayOnly, toggleUrgentSort} from './tab-urgent.js';
import {isSamePartner, renderRevokeRows, selectRevokeFilter, toggleRevokeDropdown} from './tab-revoke.js';

export let calYear=TODAY.getFullYear(),calMonth=TODAY.getMonth();

export function loadEmailFile(ev){
  const f=ev.target.files[0];if(!f)return;
  readXlsx(f,wb=>{
    const ws=wb.Sheets[wb.SheetNames[0]];
    const arr=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    const emails=getEmails();let added=0;
    // 헤더 행이 있는지 확인 (첫 행에 이메일이 없으면 헤더로 간주)
    const firstRow=arr[0]||[];
    const hasHeader=firstRow.some(c=>String(c).includes('@')===false&&String(c).trim().length>0);
    let hdrIdx={name:0,email:1,phone:-1};
    if(hasHeader){
      firstRow.forEach((h,i)=>{
        const s=String(h).trim().toLowerCase();
        if(s.includes('이름')||s.includes('컨설팅'))hdrIdx.name=i;
        if(s.includes('이메일')||s.includes('email')||s.includes('mail'))hdrIdx.email=i;
        if(s.includes('전화')||s.includes('연락처')||s.includes('번호')||s.includes('phone')||s.includes('tel'))hdrIdx.phone=i;
      });
    }
    const dataRows=hasHeader?arr.slice(1):arr;
    dataRows.forEach(row=>{
      const name=String(row[hdrIdx.name]||'').trim();
      const email=String(row[hdrIdx.email]||'').trim();
      const phone=hdrIdx.phone>=0?String(row[hdrIdx.phone]||'').trim():'';
      if(name&&email&&email.includes('@')){
        if(!emails[name])emails[name]={email,phone,aliases:[]};
        else{emails[name].email=email;if(phone)emails[name].phone=phone;}
        added++;
      }
    });
    saveEmails(emails);renderEmailTab();showMsg(`${added}개 저장됐어요.`);
  });ev.target.value='';
}


// ── 서브탭 ──

export function goSubTab(tab){
  setCurSubTab(tab);
  document.getElementById('stab-fssc').classList.toggle('on',tab==='fssc');
  document.getElementById('stab-fsms').classList.toggle('on',tab==='fsms');
  renderDashContent();
}


// ── DASHBOARD ──

export function toggleHidden(){setShowHidden(!showHidden);renderDashContent();}


export function renderDash(){renderDashContent();}


export function renderDashContent(){
  try{
    const src=curSubTab==='fssc'?'FSSC':'IGC';
    const data=merged.filter(r=>r.src===src);
    const cm=TODAY.getMonth(),cy=TODAY.getFullYear();
    const month=[],monthHidden=[],suspend=[],revoke=[],urgent=[];
    const checks=getChecks();
    const rangeStart=new Date(TODAY.getFullYear(),TODAY.getMonth()-1,1);
    const rangeEnd=new Date(TODAY.getFullYear(),TODAY.getMonth()+3,0);
    data.forEach(r=>{
      const d=diff(r.expDate);
      const st=r.status;
      const c=checks[r.jobNo]||{};
      const isActive=(st==='정상'||st==='유지'||st==='진행중'||st==='');
      if(r.expDate&&isActive){
        const baseDate=(r.src==='FSSC'&&r.initialCert)?r.initialCert:r.expDate;
        const bm=baseDate.getMonth();
        const inRange=[cm-2,cm-1,cm,cm+1,cm+2].map(m=>(m+12)%12).includes(bm);
        const d2=diff(r.expDate);
        const isExpired=d2!==null&&d2<0&&!c['monthUpload'];
        if(inRange&&!isExpired&&!c['monthUpload']){
          if(c['intent1'])monthHidden.push(r);else month.push(r);
        }
      }
      // URGENT: 만료일 D-10~D+5 또는 정지일로부터 170일 이상 경과
      const suspendElapsed=r.suspendDate?Math.round((TODAY-r.suspendDate)/86400000):null;
      const isUrgentExpire=d!==null&&d>=-5&&d<=10;
      const isUrgentSuspend=suspendElapsed!==null&&suspendElapsed>=170;
      if(isUrgentExpire||isUrgentSuspend)suspend.push(r);
      if(st==='정지'||st==='정지(Suspended)'||c.chk3||c.chk4||c.urgSuspend===2){
        if(src==='FSSC'&&((r.fsscStatus||'').toLowerCase()==='suspended'||c.chk3||c.chk4||c.urgSuspend===2))revoke.push(r);
        else if(src==='IGC'&&(st==='정지'||st==='취소'||c.chk3||c.chk4||c.urgSuspend===2))revoke.push(r);
      }
      if(d!==null&&d>=0&&d<=10)urgent.push(r);
    });
    month.sort((a,b)=>{
      if(window._monthSort==='dday'){
        return (diff(a.expDate)??999)-(diff(b.expDate)??999);
      }else{
        // 만료일 정렬 - 현재 월 기준 가까운 순 (월 우선, 그 다음 일)
        const am=(a.expDate?.getMonth()-cm+12)%12;
        const bm=(b.expDate?.getMonth()-cm+12)%12;
        if(am!==bm)return am-bm;
        return (a.expDate?.getDate()||0)-(b.expDate?.getDate()||0);
      }
    });
    monthHidden.sort((a,b)=>(a.expDate||0)-(b.expDate||0));
    revoke.sort((a,b)=>{if(!a.suspendDate&&!b.suspendDate)return 0;if(!a.suspendDate)return 1;if(!b.suspendDate)return -1;return a.suspendDate-b.suspendDate;});
    // 철회 완료 / 업로드 완료(인증 복구)된 건은 목록엔 표시하되 "정지 중" 개수에는 포함하지 않음
    const revokeActiveCount=revoke.filter(r=>{const c=checks[r.jobNo]||{};return!(c.revokeDone||c.revokeUpload);}).length;
    const mN=['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];
    const subTxt=`${mN[(cm-1+12)%12]} ~ ${mN[(cm+2)%12]} 만료 업체`;
    const isFsms=curSubTab==='fsms';

    // 요청 업체 (해당 src만)
    const reqRows=data.filter(r=>{const c=checks[r.jobNo]||{};return c.suspend||c.revoke;});

    // 숨김 버튼 HTML
    const hiddenBtn=monthHidden.length>0
      ?`<button class="tbtn" style="font-size:11px;padding:3px 8px" onclick="toggleHidden()">${showHidden?`🔓 숨김 ${monthHidden.length} (접기)`:`🔒 숨김 ${monthHidden.length}`}</button>`:'';

    const dashEl=document.getElementById('dash-content');
    const scrollWraps=[...dashEl.querySelectorAll('.tbl-wrap')];
    const scrollPositions=scrollWraps.map(w=>w.scrollTop);
    const pageScrollY=window.scrollY;

    dashEl.innerHTML=`
    <div class="dash-grid">
      <div class="col-left">
        <div class="card">
          <div class="card-hd">
            만료 예정 업체 <span class="badge">${month.length}</span>
            ${hiddenBtn}
            <button class="tbtn" style="font-size:11px;padding:3px 8px;margin-left:auto" onclick="toggleMonthSort()" title="정렬">${window._monthSort==='dday'?'D-day순':'만료일순'} ↕</button>
            <button class="tbtn" style="font-size:11px;padding:3px 8px" onclick="goTabById('cal')">📅 달력</button>
          </div>
          <div class="card-sub">${subTxt}</div>
          <div class="tbl-wrap" style="max-height:510px">
            <table>
              <thead><tr><th>컨설팅사</th><th>업체명</th><th>잡넘버</th><th>Initial certification</th><th></th></tr></thead>
              <tbody>${renderMonthRows(month,showHidden?monthHidden:[])}</tbody>
            </table>
          </div>
        </div>
        <div class="card">
          <div class="card-hd">요청 업체 <span class="badge">${reqRows.length}</span></div>
          <div class="card-sub">정지·철회 요청이 체크된 업체</div>
          <div class="tbl-wrap">
            <table>
              <thead><tr><th>컨설팅사</th><th>업체명</th><th>잡넘버</th><th>요청</th></tr></thead>
              <tbody>${renderReqRows(reqRows,checks)}</tbody>
            </table>
          </div>
        </div>
      </div>
      <div class="col-mid">
        <div class="card urg">
          <div class="card-hd">!! URGENT !! <span class="badge r">${suspend.length}</span>
            <button class="tbtn" style="font-size:11px;padding:3px 8px;margin-left:auto;border-color:#ff6b6b;color:#ff6b6b" onclick="toggleUrgentSort()">D-day 정렬 ${window._urgentSort?'↑':'↓'}</button>
            <button class="tbtn" style="font-size:11px;padding:3px 8px;${window._urgentDdayOnly?'background:#ff6b6b;color:#fff;':'border-color:#ff6b6b;color:#ff6b6b'}" onclick="toggleUrgentDdayOnly()" title="D+n(기한 경과)인 업체 숨기기">D-day</button>
          </div>
          <div class="card-sub">만료일 10일 전인 업체 + 정지일로부터 170일 경과한 업체</div>
          <div class="tbl-wrap">
            <table>
              <thead><tr><th>컨설팅사</th><th>업체명</th><th>잡넘버</th><th>만료일 / 정지일</th><th>상태</th></tr></thead>
              <tbody>${renderSuspendRows(suspend)}</tbody>
            </table>
          </div>
        </div>
        <div class="card">
          <div class="card-hd">철회 예정 업체 <span class="badge">${revokeActiveCount}</span>
            <div style="position:relative;display:inline-block;margin-left:6px">
              <button id="revoke-filter-btn" onclick="toggleRevokeDropdown()" style="background:none;border:none;color:var(--muted);cursor:pointer;font-size:12px;padding:2px 4px;display:flex;align-items:center;gap:3px">
                <span id="revoke-filter-label" style="font-size:11px;color:var(--muted)">${window._revokeFilter||'컨설팅사'} ▽</span>
              </button>
              <div id="revoke-filter-list" style="display:none;position:absolute;top:100%;left:0;background:var(--panel);border:1px solid var(--border);border-radius:6px;z-index:50;min-width:160px;box-shadow:0 4px 12px rgba(0,0,0,.3);max-height:200px;overflow-y:auto">
                <div style="padding:6px 10px;cursor:pointer;font-size:12px;border-bottom:1px solid var(--border)" onmouseover="this.style.background='var(--hover)'" onmouseout="this.style.background=''" onclick="selectRevokeFilter('','컨설팅사')">전체</div>
                ${[...new Set([...revoke.map(r=>r.partner),...Object.entries(getChecks()).filter(([k,v])=>v.chk3||v.chk4).map(([k])=>merged.find(r=>r.jobNo===k)?.partner)].filter(Boolean))].sort().map(p=>`<option value="${p}" ${window._revokeFilter===p?';color:var(--accent);font-weight:600':''}" onmouseover="this.style.background='var(--hover)'" onmouseout="this.style.background=''" onclick="selectRevokeFilter('${p.replace(/'/g,"\\'")}','${p.replace(/'/g,"\\'")}')">${p}</option>`).join('')}
              </div>
            </div>
          </div>
          <div class="card-sub">정지 상태 업체 · 정지날짜 오래된 순</div>
          <div class="tbl-wrap">
            <table>
              <thead><tr><th>컨설팅사</th><th>업체명</th><th>잡넘버</th><th>정지날짜</th><th>경과일</th></tr></thead>
              <tbody>${renderRevokeRows(window._revokeFilter?revoke.filter(r=>isSamePartner(r.partner,window._revokeFilter)):revoke)}</tbody>
            </table>
          </div>
        </div>
      </div>
      <div class="col-right">
        ${renderMiniCalHtml()}
        ${renderScheduleHtml()}
      </div>
    </div>`;

    // 재렌더로 초기화된 스크롤 위치 복원 (클릭할 때마다 맨 위로 튀는 문제 방지)
    const newWraps=[...dashEl.querySelectorAll('.tbl-wrap')];
    newWraps.forEach((w,i)=>{ if(scrollPositions[i])w.scrollTop=scrollPositions[i]; });
    if(pageScrollY)window.scrollTo(0,pageScrollY);
  }catch(e){console.error('renderDashContent 오류:',e);}
}


export function addSchedule(){
  const date=document.getElementById('sc-date')?.value;
  const memo=document.getElementById('sc-memo')?.value.trim();
  if(!memo)return;
  const arr=getSchedules();
  arr.push({date,memo});
  saveSchedules(arr);
  renderDashContent();
}

export function delSchedule(i){
  const arr=getSchedules();arr.splice(i,1);saveSchedules(arr);renderDashContent();
}


// 정지처리 / 철회처리 오버라이드

export function applyOverride(jobNo,type,val){
  const overrides=getOverrides();
  if(val){
    overrides[jobNo]=type;
  }else if(overrides[jobNo]===type){
    delete overrides[jobNo];
  }
  saveOverrides(overrides);
  // merged 재반영
  merged.forEach(r=>{
    if(r.jobNo===jobNo){
      const cur=overrides[jobNo];
      if(cur){
        r._override=cur;
        if(cur==='suspended')r.status='정지';
        if(cur==='withdrawn')r.status='취소';
      }else{
        delete r._override;
        // 원래 상태 복원 - FSSC는 업무시트, IGC는 igcStore
        const igcStore=getIgcStore();
        if(igcStore[jobNo])r.status=normalizeStatus(igcStore[jobNo].status);
      }
    }
  });
  renderDashContent();
  renderAll();
}


// ── ALL TABLE ──

export function renderAll(){
  try{
    const q=(document.getElementById('srch').value||'').toLowerCase();
    const st=document.getElementById('flt-status').value;
    const src=document.getElementById('flt-src').value;
    let rows=merged;
    if(q)rows=rows.filter(r=>[r.nameKo,r.nameEn,r.jobNo,r.partner].some(v=>(v||'').toLowerCase().includes(q)));
    if(st)rows=rows.filter(r=>normalizeStatus(r.status)===st||r.status===st);
    if(src)rows=rows.filter(r=>r.src===src);
    setTxt('lbl-cnt',rows.length+'개');
    const tb=document.getElementById('tb-all');
    if(!rows.length){tb.innerHTML='<tr class="empty"><td colspan="16">데이터 없음</td></tr>';return;}
    const checks=getChecks();
    const REQ_REASONS=['컨설팅사 요청','고객사 요청','후속심사 미진행'];
    tb.innerHTML=rows.map(r=>{
      const c=checks[r.jobNo]||{};
      const isOverseas=isOverseasCountry(r.country);
      const srcBadge=r.src==='IGC'?'<span class="chip cb" style="font-size:9px">FSMS</span>':'<span class="chip cgr" style="font-size:9px">FSSC</span>';
      const overseas=isOverseas?'<span class="chip-overseas">해외</span>':'';
      const suspendDrop=c.suspend?`<select style="font-size:10px;padding:2px 4px;border-radius:3px;border:1px solid var(--border);background:var(--panel2);color:var(--text);margin-top:2px" onchange="saveReqReason('${r.jobNo}','suspendReason',this.value)" onclick="event.stopPropagation()"><option value="">사유 선택</option>${REQ_REASONS.map(rs=>`<option ${c.suspendReason===rs?'selected':''}>${rs}</option>`).join('')}</select>`:'';
      const revokeDrop=c.revoke?`<select style="font-size:10px;padding:2px 4px;border-radius:3px;border:1px solid var(--border);background:var(--panel2);color:var(--text);margin-top:2px" onchange="saveReqReason('${r.jobNo}','revokeReason',this.value)" onclick="event.stopPropagation()"><option value="">사유 선택</option>${REQ_REASONS.map(rs=>`<option ${c.revokeReason===rs?'selected':''}>${rs}</option>`).join('')}</select>`:'';
      return`<tr><td>${srcBadge}</td><td>${statusChip(r.status)}</td><td>${overseas}</td><td title="${r.partner}">${r.partner||'-'}</td><td title="${r.nameKo}" class="mv hi" style="cursor:pointer" onclick="openModal(${r.idx})">${r.nameKo||'-'}</td><td title="${r.nameEn}">${r.nameEn||'-'}</td><td>${r.jobNo}</td><td>${r.coid||'-'}</td><td>${fmt(r.issuedDate)}</td><td>${fmt(r.initialCert)}</td><td>${fmt(r.expDate)}</td><td>${ddayChip(r.expDate)}</td><td title="${r.email}">${r.email||'-'}</td><td>${r.phone||'-'}</td>
      <td style="text-align:center" onclick="event.stopPropagation()">
        <div style="display:flex;flex-direction:column;align-items:center">
          <input type="checkbox" ${c.suspend?'checked':''} style="width:15px;height:15px;cursor:pointer;accent-color:#f0a040" onchange="saveReq('${r.jobNo}','suspend',this.checked)">
          ${suspendDrop}
        </div>
      </td>
      <td style="text-align:center" onclick="event.stopPropagation()">
        <div style="display:flex;flex-direction:column;align-items:center">
          <input type="checkbox" ${c.revoke?'checked':''} style="width:15px;height:15px;cursor:pointer;accent-color:#ff6b6b" onchange="saveReq('${r.jobNo}','revoke',this.checked)">
          ${revokeDrop}
        </div>
      </td></tr>`;
    }).join('');
  }catch(e){console.error('renderAll 오류:',e);}
}

export function saveUpdDone(jobNo,val){
  const c=getChecks();if(!c[jobNo])c[jobNo]={};
  c[jobNo].updDone=val;saveChecks(c);renderUpd();
}

export function saveReq(jobNo,field,val){
  const c=getChecks();if(!c[jobNo])c[jobNo]={};
  c[jobNo][field]=val;
  // 체크 해제 시 사유도 초기화
  if(!val){
    if(field==='suspend')delete c[jobNo].suspendReason;
    if(field==='revoke')delete c[jobNo].revokeReason;
  }
  saveChecks(c);renderDashContent();renderAll();
}

export function saveReqReason(jobNo,field,val){
  const c=getChecks();if(!c[jobNo])c[jobNo]={};
  c[jobNo][field]=val;saveChecks(c);renderDashContent();
}


// ── CALENDAR ──

export function calMove(dir){calMonth+=dir;if(calMonth>11){calMonth=0;calYear++;}else if(calMonth<0){calMonth=11;calYear--;}renderCal();}

export function calGoToday(){calYear=TODAY.getFullYear();calMonth=TODAY.getMonth();renderCal();}


export function renderCal(){
  try{
    const showFSSC=document.getElementById('cf-fssc')?.checked!==false;
    const showFSMS=document.getElementById('cf-fsms')?.checked!==false;
    const mN=['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];
    setTxt('cal-title',`${calYear}년 ${mN[calMonth]}`);
    const grid=document.getElementById('cal-grid');if(!grid)return;
    const days=['일','월','화','수','목','금','토'];
    const dCls=['sun','','','','','','sat'];
    let html=days.map((d,i)=>`<div class="cal-head ${dCls[i]}">${d}</div>`).join('');
    const firstDay=new Date(calYear,calMonth,1).getDay();
    const lastDate=new Date(calYear,calMonth+1,0).getDate();
    const prevLast=new Date(calYear,calMonth,0).getDate();
    const dateMap={};
    merged.forEach(r=>{
      if(r.src==='FSSC'&&!showFSSC)return;
      if(r.src==='IGC'&&!showFSMS)return;
      const st=r.status;
      const isActive=(st==='정상'||st==='유지'||st==='진행중'||st==='');
      const n=diff(r.expDate);
      if(r.expDate&&r.expDate.getFullYear()===calYear&&r.expDate.getMonth()===calMonth&&isActive){
        const k=r.expDate.getDate();
        if(!dateMap[k])dateMap[k]=[];
        dateMap[k].push({r,cls:(n!==null&&n>=0&&n<=10)?'ev-suspend':'ev-expire',type:(n!==null&&n>=0&&n<=10)?'정지예정':'만료예정'});
      }
      const isRevoke=(st==='정지'||st==='정지(Suspended)');
      const fsscRevoke=r.src==='FSSC'&&(r.fsscStatus||'').toLowerCase()==='suspended';
      if(isRevoke&&(r.src==='IGC'||fsscRevoke)&&r.suspendDate){
        if(r.suspendDate.getFullYear()===calYear&&r.suspendDate.getMonth()===calMonth){
          const k=r.suspendDate.getDate();
          if(!dateMap[k])dateMap[k]=[];
          dateMap[k].push({r,cls:'ev-revoke',type:'철회예정'});
        }
      }
    });
    for(let i=0;i<firstDay;i++)html+=`<div class="cal-cell other-month"><div class="cal-day">${prevLast-firstDay+1+i}</div></div>`;
    for(let d=1;d<=lastDate;d++){
      const dow=(firstDay+d-1)%7;
      const isToday=d===TODAY.getDate()&&calMonth===TODAY.getMonth()&&calYear===TODAY.getFullYear();
      const dC=dow===0?'sun':dow===6?'sat':'';
      const num=isToday?`<span class="cal-day today-num ${dC}">${d}</span>`:`<div class="cal-day ${dC}">${d}</div>`;
      const evs=dateMap[d]||[];
      const evHtml=evs.slice(0,3).map(e=>`<div class="cal-ev ${e.cls}" onclick="showCalPopup(event,${calYear},${calMonth+1},${d})" title="${e.r.nameKo||e.r.nameEn}">${e.r.nameKo||e.r.nameEn}</div>`).join('');
      const more=evs.length>3?`<div style="font-size:10px;color:var(--muted);padding:1px 4px">+${evs.length-3}개</div>`:'';
      html+=`<div class="cal-cell${isToday?' today-cell':''}">${num}${evHtml}${more}</div>`;
    }
    const remain=(7-(firstDay+lastDate)%7)%7;
    for(let i=1;i<=remain;i++)html+=`<div class="cal-cell other-month"><div class="cal-day">${i}</div></div>`;
    grid.innerHTML=html;
  }catch(e){console.error('renderCal 오류:',e);}
}


export function showCalPopup(e,y,m,d){
  e.stopPropagation();
  const showFSSC=document.getElementById('cf-fssc')?.checked!==false;
  const showFSMS=document.getElementById('cf-fsms')?.checked!==false;
  const evs=[];
  merged.forEach(r=>{
    if(r.src==='FSSC'&&!showFSSC)return;
    if(r.src==='IGC'&&!showFSMS)return;
    const st=r.status;const isActive=(st==='정상'||st==='유지'||st==='진행중'||st==='');
    const n=diff(r.expDate);
    if(r.expDate&&r.expDate.getFullYear()===y&&r.expDate.getMonth()===m-1&&r.expDate.getDate()===d&&isActive)
      evs.push({r,type:(n!==null&&n>=0&&n<=10)?'정지예정':'만료예정'});
    const isRevoke=(st==='정지'||st==='정지(Suspended)');
    const fsscRevoke=r.src==='FSSC'&&(r.fsscStatus||'').toLowerCase()==='suspended';
    if(isRevoke&&(r.src==='IGC'||fsscRevoke)&&r.suspendDate&&r.suspendDate.getFullYear()===y&&r.suspendDate.getMonth()===m-1&&r.suspendDate.getDate()===d)
      evs.push({r,type:'철회예정'});
  });
  if(!evs.length)return;
  const popup=document.getElementById('cal-popup');
  setTxt('cal-popup-title',`${y}년 ${m}월 ${d}일 (${evs.length}건)`);
  document.getElementById('cal-popup-list').innerHTML=evs.map(({r,type})=>`<div class="cal-popup-row" onclick="openModal(${r.idx})"><span class="chip ${type==='철회예정'?'cr':'co'}" style="margin-right:6px">${type}</span>${r.nameKo||r.nameEn} <span style="color:var(--muted);font-size:11px">${r.jobNo}</span></div>`).join('');
  popup.style.left=Math.min(e.pageX,window.innerWidth-340)+'px';
  popup.style.top=(e.pageY+10)+'px';
  popup.classList.add('on');
}
document.addEventListener('click',e=>{
  const p=document.getElementById('cal-popup');
  if(p&&!p.contains(e.target))p.classList.remove('on');
  const rf=document.getElementById('revoke-filter-list');
  const rb=document.getElementById('revoke-filter-btn');
  if(rf&&rb&&!rb.contains(e.target)&&!rf.contains(e.target))rf.style.display='none';
});


// ── UPDATE ──

export function renderUpd(){
  try{
    const rows=merged.filter(r=>r.src==='FSSC'&&r.status==='정지'&&(r.fsscStatus||'').toLowerCase()==='valid');
    setTxt('cnt-upd',rows.length);
    const tb=document.getElementById('tb-upd');if(!tb)return;
    if(!rows.length){tb.innerHTML='<tr class="empty"><td colspan="8">해당 없음</td></tr>';return;}
    tb.innerHTML=rows.map(r=>{
      const c=getChecks()[r.jobNo]||{};
      const done=c.updDone||false;
      const rowStyle=done?'opacity:.4;':'';
      return`<tr onclick="openModal(${r.idx})" style="background:rgba(58,40,16,.2);${rowStyle}"><td>${r.partner||'-'}</td><td>${r.nameKo||'-'}</td><td>${r.nameEn||'-'}</td><td>${r.jobNo}</td><td>${r.coid||'-'}</td><td><span class="chip co">정지 (업무시트)</span></td><td><span class="chip cg">Valid (FSSC)</span></td><td>${fmt(r.expDate)}</td>
      <td style="text-align:center" onclick="event.stopPropagation()"><input type="checkbox" ${done?'checked':''} style="width:15px;height:15px;cursor:pointer;accent-color:#4caf8a" onchange="saveUpdDone('${r.jobNo}',this.checked)"></td></tr>`;
    }).join('');
  }catch(e){console.error('renderUpd 오류:',e);}
}


// ── MODAL ──

// 수정 모드 관련

export const DATE_EDIT_FIELDS=['issuedDate','expDate','suspendDate','initialCert'];

export const getEdits=()=>{
  const raw=LS.get('edits')||{};
  // localStorage는 JSON으로만 저장되므로 Date가 문자열로 바뀜 → 여기서 다시 Date로 복원
  Object.values(raw).forEach(ed=>{
    DATE_EDIT_FIELDS.forEach(f=>{
      if(ed && ed[f]!=null && !(ed[f] instanceof Date)) ed[f]=parseDate(ed[f]);
    });
  });
  return raw;
};

export const saveEdits=o=>{LS.set('edits',o);if(window.FS)window.FS.save('edits',o);};


export function openModal(idx,ctx=''){
  try{
    const r=merged[idx];if(!r)return;
    const mBodyEl=document.getElementById('m-body');
    const mBodyScroll=mBodyEl?mBodyEl.scrollTop:0;
    // 수정된 값 있으면 반영
    const edits=getEdits();
    const ed=edits[r.jobNo]||{};
    const rc={...r,...ed}; // 수정값 오버레이
    const checks=getChecks();const c=checks[r.jobNo]||{};
    const n=diff(rc.expDate);
    const ddayTxt=n===null?'-':n<0?`D+${Math.abs(n)} (만료됨)`:n===0?'D-day':`D-${n}일 남음`;
    const ddayCol=n===null?'':n<=0?'color:#ff6b6b':n<=10?'color:#ff6b6b':n<=30?'color:#f0a040':'color:#4caf8a';
    const sub='margin-left:16px;border-left:2px solid var(--border);border-radius:0 6px 6px 0';
    const isOverseas=isOverseasCountry(rc.country);
    const isFsms=rc.src==='IGC';
    let chkHtml,chkTitle;
    if(ctx==='month'){
      chkTitle='심사주기 안내';
      const internal=c['internal']||false;
      const monthUpload=c['monthUpload']||false;
      const uploadRow=`<div class="chkrow"><input type="checkbox" id="mc-mup" ${monthUpload?'checked':''} style="accent-color:#5a9fd4" onchange="saveReqModal('${rc.jobNo}','monthUpload',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc-mup">업로드 완료</label><span class="chip ${monthUpload?'cb':'cgr'}">${monthUpload?'완료':'미완료'}</span></div>`;
      chkHtml=`
        <div class="chkrow"><input type="checkbox" id="mc-int" ${internal?'checked':''} style="accent-color:#b06aff" onchange="saveReqModal('${rc.jobNo}','internal',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc-int">내부심사</label><span class="chip" style="${internal?'background:#2a1a40;color:#b06aff':'background:var(--br-gray);color:#7a9bb5'}">${internal?'완료':'미완료'}</span></div>
        ${uploadRow}`;
    }else if(ctx==='urgent'){
      chkTitle='처리 현황';
      // 3단계: 0=미처리, 1=필요, 2=완료
      const urgSuspend=c['urgSuspend']||0;
      const urgWithdraw=c['urgWithdraw']||0;
      const urgReport=c['urgReport']||false;
      const urgUpload=c['urgUpload']||false;
      const urgPost2=c['urgPost2']||false;
      const urgNotFirstPost=c['urgNotFirstPost']||false;
      const urgCertAudit=c['urgCertAudit']||false;
      const suspendLabel=urgSuspend===2?'완료':urgSuspend===1?'필요':'미처리';
      const suspendChip=urgSuspend===2?'cg':urgSuspend===1?'co':'cgr';
      const withdrawLabel=urgWithdraw===2?'완료':urgWithdraw===1?'필요':'미처리';
      const withdrawChip=urgWithdraw===2?'cg':urgWithdraw===1?'cr':'cgr';
      chkHtml=`
        <div class="chkrow" style="cursor:pointer" onclick="cycleUrgField('${rc.jobNo}','urgSuspend',${urgSuspend},${idx},'${ctx}')">
          <span class="chklbl">정지처리</span>
          <span class="chip ${suspendChip}">${suspendLabel}</span>
        </div>
        <div class="chkrow" style="cursor:pointer" onclick="cycleUrgField('${rc.jobNo}','urgWithdraw',${urgWithdraw},${idx},'${ctx}')">
          <span class="chklbl">철회처리</span>
          <span class="chip ${withdrawChip}">${withdrawLabel}</span>
        </div>
        <div class="chkrow"><input type="checkbox" id="mc-ur" ${urgReport?'checked':''} style="accent-color:#4caf8a" onchange="saveReqModal('${rc.jobNo}','urgReport',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc-ur">보고서, 신청서 접수</label><span class="chip ${urgReport?'cg':'cgr'}" style="${urgReport?'font-weight:700':''}">${urgReport?'완료':'미완료'}</span></div>
        <div class="chkrow"><input type="checkbox" id="mc-u2" ${urgPost2?'checked':''} style="accent-color:#b06aff" onchange="saveReqModal('${rc.jobNo}','urgPost2',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc-u2">사후2차</label><span class="chip" style="${urgPost2?'background:#2a1a40;color:#b06aff':'background:var(--br-gray);color:#7a9bb5'}">${urgPost2?'완료':'미완료'}</span></div>
        <div class="chkrow"><input type="checkbox" id="mc-nf" ${urgNotFirstPost?'checked':''} style="accent-color:#b06aff" onchange="saveReqModal('${rc.jobNo}','urgNotFirstPost',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc-nf">첫 사후 아님</label><span class="chip" style="${urgNotFirstPost?'background:#2a1a40;color:#b06aff':'background:var(--br-gray);color:#7a9bb5'}">${urgNotFirstPost?'완료':'미완료'}</span></div>
        <div class="chkrow"><input type="checkbox" id="mc-uc" ${urgCertAudit?'checked':''} style="accent-color:#5a9fd4" onchange="saveReqModal('${rc.jobNo}','urgCertAudit',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc-uc">인증결정일 내 심사 완료</label><span class="chip ${urgCertAudit?'cb':'cgr'}">${urgCertAudit?'완료':'미완료'}</span></div>
        <div class="chkrow"><input type="checkbox" id="mc-uu" ${urgUpload?'checked':''} style="accent-color:#5a9fd4" onchange="saveReqModal('${rc.jobNo}','urgUpload',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc-uu">업로드 완료</label><span class="chip ${urgUpload?'cb':'cgr'}">${urgUpload?'완료':'미완료'}</span></div>`;
    }else if(ctx==='req'){
      chkTitle='요청 처리';
      chkHtml=`<div class="chkrow"><input type="checkbox" id="mc0" ${c['reqNotice']?'checked':''} style="accent-color:#4caf8a" onchange="saveReqModal('${rc.jobNo}','reqNotice',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc0">공문 발송</label><span class="chip ${c['reqNotice']?'cg':'cgr'}">${c['reqNotice']?'완료':'미완료'}</span></div>`;
    }else if(ctx==='revoke'){
      chkTitle='철회 처리';
      chkHtml=`<div class="chkrow"><input type="checkbox" id="mc0" ${c['chk3']?'checked':''} style="accent-color:#f0a040" onchange="saveReqModal('${rc.jobNo}','chk3',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc0">컨설팅사 안내</label><span class="chip ${c['chk3']?'co':'cgr'}">${c['chk3']?'완료':'미완료'}</span></div><div class="chkrow"><input type="checkbox" id="mc1" ${c['chk4']?'checked':''} style="accent-color:#ff6b6b" onchange="saveReqModal('${rc.jobNo}','chk4',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc1">철회예정 공문 발송</label><span class="chip ${c['chk4']?'cr':'cgr'}">${c['chk4']?'완료':'미완료'}</span></div><div class="chkrow"><input type="checkbox" id="mc2" ${c['revokeUpload']?'checked':''} style="accent-color:#4caf8a" onchange="saveReqModal('${rc.jobNo}','revokeUpload',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc2">업로드 완료</label><span class="chip ${c['revokeUpload']?'cg':'cgr'}">${c['revokeUpload']?'완료':'미완료'}</span></div><div class="chkrow"><input type="checkbox" id="mc3" ${c['revokeDone']?'checked':''} style="accent-color:#ff6b6b" onchange="saveReqModal('${rc.jobNo}','revokeDone',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc3">철회 완료</label><span class="chip ${c['revokeDone']?'cr':'cgr'}">${c['revokeDone']?'완료':'미완료'}</span></div>`;
    }else{
      chkTitle='요청 처리';
      chkHtml=`<div class="chkrow"><input type="checkbox" id="mc0" ${c['reqNotice']?'checked':''} style="accent-color:#4caf8a" onchange="saveReqModal('${rc.jobNo}','reqNotice',this.checked,${idx},'${ctx}')"><label class="chklbl" for="mc0">공문 발송</label><span class="chip ${c['reqNotice']?'cg':'cgr'}">${c['reqNotice']?'완료':'미완료'}</span></div>`;
    }

    // 수정 여부 표시
    const hasEdit=Object.keys(ed).length>0;

    document.getElementById('m-title').innerHTML=`${rc.nameKo||rc.nameEn||'업체 상세'}${isOverseas?' <span class="chip-overseas">해외</span>':''}${hasEdit?' <span class="chip cb" style="font-size:10px">수정됨</span>':''}`;
    document.getElementById('m-body').innerHTML=`
      <div class="msec"><div class="mstitle" style="display:flex;align-items:center;justify-content:space-between">
        기본 정보
        <button class="tbtn" id="edit-btn" style="font-size:11px;padding:3px 10px" onclick="toggleEditMode(${idx},'${ctx}')">✏️ 수정</button>
      </div>
      <div class="mgrid" id="modal-view-mode">
        <span class="ml">출처</span><span class="mv">${rc.src==='IGC'?'<span class="chip cb">ISO (FSMS)</span>':'<span class="chip cgr">FSSC</span>'}</span>
        <span class="ml">현황</span><span class="mv">${statusChip(rc.status)}</span>
        <span class="ml">국가</span><span class="mv">${rc.country||'KOREA'}${isOverseas?' <span class="chip-overseas">해외</span>':''}</span>
        <span class="ml">컨설팅사</span><span class="mv">${rc.partner||'-'}</span>
        <span class="ml">컨설팅사 이메일</span><span class="mv">${rc.consultEmail||findEmail(rc.partner)||'<span style="color:var(--muted)">미등록</span>'}</span>
        <span class="ml">잡넘버</span><span class="mv hi">${rc.jobNo}</span>
        <span class="ml">COID</span><span class="mv">${rc.coid||'-'}</span>
      </div></div>
      <div class="msec"><div class="mstitle">업체명</div><div class="mgrid">
        <span class="ml">국문</span><span class="mv">${rc.nameKo||'-'}</span>
        <span class="ml">영문</span><span class="mv">${rc.nameEn||'-'}</span>
      </div></div>
      <div class="msec"><div class="mstitle">인증 일정</div><div class="mgrid">
        <span class="ml">발행일</span><span class="mv">${fmt(rc.issuedDate)}</span>
        ${rc.initialCert?`<span class="ml">Initial certification</span><span class="mv">${fmt(rc.initialCert)}</span>`:''}
        <span class="ml">만료일</span><span class="mv">${fmt(rc.expDate)}</span>
        <span class="ml">남은 기간</span><span class="mv" style="${ddayCol};font-weight:700">${ddayTxt}</span>
        ${rc.suspendDate?`<span class="ml">정지날짜</span><span class="mv">${fmt(rc.suspendDate)}</span>`:''}
      </div></div>
      <div class="msec"><div class="mstitle">업체 연락처</div><div class="mgrid">
        <span class="ml">이메일</span><span class="mv">${rc.email||'-'}</span>
        <span class="ml">전화번호</span><span class="mv">${rc.phone||'-'}</span>
        <span class="ml">주소</span><span class="mv" style="white-space:normal">${rc.address||'-'}</span>
      </div></div>
      <div class="msec"><div class="mstitle">${chkTitle}</div>${chkHtml}</div>`;
    if(mBodyScroll)document.getElementById('m-body').scrollTop=mBodyScroll;
    document.getElementById('modal').classList.add('on');
    // 현재 열린 모달 정보 저장
    document.getElementById('modal')._idx=idx;
    document.getElementById('modal')._ctx=ctx;
  }catch(e){console.error('openModal 오류:',e);}
}

export function toggleEditMode(idx,ctx){
  const r=merged[idx];if(!r)return;
  const edits=getEdits();
  const ed=edits[r.jobNo]||{};
  const rc={...r,...ed};
  // 수정 폼으로 교체
  const grid=document.getElementById('modal-view-mode');
  const btn=document.getElementById('edit-btn');
  if(btn.textContent.includes('수정')){
    // → 수정 모드
    btn.textContent='💾 저장';
    btn.style.background='var(--accent)';btn.style.color='#fff';btn.style.borderColor='var(--accent)';
    btn.onclick=()=>saveEditMode(idx,ctx);
    grid.innerHTML=`
      <span class="ml">출처</span><span class="mv">${rc.src==='IGC'?'<span class="chip cb">IGC</span>':'<span class="chip cgr">FSSC</span>'}</span>
      <span class="ml">현황</span><span class="mv"><select id="ed-status" style="padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px">
        <option ${rc.status==='정상'?'selected':''}>정상</option>
        <option ${rc.status==='진행중'?'selected':''}>진행중</option>
        <option ${rc.status==='정지'?'selected':''}>정지</option>
        <option ${rc.status==='만료'?'selected':''}>만료</option>
        <option ${rc.status==='취소'?'selected':''}>취소</option>
      </select></span>
      <span class="ml">컨설팅사</span><span class="mv"><input id="ed-partner" value="${rc.partner||''}" style="width:100%;padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>
      <span class="ml">업체명 (국문)</span><span class="mv"><input id="ed-nameKo" value="${rc.nameKo||''}" style="width:100%;padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>
      <span class="ml">업체명 (영문)</span><span class="mv"><input id="ed-nameEn" value="${rc.nameEn||''}" style="width:100%;padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>
      <span class="ml">COID</span><span class="mv"><input id="ed-coid" value="${rc.coid||''}" style="width:100%;padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>
      <span class="ml">발행일</span><span class="mv"><input id="ed-issued" type="date" value="${fmt(rc.issuedDate)!=='-'?fmt(rc.issuedDate):''}" style="padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>
      <span class="ml">만료일</span><span class="mv"><input id="ed-exp" type="date" value="${fmt(rc.expDate)!=='-'?fmt(rc.expDate):''}" style="padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>
      <span class="ml">정지일</span><span class="mv"><input id="ed-suspend" type="date" value="${fmt(rc.suspendDate)!=='-'?fmt(rc.suspendDate):''}" style="padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>
      <span class="ml">이메일</span><span class="mv"><input id="ed-email" value="${rc.email||''}" style="width:100%;padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>
      <span class="ml">전화번호</span><span class="mv"><input id="ed-phone" value="${rc.phone||''}" style="width:100%;padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>
      <span class="ml">주소</span><span class="mv"><input id="ed-address" value="${rc.address||''}" style="width:100%;padding:4px 6px;border-radius:4px;border:1px solid var(--border);background:var(--panel2);color:var(--text);font-size:12px"></span>`;
  }
}


export function saveEditMode(idx,ctx){
  const r=merged[idx];if(!r)return;
  const edits=getEdits();
  const ed={
    status:document.getElementById('ed-status')?.value||r.status,
    partner:document.getElementById('ed-partner')?.value||r.partner,
    nameKo:document.getElementById('ed-nameKo')?.value||r.nameKo,
    nameEn:document.getElementById('ed-nameEn')?.value||r.nameEn,
    coid:document.getElementById('ed-coid')?.value||r.coid,
    issuedDate:parseDate(document.getElementById('ed-issued')?.value)||r.issuedDate,
    expDate:parseDate(document.getElementById('ed-exp')?.value)||r.expDate,
    suspendDate:parseDate(document.getElementById('ed-suspend')?.value)||r.suspendDate,
    email:document.getElementById('ed-email')?.value||r.email,
    phone:document.getElementById('ed-phone')?.value||r.phone,
    address:document.getElementById('ed-address')?.value||r.address,
  };
  edits[r.jobNo]=ed;
  saveEdits(edits);
  Object.assign(merged[idx],ed);
  renderDashContent();renderAll();
  openModal(idx,ctx);
}

const URGENT_FIELD_LABELS={
  urgSuspend:'정지처리', urgWithdraw:'철회처리', urgReport:'보고서 접수',
  urgPost2:'사후2차', urgNotFirstPost:'첫 사후 아님', urgUpload:'업로드 완료', urgCertAudit:'인증결정일 내 심사',
  monthUpload:'업로드 완료(만료예정)', internal:'내부심사',
};
const URGENT_STAGE_LABELS={0:'미처리',1:'필요',2:'완료'};

function logUrgentAction(jobNo,field,valueLabel){
  const r=merged.find(m=>m.jobNo===jobNo);
  addLog({jobNo, nameKo:r?(r.nameKo||r.nameEn):'', field, fieldLabel:URGENT_FIELD_LABELS[field]||field, valueLabel});
}

export function cycleUrgField(jobNo,field,cur,idx,ctx){
  // 0→1→2→0 순환
  const next=(cur+1)%3;
  const c=getChecks();if(!c[jobNo])c[jobNo]={};
  c[jobNo][field]=next;saveChecks(c);
  logUrgentAction(jobNo,field,URGENT_STAGE_LABELS[next]);
  // 완료(2) 상태가 되면 다른 탭(전체 목록/만료 예정/캘린더 등)에도 현황이 반영되도록 override 적용
  // 완료에서 벗어나면 override 해제
  if(field==='urgSuspend')applyOverride(jobNo,'suspended',next===2);
  else if(field==='urgWithdraw')applyOverride(jobNo,'withdrawn',next===2);
  renderDashContent();openModal(idx,ctx);
}

export function saveReqModal(jobNo,field,val,idx,ctx){
  const c=getChecks();if(!c[jobNo])c[jobNo]={};c[jobNo][field]=val;saveChecks(c);
  if(ctx==='urgent')logUrgentAction(jobNo,field,val?'체크':'해제');
  renderDashContent();renderAll();openModal(idx,ctx);
}

export function closeModal(){document.getElementById('modal').classList.remove('on');}

// ── URGENT 활동 기록 (전용 탭) ──
function fmtLogTime(ts){
  const d=new Date(ts);
  const p=n=>String(n).padStart(2,'0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}
export function renderLogPanel(){
  const logs=getLogs();
  const body=document.getElementById('log-body');
  if(!logs.length){body.innerHTML='<div style="padding:20px;text-align:center;color:var(--muted);font-size:13px">아직 기록이 없어요. 이 기능을 추가한 시점부터 URGENT 탭에서 체크한 내역이 쌓입니다 (그 이전 기록은 남아있지 않아요).</div>';return;}
  body.innerHTML=`<div style="font-size:12px;color:var(--muted);margin-bottom:10px">이 기능을 추가한 시점부터의 기록입니다. 그 이전 클릭 내역은 저장되어 있지 않아요.</div>
    <div style="max-height:75vh;overflow-y:auto">` + logs.map(l=>`
    <div style="padding:9px 4px;border-bottom:1px solid var(--border);display:flex;flex-direction:column;gap:2px">
      <div style="font-size:12px;color:var(--muted)">${fmtLogTime(l.ts)}</div>
      <div style="font-size:13px"><b>${l.nameKo||l.jobNo}</b> · ${l.fieldLabel} → <span style="color:var(--accent)">${l.valueLabel}</span></div>
    </div>`).join('') + `</div>`;
}


// ── 주기 안내 메일링 (메인 대시보드와 별개의 업로드로 동작) ──
function todayStr(){
  const p=n=>String(n).padStart(2,'0');
  return `${TODAY.getFullYear()}-${p(TODAY.getMonth()+1)}-${p(TODAY.getDate())}`;
}
function mailingFilename(){
  return `주기안내(${TODAY.getMonth()+1}월).xls`;
}

const PROC_OPTIONS=['미정','자체심사','신청은됨','입금이됨','해당X','프로세스진행중','철회예정','철회'];
const PROC_COLOR_SET={
  '미정':{bg:'#4a5568',fg:'#fff'},
  '자체심사':{bg:'#7a5fc9',fg:'#fff'},
  '신청은됨':{bg:'#29b6d8',fg:'#fff'},
  '입금이됨':{bg:'#7cb342',fg:'#fff'},
  '해당X':{bg:'#1a1a1a',fg:'#fff'},
  '프로세스진행중':{bg:'#6b4a2a',fg:'#ffd54a'},
  '철회예정':{bg:'#e04040',fg:'#fff'},
  '철회':{bg:'#8b1a1a',fg:'#fff'},
};
function procColorSet(v){return PROC_COLOR_SET[v]||PROC_COLOR_SET['미정'];}
function procColor(v){return procColorSet(v).bg;}
function hexToRgba(hex,alpha){
  const h=hex.replace('#','');
  const r=parseInt(h.substring(0,2),16),g=parseInt(h.substring(2,4),16),b=parseInt(h.substring(4,6),16);
  return`rgba(${r},${g},${b},${alpha})`;
}
// 해당X 상태인데 인증유지의사가 있으면 검정 대신 연두색으로 (텍스트는 여전히 '해당X')
function procDisplayColorSet(proc,intentKept){
  if(proc==='해당X'&&intentKept)return{bg:'#8fd19e',fg:'#153a1a'};
  return procColorSet(proc);
}
// 진행상태에 따라 총 뷰 행 전체를 옅게 물들여서 한눈에 보이게 함 (미정은 색 없음)
function procRowStyle(proc,intentKept){
  if(proc==='해당X'&&intentKept)return`background:${hexToRgba('#8fd19e',.28)}`;
  if(!proc||proc==='미정')return'';
  return`background:${hexToRgba(procColorSet(proc).bg,.28)}`;
}
// 월별안내 사유와 총 탭 진행상태의 이름이 다르지만 같은 뜻인 경우 매핑
const NOTICE_TO_PROC_MAP={'진행중':'프로세스진행중'};

export function switchMailingView(view){
  window._mailingView=view;
  window._mailingMonthFilter=null; // 뷰 전환시 월 필터 초기화
  renderMailingTab();
}

export function setMailingMonthFilter(m){
  window._mailingMonthFilter=(m===null||window._mailingMonthFilter===m)?null:m;
  renderMailingTab();
}

// 컨설팅 이메일이 매칭되지 않은 컨설팅사 안내 배너 ('자체' 포함은 원래 이메일 매칭 대상이 아니므로 제외)
function renderUnmatchedBanner(all){
  const unmatched=[...new Set(all.filter(r=>r.partner&&!r.partner.includes('자체')&&!r.consultEmail).map(r=>r.partner))];
  if(!unmatched.length)return'';
  return `<div style="margin-bottom:12px;padding:10px 12px;background:#3a2a10;border:1px solid #7a5a20;border-radius:8px;font-size:12px;color:#e0b060">
    ⚠️ 컨설팅 이메일이 매칭되지 않은 컨설팅사 ${unmatched.length}곳이 있어요 — <b>이메일 탭</b>에서 별칭을 추가해주세요:
    <div style="margin-top:4px">${unmatched.map(p=>`<span style="display:inline-block;background:#4a3a1a;padding:2px 8px;border-radius:10px;margin:2px">${p}</span>`).join('')}</div>
  </div>`;
}

export function renderMailingTab(){
  const view=window._mailingView||'month';
  const uploadHtml=`
    <div style="display:flex;gap:10px;align-items:center;margin-bottom:14px;flex-wrap:wrap">
      <button class="tbtn pri" onclick="document.getElementById('fi-mailing-fssc').click()">📂 FSSC 목록</button>
      <input type="file" id="fi-mailing-fssc" accept=".xlsx,.xls" style="display:none" onchange="loadMailingFssc(event)">
      <span id="lbl-mailing-fssc" style="font-size:12px;color:var(--muted)">${mailingFsscName||'업로드 안됨'}</span>
      <button class="tbtn pri" onclick="document.getElementById('fi-mailing-job').click()">📂 업무시트</button>
      <input type="file" id="fi-mailing-job" accept=".xlsx,.xls" style="display:none" onchange="loadMailingJob(event)">
      <span id="lbl-mailing-job" style="font-size:12px;color:var(--muted)">${mailingJobName||'업로드 안됨 (없어도 조회는 가능, 컨설팅사/국문명만 비어보임)'}</span>
    </div>
    <div style="display:flex;gap:6px;margin-bottom:14px">
      <button class="tbtn ${view==='month'?'pri':''}" onclick="switchMailingView('month')">주기안내(월별)</button>
      <button class="tbtn ${view==='total'?'pri':''}" onclick="switchMailingView('total')">주기안내(총)</button>
    </div>`;
  const body=document.getElementById('mailing-body');

  const all=buildMailingData();
  if(!all.length){
    body.innerHTML=uploadHtml+'<div style="padding:30px;text-align:center;color:var(--muted);font-size:13px">FSSC 목록 파일을 업로드해주세요.</div>';
    return;
  }

  const banner=renderUnmatchedBanner(all);
  const viewHtml=view==='total'?renderMailingTotalView(all):renderMailingMonthView(all);
  body.innerHTML=uploadHtml+banner+viewHtml;
}

const NOTICE_STATUS_OPTIONS=['자체심사','신청은됨','심사공문','입금이됨','진행중','철회요청'];
const NOTICE_LABELS={notice1:'안내메일 1차',notice2:'안내메일 2차'};

function logMailingAction(jobNo,fieldLabel,valueLabel){
  const all=buildMailingData();
  const r=all.find(x=>x.jobNo===jobNo);
  addLog({jobNo,nameKo:r?(r.nameKo||r.nameEn):'',field:'mailing',fieldLabel,valueLabel});
}

// ── 메일 초안 작성 (mailto: 링크) - 기본 메일프로그램을 열어 받는사람/제목/본문을 미리 채워줌 ──
function buildMailingMailto(r){
  const to=r.consultEmail||r.email||'';
  const cc=(r.consultEmail&&r.email&&r.consultEmail!==r.email)?r.email:'';
  const name=r.nameKo||r.nameEn||'';
  const subject=`[${name}] 인증 심사주기 안내`;
  const body=`안녕하세요, ${r.partner?r.partner+' 담당자님.':'담당자님.'}\n\n`+
    `${name} (잡넘버: ${r.jobNo}) 업체의 인증서(${r.certNo||'-'})가 ${fmt(r.expDate)}에 만료 예정이라 안내드립니다.\n\n`+
    `심사 절차 관련하여 확인 부탁드립니다.\n\n감사합니다.`;
  const params=[];
  if(cc)params.push('cc='+encodeURIComponent(cc));
  params.push('subject='+encodeURIComponent(subject));
  params.push('body='+encodeURIComponent(body));
  return`mailto:${to}?${params.join('&')}`;
}
export function openMailingMailDraft(jobNo){
  const all=buildMailingData();
  const r=all.find(x=>x.jobNo===jobNo);
  if(!r)return;
  if(!r.consultEmail&&!r.email){alert('받는사람으로 쓸 이메일 정보가 없어요 (컨설팅 이메일/업체 이메일 모두 비어있음).');return;}
  window.location.href=buildMailingMailto(r);
}

function noticeText(n){
  if(!n)return'';
  if(n.mode==='o'){
    if(!n.oDate)return'O';
    const parts=n.oDate.split('-'); // YYYY-MM-DD
    if(parts.length!==3)return`O (${n.oDate})`;
    const[,m,d]=parts;
    return`${Number(m)}월심사주기안내(${m}${d})`;
  }
  if(n.mode==='x')return n.xStatus?`X - ${n.xStatus} (${n.xDate||''})`:'X (미정)';
  return'';
}

// 행 전체 색상: 철회요청 > 인증유지의사 > 그 외 상태 순 우선. 라이트/다크 테마 모두에서 또렷이 보이되 새까맣지 않은 톤으로
function noticeRowStyle(ml){
  const n1=ml.notice1||{},n2=ml.notice2||{};
  const isLight=document.body.classList.contains('light');
  if(n1.xStatus==='철회요청'||n2.xStatus==='철회요청')
    return isLight?'background:#f5c99a':'background:#5c3a1a';
  if(ml.intentKept)
    return isLight?'background:#c8e6c9':'background:#1f4a28';
  if(n1.xStatus||n2.xStatus)
    return isLight?'background:#e3e7ec':'background:#2e3d4d';
  return'';
}

// O를 선택하면 O만, X를 선택하면 X(+사유)만 보이고, 아직 아무것도 안 골랐을 때만 둘 다 보여줌
function buildNoticeCell(jobNo,noticeKey,notice){
  const mode=notice.mode||null;
  if(mode==='o'){
    return`<td onclick="event.stopPropagation()">
      <label style="display:flex;align-items:center;gap:6px;font-size:11px;cursor:pointer">
        <input type="checkbox" checked onchange="setMailingNoticeMode('${jobNo}','${noticeKey}','o',this.checked)">O
        <span style="color:var(--muted)">${notice.oDate||''}</span>
        <button onclick="event.stopPropagation();editMailingNoticeDate('${jobNo}','${noticeKey}',this)" style="background:none;border:none;cursor:pointer;font-size:11px;padding:0" title="날짜 수정">✏️</button>
      </label>
    </td>`;
  }
  if(mode==='x'){
    const status=notice.xStatus||'';
    return`<td onclick="event.stopPropagation()">
      <div style="display:flex;align-items:center;gap:6px">
        <label style="display:flex;align-items:center;gap:4px;font-size:11px;cursor:pointer"><input type="checkbox" checked onchange="setMailingNoticeMode('${jobNo}','${noticeKey}','x',this.checked)">X</label>
        <select onclick="event.stopPropagation()" onchange="event.stopPropagation();setMailingNoticeStatus('${jobNo}','${noticeKey}',this.value)" style="background:var(--panel2);color:var(--text);border:1px solid var(--border);border-radius:4px;padding:2px 4px;font-size:11px">
          <option value="" ${!status?'selected':''}>선택</option>
          ${NOTICE_STATUS_OPTIONS.map(o=>`<option value="${o}" ${o===status?'selected':''}>${o}</option>`).join('')}
        </select>
      </div>
      ${status?`<div style="font-size:10px;color:var(--muted);margin-top:2px">${notice.xDate||''}</div>`:''}
    </td>`;
  }
  return`<td onclick="event.stopPropagation()">
    <div style="display:flex;gap:8px">
      <label style="display:flex;align-items:center;gap:3px;font-size:11px;cursor:pointer"><input type="checkbox" onchange="setMailingNoticeMode('${jobNo}','${noticeKey}','o',this.checked)">O</label>
      <label style="display:flex;align-items:center;gap:3px;font-size:11px;cursor:pointer"><input type="checkbox" onchange="setMailingNoticeMode('${jobNo}','${noticeKey}','x',this.checked)">X</label>
    </div>
  </td>`;
}

// ── 주기안내(월별): 만료일의 "월"만 기준 (연도 무관, 현재월 +1~+3) ──
function renderMailingMonthView(all){
  const mailingLog=getMailingLog();
  const collapsed=window._mailingCollapsed||{};
  const emailFilter=window._mailingEmailFilter||'';
  const targetMonths=[1,2,3].map(i=>(TODAY.getMonth()+i)%12);
  const inWindowAll=all.filter(r=>targetMonths.includes(r.expDate.getMonth()));
  // 필터용 이메일 목록은 필터를 적용하기 전 전체 기준으로 뽑음 (필터링 후엔 선택 가능한 옵션이 줄어들지 않도록)
  const emailOptions=[...new Set(inWindowAll.map(r=>r.consultEmail).filter(Boolean))].sort();
  const inWindow=emailFilter?inWindowAll.filter(r=>r.consultEmail===emailFilter):inWindowAll;
  const byMonth={};
  inWindow.forEach(r=>{
    const m=r.expDate.getMonth();
    if(!byMonth[m])byMonth[m]=[];
    byMonth[m].push(r);
  });
  const orderedMonths=targetMonths.filter(m=>byMonth[m]&&byMonth[m].length);
  // 같은 컨설팅 이메일주소끼리 묶어서 보이도록 - 이메일 우선, 그 안에서는 만료일 오름차순
  orderedMonths.forEach(m=>byMonth[m].sort((a,b)=>{
    const ea=a.consultEmail||'',eb=b.consultEmail||'';
    if(ea!==eb)return ea.localeCompare(eb);
    return a.expDate-b.expDate;
  }));

  const filterHtml=`<div style="display:flex;align-items:center;gap:8px;margin-bottom:12px">
    <label style="font-size:12px;color:var(--muted)">컨설팅 이메일주소 필터</label>
    <select onchange="setMailingEmailFilter(this.value)" style="background:var(--panel2);color:var(--text);border:1px solid var(--border);border-radius:4px;padding:4px 8px;font-size:12px">
      <option value="">전체 보기</option>
      ${emailOptions.map(e=>`<option value="${e}" ${e===emailFilter?'selected':''}>${e}</option>`).join('')}
    </select>
  </div>`;

  if(!orderedMonths.length)return filterHtml+'<div style="padding:30px;text-align:center;color:var(--muted);font-size:13px">'+(emailFilter?'이 이메일주소로 해당하는 업체가 없어요.':'다음 3개월(월 기준, 연도 무관) 내 만료 예정인 업체가 없어요.')+'</div>';

  let html=filterHtml+`<table><thead><tr>
    <th>컨설팅 이메일주소</th><th>컨설팅사</th><th>업체명</th><th>Client number</th>
    <th>이메일</th><th>만료일</th><th>인증서번호</th><th>안내메일(1차)</th><th>안내메일(2차)</th><th>인증유지의사</th><th>메일</th>
  </tr></thead><tbody>`;
  orderedMonths.forEach(m=>{
    const rows=byMonth[m];
    const isCollapsed=!!collapsed[m];
    html+=`<tr><td colspan="11" style="background:var(--panel2);font-weight:700;padding:8px 11px;cursor:pointer" onclick="toggleMailingMonthCollapse(${m})">
      <div style="display:flex;justify-content:space-between;align-items:center">
        <label onclick="event.stopPropagation()" style="display:flex;align-items:center;gap:6px;cursor:pointer"><input type="checkbox" class="mailing-month-dl" value="${m}" checked>${m+1}월 만료 (${rows.length}건, 연도 무관)</label>
        <span style="font-size:11px">${isCollapsed?'▶':'▼'}</span>
      </div>
    </td></tr>`;
    if(isCollapsed)return;
    rows.forEach(r=>{
      const ml=mailingLog[r.jobNo]||{};
      const rowStyle=noticeRowStyle(ml);
      html+=`<tr onclick="openMailingDetail('${r.jobNo}')" style="cursor:pointer;${rowStyle}">
        <td>${r.consultEmail||'-'}</td>
        <td title="${r.partner}">${r.partner||'-'}</td>
        <td title="${r.nameEn}">${r.nameKo||r.nameEn||'-'}</td>
        <td>${r.jobNo}</td>
        <td>${r.email||'-'}</td>
        <td>${fmt(r.expDate)}</td>
        <td>${r.certNo||'-'}</td>
        ${buildNoticeCell(r.jobNo,'notice1',ml.notice1||{})}
        ${buildNoticeCell(r.jobNo,'notice2',ml.notice2||{})}
        <td onclick="event.stopPropagation()"><button onclick="toggleMailingIntent('${r.jobNo}')" style="background:${ml.intentKept?'#8fd19e':'var(--br-gray)'};color:${ml.intentKept?'#153a1a':'#7a9bb5'};border:none;border-radius:14px;padding:4px 10px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap">인증유지의사 있음</button></td>
        <td onclick="event.stopPropagation()"><button onclick="openMailingMailDraft('${r.jobNo}')" title="메일 초안 작성" style="background:var(--panel2);color:var(--text);border:1px solid var(--border);border-radius:6px;padding:4px 8px;font-size:13px;cursor:pointer">✉️</button></td>
      </tr>`;
    });
  });
  html+='</tbody></table>';
  html+=`<div style="margin-top:12px"><button class="tbtn pri" onclick="exportMailingMonthExcel()">📥 엑셀 다운로드 (체크된 월만)</button></div>`;
  return html;
}

// ── 주기안내(총): 만료월 버튼으로 필터, 만료일 오름차순, 진행상태 드롭다운 + 1차/2차 안내(자유 텍스트) ──
function renderMailingTotalView(all){
  const mailingLog=getMailingLog();
  const filter=window._mailingMonthFilter;
  const monthBtns=`<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">
    ${Array.from({length:12},(_,m)=>`<button class="tbtn ${filter===m?'pri':''}" style="min-width:44px" onclick="setMailingMonthFilter(${m})">${m+1}월</button>`).join('')}
    ${filter!==null&&filter!==undefined?`<button class="tbtn" onclick="setMailingMonthFilter(null)">전체보기</button>`:''}
  </div>`;

  const filtered=(filter===null||filter===undefined)?all:all.filter(r=>r.expDate.getMonth()===filter);
  // 특정 월로 필터링된 상태에서는 연도와 무관하게 "일(day)" 기준으로 오름차순 정렬
  // (필터 자체가 연도를 무시하는데 정렬에서 연도가 우선시되면 앞뒤가 뒤바뀌어 보일 수 있음)
  const sorted=[...filtered].sort((a,b)=>{
    if(filter===null||filter===undefined)return a.expDate-b.expDate;
    return a.expDate.getDate()-b.expDate.getDate();
  });
  if(!sorted.length)return monthBtns+'<div style="padding:30px;text-align:center;color:var(--muted);font-size:13px">해당하는 업체가 없어요.</div>';

  let html=`<table><thead><tr>
    <th>만료일</th><th>업체이름</th><th>잡넘버</th>
    <th>심사 프로세스 진행 중 확인<br><span style="font-weight:400">- 보고서 검토 시트<br>- FSSC 심사 관리 시트</span></th>
    <th colspan="2" style="text-align:center">심사 주기 안내 시트 확인 (월별안내와 잡넘버 기준 자동 연동)</th><th>메일</th>
  </tr><tr><th></th><th></th><th></th><th></th><th>1차 안내</th><th>2차 안내</th><th></th></tr></thead><tbody>`;
  sorted.forEach(r=>{
    const ml=mailingLog[r.jobNo]||{};
    const proc=ml.procStatus||'미정';
    const pc=procDisplayColorSet(proc,ml.intentKept);
    html+=`<tr onclick="openMailingDetail('${r.jobNo}')" style="cursor:pointer;${procRowStyle(proc,ml.intentKept)}">
      <td>${fmt(r.expDate)}</td>
      <td title="${r.nameEn}">${r.nameKo||r.nameEn||'-'}<div style="font-size:10px;color:var(--muted)">${r.jobNo}</div></td>
      <td>${r.jobNo}</td>
      <td onclick="event.stopPropagation()" style="position:relative">
        <button onclick="toggleMailingProcDropdown('${r.jobNo}',this)" style="background:${pc.bg};color:${pc.fg};border:none;border-radius:14px;padding:4px 10px;font-size:11px;font-weight:600;cursor:pointer">${proc} ▾</button>
        <div id="proc-list-${r.jobNo}" style="display:none;z-index:9999;background:var(--panel);border:1px solid var(--border);border-radius:6px;min-width:120px;box-shadow:0 4px 12px rgba(0,0,0,.35)">
          ${PROC_OPTIONS.map(o=>`<div onclick="selectMailingProc('${r.jobNo}','${o}')" style="padding:6px 10px;font-size:11px;cursor:pointer;color:var(--text);${o===proc?'font-weight:700;background:var(--panel2)':''}">${o===proc?'✓ ':''}${o}</div>`).join('')}
        </div>
      </td>
      <td style="font-size:13px;font-weight:700">${noticeText(ml.notice1)||'<span style="color:var(--muted)">월별안내에서 설정</span>'}</td>
      <td style="font-size:13px;font-weight:700">${noticeText(ml.notice2)||'<span style="color:var(--muted)">월별안내에서 설정</span>'}</td>
      <td onclick="event.stopPropagation()"><button onclick="openMailingMailDraft('${r.jobNo}')" title="메일 초안 작성" style="background:var(--panel2);color:var(--text);border:1px solid var(--border);border-radius:6px;padding:4px 8px;font-size:13px;cursor:pointer">✉️</button></td>
    </tr>`;
  });
  html+='</tbody></table>';
  html+=`<div style="margin-top:12px"><button class="tbtn pri" onclick="exportMailingTotalExcel()">📥 엑셀 다운로드 (${filter===null||filter===undefined?'전체':`${filter+1}월만`})</button></div>`;
  return monthBtns+html;
}

export function openMailingDetail(jobNo){
  const all=buildMailingData();
  const r=all.find(x=>x.jobNo===jobNo);
  if(!r)return;
  const ml=getMailingLog()[jobNo]||{};
  const noticeTextOrDash=n=>noticeText(n)||'-';
  document.getElementById('m-title').innerHTML=`${r.nameKo||r.nameEn||'업체 상세'}`;
  document.getElementById('m-body').innerHTML=`
    <div class="mgrid">
      <span class="ml">국문 업체명</span><span class="mv">${r.nameKo||'-'}</span>
      <span class="ml">영문 업체명</span><span class="mv">${r.nameEn||'-'}</span>
      <span class="ml">잡넘버</span><span class="mv">${r.jobNo}</span>
      <span class="ml">컨설팅사</span><span class="mv">${r.partner||'-'}</span>
      <span class="ml">컨설팅 이메일</span><span class="mv">${r.consultEmail||'-'}</span>
      <span class="ml">이메일</span><span class="mv">${r.email||'-'}</span>
      <span class="ml">만료일</span><span class="mv">${fmt(r.expDate)}</span>
      <span class="ml">인증서번호</span><span class="mv">${r.certNo||'-'}</span>
      <span class="ml">진행 상태(총)</span><span class="mv">${ml.procStatus||'미정'}</span>
      <span class="ml">안내메일 1차</span><span class="mv">${noticeTextOrDash(ml.notice1)}</span>
      <span class="ml">안내메일 2차</span><span class="mv">${noticeTextOrDash(ml.notice2)}</span>
      <span class="ml">인증유지의사</span><span class="mv">${ml.intentKept?'있음':'-'}</span>
    </div>`;
  document.getElementById('modal').classList.add('on');
}

export function saveMailingField(jobNo,field,val){
  setMailingField(jobNo,field,val);
}

export function saveMailingProc(jobNo,val,selEl){
  setMailingField(jobNo,'procStatus',val);
  if(selEl)selEl.style.background=procColor(val);
}

// 총 뷰의 진행상태 커스텀 드롭다운 - 열었을 때 현재 선택된 값은 목록에서 빼서 중복으로 안 보이게 함
export function toggleMailingProcDropdown(jobNo,btn){
  document.querySelectorAll('[id^="proc-list-"]').forEach(el=>{
    if(el.id!==`proc-list-${jobNo}`)el.style.display='none';
  });
  const el=document.getElementById(`proc-list-${jobNo}`);
  if(!el)return;
  const isOpen=el.style.display==='block';
  if(isOpen){
    el.style.display='none';
  }else if(btn){
    // 표 스크롤 영역(overflow) 안에 있으면 position:absolute가 잘려서 안 보이므로
    // 화면 좌표 기준 position:fixed로 버튼 바로 아래에 띄운다
    const rect=btn.getBoundingClientRect();
    el.style.position='fixed';
    el.style.top=(rect.bottom+4)+'px';
    el.style.left=rect.left+'px';
    el.style.display='block';
  }
}
export function selectMailingProc(jobNo,val){
  saveMailingProc(jobNo,val,null);
  renderMailingTab();
}

// 체크하면 오늘 날짜 자동 기입, 해제하면 비움
export function saveMailingCheck(jobNo,field,checked){
  setMailingField(jobNo,field,checked?todayStr():'');
  renderMailingTab();
}

// O/X 체크(둘 중 하나만 선택 가능). 모드가 바뀌면 이전 X 사유(색상 원인)는 초기화해서 색이 남지 않게 함
export function setMailingNoticeMode(jobNo,noticeKey,mode,checked){
  const label=NOTICE_LABELS[noticeKey]||noticeKey;
  const cur=(getMailingLog()[jobNo]||{})[noticeKey]||{};
  if(!checked){
    if(cur.mode===mode){
      setMailingNotice(jobNo,noticeKey,{mode:null,xStatus:'',xDate:''});
      logMailingAction(jobNo,label,'선택 해제');
    }
  }else if(mode==='o'){
    const oDate=cur.oDate||todayStr();
    setMailingNotice(jobNo,noticeKey,{mode:'o',oDate,xStatus:'',xDate:''});
    logMailingAction(jobNo,label,`O (${oDate})`);
    // 주기안내 발송(O) 완료된 건은 총 탭 진행상태를 '해당X'로 자동 반영
    setMailingField(jobNo,'procStatus','해당X');
  }else{
    setMailingNotice(jobNo,noticeKey,{mode:'x'});
    logMailingAction(jobNo,label,'X 선택');
  }
  renderMailingTab();
}

// X 선택시 뜨는 사유 드롭다운
export function setMailingNoticeStatus(jobNo,noticeKey,status){
  const label=NOTICE_LABELS[noticeKey]||noticeKey;
  if(!status){
    setMailingNotice(jobNo,noticeKey,{xStatus:'',xDate:''});
  }else{
    const xDate=todayStr();
    setMailingNotice(jobNo,noticeKey,{xStatus:status,xDate});
    logMailingAction(jobNo,label,`${status} (${xDate})`);
    // 월별에서 고른 사유가 총 탭 진행상태 목록에도 같은(또는 매핑된) 이름으로 있으면 그쪽도 자동으로 맞춰줌
    if(PROC_OPTIONS.includes(status)){
      setMailingField(jobNo,'procStatus',status);
    }else if(NOTICE_TO_PROC_MAP[status]){
      setMailingField(jobNo,'procStatus',NOTICE_TO_PROC_MAP[status]);
    }
  }
  renderMailingTab();
}

// O 선택시 자동기입된 날짜를 수동으로 수정
export function editMailingNoticeDate(jobNo,noticeKey,btn){
  const span=btn.previousElementSibling;
  const cur=span.textContent;
  const input=document.createElement('input');
  input.type='date';input.value=cur;
  input.style.cssText='background:var(--panel2);color:var(--text);border:1px solid var(--border);border-radius:4px;padding:2px 4px;font-size:11px';
  input.onclick=e=>e.stopPropagation();
  input.onchange=()=>{
    setMailingNotice(jobNo,noticeKey,{oDate:input.value});
    logMailingAction(jobNo,NOTICE_LABELS[noticeKey]||noticeKey,`날짜 수정 (${input.value})`);
    renderMailingTab();
  };
  btn.parentElement.replaceChild(input,span);
  btn.style.display='none';
  input.focus();
}

// 인증유지의사 있음 - 다시 누르면 취소되는 토글
export function toggleMailingIntent(jobNo){
  const cur=(getMailingLog()[jobNo]||{}).intentKept||false;
  setMailingField(jobNo,'intentKept',!cur);
  logMailingAction(jobNo,'인증유지의사',!cur?'있음':'해제');
  renderMailingTab();
}

export function toggleMailingMonthCollapse(m){
  window._mailingCollapsed=window._mailingCollapsed||{};
  window._mailingCollapsed[m]=!window._mailingCollapsed[m];
  renderMailingTab();
}

export function setMailingEmailFilter(email){
  window._mailingEmailFilter=email;
  renderMailingTab();
}

function monthRowsToSheet(rows,mailingLog){
  return rows.map(r=>{
    const ml=mailingLog[r.jobNo]||{};
    return{
      '컨설팅 이메일주소':r.consultEmail||'',
      '컨설팅사':r.partner||'',
      '업체명':r.nameKo||r.nameEn||'',
      'Client number':r.jobNo,
      '이메일':r.email||'',
      '만료일':fmt(r.expDate),
      '인증서번호':r.certNo||'',
      '안내메일(1차)':noticeText(ml.notice1),
      '안내메일(2차)':noticeText(ml.notice2),
      '인증유지의사':ml.intentKept?'있음':'',
    };
  });
}
function mailingRowsToSheet(rows,mailingLog){
  return rows.map(r=>{
    const ml=mailingLog[r.jobNo]||{};
    return{
      '컨설팅 이메일주소':r.consultEmail||'',
      '컨설팅사':r.partner||'',
      '업체명(국문)':r.nameKo||'',
      'Organization':r.nameEn||'',
      'Client number':r.jobNo,
      '이메일':r.email||'',
      '만료일':fmt(r.expDate),
      '인증서번호':r.certNo||'',
      '진행 상태':ml.procStatus||'미정',
      '1차 안내':noticeText(ml.notice1),
      '2차 안내':noticeText(ml.notice2),
      '인증유지의사':ml.intentKept?'있음':'',
    };
  });
}
// 엑셀 열 너비/행 높이 지정: 컨설팅사 20, 업체명 25, 잡넘버/만료일은 내용 길이+앞뒤 스페이스 1칸씩, 인증서번호 15
function applyMailingLayout(ws,rows){
  if(!rows.length)return;
  const headers=Object.keys(rows[0]);
  const nameCol=headers.find(h=>h==='업체명'||h==='업체명(국문)');
  const jobCol=headers.find(h=>h==='Client number');
  const expCol=headers.find(h=>h==='만료일');
  const jobMaxLen=jobCol?Math.max(...rows.map(r=>String(r[jobCol]||'').length)):0;
  const expMaxLen=expCol?Math.max(...rows.map(r=>String(r[expCol]||'').length)):0;
  ws['!cols']=headers.map(h=>{
    if(h==='컨설팅사')return{wch:20};
    if(h===nameCol)return{wch:25};
    if(h===jobCol)return{wch:jobMaxLen+2};
    if(h==='이메일'||h==='컨설팅 이메일주소')return{wch:20};
    if(h===expCol)return{wch:expMaxLen+2};
    if(h==='인증서번호')return{wch:15};
    return{wch:14};
  });
  // 행간이 너무 빡빡하지 않도록 여유 있게 (기본 15pt보다 크게)
  ws['!rows']=new Array(rows.length+1).fill({hpt:20});
}
function downloadMailingSheet(rows){
  if(!rows.length){alert('다운로드할 데이터가 없어요.');return;}
  const ws=XLSX.utils.json_to_sheet(rows);
  applyMailingLayout(ws,rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'주기안내');
  XLSX.writeFile(wb,mailingFilename());
}

// 주기안내(월별): 화면에 표시된 현재+1~3개월 중 체크된 월만 다운로드
export function exportMailingMonthExcel(){
  const checkedMonths=new Set([...document.querySelectorAll('.mailing-month-dl:checked')].map(el=>Number(el.value)));
  if(!checkedMonths.size){alert('다운로드할 월을 하나 이상 선택해주세요.');return;}
  const all=buildMailingData();
  const mailingLog=getMailingLog();
  const targetMonths=[1,2,3].map(i=>(TODAY.getMonth()+i)%12);
  const rows=all.filter(r=>targetMonths.includes(r.expDate.getMonth())&&checkedMonths.has(r.expDate.getMonth()))
    .sort((a,b)=>a.expDate-b.expDate);
  downloadMailingSheet(monthRowsToSheet(rows,mailingLog));
}

// 주기안내(총): 현재 월 필터가 적용되어 있으면 그 달만, 없으면 전체
export function exportMailingTotalExcel(){
  const filter=window._mailingMonthFilter;
  const all=buildMailingData();
  const mailingLog=getMailingLog();
  const rows=((filter===null||filter===undefined)?all:all.filter(r=>r.expDate.getMonth()===filter))
    .sort((a,b)=>a.expDate-b.expDate);
  downloadMailingSheet(mailingRowsToSheet(rows,mailingLog));
}


// ── EMAIL ──

export function renderEmailTab(){
  const emails=getEmails();const tb=document.getElementById('tb-email');const keys=Object.keys(emails);
  if(!keys.length){tb.innerHTML='<tr class="empty"><td colspan="4">등록된 이메일 없음</td></tr>';return;}
  tb.innerHTML=keys.map(k=>{
    const entry=emails[k];
    const kEsc=k.replace(/'/g,"\\'");
    const aliases=(entry.aliases||[]).map(a=>{
      const aEsc=a.replace(/'/g,"\\'");
      return`<span class="alias-tag" title="더블클릭하여 편집" ondblclick="editAlias('${kEsc}','${aEsc}',this)">${a}<button onclick="event.stopPropagation();removeAlias('${kEsc}','${aEsc}')">×</button></span>`;
    }).join('');
    // 빈 영역 더블클릭 → 새 별칭 추가
    return`<tr>
      <td ondblclick="renameEmailEntry('${kEsc}')" title="더블클릭하여 컨설팅사명 수정" style="cursor:pointer">${k}</td>
      <td>${entry.email||''}</td>
      <td>${entry.phone||''}</td>
      <td>
        <div class="alias-tags" ondblclick="quickAddAlias(event,'${kEsc}')" title="빈 곳을 더블클릭하여 별칭 추가">
          ${aliases}
          <span class="alias-tag" style="border-style:dashed;color:var(--muted);cursor:pointer" ondblclick="event.stopPropagation();quickAddAlias(event,'${kEsc}')">+ 추가</span>
        </div>
      </td>
      <td><button class="tbtn" style="padding:3px 8px;font-size:11px" onclick="delEmail('${kEsc}')">삭제</button></td>
    </tr>`;
  }).join('');
}


export function renameEmailEntry(oldName){
  const newName=prompt('컨설팅사명 수정:',oldName);
  if(!newName||!newName.trim()||newName.trim()===oldName)return;
  const trimmed=newName.trim();
  const emails=getEmails();
  if(!emails[oldName])return;
  if(emails[trimmed]){alert(`"${trimmed}" 이름이 이미 있어요. 다른 이름을 사용해주세요.`);return;}
  emails[trimmed]=emails[oldName];
  delete emails[oldName];
  saveEmails(emails);
  renderEmailTab();
  showMsg(`"${oldName}" → "${trimmed}"(으)로 수정됐어요.`);
}

export function quickAddAlias(event, name){
  event.stopPropagation();
  const kw=prompt(`"${name}"에 추가할 별칭 키워드를 입력하세요:`);
  if(!kw||!kw.trim())return;
  const emails=getEmails();
  if(!emails[name])return;
  if(!emails[name].aliases)emails[name].aliases=[];
  if(!emails[name].aliases.includes(kw.trim()))emails[name].aliases.push(kw.trim());
  saveEmails(emails);renderEmailTab();showMsg(`별칭 "${kw.trim()}" 추가됐어요.`);
}


export function editAlias(name, oldKw, el){
  const newKw=prompt(`별칭 수정:`, oldKw);
  if(!newKw||!newKw.trim()||newKw.trim()===oldKw)return;
  const emails=getEmails();
  if(!emails[name]?.aliases)return;
  const idx=emails[name].aliases.indexOf(oldKw);
  if(idx>-1)emails[name].aliases[idx]=newKw.trim();
  saveEmails(emails);renderEmailTab();showMsg(`별칭이 "${newKw.trim()}"으로 수정됐어요.`);
}

export function addEmail(){
  const name=document.getElementById('ei-name').value.trim(),email=document.getElementById('ei-email').value.trim();
  const phone=document.getElementById('ei-phone').value.trim();
  if(!name||!email){showMsg('이름과 이메일을 입력해 주세요.');return;}
  const emails=getEmails();emails[name]={email,phone,aliases:emails[name]?.aliases||[]};saveEmails(emails);
  document.getElementById('ei-name').value='';document.getElementById('ei-email').value='';document.getElementById('ei-phone').value='';
  renderEmailTab();showMsg('저장됐어요.');
}

export function addAlias(){
  const name=document.getElementById('ei-alias-name').value.trim(),kw=document.getElementById('ei-alias-kw').value.trim();
  if(!name||!kw){showMsg('컨설팅사명과 별칭을 입력해 주세요.');return;}
  const emails=getEmails();if(!emails[name]){showMsg(`"${name}" 이름이 없어요.`);return;}
  if(!emails[name].aliases)emails[name].aliases=[];
  if(!emails[name].aliases.includes(kw))emails[name].aliases.push(kw);
  saveEmails(emails);
  document.getElementById('ei-alias-name').value='';document.getElementById('ei-alias-kw').value='';
  renderEmailTab();showMsg(`별칭 "${kw}" 추가됐어요.`);
}

export function removeAlias(name,kw){const emails=getEmails();if(emails[name]?.aliases)emails[name].aliases=emails[name].aliases.filter(a=>a!==kw);saveEmails(emails);renderEmailTab();}

export function delEmail(name){const emails=getEmails();delete emails[name];saveEmails(emails);renderEmailTab();}

export function showMsg(t){document.getElementById('emsg').textContent=t;setTimeout(()=>document.getElementById('emsg').textContent='',3000);}


// ── 수익보고 계산기 ──

export const CALC_TABLES=[
  {name:'일반',rows:{최초:2800000,사후:2300000,갱신:2300000,특별심사:2300000}},
  {name:'㈜ 비전경영기술원',rows:{최초:2300000,사후:1800000,갱신:1800000,특별심사:1800000}},
  {name:'한국 경영 정보',rows:{최초:2500000,사후:2000000,갱신:null,특별심사:null}},
  {name:'품질환경인증개발원',rows:{최초:2500000,사후:2000000,갱신:2000000,특별심사:null}},
];

export let calcMDVal=1.0;

export let calcSelectedAdmin=null; // {tableIdx, type, amount}


export function initCalcPage(){
  const wrap=document.getElementById('calc-admin-tables');
  if(!wrap||wrap.children.length>0)return;
  CALC_TABLES.forEach((tbl,ti)=>{
    const div=document.createElement('div');
    div.style.cssText='background:var(--panel);border:1px solid var(--border);border-radius:10px;overflow:hidden';
    const types=Object.keys(tbl.rows);
    div.innerHTML=`
      <div style="padding:8px 14px;background:var(--panel2);border-bottom:1px solid var(--border);font-weight:600;font-size:12px">${tbl.name}</div>
      <table style="width:100%;border-collapse:collapse">
        <tbody>
          ${types.map(type=>{
            const amt=tbl.rows[type];
            const id=`calc-r-${ti}-${type}`;
            return amt!==null?`
            <tr style="border-bottom:1px solid rgba(36,61,85,.3)">
              <td style="padding:7px 14px;font-size:12px">
                <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
                  <input type="radio" name="calc-admin" id="${id}" value="${ti}-${type}-${amt}"
                    style="width:14px;height:14px;accent-color:var(--accent)" onchange="calcSelectAdmin(${ti},'${type}',${amt})">
                  ${type}
                </label>
              </td>
              <td style="padding:7px 14px;font-size:12px;text-align:right;color:var(--text)">₩ ${amt.toLocaleString()}</td>
            </tr>`:'';
          }).join('')}
        </tbody>
      </table>`;
    wrap.appendChild(div);
  });
  calcUpdate();
}


export function calcSelectAdmin(ti,type,amt){
  calcSelectedAdmin={ti,type,amt};
  calcUpdate();
}


export function calcMD(delta){
  calcMDVal=Math.max(0.5,calcMDVal+delta);
  // 3.5 이상은 4.0으로 반올림
  if(calcMDVal>=3.5&&calcMDVal<4.0)calcMDVal=4.0;
  document.getElementById('calc-md-val').textContent=calcMDVal.toFixed(1);
  calcUpdate();
}


export function calcUpdate(){
  const md=calcMDVal;
  // 행정비
  const adminAmt=calcSelectedAdmin?calcSelectedAdmin.amt:0;
  // 지원심사비
  const supportCheck=document.getElementById('calc-support')?.checked;
  const supportAmt=supportCheck?400000*md:0;
  // 교통비
  const locVal=parseInt(document.getElementById('calc-location')?.value||'0');
  // 교통비 - M/D 반올림 후 곱하기 (1.5→2, 2.5→3)
  const mdRounded=Math.ceil(md);
  const transportAmt=locVal*mdRounded;
  const total=adminAmt+supportAmt+transportAmt;
  const fmt2=v=>v>0?'₩ '+v.toLocaleString():'-';
  document.getElementById('res-admin').textContent=fmt2(adminAmt);
  document.getElementById('res-support').textContent=supportCheck?'₩ '+supportAmt.toLocaleString():'-';
  document.getElementById('res-transport').textContent=locVal>0?'₩ '+transportAmt.toLocaleString():'-';
  document.getElementById('res-total').textContent=total>0?'₩ '+total.toLocaleString():'-';
}


export function copyCalcTotal(){
  const el=document.getElementById('res-total');
  if(!el||el.textContent==='-')return;
  // 숫자만 추출해서 복사 (₩, 콤마 제거)
  const num=el.textContent.replace(/[₩,\s]/g,'');
  navigator.clipboard.writeText(num).then(()=>{
    const msg=document.getElementById('calc-copy-msg');
    if(msg){msg.textContent='✅ 복사됐어요!';setTimeout(()=>msg.textContent='',2000);}
  });
}

export function calcReset(){
  calcMDVal=1.0;
  calcSelectedAdmin=null;
  document.getElementById('calc-md-val').textContent='1.0';
  document.querySelectorAll('input[name="calc-admin"]').forEach(r=>r.checked=false);
  const sup=document.getElementById('calc-support');if(sup)sup.checked=false;
  const loc=document.getElementById('calc-location');if(loc)loc.value='0';
  calcUpdate();
}

export function toggleTheme(){
  const isLight=document.body.classList.toggle('light');
  document.getElementById('theme-btn').textContent=isLight?'☀️':'🌙';
  LS.set('theme',isLight?'light':'dark');
}


// ── 탭 드래그 ──

export function initTabDrag(){
  const bar=document.getElementById('tab-bar');let dragSrc=null;
  bar.querySelectorAll('.tab').forEach(tab=>{
    tab.addEventListener('dragstart',e=>{dragSrc=tab;tab.classList.add('dragging');e.dataTransfer.effectAllowed='move';});
    tab.addEventListener('dragend',()=>bar.querySelectorAll('.tab').forEach(t=>t.classList.remove('dragging','drag-over')));
    tab.addEventListener('dragover',e=>{e.preventDefault();if(tab!==dragSrc)tab.classList.add('drag-over');});
    tab.addEventListener('dragleave',()=>tab.classList.remove('drag-over'));
    tab.addEventListener('drop',e=>{
      e.preventDefault();
      if(dragSrc&&dragSrc!==tab){
        const tabs=[...bar.querySelectorAll('.tab')];
        const si=tabs.indexOf(dragSrc),ti=tabs.indexOf(tab);
        if(si<ti)bar.insertBefore(dragSrc,tab.nextSibling);else bar.insertBefore(dragSrc,tab);
      }
      bar.querySelectorAll('.tab').forEach(t=>t.classList.remove('drag-over'));
    });
  });
}


// ── NAV ──

export function goTab(id,el){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  document.getElementById('page-'+id).classList.add('on');
  if(el)el.classList.add('on');
  else{const t=document.querySelector(`.tab[data-page="${id}"]`);if(t)t.classList.add('on');}
  if(id==='email')renderEmailTab();
  if(id==='mailing')renderMailingTab();
  if(id==='cal')renderCal();
  if(id==='notice')initNoticePage();
  if(id==='calc')initCalcPage();
  if(id==='log')renderLogPanel();
}

export function goTabById(id){goTab(id,null);}


// ── 단톡 공지 탭 ──

export let nfMultiItems=[]; // 복수건 업체 목록


export function initNoticePage(){
  nfMultiItems=[];
  const listEl=document.getElementById('nf-multi-list');
  if(listEl)listEl.innerHTML='';
  const nameEl=document.getElementById('nf-single-name');
  if(nameEl)nameEl.value='';
  const partEl=document.getElementById('nf-single-partner');
  if(partEl)partEl.textContent='';
  ['nf-single','nf-multi-r','nf-fssc22','nf-gic','nf-igc','nf-suspend','nf-revoke'].forEach(id=>{
    const el=document.getElementById(id);
    if(el)el.onchange=buildNotice;
  });
  buildNotice();
  // 외부 클릭 시 자동완성 닫기
  document.addEventListener('click',e=>{
    if(!e.target.closest('#nf-single-name')&&!e.target.closest('#nf-single-suggest'))
      hideSuggest('nf-single-suggest');
    if(!e.target.closest('#nf-multi-name')&&!e.target.closest('#nf-multi-suggest'))
      hideSuggest('nf-multi-suggest');
  });
}


export function hideSuggest(id){
  const el=document.getElementById(id);if(el)el.style.display='none';
}


export function noticeAutoComplete(inputId,suggestId){
  const input=document.getElementById(inputId);
  const suggest=document.getElementById(suggestId);
  if(!input||!suggest)return;
  const q=input.value.trim().toLowerCase();
  if(!q){suggest.style.display='none';return;}
  const matches=merged.filter(r=>(r.nameKo||r.nameEn||'').toLowerCase().includes(q)).slice(0,8);
  if(!matches.length){suggest.style.display='none';return;}
  suggest.innerHTML=matches.map(r=>`<div style="padding:7px 10px;cursor:pointer;font-size:12px;border-bottom:1px solid var(--border)" onmousedown="noticeSelect('${inputId}','${suggestId}','${(r.nameKo||r.nameEn).replace(/'/g,"\\'")}','${(r.partner||'').replace(/'/g,"\\'")}','${r.jobNo}')" onmouseover="this.style.background='var(--hover)'" onmouseout="this.style.background=''">${r.nameKo||r.nameEn} <span style="color:var(--muted);font-size:11px">/ ${r.partner||'-'}</span></div>`).join('');
  suggest.style.display='block';
}


export function noticeSelect(inputId,suggestId,name,partner,jobNo){
  const input=document.getElementById(inputId);
  if(input)input.value=name;
  input.dataset.jobNo=jobNo;
  input.dataset.partner=partner;
  hideSuggest(suggestId);
  if(inputId==='nf-single-name'){
    const partEl=document.getElementById('nf-single-partner');
    if(partEl)partEl.textContent=partner?`컨설팅사: ${partner}`:'';
    buildNotice();
  }else if(inputId==='nf-multi-name'){
    // 복수건: 목록에 추가
    if(!nfMultiItems.find(x=>x.jobNo===jobNo)){
      nfMultiItems.push({name,partner,jobNo});
      renderMultiList();
    }
    input.value='';
    buildNotice();
  }
}


export function noticePick(e,suggestId,inputId){
  // Enter 키로 첫 번째 항목 선택
  if(e.key==='Enter'){
    const suggest=document.getElementById(suggestId);
    if(suggest&&suggest.style.display!=='none'){
      const first=suggest.querySelector('div');
      if(first)first.dispatchEvent(new MouseEvent('mousedown'));
      e.preventDefault();
    }
  }
}


export function renderMultiList(){
  const el=document.getElementById('nf-multi-list');
  if(!el)return;
  el.innerHTML=nfMultiItems.map((item,i)=>`
    <div style="display:flex;align-items:center;gap:6px;padding:5px 8px;background:var(--panel2);border-radius:5px;font-size:12px">
      <span style="flex:1">${item.name} <span style="color:var(--muted);font-size:11px">/ ${item.partner||'-'}</span></span>
      <button style="background:none;border:none;color:var(--muted);cursor:pointer;font-size:13px" onclick="nfMultiItems.splice(${i},1);renderMultiList();buildNotice()">×</button>
    </div>`).join('');
  buildNotice();
}


export function buildNotice(){
  const isMulti=document.getElementById('nf-multi-r')?.checked;
  const type=document.querySelector('input[name="nf-type"]:checked')?.value||'정지';
  const schemes=[];
  if(document.getElementById('nf-fssc22')?.checked)schemes.push('FSSC 22000');
  if(document.getElementById('nf-gic')?.checked)schemes.push('GIC ISO 22000');
  if(document.getElementById('nf-igc')?.checked)schemes.push('ISO 22000');
  const schemeStr=schemes.join(', ')||'[규격 선택]';
  let lines=`아래 업체 ${schemeStr} 인증 ${type} 진행하겠습니다\n\n`;
  if(!isMulti){
    const nameEl=document.getElementById('nf-single-name');
    const name=nameEl?.value.trim()||'';
    const partner=nameEl?.dataset.partner||'';
    const reason=document.getElementById('nf-single-reason')?.value||'';
    if(!name){lines+='(업체명을 입력해 주세요)';}
    else{
      lines+=`- ${name}${partner?' / '+partner:''}\n`;
      if(reason)lines+=`  사유 : ${reason}`;
    }
  }else{
    const reason=document.getElementById('nf-multi-reason')?.value||'';
    if(!nfMultiItems.length){lines+='(업체를 추가해 주세요)';}
    else{
      nfMultiItems.forEach(item=>{
        lines+=`- ${item.name}${item.partner?' / '+item.partner:''}\n`;
      });
      if(reason)lines+=`  사유 : ${reason}`;
    }
  }
  const el=document.getElementById('nf-result');
  if(el)el.textContent=lines.trim();
}


export function copyNoticePage(){
  const text=document.getElementById('nf-result')?.textContent||'';
  if(!text||text==='설정 후 자동으로 생성됩니다')return;
  navigator.clipboard.writeText(text).then(()=>{
    const btn=document.querySelector('#page-notice .tbtn.pri');
    if(btn){btn.textContent='✅ 복사됨';setTimeout(()=>btn.textContent='📋 복사',2000);}
  });
}


// ── INIT ──

export function initApp() {
  setTxt('lbl-today',TODAY.toLocaleDateString('ko-KR',{year:'numeric',month:'long',day:'numeric',weekday:'short'}));
  if(LS.get('theme')==='light'){document.body.classList.add('light');document.getElementById('theme-btn').textContent='☀️';}
  renderEmailTab();renderCal();initTabDrag();

  // 로딩 표시
  document.getElementById('dash-content').innerHTML='<div style="padding:40px;text-align:center;color:var(--muted);font-size:13px">⏳ 데이터 불러오는 중...</div>';

  if(window.FS){
    // Firebase에서 모든 데이터 로드 후 렌더
    const keys=['emails2','checks','igcStore','overrides','edits','schedules','urgentLogs','mailingLog','mailingFsscRaw','mailingFsscName','mailingJobRaw','mailingJobName','fsscRaw','fsscName','jobRaw','jobName','igcName'];
    Promise.all(keys.map(async k=>{
      const val=await window.FS.load(k);
      if(val!==null){FSCache[k]=val;LS.set(k,val);}
    })).then(()=>{
      _startRender();
      // 실시간 동기화
      ['checks','overrides','igcStore','emails2','edits','schedules','urgentLogs','mailingLog'].forEach(key=>{
        window.FS.watch(key,val=>{
          FSCache[key]=val;LS.set(key,val);
          if(merged.length>0){
            if(key==='igcStore')mergeAndRender();
            else{renderDashContent();renderAll();}
          }
        });
      });
    }).catch(e=>{
      console.error('Firebase 로드 실패, 로컬로 시작:', e);
      _startRender();
    });
  } else {
    // Firebase 없으면 로컬로 시작
    _startRender();
  }
}


export function _startRender(){
  hydrateRawData();
  hydrateMailingRawData();
  if((fsscRaw&&jobRaw)||Object.keys(getIgcStore()).length>0) mergeAndRender();
  else renderDashContent(); // 빈 화면이라도 렌더
  renderEmailTab();
}

document.addEventListener('DOMContentLoaded', ()=>{
  if(window._fsReady){initApp();}
  else{
    document.addEventListener('fs-ready',initApp);
    setTimeout(()=>{if(!window._fsReady)initApp();},3000);
  }
});


// ── 하위 모듈(core/tab-*)에서 dispatch하는 재렌더 요청을 수신 ──
document.addEventListener('app:refresh-dash', ()=>{ renderDashContent(); renderAll(); });
document.addEventListener('app:refresh-all', ()=>{ renderDash(); renderAll(); renderCal(); renderUpd(); });
document.addEventListener('app:refresh-mailing', ()=>{ if(document.getElementById('page-mailing')?.classList.contains('on'))renderMailingTab(); });
document.addEventListener('click', e=>{
  if(e.target.closest && e.target.closest('[id^="proc-list-"]'))return;
  if(e.target.closest && e.target.closest('button')?.getAttribute('onclick')?.includes('toggleMailingProcDropdown'))return;
  document.querySelectorAll('[id^="proc-list-"]').forEach(el=>{el.style.display='none';});
});

// onclick="..." 문자열에서 참조하는 이 파일의 함수들을 전역에 노출
Object.assign(window, {addAlias, addEmail, addSchedule, buildNotice, calGoToday, calMove, calcMD, calcReset, calcSelectAdmin, calcUpdate, closeModal, copyCalcTotal, copyNoticePage, cycleUrgField, delEmail, delSchedule, editAlias, editMailingNoticeDate, exportMailingMonthExcel, exportMailingTotalExcel, goSubTab, goTab, goTabById, loadEmailFile, loadMailingFssc, loadMailingJob, noticeAutoComplete, noticePick, noticeSelect, openMailingDetail, openMailingMailDraft, openModal, quickAddAlias, removeAlias, renameEmailEntry, renderAll, renderCal, renderMultiList, saveMailingCheck, saveMailingField, saveMailingProc, saveReq, saveReqModal, saveReqReason, saveUpdDone, selectMailingProc, setMailingEmailFilter, setMailingMonthFilter, setMailingNoticeMode, setMailingNoticeStatus, showCalPopup, switchMailingView, toggleEditMode, toggleHidden, toggleMailingIntent, toggleMailingMonthCollapse, toggleMailingProcDropdown, toggleTheme});
