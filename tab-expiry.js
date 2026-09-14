// tab-expiry.js — 만료 예정 업체 / 요청 업체 섹션

import {TODAY, ddayChip, fmt, getChecks, requestDashRefresh, saveChecks, setShowHidden} from './core.js';

export function renderMonthRows(rows,hiddenRows){
  const all=[...rows.map(r=>({r,hidden:false})),...hiddenRows.map(r=>({r,hidden:true}))];
  if(!all.length)return'<tr class="empty"><td colspan="5">해당 없음</td></tr>';
  const checks=getChecks();
  return all.map(({r,hidden})=>{
    const st=hidden?'opacity:.45':'';
    const undo=hidden?`<td><button class="tbtn" style="padding:2px 7px;font-size:10px;color:#ff6b6b;border-color:#ff6b6b" onclick="event.stopPropagation();undoIntent('${r.jobNo}')">되돌리기</button></td>`:'<td></td>';
    const c=checks[r.jobNo]||{};
    // FSSC이고 결정일(Date decision)이 있으면 그 날짜를, 없으면 만료일을 표시
    const baseDate=(r.src==='FSSC'&&r.initialCert)?r.initialCert:r.expDate;
    // 업로드 완료 시 다음 년도 D-day
    const dday=c.monthUpload?
      `<span class="chip cg">D-${diffNextYear(r.expDate)} (갱신)</span>`:
      ddayChip(r.expDate);
    return`<tr onclick="openModal(${r.idx},'month')" style="${st}"><td title="${r.partner}">${r.partner||'-'}</td><td title="${r.nameKo}">${r.nameKo||r.nameEn||'-'}</td><td>${r.jobNo}</td><td>${dday} ${fmt(baseDate)}</td>${undo}</tr>`;
  }).join('');
}

export function renderReqRows(rows,checks){
  if(!rows.length)return'<tr class="empty"><td colspan="4">요청된 업체 없음</td></tr>';
  const isLight=document.body.classList.contains('light');
  return rows.map(r=>{
    const c=checks[r.jobNo]||{};
    const tags=[
      c.suspend?`<div><span class="chip co" style="margin-right:3px">정지</span>${c.suspendReason?`<span style="font-size:10px;color:var(--muted)">${c.suspendReason}</span>`:''}</div>`:'',
      c.revoke?`<div><span class="chip cr" style="margin-right:3px">철회</span>${c.revokeReason?`<span style="font-size:10px;color:var(--muted)">${c.revokeReason}</span>`:''}</div>`:'',
      c.suspend1?`<div><span class="chip co" style="margin-right:3px">정지(1차)</span></div>`:'',
      c.suspend2?`<div><span class="chip co">정지(2차)</span></div>`:'',
      c.revoke1?`<div><span class="chip cr" style="margin-right:3px">철회(1차)</span></div>`:'',
      c.revoke2?`<div><span class="chip cr">철회(2차)</span></div>`:''
    ].join('');
    let rowBg='';
    if(c.reqNotice){
      rowBg=isLight?'background:#a9d18e;color:#1a3a1a':'background:#1a3a1a';
    }
    return`<tr onclick="openModal(${r.idx},'req')" style="${rowBg}"><td title="${r.partner}">${r.partner||'-'}</td><td title="${r.nameKo}">${r.nameKo||r.nameEn||'-'}</td><td>${r.jobNo}</td><td>${tags}</td></tr>`;
  }).join('');
}

export function toggleMonthSort(){
  window._monthSort=window._monthSort==='dday'?'date':'dday';
  requestDashRefresh();
}

export function diffNextYear(d){
  if(!d)return null;
  // 업로드 완료 시 다음 년도 기준 D-day
  const next=new Date(TODAY.getFullYear()+1,d.getMonth(),d.getDate());
  return Math.round((next-TODAY)/86400000);
}


export function undoIntent(jobNo){
  const c=getChecks();if(c[jobNo]){c[jobNo]['intent1']=false;saveChecks(c);}
  setShowHidden(false);requestDashRefresh();
}


// onclick="..." 문자열에서 참조하는 이 파일의 함수들을 전역에 노출
Object.assign(window, {toggleMonthSort, undoIntent});
